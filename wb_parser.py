# -*- coding: utf-8 -*-
"""
Парсер Wildberries с историей цен.
Особенности:
- Всегда использует ценовые фильтры для обхода ограничения 10 000 товаров.
- Поддерживает строковые ID продавцов (например, "positronica").
- Сохраняет историю цен в par.xlsx (отдельный лист на продавца).
- Детальное логирование каждого этапа обработки Excel.
"""

import asyncio
import random
import os
import time
from datetime import datetime

import pandas as pd
from patchright.async_api import async_playwright


# ===================================================================
# КОНФИГУРАЦИЯ – измените здесь имена файлов и параметры при необходимости
# ===================================================================
SELLERS_FILE = "sellers.xlsx"       # файл со списком продавцов (первый столбец – ID, второй – флаг 1/0)
PRICE_HISTORY_FILE = "par.xlsx"     # файл для сохранения истории цен

KEEP_CARDS = 700                    # количество оставляемых карточек на странице (для лёгкого DOM)
MAX_STALL = 5                       # число итераций без новых товаров для остановки
HEADLESS = False                    # True – браузер без интерфейса
MAX_GOTO_ATTEMPTS = 3               # попытки загрузки страницы при сбоях сети
# ===================================================================



# Диапазоны цен в копейках (смежные, без пропусков: каждый следующий начинается на 1 копейку позже)
PRICE_RANGES = [
    (100, 14999),
    (15000, 19999),
    (20000, 29999),
    (30000, 49999),
    (50000, 59999),
    (60000, 69999),
    (70000, 79999),
    (80000, 89999),
    (90000, 99999),
    (100000, 129999),
    (130000, 149999),
    (150000, 199999),
    (200000, 299999),
    (300000, 399999),
    (400000, 499999),
    (500000, 799999),
    (800000, 999999),
    (1000000, 1999999),
    (2000000, 100000000000)
]



# ---------- 1. Чтение списка продавцов из Excel ----------
def get_sellers_to_parse() -> list[dict]:
    """
    Читает файл SELLERS_FILE и возвращает список активных продавцов:
    [{'seller_id': '...', 'split': True/False}, ...]

    - активный продавец: во втором столбце стоит 1;
    - split=True  (столбец «Разбивка» = 1) – парсить с разбивкой по ценовым диапазонам;
    - split=False (столбец «Разбивка» = 0) – парсить одним запросом без разбивки;
    - если столбца «Разбивка» нет или значение пустое – по умолчанию True.

    Первая строка может содержать заголовки (например «ПРОДАВЕЦ | ПАРСИТЬ? | … | Разбивка») —
    она определяется по нечисловым значениям ID и флага и пропускается.
    Если файл не найден, создаёт пример.
    """
    filename = SELLERS_FILE
    if not os.path.exists(filename):
        print(f"Файл {filename} не найден. Создаю пример файла.")
        example_df = pd.DataFrame({
            'seller_id': [4072488, 'positronica', 1234567],
            'active': [1, 1, 0],
            'split': [1, 0, 1]
        })
        example_df.to_excel(filename, index=False)
        print(f"Создан пример файла {filename}. Отредактируйте его и запустите скрипт снова.")
        return []

    df = pd.read_excel(filename, header=None, dtype=str)
    sellers = []
    for idx, row in df.iterrows():
        try:
            seller_id = str(row.iloc[0]).strip()
            if not seller_id:
                continue
            flag_str = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else '0'
            # Пропускаем строку заголовков (например «ПРОДАВЕЦ | ПАРСИТЬ?»):
            # у заголовков и ID, и флаг не являются числами
            if not seller_id.replace('.', '', 1).isdigit() and not flag_str.replace('.', '', 1).isdigit():
                print(f"Пропускаю строку {idx+1}: строка заголовков")
                continue
            flag = int(float(flag_str)) if flag_str.replace('.', '', 1).isdigit() else 0
            if flag != 1:
                continue
            # Столбец «Разбивка» (5-й, индекс 4): 1 – с разбивкой, 0 – без.
            # Если столбца нет или значение пустое – по умолчанию с разбивкой.
            split_str = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else '1'
            split = (int(float(split_str)) if split_str.replace('.', '', 1).isdigit() else 1) == 1
            sellers.append({'seller_id': seller_id, 'split': split})
        except (ValueError, IndexError, TypeError) as e:
            print(f"Пропускаю строку {idx+1}: неверный формат данных ({e})")
            continue
    print(f"Найдено активных продавцов: {len(sellers)}")
    if sellers:
        print(f"ID: {', '.join(s['seller_id'] for s in sellers[:5])}{'...' if len(sellers)>5 else ''}")
    return sellers


# ---------- 2. Сбор товаров одного продавца (с поддержкой URL-фильтра) ----------


async def create_browser(p):
    """Запускает браузер и создаёт контекст (общий для всех диапазонов)."""
    browser = await p.chromium.launch(headless=HEADLESS)
    context = await browser.new_context(
        viewport={'width': 1280, 'height': 720},
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    )
    return browser, context


async def get_seller_products_fast(seller_id: str, url_override: str = None, log_file=None, context=None) -> list[dict]:
    """
    Собирает товары продавца со страницы с ценовым фильтром.

    Оптимизации и устойчивость:
    - Инкрементальный сбор: каждая итерация обрабатывает ТОЛЬКО новые карточки,
      добавленные в конец DOM (счётчик window.__wb_processed), а не все ~700 заново.
    - Сбор и удаление старых карточек выполняются одним вызовом page.evaluate.
    - textContent вместо innerText (не вызывает принудительный reflow страницы).
    - Для каждого вызова создаётся СВЕЖАЯ страница (context.new_page()), поэтому
      «протухшие» соединения и состояние предыдущего диапазона не переносятся.
    - При сбое сети/таймауте навигация повторяется (до MAX_GOTO_ATTEMPTS) с паузой.
    - Если context не передан, создаётся собственный браузер (автономный запуск).
    """
    if url_override:
        url = url_override
    else:
        url = f"https://www.wildberries.ru/seller/{seller_id}"

    start_total = time.time()
    start_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def log(msg):
        print(msg)
        if log_file:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            log_file.write(f"[{timestamp}] {msg}\n")
            log_file.flush()

    log(f"Начинаю сбор для продавца {seller_id} (URL: {url}) (старт: {start_str})")

    owns_browser = context is None
    async with async_playwright() as p:
        if owns_browser:
            browser, context = await create_browser(p)
        else:
            browser = None

        # Свежая страница на каждый вызов: «протухшие» соединения и состояние
        # предыдущего диапазона не переносятся между диапазонами
        page = await context.new_page()
        # Блокируем изображения, шрифты, медиа
        await page.route(
            '**/*',
            lambda route: route.abort() if route.request.resource_type in ['image', 'font', 'media'] else route.continue_()
        )

        log("Загружаю страницу...")
        load_start = time.time()
        # На WB networkidle не срабатывает (постоянные фоновые запросы) — ждём только DOM.
        # При сбое сети (таймаут/обрыв соединения) повторяем навигацию с нарастающей паузой.
        goto_ok = False
        for attempt in range(1, MAX_GOTO_ATTEMPTS + 1):
            try:
                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                except Exception as e:
                    log(f"Не дождался domcontentloaded (попытка {attempt}/{MAX_GOTO_ATTEMPTS}, {type(e).__name__}), пробую commit...")
                    await page.goto(url, wait_until="commit", timeout=30000)
                goto_ok = True
                break
            except Exception as e:
                log(f"Попытка {attempt}/{MAX_GOTO_ATTEMPTS}: {type(e).__name__}: {e}")
                if attempt < MAX_GOTO_ATTEMPTS:
                    # сбрасываем состояние навигации страницы перед следующей попыткой
                    try:
                        await page.goto("about:blank", wait_until="commit", timeout=10000)
                    except Exception:
                        pass
                    await asyncio.sleep(5 * attempt)
        if not goto_ok:
            log(f"Не удалось загрузить {url} за {MAX_GOTO_ATTEMPTS} попыток. Диапазон пропущен.")
            try:
                await page.close()
            except Exception:
                pass
            log("----------------------------------------\n")
            return []
        # Ждём появления карточек товаров, иначе первая итерация будет пустой
        try:
            await page.wait_for_selector(
                '.product-card, .j-card, [data-card-id]',
                state='attached',
                timeout=20000
            )
            await page.wait_for_timeout(1000)
            log("Карточки товаров появились в DOM")
        except Exception as e:
            log(f"Карточки не появились в DOM за 20 сек ({type(e).__name__}: {e})")
        log(f"Страница загружена за {time.time() - load_start:.2f} сек")

        # Сбрасываем счётчики обработки (важно при переиспользовании страницы между диапазонами)
        await page.evaluate(f"window.__wb_processed = 0; window.__wb_keep = {KEEP_CARDS};")

        all_data = []
        seen_ids = set()
        stall_count = 0
        iter_num = 0

        log(f"--- Начинаю сбор (keep_cards={KEEP_CARDS}, max_stall={MAX_STALL}) ---")

        while stall_count < MAX_STALL:
            iter_start = time.time()
            iter_num += 1

            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            # Умеренная задержка между прокрутками (для снижения риска блокировки);
            # на «простаивающих» (пустых) итерациях ждём меньше
            await page.wait_for_timeout(random.randint(400, 700) if stall_count > 0 else random.randint(800, 1500))

            collect_start = time.time()
            cards_data = await page.evaluate('''
                () => {
                    const KEEP = (typeof window.__wb_keep === 'number') ? window.__wb_keep : 700;
                    const cards = document.querySelectorAll('.product-card, .j-card, [data-card-id]');
                    const total = cards.length;
                    if (typeof window.__wb_processed !== 'number') window.__wb_processed = 0;
                    const start = Math.min(window.__wb_processed, total);

                    // Обрабатываем только НОВЫЕ карточки (добавленные в конец DOM)
                    const items = [];
                    for (let i = start; i < total; i++) {
                        const card = cards[i];
                        try {
                            let id = card.getAttribute('data-id') || card.getAttribute('data-nm-id');
                            if (!id) {
                                const link = card.querySelector('a');
                                if (link && link.href) {
                                    const match = link.href.match(/\\/catalog\\/(\\d+)\\//);
                                    if (match) id = match[1];
                                }
                            }
                            if (!id) continue;

                            let name = '';
                            const nameEl = card.querySelector('[data-testid="productName"]');
                            if (nameEl) name = (nameEl.textContent || '').replace(/\\s+/g, ' ').trim();
                            if (!name) {
                                const selectors = ['.product-card__name', '.goods-name', '[data-name]', '.j-card__name', '.product__name'];
                                for (let sel of selectors) {
                                    let el = card.querySelector(sel);
                                    if (el) {
                                        name = (el.textContent || '').replace(/\\s+/g, ' ').trim();
                                        if (name) break;
                                    }
                                }
                            }
                            if (!name) {
                                const lines = (card.textContent || '').split('\\n').map(s => s.trim()).filter(s => s.length > 10);
                                for (let line of lines) {
                                    if (!/^[\\d\\s₽$€%#]+$/.test(line)) {
                                        name = line.replace(/\\s+/g, ' ');
                                        break;
                                    }
                                }
                            }

                            let brand = '';
                            const brandEl = card.querySelector('[data-testid="productBrand"]');
                            if (brandEl) brand = (brandEl.textContent || '').trim();
                            if (!brand) {
                                const selectors = ['.product-card__brand', '.brand-name', '[data-brand]', '.j-card__brand', '.product__brand', '.brand'];
                                for (let sel of selectors) {
                                    let el = card.querySelector(sel);
                                    if (el) {
                                        brand = (el.textContent || '').trim();
                                        if (brand) break;
                                    }
                                }
                            }

                            let price = 0;
                            const priceElem = card.querySelector('.price__lower-price, .price__wrap, .product-card__price, .price, .sale-price');
                            if (priceElem) {
                                const match = (priceElem.textContent || '').match(/(\\d[\\d\\s]*\\d)/);
                                if (match) {
                                    const num = parseFloat(match[1].replace(/\\s/g, ''));
                                    if (!isNaN(num)) price = num;
                                }
                            }

                            items.push({ id, name, brand, price });
                        } catch(e) {
                            console.error('[DEBUG] Ошибка при обработке карточки ' + i + ':', e);
                        }
                    }
                    window.__wb_processed = total;

                    // Удаляем старые карточки с начала списка, чтобы DOM оставался лёгким
                    if (total > KEEP) {
                        const toRemove = total - KEEP;
                        for (let i = 0; i < toRemove; i++) {
                            if (cards[i] && cards[i].parentNode) {
                                cards[i].remove();
                            }
                        }
                        window.__wb_processed -= toRemove;
                        if (window.__wb_processed < 0) window.__wb_processed = 0;
                    }
                    return items;
                }
            ''')
            collect_elapsed = time.time() - collect_start

            new_items = []
            for item in cards_data:
                if item['id'] not in seen_ids:
                    seen_ids.add(item['id'])
                    all_data.append(item)
                    new_items.append(item)

            log(f"Итерация {iter_num}: собрано {len(cards_data)} карточек, "
                f"из них новых: {len(new_items)} (всего {len(all_data)}) | "
                f"сбор занял {collect_elapsed:.2f} сек")

            if new_items:
                stall_count = 0
            else:
                stall_count += 1
                log(f"Новых товаров нет ({stall_count}/{MAX_STALL})")

            total_iter_elapsed = time.time() - iter_start
            log(f"Итерация заняла {total_iter_elapsed:.2f} сек")

        await page.wait_for_timeout(1000)
        log("--- Сбор завершён ---")

        # Закрываем страницу (браузер/контекст остаются жить для следующих диапазонов)
        try:
            await page.close()
        except Exception:
            pass
        if browser is not None:
            await browser.close()

        total_elapsed = time.time() - start_total
        log(f"Общее время: {total_elapsed:.2f} сек")
        log(f"Собрано товаров: {len(all_data)}")
        log("----------------------------------------\n")
        return all_data








# ---------- 3. Обновление истории цен в Excel (с детальным логированием) ----------

def update_price_history(new_products: list[dict], seller_id: str, log_file=None):
    """
    Обновляет Excel-файл с историей цен, используя векторизованные операции pandas.
    Если файл повреждён, переименовывает его и создаёт новый.
    Гарантированно удаляет столбец "price" из итогового DataFrame.
    """
    start_time = time.time()

    def log(msg):
        print(msg)
        if log_file:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            log_file.write(f"[{timestamp}] {msg}\n")
            log_file.flush()

    log(f"Начинаю обновление истории цен для продавца {seller_id} (товаров: {len(new_products)})")
    filename = PRICE_HISTORY_FILE
    sheet_name = str(seller_id)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # --- Попытка загрузить существующий файл ---
    all_sheets = {}
    file_exists = os.path.exists(filename)
    file_ok = False

    if file_exists:
        try:
            load_start = time.time()
            with pd.ExcelFile(filename) as xls:
                for sheet in xls.sheet_names:
                    all_sheets[sheet] = pd.read_excel(
                        xls,
                        sheet_name=sheet,
                        dtype={'id': str, 'seller_id': str}
                    )
            file_ok = True
            log(f"Загружен файл {filename}. Листы: {list(all_sheets.keys())} (время: {time.time()-load_start:.2f} сек)")
        except Exception as e:
            log(f"Ошибка при открытии {filename}: {e}. Файл повреждён, переименовываю.")
            backup_name = f"par_broken_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            os.rename(filename, backup_name)
            log(f"Файл переименован в {backup_name}. Будет создан новый файл.")
            file_exists = False
            all_sheets = {}

    if not file_exists:
        log(f"Файл {filename} не найден или повреждён, создаю новый.")

    # --- Получение или создание DataFrame для текущего продавца ---
    if file_ok and sheet_name in all_sheets:
        df = all_sheets[sheet_name]
        # Удаляем лишний столбец "price", если он есть
        if 'price' in df.columns:
            df.drop(columns=['price'], inplace=True)
            log("Удалён столбец 'price' из загруженного листа.")
        log(f"Загружен лист '{sheet_name}'. Записей: {len(df)}")
    else:
        df = pd.DataFrame(columns=['id', 'name', 'brand', 'seller_id'])
        log(f"Создан новый лист '{sheet_name}'.")

    # --- Приведение типов ---
    prep_start = time.time()
    for col in ['id', 'name', 'brand', 'seller_id']:
        if col not in df.columns:
            df[col] = ''
    df['id'] = df['id'].astype(str)
    df['seller_id'] = df['seller_id'].astype(str)
    log(f"Подготовка DataFrame заняла {time.time()-prep_start:.2f} сек")

    # --- Создание DataFrame из новых товаров ---
    dict_start = time.time()
    new_df = pd.DataFrame(new_products)
    new_df['id'] = new_df['id'].astype(str)
    new_df['seller_id'] = str(seller_id)

    # Сохраняем цены отдельно, затем удаляем столбец price из new_df.
    # ВАЖНО: индексируем по id ДО копирования цен, иначе у price_series
    # останется позиционный RangeIndex и цены никогда не попадут в историю.
    new_df.set_index('id', inplace=True)
    price_series = new_df['price'].copy()
    new_df.drop(columns=['price'], inplace=True)

    log(f"Создание DataFrame из новых товаров заняло {time.time()-dict_start:.2f} сек")

    # --- Векторизованное обновление ---
    update_start = time.time()
    df.set_index('id', inplace=True)

    # Обновляем существующие записи (name, brand, seller_id)
    df.update(new_df[['name', 'brand', 'seller_id']])

    # Добавляем новые записи (без столбца price)
    new_ids = new_df.index.difference(df.index)
    if len(new_ids) > 0:
        new_rows = new_df.loc[new_ids].copy()
        date_cols = [col for col in df.columns if col not in ['name', 'brand', 'seller_id']]
        for col in date_cols:
            new_rows[col] = 0.0
        df = pd.concat([df, new_rows])

    # Добавляем новый столбец с датой и записываем цены из price_series
    df[now_str] = 0.0
    common_ids = df.index.intersection(price_series.index)
    df.loc[common_ids, now_str] = price_series[common_ids]

    df.reset_index(inplace=True)
    log(f"Обновление и добавление записей (векторизовано) заняло {time.time()-update_start:.2f} сек")

    # --- Сортировка ---
    sort_start = time.time()
    try:
        df['id_num'] = pd.to_numeric(df['id'], errors='coerce')
        df = df.sort_values('id_num').drop(columns=['id_num'])
        log(f"Сортировка заняла {time.time()-sort_start:.2f} сек")
    except Exception as e:
        log(f"Сортировка не удалась: {e} (время: {time.time()-sort_start:.2f} сек)")

    # --- Переупорядочивание колонок ---
    reorder_start = time.time()
    date_cols = [col for col in df.columns if col not in ['id', 'name', 'brand', 'seller_id']]
    ordered_columns = ['id', 'name', 'brand', 'seller_id'] + date_cols
    df = df[ordered_columns]
    log(f"Переупорядочивание колонок заняло {time.time()-reorder_start:.2f} сек")

    # --- ФИНАЛЬНАЯ ЗАЩИТА: удаляем столбец price, если он всё ещё есть ---
    if 'price' in df.columns:
        df.drop(columns=['price'], inplace=True)
        log("Удалён случайный столбец 'price' перед записью.")

    # --- Запись в Excel ---
    write_start = time.time()
    if file_ok:
        all_sheets[sheet_name] = df
    else:
        all_sheets = {sheet_name: df}

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet, data_df in all_sheets.items():
            data_df.to_excel(writer, sheet_name=sheet, index=False)
            worksheet = writer.sheets[sheet]
            for column in worksheet.columns:
                max_len = 0
                col_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except Exception:
                        pass
                adjusted = min(max_len + 2, 50)
                worksheet.column_dimensions[col_letter].width = adjusted
    log(f"Запись в Excel заняла {time.time()-write_start:.2f} сек")

    total_elapsed = time.time() - start_time
    log(f"Обновление истории цен для продавца {seller_id} завершено за {total_elapsed:.2f} сек")
    print(f"История цен сохранена в {filename} на листе '{sheet_name}'. Добавлен столбец: {now_str}")


# ---------- 4. Вспомогательные функции и главная функция ----------

def merge_products(products_range: list, all_products: list, global_seen_ids: set, log_file) -> int:
    """Добавляет новые (ещё не встречавшиеся) товары в общий список. Возвращает число добавленных."""
    added = 0
    for prod in products_range:
        if prod['id'] not in global_seen_ids:
            global_seen_ids.add(prod['id'])
            all_products.append(prod)
            added += 1
    print(f"Добавлено {added} новых товаров. Всего: {len(all_products)}")
    log_file.write(f"Добавлено {added} новых товаров. Всего: {len(all_products)}\n")
    return added


async def scrape_url(seller_id: str, url: str, context, log_file) -> list[dict]:
    """Парсит один URL и возвращает список товаров (или [] при ошибке)."""
    try:
        return await get_seller_products_fast(
            seller_id,
            url_override=url,
            log_file=log_file,
            context=context
        )
    except Exception as e:
        print(f"Ошибка при парсинге {url}: {type(e).__name__}: {e}")
        log_file.write(f"Ошибка при парсинге {url}: {type(e).__name__}: {e}\n")
        return []


async def handle_empty_result(browser, context, p, consecutive_failures: int, log_file):
    """
    Реакция на пустой результат запроса: растущая пауза, а при серии сбоев —
    пересоздание браузера (сброс соединений). Возвращает (browser, context, consecutive_failures).
    """
    consecutive_failures += 1
    delay = min(30 * consecutive_failures, 120)
    print(f"Запрос не дал результатов ({consecutive_failures} подряд). Пауза {delay} сек...")
    log_file.write(f"Запрос пуст ({consecutive_failures} подряд). Пауза {delay} сек.\n")
    await asyncio.sleep(delay)
    if consecutive_failures >= 2:
        print("Пересоздаю браузер (сброс соединений)...")
        log_file.write("Пересоздаю браузер (сброс соединений).\n")
        await browser.close()
        browser, context = await create_browser(p)
        consecutive_failures = 0
    return browser, context, consecutive_failures


async def main():
    """
    Запускает парсинг для всех активных продавцов.

    Режимы (столбец «Разбивка» в sellers.xlsx):
    - split=True  – парсинг по ценовым диапазонам (для больших каталогов);
    - split=False – парсинг одним запросом без разбивки (для небольших каталогов).
    """
    log_filename = f"timing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(log_filename, 'w', encoding='utf-8') as log_file:
        log_file.write(f"=== Лог таймингов парсинга Wildberries ===\n")
        log_file.write(f"Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_file.write(f"Используемые ценовые диапазоны: {len(PRICE_RANGES)}\n")
        log_file.write("============================================\n\n")

        sellers = get_sellers_to_parse()
        if not sellers:
            print("Нет продавцов для парсинга.")
            return

        print(f"\nНачинаю парсинг. Продавцов: {len(sellers)}.\n")

        # Один браузер на весь прогон; для каждого запроса создаётся свежая страница
        async with async_playwright() as p:
            browser, context = await create_browser(p)
            try:
                consecutive_failures = 0
                for idx, seller in enumerate(sellers, 1):
                    seller_id = seller['seller_id']
                    split = seller['split']

                    print(f"\n{'='*50}")
                    print(f"Продавец {idx}/{len(sellers)}: ID = {seller_id}")
                    print(f"{'='*50}")
                    log_file.write(f"\n--- Продавец {idx}/{len(sellers)}: {seller_id} ---\n")

                    if split:
                        print(f"Режим: разбивка по {len(PRICE_RANGES)} ценовым диапазонам")
                        log_file.write(f"Режим: разбивка по {len(PRICE_RANGES)} ценовым диапазонам\n")
                    else:
                        print("Режим: без разбивки (один запрос)")
                        log_file.write("Режим: без разбивки (один запрос)\n")

                    all_products = []
                    global_seen_ids = set()

                    if split:
                        for min_price, max_price in PRICE_RANGES:
                            filter_url = f"https://www.wildberries.ru/seller/{seller_id}?priceU={min_price}%3B{max_price}"
                            print(f"\n--- Парсинг диапазона: {min_price/100:.0f}–{max_price/100:.0f} руб. ---")
                            log_file.write(f"--- Диапазон: {min_price/100:.0f}–{max_price/100:.0f} руб. ---\n")

                            products_range = await scrape_url(seller_id, filter_url, context, log_file)
                            merge_products(products_range, all_products, global_seen_ids, log_file)

                            if not products_range:
                                browser, context, consecutive_failures = await handle_empty_result(
                                    browser, context, p, consecutive_failures, log_file
                                )
                            else:
                                consecutive_failures = 0

                            await asyncio.sleep(random.randint(5, 10))
                    else:
                        url = f"https://www.wildberries.ru/seller/{seller_id}"
                        print(f"\n--- Парсинг без разбивки: {url} ---")
                        log_file.write(f"--- Парсинг без разбивки: {url} ---\n")

                        products_range = await scrape_url(seller_id, url, context, log_file)
                        merge_products(products_range, all_products, global_seen_ids, log_file)

                        if not products_range:
                            browser, context, consecutive_failures = await handle_empty_result(
                                browser, context, p, consecutive_failures, log_file
                            )
                        else:
                            consecutive_failures = 0

                        await asyncio.sleep(random.randint(5, 10))

                    if all_products:
                        update_price_history(all_products, seller_id, log_file)
                        print(f"✅ Данные для продавца {seller_id} успешно сохранены (всего {len(all_products)} товаров).")
                        log_file.write(f"✅ Данные для продавца {seller_id} сохранены (всего {len(all_products)}).\n")
                    else:
                        print(f"❌ Не удалось собрать данные для продавца {seller_id}.")
                        log_file.write(f"❌ Не удалось собрать данные для продавца {seller_id}.\n")

                    if idx < len(sellers):
                        delay = random.randint(10, 20)
                        print(f"Ожидание {delay} секунд перед следующим продавцом...")
                        log_file.write(f"Ожидание {delay} секунд перед следующим продавцом...\n")
                        await asyncio.sleep(delay)
            finally:
                await browser.close()

        print("\n🎉 Все продавцы обработаны!")
        log_file.write("\n🎉 Все продавцы обработаны!\n")

    print(f"\nЛог таймингов сохранён в файл: {log_filename}")


# ---------- Точка входа ----------
if __name__ == "__main__":
    asyncio.run(main())