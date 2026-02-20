import time
import random
import re
import json
import csv
from typing import Optional, List, Tuple
import requests
from tqdm import tqdm
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from openpyxl import Workbook
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

import config
from config import setup_logging
from uc_wire_tunnel import UCWithTunnel
from proxy_manager import ProxyManager
from gsheets import safe_batch_update, col_letter_to_index, apply_cell_colors, get_sheet_client, col_index_to_letter

logger = setup_logging("wb_parser")


def random_delay(min_sec=None, max_sec=None):
    min_sec = min_sec or config.RANDOM_DELAY_MIN
    max_sec = max_sec or config.RANDOM_DELAY_MAX
    time.sleep(random.uniform(min_sec, max_sec))


def detect_link_type(url: str) -> str:
    v = url.strip()
    if v.isdigit():
        return 'wb'
    if not v.lower().startswith('http'):
        return 'skip'
    if 'wildberries' in v.lower() or 'wb.ru' in v.lower():
        return 'wb'
    if 'ozon.ru' in v.lower() or 'ozon.by' in v.lower():
        return 'ozon'
    return 'skip'


def build_wb_url(value: str) -> str:
    value = value.strip()
    if value.startswith("http"):
        return value
    return f"https://www.wildberries.ru/catalog/{value}/detail.aspx"


def extract_nm_id(url: str) -> Optional[str]:
    match = re.search(r'/catalog/(\d+)/', url)
    return match.group(1) if match else None


def get_sku_url_data(sku: str):
    sku = str(sku)
    part = sku[:-3]
    vol = int(part[:-2]) if len(part) > 2 else 0

    if vol <= 143:
        basket = "01"
    elif vol <= 287:
        basket = "02"
    elif vol <= 431:
        basket = "03"
    elif vol <= 719:
        basket = "04"
    elif vol <= 1007:
        basket = "05"
    elif vol <= 1061:
        basket = "06"
    elif vol <= 1115:
        basket = "07"
    elif vol <= 1169:
        basket = "08"
    elif vol <= 1313:
        basket = "09"
    elif vol <= 1601:
        basket = "10"
    elif vol <= 1655:
        basket = "11"
    elif vol <= 1919:
        basket = "12"
    elif vol <= 2045:
        basket = "13"
    elif vol <= 2189:
        basket = "14"
    elif vol <= 2405:
        basket = "15"
    elif vol <= 2621:
        basket = "16"
    elif vol <= 2837:
        basket = "17"
    elif vol <= 3053:
        basket = "18"
    elif vol <= 3269:
        basket = "19"
    elif vol <= 3485:
        basket = "20"
    elif vol <= 3701:
        basket = "21"
    elif vol <= 3917:
        basket = "22"
    elif vol <= 4133:
        basket = "23"
    elif vol <= 4349:
        basket = "24"
    elif vol <= 4565:
        basket = "25"
    elif vol <= 4877:
        basket = "26"
    elif vol <= 5189:
        basket = "27"
    elif vol <= 5501:
        basket = "28"
    elif vol <= 5813:
        basket = "29"
    elif vol <= 6125:
        basket = "30"
    elif vol <= 6437:
        basket = "31"
    elif vol <= 6749:
        basket = "32"
    elif vol <= 7061:
        basket = "33"
    elif vol <= 7373:
        basket = "34"
    elif vol <= 7685:
        basket = "35"
    elif vol <= 7997:
        basket = "36"
    elif vol <= 8309:
        basket = "37"
    else:
        basket = "38"

    return basket, str(vol), part


def init_driver(headless=None):
    if headless is None:
        headless = config.HEADLESS_MODE
    logger.info("Инициализация драйвера для WB...")
    proxy_config = None
    if config.USE_PROXY:
        pm = ProxyManager(str(config.PROXY_FILE))
        if pm.has_proxies():
            proxy = pm.get_first()
            if proxy:
                proxy_config = pm.format_for_selenium_wire(proxy)
    tunnel = UCWithTunnel(proxy_config=proxy_config)
    driver = tunnel.create_driver(headless=headless, user_data_dir=str(config.CHROME_PROFILE_WB))
    driver.set_page_load_timeout(config.PAGE_LOAD_TIMEOUT)
    driver.implicitly_wait(config.IMPLICIT_WAIT)
    return driver, tunnel


def get_cookies_from_wb(driver, max_attempts=3) -> Optional[dict]:
    """Загружает главную страницу и возвращает словарь кук, с повторными попытками"""
    for attempt in range(1, max_attempts + 1):
        try:
            logger.info(f"Загружаем wildberries.ru (попытка {attempt})...")
            driver.get("https://www.wildberries.ru/")
            # Ждём загрузки body
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            # Имитация действий пользователя, чтобы стимулировать установку кук
            random_delay(2, 3)
            # Прокрутка страницы
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/3);")
            time.sleep(1)
            driver.execute_script("window.scrollTo(0, 0);")
            
            # Небольшое движение мыши
            try:
                actions = ActionChains(driver)
                actions.move_by_offset(100, 100).perform()
                actions.move_by_offset(-50, -50).perform()
            except:
                pass
            
            time.sleep(2)
            
            cookies = {c['name']: c['value'] for c in driver.get_cookies()}
            logger.info(f"Получено кук: {len(cookies)}")
            
            if cookies:
                return cookies
            else:
                logger.warning(f"Куки не получены, повтор через 2 сек...")
                time.sleep(2)
        except Exception as e:
            logger.error(f"Ошибка при загрузке WB: {e}")
            if attempt == max_attempts:
                return None
            time.sleep(3)
    return None


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((requests.RequestException, json.JSONDecodeError)),
    reraise=True
)
def fetch_wb_card(nm_id: str, cookies: dict) -> dict:
    basket, vol, part = get_sku_url_data(nm_id)
    url = f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/info/ru/card.json"
    headers = {'User-Agent': random.choice(config.USER_AGENTS)}
    resp = requests.get(url, headers=headers, cookies=cookies, timeout=15)
    resp.raise_for_status()
    return resp.json()


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((requests.RequestException, json.JSONDecodeError)),
    reraise=True
)
def fetch_wb_detail(nm_id: str, cookies: dict) -> dict:
    url = f"https://www.wildberries.ru/__internal/u-card/cards/v4/detail?appType=1&curr=rub&dest=-1257786&spp=30&hide_vflags=4294967296&hide_dtype=9;11&ab_testing=false&lang=ru&nm={nm_id}"
    headers = {
        'User-Agent': random.choice(config.USER_AGENTS),
        'X-Requested-With': 'XMLHttpRequest',
    }
    resp = requests.get(url, headers=headers, cookies=cookies, timeout=15)
    resp.raise_for_status()
    return resp.json()


def parse_wb_product(nm_id: str, cookies: dict) -> dict:
    result = {
        "price": "",
        "rating_reviews": "",
        "display_type": "",
        "battery_type": "",
        "promo": "",
        "has_promo": False,
        "seller": "",
        "error": None
    }

    try:
        card = fetch_wb_card(nm_id, cookies)
        detail = fetch_wb_detail(nm_id, cookies)
    except Exception as e:
        result["error"] = str(e)[:200]
        return result

    products = detail.get("products", [])
    if products:
        p = products[0]
        sizes = p.get("sizes", [])
        for size in sizes:
            price_info = size.get("price", {})
            product_price = price_info.get("product")
            if product_price:
                result["price"] = str(int(product_price / 100))
                break

        rating = p.get("rating")
        feedbacks = p.get("feedbacks") or p.get("nmFeedbacks", 0)
        if rating and feedbacks:
            result["rating_reviews"] = f"{rating} / {feedbacks}"
        elif rating:
            result["rating_reviews"] = str(rating)

        for size in sizes:
            price_info = size.get("price", {})
            basic = price_info.get("basic")
            product_price = price_info.get("product")
            if basic and product_price and basic > product_price:
                result["has_promo"] = True
                break
        if not result["has_promo"]:
            promo_text = p.get("promoTextCard") or p.get("promoTextCat")
            if promo_text:
                result["has_promo"] = True

        result["seller"] = p.get("brand", "")

    options = card.get("options", [])
    for opt in options:
        if isinstance(opt, dict):
            name = opt.get("name", "").lower()
            value = opt.get("value", "")
            if not value:
                continue
            if not result["display_type"] and ("дисплей" in name or "экран" in name):
                result["display_type"] = str(value)
            if not result["battery_type"] and ("аккумулятор" in name or "батарея" in name):
                result["battery_type"] = str(value)

    return result


def save_to_local_files(updates, promo_cells, sheet):
    """Сохраняет данные в локальные CSV и XLSX файлы при ошибке записи в Google Sheets."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"wb_results_{timestamp}.csv"
    xlsx_filename = f"wb_results_{timestamp}.xlsx"

    # Собираем данные в словарь по строкам
    rows_data = {}
    for row, col, val in updates:
        col_letter = col_index_to_letter(col)
        if row not in rows_data:
            rows_data[row] = {}
        rows_data[row][col_letter] = val

    promo_info = {}
    for row, col, color in promo_cells:
        promo_info[row] = color

    # Сохраняем CSV
    try:
        with open(csv_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Row', 'Column', 'Value', 'Promo'])
            for row, cols in rows_data.items():
                for col, val in cols.items():
                    promo = 'Yes' if row in promo_info else ''
                    writer.writerow([row, col, val, promo])
        logger.info(f"✅ Данные сохранены в CSV: {csv_filename}")
    except Exception as e:
        logger.error(f"Ошибка сохранения CSV: {e}")

    # Сохраняем XLSX
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "WB Results"
        ws.append(['Row', 'Column', 'Value', 'Promo'])
        for row, cols in rows_data.items():
            for col, val in cols.items():
                promo = 'Yes' if row in promo_info else ''
                ws.append([row, col, val, promo])
        wb.save(xlsx_filename)
        logger.info(f"✅ Данные сохранены в XLSX: {xlsx_filename}")
    except Exception as e:
        logger.error(f"Ошибка сохранения XLSX: {e}")


def main():
    logger.info("=" * 70)
    logger.info("🚀 ЗАПУСК WILDBERRIES PARSER (STANDALONE)")
    logger.info("=" * 70)

    _, sheet = get_sheet_client()
    driver, tunnel = init_driver()
    all_updates = []
    promo_cells = []

    try:
        cookies = get_cookies_from_wb(driver)
        if not cookies:
            logger.error("Не удалось получить куки после нескольких попыток")
            driver.quit()
            tunnel.close()
            return

        all_values = sheet.get_all_values()
        if len(all_values) < 2:
            logger.warning("Нет данных")
            driver.quit()
            tunnel.close()
            return

        col_wb = col_letter_to_index(config.WB_SKU_COLUMN)        # K
        col_ozon = col_letter_to_index(config.OZON_INPUT_COLUMN)  # K (по конфигу)
        col_price = col_letter_to_index(config.WB_PRICE_COLUMN)   # M
        col_rating = col_letter_to_index(config.WB_RATING_REVIEWS_COLUMN)  # Y
        col_display_battery = col_letter_to_index(config.WB_DISPLAY_BATTERY_COLUMN)  # H
        col_promo = col_letter_to_index(config.WB_PROMO_COLUMN)   # AD
        col_seller = col_letter_to_index(config.WB_SELLER_COLUMN) # I

        wb_tasks = []

        for row_idx in range(2, len(all_values) + 1):
            row = all_values[row_idx - 1]

            if len(row) >= col_wb:
                val = row[col_wb - 1].strip()
                if val and detect_link_type(val) == 'wb':
                    wb_tasks.append((row_idx, val))

            if len(row) >= col_ozon:
                val = row[col_ozon - 1].strip()
                if val and detect_link_type(val) == 'wb':
                    wb_tasks.append((row_idx, val))

        wb_tasks = list(dict.fromkeys(wb_tasks))

        if not wb_tasks:
            logger.warning("Нет WB ссылок")
            driver.quit()
            tunnel.close()
            return

        total = len(wb_tasks)
        parsed = 0
        errors = 0

        logger.info(f"Найдено WB ссылок: {total}")

        pbar = tqdm(total=total, desc="Парсинг WB", unit="товаров", colour="green")

        for row_idx, raw in wb_tasks:
            nm_id = extract_nm_id(raw) or (raw if raw.isdigit() else None)
            if not nm_id:
                errors += 1
                all_updates.extend([
                    (row_idx, col_price, "INVALID"),
                    (row_idx, col_rating, ""),
                    (row_idx, col_display_battery, ""),
                    (row_idx, col_promo, ""),
                    (row_idx, col_seller, "")
                ])
                pbar.update(1)
                continue

            pbar.set_postfix_str(f"{raw[:20]}...")
            data = parse_wb_product(nm_id, cookies)

            if data.get("error"):
                errors += 1
                err = data["error"][:20]
                all_updates.append((row_idx, col_price, f"ERR: {err}"))
                all_updates.append((row_idx, col_rating, ""))
                all_updates.append((row_idx, col_display_battery, ""))
                all_updates.append((row_idx, col_promo, ""))
                all_updates.append((row_idx, col_seller, ""))
            else:
                parsed += 1
                display = data.get("display_type") or data.get("battery_type") or ""
                all_updates.append((row_idx, col_price, data.get("price", "")))
                all_updates.append((row_idx, col_rating, data.get("rating_reviews", "")))
                all_updates.append((row_idx, col_display_battery, display))
                all_updates.append((row_idx, col_promo, data.get("promo", "")))
                all_updates.append((row_idx, col_seller, data.get("seller", "")))
                if data.get("has_promo"):
                    promo_cells.append((row_idx, col_promo, "#b7e1cd"))

            pbar.update(1)

        pbar.close()

        if all_updates:
            logger.info(f"Запись {len(all_updates)} обновлений в Google Sheets...")
            try:
                safe_batch_update(sheet, all_updates)
                if promo_cells:
                    logger.info(f"Заливка {len(promo_cells)} ячеек цветом")
                    apply_cell_colors(sheet, promo_cells)
            except Exception as e:
                logger.error(f"Не удалось записать в Google Sheets: {e}")
                # Сохраняем результаты локально
                save_to_local_files(all_updates, promo_cells, sheet)
        else:
            logger.info("Нет данных для записи.")

        logger.info(f"Готово! Обработано: {parsed}, Ошибок: {errors}")

    except KeyboardInterrupt:
        logger.warning("\nПрервано пользователем")
        if all_updates:
            safe_batch_update(sheet, all_updates)
            if promo_cells:
                apply_cell_colors(sheet, promo_cells)
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        if all_updates:
            safe_batch_update(sheet, all_updates)
    finally:
        driver.quit()
        tunnel.close()


if __name__ == "__main__":
    main()
