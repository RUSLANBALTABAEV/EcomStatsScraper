import time
import random
import re
import json
from typing import Optional, List
import requests
from tqdm import tqdm
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config
from config import setup_logging
from uc_wire_tunnel import UCWithTunnel
from proxy_manager import ProxyManager
from gsheets import safe_batch_update, col_letter_to_index, get_sheet_client

logger = setup_logging("ozon_parser")


def random_pause(min_sec=None, max_sec=None):
    min_sec = min_sec or config.RANDOM_DELAY_MIN
    max_sec = max_sec or config.RANDOM_DELAY_MAX
    time.sleep(random.uniform(min_sec, max_sec))


def build_ozon_url(value: str) -> str:
    value = value.strip()
    if value.startswith("http"):
        return value
    return f"https://www.ozon.ru/product/{value}/"


def detect_link_type(url: str) -> str:
    v = url.strip()
    if v.isdigit():
        return 'ozon'
    if not v.lower().startswith('http'):
        return 'skip'
    if 'ozon.ru' in v.lower() or 'ozon.by' in v.lower():
        return 'ozon'
    if 'wildberries' in v.lower() or 'wb.ru' in v.lower():
        return 'wb'
    return 'skip'


def init_driver(headless=None):
    if headless is None:
        headless = config.HEADLESS_MODE
    logger.info("Инициализация драйвера для Ozon...")
    proxy_config = None
    if config.USE_PROXY:
        pm = ProxyManager(str(config.PROXY_FILE))
        if pm.has_proxies():
            proxy = pm.get_first()
            if proxy:
                proxy_config = pm.format_for_selenium_wire(proxy)
    tunnel = UCWithTunnel(proxy_config=proxy_config)
    driver = tunnel.create_driver(headless=headless, user_data_dir=str(config.CHROME_PROFILE_OZON))
    driver.set_page_load_timeout(config.PAGE_LOAD_TIMEOUT)
    driver.implicitly_wait(config.IMPLICIT_WAIT)
    return driver, tunnel


def get_cookies_from_ozon(driver) -> Optional[dict]:
    try:
        logger.info("Загружаем ozon.ru для получения кук...")
        driver.get("https://www.ozon.ru/")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        random_pause(1, 2)
        cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        logger.info(f"Получено кук: {len(cookies)}")
        return cookies
    except Exception as e:
        logger.error(f"Не удалось получить куки: {e}")
        return None


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((requests.RequestException, json.JSONDecodeError)),
    reraise=True
)
def fetch_ozon_price(article: str, cookies: dict) -> Optional[str]:
    url = f"https://www.ozon.ru/api/composer-api.bx/page/json/v2?url=/product/{article}"
    headers = {
        'User-Agent': random.choice(config.USER_AGENTS),
        'Accept': 'application/json',
        'X-Requested-With': 'XMLHttpRequest',
    }
    resp = requests.get(url, headers=headers, cookies=cookies, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    widget_states = data.get('widgetStates', {})
    for key, value in widget_states.items():
        if key.startswith('webPrice'):
            if isinstance(value, str):
                value = json.loads(value)
            if not value.get('isAvailable'):
                logger.debug("Товар недоступен")
                return None
            price = value.get('cardPrice') or value.get('price')
            if price:
                clean = re.sub(r'[^\d]', '', price)
                return clean
    logger.debug("Цена не найдена в ответе")
    return None


def parse_ozon_price(article: str, cookies: dict) -> Optional[str]:
    try:
        return fetch_ozon_price(article, cookies)
    except Exception as e:
        logger.warning(f"Ошибка получения цены для {article}: {e}")
        return None


def main():
    logger.info("=" * 70)
    logger.info("🚀 ЗАПУСК OZON PARSER (STANDALONE)")
    logger.info("=" * 70)

    _, sheet = get_sheet_client()
    driver, tunnel = init_driver()
    all_updates = []

    try:
        cookies = get_cookies_from_ozon(driver)
        if not cookies:
            logger.error("Не удалось получить куки, выходим")
            driver.quit()
            tunnel.close()
            return

        all_values = sheet.get_all_values()
        if len(all_values) < 2:
            logger.warning("Нет данных")
            driver.quit()
            tunnel.close()
            return

        col_wb = col_letter_to_index(config.WB_SKU_COLUMN)          # K
        col_price = col_letter_to_index(config.OZON_PRICE_COLUMN)   # V

        ozon_tasks = []

        for row_idx in range(2, len(all_values) + 1):
            row = all_values[row_idx - 1]

            if len(row) >= col_wb:
                val = row[col_wb - 1].strip()
                if val and detect_link_type(val) == 'ozon':
                    ozon_tasks.append((row_idx, val))

        if not ozon_tasks:
            logger.warning("Нет Ozon ссылок для обработки")
            driver.quit()
            tunnel.close()
            return

        total = len(ozon_tasks)
        parsed = 0
        errors = 0

        logger.info(f"Найдено Ozon ссылок: {total}")

        pbar = tqdm(total=total, desc="Парсинг Ozon", unit="товаров", colour="blue")

        for row_idx, raw in ozon_tasks:
            article = re.sub(r'.*/product/(\d+).*', r'\1', raw)
            if not article.isdigit():
                article = raw
            pbar.set_postfix_str(f"{raw[:20]}...")

            price = parse_ozon_price(article, cookies)
            if price:
                parsed += 1
            else:
                errors += 1
                price = ""

            all_updates.append((row_idx, col_price, price))
            pbar.update(1)

        pbar.close()

        if all_updates:
            logger.info(f"Запись {len(all_updates)} обновлений...")
            safe_batch_update(sheet, all_updates)

        logger.info(f"Готово! Обработано: {parsed}, Ошибок: {errors}")

    except KeyboardInterrupt:
        logger.warning("\nПрервано пользователем")
        if all_updates:
            safe_batch_update(sheet, all_updates)
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        if all_updates:
            safe_batch_update(sheet, all_updates)
    finally:
        driver.quit()
        tunnel.close()


if __name__ == "__main__":
    main()
