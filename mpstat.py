import sys
import time
import random
from pathlib import Path
from typing import List, Tuple, Optional

import pandas as pd
from tqdm import tqdm

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.action_chains import ActionChains

import config
from config import setup_logging
from uc_wire_tunnel import UCWithTunnel
from proxy_manager import ProxyManager
from gsheets import safe_batch_update, col_letter_to_index, get_sheet_client

logger = setup_logging("mpstats_parser")


def random_delay(min_sec=None, max_sec=None):
    min_sec = min_sec or config.RANDOM_DELAY_MIN
    max_sec = max_sec or config.RANDOM_DELAY_MAX
    time.sleep(random.uniform(min_sec, max_sec))


def human_like_actions(driver):
    """Имитация действий человека для обхода бот-детекции"""
    try:
        # Прокрутка страницы
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight/3);")
        time.sleep(0.5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
        time.sleep(0.5)
        driver.execute_script("window.scrollTo(0, 0);")
        
        # Движение мыши
        actions = ActionChains(driver)
        actions.move_by_offset(100, 100).perform()
        time.sleep(0.3)
        actions.move_by_offset(-50, -50).perform()
    except:
        pass


def setup_browser(headless=False):
    logger.info("Инициализация UC драйвера для MPStats...")
    proxy_config = None
    if config.USE_PROXY:
        pm = ProxyManager(str(config.PROXY_FILE))
        if pm.has_proxies():
            proxy = pm.get_first()
            if proxy:
                proxy_config = pm.format_for_selenium_wire(proxy)
                logger.info(f"Использование прокси: {proxy['host']}:{proxy['port']}")

    tunnel = UCWithTunnel(proxy_config=proxy_config)
    driver = tunnel.create_driver(
        headless=headless,
        user_data_dir=str(config.CHROME_PROFILE_MPSTATS)
    )

    # Настройка папки загрузок
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": str(config.DOWNLOAD_DIR)
    })

    driver.set_page_load_timeout(60)
    driver.implicitly_wait(config.IMPLICIT_WAIT)
    logger.info("UC драйвер успешно инициализирован")
    return driver, tunnel


def check_and_login_mpstats(driver) -> bool:
    try:
        logger.info("Проверка авторизации на MPStats...")
        driver.get("https://mpstats.io/login")
        time.sleep(3)

        if not config.MPSTATS_EMAIL or not config.MPSTATS_PASSWORD:
            logger.error("MPSTATS_EMAIL/PASSWORD не заданы")
            return False

        wait = WebDriverWait(driver, 20)

        # Ищем поля ввода
        email_input = None
        for sel in [
            (By.NAME, "mpstats-login-form-name"),
            (By.CSS_SELECTOR, "input[type='email']"),
            (By.CSS_SELECTOR, "input[placeholder*='Email']"),
        ]:
            try:
                email_input = wait.until(EC.element_to_be_clickable(sel))
                break
            except TimeoutException:
                continue

        password_input = None
        for sel in [
            (By.NAME, "mpstats-login-form-password"),
            (By.CSS_SELECTOR, "input[type='password']"),
        ]:
            try:
                password_input = wait.until(EC.element_to_be_clickable(sel))
                break
            except TimeoutException:
                continue

        if not email_input or not password_input:
            if "/login" in driver.current_url:
                logger.error("Форма логина не найдена")
                driver.save_screenshot("login_form_not_found.png")
                return False
            logger.info("Уже авторизованы")
            return True

        logger.info("Выполняю вход...")
        email_input.clear()
        email_input.send_keys(config.MPSTATS_EMAIL)
        time.sleep(0.5)
        password_input.clear()
        password_input.send_keys(config.MPSTATS_PASSWORD)
        time.sleep(0.5)

        # Имитация человеческого поведения перед кликом
        human_like_actions(driver)

        # Пытаемся отправить форму
        try:
            submit = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            submit.click()
        except Exception:
            password_input.send_keys(Keys.RETURN)

        # Ждём до 30 секунд и проверяем наличие элементов кабинета
        time.sleep(5)

        # Проверяем элементы, характерные для личного кабинета MPStats
        dashboard_selectors = [
            (By.CSS_SELECTOR, "[href*='/profile']"),
            (By.CSS_SELECTOR, ".user-menu"),
            (By.XPATH, "//*[contains(text(), 'Динамика товаров')]"),
            (By.XPATH, "//*[contains(text(), 'Финансовая сводка')]"),
            (By.CSS_SELECTOR, ".ag-root"),  # таблица
        ]

        for by, selector in dashboard_selectors:
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((by, selector)))
                logger.info(f"Элемент кабинета найден: {selector}")
                logger.info("Авторизация выполнена успешно")
                return True
            except TimeoutException:
                continue

        # Если элементы не найдены, но URL изменился – возможно, всё равно вход выполнен
        if "/login" not in driver.current_url:
            logger.info("URL изменился, хотя элементы кабинета не найдены – считаем вход успешным")
            return True

        # Если ничего не помогло – ошибка
        logger.error("Авторизация не выполнена: не найдены элементы кабинета и URL не изменился")
        driver.save_screenshot("login_error.png")
        return False

    except Exception as e:
        logger.error(f"Ошибка авторизации: {e}")
        driver.save_screenshot("login_exception.png")
        return False


def clear_all_filters(driver):
    try:
        logger.debug("Очистка фильтров...")
        inp = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.ag-input-field-input[aria-label*='Название']"))
        )
        driver.execute_script("arguments[0].value = '';", inp)
        time.sleep(0.5)
        logger.debug("Фильтр очищен")
    except Exception as e:
        logger.debug(f"Не удалось очистить фильтры: {e}")


def fill_name_filter(driver, name):
    try:
        logger.debug(f"Заполнение фильтра: {name}")
        inp = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.ag-input-field-input[aria-label*='Название']"))
        )
        inp.clear()
        inp.send_keys(name)
        inp.send_keys(Keys.ENTER)
        logger.debug("Фильтр применён")
        time.sleep(2)
    except Exception as e:
        logger.warning(f"Ошибка при заполнении фильтра: {e}")


def click_download_csv(driver):
    logger.info("📥 Скачивание CSV...")
    wait = WebDriverWait(driver, 30)

    try:
        download_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[.//div[normalize-space()='Скачать']]")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", download_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", download_btn)
        logger.info("✅ Нажата кнопка 'Скачать'")

        time.sleep(2)

        csv_item = None
        selectors = [
            (By.XPATH, "//span[starts-with(., 'Скачать только включенные колонки')]"),
            (By.XPATH, "//*[contains(., 'Скачать только включенные колонки')]"),
        ]
        for by, sel in selectors:
            try:
                csv_item = wait.until(EC.element_to_be_clickable((by, sel)))
                logger.info("✅ Пункт меню найден")
                break
            except TimeoutException:
                continue

        if not csv_item:
            logger.error("❌ Пункт меню не найден")
            menu_items = driver.find_elements(By.XPATH, "//*[contains(text(), 'Скачать')]")
            logger.info(f"Найдено элементов с 'Скачать': {len(menu_items)}")
            for item in menu_items[:5]:
                logger.info(f"  Текст: {item.text[:100]}")
            raise TimeoutException("Не найден пункт меню")

        driver.execute_script("arguments[0].click();", csv_item)
        logger.info("✅ Загрузка начата")
        time.sleep(2)

    except Exception as e:
        logger.error(f"Ошибка при скачивании: {e}")
        raise


def wait_new_file(timeout=60):
    folder = Path(config.DOWNLOAD_DIR)
    end = time.time() + timeout
    last_mtime = 0
    last_file = None

    while time.time() < end:
        files = list(folder.glob("*.csv"))
        if files:
            newest = max(files, key=lambda f: f.stat().st_mtime)
            if newest.stat().st_mtime != last_mtime:
                last_file = newest
                last_mtime = newest.stat().st_mtime
                time.sleep(2)
                return str(newest)
        time.sleep(1)
    return None


def parse_csv(file_path):
    logger.debug(f"Парсинг CSV: {file_path}")
    separators = [';', ',']
    encodings = ['utf-8', 'cp1251']

    df = None
    for enc in encodings:
        for sep in separators:
            try:
                df = pd.read_csv(file_path, encoding=enc, sep=sep, on_bad_lines='skip')
                if len(df.columns) > 5:
                    logger.debug(f"✅ Прочитан с encoding={enc}, sep='{sep}'")
                    break
            except Exception:
                continue
        if df is not None and len(df.columns) > 5:
            break

    if df is None or len(df.columns) <= 5:
        logger.error("Не удалось прочитать CSV")
        return []

    price_col = None
    sales_col = None
    for col in df.columns:
        low = col.lower()
        if not price_col and ('price' in low or 'цена' in low):
            price_col = col
        if not sales_col and ('sales' in low or 'продаж' in low):
            sales_col = col

    if not price_col or not sales_col:
        logger.error(f"Колонки не найдены: {list(df.columns)}")
        return []

    items = []
    for _, r in df.iterrows():
        try:
            price_str = str(r[price_col]).strip()
            if not price_str or price_str.lower() in ('nan', 'none', ''):
                continue
            price = float(price_str.replace(",", ".").replace(" ", ""))

            sales_str = str(r[sales_col]).strip()
            sales = 0 if sales_str.lower() in ('nan', 'none', '') else int(float(sales_str.replace(",", ".").replace(" ", "")))

            if sales > 0 and price > 0:
                items.append({"price": price, "sales": sales})
        except Exception:
            continue

    logger.debug(f"Найдено {len(items)} товаров")
    return items


def calculate(items):
    if not items:
        return "0", "0 / 0"
    avg = int(sum(i["price"] for i in items[:10]) / min(10, len(items)))
    sales = sum(i["sales"] for i in items)
    return str(avg), f"{sales} / {len(items)}"


def get_all_filled_rows(sheet, column_letter: str) -> List[Tuple[int, str]]:
    col_idx = col_letter_to_index(column_letter) - 1
    values = sheet.get_all_values()
    rows = []
    for i, row in enumerate(values, start=1):
        if i == 1:
            continue
        if len(row) > col_idx and row[col_idx].strip():
            rows.append((i, row[col_idx].strip()))
    return rows


def get_name_filter(sheet, row: int) -> Optional[str]:
    try:
        val = sheet.cell(row, col_letter_to_index(config.MPSTATS_FILTER_NAME_COLUMN)).value
        return val.strip() if val else None
    except:
        return None


def wait_for_table(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".ag-root, .ag-grid, table"))
        )
        logger.debug("Таблица загружена")
        return True
    except TimeoutException:
        logger.warning("Таблица не загрузилась за отведённое время")
        return False


def main():
    logger.info("=" * 70)
    logger.info("🚀 ЗАПУСК MPSTATS PARSER (STANDALONE)")
    logger.info("=" * 70)

    # Подключаемся к Google Sheets с повторными попытками
    max_retries = 3
    sheet = None
    for attempt in range(1, max_retries + 1):
        try:
            _, sheet = get_sheet_client()
            logger.info("Подключение к Google Sheets успешно")
            break
        except Exception as e:
            logger.error(f"Ошибка подключения к Google Sheets (попытка {attempt}): {e}")
            if attempt == max_retries:
                logger.critical("Не удалось подключиться к Google Sheets после нескольких попыток. Выход.")
                sys.exit(1)
            time.sleep(5)

    driver, tunnel = setup_browser(headless=False)
    all_updates = []

    try:
        if not check_and_login_mpstats(driver):
            logger.error("Не удалось авторизоваться")
            driver.quit()
            tunnel.close()
            sys.exit(1)

        rows = get_all_filled_rows(sheet, config.MPSTATS_LINK_COLUMN)
        if not rows:
            logger.warning("Нет данных для обработки")
            driver.quit()
            tunnel.close()
            sys.exit(0)

        total = len(rows)
        parsed = 0
        errors = 0
        logger.info(f"Найдено фильтров: {total}")

        col_price = col_letter_to_index(config.MPSTATS_AVG_PRICE_COLUMN)
        col_sales = col_letter_to_index(config.MPSTATS_SALES_COLUMN)

        pbar = tqdm(total=total, desc="Парсинг MPStats", unit="фильтров", colour="yellow")

        for row_num, link_value in rows:
            filter_name = get_name_filter(sheet, row_num)
            display = filter_name or link_value or ""
            pbar.set_postfix_str(f"Фильтр: {display[:20]}...")

            try:
                if link_value and link_value.startswith(("http://", "https://")):
                    logger.info(f"Переход по ссылке: {link_value}")
                    driver.get(link_value)
                    time.sleep(5)
                    if not wait_for_table(driver):
                        raise Exception("Таблица не загрузилась после перехода по ссылке")
                else:
                    logger.warning(f"Пропускаем строку {row_num}: нет ссылки для перехода")
                    errors += 1
                    all_updates.append((row_num, col_price, "Нет ссылки"))
                    all_updates.append((row_num, col_sales, "Нет ссылки"))
                    pbar.update(1)
                    continue

                if filter_name:
                    clear_all_filters(driver)
                    fill_name_filter(driver, filter_name)
                    time.sleep(3)

                click_download_csv(driver)
                file_path = wait_new_file(timeout=30)

                if not file_path:
                    errors += 1
                    all_updates.append((row_num, col_price, "Ошибка скачивания"))
                    all_updates.append((row_num, col_sales, "Ошибка скачивания"))
                    pbar.update(1)
                    continue

                items = parse_csv(file_path)
                avg_price, sales_str = calculate(items)

                parsed += 1
                all_updates.append((row_num, col_price, avg_price))
                all_updates.append((row_num, col_sales, sales_str))

            except Exception as e:
                logger.error(f"Ошибка обработки строки {row_num}: {e}")
                errors += 1
                all_updates.append((row_num, col_price, "Ошибка"))
                all_updates.append((row_num, col_sales, "Ошибка"))

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
