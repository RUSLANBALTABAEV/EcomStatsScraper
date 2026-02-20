"""
Централизованная конфигурация для парсеров zakaz2
Все настройки загружаются из .env файла для безопасности
"""

import os
import logging
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).parent.absolute()
DOWNLOAD_DIR = BASE_DIR / "downloads"
CREDENTIALS_FILE = BASE_DIR / os.getenv("CREDENTIALS_FILE", "credentials.json")

# Профили Chrome для разных парсеров (чтобы избежать конфликтов)
CHROME_PROFILE_MPSTATS = BASE_DIR / "chrome_profile_mpstats"
CHROME_PROFILE_WB = BASE_DIR / "chrome_profile_wb"
CHROME_PROFILE_OZON = BASE_DIR / "chrome_profile_ozon"

# Google Sheets
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1MVT3aRaquEqTT2u7kWJpIHAnnZxIEcLVdO1x1Oq1vqc")
SHEET_GID = int(os.getenv("SHEET_GID", "640866547"))
GOOGLE_SHEETS_SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# MPStats
MPSTATS_EMAIL = os.getenv("MPSTATS_EMAIL", "")
MPSTATS_PASSWORD = os.getenv("MPSTATS_PASSWORD", "")
MPSTATS_API_TOKEN = os.getenv("MPSTATS_API_TOKEN", "")

# Колонки Wildberries
WB_SKU_COLUMN = "K"
WB_LINK_COLUMN = "K"
WB_DISPLAY_BATTERY_COLUMN = "H"
WB_PRICE_COLUMN = "M"
WB_RATING_REVIEWS_COLUMN = "Y"
WB_PROMO_COLUMN = "AD"
WB_SELLER_COLUMN = "I"

# Колонки Ozon
OZON_INPUT_COLUMN = "K"
OZON_PRICE_COLUMN = "V"

# Колонки MPStats
MPSTATS_LINK_COLUMN = "AG"
MPSTATS_FILTER_NAME_COLUMN = "AH"
MPSTATS_AVG_PRICE_COLUMN = "L"
MPSTATS_SALES_COLUMN = "Z"

# Прокси
PROXY_FILE = BASE_DIR / os.getenv("PROXY_FILE", "proxies.txt")
USE_PROXY = os.getenv("USE_PROXY", "False").lower() == "true"

# Selenium/UC
HEADLESS_MODE = os.getenv("HEADLESS_MODE", "False").lower() == "true"
PAGE_LOAD_TIMEOUT = int(os.getenv("PAGE_LOAD_TIMEOUT", "30"))
IMPLICIT_WAIT = int(os.getenv("IMPLICIT_WAIT", "10"))

# User Agents
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
]

# Задержки
DELAY_BETWEEN_SCRIPTS = int(os.getenv("DELAY_BETWEEN_SCRIPTS", "10"))
RANDOM_DELAY_MIN = float(os.getenv("RANDOM_DELAY_MIN", "0.4"))
RANDOM_DELAY_MAX = float(os.getenv("RANDOM_DELAY_MAX", "1.2"))

# Логирование
LOG_FILE = BASE_DIR / "parser.log"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# Автосоздание директорий
DOWNLOAD_DIR.mkdir(exist_ok=True)
CHROME_PROFILE_MPSTATS.mkdir(exist_ok=True)
CHROME_PROFILE_WB.mkdir(exist_ok=True)
CHROME_PROFILE_OZON.mkdir(exist_ok=True)


def setup_logging(script_name: str = "zakaz2"):
    """Настраивает логирование для скрипта"""
    logger = logging.getLogger(script_name)
    logger.setLevel(getattr(logging, LOG_LEVEL))
    logger.handlers.clear()

    formatter = logging.Formatter(LOG_FORMAT, LOG_DATE_FORMAT)

    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setLevel(getattr(logging, LOG_LEVEL))
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, LOG_LEVEL))
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


def validate_config():
    """Проверяет необходимые настройки"""
    errors = []
    if not CREDENTIALS_FILE.exists():
        errors.append(f"Файл {CREDENTIALS_FILE.name} не найден!")
    if not MPSTATS_EMAIL:
        errors.append("MPSTATS_EMAIL не задан")
    if not MPSTATS_PASSWORD:
        errors.append("MPSTATS_PASSWORD не задан")
    if not SPREADSHEET_ID or SPREADSHEET_ID == "your_spreadsheet_id_here":
        errors.append("SPREADSHEET_ID не задан")
    if USE_PROXY and not PROXY_FILE.exists():
        errors.append(f"Прокси включены, но файл {PROXY_FILE.name} не найден")
    return len(errors) == 0, errors


def print_config_info():
    """Выводит информацию о конфигурации"""
    print("\n" + "=" * 70)
    print("КОНФИГУРАЦИЯ ZAKAZ2")
    print("=" * 70)
    print(f"Базовая директория: {BASE_DIR}")
    print(f"Папка загрузок: {DOWNLOAD_DIR}")
    print(f"Профиль MPStats: {CHROME_PROFILE_MPSTATS}")
    print(f"Профиль WB: {CHROME_PROFILE_WB}")
    print(f"Профиль Ozon: {CHROME_PROFILE_OZON}")
    print(f"Credentials: {CREDENTIALS_FILE.name} ({'найден' if CREDENTIALS_FILE.exists() else 'не найден'})")
    print(f"Spreadsheet ID: {SPREADSHEET_ID[:20]}...")
    print(f"MPStats email: {MPSTATS_EMAIL if MPSTATS_EMAIL else 'не задан'}")
    print(f"Прокси: {'включены' if USE_PROXY else 'выключены'}")
    if USE_PROXY:
        print(f"   Файл прокси: {PROXY_FILE.name} ({'найден' if PROXY_FILE.exists() else 'не найден'})")
    print(f"Headless режим: {'включен' if HEADLESS_MODE else 'выключен'}")
    print(f"Лог файл: {LOG_FILE.name}")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    print_config_info()
    is_valid, errors = validate_config()
    if is_valid:
        print("[OK] Конфигурация валидна!")
    else:
        print("[ERROR] Ошибки конфигурации:")
        for error in errors:
            print(f"   {error}")
