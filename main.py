import subprocess
import sys
import time
import io
from pathlib import Path
from tqdm import tqdm

import config
from config import setup_logging

# Устанавливаем кодировку stdout для корректного отображения русских символов
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

logger = setup_logging("main")

PYTHON_EXECUTABLE = sys.executable

SCRIPTS = [
    {"name": "MPStats parser", "path": "mpstat.py"},
    {"name": "Wildberries parser", "path": "wb.py"},
    {"name": "Ozon parser", "path": "ozon.py"}
]


def run_script(script_path: str, name: str) -> bool:
    logger.info("\n" + "=" * 70)
    logger.info(f"🚀 ЗАПУСК: {name}")
    logger.info("=" * 70)

    if not Path(script_path).exists():
        logger.error(f"Файл не найден: {script_path}")
        return False

    try:
        start = time.time()
        process = subprocess.run(
            [PYTHON_EXECUTABLE, script_path],
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        duration = time.time() - start
        if process.returncode != 0:
            logger.error(f"{name} завершился с ошибкой (код {process.returncode})")
            return False
        logger.info(f"{name} завершён успешно за {duration:.1f} сек")
        return True
    except Exception as e:
        logger.error(f"Ошибка запуска {name}: {e}")
        return False


def check_dependencies():
    logger.info("\n" + "=" * 70)
    logger.info("🔍 ПРОВЕРКА ЗАВИСИМОСТЕЙ")
    logger.info("=" * 70)

    is_valid, errors = config.validate_config()
    if not is_valid:
        logger.error("Ошибки конфигурации:")
        for e in errors:
            logger.error(f"   {e}")
        return False

    required = ["selenium", "seleniumwire", "undetected_chromedriver", "gspread", "oauth2client", "pandas", "dotenv", "requests", "tenacity", "tqdm"]
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)

    if missing:
        logger.error(f"Не установлены: {', '.join(missing)}")
        logger.info("Установите: pip install -r requirements.txt")
        return False

    logger.info("✅ Все зависимости установлены")
    logger.info("✅ Конфигурация валидна")
    return True


def main():
    logger.info("\n🔥 ОБЩИЙ ЗАПУСК ВСЕХ ПАРСЕРОВ")
    logger.info("=" * 70)

    if not check_dependencies():
        sys.exit(1)

    config.print_config_info()

    results = []
    pbar = tqdm(total=len(SCRIPTS), desc="Общий прогресс", unit="парсер", colour="cyan")

    for idx, script in enumerate(SCRIPTS, 1):
        pbar.set_postfix_str(f"Запуск: {script['name']}")
        success = run_script(script["path"], script["name"])
        results.append((script["name"], success))
        pbar.update(1)
        if script != SCRIPTS[-1]:
            logger.info(f"⏸️ Пауза {config.DELAY_BETWEEN_SCRIPTS} сек...")
            time.sleep(config.DELAY_BETWEEN_SCRIPTS)

    pbar.close()

    logger.info("\n" + "=" * 70)
    logger.info("📊 ИТОГ ЗАПУСКА")
    logger.info("=" * 70)
    for name, ok in results:
        logger.info(f"{'✅' if ok else '❌'} {name}")
    logger.info("\n🏁 ВСЕ СКРИПТЫ ОТРАБОТАЛИ")
    logger.info("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("\nПрервано пользователем")
        sys.exit(130)
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        sys.exit(1)
