"""
UC (Undetected Chrome) с прокси-туннелем от selenium-wire
"""
import atexit
import logging
import signal
import sys
from typing import Optional, Dict

import undetected_chromedriver as uc
from seleniumwire import backend

logger = logging.getLogger(__name__)


def _detect_chrome_major_version() -> Optional[int]:
    """Определяет major-версию Chrome (упрощённо)"""
    try:
        import winreg
        for root in [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]:
            try:
                with winreg.OpenKey(root, r"Software\Google\Chrome\BLBeacon") as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    return int(version.split('.')[0])
            except Exception:
                continue
    except ImportError:
        pass
    return None


def _sanitize_chrome_profile(user_data_dir: str) -> None:
    """Проверяет целостность профиля Chrome (Preferences)"""
    import json
    import time
    from pathlib import Path

    try:
        prefs = Path(user_data_dir) / "Default" / "Preferences"
        if prefs.exists():
            with open(prefs, 'r', encoding='utf-8') as f:
                json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        # Создаём бэкап и новый пустой Preferences
        backup = prefs.with_suffix(f".corrupt.{int(time.time())}")
        try:
            prefs.rename(backup)
        except Exception:
            prefs.unlink(missing_ok=True)
        prefs.parent.mkdir(parents=True, exist_ok=True)
        with open(prefs, 'w', encoding='utf-8') as f:
            json.dump({}, f)
        logger.warning(f"Preferences повреждён, создан новый")
    except Exception:
        pass


class UCWithTunnel:
    def __init__(self, proxy_config: Optional[Dict] = None):
        self.proxy_config = proxy_config or {}
        self.backend = None
        self.local_proxy_address = None
        self.is_active = False
        # Регистрируем закрытие при выходе
        atexit.register(self.close)
        # Обработка Ctrl+C
        signal.signal(signal.SIGINT, self._signal_handler)

    def _signal_handler(self, sig, frame):
        logger.info("Получен сигнал прерывания, закрываю туннель...")
        self.close()
        sys.exit(0)

    def _start_proxy_backend(self):
        if self.is_active:
            return self.local_proxy_address

        logger.info("Запуск прокси-туннеля...")
        self.backend = backend.create(
            addr='127.0.0.1',
            port=0,
            options={
                'proxy': self.proxy_config,
                'disable_encoding': False,
                'verify_ssl': False,
            }
        )
        addr = self.backend.address()
        self.local_proxy_address = f"{addr[0]}:{addr[1]}"
        self.is_active = True
        logger.info(f"Туннель запущен на {self.local_proxy_address}")
        return self.local_proxy_address

    def create_driver(self, headless: bool = False, **uc_kwargs) -> uc.Chrome:
        local_proxy = self._start_proxy_backend()

        options = uc.ChromeOptions()
        options.add_argument(f'--proxy-server=http://{local_proxy}')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        options.add_argument('--log-level=3')
        options.page_load_strategy = 'eager'  # ускоряем загрузку

        prefs = {
            "profile.default_content_setting_values.geolocation": 1,
            "profile.default_content_setting_values.notifications": 2,
        }
        options.add_experimental_option("prefs", prefs)

        if headless:
            options.add_argument('--headless=new')

        user_data_dir = uc_kwargs.get("user_data_dir")
        if user_data_dir:
            _sanitize_chrome_profile(str(user_data_dir))

        if 'version_main' not in uc_kwargs:
            detected = _detect_chrome_major_version()
            if detected:
                uc_kwargs['version_main'] = detected
                logger.info(f"Определена версия Chrome: {detected}")

        driver = uc.Chrome(options=options, **uc_kwargs)

        # Добавляем метод для получения запросов (может пригодиться)
        def get_requests():
            try:
                if self.backend and self.backend.storage:
                    return list(self.backend.storage.iter_requests())
                return []
            except Exception:
                return []

        driver.get_requests = get_requests

        driver.set_page_load_timeout(15)
        driver.implicitly_wait(3)

        if not headless:
            try:
                driver.minimize_window()
            except Exception:
                pass

        logger.info("UC драйвер создан с прокси-туннелем")
        return driver

    def close(self):
        if self.backend and self.is_active:
            try:
                self.backend.shutdown()
                self.is_active = False
                logger.info("Прокси-туннель закрыт")
            except Exception as e:
                logger.warning(f"Ошибка при закрытии туннеля: {e}")

    def __del__(self):
        self.close()
