"""
Менеджер для управления прокси серверами
"""
import random
from typing import Optional, Dict
from urllib.parse import urlparse


class ProxyManager:
    def __init__(self, proxy_file: str = "proxies.txt"):
        self.proxy_file = proxy_file
        self.proxies = []
        self._load_proxies()

    def _load_proxies(self):
        try:
            with open(self.proxy_file, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip() and not line.startswith('#')]
                self.proxies = [self._parse_proxy(line) for line in lines]
                self.proxies = [p for p in self.proxies if p]
        except FileNotFoundError:
            self.proxies = []

    def _parse_proxy(self, proxy_str: str) -> Optional[Dict]:
        proxy_str = proxy_str.strip()
        # protocol://user:pass@host:port
        if '://' in proxy_str:
            parsed = urlparse(proxy_str)
            return {
                'protocol': parsed.scheme or 'http',
                'host': parsed.hostname,
                'port': parsed.port or 8080,
                'username': parsed.username,
                'password': parsed.password
            }
        # user:pass@host:port
        if '@' in proxy_str:
            try:
                auth, server = proxy_str.split('@', 1)
                username, password = auth.split(':', 1)
                host, port = server.split(':', 1)
                return {
                    'protocol': 'http',
                    'host': host,
                    'port': int(port),
                    'username': username,
                    'password': password
                }
            except ValueError:
                pass
        # host:port
        parts = proxy_str.split(':')
        if len(parts) == 2:
            return {
                'protocol': 'http',
                'host': parts[0],
                'port': int(parts[1]),
                'username': None,
                'password': None
            }
        # host:port:user:pass
        if len(parts) == 4:
            return {
                'protocol': 'http',
                'host': parts[0],
                'port': int(parts[1]),
                'username': parts[2],
                'password': parts[3]
            }
        return None

    def has_proxies(self) -> bool:
        return len(self.proxies) > 0

    def count(self) -> int:
        return len(self.proxies)

    def get_first(self) -> Optional[Dict]:
        return self.proxies[0] if self.proxies else None

    def get_random(self) -> Optional[Dict]:
        return random.choice(self.proxies) if self.proxies else None

    def format_for_selenium_wire(self, proxy: Dict) -> Dict:
        """Форматирует прокси для selenium-wire"""
        if not proxy:
            return {}
        if proxy['username'] and proxy['password']:
            url = f"{proxy['protocol']}://{proxy['username']}:{proxy['password']}@{proxy['host']}:{proxy['port']}"
        else:
            url = f"{proxy['protocol']}://{proxy['host']}:{proxy['port']}"
        return {
            'http': url,
            'https': url,
            'no_proxy': 'localhost,127.0.0.1'
        }
