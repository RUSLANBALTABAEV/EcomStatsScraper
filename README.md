# EcomStatsScraper — Парсер маркетплейсов

Автоматизированный инструмент для сбора данных с **Wildberries**, **Ozon** и **MPStats** с записью результатов в **Google Sheets**.

---

## Возможности

- **Wildberries** — цена, рейтинг/отзывы, тип дисплея/аккумулятора, наличие акции, продавец
- **Ozon** — текущая цена товара
- **MPStats** — средняя цена и статистика продаж по фильтру
- Обход антибот-защиты через `undetected-chromedriver`
- Поддержка прокси (HTTP/SOCKS)
- Пакетная запись в Google Sheets с обработкой rate-limit
- Цветовая разметка ячеек (акционные товары)
- Локальное сохранение результатов в CSV/XLSX при недоступности Google Sheets

---

## Структура проекта

```
EcomStatsScraper/
├── main.py              # Точка входа — запуск всех парсеров
├── wb.py                # Парсер Wildberries
├── ozon.py              # Парсер Ozon
├── mpstat.py            # Парсер MPStats
├── gsheets.py           # Работа с Google Sheets API
├── config.py            # Централизованная конфигурация
├── uc_wire_tunnel.py    # UC Chrome + прокси-туннель
├── proxy_manager.py     # Менеджер прокси
├── requirements.txt     # Зависимости Python
└── .env                 # Переменные окружения (не в репозитории)
```

---

## Установка

### Требования

- Python 3.9+
- Google Chrome (актуальная версия)
- Аккаунт Google с доступом к Google Sheets API
- Аккаунт MPStats

### Шаги

1. **Клонируйте репозиторий:**

   ```bash
   git clone https://github.com/your-username/zakaz2.git
   cd zakaz2
   ```

2. **Установите зависимости:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Настройте Google Sheets API:**

   - Создайте сервисный аккаунт в [Google Cloud Console](https://console.cloud.google.com/)
   - Включите Google Sheets API и Google Drive API
   - Скачайте JSON-ключ и сохраните как `credentials.json` в корне проекта
   - Дайте сервисному аккаунту доступ к вашей таблице

4. **Создайте файл `.env`:**

   ```env
   # Google Sheets
   SPREADSHEET_ID=your_spreadsheet_id_here
   SHEET_GID=0

   # MPStats
   MPSTATS_EMAIL=your@email.com
   MPSTATS_PASSWORD=yourpassword
   MPSTATS_API_TOKEN=your_token_here

   # Прокси (опционально)
   USE_PROXY=False
   PROXY_FILE=proxies.txt

   # Chrome
   HEADLESS_MODE=False
   PAGE_LOAD_TIMEOUT=30
   IMPLICIT_WAIT=10

   # Задержки
   DELAY_BETWEEN_SCRIPTS=10
   RANDOM_DELAY_MIN=0.4
   RANDOM_DELAY_MAX=1.2

   # Логирование
   LOG_LEVEL=INFO
   ```

---

## Настройка таблицы Google Sheets

Парсеры ожидают следующие колонки:

| Колонка | Содержимое |
|---------|-----------|
| K | Ссылка/артикул WB или Ozon |
| H | Тип дисплея / аккумулятора (WB) |
| I | Продавец (WB) |
| L | Средняя цена (MPStats) |
| M | Цена WB |
| V | Цена Ozon |
| Y | Рейтинг / отзывы (WB) |
| Z | Продажи (MPStats) |
| AD | Акция (WB) |
| AG | Ссылка MPStats |
| AH | Фильтр по названию (MPStats) |

Колонки можно переопределить в `config.py`.

---

## Использование

### Запустить все парсеры последовательно:

```bash
python main.py
```

### Запустить отдельный парсер:

```bash
python wb.py       # Только Wildberries
python ozon.py     # Только Ozon
python mpstat.py   # Только MPStats
```

### Проверить конфигурацию:

```bash
python config.py
```

---

## Прокси

Поместите прокси в файл `proxies.txt` (один на строку) в любом из форматов:

```
http://user:pass@host:port
user:pass@host:port
host:port
host:port:user:pass
```

Включите использование прокси: `USE_PROXY=True` в `.env`.

---

## Логирование

Все события пишутся в `parser.log` и выводятся в консоль. Уровень логирования настраивается через `LOG_LEVEL` в `.env` (DEBUG, INFO, WARNING, ERROR).

---

## Зависимости

| Пакет | Назначение |
|-------|-----------|
| `selenium` | Управление браузером |
| `undetected-chromedriver` | Обход антибот-защиты |
| `selenium-wire` | Перехват/проксирование трафика |
| `gspread` | Google Sheets API |
| `oauth2client` | Авторизация Google |
| `pandas` | Обработка CSV |
| `requests` | HTTP-запросы к API |
| `tenacity` | Retry-логика |
| `tqdm` | Прогресс-бар |
| `python-dotenv` | Загрузка `.env` |

---

## Лицензия

MIT — подробнее в файле [LICENSE](LICENSE).
