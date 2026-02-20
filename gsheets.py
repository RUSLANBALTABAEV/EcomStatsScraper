"""
Общие функции для работы с Google Sheets
"""
import time
import logging
from typing import List, Tuple

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import config

logger = logging.getLogger("gsheets")


def col_letter_to_index(letter: str) -> int:
    result = 0
    for c in letter.upper():
        result = result * 26 + (ord(c) - 64)
    return result


def col_index_to_letter(col: int) -> str:
    result = ""
    col -= 1
    while col >= 0:
        result = chr(col % 26 + ord('A')) + result
        col = col // 26 - 1
    return result


def get_sheet_client(max_retries=3, delay=2):
    for attempt in range(1, max_retries + 1):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                str(config.CREDENTIALS_FILE), config.GOOGLE_SHEETS_SCOPE
            )
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_key(config.SPREADSHEET_ID)
            sheet = spreadsheet.get_worksheet_by_id(config.SHEET_GID)
            logger.info("Подключение к Google Sheets успешно")
            return client, sheet
        except Exception as e:
            logger.warning(f"Ошибка подключения (попытка {attempt}/{max_retries}): {e}")
            if hasattr(e, 'response') and e.response:
                logger.warning(f"Статус: {e.response.status_code}, тело: {e.response.text[:200]}")
            if attempt == max_retries:
                logger.error("Не удалось подключиться к Google Sheets")
                raise
            time.sleep(delay * attempt)


def safe_batch_update(sheet, updates: List[Tuple[int, int, str]], max_retries=3) -> bool:
    if not updates:
        return True

    batch_data = []
    for row, col, val in updates:
        col_letter = col_index_to_letter(col)
        batch_data.append({
            'range': f"{col_letter}{row}",
            'values': [[val]]
        })

    for attempt in range(max_retries):
        try:
            sheet.batch_update(batch_data, value_input_option='USER_ENTERED')
            return True
        except Exception as e:
            if "quota" in str(e).lower() or "rate" in str(e).lower():
                wait = 2 ** attempt
                logger.warning(f"Лимит API, ожидание {wait} сек...")
                time.sleep(wait)
            else:
                logger.error(f"Ошибка batch update: {e}")
                if attempt == max_retries - 1:
                    raise
    return False


def apply_cell_colors(sheet, color_requests: List[Tuple[int, int, str]], max_retries=3) -> bool:
    if not color_requests:
        return True

    for attempt in range(max_retries):
        try:
            requests = []
            for row, col, hex_color in color_requests:
                hex_color = hex_color.lstrip('#')
                r = int(hex_color[0:2], 16) / 255.0
                g = int(hex_color[2:4], 16) / 255.0
                b = int(hex_color[4:6], 16) / 255.0
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": row - 1,
                            "endRowIndex": row,
                            "startColumnIndex": col - 1,
                            "endColumnIndex": col
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": r, "green": g, "blue": b}
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
            sheet.spreadsheet.batch_update({"requests": requests})
            return True
        except Exception as e:
            if "quota" in str(e).lower() or "rate" in str(e).lower():
                wait = 2 ** attempt
                logger.warning(f"Лимит API (форматирование), ожидание {wait} сек...")
                time.sleep(wait)
            else:
                logger.error(f"Ошибка заливки цветом: {e}")
                if attempt == max_retries - 1:
                    return False
    return False
