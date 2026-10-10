"""
iiko KPI Report — «Здрасте» корпоративный стиль
Забирает данные из iikoChain каждый вечер в 23:00.
Показывает всех сотрудников: работавших и нет. Отслеживает открытие/закрытие смены.
"""

import html
import json
import os
import sys
import argparse
from datetime import datetime, date, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

# ─────────────────────────────────────────────
# НАСТРОЙКИ
# ─────────────────────────────────────────────
# Секреты и идентификаторы внешних систем задаются только через окружение.
IIKO_SERVER = os.environ.get("IIKO_SERVER", "").rstrip("/")
API_LOGIN = os.environ.get("IIKO_LOGIN", "")
API_PASSWORD = os.environ.get("IIKO_PASSWORD", "")
ORG_ID = os.environ.get("IIKO_ORG_ID", "")
ORG_NAME = os.environ.get(
    "IIKO_ORG_NAME",
    "Волоколамское шоссе 95/2с8",
).strip()
OLAP_DEPARTMENT_FIELD = os.environ.get(
    "IIKO_OLAP_DEPARTMENT_FIELD",
    "Department.Id",
).strip()
IIKO_CA_BUNDLE = os.environ.get("IIKO_CA_BUNDLE", "")


def _safe_error(exc: Exception) -> str:
    """Описание ошибки без URL и query-параметров, где могут быть токены."""
    status = getattr(getattr(exc, "response", None), "status_code", None)
    suffix = f" (HTTP {status})" if status else ""
    return f"{type(exc).__name__}{suffix}"

# Коды ролей для API-режима (iiko role codes)
BARISTA_ROLES = ["2145124279"]
WAITER_ROLES  = ["2145125755",  "2145125516", "1680747245"]

# ── Роли сотрудников для Excel-режима ──────────────────────────────────────
# ФИО не хранятся в репозитории. Формат переменных: "Фамилия Имя;Фамилия Имя".
def _env_names(variable_name: str) -> list[str]:
    return [
        value.strip()
        for value in os.environ.get(variable_name, "").split(";")
        if value.strip()
    ]


BARISTA_NAMES = _env_names("BARISTA_NAMES")
WAITER_NAMES = _env_names("WAITER_NAMES")
# ───────────────────────────────────────────────────────────────────────────

OUTPUT_DIR = Path(__file__).parent / "reports"
DATA_DIR = Path(__file__).parent / "data"
APP_TIMEZONE = ZoneInfo(os.environ.get("APP_TIMEZONE", "Europe/Moscow"))


def _now() -> datetime:
    return datetime.now(APP_TIMEZONE)


def _today() -> date:
    return _now().date()

# Нормативы времени
OPEN_DEADLINE  = "08:15"   # открытие не позднее
CLOSE_DEADLINE = "22:15"   # закрытие не позднее
# ─────────────────────────────────────────────

KPI = {
    "barista": [
        {"threshold": 1600, "bonus": 10000, "label": "≥ 1600 ₽", "color": "#c8891a"},
        {"threshold": 1350, "bonus": 5000,  "label": "≥ 1350 ₽", "color": "#7A8C5E"},
        {"threshold": 1200, "bonus": 2000,  "label": "≥ 1200 ₽", "color": "#E8A898"},
    ],
    "waiter": [
        {"threshold": 2000, "bonus": 10000, "label": "≥ 2000 ₽", "color": "#c8891a"},
        {"threshold": 1650, "bonus": 5000,  "label": "≥ 1650 ₽", "color": "#7A8C5E"},
        {"threshold": 1500, "bonus": 2000,  "label": "≥ 1500 ₽", "color": "#E8A898"},
    ],
}

MONTH_NAMES = {
    1:"Январь",2:"Февраль",3:"Март",4:"Апрель",5:"Май",6:"Июнь",
    7:"Июль",8:"Август",9:"Сентябрь",10:"Октябрь",11:"Ноябрь",12:"Декабрь"
}

# ─── iiko API ──────────────────────────────────────────────────────────────

def _session():
    """requests-сессия с обязательной проверкой TLS."""
    session = requests.Session()
    session.headers.update({"User-Agent": "Zdraste-KPI/1.0"})
    if IIKO_CA_BUNDLE:
        session.verify = IIKO_CA_BUNDLE
    return session


def get_token():
    import hashlib

    required = {
        "IIKO_SERVER": IIKO_SERVER,
        "IIKO_LOGIN": API_LOGIN,
        "IIKO_PASSWORD": API_PASSWORD,
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        raise RuntimeError(
            "Не заданы обязательные переменные окружения: " + ", ".join(missing)
        )

    pass_hash = hashlib.sha1(API_PASSWORD.encode()).hexdigest()
    url = f"{IIKO_SERVER}/api/auth"
    r = _session().get(url, params={"login": API_LOGIN, "pass": pass_hash}, timeout=60)
    r.raise_for_status()
    token = r.text.strip().strip('"')
    if not token:
        raise RuntimeError("iiko вернул пустой токен")
    print("[OK] Авторизация iiko выполнена")
    return token

def get_organizations(token):
    import xml.etree.ElementTree as ET
    r = _session().get(f"{IIKO_SERVER}/api/corporation/departments",
                       params={"key": token}, timeout=60)
    r.raise_for_status()
    root = ET.fromstring(r.text)
    def local_name(tag):
        return tag.rsplit("}", 1)[-1]

    def child_text(item, field):
        for child in item:
            if local_name(child.tag) == field:
                return (child.text or "").strip()
        return ""

    result = []
    for item in root.iter():
        if local_name(item.tag) != "corporateItemDto":
            continue
        eid = child_text(item, "id")
        name = child_text(item, "name")
        typ = child_text(item, "type")
        deleted = child_text(item, "deleted").casefold() == "true"
        active_text = child_text(item, "active").casefold()
        active = not deleted and active_text != "false"
        if eid:
            result.append({
                "id": eid,
                "name": name,
                "type": typ,
                "active": active,
            })
    return result


def resolve_organization(organizations):
    """Однозначно выбрать активное подразделение; никогда не брать первое."""
    departments = [
        item for item in organizations
        if item.get("active", True)
        and str(item.get("type", "")).upper() == "DEPARTMENT"
    ]
    if ORG_ID:
        matches = [item for item in departments if item.get("id") == ORG_ID]
        selector = f"IIKO_ORG_ID={ORG_ID}"
    elif ORG_NAME:
        expected = " ".join(ORG_NAME.split()).casefold()
        matches = [
            item for item in departments
            if " ".join(str(item.get("name", "")).split()).casefold() == expected
        ]
        selector = f"IIKO_ORG_NAME={ORG_NAME}"
    else:
        raise RuntimeError("Нужно задать IIKO_ORG_ID или IIKO_ORG_NAME")

    if len(matches) != 1:
        raise RuntimeError(
            f"Подразделение по {selector} найдено неоднозначно: {len(matches)}"
        )
    selected = matches[0]
    if ORG_ID and ORG_NAME:
        actual = " ".join(str(selected.get("name", "")).split()).casefold()
        expected = " ".join(ORG_NAME.split()).casefold()
        if actual != expected:
            raise RuntimeError("IIKO_ORG_ID и IIKO_ORG_NAME указывают на разные подразделения")
    return selected


def _find_column_metadata(value, field_name):
    if isinstance(value, dict):
        direct = value.get(field_name)
        if isinstance(direct, dict):
            return direct
        identifiers = (
            value.get("name"),
            value.get("fieldName"),
            value.get("id"),
            value.get("column"),
        )
        if field_name in identifiers:
            return value
        for nested in value.values():
            found = _find_column_metadata(nested, field_name)
            if found is not None:
                return found
    elif isinstance(value, list):
        for nested in value:
            found = _find_column_metadata(nested, field_name)
            if found is not None:
                return found
    return None


def validate_olap_department_filter(token):
    """Проверить поддержку фильтра филиала в живой схеме iiko."""
    r = _session().get(
        f"{IIKO_SERVER}/api/v2/reports/olap/columns",
        params={"key": token, "reportType": "SALES"},
        timeout=60,
    )
    r.raise_for_status()
    metadata = _find_column_metadata(r.json(), OLAP_DEPARTMENT_FIELD)
    if metadata is None:
        raise RuntimeError(
            f"iiko OLAP не содержит поле {OLAP_DEPARTMENT_FIELD}"
        )
    filtering_allowed = metadata.get(
        "filteringAllowed",
        metadata.get("filterAllowed"),
    )
    if filtering_allowed is not True:
        raise RuntimeError(
            f"iiko OLAP не разрешает фильтрацию по {OLAP_DEPARTMENT_FIELD}"
        )
    field_type = str(
        metadata.get("type")
        or metadata.get("fieldType")
        or metadata.get("valueType")
        or metadata.get("columnType")
        or ""
    ).upper()
    if "ID" not in field_type:
        raise RuntimeError(
            f"Неожиданный тип {OLAP_DEPARTMENT_FIELD}: {field_type or 'не указан'}"
        )


def _sales_filters(org_id, date_from, date_to):
    if not org_id:
        raise RuntimeError("Не определён UUID подразделения iiko")
    return {
        "OpenDate.Typed": {
            "filterType": "DateRange",
            "periodType": "CUSTOM",
            "from": date_from[:10],
            "to": date_to[:10],
            "includeLow": True,
            "includeHigh": True,
        },
        OLAP_DEPARTMENT_FIELD: {
            "filterType": "IncludeValues",
            "values": [org_id],
        },
    }

def get_employees(token, org_id):
    import xml.etree.ElementTree as ET
    r = _session().get(f"{IIKO_SERVER}/api/v2/employees",
                       params={"key": token}, timeout=60)
    r.raise_for_status()
    root = ET.fromstring(r.text)
    result = []
    for emp in root.iter("employee"):
        eid       = emp.findtext("id",           "")
        name      = emp.findtext("name",         "")
        role_code = emp.findtext("mainRoleCode", "") or emp.findtext("mainRoleId", "")
        if eid:
            result.append({"id": eid, "name": name, "mainRoleCode": role_code})
    return result

def get_worked_hours(token, org_id, date_from, date_to):
    # EMPLOYEE_WORKED_HOURS не поддерживается этой версией iiko API.
    # Время смены берём из данных продаж (первый/последний чек).
    return None

def get_sales(token, org_id, date_from, date_to):
    payload = {
        "reportType": "SALES",
        "groupByRowFields": ["WaiterName"],
        "aggregateFields":  [
            "DishDiscountSumInt",        # Сумма со скидкой (выручка)
            "UniqOrderId",               # Чеков
            "DishDiscountSumInt.average",# Средняя сумма заказа
        ],
        "filters": _sales_filters(org_id, date_from, date_to),
        "buildSummary": False,
    }
    try:
        r = _session().post(f"{IIKO_SERVER}/api/v2/reports/olap",
                            params={"key": token}, json=payload, timeout=90)
        if r.status_code in (400, 500):
            print(f"[WARN] Продажи: HTTP {r.status_code}")
            return None
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[WARN] Продажи: {_safe_error(e)}"); return None

def get_shift_times(token, org_id, date_from, date_to):
    """Время первого и последнего чека сотрудника за день."""
    # OpenTime/CloseTime — поля группировки, берём их как строки
    # и потом находим min/max по каждому сотруднику
    payload = {
        "reportType": "SALES",
        "groupByRowFields": ["WaiterName", "OpenTime", "CloseTime"],
        "aggregateFields":  ["UniqOrderId"],
        "filters": _sales_filters(org_id, date_from, date_to),
        "buildSummary": False,
    }
    try:
        r = _session().post(f"{IIKO_SERVER}/api/v2/reports/olap",
                            params={"key": token}, json=payload, timeout=90)
        if r.status_code in (400, 500):
            print(f"[WARN] Время смены: HTTP {r.status_code}")
            return None
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[WARN] Время смены: {_safe_error(e)}"); return None

def get_category_sales(token, org_id, date_from, date_to):
    """Продажи по категориям блюд для каждого сотрудника."""
    payload = {
        "reportType": "SALES",
        "groupByRowFields": ["WaiterName", "DishCategory"],
        "aggregateFields":  ["DishDiscountSumInt"],
        "filters": _sales_filters(org_id, date_from, date_to),
        "buildSummary": False,
    }
    try:
        r = _session().post(f"{IIKO_SERVER}/api/v2/reports/olap",
                            params={"key": token}, json=payload, timeout=90)
        if r.status_code in (400, 500):
            return None
        r.raise_for_status()
        # Строим словарь: name → {category: sum}
        result = {}
        for row in r.json().get("data", []):
            raw_name = row.get("WaiterName", "") or ""
            parts = raw_name.split()
            name = " ".join(parts[:-1]) if parts and parts[-1].isdigit() else raw_name
            name = name.strip()
            if not name:
                continue
            cat = row.get("DishCategory", "") or ""
            amt = float(row.get("DishDiscountSumInt", 0) or 0)
            if name not in result:
                result[name] = {}
            result[name][cat] = result[name].get(cat, 0) + amt
        print(f"[OK] Категории блюд: {len(result)} сотрудников")
        return result
    except Exception as e:
        print(f"[WARN] Категории: {_safe_error(e)}"); return None

# ─── Настройки KPI по категориям ────────────────────────────────────────────
DOBY_BAR_TARGET    = 0.08   # 8% от выручки для бариста
DOBY_KITCHEN_TARGET = 0.08  # 8% от выручки для официантов
DESSERTS_TARGET    = 0.05   # 5% от выручки для всех

# ─── Чтение графика из Google Sheets ────────────────────────────────────────

SCHEDULE_SHEET_ID = os.environ.get("SCHEDULE_SHEET_ID", "")
SCHEDULE_SHEET_GID = os.environ.get("SCHEDULE_SHEET_GID", "")

MONTH_SHEET_NAMES = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
}


def _get_schedule_gid(target_date: date) -> str:
    """Получить GID вкладки графика для месяца и года."""
    try:
        month_name = MONTH_SHEET_NAMES.get(target_date.month, "")
        sheet_name = f"{month_name} {target_date.year}" if month_name else ""
        if not sheet_name:
            return SCHEDULE_SHEET_GID
        # Получаем список листов через экспорт метаданных
        url = f"https://docs.google.com/spreadsheets/d/{SCHEDULE_SHEET_ID}/edit"
        r = _session().get(url, timeout=15)
        # Ищем gid для нужного листа в HTML
        import re
        pattern = rf'"name":"{re.escape(sheet_name)}"[^}}]*?"sheetId":(\d+)'
        m = re.search(pattern, r.text)
        if not m:
            # Попробуем другой формат
            pattern2 = rf'gid=(\d+)[^"]*"[^>]*>{re.escape(sheet_name)}'
            m = re.search(pattern2, r.text)
        if m:
            gid = m.group(1)
            print(f"[OK] График вкладка «{sheet_name}» gid={gid}")
            return gid
    except Exception as e:
        print(f"[WARN] Не удалось определить GID листа: {_safe_error(e)}")
    return SCHEDULE_SHEET_GID

def load_schedule(target_date: date) -> dict:
    """
    Читает график из Google Sheets.
    Структура: строка 1 — «Сити Бэй», строка 2 — дни недели,
    строки 3+ — дата в col A (1-июн. и т.д.), сотрудник в col A, смены в колонках дат.
    Возвращает dict: фамилия_lower → {open: "HH:MM", close: "HH:MM"}
    """
    if not PANDAS_OK or not SCHEDULE_SHEET_ID:
        return {}
    try:
        gid = _get_schedule_gid(target_date)
        if not gid:
            print("[WARN] График: не задан SCHEDULE_SHEET_GID")
            return {}
        url = (f"https://docs.google.com/spreadsheets/d/{SCHEDULE_SHEET_ID}"
               f"/export?format=csv&gid={gid}")
        import io
        r = _session().get(url, timeout=30)
        r.raise_for_status()
        content = r.content.decode("utf-8-sig")
        df = pd.read_csv(io.StringIO(content), header=None)

        # Найти колонку для нужной даты
        # Строка 0: "Сити Бэй", "1-июн.", "2-июн." ...
        months_ru = {
            "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "июн": 6,
            "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12
        }
        import re as _re
        target_col = None
        # Печатаем первую строку для диагностики
        header_vals = [str(df.iloc[0, c]) for c in range(min(10, len(df.columns)))]
        print(f"[DEBUG] график заголовки: {header_vals}")
        for col in range(1, len(df.columns)):
            cell = str(df.iloc[0, col]).strip().lower().replace(".", "").replace("\xa0", "")
            m = _re.search(r"(\d{1,2})[^\d]*(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)", cell)
            if m:
                try:
                    day = int(m.group(1))
                    mon = months_ru.get(m.group(2)[:3])
                    if mon and day == target_date.day and mon == target_date.month:
                        target_col = col
                        break
                except Exception:
                    pass

        if target_col is None:
            print(f"[WARN] График: дата {target_date} не найдена в таблице")
            return {}

        # Собираем расписание: имя → смена
        result = {}
        known = set(BARISTA_NAMES + WAITER_NAMES)
        known_last = {n.split()[0].lower(): n for n in known}  # фамилия → полное имя

        for row in range(3, len(df)):
            name_raw = str(df.iloc[row, 0]).strip()
            if not name_raw or name_raw == "nan":
                continue
            shift_raw = str(df.iloc[row, target_col]).strip()
            if not shift_raw or shift_raw in ("nan", "", "-"):
                continue
            # Пропускаем заголовки групп (Бариста, Официант и т.д.)
            if shift_raw.lower() in ("бариста", "официант", "повар", "уборщица",
                                     "су-шеф", "стажер", "замена", "nan"):
                continue
            # Убираем суффиксы типа "см.", "тренинг"
            shift_clean = shift_raw.lower().replace("см.", "").replace("тренинг", "").strip()
            if not shift_clean or "-" not in shift_clean:
                continue

            # Сопоставляем имя по фамилии
            name_parts = name_raw.split()
            last_name = name_parts[0].lower() if name_parts else ""
            full_name = known_last.get(last_name)
            if not full_name:
                # Попробуем по второму слову (если «Виктория Милосердина»)
                if len(name_parts) > 1:
                    last_name2 = name_parts[1].lower()
                    full_name = known_last.get(last_name2)
            if not full_name:
                continue

            # Парсим смену "8-22" → open="08:00", close="22:00"
            try:
                parts = shift_clean.split("-")
                open_h  = int(parts[0].strip())
                close_h = int(parts[1].strip())
                result[full_name] = {
                    "open":  f"{open_h:02d}:00",
                    "close": f"{close_h:02d}:00",
                }
            except Exception:
                pass

        print(f"[OK] График: {len(result)} сотрудников на {target_date}")
        return result
    except Exception as e:
        print(f"[WARN] График: {_safe_error(e)}")
        return {}

GSHEET_ID = os.environ.get("ATTENDANCE_SHEET_ID", "")
GSHEET_GID = os.environ.get("ATTENDANCE_SHEET_GID", "")

def load_attendance(target_date: date) -> dict:
    """
    Читает табель из Google Sheets (доступ по ссылке — Читатель).
    Возвращает dict: name → {open_time, close_time, hours}
    только для указанной даты.
    """
    if not PANDAS_OK:
        print("[WARN] Табель: pandas не установлен")
        return {}
    if not GSHEET_ID or not GSHEET_GID:
        return {}
    try:
        url = (f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}"
               f"/export?format=csv&gid={GSHEET_GID}")
        import io
        r = _session().get(url, timeout=30)
        r.raise_for_status()
        content = r.content.decode("utf-8-sig")
        df = pd.read_csv(io.StringIO(content))
        df.columns = [str(c).strip() for c in df.columns]

        result = {}
        for _, row in df.iterrows():
            raw_date = row.get("Дата") or row.get(df.columns[0])
            if not raw_date or str(raw_date).strip() in ("", "nan"):
                continue
            try:
                if hasattr(raw_date, "date"):
                    row_date = raw_date.date()
                else:
                    from dateutil.parser import parse
                    row_date = parse(str(raw_date), dayfirst=True).date()
            except Exception:
                continue
            if row_date != target_date:
                continue

            name = str(row.get("Сотрудник") or row.get(df.columns[1]) or "").strip()
            if not name or name == "nan":
                continue

            def _t(val):
                if val is None or str(val).strip() in ("", "nan"):
                    return None
                if hasattr(val, "strftime"):
                    return val.strftime("%H:%M")
                s = str(val).strip()
                return s[:5] if len(s) >= 5 else s

            open_t  = _t(row.get("Приход") or row.get(df.columns[2]))
            close_t = _t(row.get("Уход")   or row.get(df.columns[3]))

            hours = 0.0
            raw_h = row.get("Часов") or row.get(df.columns[4] if len(df.columns) > 4 else "")
            try:
                if raw_h and str(raw_h).strip() not in ("", "nan"):
                    hours = float(str(raw_h).replace(",", "."))
            except Exception:
                if open_t and close_t:
                    try:
                        from datetime import datetime as dt
                        h = (dt.strptime(close_t, "%H:%M") - dt.strptime(open_t, "%H:%M")).seconds / 3600
                        hours = round(h, 2)
                    except Exception:
                        pass

            result[name] = {"open_time": open_t, "close_time": close_t, "hours": hours}

        print(f"[OK] Табель Google Sheets: {len(result)} записей за {target_date}")
        return result
    except Exception as e:
        print(f"[WARN] Табель Google Sheets: {_safe_error(e)}")
        return {}

def classify_role(role_code: str) -> str:
    rc = (role_code or "").lower()
    for r in BARISTA_ROLES:
        if r.lower() in rc: return "barista"
    for r in WAITER_ROLES:
        if r.lower() in rc: return "waiter"
    return "other"

# ─── Чтение Excel-выгрузок из iiko ─────────────────────────────────────────

# Возможные названия колонок в разных версиях iiko (регистр игнорируется)
_SALES_COL_MAPS = {
    "name":    ["кассир", "сотрудник", "cashier", "employee", "имя", "name"],
    "revenue": ["сумма со скидками", "выручка", "сумма", "sumafterdiscounts", "revenue", "продажи"],
    "orders":  ["количество заказов", "заказы", "orderscount", "orders", "кол-во заказов"],
}
_HOURS_COL_MAPS = {
    "name":       ["сотрудник", "employee", "имя", "name", "кассир"],
    "role":       ["должность", "роль", "role", "position", "код должности"],
    "hours":      ["отработано", "часы", "workedhours", "рабочее время", "часов"],
    "open_time":  ["приход", "начало смены", "clockin", "clock in", "время прихода", "открытие"],
    "close_time": ["уход", "конец смены", "clockout", "clock out", "время ухода", "закрытие"],
}

def _find_col(df_cols: list, variants: list) -> str | None:
    """Найти колонку датафрейма по списку возможных названий (без учёта регистра и пробелов)."""
    normalized = {c.lower().strip(): c for c in df_cols}
    for v in variants:
        if v.lower() in normalized:
            return normalized[v.lower()]
    # мягкий поиск — вхождение
    for v in variants:
        for c_low, c_orig in normalized.items():
            if v.lower() in c_low:
                return c_orig
    return None

def _parse_time_val(val) -> str | None:
    """Преобразовать значение ячейки (datetime, time, строка) в HH:MM."""
    if val is None or (isinstance(val, float) and __import__("math").isnan(val)):
        return None
    import math
    if hasattr(val, "strftime"):          # datetime / time
        return val.strftime("%H:%M")
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    # "2025-06-04 08:05:00" или "08:05:00" или "08:05"
    if "T" in s:
        return s[11:16]
    if " " in s and len(s) > 10:
        return s[11:16]
    return s[:5]

def _parse_hours_val(val) -> float:
    """Преобразовать значение ячейки в часы (float)."""
    if val is None:
        return 0.0
    try:
        if hasattr(val, "hour"):          # time object (0:08:30)
            return round(val.hour + val.minute / 60 + val.second / 3600, 2)
        v = float(str(val).replace(",", "."))
        # iiko иногда отдаёт дробь суток (Excel time serial)
        if 0 < v < 1:
            v = round(v * 24, 2)
        return round(v, 2)
    except Exception:
        return 0.0

def load_sales_excel(path: str | Path) -> dict:
    """
    Разобрать Excel-выгрузку «Продажи по кассирам» из iiko.
    Возвращает dict: name -> {revenue, orders, avg_check}
    """
    if not PANDAS_OK:
        print("[ERROR] Установите pandas: pip install pandas openpyxl"); return {}
    try:
        df = pd.read_excel(path, header=None)
    except Exception as e:
        print(f"[ERROR] Не удалось открыть {path}: {e}"); return {}

    # Найти строку заголовка (первая строка, где есть хоть одно ключевое слово)
    header_row = None
    keywords = set(v for vals in _SALES_COL_MAPS.values() for v in vals)
    for i, row in df.iterrows():
        if any(isinstance(c, str) and c.lower().strip() in keywords for c in row):
            header_row = i; break
    if header_row is None:
        print(f"[WARN] sales: не найдена строка заголовка в {path}"); return {}

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c) for c in df.columns]

    c_name = _find_col(list(df.columns), _SALES_COL_MAPS["name"])
    c_rev  = _find_col(list(df.columns), _SALES_COL_MAPS["revenue"])
    c_ord  = _find_col(list(df.columns), _SALES_COL_MAPS["orders"])

    if not c_name:
        print(f"[WARN] sales: не найдена колонка «Кассир» в {path}"); return {}

    result = {}
    for _, row in df.iterrows():
        name = str(row.get(c_name, "") or "").strip()
        if not name or name.lower() in ("итого", "total", "nan", ""):
            continue
        try:
            rev    = float(str(row.get(c_rev, 0) or 0).replace(" ", "").replace(",", ".")) if c_rev else 0.0
            orders = int(float(str(row.get(c_ord, 0) or 0).replace(" ", "").replace(",", "."))) if c_ord else 0
        except Exception:
            rev, orders = 0.0, 0
        result[name] = {
            "revenue":   round(rev, 2),
            "orders":    orders,
            "avg_check": round(rev / orders, 2) if orders > 0 else 0.0,
        }
    print(f"[OK] sales excel: {len(result)} кассиров из {path}")
    return result

def load_hours_excel(path: str | Path) -> dict:
    """
    Разобрать Excel-выгрузку «Рабочее время» из iiko.
    Возвращает dict: name -> {hours, open_time, close_time, role_raw}
    """
    if not PANDAS_OK:
        print("[ERROR] Установите pandas: pip install pandas openpyxl"); return {}
    try:
        df = pd.read_excel(path, header=None)
    except Exception as e:
        print(f"[ERROR] Не удалось открыть {path}: {e}"); return {}

    header_row = None
    keywords = set(v for vals in _HOURS_COL_MAPS.values() for v in vals)
    for i, row in df.iterrows():
        if any(isinstance(c, str) and c.lower().strip() in keywords for c in row):
            header_row = i; break
    if header_row is None:
        print(f"[WARN] hours: не найдена строка заголовка в {path}"); return {}

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c) for c in df.columns]

    c_name  = _find_col(list(df.columns), _HOURS_COL_MAPS["name"])
    c_role  = _find_col(list(df.columns), _HOURS_COL_MAPS["role"])
    c_hours = _find_col(list(df.columns), _HOURS_COL_MAPS["hours"])
    c_open  = _find_col(list(df.columns), _HOURS_COL_MAPS["open_time"])
    c_close = _find_col(list(df.columns), _HOURS_COL_MAPS["close_time"])

    if not c_name:
        print(f"[WARN] hours: не найдена колонка «Сотрудник» в {path}"); return {}

    result = {}
    for _, row in df.iterrows():
        name = str(row.get(c_name, "") or "").strip()
        if not name or name.lower() in ("итого", "total", "nan", ""):
            continue
        result[name] = {
            "hours":      _parse_hours_val(row.get(c_hours)) if c_hours else 0.0,
            "open_time":  _parse_time_val(row.get(c_open))   if c_open  else None,
            "close_time": _parse_time_val(row.get(c_close))  if c_close else None,
        }
    print(f"[OK] hours excel: {len(result)} сотрудников из {path}")
    return result

def _role_by_name(name: str) -> str:
    """Определить роль сотрудника по имени из BARISTA_NAMES / WAITER_NAMES."""
    normalized = str(name or "").lower().strip()
    if not normalized:
        return "other"
    surname = normalized.split()[0]
    for configured_name in BARISTA_NAMES:
        configured = configured_name.lower().strip()
        if normalized == configured or surname == configured.split()[0]:
            return "barista"
    for configured_name in WAITER_NAMES:
        configured = configured_name.lower().strip()
        if normalized == configured or surname == configured.split()[0]:
            return "waiter"
    return "other"

def build_stats_from_excel(sales_map: dict, hours_map: dict) -> list:
    """
    Собрать all_stats из двух Excel-словарей (без API).
    Роль определяется по BARISTA_NAMES / WAITER_NAMES из настроек.
    Имена сопоставляются нечётко: сначала точно, потом по первому слову (фамилии).
    """
    all_names = set(sales_map) | set(hours_map)

    def _match(name: str, other: dict):
        if name in other:
            return other[name]
        first_word = name.split()[0].lower() if name.split() else ""
        for k, v in other.items():
            if k.split()[0].lower() == first_word:
                return v
        return {}

    result = []
    for name in sorted(all_names):
        s      = _match(name, sales_map)
        h      = _match(name, hours_map)
        role   = _role_by_name(name)
        worked = bool(h.get("hours", 0) > 0 or s.get("orders", 0) > 0)

        if role == "other":
            print(f"[WARN] Сотрудник «{name}» не найден в BARISTA_NAMES / WAITER_NAMES — попадёт в «Остальные»")

        result.append({
            "id":           name,
            "name":         name,
            "role":         role,
            "worked_today": worked,
            "worked_hours": h.get("hours", 0.0),
            "open_time":    h.get("open_time")  if worked else None,
            "close_time":   h.get("close_time") if worked else None,
            "revenue":      s.get("revenue", 0.0),
            "orders_count": s.get("orders", 0),
            "avg_check":    s.get("avg_check", 0.0),
        })
    return result

# ─── Время смены ────────────────────────────────────────────────────────────

def parse_shift_time(raw) -> str | None:
    """Привести время из iiko к формату HH:MM или None."""
    if not raw: return None
    try:
        # iiko может отдавать "2025-06-04T08:05:00" или "08:05:00"
        s = str(raw).strip()
        if "T" in s:
            return s[11:16]
        return s[:5]
    except Exception:
        return None

def check_shift_violation(open_time: str | None, close_time: str | None,
                          planned_open: str | None = None, planned_close: str | None = None):
    """
    Нарушения:
    - Приход: опоздал (пришёл позже плана + 10 мин)
    - Уход:   ушёл раньше (ушёл раньше плана - 10 мин)
    """
    TOLERANCE = 10  # минут допуска

    def to_minutes(t):
        if not t: return None
        try:
            h, m = map(int, t.split(":"))
            return h * 60 + m
        except Exception:
            return None

    open_min  = to_minutes(open_time)
    close_min = to_minutes(close_time)

    # Приход: опоздание = факт > план + допуск
    ol_min = to_minutes(planned_open) + TOLERANCE if planned_open else to_minutes(OPEN_DEADLINE)
    open_late = bool(open_min and ol_min and open_min > ol_min)

    # Уход: ранний уход = факт < план - допуск
    cl_min = to_minutes(planned_close) - TOLERANCE if planned_close else to_minutes(CLOSE_DEADLINE)
    close_early = bool(close_min and cl_min and close_min < cl_min)

    return {
        "open_late":   open_late,
        "close_late":  close_early,  # переименуем в логике, в HTML оставим close_late
    }

# ─── JSON-хранилище за месяц ────────────────────────────────────────────────

def get_month_key(d: date) -> str:
    """Ключ месяца, изолированный от старых и чужих подразделений."""
    import hashlib

    selector = " ".join((ORG_NAME or ORG_ID).split()).casefold()
    if not selector:
        raise RuntimeError("Нельзя создать ключ истории без филиала iiko")
    scope = hashlib.sha256(selector.encode("utf-8")).hexdigest()[:16]
    return f"{d.strftime('%Y-%m')}_{scope}"

def load_month_data(month_key: str) -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"month_{month_key}.json"
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_month_data(month_key: str, data: dict):
    """Атомарно сохранить локальную резервную копию."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"month_{month_key}.json"
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with open(temp_path, "w", encoding="utf-8") as file_obj:
        json.dump(data, file_obj, ensure_ascii=False, indent=2)
        file_obj.flush()
        os.fsync(file_obj.fileno())
    os.replace(temp_path, path)

# ─── PostgreSQL хранилище ────────────────────────────────────────────────────

DATABASE_URL = os.environ.get("DATABASE_URL", "")
PG_HOST = os.environ.get("PG_HOST", "")
PG_PORT = int(os.environ.get("PG_PORT", "5432"))
PG_DB = os.environ.get("PG_DB", "")
PG_USER = os.environ.get("PG_USER", "")
PG_PASS = os.environ.get("PG_PASS", "")


def _pg_configured() -> bool:
    return bool(DATABASE_URL or all((PG_HOST, PG_DB, PG_USER, PG_PASS)))


def _pg_conn():
    import psycopg2

    if DATABASE_URL:
        return psycopg2.connect(DATABASE_URL, connect_timeout=10)
    if not _pg_configured():
        raise RuntimeError("PostgreSQL не настроен")
    return psycopg2.connect(
        host=PG_HOST,
        port=PG_PORT,
        dbname=PG_DB,
        user=PG_USER,
        password=PG_PASS,
        connect_timeout=10,
        sslmode=os.environ.get("PGSSLMODE", "require"),
    )

_PG_READY = False


def _pg_init():
    """Создать таблицу при первом обращении к настроенной БД."""
    global _PG_READY
    if _PG_READY:
        return True
    if not _pg_configured():
        return False
    try:
        with _pg_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS month_data_v2 (
                        month_key VARCHAR(32) PRIMARY KEY,
                        data      JSONB NOT NULL,
                        updated   TIMESTAMP DEFAULT NOW()
                    )
                """)
            conn.commit()
        _PG_READY = True
        return True
    except Exception as e:
        print(f"[WARN] PostgreSQL недоступен: {_safe_error(e)}")
        return False

def load_month_data_pg(month_key: str) -> dict:
    """Загрузить историю из PostgreSQL, с fallback на JSON."""
    if not _pg_configured() or not _pg_init():
        return load_month_data(month_key)
    try:
        with _pg_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT data FROM month_data_v2 WHERE month_key = %s", (month_key,))
                row = cur.fetchone()
                if row:
                    print(f"[OK] История из PostgreSQL: {month_key}")
                    return row[0]
    except Exception as e:
        print(f"[WARN] PostgreSQL чтение: {_safe_error(e)}")
    # Fallback на локальный JSON
    return load_month_data(month_key)

def save_month_data_pg(month_key: str, data: dict):
    """Сохранить историю в PostgreSQL и локально."""
    # Локальная копия — резерв, но её ошибка не должна блокировать PostgreSQL.
    try:
        save_month_data(month_key, data)
    except OSError as exc:
        print(f"[WARN] Локальная резервная копия: {_safe_error(exc)}")
        if not _pg_configured():
            raise
    if not _pg_configured() or not _pg_init():
        return
    try:
        with _pg_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO month_data_v2 (month_key, data, updated)
                    VALUES (%s, %s, NOW())
                    ON CONFLICT (month_key) DO UPDATE
                    SET data = EXCLUDED.data, updated = NOW()
                """, (month_key, json.dumps(data, ensure_ascii=False)))
            conn.commit()
        print(f"[OK] История сохранена в PostgreSQL: {month_key}")
    except Exception as e:
        print(f"[WARN] PostgreSQL запись: {_safe_error(e)}")

def update_month_data(month_data: dict, all_emps: list, today: date):
    """Обновить историю месяца. Записывает ВСЕХ сотрудников (в т.ч. не работавших сегодня)."""
    date_str = today.isoformat()
    for emp in all_emps:
        eid = emp["id"]
        if eid not in month_data:
            month_data[eid] = {
                "name": emp["name"], "role": emp["role"],
                "days": [], "violations": []
            }
        # Убедиться что поле violations есть (для старых записей)
        if "violations" not in month_data[eid]:
            month_data[eid]["violations"] = []

        days = month_data[eid]["days"]
        existing = next((d for d in days if d["date"] == date_str), None)
        day_record = {
            "date":        date_str,
            "worked":      emp["worked_today"],
            "hours":       emp["worked_hours"],
            "avg_check":   emp["avg_check"],
            "revenue":     emp["revenue"],
            "orders":      emp["orders_count"],
            "open_time":   emp.get("open_time"),
            "close_time":  emp.get("close_time"),
            "dop_bar":     emp.get("dop_bar", 0.0),
            "dop_kitchen": emp.get("dop_kitchen", 0.0),
            "dop_water":   emp.get("dop_water", 0.0),
            "dop_alcohol": emp.get("dop_alcohol", 0.0),
            "dops":        emp.get("dop_bar", 0.0) + emp.get("dop_kitchen", 0.0) + emp.get("dop_water", 0.0) + emp.get("dop_alcohol", 0.0),
            "desserts":    emp.get("desserts", 0.0),  # Десерты + Выпечка + Прикассовая зона
        }
        if existing:
            existing.update(day_record)
        else:
            days.append(day_record)

        # Зафиксировать нарушения с учётом индивидуального графика
        viol = check_shift_violation(
            emp.get("open_time"), emp.get("close_time"),
            emp.get("planned_open"), emp.get("planned_close")
        )
        viols_list = month_data[eid]["violations"]
        # Удалить старую запись за сегодня, добавить свежую
        month_data[eid]["violations"] = [v for v in viols_list if v.get("date") != date_str]
        if viol["open_late"] or viol["close_late"]:
            month_data[eid]["violations"].append({
                "date": date_str,
                "open_late":  viol["open_late"],
                "close_late": viol["close_late"],
                "open_time":  emp.get("open_time"),
                "close_time": emp.get("close_time"),
            })
    return month_data

def calc_monthly_avg_check(days: list) -> float:
    total_rev    = sum(d.get("revenue", 0) for d in days if d.get("worked"))
    total_orders = sum(d.get("orders",  0) for d in days if d.get("worked"))
    return round(total_rev / total_orders, 2) if total_orders > 0 else 0.0

def calc_monthly_category_pct(days: list, category: str) -> float:
    """Процент категории от общей выручки за месяц."""
    total_rev = sum(d.get("revenue", 0) for d in days if d.get("worked"))
    total_cat = sum(d.get(category, 0) for d in days if d.get("worked"))
    return round(total_cat / total_rev, 4) if total_rev > 0 else 0.0

def get_kpi_status(avg_check: float, role: str):
    """Текущая и ближайшая следующая ступени KPI."""
    levels = KPI.get(role, [])
    current = next(
        (level for level in levels if avg_check >= level["threshold"]),
        None,
    )
    higher_levels = [
        level for level in levels if avg_check < level["threshold"]
    ]
    next_level = min(
        higher_levels,
        key=lambda level: level["threshold"],
        default=None,
    )
    return current, next_level

# ─── Сборка статистики ──────────────────────────────────────────────────────

def build_stats(employees, hours_data, sales_data, shift_data=None, attendance=None, schedule=None, category_data=None):
    """
    Собирает статистику. Сопоставление по имени сотрудника.
    Показывает только сотрудников из BARISTA_NAMES + WAITER_NAMES.
    """
    # Продажи: WaiterName → {revenue, orders, avg_check}
    sales_map = {}
    if sales_data and "data" in sales_data:
        for row in sales_data["data"]:
            raw_name = row.get("WaiterName", "") or ""
            parts = raw_name.split()
            name = " ".join(parts[:-1]) if parts and parts[-1].isdigit() else raw_name
            name = name.strip()
            if not name:
                continue
            rev     = float(row.get("DishDiscountSumInt",         0) or 0)
            orders  = int(  row.get("UniqOrderId",                0) or 0)
            avg     = float(row.get("DishDiscountSumInt.average", 0) or 0)
            if name in sales_map:
                sales_map[name]["revenue"] += rev
                sales_map[name]["orders"]  += orders
            else:
                sales_map[name] = {"revenue": rev, "orders": orders, "avg_check": round(avg, 2)}
    for v in sales_map.values():
        if v.get("avg_check", 0) == 0 and v["orders"] > 0:
            v["avg_check"] = round(v["revenue"] / v["orders"], 2)

    # Часы: пока пустые (разберёмся с полями позже)
    hours_map = {}
    if hours_data and "data" in hours_data:
        for row in hours_data["data"]:
            raw_name = row.get("Employee", "") or row.get("EmployeeCard", "") or ""
            name = " ".join(raw_name.split()[:-1]) if raw_name and raw_name.split()[-1].isdigit() else raw_name
            name = name.strip()
            if not name:
                continue
            h = float(row.get("WorkingHours", 0) or row.get("WorkedHours", 0) or 0)
            hours_map[name] = {
                "hours":      round(h, 2),
                "open_time":  parse_shift_time(row.get("FirstClockIn")  or row.get("ClockInTime")),
                "close_time": parse_shift_time(row.get("LastClockOut") or row.get("ClockOutTime")),
            }

    # Время смены: берём минимальное время открытия и максимальное закрытия
    shift_map = {}
    if shift_data and "data" in shift_data:
        for row in shift_data["data"]:
            raw_name = row.get("WaiterName", "") or ""
            parts = raw_name.split()
            name = " ".join(parts[:-1]) if parts and parts[-1].isdigit() else raw_name
            name = name.strip()
            if not name:
                continue
            open_t  = parse_shift_time(row.get("OpenTime"))
            close_t = parse_shift_time(row.get("CloseTime"))
            if name not in shift_map:
                shift_map[name] = {"open_time": open_t, "close_time": close_t}
            else:
                # Берём минимальное открытие и максимальное закрытие
                cur = shift_map[name]
                if open_t and (not cur["open_time"] or open_t < cur["open_time"]):
                    cur["open_time"] = open_t
                if close_t and (not cur["close_time"] or close_t > cur["close_time"]):
                    cur["close_time"] = close_t

    # Собираем только нужных сотрудников
    all_names = set(BARISTA_NAMES) | set(WAITER_NAMES)
    result = []
    for name in sorted(all_names):
        role = _role_by_name(name)

        # Мягкое сопоставление — точно или по фамилии
        def _find(d, n):
            if n in d: return d[n]
            first = n.split()[0].lower()
            for k, v in d.items():
                if k.split()[0].lower() == first:
                    return v
            return {}

        s  = _find(sales_map, name)
        h  = _find(hours_map, name)
        sh = _find(shift_map, name)
        at = _find(attendance or {}, name)
        sc = _find(schedule  or {}, name)
        cats = _find(category_data or {}, name)

        # Только из табеля Google Sheets — никаких данных из iiko для времени
        open_t  = at.get("open_time")
        close_t = at.get("close_time")
        hours   = at.get("hours", 0.0)
        # Работал если есть продажи ИЛИ в табеле указано время прихода
        worked  = bool(s.get("orders", 0) > 0 or bool(open_t))

        revenue = s.get("revenue", 0.0)

        result.append({
            "id":            name,
            "name":          name,
            "role":          role,
            "worked_today":  worked,
            "worked_hours":  hours,
            "open_time":     open_t  if worked else None,
            "close_time":    close_t if worked else None,
            "revenue":       revenue,
            "orders_count":  s.get("orders",    0),
            "avg_check":     s.get("avg_check", 0.0),
            "planned_open":  sc.get("open"),
            "planned_close": sc.get("close"),
            "dop_bar":       cats.get("Допы Бар",   0.0) if cats else 0.0,
            "dop_kitchen":   cats.get("Допы Кухня", 0.0) if cats else 0.0,
            "dop_water":     cats.get("Бутилированные напитки", 0.0) if cats else 0.0,
            "dop_alcohol":   cats.get("Алкоголь", 0.0) if cats else 0.0,
            "desserts":    (cats.get("Десерты", 0.0) + cats.get("Выпечка", 0.0) + cats.get("Прикассовая зона", 0.0) + cats.get("Мороженое", 0.0)) if cats else 0.0,
        })
    return result

# ─── HTML ────────────────────────────────────────────────────────────────────

WAVE_SVG = """<svg xmlns='http://www.w3.org/2000/svg' width='120' height='60' opacity='0.18'>
  <path d='M0 20 Q15 5 30 20 Q45 35 60 20 Q75 5 90 20 Q105 35 120 20' fill='none' stroke='%s' stroke-width='3'/>
  <path d='M0 35 Q15 20 30 35 Q45 50 60 35 Q75 20 90 35 Q105 50 120 35' fill='none' stroke='%s' stroke-width='3'/>
</svg>"""


def _shift_cell(open_t, close_t, worked, planned_open=None, planned_close=None):
    """HTML для ячейки времени смены."""
    if not worked:
        return '<span class="dash">—</span>', '<span class="dash">—</span>'

    def fmt(t, label):
        if not t:
            return f'<span class="time-unknown">н/д</span>'
        viol = check_shift_violation(
            t if label == "open" else None,
            t if label == "close" else None,
            planned_open  if label == "open"  else None,
            planned_close if label == "close" else None,
        )
        late = viol["open_late"] if label == "open" else viol["close_late"]
        cls  = "time-late" if late else "time-ok"
        icon = " ⚠️" if late else ""
        # Показываем плановое время в подсказке
        plan = planned_open if label == "open" else planned_close
        if label == "open":
            plan_hint = f" (план {plan})" if plan and late else ""
        else:
            plan_hint = f" (план {plan})" if plan and late else ""
        return f'<span class="{cls}">{t}{icon}{plan_hint}</span>'

    return fmt(open_t, "open"), fmt(close_t, "close")


def _violation_count(viols: list) -> tuple[int, int]:
    open_v  = sum(1 for v in viols if v.get("open_late"))
    close_v = sum(1 for v in viols if v.get("close_late"))
    return open_v, close_v


def _role_section(role_key, role_label, emps, month_data, today, icon):
    if not emps:
        return ""

    kpi_levels   = KPI.get(role_key, [])
    avg_legend  = "".join(
        f'<div class="kpi-legend-item">'
        f'<span class="kpi-dot" style="background:{l["color"]}"></span>'
        f'<span class="kpi-threshold">{l["label"]}</span>'
        f'<span class="kpi-sep">→</span>'
        f'<span class="kpi-bonus">+{l["bonus"]//1000} т.р.</span>'
        f'</div>'
        for l in reversed(kpi_levels)
    )
    dop_label  = "Допы"
    legend_html = (
        f'<div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center">'
        f'<div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:0.5px">Ср. чек:</div>'
        f'{avg_legend}'
        f'<div style="width:1px;height:20px;background:var(--border)"></div>'
        f'<div class="kpi-legend-item"><span class="kpi-dot" style="background:var(--olive)"></span>'
        f'<span class="kpi-threshold">Допы ≥ {int((DOBY_BAR_TARGET if role_key == "barista" else DOBY_KITCHEN_TARGET)*100)}%</span>'
        f'<span class="kpi-sep">→</span><span class="kpi-bonus">+6 т.р.</span></div>'
        f'<div class="kpi-legend-item"><span class="kpi-dot" style="background:var(--olive)"></span>'
        f'<span class="kpi-threshold">Десерты ≥ {int(DESSERTS_TARGET*100)}%</span>'
        f'<span class="kpi-sep">→</span><span class="kpi-bonus">+6 т.р.</span></div>'
        f'</div>'
    )

    rows = ""
    cards = ""
    for emp in sorted(emps, key=lambda x: (-calc_monthly_avg_check(
            month_data.get(x["id"], {}).get("days", [])), x["name"])):

        eid        = emp["id"]
        emp_month  = month_data.get(eid, {})
        month_days = emp_month.get("days", [])
        viols      = emp_month.get("violations", [])
        monthly_avg = calc_monthly_avg_check(month_days)
        worked_days = sum(1 for d in month_days if d.get("worked"))

        current_kpi, next_kpi = get_kpi_status(monthly_avg, role_key)
        worked = emp["worked_today"]

        # Прогресс
        if next_kpi and monthly_avg > 0:
            prev_t = current_kpi["threshold"] if current_kpi else 0
            pct = min(100, int((monthly_avg - prev_t) / (next_kpi["threshold"] - prev_t) * 100))
            bar_color = next_kpi["color"]
        elif current_kpi:
            pct, bar_color = 100, current_kpi["color"]
        else:
            min_t = kpi_levels[-1]["threshold"] if kpi_levels else 0
            pct = min(100, int(monthly_avg / min_t * 100)) if min_t else 0
            bar_color = "#C4A882"

        bonus_html = (f'<span class="bonus-badge" style="border-color:{current_kpi["color"]};'
                      f'color:{current_kpi["color"]}">+{current_kpi["bonus"]//1000}т.р.</span>'
                      if current_kpi else "")

        next_html = ""
        if next_kpi and not (current_kpi and current_kpi == kpi_levels[0]):
            gap = next_kpi["threshold"] - monthly_avg
            next_html = f'<span class="next-target">до {next_kpi["label"]}: ещё {gap:.0f} ₽</span>'

        # Спарклайн
        recent = sorted([d for d in month_days if d.get("worked")], key=lambda d: d["date"])[-7:]
        sparkline = ""
        if len(recent) > 1:
            vals = [d["avg_check"] for d in recent if d.get("avg_check", 0) > 0]
            if vals:
                mn, mx = min(vals), max(vals)
                rng = mx - mn or 1
                pts = [f"{i*40/max(len(vals)-1,1):.1f},{16-((v-mn)/rng*14):.1f}" for i, v in enumerate(vals)]
                sparkline = (f'<svg viewBox="0 0 40 18" class="sparkline">'
                             f'<polyline points="{" ".join(pts)}" fill="none" '
                             f'stroke="{bar_color}" stroke-width="1.8" '
                             f'stroke-linecap="round" stroke-linejoin="round"/></svg>')

        # Нарушения
        open_v, close_v = _violation_count(viols)
        viol_html = ""
        if open_v:
            viol_html += f'<span class="viol-chip">⏰ откр. ×{open_v}</span>'
        if close_v:
            viol_html += f'<span class="viol-chip">🔒 закр. ×{close_v}</span>'

        # Сегодняшние значения
        open_cell, close_cell = _shift_cell(
            emp.get("open_time"), emp.get("close_time"), worked,
            emp.get("planned_open"), emp.get("planned_close")
        )
        hours_fmt  = f"{emp['worked_hours']:.1f} ч" if worked and emp["worked_hours"] > 0 else '<span class="dash">—</span>'
        today_chk  = f"{emp['avg_check']:,.0f} ₽"   if worked and emp["avg_check"] > 0   else '<span class="dash">—</span>'
        month_chk  = f"{monthly_avg:,.0f} ₽"         if monthly_avg > 0                   else '<span class="dash">—</span>'

        # Экранируем внешние данные перед вставкой в HTML.
        safe_name = html.escape(str(emp.get("name", "")), quote=True)

        # Плановое время из графика
        plan_open  = emp.get("planned_open")
        plan_close = emp.get("planned_close")
        plan_html  = ""
        if plan_open or plan_close:
            po = plan_open  or "—"
            pc = plan_close or "—"
            plan_html = f'<span style="color:var(--muted);font-size:10px">📅 план: {po}–{pc}</span>'

        # KPI по категориям — берём месячные данные
        revenue     = emp.get("revenue", 0.0)
        dop_bar     = emp.get("dop_bar", 0.0)
        dop_kitchen = emp.get("dop_kitchen", 0.0)
        desserts    = emp.get("desserts", 0.0)
        role_key    = emp.get("role", "other")
        dop_amount  = dop_bar if role_key == "barista" else dop_kitchen
        dop_label   = "Допы Бар" if role_key == "barista" else "Допы Кухня"

        # Месячные показатели по категориям
        # Допы = Допы Бар + Допы Кухня + бутилированные напитки + алкоголь.
        monthly_dop  = calc_monthly_category_pct(month_days, "dops")
        monthly_des  = calc_monthly_category_pct(month_days, "desserts")
        monthly_rev  = sum(d.get("revenue", 0) for d in month_days if d.get("worked"))

        def _kpi_cell(actual_pct, target_pct, bonus=6000):
            if monthly_rev <= 0:
                return '<span class="dash">—</span>'
            display_pct = round(actual_pct * 100, 1)
            target_display = int(target_pct * 100)
            bar_w = min(100, int(actual_pct / target_pct * 100)) if target_pct > 0 else 0
            ok = actual_pct >= target_pct
            color = "var(--olive)" if ok else "var(--terra2)"
            badge = f'<span class="bonus-badge" style="border-color:{color};color:{color}">+{bonus//1000}т.р.</span>' if ok else ""
            gap_pct = max(0, target_pct - actual_pct)
            gap_str = f'<span class="next-target">ещё {gap_pct*100:.1f}%</span>' if not ok else ""
            return f'''<div class="month-row">
              <span class="month-val" style="color:{color}">{display_pct}%</span>
              {badge}
            </div>
            <div class="progress-wrap">
              <div class="progress-bar" style="width:{bar_w}%;background:{color}"></div>
            </div>
            <div class="progress-meta"><span style="color:var(--muted);font-size:10px">цель {target_display}%{" · " + gap_str if gap_str else ""}</span></div>'''

        dop_target = DOBY_BAR_TARGET if role_key == "barista" else DOBY_KITCHEN_TARGET
        dop_cell = _kpi_cell(monthly_dop, dop_target)
        desserts_cell = _kpi_cell(monthly_des, DESSERTS_TARGET)

        row_cls  = "" if worked else "row-absent"
        card_cls = "card-absent" if not worked else ""

        # Десктопная строка таблицы
        rows += f"""
        <tr class="{row_cls}">
          <td class="emp-name-cell">
            <div class="name-main">{safe_name}</div>
            <div class="name-meta">{worked_days} дн.{"&nbsp;&nbsp;" + viol_html if viol_html else ""}</div>
            {f'<div class="name-meta">{plan_html}</div>' if plan_html else ""}
          </td>
          <td class="cell-time">{open_cell}</td>
          <td class="cell-time">{close_cell}</td>
          <td class="cell-hours">{hours_fmt}</td>
          <td class="cell-today">{today_chk}</td>
          <td class="cell-month">{dop_cell}</td>
          <td class="cell-month">{desserts_cell}</td>
          <td class="cell-month">
            <div class="month-row">
              <span class="month-val">{month_chk}</span>
              {bonus_html}
            </div>
            <div class="progress-wrap">
              <div class="progress-bar" style="width:{pct}%;background:{bar_color}"></div>
            </div>
            {f'<div class="progress-meta">{next_html}</div>' if next_html else ""}
          </td>
        </tr>"""

        # Мобильная карточка
        absent_label = '<span style="color:var(--muted);font-size:12px;">не работал сегодня</span>' if not worked else ""
        cards += f"""
        <div class="emp-card {card_cls}">
          <div class="card-top">
            <div>
              <div class="card-name">{safe_name}</div>
              <div class="card-meta">{worked_days} дн. {viol_html} {absent_label}</div>
            </div>
            <div class="card-kpi-block">
              <div class="card-kpi-val">{month_chk}</div>
              <div class="card-kpi-label">ср. чек за месяц</div>
              {"<div style='margin-top:4px'>" + bonus_html + "</div>" if bonus_html else ""}
            </div>
          </div>
          <div class="card-grid">
            <div class="card-cell">
              <div class="card-cell-label">Приход</div>
              <div class="card-cell-val">{open_cell if worked else '<span class="dash">—</span>'}</div>
            </div>
            <div class="card-cell">
              <div class="card-cell-label">Уход</div>
              <div class="card-cell-val">{close_cell if worked else '<span class="dash">—</span>'}</div>
            </div>
            <div class="card-cell">
              <div class="card-cell-label">Часов сегодня</div>
              <div class="card-cell-val cell-hours">{hours_fmt}</div>
            </div>
            <div class="card-cell">
              <div class="card-cell-label">Ср. чек сегодня</div>
              <div class="card-cell-val">{today_chk}</div>
            </div>
            <div class="card-cell">
              <div class="card-cell-label" title="Допы Бар + Допы Кухня + Бутилированные напитки + Алкоголь">Допы ≥{int(dop_target * 100)}%</div>
              <div class="card-cell-val">{dop_cell}</div>
            </div>
            <div class="card-cell">
              <div class="card-cell-label">Десерты ≥5%</div>
              <div class="card-cell-val">{desserts_cell}</div>
            </div>
          </div>
          <div class="card-progress">
            <div class="card-progress-bar-wrap">
              <div class="card-progress-bar" style="width:{pct}%;background:{bar_color}"></div>
            </div>
            {f'<div class="card-next">{next_html}</div>' if next_html else ""}
          </div>
        </div>"""

    return f"""
  <div class="role-section">
    <div class="role-header">
      <div class="role-header-top">
        <div class="role-title">{icon} {role_label}</div>
      </div>
      <div class="kpi-legend">{legend_html}</div>
    </div>
    <div class="table-desktop">
    <table>
      <thead><tr>
        <th>Сотрудник</th>
        <th>Приход<br><span class="th-norm">план +10 мин</span></th>
        <th>Уход<br><span class="th-norm">план +10 мин</span></th>
        <th>Часов<br><span class="th-sub">сегодня</span></th>
        <th>Ср. чек<br><span class="th-sub">сегодня</span></th>
        <th title="Допы Бар + Допы Кухня + Бутилированные напитки + Алкоголь">Допы<br><span class="th-norm">цель по роли</span></th>
        <th>Десерты<br><span class="th-norm">≥ 5%</span></th>
        <th>Ср. чек<br><span class="th-sub">за месяц / KPI</span></th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
    </div>
    <div class="table-mobile">
      <div class="emp-cards">{cards}</div>
    </div>
  </div>"""


def generate_html(all_stats: list, month_data: dict, today: date, org_name: str = "") -> str:
    baristas = [e for e in all_stats if e["role"] == "barista"]
    waiters  = [e for e in all_stats if e["role"] == "waiter"]
    others   = [e for e in all_stats if e["role"] == "other"]

    worked_today = sum(1 for e in all_stats if e["worked_today"])
    date_str     = today.strftime("%d.%m.%Y")
    month_label  = MONTH_NAMES[today.month]
    gen_at = _now().strftime("%d.%m.%Y %H:%M")

    b_sec = _role_section("barista", "Бариста",   baristas, month_data, today, "☕")
    w_sec = _role_section("waiter",  "Официанты", waiters,  month_data, today, "🍽️")
    o_sec = _role_section("other",   "Остальные", others,   month_data, today, "👤")

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Здрасте — KPI {date_str}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600&display=swap');

:root {{
  --terra:   #B85C38;
  --terra2:  #8B3E20;
  --beige:   #F2E0CC;
  --beige2:  #EDD4B8;
  --pink:    #E8A898;
  --olive:   #7A8C5E;
  --brown:   #6B3E26;
  --cream:   #FBF4EC;
  --text:    #2C1A0E;
  --muted:   #9C7B6A;
  --border:  #E2C9B4;
  --absent:  #F5EDE4;
  --font:    'Inter', sans-serif;
  --head:    'Montserrat', sans-serif;
  --gap:     24px;
  --pad:     clamp(12px, 4vw, 48px);
}}

* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:var(--cream); color:var(--text); font-family:var(--font); min-height:100vh; }}

/* ── HEADER ── */
header {{
  background: var(--terra);
  padding: 0 var(--pad);
  position: relative; overflow: hidden;
}}
.header-wave {{
  position: absolute; inset: 0; opacity: 0.12;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='200' height='80'%3E%3Cpath d='M0 30 Q25 10 50 30 Q75 50 100 30 Q125 10 150 30 Q175 50 200 30' fill='none' stroke='%23FBF4EC' stroke-width='4'/%3E%3Cpath d='M0 55 Q25 35 50 55 Q75 75 100 55 Q125 35 150 55 Q175 75 200 55' fill='none' stroke='%23FBF4EC' stroke-width='4'/%3E%3C/svg%3E");
  background-size: 200px 80px;
}}
.header-inner {{
  position: relative; z-index: 1;
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 12px;
  padding: 20px 0;
}}
.brand {{ display:flex; align-items:center; gap:14px; }}
.brand-logo {{
  background: var(--beige); border-radius: 12px;
  width: 48px; height: 48px; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  font-family: var(--head); font-weight: 900; font-size: 20px; color: var(--terra);
}}
.brand-name {{
  font-family: var(--head); font-weight: 900; font-size: 22px;
  color: #fff; letter-spacing: 1px; text-transform: uppercase;
}}
.brand-sub {{ font-size: 11px; color: rgba(255,255,255,0.65); margin-top: 2px; }}
.header-meta {{
  display: flex; gap: 20px; flex-wrap: wrap;
  color: rgba(255,255,255,0.75); font-size: 12px;
}}
.header-meta-item {{ display:flex; flex-direction:column; align-items:center; }}
.header-meta-item .hm-label {{ font-size:10px; opacity:0.7; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px; }}
.header-meta-item .hm-val {{ color:#fff; font-weight:700; font-size:15px; font-family:var(--head); }}

/* ── MONTH STRIP ── */
.month-strip {{
  background: var(--beige2);
  border-bottom: 2px solid var(--border);
  padding: 10px var(--pad);
  display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
}}
.month-pill {{
  background: var(--terra); color: #fff;
  border-radius: 20px; padding: 4px 14px;
  font-family: var(--head); font-size: 11px; font-weight: 700; letter-spacing: 0.5px;
  text-transform: uppercase; white-space: nowrap;
}}
.month-info {{ font-size: 12px; color: var(--brown); }}
.month-info b {{ color: var(--terra2); }}
.ms-sep {{ width:1px; height:18px; background:var(--border); flex-shrink:0; }}

/* ── ROLE SECTION ── */
.role-section {{
  margin: var(--gap) var(--pad) 0;
  background: #fff;
  border: 1px solid var(--border);
  border-radius: 16px;
  overflow: hidden;
  box-shadow: 0 2px 12px rgba(184,92,56,0.06);
}}
.role-header {{
  padding: 14px 20px;
  background: var(--beige);
  border-bottom: 1px solid var(--border);
}}
.role-header-top {{
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 8px; margin-bottom: 10px;
}}
.role-title {{
  font-family: var(--head); font-size: 15px; font-weight: 800;
  color: var(--terra2); text-transform: uppercase; letter-spacing: 0.8px;
}}

/* KPI Legend — крупнее и заметнее */
.kpi-legend {{ display:flex; gap:10px; flex-wrap:wrap; }}
.kpi-legend-item {{
  display: flex; align-items: center; gap: 8px;
  background: #fff; border: 1.5px solid var(--border);
  border-radius: 10px; padding: 6px 12px;
  font-size: 13px; color: var(--text);
}}
.kpi-dot {{ width:10px; height:10px; border-radius:50%; flex-shrink:0; }}
.kpi-threshold {{ font-weight: 700; color: var(--text); font-family: var(--head); font-size:13px; }}
.kpi-bonus {{
  font-weight: 800; font-size: 14px; font-family: var(--head);
  color: var(--terra2);
}}
.kpi-sep {{ color: var(--border); margin: 0 2px; }}

/* ── DESKTOP TABLE ── */
.table-desktop {{ display: block; }}
.table-mobile  {{ display: none; }}

table {{ width:100%; border-collapse:collapse; font-size:13px; }}
thead th {{
  padding: 10px 16px; text-align:left;
  font-size: 10px; text-transform: uppercase; letter-spacing: 0.8px;
  color: var(--muted); font-weight: 600; font-family: var(--head);
  border-bottom: 1px solid var(--border);
  background: rgba(242,224,204,0.25);
  white-space: nowrap;
}}
.th-sub  {{ font-size:9px; opacity:0.7; text-transform:none; letter-spacing:0; font-family:var(--font); font-weight:400; }}
.th-norm {{ font-size:9px; color:var(--olive); text-transform:none; letter-spacing:0; font-family:var(--font); font-weight:600; }}
tbody tr {{ border-bottom:1px solid var(--border); transition:background 0.15s; }}
tbody tr:last-child {{ border-bottom:none; }}
tbody tr:hover {{ background: rgba(242,224,204,0.3); }}
tbody tr.row-absent {{ background: var(--absent); }}
td {{ padding:14px 16px; vertical-align:middle; }}

/* ── MOBILE CARDS ── */
.emp-cards {{ display: flex; flex-direction: column; gap: 0; }}
.emp-card {{
  padding: 16px 16px;
  border-bottom: 1px solid var(--border);
}}
.emp-card:last-child {{ border-bottom: none; }}
.emp-card.card-absent {{ background: var(--absent); }}
.card-top {{
  display: flex; align-items: flex-start; justify-content: space-between;
  gap: 8px; margin-bottom: 12px;
}}
.card-name {{ font-weight: 700; font-size: 15px; color: var(--text); }}
.card-meta {{ font-size: 11px; color: var(--muted); margin-top: 3px; display:flex; gap:6px; flex-wrap:wrap; }}
.card-kpi-block {{ text-align: right; flex-shrink: 0; }}
.card-kpi-val {{
  font-family: var(--head); font-weight: 900; font-size: 20px; color: var(--terra2);
}}
.card-kpi-label {{ font-size: 10px; color: var(--muted); margin-top: 1px; }}
.card-grid {{
  display: grid; grid-template-columns: 1fr 1fr;
  gap: 8px; margin-bottom: 10px;
}}
.card-cell {{
  background: rgba(242,224,204,0.3);
  border-radius: 8px; padding: 8px 10px;
}}
.card-cell-label {{ font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
.card-cell-val {{ font-size: 14px; font-weight: 600; }}
.card-progress {{ margin-top: 2px; }}
.card-progress-bar-wrap {{
  height: 5px; background: rgba(107,62,38,0.1); border-radius: 3px; overflow: hidden; margin: 6px 0 4px;
}}
.card-progress-bar {{ height: 100%; border-radius: 3px; }}
.card-next {{ font-size: 11px; color: var(--muted); }}

/* ── SHARED ELEMENTS ── */
.emp-name-cell .name-main {{ font-weight:600; font-size:14px; color:var(--text); }}
.emp-name-cell .name-meta  {{
  font-size:11px; color:var(--muted); margin-top:3px;
  display:flex; align-items:center; gap:6px; flex-wrap:wrap;
}}
.viol-chip {{
  background: #FFF0ED; color: #C0392B;
  border: 1px solid #F5C6BB;
  border-radius: 10px; padding: 1px 7px; font-size: 10px; font-weight: 600;
  white-space: nowrap;
}}
.time-ok      {{ color: var(--olive); font-weight:600; }}
.time-late    {{ color: #C0392B; font-weight:700; }}
.time-unknown {{ color: var(--muted); }}
.cell-hours   {{ color: var(--olive); font-weight:600; white-space:nowrap; }}
.cell-today   {{ font-weight:600; white-space:nowrap; }}
.dash         {{ color: var(--muted); font-weight:400; }}

.month-row {{ display:flex; align-items:center; gap:8px; margin-bottom:5px; flex-wrap:wrap; }}
.month-val  {{ font-family:var(--head); font-size:17px; font-weight:900; color:var(--terra2); }}
.bonus-badge {{
  font-size:11px; font-weight:700; padding:3px 8px; border-radius:20px;
  border:1.5px solid; white-space:nowrap; font-family:var(--head);
}}
.progress-wrap {{
  height:5px; background:rgba(107,62,38,0.1); border-radius:3px; overflow:hidden;
}}
.progress-bar {{ height:100%; border-radius:3px; }}
.progress-meta {{ margin-top:4px; }}
.next-target {{ font-size:11px; color:var(--muted); }}
.cell-spark {{ width:52px; }}
.sparkline  {{ width:44px; height:20px; display:block; }}

/* ── FOOTER ── */
footer {{
  margin: var(--gap) var(--pad) 40px;
  padding-top: 14px; border-top: 1px solid var(--border);
  display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px;
  font-size:11px; color:var(--muted);
}}
footer b {{ color:var(--terra); }}

/* ── LEGEND BOX ── */
.legend-box {{
  margin: var(--gap) var(--pad) 0;
  background: #fff;
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px 18px;
  display: flex; gap: 16px; flex-wrap: wrap; align-items: center;
  font-size: 12px;
}}
.legend-box b {{ color:var(--terra2); font-family:var(--head); font-size:11px; text-transform:uppercase; letter-spacing:0.5px; margin-right:2px; }}
.leg-item {{ display:flex; align-items:center; gap:6px; color:var(--muted); }}
.leg-ok   {{ color:var(--olive); font-weight:700; }}
.leg-late {{ color:#C0392B; font-weight:700; }}
.leg-abs  {{ width:14px; height:10px; background:var(--absent); border:1px solid var(--border); border-radius:2px; }}

/* ── RESPONSIVE ── */
@media(max-width:700px){{
  .header-inner {{ flex-direction: column; align-items: flex-start; gap:10px; }}
  .header-meta {{ width: 100%; justify-content: space-between; }}
  .ms-sep {{ display:none; }}
  .month-strip {{ gap:8px; }}
  .table-desktop {{ display: none !important; }}
  .table-mobile  {{ display: block; }}
  .legend-box {{ gap:10px; }}
}}
</style>
</head>
<body>

<header>
  <div class="header-wave"></div>
  <div class="header-inner">
    <div class="brand">
      <div class="brand-logo">З</div>
      <div>
        <div class="brand-name">Здрасте</div>
        <div class="brand-sub">KPI — мониторинг среднего чека</div>
      </div>
    </div>
    <div class="header-meta">
      <div class="header-meta-item">
        <span class="hm-label">Дата</span>
        <span class="hm-val">{date_str}</span>
      </div>
      <div class="header-meta-item">
        <span class="hm-label">Работало</span>
        <span class="hm-val">{worked_today} чел.</span>
      </div>
      <div class="header-meta-item">
        <span class="hm-label">Сформирован</span>
        <span class="hm-val" style="font-size:12px;font-weight:600">{gen_at}</span>
      </div>
    </div>
  </div>
</header>

<div class="month-strip">
  <div class="month-pill">📅 {month_label} {today.year}</div>
  <div class="ms-sep"></div>
  <div class="month-info"><b>{today.day}-й день</b> месяца</div>
  <div class="ms-sep"></div>
  <div class="month-info">Средний чек за месяц = <b>выручка ÷ заказы</b> нарастающим итогом. KPI сбрасываются 1-го числа.</div>
</div>

<div class="legend-box">
  <b>Обозначения:</b>
  <div class="leg-item"><span class="leg-ok">08:05</span> — вовремя</div>
  <div class="leg-item"><span class="leg-late">08:20 ⚠️</span> — нарушение</div>
  <div class="leg-item"><div class="leg-abs"></div> серая строка — не работал сегодня</div>
  <div class="leg-item">⏰ откр. ×N / 🔒 закр. ×N — нарушений за месяц</div>
</div>

{b_sec}
{w_sec}
{o_sec}

<footer>
  <span><b>Здрасте</b> · Автоматический KPI-отчёт · запуск ежедневно в 23:00</span>
  <span>Норматив: приход/уход по графику ±10 мин</span>
</footer>

</body>
</html>"""


def _stats_from_month_data(month_data: dict, today: date) -> list:
    """Собрать last-known-good представление без обращения к внешним API."""
    all_stats = []
    for employee_id, employee_data in month_data.items():
        days = employee_data.get("days", [])
        today_day = next(
            (day for day in days if day.get("date") == today.isoformat()),
            None,
        )
        all_stats.append({
            "id": employee_id,
            "name": employee_data.get("name", employee_id),
            "role": employee_data.get("role") or _role_by_name(
                employee_data.get("name", "")
            ),
            "worked_today": bool(today_day and today_day.get("worked")),
            "worked_hours": today_day.get("hours", 0) if today_day else 0,
            "open_time": today_day.get("open_time") if today_day else None,
            "close_time": today_day.get("close_time") if today_day else None,
            "revenue": today_day.get("revenue", 0) if today_day else 0,
            "orders_count": today_day.get("orders", 0) if today_day else 0,
            "avg_check": today_day.get("avg_check", 0) if today_day else 0,
            "planned_open": None,
            "planned_close": None,
        })
    return all_stats


def _cached_report_or_error(month_data: dict, today: date) -> str:
    if month_data:
        return generate_html(_stats_from_month_data(month_data, today), month_data, today, "")
    raise RuntimeError("Нет корректной сохранённой версии отчёта")


def _olap_response_ok(value) -> bool:
    return isinstance(value, dict) and isinstance(value.get("data"), list)


def generate_report_html() -> str:
    """Собрать отчёт и вернуть HTML — вызывается сервером при каждом запросе."""
    today = _today()
    month_key = get_month_key(today)
    date_from = today.strftime("%Y-%m-%dT00:00:00")
    date_to   = today.strftime("%Y-%m-%dT23:59:59")

    # Загружаем историю из PostgreSQL
    month_data = load_month_data_pg(month_key)

    try:
        token = get_token()
    except Exception as e:
        print(f"[WARN] iiko недоступна: {_safe_error(e)} — показываем данные из базы")
        # При сбое не перезаписываем корректный день нулями.
        return _cached_report_or_error(month_data, today)

    try:
        org = resolve_organization(get_organizations(token))
        org_id = org["id"]
        org_name = org["name"]
        validate_olap_department_filter(token)
    except Exception as e:
        print(f"[WARN] Филиал iiko не подтверждён: {_safe_error(e)}")
        return _cached_report_or_error(month_data, today)

    try:
        employees = get_employees(token, org_id)
    except Exception:
        employees = []

    hours_data    = get_worked_hours(token, org_id, date_from, date_to)
    sales_data    = get_sales(token, org_id, date_from, date_to)
    shift_data    = get_shift_times(token, org_id, date_from, date_to)
    category_data = get_category_sales(token, org_id, date_from, date_to)
    attendance = load_attendance(today)
    schedule = load_schedule(today)

    required_sources_ok = (
        bool(employees)
        and _olap_response_ok(sales_data)
        and _olap_response_ok(shift_data)
        and isinstance(category_data, dict)
    )
    if not required_sources_ok:
        print("[WARN] Неполная загрузка источников — история не обновлена")
        return _cached_report_or_error(month_data, today)

    all_stats = build_stats(
        employees,
        hours_data,
        sales_data,
        shift_data,
        attendance,
        schedule,
        category_data,
    )

    month_data = update_month_data(month_data, all_stats, today)
    save_month_data_pg(month_key, month_data)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    html = generate_html(all_stats, month_data, today, org_name)
    with open(OUTPUT_DIR / f"kpi_{today.strftime('%Y-%m-%d')}.html", "w", encoding="utf-8") as f:
        f.write(html)
    with open(OUTPUT_DIR / "latest.html", "w", encoding="utf-8") as f:
        f.write(html)
    return html


# ─── MAIN ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Здрасте KPI Report")
    parser.add_argument("--sales",  metavar="FILE.xlsx", help="Excel-выгрузка продаж по кассирам")
    parser.add_argument("--hours",  metavar="FILE.xlsx", help="Excel-выгрузка рабочего времени")
    parser.add_argument("--date",   metavar="YYYY-MM-DD", help="Дата отчёта (по умолчанию — сегодня)")
    args = parser.parse_args()

    today = date.fromisoformat(args.date) if args.date else _today()
    month_key = get_month_key(today)

    print(f"\n{'='*54}")
    print(f"  Здрасте KPI Report | {today.strftime('%d.%m.%Y')}")
    print(f"{'='*54}\n")

    # ── Режим Excel ──────────────────────────────────────────────────────────
    if args.sales or args.hours:
        if not PANDAS_OK:
            print("[ERROR] Для Excel-режима нужен pandas:")
            print("        pip install pandas openpyxl")
            sys.exit(1)
        if not args.sales:
            print("[ERROR] Укажите --sales FILE.xlsx (выгрузка продаж)")
            sys.exit(1)
        if not args.hours:
            print("[ERROR] Укажите --hours FILE.xlsx (выгрузка рабочего времени)")
            sys.exit(1)

        print("[INFO] Режим: Excel-файлы (без подключения к iiko)")
        sales_map = load_sales_excel(args.sales)
        hours_map = load_hours_excel(args.hours)
        all_stats = build_stats_from_excel(sales_map, hours_map)
        org_name  = ""

    # ── Режим API iiko ───────────────────────────────────────────────────────
    else:
        print("[INFO] Режим: API iiko")
        date_from = today.strftime("%Y-%m-%dT00:00:00")
        date_to   = today.strftime("%Y-%m-%dT23:59:59")

        try:
            token = get_token()
        except Exception as e:
            print(f"[ERROR] Авторизация: {_safe_error(e)}"); sys.exit(1)

        try:
            org = resolve_organization(get_organizations(token))
            org_id = org["id"]
            org_name = org["name"]
            validate_olap_department_filter(token)
        except Exception as e:
            print(f"[ERROR] Филиал iiko не подтверждён: {_safe_error(e)}")
            sys.exit(2)

        try:
            employees = get_employees(token, org_id)
            print(f"[OK] Сотрудников: {len(employees)}")
        except Exception as e:
            print(f"[ERROR] Сотрудники: {_safe_error(e)}"); employees = []

        hours_data    = get_worked_hours(token, org_id, date_from, date_to)
        sales_data    = get_sales(token, org_id, date_from, date_to)
        shift_data    = get_shift_times(token, org_id, date_from, date_to)
        category_data = get_category_sales(token, org_id, date_from, date_to)
        attendance = load_attendance(today)
        schedule = load_schedule(today)
        required_sources_ok = (
            bool(employees)
            and _olap_response_ok(sales_data)
            and _olap_response_ok(shift_data)
            and isinstance(category_data, dict)
        )
        if not required_sources_ok:
            print("[ERROR] Неполная загрузка источников; история не изменена")
            sys.exit(2)
        all_stats = build_stats(
            employees,
            hours_data,
            sales_data,
            shift_data,
            attendance,
            schedule,
            category_data,
        )

    # ── Общая часть ──────────────────────────────────────────────────────────
    print(f"[OK] Всего: {len(all_stats)}, работали сегодня: {sum(e['worked_today'] for e in all_stats)}")

    month_data = load_month_data_pg(month_key)
    month_data = update_month_data(month_data, all_stats, today)
    save_month_data_pg(month_key, month_data)
    print(f"[OK] История обновлена: {month_key}")

    html = generate_html(all_stats, month_data, today, org_name)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for path in (OUTPUT_DIR / f"kpi_{today.strftime('%Y-%m-%d')}.html",
                 OUTPUT_DIR / "latest.html"):
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[✓] {path}")

    print(f"{'='*54}\n")


if __name__ == "__main__":
    main()
