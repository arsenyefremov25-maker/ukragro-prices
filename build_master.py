from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "UKRAGRO_ALL_PRICES.xlsx"

# Накопичена база (зберігається в репозиторії після кожного запуску)
DATA_DIR = Path("data")
DB_FILE = DATA_DIR / "prices_db.csv"
MANIFEST_FILE = DATA_DIR / "processed_files.csv"

SCRIPT_VERSION = "3.0-front-month-aliases-db"

FINAL_SHEET_ORDER = [
    "Навігація",
    "Усі ціни",
    "Зернові",
    "Олійні",
    "Рослинні олії",
    "Шроти",
    "Ключові індикатори",
    "Україна — експорт",
    "Україна — внутрішній ринок",
    "Портові спреди",
    "Пшениця — премії якості",
    "Чорне море — конкуренти",
    "База даних",
    "Журнал",
    "Контроль якості",
]

# ---------------------------------------------------------------------------
# Заголовки вхідних файлів.
# Ключ — уніфікована назва колонки, значення — усі відомі варіанти заголовка
# (російські, англійські). Порівняння йде без урахування регістру, пробілів,
# дефісів, ком і крапок, тому "Seller, min" = "seller min" = "SELLER-MIN".
# Якщо UkrAgroConsult знову змінить заголовок — достатньо дописати його сюди.
# ---------------------------------------------------------------------------
HEADER_ALIASES: dict[str, list[str]] = {
    "Дата": ["date", "дата", "day", "report date"],
    "Товар": ["товар", "commodity", "product", "продукт", "культура"],
    "Якість": ["качество", "quality", "grade", "якість", "specification", "spec"],
    "Країна": ["страна", "country", "origin", "країна", "country of origin"],
    "Порт": ["порт", "port", "location", "place", "market", "destination", "port / market"],
    "Базис": [
        "базис", "delivery terms", "delivery term", "basis", "terms", "incoterms",
        "terms of delivery", "базис поставки", "умови поставки",
    ],
    "Валюта": ["валюта", "currency", "curr"],
    "Продавець min": ["seller min", "sellers min", "seller minimum", "offer min", "ask min", "продавец мин", "продавець мін"],
    "Продавець max": ["seller max", "sellers max", "seller maximum", "offer max", "ask max", "продавец макс", "продавець макс"],
    "Покупець min": ["buyer min", "buyers min", "buyer minimum", "bid min", "покупатель мин", "покупець мін"],
    "Покупець max": ["buyer max", "buyers max", "buyer maximum", "bid max", "покупатель макс", "покупець макс"],
    "Місяць поставки": [
        "месяц", "month", "delivery month", "shipment month", "shipment", "delivery period",
        "period", "месяц поставки", "місяць", "місяць поставки",
    ],
    "Сезон": ["сезон", "season", "marketing year", "crop year", "my", "crop"],
    "Тип товару": [
        "тип товара", "commodity type", "type", "product type", "group", "category",
        "commodity group", "тип товару", "type of commodity",
    ],
}

# Без цих колонок аркуш вважається нечитабельним (помилка).
REQUIRED_COLUMNS = ["Товар", "Країна", "Базис", "Валюта"]

# Якщо у файлі немає колонки «тип товару» — визначаємо категорію за товаром.
COMMODITY_TYPE_FALLBACK = {
    "wheat": "Grain", "corn": "Grain", "maize": "Grain", "barley": "Grain",
    "rye": "Grain", "oats": "Grain", "sorghum": "Grain", "peas": "Grain",
    "rapeseed": "Oilseeds", "soybeans": "Oilseeds", "soybean": "Oilseeds",
    "sunflower seed": "Oilseeds", "sunflower seeds": "Oilseeds", "linseed": "Oilseeds",
    "sunflower oil": "Vegoil", "soybean oil": "Vegoil", "rapeseed oil": "Vegoil",
    "palm oil": "Vegoil",
    "sunflower meal": "Meals", "soybean meal": "Meals", "rapeseed meal": "Meals",
    "sunflower cake": "Meals", "soybean cake": "Meals", "rapeseed cake": "Meals",
}

# ---------------------------------------------------------------------------
# Синоніми портів. Кортеж = «будь-яке з цих значень», першим іде пріоритетне.
# Порівняння не залежить від порядку портів через кому:
# "Odesa, Chornomorsk, Pivdennyi" == "Chornomorsk, Odesa, Pivdennyi".
# ---------------------------------------------------------------------------
ODESA = ("Odesa", "Odessa")
RENI = ("Reni",)
BLACK_SEA_PORTS = (
    "Chornomorsk, Odesa, Pivdennyi",  # назва з 2025 року
    "Chornomorsk",                    # назва у 2024 році
    "Chornomorsk, Odesa",
    "Big Odesa", "Greater Odesa", "Odesa ports",
)
WESTERN_BORDER = ("Western border", "West border", "Western Ukraine border")
NOVOROSSIYSK = ("Novorossiysk", "Novorossiisk")
ROSTOV_AZOV = ("Rostov/Azov", "Azov/Rostov", "Rostov, Azov", "Azov")
EU_BLACK_SEA_PORTS = ("Constanta, Varna, Burgas", "Constanta")

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "мая": 5, "май": 5, "июн": 6,
    "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12,
    "січ": 1, "лют": 2, "бер": 3, "кві": 4, "тра": 5, "чер": 6,
    "лип": 7, "сер": 8, "вер": 9, "жов": 10, "лис": 11, "гру": 12,
}

TEXT_COLUMNS = [
    "Товар",
    "Якість",
    "Країна",
    "Порт",
    "Базис",
    "Валюта",
    "Місяць поставки",
    "Сезон",
    "Тип товару",
]

PRICE_COLUMNS = [
    "Продавець min",
    "Продавець max",
    "Покупець min",
    "Покупець max",
]

DERIVED_PRICE_COLUMNS = [
    "Середня ціна продавця",
    "Середня ціна покупця",
]

ALL_VALUE_COLUMNS = PRICE_COLUMNS + DERIVED_PRICE_COLUMNS

KEY_COLUMNS = [
    "Дата",
    "Товар",
    "Якість",
    "Країна",
    "Порт",
    "Базис",
    "Валюта",
    "Місяць поставки",
    "Сезон",
    "Тип товару",
]

DB_COLUMNS = (
    ["Дата", "Товар", "Якість", "Країна", "Порт", "Базис", "Валюта"]
    + PRICE_COLUMNS
    + ["Місяць поставки", "Сезон", "Тип товару"]
    + DERIVED_PRICE_COLUMNS
    + ["Джерело — файл", "Джерело — аркуш", "_rank"]
)

SERIES_COLUMNS = [
    "Тип товару",
    "Товар",
    "Якість",
    "Країна",
    "Порт",
    "Базис",
    "Валюта",
    "Місяць поставки",
    "Сезон",
]

CATEGORY_NAMES = {
    "Grain": "Зернові",
    "Oilseeds": "Олійні",
    "Vegoil": "Рослинні олії",
    "Meals": "Шроти",
}

CATEGORY_ORDER = {
    "Grain": 0,
    "Oilseeds": 1,
    "Vegoil": 2,
    "Meals": 3,
}

METRIC_NAMES = {
    "Продавець min": "Продавець — min",
    "Продавець max": "Продавець — max",
    "Покупець min": "Покупець — min",
    "Покупець max": "Покупець — max",
    "Середня ціна продавця": "Продавець — середня",
    "Середня ціна покупця": "Покупець — середня",
}

METRIC_ORDER = {name: index for index, name in enumerate(ALL_VALUE_COLUMNS)}

TITLE_FILL = "17365D"
HEADER_FILL = "1F4E78"
DATE_FILL = "70AD47"
NOTE_FILL = "D9EAF7"
GROUP_FILL = "F4F8FC"
SPREAD_FILL = "FFF2CC"
POSITIVE_FILL = "E2F0D9"
NEGATIVE_FILL = "FCE4D6"
WHITE = "FFFFFF"
TEXT_COLOR = "1F2937"
BORDER_COLOR = "B8CCE4"


# ===========================================================================
# Нормалізація
# ===========================================================================


def header_key(value: object) -> str:
    """Ключ для порівняння заголовків: лише літери та цифри, нижній регістр."""
    return re.sub(r"[^0-9a-zа-яёіїєґ]+", "", str(value).lower())


HEADER_LOOKUP: dict[str, str] = {}
for _canonical, _aliases in HEADER_ALIASES.items():
    for _alias in _aliases:
        HEADER_LOOKUP[header_key(_alias)] = _canonical


def norm_text(value: object) -> str:
    """Нормалізація значень для зіставлення: регістр, пробіли, дефіси."""
    text = str(value).lower().strip()
    text = re.sub(r"[\s\-–—_]+", " ", text)
    return text.strip()


def norm_port(value: object) -> str:
    """Порт без залежності від порядку: 'B, A' == 'A, B'."""
    parts = [norm_text(part) for part in re.split(r"[,;/]", str(value))]
    return ", ".join(sorted(part for part in parts if part))


def norm_for(column: str, value: object) -> str:
    return norm_port(value) if column == "Порт" else norm_text(value)


def display_value(value: object) -> object:
    if isinstance(value, tuple):
        return value[0]
    return "—" if value is None else value


# ===========================================================================
# Читання вхідних файлів
# ===========================================================================


def parse_sheet_date(sheet_name: str) -> pd.Timestamp | None:
    sheet_name = sheet_name.strip()
    match = re.fullmatch(r"(\d{2}\.\d{2}\.\d{4})_min_max", sheet_name)
    if not match:
        return None
    return pd.to_datetime(match.group(1), format="%d.%m.%Y", errors="coerce")


def map_header_row(row: list[object]) -> dict[int, str]:
    """Повертає {номер колонки: уніфікована назва} для рядка-кандидата."""
    mapping: dict[int, str] = {}
    used: set[str] = set()
    for position, value in enumerate(row):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            continue
        canonical = HEADER_LOOKUP.get(header_key(value))
        if canonical and canonical not in used:
            mapping[position] = canonical
            used.add(canonical)
    return mapping


def find_header(raw: pd.DataFrame) -> tuple[int, dict[int, str]]:
    best_row, best_mapping = -1, {}
    for row_number in range(min(40, len(raw))):
        mapping = map_header_row(list(raw.iloc[row_number]))
        names = set(mapping.values())
        if "Товар" in names and len(names) > len(best_mapping):
            best_row, best_mapping = row_number, mapping
    if best_row < 0 or len(best_mapping) < 5:
        raise ValueError(
            "Не знайдено рядок заголовків (шукали варіанти: товар/commodity, "
            "базис/delivery terms тощо). Якщо формат змінився — додайте нові "
            "назви в HEADER_ALIASES."
        )
    # Колонка дати зазвичай перша і має порожній заголовок (' ').
    if "Дата" not in best_mapping.values() and 0 not in best_mapping:
        best_mapping[0] = "Дата"
    return best_row, best_mapping


def workbook_rank(sheet_names: list[str]) -> pd.Timestamp:
    dates = [parse_sheet_date(name) for name in sheet_names]
    dates = [date for date in dates if date is not None and pd.notna(date)]
    return max(dates) if dates else pd.Timestamp.min


def clean_text(series: pd.Series) -> pd.Series:
    cleaned = series.astype("string").str.strip()
    return cleaned.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "-": pd.NA, "—": pd.NA, "NaT": pd.NA})


def read_price_sheet(
    file_path: Path,
    sheet_name: str,
    report_date: pd.Timestamp,
    file_rank: pd.Timestamp,
) -> tuple[pd.DataFrame, dict[str, object]]:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row, mapping = find_header(raw)

    body = raw.iloc[header_row + 1 :, list(mapping.keys())].copy()
    body.columns = [mapping[position] for position in mapping]

    missing_required = [column for column in REQUIRED_COLUMNS if column not in body.columns]
    if missing_required:
        raise ValueError(f"Відсутні обов'язкові колонки: {', '.join(missing_required)}")
    if not any(column in body.columns for column in PRICE_COLUMNS):
        raise ValueError("Не знайдено жодної колонки з цінами (seller/buyer min/max)")

    meta: dict[str, object] = {"missing_optional": [], "inner_dates": []}
    for column in TEXT_COLUMNS + PRICE_COLUMNS:
        if column not in body.columns:
            body[column] = pd.NA
            if column != "Тип товару":
                meta["missing_optional"].append(column)

    for column in TEXT_COLUMNS:
        body[column] = clean_text(body[column])
    for column in PRICE_COLUMNS:
        body[column] = pd.to_numeric(body[column], errors="coerce")

    body = body[body["Товар"].notna()].copy()

    # Дата всередині аркуша (колонка Date) — для перевірки копій.
    if "Дата" in body.columns:
        inner = pd.to_datetime(body["Дата"], errors="coerce", dayfirst=True).dropna()
        meta["inner_dates"] = sorted({pd.Timestamp(value).normalize() for value in inner})

    # Тип товару: з файлу або за довідником.
    fallback = body["Товар"].map(lambda value: COMMODITY_TYPE_FALLBACK.get(norm_text(value)))
    body["Тип товару"] = body["Тип товару"].fillna(fallback.astype("string")).fillna("Other")

    data = body[TEXT_COLUMNS + PRICE_COLUMNS].copy()
    data.insert(0, "Дата", report_date)
    data["Середня ціна продавця"] = data[["Продавець min", "Продавець max"]].mean(axis=1, skipna=True)
    data["Середня ціна покупця"] = data[["Покупець min", "Покупець max"]].mean(axis=1, skipna=True)
    data["Джерело — файл"] = file_path.name
    data["Джерело — аркуш"] = sheet_name
    data["_rank"] = file_rank
    return data.reset_index(drop=True), meta


def sheet_fingerprint(frame: pd.DataFrame) -> str:
    """Відбиток вмісту аркуша без урахування дати — для пошуку копій."""
    if frame.empty:
        return ""
    part = frame[TEXT_COLUMNS + PRICE_COLUMNS].copy()
    for column in PRICE_COLUMNS:
        part[column] = part[column].round(4)
    part = part.astype("string").fillna("")
    part = part.sort_values(list(part.columns)).reset_index(drop=True)
    return hashlib.sha1(part.to_csv(index=False).encode("utf-8")).hexdigest()


def file_sha256(file_path: Path) -> str:
    digest = hashlib.sha256()
    with open(file_path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ===========================================================================
# Накопичена база (CSV)
# ===========================================================================


def load_database() -> pd.DataFrame:
    if not DB_FILE.exists():
        return pd.DataFrame(columns=DB_COLUMNS)
    db = pd.read_csv(DB_FILE, dtype=str, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
    for column in DB_COLUMNS:
        if column not in db.columns:
            db[column] = pd.NA
    db["Дата"] = pd.to_datetime(db["Дата"], format="%Y-%m-%d", errors="coerce")
    db["_rank"] = pd.to_datetime(db["_rank"], format="%Y-%m-%d", errors="coerce").fillna(pd.Timestamp.min)
    for column in TEXT_COLUMNS + ["Джерело — файл", "Джерело — аркуш"]:
        db[column] = clean_text(db[column])
    for column in ALL_VALUE_COLUMNS:
        db[column] = pd.to_numeric(db[column], errors="coerce")
    return db[DB_COLUMNS]


def save_database(data: pd.DataFrame) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    out = data[DB_COLUMNS].copy()
    out["Дата"] = out["Дата"].dt.strftime("%Y-%m-%d")
    out["_rank"] = out["_rank"].map(
        lambda value: "" if pd.isna(value) or value == pd.Timestamp.min else value.strftime("%Y-%m-%d")
    )
    out.to_csv(DB_FILE, index=False, encoding="utf-8-sig", float_format="%.4f")


def load_manifest() -> pd.DataFrame:
    columns = ["Файл", "SHA256", "Оброблено", "Рядків"]
    if not MANIFEST_FILE.exists():
        return pd.DataFrame(columns=columns)
    return pd.read_csv(MANIFEST_FILE, dtype=str, encoding="utf-8-sig")


# ===========================================================================
# Збір даних
# ===========================================================================


def deduplicate(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.sort_values(by=["Дата", "_rank", "Джерело — файл"], kind="stable")
    frame = frame.drop_duplicates(subset=KEY_COLUMNS, keep="last")
    return frame.sort_values(
        by=["Дата", "Тип товару", "Товар", "Країна", "Порт", "Базис"],
        na_position="last",
        kind="stable",
    ).reset_index(drop=True)


def collect_data(full_rebuild: bool) -> tuple[pd.DataFrame, pd.DataFrame, list[dict[str, object]]]:
    files = sorted(INPUT_DIR.glob("*.xlsx"))
    files = [path for path in files if not path.name.startswith("~$")]

    database = pd.DataFrame(columns=DB_COLUMNS) if full_rebuild else load_database()
    manifest = pd.DataFrame(columns=["Файл", "SHA256", "Оброблено", "Рядків"]) if full_rebuild else load_manifest()
    known = set(zip(manifest["Файл"], manifest["SHA256"]))

    if not files and database.empty:
        raise FileNotFoundError("У папці input немає Excel-файлів, і накопичена база порожня")

    frames: list[pd.DataFrame] = []
    log_rows: list[dict[str, object]] = []
    issues: list[dict[str, object]] = []
    new_manifest_rows: list[dict[str, object]] = []

    # Відбитки вже відомих дат (з бази) — щоб ловити копії між запусками.
    fingerprints: dict[str, set[pd.Timestamp]] = {}
    if not database.empty:
        for date, group in database.groupby("Дата"):
            fingerprints.setdefault(sheet_fingerprint(group), set()).add(pd.Timestamp(date))

    for file_path in files:
        sha = file_sha256(file_path)
        if (file_path.name, sha) in known:
            log_rows.append({
                "Файл": file_path.name, "Аркуш": "—", "Статус": "Вже в базі",
                "Рядків прочитано": 0, "Повідомлення": "Файл не змінювався — дані взято з data/prices_db.csv",
            })
            continue

        excel_file = pd.ExcelFile(file_path, engine="openpyxl")
        rank = workbook_rank(excel_file.sheet_names)
        matching = [sheet for sheet in excel_file.sheet_names if parse_sheet_date(sheet) is not None]

        if not matching:
            log_rows.append({
                "Файл": file_path.name, "Аркуш": "—", "Статус": "Пропущено",
                "Рядків прочитано": 0, "Повідомлення": "Не знайдено аркушів *_min_max",
            })
            issues.append({"Рівень": "Попередження", "Розділ": "Вхідні файли", "Об'єкт": file_path.name,
                           "Повідомлення": "Не знайдено аркушів *_min_max"})
            continue

        file_has_errors = False
        file_rows = 0
        for sheet_name in matching:
            report_date = parse_sheet_date(sheet_name)
            try:
                frame, meta = read_price_sheet(file_path, sheet_name, report_date, rank)
            except Exception as error:  # noqa: BLE001 — помилку фіксуємо і рахуємо в кінці
                file_has_errors = True
                log_rows.append({
                    "Файл": file_path.name, "Аркуш": sheet_name, "Статус": "Помилка",
                    "Рядків прочитано": 0, "Повідомлення": str(error),
                })
                issues.append({"Рівень": "Помилка", "Розділ": "Вхідні файли",
                               "Об'єкт": f"{file_path.name} / {sheet_name}", "Повідомлення": str(error)})
                continue

            # --- Перевірка 1: дата всередині аркуша має збігатися з назвою аркуша
            inner_dates: list[pd.Timestamp] = meta["inner_dates"]  # type: ignore[assignment]
            if inner_dates and report_date.normalize() not in inner_dates:
                inner_text = ", ".join(date.strftime("%d.%m.%Y") for date in inner_dates)
                message = (
                    f"Назва аркуша — {report_date:%d.%m.%Y}, але в колонці Date стоїть {inner_text}. "
                    "Схоже на копію попереднього дня — аркуш НЕ включено в базу."
                )
                log_rows.append({
                    "Файл": file_path.name, "Аркуш": sheet_name, "Статус": "Пропущено (копія)",
                    "Рядків прочитано": len(frame), "Повідомлення": message,
                })
                issues.append({"Рівень": "Попередження", "Розділ": "Копії аркушів",
                               "Об'єкт": f"{file_path.name} / {sheet_name}", "Повідомлення": message})
                # Файл не позначаємо як оброблений — попередження показуватиметься,
                # доки файл не виправлять або не приберуть з input/.
                file_has_errors = True
                continue

            # --- Перевірка 2: вміст ідентичний іншій даті
            fingerprint = sheet_fingerprint(frame)
            same_as = sorted(date for date in fingerprints.get(fingerprint, set()) if date != report_date)
            message = ""
            if same_as:
                message = (
                    "Ціни повністю ідентичні аркушу за "
                    + ", ".join(date.strftime("%d.%m.%Y") for date in same_as)
                    + ". Перевірте, чи це не копія."
                )
                issues.append({"Рівень": "Попередження", "Розділ": "Копії аркушів",
                               "Об'єкт": f"{file_path.name} / {sheet_name}", "Повідомлення": message})
            fingerprints.setdefault(fingerprint, set()).add(report_date)

            missing_optional = meta["missing_optional"]
            if missing_optional:
                note = "Немає колонок: " + ", ".join(missing_optional)  # type: ignore[arg-type]
                message = f"{message} {note}".strip()

            frames.append(frame)
            file_rows += len(frame)
            log_rows.append({
                "Файл": file_path.name, "Аркуш": sheet_name, "Статус": "Оброблено",
                "Рядків прочитано": len(frame), "Повідомлення": message,
            })

        if not file_has_errors:  # файли з помилками/копіями перечитуються щоразу
            new_manifest_rows.append({
                "Файл": file_path.name, "SHA256": sha,
                "Оброблено": datetime.now().strftime("%Y-%m-%d %H:%M"), "Рядків": file_rows,
            })

    new_rows = sum(len(frame) for frame in frames)
    parts = [frame for frame in [database] + frames if not frame.empty]
    if not parts:
        raise RuntimeError("Не вдалося прочитати жодної таблиці з цінами")
    combined = pd.concat(parts, ignore_index=True)
    before = len(combined)
    combined = deduplicate(combined)

    log_rows.append({
        "Файл": "УСІ ФАЙЛИ", "Аркуш": "—", "Статус": "Підсумок",
        "Рядків прочитано": new_rows,
        "Повідомлення": (
            f"Рядків із бази: {len(database)}; нових рядків: {new_rows}; "
            f"унікальних рядків після об'єднання: {len(combined)}; "
            f"дублікатів вилучено: {before - len(combined)}"
            + ("; режим: повна перебудова" if full_rebuild else "")
        ),
    })

    # Оновлений маніфест: старі записи + нові (за назвою файлу перезаписуються).
    if new_manifest_rows:
        updated = pd.DataFrame(new_manifest_rows)
        manifest = manifest[~manifest["Файл"].isin(updated["Файл"])]
        manifest = pd.concat([manifest, updated], ignore_index=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    save_database(combined)
    manifest.sort_values("Файл").to_csv(MANIFEST_FILE, index=False, encoding="utf-8-sig")

    combined = combined.drop(columns=["_rank"])
    return combined, pd.DataFrame(log_rows), issues


# ===========================================================================
# Вибір серій для аналітики: найближчий місяць поставки замість зашитого "July"
# ===========================================================================


def parse_month(value: object) -> int | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().lower()
    match = re.search(r"[a-zа-яёіїєґ]{3,}", text)
    if not match:
        number = re.fullmatch(r"(\d{1,2})", text)
        return int(number.group(1)) if number and 1 <= int(number.group(1)) <= 12 else None
    return MONTHS.get(match.group(0)[:3])


def parse_season(value: object) -> tuple[int, int] | None:
    if value is None or pd.isna(value):
        return None
    match = re.search(r"(\d{4})\s*[/\-]\s*(\d{2,4})", str(value))
    if not match:
        return None
    first = int(match.group(1))
    second = int(match.group(2))
    if second < 100:
        second = (first // 100) * 100 + second
    return first, second


def delivery_date(month: object, season: object, report_date: pd.Timestamp) -> pd.Timestamp | None:
    """Місяць поставки → конкретна дата (1-ше число місяця).

    Рік беремо з сезону ("2025/26" → 2025 або 2026) — той, що ближчий до дати звіту.
    Так коректно працює і для пшениці (сезон липень–червень), і для кукурудзи
    (сезон жовтень–вересень): "August 2025/26" у серпні 2026 → серпень 2026.
    """
    month_number = parse_month(month)
    if month_number is None:
        return None
    base = pd.Timestamp(report_date.year, report_date.month, 1)
    parsed = parse_season(season)
    years = list(parsed) if parsed else [report_date.year - 1, report_date.year, report_date.year + 1]
    candidates = [pd.Timestamp(year, month_number, 1) for year in years]
    return min(candidates, key=lambda date: abs((date - base).days))


def add_selection_columns(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    combos = data[["Місяць поставки", "Сезон", "Дата"]].drop_duplicates()
    lookup = {
        (row["Місяць поставки"], row["Сезон"], row["Дата"]): delivery_date(
            row["Місяць поставки"], row["Сезон"], pd.Timestamp(row["Дата"])
        )
        for _, row in combos.iterrows()
    }
    data["_delivery"] = [
        lookup.get((month, season, date))
        for month, season, date in zip(data["Місяць поставки"], data["Сезон"], data["Дата"])
    ]
    data["_delivery"] = pd.to_datetime(data["_delivery"])
    data["_delivery_label"] = (
        data["Місяць поставки"].fillna("—").astype(str) + " " + data["Сезон"].fillna("").astype(str)
    ).str.strip()
    for column in ["Товар", "Якість", "Країна", "Порт", "Базис", "Валюта"]:
        data[f"_n_{column}"] = data[column].map(
            lambda value, column=column: pd.NA if pd.isna(value) else norm_for(column, value)
        )
    return data


def match_rows(data: pd.DataFrame, criteria: dict[str, object]) -> pd.DataFrame:
    """Рядки, що відповідають критеріям, + пріоритет синоніма (_alias)."""
    mask = pd.Series(True, index=data.index)
    alias_rank = pd.Series(0, index=data.index)
    for column, value in criteria.items():
        normalized = data[f"_n_{column}"]
        if value is None:
            mask &= normalized.isna()
        elif isinstance(value, tuple):
            options = [norm_for(column, option) for option in value]
            mask &= normalized.isin(options)
            rank_map = {option: index for index, option in enumerate(options)}
            alias_rank = alias_rank + normalized.map(rank_map).fillna(0)
        else:
            mask &= normalized.eq(norm_for(column, value)).fillna(False)
    selected = data.loc[mask].copy()
    selected["_alias"] = alias_rank[mask]
    return selected


def front_key(delivery: pd.Timestamp | None, report_date: pd.Timestamp) -> tuple[int, int]:
    """Чим менше — тим «ближчий» контракт. Минулі місяці — в кінець."""
    base = pd.Timestamp(report_date.year, report_date.month, 1)
    if delivery is None or pd.isna(delivery):
        return (0, 0)  # без місяця = спот
    days = (delivery - base).days
    return (0, days) if days >= 0 else (1, -days)


class Selection:
    """Результат вибору для однієї серії: значення та підписи поставки по датах."""

    def __init__(self, count: int) -> None:
        self.values: list[float | None] = [None] * count
        self.labels: list[str | None] = [None] * count

    @property
    def last_label(self) -> str:
        for label in reversed(self.labels):
            if label:
                return label
        return "—"

    @property
    def filled(self) -> int:
        return sum(value is not None for value in self.values)


def pick_row(candidates: pd.DataFrame, report_date: pd.Timestamp, previous_label: str | None) -> pd.Series:
    ordered = candidates.assign(
        _front=[front_key(value, report_date) for value in candidates["_delivery"]],
        _same=[0 if label == previous_label else 1 for label in candidates["_delivery_label"]],
    ).sort_values(["_alias", "_front", "_same", "Сезон"], kind="stable")
    return ordered.iloc[0]


def select_group(
    data: pd.DataFrame,
    criteria_list: list[dict[str, object]],
    dates: list[pd.Timestamp],
    value_column: str = "Середня ціна продавця",
) -> list[Selection]:
    """Вибір найближчого місяця поставки, спільного для групи серій.

    Для спредів/премій важливо порівнювати однаковий місяць поставки, тому для
    кожної дати шукаємо місяць, який є в першої (опорної) серії і покриває
    найбільше серій групи. Серія без цього місяця на дату лишається порожньою.
    """
    matched = [match_rows(data, criteria) for criteria in criteria_list]
    matched = [frame[frame[value_column].notna()] for frame in matched]
    selections = [Selection(len(dates)) for _ in criteria_list]

    for date_index, date in enumerate(dates):
        per_series = [frame[frame["Дата"] == date] for frame in matched]
        if all(frame.empty for frame in per_series):
            continue

        anchor = per_series[0] if not per_series[0].empty else pd.concat(per_series)
        options = {
            (None if pd.isna(value) else pd.Timestamp(value)) for value in anchor["_delivery"]
        }

        def coverage(option: pd.Timestamp | None) -> int:
            count = 0
            for frame in per_series:
                if option is None:
                    count += int(frame["_delivery"].isna().any())
                else:
                    count += int((frame["_delivery"] == option).any())
            return count

        best = sorted(options, key=lambda option: (-coverage(option), front_key(option, date)))[0]

        for series_index, frame in enumerate(per_series):
            candidates = frame[frame["_delivery"].isna()] if best is None else frame[frame["_delivery"] == best]
            if candidates.empty:
                continue
            previous = selections[series_index].labels[date_index - 1] if date_index else None
            row = pick_row(candidates, date, previous)
            selections[series_index].values[date_index] = float(row[value_column])
            selections[series_index].labels[date_index] = row["_delivery_label"]

    return selections


def pair_label(left: Selection, right: Selection) -> str:
    """Місяць поставки на останню дату, де є обидві ціни (для рядка спреду)."""
    for index in range(len(left.values) - 1, -1, -1):
        if left.values[index] is not None and right.values[index] is not None:
            return left.labels[index] or "—"
    return "—"


def select_series(data: pd.DataFrame, criteria: dict[str, object], dates: list[pd.Timestamp]) -> Selection:
    return select_group(data, [criteria], dates)[0]


def coverage_issue(sheet: str, name: str, selection: Selection, dates: list[pd.Timestamp]) -> dict[str, object]:
    total = len(dates)
    filled = selection.filled
    last_index = max((index for index, value in enumerate(selection.values) if value is not None), default=None)
    last_text = dates[last_index].strftime("%d.%m.%Y") if last_index is not None else "—"
    if filled == 0:
        level = "Попередження"
        message = "Жодної дати з даними — перевірте назви товару/порту/якості у специфікації"
    elif selection.values[-1] is None:
        level = "Увага"
        message = f"Заповнено {filled} із {total} дат; на останню дату даних немає (остання з даними: {last_text})"
    else:
        level = "ОК"
        message = f"Заповнено {filled} із {total} дат; поставка на останню дату: {selection.last_label}"
    return {"Рівень": level, "Розділ": f"Покриття: {sheet}", "Об'єкт": name, "Повідомлення": message}


# ===========================================================================
# Таблиці
# ===========================================================================


def date_list(data: pd.DataFrame) -> list[pd.Timestamp]:
    return sorted(pd.Timestamp(value) for value in data["Дата"].dropna().unique())


def date_labels(dates: list[pd.Timestamp]) -> list[str]:
    return [date.strftime("%d.%m.%Y") for date in dates]
def build_wide_table(data: pd.DataFrame, include_category: bool) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame()

    index_columns = SERIES_COLUMNS.copy()
    if not include_category:
        index_columns.remove("Тип товару")

    melted = data.melt(
        id_vars=["Дата"] + index_columns,
        value_vars=ALL_VALUE_COLUMNS,
        var_name="_metric",
        value_name="_value",
    )
    melted = melted[melted["_value"].notna()].copy()
    for column in index_columns:
        melted[column] = melted[column].fillna("—")
    melted["Показник"] = melted["_metric"].map(METRIC_NAMES)
    melted["_metric_order"] = melted["_metric"].map(METRIC_ORDER)

    wide = melted.pivot_table(
        index=index_columns + ["Показник", "_metric_order"],
        columns="Дата",
        values="_value",
        aggfunc="last",
        dropna=True,
        observed=True,
    ).reset_index()

    if include_category:
        wide["_category_order"] = wide["Тип товару"].map(CATEGORY_ORDER).fillna(99)
        sort_columns = ["_category_order"] + index_columns[1:] + ["_metric_order"]
    else:
        sort_columns = index_columns + ["_metric_order"]

    wide = wide.sort_values(sort_columns, na_position="last", kind="stable")

    if include_category:
        wide["Тип товару"] = wide["Тип товару"].map(CATEGORY_NAMES).fillna(wide["Тип товару"])
        wide = wide.rename(columns={"Тип товару": "Категорія"})

    wide = wide.drop(columns=["_metric_order", "_category_order"], errors="ignore")

    renamed_dates = {
        column: column.strftime("%d.%m.%Y")
        for column in wide.columns
        if isinstance(column, pd.Timestamp)
    }
    wide = wide.rename(columns=renamed_dates)

    descriptor_order = (
        ["Категорія"] if include_category else []
    ) + [
        "Товар",
        "Якість",
        "Країна",
        "Порт",
        "Базис",
        "Валюта",
        "Місяць поставки",
        "Сезон",
        "Показник",
    ]
    date_columns = sorted(
        [column for column in wide.columns if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", str(column))],
        key=lambda value: pd.to_datetime(value, format="%d.%m.%Y"),
    )
    return wide[descriptor_order + date_columns]


def build_compact_market_table(
    data: pd.DataFrame,
    dates: list[pd.Timestamp],
    mode: str,
) -> pd.DataFrame:
    if mode == "export":
        mask = (
            data["Країна"].eq("Ukraine")
            & data["Валюта"].eq("USD")
            & data["Базис"].isin(["CPT", "FOB", "DAP"])
            & data["Порт"].notna()
        )
    elif mode == "internal":
        mask = (
            data["Країна"].eq("Ukraine")
            & data["Валюта"].eq("UAH")
            & data["Базис"].isin(["CPT", "EXW"])
        )
    else:
        raise ValueError(f"Невідомий режим: {mode}")

    columns = [
        "Тип товару",
        "Товар",
        "Якість",
        "Порт",
        "Базис",
        "Валюта",
        "Місяць поставки",
        "Сезон",
    ]
    source = data.loc[mask, columns + ["Дата", "Середня ціна продавця"]].copy()
    source = source[source["Середня ціна продавця"].notna()]
    for column in columns:
        source[column] = source[column].fillna("—")

    wide = source.pivot_table(
        index=columns,
        columns="Дата",
        values="Середня ціна продавця",
        aggfunc="last",
        observed=True,
    ).reset_index()

    wide["_category_order"] = wide["Тип товару"].map(CATEGORY_ORDER).fillna(99)
    wide = wide.sort_values(
        [
            "_category_order",
            "Товар",
            "Якість",
            "Порт",
            "Базис",
            "Місяць поставки",
            "Сезон",
        ],
        kind="stable",
    )
    wide["Категорія"] = wide["Тип товару"].map(CATEGORY_NAMES).fillna(wide["Тип товару"])
    wide = wide.rename(columns={"Порт": "Порт / ринок"})
    wide["Показник"] = "Середня ціна продавця"

    for date in dates:
        if date not in wide.columns:
            wide[date] = pd.NA

    wide = wide.rename(columns={date: date.strftime("%d.%m.%Y") for date in dates})
    final_columns = [
        "Категорія",
        "Товар",
        "Якість",
        "Порт / ринок",
        "Базис",
        "Валюта",
        "Місяць поставки",
        "Сезон",
        "Показник",
    ] + date_labels(dates)
    return wide[final_columns].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Специфікації аналітичних аркушів.
# Місяць поставки і сезон НЕ задаються: на кожну дату береться найближчий
# доступний місяць поставки (для спредів — спільний для обох серій).
# ---------------------------------------------------------------------------


def spec(
    commodity: str,
    quality: object,
    port: object,
    basis: str,
    currency: str = "USD",
    country: str = "Ukraine",
) -> dict[str, object]:
    return {
        "Товар": commodity,
        "Якість": quality,
        "Країна": country,
        "Порт": port,
        "Базис": basis,
        "Валюта": currency,
    }


WHEAT_125 = ("Milling - 12.5%", "Milling 12.5%", "12.5%")
WHEAT_115 = ("Milling - 11.5%", "Milling 11.5%", "11.5%")
FEED = ("Feed",)

DeliveryNotes = dict[int, list[str | None]]


def build_key_indicators(
    data: pd.DataFrame,
    dates: list[pd.Timestamp],
    issues: list[dict[str, object]],
) -> tuple[pd.DataFrame, DeliveryNotes]:
    specs = [
        ("Зернові — експорт", "Пшениця 12,5% — Одеса CPT", "Україна, Одеса", "CPT", "USD/т", spec("Wheat", WHEAT_125, ODESA, "CPT")),
        ("Зернові — експорт", "Пшениця 12,5% — Одеса FOB", "Україна, Одеса", "FOB", "USD/т", spec("Wheat", WHEAT_125, ODESA, "FOB")),
        ("Зернові — експорт", "Пшениця 11,5% — Одеса CPT", "Україна, Одеса", "CPT", "USD/т", spec("Wheat", WHEAT_115, ODESA, "CPT")),
        ("Зернові — експорт", "Пшениця фуражна — Одеса CPT", "Україна, Одеса", "CPT", "USD/т", spec("Wheat", FEED, ODESA, "CPT")),
        ("Зернові — експорт", "Кукурудза — Одеса CPT", "Україна, Одеса", "CPT", "USD/т", spec("Corn", FEED, ODESA, "CPT")),
        ("Зернові — експорт", "Кукурудза — Одеса FOB", "Україна, Одеса", "FOB", "USD/т", spec("Corn", FEED, ODESA, "FOB")),
        ("Зернові — експорт", "Ячмінь — Одеса CPT", "Україна, Одеса", "CPT", "USD/т", spec("Barley", FEED, ODESA, "CPT")),
        ("Олійні — експорт", "Ріпак — чорноморські порти CPT", "Україна, Чорне море", "CPT", "USD/т", spec("Rapeseed", None, BLACK_SEA_PORTS, "CPT")),
        ("Олійні — експорт", "Соя — чорноморські порти CPT", "Україна, Чорне море", "CPT", "USD/т", spec("Soybeans", None, BLACK_SEA_PORTS, "CPT")),
        ("Олії та шроти — експорт", "Соняшникова олія — порти FOB", "Україна, Чорне море", "FOB", "USD/т", spec("Sunflower oil", ("Crude",), BLACK_SEA_PORTS, "FOB")),
        ("Олії та шроти — експорт", "Соняшниковий шрот — західний кордон DAP", "Україна, західний кордон", "DAP", "USD/т", spec("Sunflower meal", None, WESTERN_BORDER, "DAP")),
        ("Внутрішній ринок", "Пшениця 12,5% — CPT", "Україна", "CPT", "UAH/т", spec("Wheat", WHEAT_125, None, "CPT", "UAH")),
        ("Внутрішній ринок", "Кукурудза — CPT", "Україна", "CPT", "UAH/т", spec("Corn", FEED, None, "CPT", "UAH")),
        ("Внутрішній ринок", "Соняшник — CPT", "Україна", "CPT", "UAH/т", spec("Sunflower seed", None, None, "CPT", "UAH")),
        ("Внутрішній ринок", "Соя — CPT", "Україна", "CPT", "UAH/т", spec("Soybeans", None, None, "CPT", "UAH")),
        ("Внутрішній ринок", "Ріпак non-GMO — CPT", "Україна", "CPT", "UAH/т", spec("Rapeseed", ("non-GMO", "non GMO"), None, "CPT", "UAH")),
    ]

    rows: list[dict[str, object]] = []
    notes: DeliveryNotes = {}
    labels = date_labels(dates)
    for group, indicator, market, basis, unit, criteria in specs:
        selection = select_series(data, criteria, dates)
        issues.append(coverage_issue("Ключові індикатори", indicator, selection, dates))
        notes[len(rows)] = selection.labels
        row: dict[str, object] = {
            "Група": group,
            "Індикатор": indicator,
            "Ринок": market,
            "Базис": basis,
            "Одиниця": unit,
            "Поставка (остання дата)": selection.last_label,
            "Остання ціна": pd.NA,
            "Δ день": pd.NA,
            "Δ період": pd.NA,
            "Δ період, %": pd.NA,
        }
        row.update(dict(zip(labels, selection.values)))
        rows.append(row)
    return pd.DataFrame(rows), notes


def build_port_spreads(
    data: pd.DataFrame,
    dates: list[pd.Timestamp],
    issues: list[dict[str, object]],
) -> tuple[pd.DataFrame, list[tuple[int, int, int]], DeliveryNotes]:
    specs = [
        ("Одеса CPT проти Рені CPT", "Пшениця 12,5%", spec("Wheat", WHEAT_125, ODESA, "CPT"), spec("Wheat", WHEAT_125, RENI, "CPT"), "Одеса CPT", "Рені CPT"),
        ("Одеса CPT проти Рені CPT", "Пшениця 11,5%", spec("Wheat", WHEAT_115, ODESA, "CPT"), spec("Wheat", WHEAT_115, RENI, "CPT"), "Одеса CPT", "Рені CPT"),
        ("Одеса CPT проти Рені CPT", "Пшениця фуражна", spec("Wheat", FEED, ODESA, "CPT"), spec("Wheat", FEED, RENI, "CPT"), "Одеса CPT", "Рені CPT"),
        ("Одеса CPT проти Рені CPT", "Кукурудза", spec("Corn", FEED, ODESA, "CPT"), spec("Corn", FEED, RENI, "CPT"), "Одеса CPT", "Рені CPT"),
        ("Одеса CPT проти Рені CPT", "Ячмінь", spec("Barley", FEED, ODESA, "CPT"), spec("Barley", FEED, RENI, "CPT"), "Одеса CPT", "Рені CPT"),
        ("FOB проти CPT в Одесі", "Пшениця 12,5%", spec("Wheat", WHEAT_125, ODESA, "FOB"), spec("Wheat", WHEAT_125, ODESA, "CPT"), "Одеса FOB", "Одеса CPT"),
        ("FOB проти CPT в Одесі", "Пшениця 11,5%", spec("Wheat", WHEAT_115, ODESA, "FOB"), spec("Wheat", WHEAT_115, ODESA, "CPT"), "Одеса FOB", "Одеса CPT"),
        ("FOB проти CPT в Одесі", "Пшениця фуражна", spec("Wheat", FEED, ODESA, "FOB"), spec("Wheat", FEED, ODESA, "CPT"), "Одеса FOB", "Одеса CPT"),
        ("FOB проти CPT в Одесі", "Кукурудза", spec("Corn", FEED, ODESA, "FOB"), spec("Corn", FEED, ODESA, "CPT"), "Одеса FOB", "Одеса CPT"),
        ("FOB проти CPT в Одесі", "Ячмінь", spec("Barley", FEED, ODESA, "FOB"), spec("Barley", FEED, ODESA, "CPT"), "Одеса FOB", "Одеса CPT"),
        ("FOB проти CPT — чорноморські порти", "Ріпак", spec("Rapeseed", None, BLACK_SEA_PORTS, "FOB"), spec("Rapeseed", None, BLACK_SEA_PORTS, "CPT"), "Порти FOB", "Порти CPT"),
        ("FOB проти CPT — чорноморські порти", "Соняшникова олія", spec("Sunflower oil", ("Crude",), BLACK_SEA_PORTS, "FOB"), spec("Sunflower oil", ("Crude",), BLACK_SEA_PORTS, "CPT"), "Порти FOB", "Порти CPT"),
    ]

    labels = date_labels(dates)
    rows: list[dict[str, object]] = []
    formulas: list[tuple[int, int, int]] = []
    notes: DeliveryNotes = {}
    for block, commodity, left_criteria, right_criteria, left_label, right_label in specs:
        left_sel, right_sel = select_group(data, [left_criteria, right_criteria], dates)
        name = f"{block}: {commodity}"
        issues.append(coverage_issue("Портові спреди", f"{name} — {left_label}", left_sel, dates))
        issues.append(coverage_issue("Портові спреди", f"{name} — {right_label}", right_sel, dates))

        indexes = []
        for selection, label, explanation in [
            (left_sel, left_label, "Перша серія"),
            (right_sel, right_label, "Друга серія"),
        ]:
            indexes.append(len(rows))
            notes[len(rows)] = selection.labels
            row = {
                "Блок": block,
                "Товар": commodity,
                "Тип рядка": "Ціна",
                "Серія / формула": label,
                "Пояснення": explanation,
                "Одиниця": "USD/т",
                "Поставка (остання дата)": selection.last_label,
            }
            row.update(dict(zip(labels, selection.values)))
            rows.append(row)

        spread_index = len(rows)
        spread = {
            "Блок": block,
            "Товар": commodity,
            "Тип рядка": "Спред",
            "Серія / формула": f"{left_label} – {right_label}",
            "Пояснення": "Перша серія мінус друга (однаковий місяць поставки)",
            "Одиниця": "USD/т",
            "Поставка (остання дата)": pair_label(left_sel, right_sel),
        }
        spread.update({label: pd.NA for label in labels})
        rows.append(spread)
        formulas.append((spread_index, indexes[0], indexes[1]))

    return pd.DataFrame(rows), formulas, notes


def build_wheat_premiums(
    data: pd.DataFrame,
    dates: list[pd.Timestamp],
    issues: list[dict[str, object]],
) -> tuple[pd.DataFrame, list[tuple[int, int, int]], DeliveryNotes]:
    markets = [
        ("Одеса CPT", "CPT", ODESA),
        ("Рені CPT", "CPT", RENI),
        ("Одеса FOB", "FOB", ODESA),
    ]
    quality_map = {
        "Фуражна": FEED,
        "11,5%": WHEAT_115,
        "12,5%": WHEAT_125,
    }
    labels = date_labels(dates)
    rows: list[dict[str, object]] = []
    formulas: list[tuple[int, int, int]] = []
    notes: DeliveryNotes = {}

    for market, basis, port in markets:
        qualities = ["Фуражна", "11,5%", "12,5%"]
        selections = select_group(
            data, [spec("Wheat", quality_map[quality], port, basis) for quality in qualities], dates
        )
        raw_indexes: dict[str, int] = {}
        for quality_label, selection in zip(qualities, selections):
            issues.append(coverage_issue("Пшениця — премії якості", f"{market}: {quality_label}", selection, dates))
            raw_indexes[quality_label] = len(rows)
            notes[len(rows)] = selection.labels
            row = {
                "Ринок": market,
                "Товар": "Пшениця",
                "Тип рядка": "Ціна",
                "Якість / премія": quality_label,
                "Пояснення": "Середня ціна продавця",
                "Базис": basis,
                "Одиниця": "USD/т",
                "Поставка (остання дата)": selection.last_label,
            }
            row.update(dict(zip(labels, selection.values)))
            rows.append(row)

        premium_specs = [
            ("11,5% – фуражна", "Премія 11,5% до фуражної", "11,5%", "Фуражна"),
            ("12,5% – 11,5%", "Премія 12,5% до 11,5%", "12,5%", "11,5%"),
            ("12,5% – фуражна", "Премія 12,5% до фуражної", "12,5%", "Фуражна"),
        ]
        for label, explanation, left_quality, right_quality in premium_specs:
            spread_index = len(rows)
            row = {
                "Ринок": market,
                "Товар": "Пшениця",
                "Тип рядка": "Премія",
                "Якість / премія": label,
                "Пояснення": explanation,
                "Базис": basis,
                "Одиниця": "USD/т",
                "Поставка (остання дата)": pair_label(
                    selections[qualities.index(left_quality)], selections[qualities.index(right_quality)]
                ),
            }
            row.update({date: pd.NA for date in labels})
            rows.append(row)
            formulas.append((spread_index, raw_indexes[left_quality], raw_indexes[right_quality]))

    return pd.DataFrame(rows), formulas, notes


def build_competitors(
    data: pd.DataFrame,
    dates: list[pd.Timestamp],
    issues: list[dict[str, object]],
) -> tuple[pd.DataFrame, list[tuple[int, int, int]], DeliveryNotes]:
    def competitor(commodity, quality, country, port):
        return spec(commodity, quality, port, "FOB", "USD", country)

    blocks = [
        {
            "block": "Пшениця 12,5% — FOB",
            "commodity": "Пшениця",
            "quality": "12,5%",
            "rows": [
                ("Україна — Одеса", competitor("Wheat", WHEAT_125, "Ukraine", ODESA)),
                ("Росія — Новоросійськ", competitor("Wheat", WHEAT_125, "Russia", NOVOROSSIYSK)),
                ("Росія — Ростов/Азов", competitor("Wheat", WHEAT_125, "Russia", ROSTOV_AZOV)),
            ],
            "spreads": [("Україна мінус Новоросійськ", 0, 1), ("Україна мінус Ростов/Азов", 0, 2)],
        },
        {
            "block": "Пшениця фуражна — FOB",
            "commodity": "Пшениця",
            "quality": "Фуражна",
            "rows": [
                ("Україна — Одеса", competitor("Wheat", FEED, "Ukraine", ODESA)),
                ("Росія — Новоросійськ", competitor("Wheat", FEED, "Russia", NOVOROSSIYSK)),
                ("ЄС — Чорне море", competitor("Wheat", FEED, "EU Black Sea", EU_BLACK_SEA_PORTS)),
            ],
            "spreads": [("Україна мінус Новоросійськ", 0, 1), ("Україна мінус ЄС Чорне море", 0, 2)],
        },
        {
            "block": "Кукурудза фуражна — FOB",
            "commodity": "Кукурудза",
            "quality": "Фуражна",
            "rows": [
                ("Україна — Одеса", competitor("Corn", FEED, "Ukraine", ODESA)),
                ("Росія — Новоросійськ", competitor("Corn", FEED, "Russia", NOVOROSSIYSK)),
                ("Росія — Ростов/Азов", competitor("Corn", FEED, "Russia", ROSTOV_AZOV)),
                ("ЄС — Чорне море", competitor("Corn", FEED, "EU Black Sea", EU_BLACK_SEA_PORTS)),
            ],
            "spreads": [("Україна мінус Новоросійськ", 0, 1), ("Україна мінус Ростов/Азов", 0, 2), ("Україна мінус ЄС Чорне море", 0, 3)],
        },
        {
            "block": "Ячмінь фуражний — FOB",
            "commodity": "Ячмінь",
            "quality": "Фуражний",
            "rows": [
                ("Україна — Одеса", competitor("Barley", FEED, "Ukraine", ODESA)),
                ("Росія — Новоросійськ", competitor("Barley", FEED, "Russia", NOVOROSSIYSK)),
                ("Росія — Ростов/Азов", competitor("Barley", FEED, "Russia", ROSTOV_AZOV)),
                ("ЄС — Чорне море", competitor("Barley", FEED, "EU Black Sea", EU_BLACK_SEA_PORTS)),
                ("Франція — Руан", competitor("Barley", FEED, "France", ("Rouen",))),
            ],
            "spreads": [("Україна мінус Новоросійськ", 0, 1), ("Україна мінус Ростов/Азов", 0, 2), ("Україна мінус ЄС Чорне море", 0, 3), ("Україна мінус Франція", 0, 4)],
        },
    ]

    labels = date_labels(dates)
    rows: list[dict[str, object]] = []
    formulas: list[tuple[int, int, int]] = []
    notes: DeliveryNotes = {}

    for block in blocks:
        selections = select_group(data, [criteria for _, criteria in block["rows"]], dates)
        raw_indexes: list[int] = []
        for (label, criteria), selection in zip(block["rows"], selections):
            issues.append(coverage_issue("Чорне море — конкуренти", f"{block['block']}: {label}", selection, dates))
            raw_indexes.append(len(rows))
            notes[len(rows)] = selection.labels
            row = {
                "Блок": block["block"],
                "Товар": block["commodity"],
                "Якість": block["quality"],
                "Тип рядка": "Ціна",
                "Ринок / спред": label,
                "Країна / пояснення": criteria["Країна"],
                "Порт": display_value(criteria["Порт"]),
                "Базис": "FOB",
                "Одиниця": "USD/т",
                "Поставка (остання дата)": selection.last_label,
            }
            row.update(dict(zip(labels, selection.values)))
            rows.append(row)

        for label, left_index, right_index in block["spreads"]:
            spread_index = len(rows)
            row = {
                "Блок": block["block"],
                "Товар": block["commodity"],
                "Якість": block["quality"],
                "Тип рядка": "Спред",
                "Ринок / спред": label,
                "Країна / пояснення": "Україна мінус конкурент",
                "Порт": "—",
                "Базис": "FOB",
                "Одиниця": "USD/т",
                "Поставка (остання дата)": pair_label(selections[left_index], selections[right_index]),
            }
            row.update({date: pd.NA for date in labels})
            rows.append(row)
            formulas.append((spread_index, raw_indexes[left_index], raw_indexes[right_index]))

    return pd.DataFrame(rows), formulas, notes


def add_delivery_comments(
    worksheet,
    notes: DeliveryNotes,
    date_start_column: int,
    data_start_row: int = 5,
) -> None:
    """Примітка в комірці, де змінився місяць поставки (перехід на наступний контракт)."""
    for row_index, labels in notes.items():
        previous: str | None = None
        for offset, label in enumerate(labels):
            if not label:
                continue
            if previous and label != previous:
                cell = worksheet.cell(row=data_start_row + row_index, column=date_start_column + offset)
                comment = Comment(
                    f"Змінився місяць поставки: {previous} → {label}. "
                    "Зміна ціни відносно попередньої дати частково пояснюється переходом контракту.",
                    "UkrAgro script",
                )
                comment.width = 260
                comment.height = 80
                cell.comment = comment
                cell.font = Font(color="C00000", bold=True, size=10)
            previous = label


def style_title_and_header(
    worksheet,
    dataframe: pd.DataFrame,
    title: str,
    note: str,
    date_start_column: int,
) -> None:
    worksheet.sheet_view.showGridLines = False
    max_column = dataframe.shape[1]
    max_row = dataframe.shape[0] + 4

    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_column)
    title_cell = worksheet.cell(row=1, column=1)
    title_cell.value = title
    title_cell.fill = PatternFill("solid", fgColor=TITLE_FILL)
    title_cell.font = Font(color=WHITE, bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    worksheet.row_dimensions[1].height = 30

    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_column)
    note_cell = worksheet.cell(row=2, column=1)
    note_cell.value = note
    note_cell.fill = PatternFill("solid", fgColor=NOTE_FILL)
    note_cell.font = Font(color=TEXT_COLOR, italic=True, size=10)
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    worksheet.row_dimensions[2].height = 34

    for column in range(1, max_column + 1):
        cell = worksheet.cell(row=4, column=column)
        cell.fill = PatternFill(
            "solid",
            fgColor=HEADER_FILL if column < date_start_column else DATE_FILL,
        )
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[4].height = 36
    worksheet.freeze_panes = worksheet.cell(row=5, column=date_start_column)
    worksheet.auto_filter.ref = f"A4:{get_column_letter(max_column)}{max_row}"


def apply_group_banding(
    worksheet,
    start_row: int,
    end_row: int,
    max_column: int,
    date_start_column: int,
    group_column: int = 1,
    special_rows: set[int] | None = None,
) -> None:
    special_rows = special_rows or set()
    thin_side = Side(style="thin", color=BORDER_COLOR)
    previous_key = object()
    group_number = -1

    for row in range(start_row, end_row + 1):
        key = worksheet.cell(row=row, column=group_column).value
        if key != previous_key:
            group_number += 1
            previous_key = key

        fill_color = GROUP_FILL if group_number % 2 == 0 else WHITE
        if row in special_rows:
            fill_color = SPREAD_FILL

        for column in range(1, max_column + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.font = Font(
                color="7F6000" if row in special_rows else TEXT_COLOR,
                bold=row in special_rows,
                size=10,
            )
            cell.alignment = Alignment(
                horizontal="right" if column >= date_start_column else "left",
                vertical="center",
                wrap_text=True,
            )
            cell.border = Border(bottom=thin_side)


def set_column_widths(worksheet, widths: list[float], date_start_column: int, max_column: int) -> None:
    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width
    for column in range(date_start_column, max_column + 1):
        worksheet.column_dimensions[get_column_letter(column)].width = 13


def style_key_indicators(
    worksheet,
    dataframe: pd.DataFrame,
    dates: list[pd.Timestamp],
) -> None:
    date_start = 11
    style_title_and_header(
        worksheet,
        dataframe,
        "Ключові індикатори UkrAgroConsult",
        "Компактна добірка ключових експортних і внутрішніх цін. На кожну дату береться найближчий "
        "доступний місяць поставки; червона комірка з приміткою — дата переходу на наступний контракт.",
        date_start,
    )

    start_row = 5
    end_row = dataframe.shape[0] + 4
    max_column = dataframe.shape[1]
    first_date_column = date_start
    last_date_column = date_start + len(dates) - 1
    previous_date_column = max(first_date_column, last_date_column - 1)
    first_letter = get_column_letter(first_date_column)
    last_letter = get_column_letter(last_date_column)
    previous_letter = get_column_letter(previous_date_column)

    for row in range(start_row, end_row + 1):
        values = f"${first_letter}{row}:${last_letter}{row}"
        last_value = f'LOOKUP(2,1/({values}<>""),{values})'
        first_value = f'INDEX({values},MATCH(TRUE,INDEX({values}<>"",0),0))'
        # Остання ціна — останнє непорожнє значення в рядку
        worksheet.cell(row=row, column=7).value = f'=IFERROR({last_value},"")'
        if len(dates) >= 2:
            worksheet.cell(row=row, column=8).value = (
                f'=IF(OR({last_letter}{row}="",{previous_letter}{row}=""),"",'
                f"{last_letter}{row}-{previous_letter}{row})"
            )
        worksheet.cell(row=row, column=9).value = f'=IFERROR({last_value}-{first_value},"")'
        worksheet.cell(row=row, column=10).value = f'=IFERROR(I{row}/{first_value},"")'

    apply_group_banding(worksheet, start_row, end_row, max_column, date_start, group_column=1)
    set_column_widths(
        worksheet,
        [20, 34, 24, 10, 11, 20, 14, 12, 12, 13],
        date_start,
        max_column,
    )

    for row in range(start_row, end_row + 1):
        unit = worksheet.cell(row=row, column=5).value
        number_format = "0" if str(unit).startswith("UAH") else "0.00"
        for column in [7, 8, 9] + list(range(date_start, max_column + 1)):
            worksheet.cell(row=row, column=column).number_format = number_format
        worksheet.cell(row=row, column=10).number_format = "0.0%"

    positive_fill = PatternFill("solid", fgColor=POSITIVE_FILL)
    negative_fill = PatternFill("solid", fgColor=NEGATIVE_FILL)
    for column in [8, 9, 10]:
        cell_range = f"{get_column_letter(column)}{start_row}:{get_column_letter(column)}{end_row}"
        worksheet.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="greaterThan", formula=["0"], fill=positive_fill),
        )
        worksheet.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="lessThan", formula=["0"], fill=negative_fill),
        )


def style_compact_market(
    worksheet,
    dataframe: pd.DataFrame,
    title: str,
    note: str,
) -> None:
    date_start = 10
    style_title_and_header(worksheet, dataframe, title, note, date_start)
    start_row = 5
    end_row = dataframe.shape[0] + 4
    max_column = dataframe.shape[1]
    apply_group_banding(worksheet, start_row, end_row, max_column, date_start, group_column=1)
    set_column_widths(
        worksheet,
        [16, 18, 18, 30, 10, 10, 16, 12, 24],
        date_start,
        max_column,
    )
    for row in range(start_row, end_row + 1):
        currency = worksheet.cell(row=row, column=6).value
        number_format = "0" if currency == "UAH" else "0.00"
        for column in range(date_start, max_column + 1):
            worksheet.cell(row=row, column=column).number_format = number_format


def write_formulas_for_comparison(
    worksheet,
    formulas: list[tuple[int, int, int]],
    date_start_column: int,
    date_count: int,
    data_start_row: int = 5,
) -> set[int]:
    special_rows: set[int] = set()
    for target_index, left_index, right_index in formulas:
        target_row = data_start_row + target_index
        left_row = data_start_row + left_index
        right_row = data_start_row + right_index
        special_rows.add(target_row)
        for column in range(date_start_column, date_start_column + date_count):
            letter = get_column_letter(column)
            worksheet.cell(row=target_row, column=column).value = (
                f'=IF(OR({letter}{left_row}="",{letter}{right_row}=""),"",'
                f'{letter}{left_row}-{letter}{right_row})'
            )
    return special_rows


def style_comparison_sheet(
    worksheet,
    dataframe: pd.DataFrame,
    title: str,
    note: str,
    date_start_column: int,
    widths: list[float],
    special_rows: set[int],
) -> None:
    style_title_and_header(worksheet, dataframe, title, note, date_start_column)
    start_row = 5
    end_row = dataframe.shape[0] + 4
    max_column = dataframe.shape[1]
    apply_group_banding(
        worksheet,
        start_row,
        end_row,
        max_column,
        date_start_column,
        group_column=1,
        special_rows=special_rows,
    )
    set_column_widths(worksheet, widths, date_start_column, max_column)
    for row in range(start_row, end_row + 1):
        for column in range(date_start_column, max_column + 1):
            worksheet.cell(row=row, column=column).number_format = "0.00"


def style_wide_sheet(
    worksheet,
    dataframe: pd.DataFrame,
    title: str,
    include_category: bool,
) -> None:
    worksheet.sheet_view.showGridLines = False

    max_column = dataframe.shape[1]
    max_row = dataframe.shape[0] + 4
    descriptor_count = 10 if include_category else 9
    first_date_column = descriptor_count + 1

    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_column)
    title_cell = worksheet.cell(row=1, column=1)
    title_cell.value = f"{title}: щоденні цінові ряди UkrAgroConsult"
    title_cell.fill = PatternFill("solid", fgColor=TITLE_FILL)
    title_cell.font = Font(color=WHITE, bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    worksheet.row_dimensions[1].height = 30

    date_columns = dataframe.columns[first_date_column - 1 :]
    series_columns = dataframe.columns[: first_date_column - 1]
    series_count = dataframe[list(series_columns[:-1])].drop_duplicates().shape[0]
    period = "—"
    if len(date_columns) > 0:
        period = f"{date_columns[0]}–{date_columns[-1]}"

    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_column)
    note_cell = worksheet.cell(row=2, column=1)
    note_cell.value = (
        "Дати розміщені по горизонталі; кожен ціновий показник — окремим рядком. "
        f"Період: {period}. Цінових рядів: {series_count}."
    )
    note_cell.fill = PatternFill("solid", fgColor=NOTE_FILL)
    note_cell.font = Font(color=TEXT_COLOR, italic=True, size=10)
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    worksheet.row_dimensions[2].height = 28

    for column in range(1, max_column + 1):
        cell = worksheet.cell(row=4, column=column)
        cell.fill = PatternFill(
            "solid",
            fgColor=HEADER_FILL if column < first_date_column else DATE_FILL,
        )
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[4].height = 34

    worksheet.freeze_panes = worksheet.cell(row=5, column=first_date_column)
    worksheet.auto_filter.ref = f"A4:{get_column_letter(max_column)}{max_row}"

    widths = ([15] if include_category else []) + [18, 18, 16, 24, 10, 10, 16, 12, 22]
    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width
    for column in range(first_date_column, max_column + 1):
        worksheet.column_dimensions[get_column_letter(column)].width = 13

    metric_column = descriptor_count
    thin_side = Side(style="thin", color=BORDER_COLOR)
    group_columns = list(range(1, metric_column))

    previous_key = None
    group_number = -1
    for row in range(5, max_row + 1):
        key = tuple(worksheet.cell(row=row, column=column).value for column in group_columns)
        if key != previous_key:
            group_number += 1
            previous_key = key

        fill = PatternFill("solid", fgColor=GROUP_FILL if group_number % 2 == 0 else WHITE)
        for column in range(1, max_column + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.fill = fill
            cell.font = Font(
                color="17365D" if column == metric_column else TEXT_COLOR,
                bold=column == metric_column,
                size=10,
            )
            cell.alignment = Alignment(
                horizontal="right" if column >= first_date_column else "left",
                vertical="center",
                wrap_text=column < first_date_column,
            )
            cell.border = Border(bottom=thin_side)
            if column >= first_date_column and isinstance(cell.value, (int, float)):
                cell.number_format = "0.00"


def style_database_sheet(worksheet, dataframe: pd.DataFrame) -> None:
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "H2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 34

    widths = {
        "A": 12,
        "B": 18,
        "C": 18,
        "D": 16,
        "E": 24,
        "F": 10,
        "G": 10,
        "H": 14,
        "I": 14,
        "J": 14,
        "K": 14,
        "L": 16,
        "M": 12,
        "N": 18,
        "O": 16,
        "P": 16,
        "Q": 34,
        "R": 24,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(color=TEXT_COLOR, size=9)
            cell.alignment = Alignment(vertical="center", wrap_text=False)

    for cell in worksheet["A"][1:]:
        if cell.value is not None:
            cell.number_format = "dd.mm.yyyy"

    for column in range(8, 17):
        for cell in worksheet.iter_cols(min_col=column, max_col=column, min_row=2):
            for value_cell in cell:
                if isinstance(value_cell.value, (int, float)):
                    value_cell.number_format = "0.00"


def style_log_sheet(worksheet) -> None:
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(color=TEXT_COLOR, size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = [34, 24, 14, 16, 60]
    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width


def style_issues_sheet(worksheet) -> None:
    style_log_sheet(worksheet)
    for index, width in enumerate([18, 34, 52, 90], start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width
    fills = {"Помилка": "F8CBAD", "Попередження": "FFE699", "Увага": "FFF2CC", "ОК": "E2F0D9"}
    for row in worksheet.iter_rows(min_row=2):
        color = fills.get(row[0].value)
        if color:
            row[0].fill = PatternFill("solid", fgColor=color)
            row[0].font = Font(color=TEXT_COLOR, bold=True, size=10)


def create_navigation_sheet(workbook, dates: list[pd.Timestamp]) -> None:
    """Створює перший довідковий аркуш із внутрішніми посиланнями."""

    if "Навігація" in workbook.sheetnames:
        del workbook["Навігація"]

    worksheet = workbook.create_sheet("Навігація", 0)
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A4"

    worksheet.merge_cells("A1:H1")
    title_cell = worksheet["A1"]
    title_cell.value = "UKRAGRO — НАВІГАЦІЯ"
    title_cell.fill = PatternFill("solid", fgColor=TITLE_FILL)
    title_cell.font = Font(color=WHITE, bold=True, size=18)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    worksheet.row_dimensions[1].height = 32

    worksheet.merge_cells("A2:H2")
    subtitle_cell = worksheet["A2"]
    subtitle_cell.value = (
        "Швидкий перехід до основних таблиць і аналітичних зрізів. "
        "Дати в робочих аркушах розміщені по горизонталі."
    )
    subtitle_cell.fill = PatternFill("solid", fgColor=NOTE_FILL)
    subtitle_cell.font = Font(color=TEXT_COLOR, italic=True, size=11)
    subtitle_cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )
    worksheet.row_dimensions[2].height = 38

    period = "—"
    if dates:
        period = f"{dates[0].strftime('%d.%m.%Y')}–{dates[-1].strftime('%d.%m.%Y')}"

    worksheet.merge_cells("A3:H3")
    period_cell = worksheet["A3"]
    period_cell.value = f"Період даних: {period}"
    period_cell.fill = PatternFill("solid", fgColor="EEF5FB")
    period_cell.font = Font(color=TITLE_FILL, bold=True, size=10)
    period_cell.alignment = Alignment(horizontal="left", vertical="center")
    worksheet.row_dimensions[3].height = 24

    sections = [
        (
            "Основні дані",
            [
                ("Усі ціни", "Повний масив усіх цінових рядів із датами по горизонталі."),
                ("Зернові", "Окремий зріз за зерновими культурами."),
                ("Олійні", "Окремий зріз за олійними культурами."),
                ("Рослинні олії", "Цінові ряди за рослинними оліями."),
                ("Шроти", "Цінові ряди за шротами та продуктами переробки."),
            ],
            DATE_FILL,
        ),
        (
            "Аналітичні зрізи",
            [
                ("Ключові індикатори", "Добірка головних ринкових показників і їх зміни."),
                ("Україна — експорт", "Українські експортні котирування у USD."),
                ("Україна — внутрішній ринок", "Внутрішні котирування України у гривні."),
                ("Портові спреди", "Різниця між портами та базисами постачання."),
                ("Пшениця — премії якості", "Премії між фуражною пшеницею, 11,5% і 12,5%."),
                ("Чорне море — конкуренти", "Порівняння України з основними конкурентами."),
            ],
            "5B9BD5",
        ),
        (
            "Службові аркуші",
            [
                ("База даних", "Повний технічний масив спостережень без розвороту дат."),
                ("Журнал", "Результати обробки файлів і вилучення дублікатів."),
                ("Контроль якості", "Помилки, підозри на копії аркушів і покриття аналітичних серій."),
            ],
            "7F8C8D",
        ),
    ]

    thin_side = Side(style="thin", color=BORDER_COLOR)
    row = 5

    for section_name, items, link_fill in sections:
        worksheet.merge_cells(
            start_row=row,
            start_column=1,
            end_row=row,
            end_column=8,
        )
        section_cell = worksheet.cell(row=row, column=1)
        section_cell.value = section_name
        section_cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        section_cell.font = Font(color=WHITE, bold=True, size=12)
        section_cell.alignment = Alignment(horizontal="left", vertical="center")
        worksheet.row_dimensions[row].height = 26
        row += 1

        for item_number, (sheet_name, description) in enumerate(items):
            worksheet.merge_cells(
                start_row=row,
                start_column=2,
                end_row=row,
                end_column=3,
            )
            worksheet.merge_cells(
                start_row=row,
                start_column=4,
                end_row=row,
                end_column=8,
            )

            link_cell = worksheet.cell(row=row, column=2)
            link_cell.value = sheet_name
            escaped = sheet_name.replace("'", "''")
            link_cell.hyperlink = Hyperlink(
                ref=link_cell.coordinate,
                location=f"'{escaped}'!A1",
                display=sheet_name,
            )
            link_cell.fill = PatternFill("solid", fgColor=link_fill)
            link_cell.font = Font(
                color=WHITE,
                bold=True,
                underline="single",
                size=11,
            )
            link_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            description_cell = worksheet.cell(row=row, column=4)
            description_cell.value = description
            description_cell.fill = PatternFill(
                "solid",
                fgColor=GROUP_FILL if item_number % 2 == 0 else WHITE,
            )
            description_cell.font = Font(color=TEXT_COLOR, size=10)
            description_cell.alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True,
            )

            for column in range(2, 9):
                worksheet.cell(row=row, column=column).border = Border(
                    bottom=thin_side
                )

            worksheet.row_dimensions[row].height = 34
            row += 1

        row += 1

    worksheet.merge_cells(
        start_row=row,
        start_column=1,
        end_row=row + 1,
        end_column=8,
    )
    note_cell = worksheet.cell(row=row, column=1)
    note_cell.value = (
        "Як читати файл: кожен рядок — окремий ціновий показник або "
        "аналітичний розрахунок; кожна нова дата автоматично додається праворуч."
    )
    note_cell.fill = PatternFill("solid", fgColor=SPREAD_FILL)
    note_cell.font = Font(color=TEXT_COLOR, italic=True, size=10)
    note_cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )
    worksheet.row_dimensions[row].height = 28
    worksheet.row_dimensions[row + 1].height = 28

    widths = {
        "A": 3,
        "B": 18,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 18,
        "G": 18,
        "H": 18,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    workbook.active = 0


def reorder_sheets(workbook) -> None:
    """Явно встановлює погоджений порядок вкладок перед збереженням."""

    missing = [name for name in FINAL_SHEET_ORDER if name not in workbook.sheetnames]
    if missing:
        raise RuntimeError(
            "Неможливо встановити порядок вкладок. Відсутні аркуші: "
            + ", ".join(missing)
        )

    unexpected = [name for name in workbook.sheetnames if name not in FINAL_SHEET_ORDER]
    if unexpected:
        raise RuntimeError(
            "У книзі є неочікувані аркуші: " + ", ".join(unexpected)
        )

    workbook._sheets = [workbook[name] for name in FINAL_SHEET_ORDER]
    workbook.active = 0


def report_status(log: pd.DataFrame, issues: pd.DataFrame) -> int:
    """Друкує підсумок, пише його в GitHub Actions і повертає код завершення."""
    errors = issues[issues["Рівень"] == "Помилка"]
    warnings = issues[issues["Рівень"] == "Попередження"]
    attention = issues[issues["Рівень"] == "Увага"]
    processed = int((log["Статус"] == "Оброблено").sum())
    in_db = int((log["Статус"] == "Вже в базі").sum())

    lines = [
        f"Оброблено нових аркушів: {processed}; файлів без змін (взято з бази): {in_db}",
        f"Помилок: {len(errors)}; попереджень: {len(warnings)}; серій без даних на останню дату: {len(attention)}",
    ]
    for line in lines:
        print(line)

    in_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    for _, issue in pd.concat([errors, warnings]).iterrows():
        target = issue["Об'єкт"]
        text = f"{target}: {issue['Повідомлення']}".replace("\n", " ")
        if in_actions:
            kind = "error" if issue["Рівень"] == "Помилка" else "warning"
            print(f"::{kind} title={issue['Розділ']}::{text}")
        else:
            print(f"  [{issue['Рівень']}] {issue['Розділ']} — {text}")

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as summary:
            summary.write("## UkrAgro: результат запуску\n\n")
            for line in lines:
                summary.write(f"- {line}\n")
            if not errors.empty or not warnings.empty:
                summary.write("\n| Рівень | Розділ | Об'єкт | Повідомлення |\n|---|---|---|---|\n")
                for _, issue in pd.concat([errors, warnings]).iterrows():
                    cells = [str(issue[col]).replace("|", "/") for col in ["Рівень", "Розділ", "Об'єкт", "Повідомлення"]]
                    summary.write("| " + " | ".join(cells) + " |\n")

    if not errors.empty:
        print(f"ЗАВЕРШЕНО З ПОМИЛКАМИ: {len(errors)} аркуш(ів) не прочитано — див. аркуш «Контроль якості».")
        return 1
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Зведення щоденних цін UkrAgroConsult")
    parser.add_argument(
        "--full-rebuild",
        action="store_true",
        help="Ігнорувати накопичену базу data/prices_db.csv і перечитати всі файли з input/",
    )
    args = parser.parse_args()

    print(f"Версія скрипта: {SCRIPT_VERSION}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_prices, log, issue_rows = collect_data(full_rebuild=args.full_rebuild)
    dates = date_list(all_prices)
    selection_data = add_selection_columns(all_prices)

    key_indicators, key_notes = build_key_indicators(selection_data, dates, issue_rows)
    ukraine_export = build_compact_market_table(all_prices, dates, mode="export")
    ukraine_internal = build_compact_market_table(all_prices, dates, mode="internal")
    port_spreads, port_spread_formulas, port_notes = build_port_spreads(selection_data, dates, issue_rows)
    wheat_premiums, wheat_premium_formulas, wheat_notes = build_wheat_premiums(selection_data, dates, issue_rows)
    competitors, competitor_formulas, competitor_notes = build_competitors(selection_data, dates, issue_rows)

    level_order = {"Помилка": 0, "Попередження": 1, "Увага": 2, "ОК": 3}
    issues = pd.DataFrame(issue_rows, columns=["Рівень", "Розділ", "Об'єкт", "Повідомлення"])
    issues = issues.sort_values(
        by="Рівень", key=lambda column: column.map(level_order), kind="stable"
    ).reset_index(drop=True)

    wide_sheets = [
        ("Усі ціни", all_prices, True),
        ("Зернові", all_prices[all_prices["Тип товару"] == "Grain"], False),
        ("Олійні", all_prices[all_prices["Тип товару"] == "Oilseeds"], False),
        ("Рослинні олії", all_prices[all_prices["Тип товару"] == "Vegoil"], False),
        ("Шроти", all_prices[all_prices["Тип товару"] == "Meals"], False),
    ]

    prepared_wide: dict[str, tuple[pd.DataFrame, bool]] = {}

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet_name, source, include_category in wide_sheets:
            wide = build_wide_table(source, include_category=include_category)
            prepared_wide[sheet_name] = (wide, include_category)
            wide.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)

        key_indicators.to_excel(writer, sheet_name="Ключові індикатори", index=False, startrow=3)
        ukraine_export.to_excel(writer, sheet_name="Україна — експорт", index=False, startrow=3)
        ukraine_internal.to_excel(writer, sheet_name="Україна — внутрішній ринок", index=False, startrow=3)
        port_spreads.to_excel(writer, sheet_name="Портові спреди", index=False, startrow=3)
        wheat_premiums.to_excel(writer, sheet_name="Пшениця — премії якості", index=False, startrow=3)
        competitors.to_excel(writer, sheet_name="Чорне море — конкуренти", index=False, startrow=3)

        all_prices.to_excel(writer, sheet_name="База даних", index=False)
        log.to_excel(writer, sheet_name="Журнал", index=False)
        issues.to_excel(writer, sheet_name="Контроль якості", index=False)

    workbook = load_workbook(OUTPUT_FILE)

    style_key_indicators(workbook["Ключові індикатори"], key_indicators, dates)
    add_delivery_comments(workbook["Ключові індикатори"], key_notes, date_start_column=11)
    style_compact_market(
        workbook["Україна — експорт"],
        ukraine_export,
        "Україна — експортні ціни",
        "Лише українські котирування в USD на визначених експортних точках: порти, Рені та західний кордон. "
        "Для компактності показано середню ціну продавця.",
    )
    style_compact_market(
        workbook["Україна — внутрішній ринок"],
        ukraine_internal,
        "Україна — внутрішній ринок",
        "Внутрішні котирування України у гривні за базисами CPT та EXW. "
        "Для компактності показано середню ціну продавця.",
    )

    comparison_sheets = [
        (
            "Портові спреди", port_spreads, port_spread_formulas, port_notes, 8,
            "Портові спреди",
            "Для кожного порівняння наведено дві вихідні ціни та розрахунковий спред. Обидві ціни беруться "
            "за однаковий (найближчий спільний) місяць поставки. Додатне значення — перша серія дорожча.",
            [28, 22, 12, 30, 34, 11, 20],
        ),
        (
            "Пшениця — премії якості", wheat_premiums, wheat_premium_formulas, wheat_notes, 9,
            "Пшениця — премії за якість",
            "Вихідні ціни фуражної пшениці, 11,5% і 12,5% за однаковий місяць поставки та премії між класами. "
            "Усі значення — USD/т.",
            [18, 14, 12, 20, 28, 10, 11, 20],
        ),
        (
            "Чорне море — конкуренти", competitors, competitor_formulas, competitor_notes, 11,
            "Чорне море — конкуренти",
            "Порівняння українських FOB-котирувань з Росією, ЄС Black Sea та Францією за однаковий місяць "
            "поставки. Спреди — ціна України мінус ціна конкурента.",
            [28, 16, 14, 12, 28, 22, 24, 10, 11, 20],
        ),
    ]
    for sheet_name, frame, formulas, notes, date_start, title, note, widths in comparison_sheets:
        special_rows = write_formulas_for_comparison(
            workbook[sheet_name], formulas, date_start_column=date_start, date_count=len(dates)
        )
        style_comparison_sheet(
            workbook[sheet_name], frame, title, note,
            date_start_column=date_start, widths=widths, special_rows=special_rows,
        )
        add_delivery_comments(workbook[sheet_name], notes, date_start_column=date_start)

    for sheet_name, (wide, include_category) in prepared_wide.items():
        style_wide_sheet(workbook[sheet_name], wide, title=sheet_name, include_category=include_category)

    style_database_sheet(workbook["База даних"], all_prices)
    style_log_sheet(workbook["Журнал"])
    style_issues_sheet(workbook["Контроль якості"])

    create_navigation_sheet(workbook, dates)
    reorder_sheets(workbook)
    workbook.save(OUTPUT_FILE)

    print(f"Готово: {OUTPUT_FILE}")
    print(f"Накопичена база: {DB_FILE} ({len(all_prices)} рядків, {len(dates)} дат)")
    return report_status(log, issues)


if __name__ == "__main__":
    sys.exit(main())
