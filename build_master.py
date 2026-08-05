from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "UKRAGRO_ALL_PRICES.xlsx"

COLUMN_MAP = {
    "товар": "Товар",
    "качество": "Якість",
    "страна": "Країна",
    "порт": "Порт",
    "базис": "Базис",
    "валюта": "Валюта",
    "seller min": "Продавець min",
    "seller max": "Продавець max",
    "buyer min": "Покупець min",
    "buyer max": "Покупець max",
    "месяц": "Місяць поставки",
    "сезон": "Сезон",
    "тип товара": "Тип товару",
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


def parse_sheet_date(sheet_name: str) -> pd.Timestamp | None:
    match = re.fullmatch(r"(\d{2}\.\d{2}\.\d{4})_min_max", sheet_name)
    if not match:
        return None
    return pd.to_datetime(match.group(1), format="%d.%m.%Y", errors="coerce")


def find_header_row(file_path: Path, sheet_name: str) -> int:
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=25,
        engine="openpyxl",
    )

    for row_number, row in preview.iterrows():
        normalized = {str(value).strip().lower() for value in row if pd.notna(value)}
        if "товар" in normalized and "тип товара" in normalized:
            return int(row_number)

    raise ValueError("Не знайдено рядок заголовків")


def workbook_rank(file_path: Path) -> pd.Timestamp:
    excel_file = pd.ExcelFile(file_path, engine="openpyxl")
    dates = [parse_sheet_date(name) for name in excel_file.sheet_names]
    dates = [date for date in dates if date is not None and pd.notna(date)]
    return max(dates) if dates else pd.Timestamp.min


def read_price_sheet(
    file_path: Path,
    sheet_name: str,
    report_date: pd.Timestamp,
    file_rank: pd.Timestamp,
) -> pd.DataFrame:
    header_row = find_header_row(file_path, sheet_name)
    data = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl",
    )

    data.columns = [str(column).strip().lower() for column in data.columns]
    data = data.rename(columns=COLUMN_MAP)

    required = list(COLUMN_MAP.values())
    missing = [column for column in required if column not in data.columns]
    if missing:
        raise ValueError(f"Відсутні колонки: {', '.join(missing)}")

    data = data[required].copy()
    data.insert(0, "Дата", report_date)

    for column in TEXT_COLUMNS:
        data[column] = (
            data[column]
            .astype("string")
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )

    for column in PRICE_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="coerce")

    data = data[data["Товар"].notna()].copy()
    data["Середня ціна продавця"] = data[["Продавець min", "Продавець max"]].mean(
        axis=1, skipna=True
    )
    data["Середня ціна покупця"] = data[["Покупець min", "Покупець max"]].mean(
        axis=1, skipna=True
    )
    data["Джерело — файл"] = file_path.name
    data["Джерело — аркуш"] = sheet_name
    data["_rank"] = file_rank

    return data


def collect_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    files = sorted(INPUT_DIR.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError("У папці input немає Excel-файлів")

    frames: list[pd.DataFrame] = []
    log_rows: list[dict[str, object]] = []

    for file_path in files:
        rank = workbook_rank(file_path)
        excel_file = pd.ExcelFile(file_path, engine="openpyxl")
        matching_sheets = [
            sheet for sheet in excel_file.sheet_names if parse_sheet_date(sheet) is not None
        ]

        if not matching_sheets:
            log_rows.append(
                {
                    "Файл": file_path.name,
                    "Аркуш": "—",
                    "Статус": "Пропущено",
                    "Рядків прочитано": 0,
                    "Повідомлення": "Не знайдено аркушів *_min_max",
                }
            )
            continue

        for sheet_name in matching_sheets:
            report_date = parse_sheet_date(sheet_name)
            try:
                frame = read_price_sheet(file_path, sheet_name, report_date, rank)
                frames.append(frame)
                log_rows.append(
                    {
                        "Файл": file_path.name,
                        "Аркуш": sheet_name,
                        "Статус": "Оброблено",
                        "Рядків прочитано": len(frame),
                        "Повідомлення": "",
                    }
                )
            except Exception as error:
                log_rows.append(
                    {
                        "Файл": file_path.name,
                        "Аркуш": sheet_name,
                        "Статус": "Помилка",
                        "Рядків прочитано": 0,
                        "Повідомлення": str(error),
                    }
                )

    if not frames:
        raise RuntimeError("Не вдалося прочитати жодної таблиці з цінами")

    combined = pd.concat(frames, ignore_index=True)
    before_deduplication = len(combined)

    combined = combined.sort_values(
        by=["Дата", "_rank", "Джерело — файл"],
        kind="stable",
    )
    combined = combined.drop_duplicates(subset=KEY_COLUMNS, keep="last")
    combined = combined.sort_values(
        by=["Дата", "Тип товару", "Товар", "Країна", "Порт", "Базис"],
        na_position="last",
        kind="stable",
    ).reset_index(drop=True)

    duplicates_removed = before_deduplication - len(combined)
    log_rows.append(
        {
            "Файл": "УСІ ФАЙЛИ",
            "Аркуш": "—",
            "Статус": "Підсумок",
            "Рядків прочитано": before_deduplication,
            "Повідомлення": f"Унікальних рядків: {len(combined)}; дублікатів вилучено: {duplicates_removed}",
        }
    )

    combined = combined.drop(columns=["_rank"])
    log = pd.DataFrame(log_rows)
    return combined, log


def style_output(file_path: Path) -> None:
    workbook = load_workbook(file_path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.sheet_view.showGridLines = False

        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=False)

        for column_cells in sheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in column_cells[:500]
            )
            sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 32)

        for cell in sheet["A"][1:]:
            if cell.value is not None:
                cell.number_format = "dd.mm.yyyy"

        header_names = {cell.value: cell.column for cell in sheet[1]}
        for column_name in PRICE_COLUMNS + [
            "Середня ціна продавця",
            "Середня ціна покупця",
        ]:
            column_index = header_names.get(column_name)
            if column_index:
                for cell in sheet.iter_cols(
                    min_col=column_index,
                    max_col=column_index,
                    min_row=2,
                ):
                    for price_cell in cell:
                        price_cell.number_format = "0.00"

    workbook.save(file_path)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_prices, log = collect_data()

    sheets = {
        "Усі ціни": all_prices,
        "Зернові": all_prices[all_prices["Тип товару"] == "Grain"],
        "Олійні": all_prices[all_prices["Тип товару"] == "Oilseeds"],
        "Рослинні олії": all_prices[all_prices["Тип товару"] == "Vegoil"],
        "Шроти": all_prices[all_prices["Тип товару"] == "Meals"],
        "Журнал": log,
    }

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet_name, dataframe in sheets.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    style_output(OUTPUT_FILE)
    print(f"Готово: {OUTPUT_FILE}")
    print(f"Унікальних цінових рядків: {len(all_prices)}")


if __name__ == "__main__":
    main()
