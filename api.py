from __future__ import annotations

import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse


app = FastAPI(title = "Excel GO Parser")


VACATION_CODES = {"GO", "G.O", "G.O.", "G O"}
WORK_TYPE_MARKERS = {
    "REDOVAN RAD",
    "RAD NA TERENU",
    "POCETAK RADA",
    "ZAVRSETAK RADA",
    "UKUPNO DNEVNO RADNO VRIJEME",
}
NON_VACATION_CODES = {
    "BO",  # bolovanje
    "SD",  # placeni/slobodni dan in these sheets, not GO
    "N",  # night/non-working marker used on Sundays
    "DP",  # drzavni praznik
    "RO",  # roditeljski/ocinski
    "0",
    "8",
    "8.0",
}
WEEKDAY_KEYS = {
    "PON": 0,
    "UTO": 1,
    "SRI": 2,
    "CET": 3,
    "ČET": 3,
    "PE": 4,
    "PET": 4,
    "SUB": 5,
    "NED": 6,
}


@dataclass(frozen = True)
class SheetLayout:
    header_row: int
    day_columns: dict[int, int]
    name_col: int
    work_type_col: int | None


def strip_accents(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_key(value: Any) -> str:
    return strip_accents(normalize_text(value)).upper()


def detect_file_type(filename: str) -> str:
    key = normalize_key(filename)
    if "URED" in key:
        return "URED"
    if "GRADILIST" in key:
        return "GRADILISTE"
    return "UNKNOWN"


def infer_file_type_from_sheets(sheets: dict[str, pd.DataFrame]) -> str:
    sheet_names = [normalize_key(name) for name in sheets]
    if any("URED" in name for name in sheet_names):
        return "URED"
    if any("GRADILIST" in name for name in sheet_names):
        return "GRADILISTE"

    parseable_sheets = [name for name, df in sheets.items() if detect_sheet_layout(df) is not None]
    if len(parseable_sheets) > 1:
        return "GRADILISTE"
    if len(parseable_sheets) == 1:
        return "URED"
    return "UNKNOWN"


def first_detected_file_type(*filenames: str | None) -> str:
    for filename in filenames:
        if not filename:
            continue
        file_type = detect_file_type(filename)
        if file_type != "UNKNOWN":
            return file_type
    return "UNKNOWN"


def first_detected_extension(*filenames: str | None) -> str:
    for filename in filenames:
        if not filename:
            continue
        extension = os.path.splitext(filename)[1].lower()
        if extension in {".xls", ".xlsx"}:
            return extension
    return ""


def extract_month_year_from_filename(filename: str) -> tuple[int, int]:
    key = normalize_key(os.path.basename(filename)).lower()

    year_match = re.search(r"\b(20\d{2})\b", key)
    year = int(year_match.group(1)) if year_match else date.today().year

    # Covers "01 mj", "01. mj", "1 mjesec", etc.
    numeric_month = re.search(r"(?<!\d)(0?[1-9]|1[0-2])\s*\.?\s*(?:mj|mjesec)\b", key)
    if numeric_month:
        return int(numeric_month.group(1)), year

    # Covers "01-2026", "1.2026", "01_2026".
    month_year = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[\s._-]+(20\d{2})\b", key)
    if month_year:
        return int(month_year.group(1)), int(month_year.group(2))

    months = {
        "sijecanj": 1,
        "januar": 1,
        "veljaca": 2,
        "februar": 2,
        "ozujak": 3,
        "mart": 3,
        "travanj": 4,
        "april": 4,
        "svibanj": 5,
        "maj": 5,
        "lipanj": 6,
        "jun": 6,
        "srpanj": 7,
        "jul": 7,
        "kolovoz": 8,
        "august": 8,
        "rujan": 9,
        "septembar": 9,
        "listopad": 10,
        "oktobar": 10,
        "studeni": 11,
        "novembar": 11,
        "prosinac": 12,
        "decembar": 12,
    }
    for month_name, month_num in months.items():
        if month_name in key:
            return month_num, year

    return None, year


def filename_contains_month(filename: str) -> bool:
    key = normalize_key(os.path.basename(filename)).lower()
    if re.search(r"(?<!\d)(0?[1-9]|1[0-2])\s*\.?\s*(?:mj|mjesec)\b", key):
        return True
    if re.search(r"(?<!\d)(0?[1-9]|1[0-2])[\s._-]+(20\d{2})\b", key):
        return True
    month_words = {
        "sijecanj",
        "januar",
        "veljaca",
        "februar",
        "ozujak",
        "mart",
        "travanj",
        "april",
        "svibanj",
        "maj",
        "lipanj",
        "jun",
        "srpanj",
        "jul",
        "kolovoz",
        "august",
        "rujan",
        "septembar",
        "listopad",
        "oktobar",
        "studeni",
        "novembar",
        "prosinac",
        "decembar",
    }
    return any(month_word in key for month_word in month_words)


def parse_day_header(value: Any) -> int | None:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.day
    if isinstance(value, (int, float)) and float(value).is_integer():
        day = int(value)
        return day if 1 <= day <= 31 else None

    text = normalize_text(value)
    if re.fullmatch(r"\d{1,2}", text):
        day = int(text)
        return day if 1 <= day <= 31 else None
    return None


def is_probable_employee_name(value: Any) -> bool:
    text = normalize_text(value)
    if len(text) < 5:
        return False
    key = normalize_key(text)
    blocked = {
        "IME I PREZIME",
        "REDOVAN RAD",
        "RAD NA TERENU",
        "POCETAK RADA",
        "ZAVRSETAK RADA",
        "UKUPNO DNEVNO RADNO VRIJEME",
        "GRADILISTE",
        "SDK",
    }
    if key in blocked or any(marker in key for marker in ("PRAZNIK", "PRIJAVA", "DNEVNICA")):
        return False
    if len(text.split()) < 2:
        return False
    letters = re.sub(r"[^A-Za-zČĆŽŠĐčćžšđ ]", "", text)
    return len(letters.strip()) >= 5


def is_vacation_code(value: Any) -> bool:
    text = normalize_text(value).upper()
    if not text:
        return False
    compact = text.replace(" ", "")
    if compact in NON_VACATION_CODES:
        return False
    return text in VACATION_CODES or compact in {"GO", "G.O", "G.O."}


def find_longest_day_run(row: pd.Series) -> dict[int, int]:
    best: list[tuple[int, int]] = []
    current: list[tuple[int, int]] = []
    previous_day: int | None = None

    for col_idx, value in enumerate(row.tolist()):
        day = parse_day_header(value)
        continues = day is not None and previous_day is not None and day == previous_day + 1
        starts_month = day == 1

        if starts_month:
            current = [(col_idx, day)]
        elif continues:
            current.append((col_idx, day))
        else:
            if len(current) > len(best):
                best = current
            current = []

        previous_day = day

    if len(current) > len(best):
        best = current

    return dict(best)


def detect_work_type_col(df: pd.DataFrame, header_row: int, day_start_col: int) -> int | None:
    best_col: int | None = None
    best_score = 0
    for col_idx in range(max(0, day_start_col)):
        score = 0
        for row_idx in range(header_row + 1, min(len(df), header_row + 40)):
            key = normalize_key(df.iat[row_idx, col_idx])
            if key in WORK_TYPE_MARKERS or key.startswith("UKUPNO DNEVNO RADNO"):
                score += 1
        if score > best_score:
            best_col = col_idx
            best_score = score
    return best_col if best_score else None


def detect_name_col(df: pd.DataFrame, header_row: int, day_start_col: int) -> int:
    header = df.iloc[header_row]
    candidates = [
        idx
        for idx, value in enumerate(header.tolist()[:day_start_col])
        if normalize_key(value) == "IME I PREZIME"
    ]
    if candidates:
        return candidates[0]

    best_col = 0
    best_score = 0
    for col_idx in range(day_start_col):
        score = sum(
            1
            for row_idx in range(header_row + 1, min(len(df), header_row + 80))
            if is_probable_employee_name(df.iat[row_idx, col_idx])
        )
        if score > best_score:
            best_col = col_idx
            best_score = score
    return best_col


def detect_sheet_layout(df: pd.DataFrame) -> SheetLayout | None:
    best_row: int | None = None
    best_days: dict[int, int] = {}

    for row_idx in range(min(len(df), 25)):
        day_columns = find_longest_day_run(df.iloc[row_idx])
        if len(day_columns) > len(best_days):
            best_row = row_idx
            best_days = day_columns

    if best_row is None or len(best_days) < 20:
        return None

    day_start_col = min(best_days)
    name_col = detect_name_col(df, best_row, day_start_col)
    work_type_col = detect_work_type_col(df, best_row, day_start_col)

    return SheetLayout(
        header_row = best_row,
        day_columns = best_days,
        name_col = name_col,
        work_type_col = work_type_col,
    )


def weekday_value(value: Any) -> int | None:
    key = normalize_key(value)
    return WEEKDAY_KEYS.get(key)


def infer_month_from_sheet(df: pd.DataFrame, year: int) -> tuple[int | None, int]:
    layout = detect_sheet_layout(df)
    if layout is None:
        return None, 0

    weekday_row = layout.header_row + 1
    if weekday_row >= len(df):
        return None, 0

    best_month: int | None = None
    best_score = 0
    for month in range(1, 13):
        score = 0
        for col_idx, day in layout.day_columns.items():
            if col_idx >= df.shape[1]:
                continue
            observed = weekday_value(df.iat[weekday_row, col_idx])
            if observed is None:
                continue
            try:
                expected = date(year, month, day).weekday()
            except ValueError:
                continue
            if observed == expected:
                score += 1
        if score > best_score:
            best_month = month
            best_score = score

    return best_month, best_score


def infer_month_from_sheets(sheets: dict[str, pd.DataFrame], year: int) -> int | None:
    scores: dict[int, int] = {}
    for df in sheets.values():
        month, score = infer_month_from_sheet(df, year)
        if month is not None:
            scores[month] = scores.get(month, 0) + score

    if not scores:
        return None

    month, score = max(scores.items(), key=lambda item: item[1])
    # A valid month grid has about 28-31 weekday labels. Keep this conservative so
    # a sparse note row cannot accidentally override the filename/default.
    return month if score >= 20 else None


def is_employee_row(df: pd.DataFrame, row_idx: int, layout: SheetLayout) -> bool:
    if not is_probable_employee_name(df.iat[row_idx, layout.name_col]):
        return False
    if layout.work_type_col is None:
        return True
    return normalize_key(df.iat[row_idx, layout.work_type_col]) == "REDOVAN RAD"


def parse_sheet(df: pd.DataFrame, sheet_name: str, month: int, year: int) -> list[dict[str, Any]]:
    layout = detect_sheet_layout(df)
    if layout is None:
        return []

    workers: list[dict[str, Any]] = []
    for row_idx in range(layout.header_row + 1, len(df)):
        if not is_employee_row(df, row_idx, layout):
            continue

        name = normalize_text(df.iat[row_idx, layout.name_col])
        go_days = sorted(
            day
            for col_idx, day in layout.day_columns.items()
            if col_idx < df.shape[1] and is_vacation_code(df.iat[row_idx, col_idx])
        )

        if not go_days:
            continue

        workers.append(
            {
                "name": name,
                "go_days_count": len(go_days),
                "dates": [f"{day}.{month}.{year}." for day in go_days],
                "sheet": normalize_text(sheet_name),
            }
        )

    return workers


def format_dates(dates: list[str]) -> str:
    parsed: list[date] = []
    for value in dates:
        match = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(20\d{2})\.", value)
        if match:
            parsed.append(date(int(match.group(3)), int(match.group(2)), int(match.group(1))))

    if not parsed:
        return ", ".join(dates)

    parsed = sorted(set(parsed))
    groups: list[list[date]] = [[parsed[0]]]
    for item in parsed[1:]:
        if (item - groups[-1][-1]).days == 1:
            groups[-1].append(item)
        else:
            groups.append([item])

    parts: list[str] = []
    for group in groups:
        if len(group) == 1:
            item = group[0]
            parts.append(f"{item.day}.{item.month}.{item.year}.")
        else:
            start, end = group[0], group[-1]
            parts.append(f"od {start.day}.{start.month}.{start.year}. do {end.day}.{end.month}.{end.year}.")
    return ", ".join(parts)


def build_telegram_message(file_type: str, workers: list[dict[str, Any]]) -> str:
    if not workers:
        return "Nisu pronađeni GO dani u dokumentu."

    lines = ["Obrada završena.", f"Tip datoteke: {file_type}", ""]
    for worker in workers:
        lines.append(f"Radnik: {worker['name']}")
        lines.append(f"GO dana: {worker['go_days_count']}")
        lines.append(f"Datumi: {format_dates(worker['dates'])}")
        lines.append(f"Sheet/gradilište: {worker['sheet']}")
        lines.append("")
    return "\n".join(lines).strip()


def excel_engine_for(extension: str) -> str:
    if extension == ".xls":
        return "xlrd"
    if extension == ".xlsx":
        return "openpyxl"
    raise ValueError("Podržani su samo .xls i .xlsx dokumenti.")


def parse_excel_file(path: str, original_filename: str, file_type: str | None = None) -> dict[str, Any]:
    file_type = file_type or detect_file_type(original_filename)
    extension = os.path.splitext(original_filename)[1].lower()
    engine = excel_engine_for(extension)

    month, year = extract_month_year_from_filename(original_filename)

    sheets = pd.read_excel(
        path,
        sheet_name = None,
        header = None,
        engine = engine,
        dtype = object
    )

    if month is None:
        inferred_month = infer_month_from_sheets(sheets, year)
        if inferred_month is not None:
            month = inferred_month

    if file_type == "UNKNOWN":
        file_type = infer_file_type_from_sheets(sheets)
    if file_type == "UNKNOWN":
        raise ValueError("Nije moguće odrediti tip datoteke. Pošaljite originalni naziv s URED ili GRADILIŠTE.")

    selected_sheets = list(sheets.items())[:1] if file_type == "URED" else list(sheets.items())

    workers: list[dict[str, Any]] = []
    for sheet_name, df in selected_sheets:
        workers.extend(parse_sheet(df, str(sheet_name), month, year))

    return {
        "status": "ok",
        "file_type": file_type,
        "workers": workers,
        "message": build_telegram_message(file_type, workers),
    }


@app.get("/")
def home() -> dict[str, str]:
    return {"status": "ok", "message": "Excel GO Parser API radi."}


@app.post("/process-excel")
async def process_excel(file: UploadFile = File(...), filename: str | None = Form(None)):
    original_filename = filename or file.filename

    print("DEBUG filename field:", repr(filename))
    print("DEBUG upload filename:", repr(file.filename))
    print("DEBUG original_filename used:", repr(original_filename))

    if not original_filename:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Filename nije poslan."})

    file_type = first_detected_file_type(filename, file.filename)
    extension = first_detected_extension(filename, file.filename)
    if extension not in {".xls", ".xlsx"}:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Podržani su samo .xls i .xlsx dokumenti."})

    parser_filename = original_filename
    if os.path.splitext(parser_filename)[1].lower() not in {".xls", ".xlsx"}:
        parser_filename = f"{parser_filename}{extension}"

    tmp_path = ""
    try:
        with tempfile.NamedTemporaryFile(delete = False, suffix = extension) as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name

        return parse_excel_file(tmp_path, parser_filename, file_type = file_type)

    except ImportError as exc:
        return JSONResponse(
            status_code = 500,
            content = {
                "status": "error",
                "message": "Nedostaje Python paket za čitanje Excel dokumenta.",
                "details": str(exc),
            },
        )
    except ValueError as exc:
        return JSONResponse(status_code = 400, content = {"status": "error", "message": str(exc)})
    except Exception as exc:
        return JSONResponse(
            status_code = 500,
            content = {"status": "error", "message": "Greška pri obradi Excel dokumenta.", "details": str(exc)},
        )
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
