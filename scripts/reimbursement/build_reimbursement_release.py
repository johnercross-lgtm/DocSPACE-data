#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import shutil
import sqlite3
import sys
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


DATASET_ID = "21a6930e-4346-461c-8d80-abeb6f9c0ae2"
DATASET_API_URL = f"https://data.gov.ua/api/3/action/package_show?id={DATASET_ID}"
SCHEMA_VERSION = 1
MINIMUM_APP_VERSION = None
SQLITE_NAME = "reimbursement.sqlite"
ZIP_NAME = "reimbursement.zip"
MANIFEST_NAME = "reimbursement_manifest.json"
PIPELINE_ID = "docspace-reimbursement-xlsx-v1"
CANONICAL_SOURCE_ID = f"data.gov.ua:dataset:{DATASET_ID}:xlsx"
MINIMUM_RELATIVE_COUNT_RATIO = 0.80
VOLUME_CHECK_SECTIONS = ("medicines", "insulins", "combined_medicines")

UA_MONTHS = {
    "січня": 1,
    "лютого": 2,
    "березня": 3,
    "квітня": 4,
    "травня": 5,
    "червня": 6,
    "липня": 7,
    "серпня": 8,
    "вересня": 9,
    "жовтня": 10,
    "листопада": 11,
    "грудня": 12,
}

SECTIONS = {
    "medicines": "I. Лікарські засоби, крім препаратів інсуліну та комбінованих лікарських засобів",
    "insulins": "II. Препарати інсуліну",
    "combined_medicines": "III. Комбіновані лікарські засоби",
}

ITEM_COLUMNS = [
    "source_row",
    "section",
    "section_uk",
    "ordinal_no",
    "inn_name",
    "trade_name",
    "dosage_form",
    "dosage",
    "pack_units",
    "atc_code",
    "manufacturer_country",
    "registration_number",
    "registration_expiry",
    "wholesale_price_uah",
    "retail_price_uah",
    "who_daily_dose",
    "reimbursement_daily_dose_uah",
    "reimbursement_primary_pack_uah",
    "reimbursement_release_form_or_primary_pack_uah",
    "reimbursement_pack_uah",
    "patient_copay_uah",
    "copay_type",
    "release_form_or_primary_pack",
    "icd10_codes",
]

STABLE_KEY_FIELDS = (
    "section",
    "registration_number",
    "trade_name",
    "dosage_form",
    "dosage",
    "pack_units",
    "atc_code",
)

REQUIRED_ITEM_COLUMNS = {
    "id",
    "source_row",
    "section",
    "section_uk",
    "inn_name",
    "trade_name",
    "dosage_form",
    "dosage",
    "pack_units",
    "atc_code",
    "manufacturer_country",
    "registration_number",
    "registration_expiry",
    "retail_price_uah",
    "reimbursement_pack_uah",
    "patient_copay_uah",
    "icd10_codes",
}

EXPECTED_HEADER_TOKENS = {
    "inn": ("міжнарод", "непатент"),
    "trade": ("торгов", "назв"),
    "form": ("форм",),
    "dosage": ("доз",),
    "pack": ("кільк", "упаков"),
    "atc": ("атх",),
    "registration": ("реєстрац", "посвід"),
    "retail": ("роздріб",),
    "reimbursement": ("відшкодуван",),
}
FREE_COPAY_LABELS = {"безоплатно"}


@dataclass(frozen=True)
class SourceCandidate:
    name: str
    url: str
    resource_id: str
    register_date: date
    description: str
    last_modified: str | None
    size: int | None


@dataclass(frozen=True)
class BuildStats:
    source_version: str
    rows_downloaded: int
    medicines_count: int
    insulins_count: int
    combined_medicines_count: int
    medical_devices_count: int
    database_version: int
    sha256: str
    published: bool


def log(message: str) -> None:
    print(message, flush=True)


def fetch_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(url, headers={"User-Agent": "DocSPACE-data updater"})
    with urllib.request.urlopen(request, timeout=45) as response:
        return json.loads(response.read().decode("utf-8"))


def download(url: str, destination: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "DocSPACE-data updater"})
    with urllib.request.urlopen(request, timeout=90) as response, destination.open("wb") as handle:
        shutil.copyfileobj(response, handle)


def parse_date_from_text(text: str) -> date | None:
    compact = re.search(r"(\d{1,2})[._-](\d{1,2})[._-](20\d{2})", text)
    if compact:
        day, month, year = map(int, compact.groups())
        return date(year, month, day)

    spoken = re.search(
        r"(\d{1,2})\s+("
        + "|".join(UA_MONTHS)
        + r")\s+(20\d{2})(?:\s+року)?",
        text.casefold(),
    )
    if spoken:
        day = int(spoken.group(1))
        month = UA_MONTHS[spoken.group(2)]
        year = int(spoken.group(3))
        return date(year, month, day)
    return None


def parse_register_date_from_text(text: str) -> date | None:
    match = re.search(r"станом\s+на\s+(.{0,60})", text.casefold())
    if not match:
        return parse_date_from_text(text)
    return parse_date_from_text(match.group(1))


def parse_strict_register_date_from_text(text: str) -> date | None:
    match = re.search(r"станом\s+на\s+(.{0,60})", text.casefold())
    if not match:
        return None
    return parse_date_from_text(match.group(1))


def parse_manifest_date(value: str | None) -> date | None:
    if not value:
        return None
    for candidate in (value[:10], value):
        try:
            return datetime.fromisoformat(candidate).date()
        except ValueError:
            pass
    return None


def latest_published_manifest(path: Path | None) -> dict[str, Any] | None:
    if path is None or not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def discover_sources() -> list[SourceCandidate]:
    payload = fetch_json(DATASET_API_URL)
    if not payload.get("success"):
        raise RuntimeError("data.gov.ua package_show returned success=false")

    dataset = payload["result"]
    organization = dataset.get("organization") or {}
    if organization.get("name") != "ministerstvo-okhorony-zdorovia-ukrayiny":
        raise RuntimeError("Unexpected data.gov.ua dataset organization")
    dataset_title = str(dataset.get("title") or "").casefold()
    if "реімбурсац" not in dataset_title or "лікар" not in dataset_title:
        raise RuntimeError("Unexpected data.gov.ua dataset title")

    candidates: list[SourceCandidate] = []
    for resource in dataset.get("resources", []):
        name = str(resource.get("name") or "")
        description = str(resource.get("description") or "")
        url = str(resource.get("url") or "")
        fmt = str(resource.get("format") or resource.get("mimetype") or "").casefold()
        haystack = " ".join([name, description, url]).casefold()
        if "xlsx" not in fmt and not url.casefold().endswith(".xlsx"):
            continue
        if "реімбурсац" not in haystack and "reimburs" not in haystack:
            continue
        if "лікар" not in haystack and "lik" not in haystack:
            continue

        register_date = parse_register_date_from_text(" ".join([name, description, url]))
        if register_date is None:
            continue

        size = resource.get("size")
        candidates.append(
            SourceCandidate(
                name=name,
                url=url,
                resource_id=str(resource.get("id") or ""),
                register_date=register_date,
                description=description,
                last_modified=resource.get("last_modified"),
                size=int(size) if isinstance(size, int) else None,
            )
        )

    return sorted(candidates, key=lambda item: (item.register_date, item.name), reverse=True)


def normalized_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ").strip())


def normalized_cell(value: Any) -> str:
    return normalized_text(value).casefold().replace("’", "'")


def as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return normalized_text(value)


def as_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(str(value).replace(",", ".")))
    except ValueError:
        return None


def as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if not math.isfinite(number):
            raise RuntimeError(f"Unexpected non-finite numeric value: {value!r}")
        return number

    text = normalized_text(value).replace("\u00a0", " ")
    compact = re.sub(r"\s+", "", text)
    if "," in compact and "." in compact:
        normalized = (
            compact.replace(".", "").replace(",", ".")
            if compact.rfind(",") > compact.rfind(".")
            else compact.replace(",", "")
        )
    else:
        normalized = compact.replace(",", ".")
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", normalized):
        raise RuntimeError(f"Unexpected numeric value: {value!r}")
    return float(normalized)


def parse_patient_copay(value: Any) -> tuple[float | None, str]:
    if normalized_cell(value) in FREE_COPAY_LABELS:
        return 0.0, normalized_text(value)
    return as_float(value), ""


def normalized_key_number(value: Any) -> str:
    number = as_float(value)
    if number is None:
        return ""
    return format(number, ".15g")


def stable_item_key(row: dict[str, Any] | sqlite3.Row) -> tuple[str, ...] | None:
    parts: list[str] = []
    for field in STABLE_KEY_FIELDS:
        value = row[field]
        if field == "pack_units":
            normalized = normalized_key_number(value)
        else:
            normalized = normalized_cell(value)
        if not normalized:
            return None
        parts.append(normalized)
    return tuple(parts)


def row_has_header_tokens(row: tuple[Any, ...]) -> bool:
    row_text = " ".join(normalized_cell(cell) for cell in row if cell is not None)
    return sum(all(token in row_text for token in tokens) for tokens in EXPECTED_HEADER_TOKENS.values()) >= 7


def validate_expected_headers(worksheet: Any) -> None:
    for row in worksheet.iter_rows(min_row=1, max_row=12, values_only=True):
        if row_has_header_tokens(row):
            return
    raise RuntimeError("XLSX header does not contain the expected reimbursement columns")


def detect_section(row_values: list[Any], current_section: str) -> str:
    text = " ".join(normalized_cell(value) for value in row_values)
    if "лікарські засоби" in text and "крім" in text:
        return "medicines"
    if "препарати інсуліну" in text:
        return "insulins"
    if "комбінован" in text:
        return "combined_medicines"
    return current_section


def validate_required_numeric_fields(rows: list[dict[str, Any]]) -> None:
    required_by_section = {
        "medicines": (
            "pack_units",
            "wholesale_price_uah",
            "retail_price_uah",
            "reimbursement_daily_dose_uah",
            "reimbursement_pack_uah",
            "patient_copay_uah",
        ),
        "insulins": (
            "pack_units",
            "wholesale_price_uah",
            "retail_price_uah",
            "reimbursement_primary_pack_uah",
            "patient_copay_uah",
        ),
        "combined_medicines": (
            "pack_units",
            "wholesale_price_uah",
            "retail_price_uah",
            "reimbursement_release_form_or_primary_pack_uah",
            "reimbursement_pack_uah",
            "patient_copay_uah",
        ),
    }
    for row in rows:
        missing = [field for field in required_by_section[row["section"]] if row[field] is None]
        if missing:
            raise RuntimeError(
                f"Missing required numeric values in {row['section']} row {row['source_row']}: {missing}"
            )


def parse_xlsx(path: Path) -> tuple[list[dict[str, Any]], date | None]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    if not workbook.worksheets:
        raise RuntimeError("XLSX has no worksheets")

    worksheet = workbook.worksheets[0]
    validate_expected_headers(worksheet)
    workbook_register_date: date | None = None
    for row in worksheet.iter_rows(min_row=1, max_row=20, values_only=True):
        workbook_register_date = parse_strict_register_date_from_text(" ".join(normalized_text(value) for value in row))
        if workbook_register_date:
            break
    if workbook_register_date is None:
        raise RuntimeError("XLSX register date is missing from the workbook")

    rows: list[dict[str, Any]] = []
    current_section = "medicines"
    seen_sections: set[str] = set()
    for excel_row_no, values in enumerate(worksheet.iter_rows(min_row=1, values_only=True), start=1):
        row = list(values[1:18])
        if not any(normalized_text(value) for value in row):
            continue
        current_section = detect_section(row, current_section)

        ordinal = as_int(row[0] if row else None)
        if ordinal is None or len(row) < 16:
            continue

        section = current_section
        seen_sections.add(section)
        if section == "insulins":
            patient_copay = as_float(row[13])
            copay_type = as_text(row[14]) if len(row) > 14 else ""
            reimbursement_primary_pack = as_float(row[12])
        else:
            patient_copay, copay_type = parse_patient_copay(row[15])
            if len(row) > 16 and normalized_text(row[16]):
                copay_type = as_text(row[16])
            reimbursement_primary_pack = None
        item = {
            "source_row": excel_row_no,
            "section": section,
            "section_uk": SECTIONS[section],
            "ordinal_no": ordinal,
            "inn_name": as_text(row[1]),
            "trade_name": as_text(row[2]),
            "dosage_form": as_text(row[3]),
            "dosage": as_text(row[4]),
            "pack_units": as_float(row[5]),
            "atc_code": as_text(row[6]),
            "manufacturer_country": as_text(row[7]),
            "registration_number": as_text(row[8]),
            "registration_expiry": as_text(row[9]),
            "wholesale_price_uah": as_float(row[10]),
            "retail_price_uah": as_float(row[11]),
            "reimbursement_pack_uah": None if section == "insulins" else as_float(row[14]),
            "patient_copay_uah": patient_copay,
            "copay_type": copay_type,
            "icd10_codes": "",
        }
        if section == "combined_medicines":
            item.update(
                {
                    "who_daily_dose": None,
                    "reimbursement_daily_dose_uah": None,
                    "reimbursement_primary_pack_uah": None,
                    "reimbursement_release_form_or_primary_pack_uah": as_float(row[13]),
                    "release_form_or_primary_pack": as_text(row[12]),
                }
            )
        elif section == "insulins":
            item.update(
                {
                    "who_daily_dose": None,
                    "reimbursement_daily_dose_uah": None,
                    "reimbursement_primary_pack_uah": reimbursement_primary_pack,
                    "reimbursement_release_form_or_primary_pack_uah": None,
                    "release_form_or_primary_pack": "",
                }
            )
        else:
            item.update(
                {
                    "who_daily_dose": as_float(row[12]),
                    "reimbursement_daily_dose_uah": as_float(row[13]),
                    "reimbursement_primary_pack_uah": None,
                    "reimbursement_release_form_or_primary_pack_uah": None,
                    "release_form_or_primary_pack": "",
                }
            )
        rows.append(item)

    missing_sections = set(VOLUME_CHECK_SECTIONS) - seen_sections
    if missing_sections:
        raise RuntimeError(f"XLSX missing expected sections: {sorted(missing_sections)}")
    if len(rows) < 100:
        raise RuntimeError(f"XLSX row count is anomalously small: {len(rows)}")
    validate_required_numeric_fields(rows)
    return rows, workbook_register_date


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE IF NOT EXISTS reimbursement_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_row INTEGER NOT NULL,
            section TEXT NOT NULL,
            section_uk TEXT NOT NULL,
            ordinal_no INTEGER,
            inn_name TEXT,
            trade_name TEXT,
            dosage_form TEXT,
            dosage TEXT,
            pack_units REAL,
            atc_code TEXT,
            manufacturer_country TEXT,
            registration_number TEXT,
            registration_expiry TEXT,
            wholesale_price_uah REAL,
            retail_price_uah REAL,
            who_daily_dose REAL,
            reimbursement_daily_dose_uah REAL,
            reimbursement_primary_pack_uah REAL,
            reimbursement_release_form_or_primary_pack_uah REAL,
            reimbursement_pack_uah REAL,
            patient_copay_uah REAL,
            copay_type TEXT,
            release_form_or_primary_pack TEXT,
            icd10_codes TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_reimbursement_items_atc ON reimbursement_items(atc_code);
        CREATE INDEX IF NOT EXISTS idx_reimbursement_items_inn ON reimbursement_items(inn_name);
        CREATE INDEX IF NOT EXISTS idx_reimbursement_items_trade ON reimbursement_items(trade_name);
        CREATE INDEX IF NOT EXISTS idx_reimbursement_items_section ON reimbursement_items(section);
        CREATE TABLE IF NOT EXISTS reimbursement_medical_devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_row INTEGER,
            medical_device_name TEXT,
            medical_device_form TEXT,
            manufacturer TEXT,
            device_description TEXT,
            declaration_info TEXT,
            classifier_code_name TEXT,
            wholesale_price_uah REAL,
            retail_price_uah REAL,
            reimbursement_pack_uah REAL,
            patient_copay_uah REAL
        );
        CREATE TABLE IF NOT EXISTS reference_price_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_row INTEGER NOT NULL,
            section TEXT NOT NULL,
            section_uk TEXT NOT NULL,
            ordinal_no INTEGER,
            inn_name TEXT,
            trade_name TEXT,
            dosage_form TEXT,
            dosage TEXT,
            pack_units REAL,
            atc_code TEXT,
            manufacturer_country TEXT,
            registration_number TEXT,
            registration_expiry TEXT,
            wholesale_price_uah REAL,
            retail_price_uah REAL,
            who_daily_dose REAL,
            reimbursement_daily_dose_uah REAL,
            reimbursement_primary_pack_uah REAL,
            reimbursement_release_form_or_primary_pack_uah REAL,
            reimbursement_pack_uah REAL,
            patient_copay_uah REAL,
            copay_type TEXT,
            release_form_or_primary_pack TEXT,
            icd10_codes TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_reference_price_items_atc ON reference_price_items(atc_code);
        CREATE INDEX IF NOT EXISTS idx_reference_price_items_inn ON reference_price_items(inn_name);
        CREATE INDEX IF NOT EXISTS idx_reference_price_items_trade ON reference_price_items(trade_name);
        CREATE INDEX IF NOT EXISTS idx_reference_price_items_section ON reference_price_items(section);
        """
    )


def table_exists(connection: sqlite3.Connection, table: str) -> bool:
    row = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table,),
    ).fetchone()
    return row is not None


def table_columns(connection: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in connection.execute(f"PRAGMA table_info({table})").fetchall()}


def prepare_database(database_path: Path, previous_zip: Path | None) -> None:
    if previous_zip and previous_zip.exists():
        with zipfile.ZipFile(previous_zip) as archive:
            names = archive.namelist()
            if names.count(SQLITE_NAME) != 1:
                raise RuntimeError(f"Previous ZIP must contain exactly one {SQLITE_NAME}")
            archive.extract(SQLITE_NAME, database_path.parent)
        extracted = database_path.parent / SQLITE_NAME
        if extracted != database_path:
            extracted.replace(database_path)
    with sqlite3.connect(database_path) as connection:
        create_schema(connection)


def item_section_counts(connection: sqlite3.Connection) -> dict[str, int]:
    return {
        row["section"]: int(row["count"])
        for row in connection.execute(
            "SELECT section, COUNT(*) AS count FROM reimbursement_items GROUP BY section"
        ).fetchall()
    }


def metadata_value(connection: sqlite3.Connection, key: str) -> str | None:
    if not table_exists(connection, "metadata"):
        return None
    row = connection.execute("SELECT value FROM metadata WHERE key=?", (key,)).fetchone()
    return normalized_text(row[0]) if row and row[0] is not None else None


def pipeline_baseline_counts(connection: sqlite3.Connection) -> dict[str, int] | None:
    if metadata_value(connection, "pipeline_id") != PIPELINE_ID:
        return None
    if metadata_value(connection, "canonical_source_id") != CANONICAL_SOURCE_ID:
        return None
    return item_section_counts(connection)


def transfer_icd10_codes(connection: sqlite3.Connection, rows: list[dict[str, Any]]) -> int:
    if "icd10_codes" not in table_columns(connection, "reimbursement_items"):
        return 0
    connection.row_factory = sqlite3.Row
    previous_by_key: dict[tuple[str, ...], list[str]] = {}
    for previous in connection.execute(
        "SELECT section, registration_number, trade_name, dosage_form, dosage, pack_units, atc_code, icd10_codes "
        "FROM reimbursement_items"
    ).fetchall():
        key = stable_item_key(previous)
        code = normalized_text(previous["icd10_codes"])
        if key is not None and code:
            previous_by_key.setdefault(key, []).append(code)

    new_key_counts: dict[tuple[str, ...], int] = {}
    new_keys: list[tuple[str, ...] | None] = []
    for row in rows:
        key = stable_item_key(row)
        new_keys.append(key)
        if key is not None:
            new_key_counts[key] = new_key_counts.get(key, 0) + 1

    transferred = 0
    for row, key in zip(rows, new_keys):
        if key is None or new_key_counts[key] != 1:
            continue
        previous_codes = previous_by_key.get(key, [])
        if len(previous_codes) != 1:
            continue
        row["icd10_codes"] = previous_codes[0]
        transferred += 1
    return transferred


def validate_volume_against_previous(
    previous_counts: dict[str, int],
    current_counts: dict[str, int],
) -> None:
    previous_total = sum(previous_counts.get(section, 0) for section in VOLUME_CHECK_SECTIONS)
    current_total = sum(current_counts.get(section, 0) for section in VOLUME_CHECK_SECTIONS)
    if previous_total > 0:
        minimum_total = math.ceil(previous_total * MINIMUM_RELATIVE_COUNT_RATIO)
        if current_total < minimum_total:
            raise RuntimeError(
                "Total reimbursement_items count dropped unexpectedly: "
                f"previous={previous_total}, current={current_total}, minimum={minimum_total}"
            )

    for section in VOLUME_CHECK_SECTIONS:
        previous_count = previous_counts.get(section, 0)
        if previous_count == 0:
            continue
        current_count = current_counts.get(section, 0)
        minimum_count = math.ceil(previous_count * MINIMUM_RELATIVE_COUNT_RATIO)
        if current_count < minimum_count:
            raise RuntimeError(
                f"{section} count dropped unexpectedly: "
                f"previous={previous_count}, current={current_count}, minimum={minimum_count}"
            )


def write_items(connection: sqlite3.Connection, rows: list[dict[str, Any]]) -> None:
    missing_columns = set(ITEM_COLUMNS) - table_columns(connection, "reimbursement_items")
    for column in sorted(missing_columns):
        connection.execute(f"ALTER TABLE reimbursement_items ADD COLUMN {column} TEXT")
    connection.execute("DELETE FROM reimbursement_items")
    insert_columns = ", ".join(ITEM_COLUMNS)
    insert_values = ", ".join(f":{column}" for column in ITEM_COLUMNS)
    connection.executemany(
        f"INSERT INTO reimbursement_items ({insert_columns}) VALUES ({insert_values})",
        rows,
    )


def write_metadata(connection: sqlite3.Connection, source: SourceCandidate, rows: list[dict[str, Any]]) -> None:
    section_counts = item_section_counts(connection)
    devices_count = (
        connection.execute("SELECT COUNT(*) FROM reimbursement_medical_devices").fetchone()[0]
        if table_exists(connection, "reimbursement_medical_devices")
        else 0
    )
    metadata = {
        "pipeline_id": PIPELINE_ID,
        "canonical_source_id": CANONICAL_SOURCE_ID,
        "document_title": "Перелік лікарських засобів, які підлягають реімбурсації",
        "register_date": source.register_date.isoformat(),
        "source_file": source.name,
        "source_url": source.url,
        "source_resource_id": source.resource_id,
        "source_description": source.description,
        "items_count": str(len(rows)),
        "devices_count": str(devices_count),
        "record_count_by_section": json.dumps(section_counts, ensure_ascii=False, sort_keys=True),
        "sections": json.dumps(SECTIONS, ensure_ascii=False, sort_keys=True),
        "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    connection.executemany(
        """
        INSERT INTO metadata(key, value)
        VALUES(?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        metadata.items(),
    )


def validate_sqlite(database_path: Path) -> dict[str, int]:
    with sqlite3.connect(database_path) as connection:
        connection.row_factory = sqlite3.Row
        integrity = connection.execute("PRAGMA integrity_check").fetchone()[0]
        if integrity != "ok":
            raise RuntimeError(f"SQLite integrity_check failed: {integrity}")
        for table in ("metadata", "reimbursement_items"):
            if not table_exists(connection, table):
                raise RuntimeError(f"Missing required table: {table}")
        columns = table_columns(connection, "reimbursement_items")
        missing = REQUIRED_ITEM_COLUMNS - columns
        if missing:
            raise RuntimeError(f"Missing reimbursement_items columns: {sorted(missing)}")
        section_counts = {
            row["section"]: int(row["count"])
            for row in connection.execute(
                "SELECT section, COUNT(*) AS count FROM reimbursement_items GROUP BY section"
            ).fetchall()
        }
        items_count = sum(section_counts.values())
        if items_count < 100:
            raise RuntimeError(f"reimbursement_items count is anomalously small: {items_count}")
        if section_counts.get("medicines", 0) < 50:
            raise RuntimeError("medicines section is anomalously small")
        if section_counts.get("insulins", 0) < 10:
            raise RuntimeError("insulins section is anomalously small")
        devices_count = 0
        if table_exists(connection, "reimbursement_medical_devices"):
            devices_count = connection.execute("SELECT COUNT(*) FROM reimbursement_medical_devices").fetchone()[0]
        if table_exists(connection, "reference_price_items"):
            _ = connection.execute("SELECT COUNT(*) FROM reference_price_items").fetchone()[0]
        return {
            "items": items_count,
            "medicines": section_counts.get("medicines", 0),
            "insulins": section_counts.get("insulins", 0),
            "combined_medicines": section_counts.get("combined_medicines", 0),
            "medical_devices": devices_count,
        }


def make_zip(database_path: Path, zip_path: Path) -> str:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        archive.write(database_path, SQLITE_NAME)
    with zipfile.ZipFile(zip_path) as archive:
        if archive.namelist() != [SQLITE_NAME]:
            raise RuntimeError(f"ZIP must contain only {SQLITE_NAME}")
    digest = hashlib.sha256()
    with zip_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_manifest(path: Path, *, database_version: int, source_date: date, repo: str, sha256: str) -> None:
    tag = f"reimbursement-v{database_version}"
    manifest = {
        "databaseVersion": database_version,
        "schemaVersion": SCHEMA_VERSION,
        "updatedAt": source_date.isoformat(),
        "minimumAppVersion": MINIMUM_APP_VERSION,
        "downloadURL": f"https://github.com/{repo}/releases/download/{tag}/{ZIP_NAME}",
        "sha256": sha256,
    }
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_summary(path: Path, stats: BuildStats, reason: str) -> None:
    path.write_text(
        json.dumps(
            {
                "publish": stats.published,
                "source_version": stats.source_version,
                "rows_downloaded": stats.rows_downloaded,
                "medicines_count": stats.medicines_count,
                "insulins_count": stats.insulins_count,
                "combined_medicines_count": stats.combined_medicines_count,
                "medical_devices_count": stats.medical_devices_count,
                "database_version": stats.database_version,
                "sha256": stats.sha256,
                "reason": reason,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build DocSPACE reimbursement release assets.")
    parser.add_argument("--work-dir", type=Path, required=True)
    parser.add_argument("--repo", required=True, help="GitHub repository, for example owner/name")
    parser.add_argument("--previous-manifest", type=Path)
    parser.add_argument("--previous-zip", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    args.work_dir.mkdir(parents=True, exist_ok=True)

    manifest = latest_published_manifest(args.previous_manifest)
    previous_version = int(manifest["databaseVersion"]) if manifest else 0
    previous_source_date = parse_manifest_date(manifest.get("updatedAt") if manifest else None)

    sources = discover_sources()
    if not sources:
        raise RuntimeError("No valid XLSX reimbursement resources were found")
    source = sources[0]

    log(f"Source version: {source.register_date.isoformat()} ({source.name})")
    if previous_source_date and source.register_date <= previous_source_date:
        stats = BuildStats(source.register_date.isoformat(), 0, 0, 0, 0, 0, previous_version, "", False)
        reason = f"No changes: latest valid XLSX source {source.register_date.isoformat()} is not newer than published {previous_source_date.isoformat()}"
        log(reason)
        write_summary(args.work_dir / "build_summary.json", stats, reason)
        return 0

    xlsx_path = args.work_dir / "source.xlsx"
    download(source.url, xlsx_path)
    rows, workbook_register_date = parse_xlsx(xlsx_path)
    if workbook_register_date != source.register_date:
        raise RuntimeError(
            "XLSX register date mismatch: "
            f"resource={source.register_date.isoformat()} workbook={workbook_register_date.isoformat()}"
        )
    database_version = previous_version + 1
    database_path = args.work_dir / SQLITE_NAME
    prepare_database(database_path, args.previous_zip)

    with sqlite3.connect(database_path) as connection:
        connection.row_factory = sqlite3.Row
        baseline_counts = pipeline_baseline_counts(connection)
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("BEGIN")
        transferred_icd10_codes = transfer_icd10_codes(connection, rows)
        write_items(connection, rows)
        write_metadata(connection, source, rows)
        connection.commit()
        connection.execute("VACUUM")

    counts = validate_sqlite(database_path)
    if baseline_counts is None:
        log("Volume baseline: none (first release for this XLSX pipeline/source)")
    else:
        log("Volume baseline: previous release from this XLSX pipeline/source")
        validate_volume_against_previous(baseline_counts, counts)
    zip_path = args.work_dir / ZIP_NAME
    sha256 = make_zip(database_path, zip_path)
    write_manifest(
        args.work_dir / MANIFEST_NAME,
        database_version=database_version,
        source_date=source.register_date,
        repo=args.repo,
        sha256=sha256,
    )

    stats = BuildStats(
        source_version=source.register_date.isoformat(),
        rows_downloaded=len(rows),
        medicines_count=counts["medicines"],
        insulins_count=counts["insulins"],
        combined_medicines_count=counts["combined_medicines"],
        medical_devices_count=counts["medical_devices"],
        database_version=database_version,
        sha256=sha256,
        published=True,
    )
    log(f"Rows downloaded: {stats.rows_downloaded}")
    log(f"Medicines count: {stats.medicines_count}")
    log(f"Insulins count: {stats.insulins_count}")
    log(f"Combined medicines count: {stats.combined_medicines_count}")
    log(f"Medical devices count: {stats.medical_devices_count}")
    log(f"ICD10 codes transferred: {transferred_icd10_codes}")
    log(f"Database version: {stats.database_version}")
    log(f"SHA256: {stats.sha256}")
    log("Published: release assets are ready")
    write_summary(args.work_dir / "build_summary.json", stats, "Published")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeError, urllib.error.URLError, zipfile.BadZipFile, sqlite3.Error) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
