"""Fixture builders for the reimbursement pipeline tests.

Everything is generated at run time: no binary fixtures in the repository.
"""
from __future__ import annotations

import json
import sqlite3
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts" / "reimbursement"))

import build_reimbursement_release as builder  # noqa: E402

REGISTER_LINE = (
    "Перелік лікарських засобів, які підлягають реімбурсації за програмою державних "
    "гарантій медичного обслуговування населення, станом на 28 серпня 2025 року"
)

MEDICINES_HEADER = [
    "Порядковий номер", "Міжнародна непатентована назва", "Торговельна назва лікарського засобу",
    "Форма випуску", "Дозування", "Кількість одиниць лікарського засобу в упаковці", "Код АТХ",
    "Найменування виробника, країни", "Номер реєстраційного посвідчення",
    "Дата закінчення строку дії реєстраційного посвідчення", "Оптово-відпускна ціна за упаковку, гривень",
    "Роздрібна ціна за упаковку, гривень", "Добова доза лікарського засобу",
    "Розмір реімбурсації добової дози, гривень", "Розмір реімбурсації за упаковку, гривень",
    "Сума доплати за упаковку, гривень",
]
INSULINS_HEADER = [
    "№", "Міжнародна непатентована назва", "Торговельна назва препарату інсуліну", "Форма випуску",
    "Дозування", "Кількість МО в первинній упаковці", "Код АТХ", "Найменування виробника, країни",
    "Номер реєстраційного посвідчення", "Дата закінчення строку дії реєстраційного посвідчення",
    "Оптово-відпускна ціна за первинну упаковку, гривень",
    "Роздрібна ціна за первинну упаковку, гривень",
    "Розмір реімбурсації за первинну упаковку, гривень",
    "Сума доплати за первинну упаковку, гривень", "Тип доплати згідно з визначеною межею",
]
COMBINED_HEADER = [
    "№ з/п", "Міжнародна непатентована назва", "Торговельна назва лікарського засобу",
    "Форма випуску", "Дозування", "Кількість одиниць лікарського засобу в упаковці", "Код АТХ",
    "Найменування виробника, країни", "Номер реєстраційного посвідчення",
    "Дата закінчення строку дії реєстраційного посвідчення", "Оптово-відпускна ціна за упаковку, гривень",
    "Роздрібна ціна за упаковку, гривень", "Форма випуску/ первинна упаковка",
    "Розмір реімбурсації форми випуску/ первинної упаковки, гривень",
    "Розмір реімбурсації за упаковку, гривень", "Сума доплати за упаковку, гривень",
]


def _medicine_row(n: int) -> list:
    return [n, f"МНН {n} (Inn{n})", f"ТОРГ {n}", "таблетки", "100", 30, "A10AB01",
            "Виробник, Україна", f"UA/{1000 + n}/01/01", "необмежений", 100.0 + n, 133.0 + n,
            400, 5.5, 133.0 + n, 0]


def _insulin_row(n: int) -> list:
    return [n, f"Інсулін {n} (Insulin{n})", f"ІНС {n}", "розчин для ін'єкцій", "100 МО/мл", 1000,
            "A10AB01", "Виробник, Україна", f"UA/{2000 + n}/01/01", "необмежений", 454.6, 577.87,
            577.87, 0, "безоплатно"]


def _combined_row(n: int) -> list:
    return [n, f"Комбо {n} (Combo{n})", f"КОМБ {n}", "порошок для інгаляцій", "50 мкг/дозу", 60,
            "R03AK06", "Виробник, Україна", f"UA/{3000 + n}/01/01", "необмежений", 537.5, 714.3,
            "дискус", 714.3, 714.3, 0]


def write_workbook(
    path: Path,
    *,
    medicines: int = 60,
    insulins: int = 15,
    combined: int = 30,
    register_line: str | None = REGISTER_LINE,
    section_titles: tuple[str, str, str] = (
        "І. Лікарські засоби, крім препаратів інсуліну та комбінованих лікарських засобів",
        "II. Препарати інсуліну",
        "III. Комбіновані лікарські засоби",
    ),
    medicines_header: list | None = None,
    inject_rows: list[tuple[int, list]] | None = None,
    numbering_rows: bool = True,
) -> Path:
    """Build a workbook shaped like the official one: column A empty, data from column B."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Лікарські засоби"

    def emit(cells: list | None) -> None:
        worksheet.append([None] + list(cells) if cells is not None else [None])

    emit(["ЗАТВЕРДЖЕНО Наказ МОЗ"])
    if register_line is not None:
        emit([register_line])
    else:
        emit(["Перелік лікарських засобів без дати"])

    blocks = (
        (section_titles[0], medicines_header or MEDICINES_HEADER, _medicine_row, medicines),
        (section_titles[1], INSULINS_HEADER, _insulin_row, insulins),
        (section_titles[2], COMBINED_HEADER, _combined_row, combined),
    )
    for title, header, factory, count in blocks:
        emit([title])
        emit(None)
        emit(header)
        if numbering_rows:
            emit(list(range(1, len(header) + 1)))
        for index in range(1, count + 1):
            emit(factory(index))

    for _, extra in inject_rows or []:
        emit(extra)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def write_input_dir(directory: Path, **kwargs) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    return write_workbook(directory / "perelik.xlsx", **kwargs)


def make_baseline(
    directory: Path,
    *,
    medicines: int = 60,
    insulins: int = 15,
    combined: int = 30,
    pipeline_id: str | None = builder.PIPELINE_ID,
    canonical_source_id: str | None = builder.CANONICAL_SOURCE_ID,
    icd10_rows_per_section: int = 0,
    reference_prices: int = 12,
    devices: int = 5,
    extra_metadata: dict[str, str] | None = None,
    with_view: bool = True,
    database_version: int = 5,
    updated_at: str = "2025-08-01",
) -> tuple[Path, Path]:
    """Create a baseline SQLite + ZIP + manifest that mirrors a published release."""
    directory.mkdir(parents=True, exist_ok=True)
    database_path = directory / builder.SQLITE_NAME
    if database_path.exists():
        database_path.unlink()

    with sqlite3.connect(database_path) as connection:
        builder.create_schema(connection)
        for section, count, factory in (
            ("medicines", medicines, _medicine_row),
            ("insulins", insulins, _insulin_row),
            ("combined_medicines", combined, _combined_row),
        ):
            for index in range(1, count + 1):
                cells = factory(index)
                connection.execute(
                    "INSERT INTO reimbursement_items (source_row, section, section_uk, ordinal_no,"
                    " inn_name, trade_name, dosage_form, dosage, pack_units, atc_code,"
                    " manufacturer_country, registration_number, registration_expiry,"
                    " retail_price_uah, reimbursement_pack_uah, patient_copay_uah, icd10_codes)"
                    " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (
                        index, section, builder.SECTIONS[section], index,
                        cells[1], cells[2], cells[3], str(cells[4]), float(cells[5]), cells[6],
                        cells[7], cells[8], cells[9], 133.0, 133.0, 0.0,
                        "J45" if index <= icd10_rows_per_section else "",
                    ),
                )
        for index in range(reference_prices):
            connection.execute(
                "INSERT INTO reference_price_items (source_row, section, section_uk, trade_name)"
                " VALUES (?,?,?,?)",
                (index, "medicines", builder.SECTIONS["medicines"], f"REF {index}"),
            )
        for index in range(devices):
            connection.execute(
                "INSERT INTO reimbursement_medical_devices (source_row, medical_device_name)"
                " VALUES (?,?)",
                (index, f"DEVICE {index}"),
            )
        metadata = dict(extra_metadata or {})
        if pipeline_id is not None:
            metadata["pipeline_id"] = pipeline_id
        if canonical_source_id is not None:
            metadata["canonical_source_id"] = canonical_source_id
        for key, value in metadata.items():
            connection.execute("INSERT INTO metadata(key, value) VALUES (?,?)", (key, value))
        if with_view:
            connection.execute(
                "CREATE VIEW v_free_items AS SELECT * FROM reimbursement_items WHERE patient_copay_uah = 0"
            )

    zip_path = directory / builder.ZIP_NAME
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(database_path, builder.SQLITE_NAME)

    manifest_path = directory / builder.MANIFEST_NAME
    manifest_path.write_text(
        json.dumps(
            {
                "databaseVersion": database_version,
                "schemaVersion": 1,
                "updatedAt": updated_at,
                "minimumAppVersion": None,
                "downloadURL": "https://example.invalid/reimbursement.zip",
                "sha256": "0" * 64,
            }
        ),
        encoding="utf-8",
    )
    return manifest_path, zip_path
