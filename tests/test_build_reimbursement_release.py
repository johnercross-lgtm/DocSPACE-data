"""Regression tests for the manual reimbursement pipeline.

Run with:  python -m unittest discover -s tests -v
"""
from __future__ import annotations

import json
import sqlite3
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from tests.support import (
    COMBINED_HEADER,
    MEDICINES_HEADER,
    builder,
    make_baseline,
    write_input_dir,
    write_workbook,
)


class BuilderTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.input_dir = self.root / "input"
        self.work_dir = self.root / "dist"
        self.baseline_dir = self.root / "baseline"

    def run_build(self, *extra: str, expect_success: bool = True) -> int:
        argv = [
            "--work-dir", str(self.work_dir),
            "--repo", "owner/repo",
            "--input-dir", str(self.input_dir),
            *extra,
        ]
        import sys

        original = sys.argv
        sys.argv = ["build_reimbursement_release.py", *argv]
        try:
            code = builder.main()
        finally:
            sys.argv = original
        if expect_success:
            self.assertEqual(code, 0)
        return code

    def summary(self) -> dict:
        return json.loads((self.work_dir / "build_summary.json").read_text(encoding="utf-8"))

    def built_database(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.work_dir / builder.SQLITE_NAME)
        connection.row_factory = sqlite3.Row
        self.addCleanup(connection.close)
        return connection


# --------------------------------------------------------------------------- parser guards


class NumericParsingTests(unittest.TestCase):
    def test_plain_and_decimal_comma(self) -> None:
        self.assertEqual(builder.as_float(12.5), 12.5)
        self.assertEqual(builder.as_float("12,5"), 12.5)
        self.assertEqual(builder.as_float("12.5"), 12.5)
        self.assertEqual(builder.as_float("0,3555"), 0.3555)

    def test_space_grouping_is_unambiguous(self) -> None:
        self.assertEqual(builder.as_float("1 234,56"), 1234.56)
        self.assertEqual(builder.as_float("1 234,56"), 1234.56)

    def test_mixed_separators_resolve_by_position(self) -> None:
        self.assertEqual(builder.as_float("1.234,56"), 1234.56)
        self.assertEqual(builder.as_float("1,234.56"), 1234.56)

    def test_repeated_separator_is_grouping(self) -> None:
        self.assertEqual(builder.as_float("1,234,567"), 1234567.0)
        self.assertEqual(builder.as_float("1.234.567"), 1234567.0)

    def test_ambiguous_thousands_group_fails_closed(self) -> None:
        """B7: '1,234' must never silently become 1.234."""
        for value in ("1,234", "12,345", "123,456", "1.234", "999.999"):
            with self.subTest(value=value):
                with self.assertRaises(RuntimeError) as ctx:
                    builder.as_float(value)
                self.assertIn("Ambiguous numeric value", str(ctx.exception))

    def test_garbage_fails(self) -> None:
        for value in ("abc", "12-34", "1,2,3.4.5"):
            with self.subTest(value=value):
                with self.assertRaises(RuntimeError):
                    builder.as_float(value)

    def test_non_finite_numeric_cell_fails(self) -> None:
        with self.assertRaises(RuntimeError):
            builder.as_float(float("inf"))


class SectionDetectionTests(unittest.TestCase):
    def test_roman_headings_switch_section(self) -> None:
        cases = {
            "І. Лікарські засоби, крім препаратів інсуліну та комбінованих лікарських засобів": "medicines",
            "II. Препарати інсуліну": "insulins",
            "III. Комбіновані лікарські засоби": "combined_medicines",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(builder.detect_section_header((None, text)), expected)

    def test_data_rows_never_switch_section(self) -> None:
        """B7: a trade name containing 'комбінований' must not reclassify the rest of the sheet."""
        data_row = (
            None, 12, "Будесонід + Формотерол", "КОМБІНОВАНИЙ ПРЕПАРАТ",
            "порошок для інгаляцій", "160мкг", 60, "R03AK07", "Виробник",
            "UA/1/01/01", "необмежений", 583.0, 774.77, "інгалятор", 774.77, 774.77, 0,
        )
        self.assertIsNone(builder.detect_section_header(data_row))

        insulin_mention = (
            None, 5, "Інсулін гларгін", "ПРЕПАРАТИ ІНСУЛІНУ ДОВГОЇ ДІЇ", "розчин",
            "100 МО/мл", 1000, "A10AE04", "Виробник", "UA/2/01/01", "необмежений",
            454.6, 577.87, 577.87, 0, "безоплатно",
        )
        self.assertIsNone(builder.detect_section_header(insulin_mention))

    def test_title_rows_are_not_sections(self) -> None:
        for text in (
            "Перелік лікарських засобів, які підлягають реімбурсації за програмою",
            "населення, станом на 28 серпня 2025 року",
            "ЗАТВЕРДЖЕНО Наказ Міністерства охорони здоров’я України",
        ):
            with self.subTest(text=text):
                self.assertIsNone(builder.detect_section_header((None, text)))

    def test_unknown_roman_heading_fails_closed(self) -> None:
        with self.assertRaises(RuntimeError):
            builder.detect_section_header((None, "II. Медичні вироби"))


class ColumnMappingTests(unittest.TestCase):
    def test_medicines_header_resolves(self) -> None:
        columns = builder.resolve_section_columns(tuple([None] + MEDICINES_HEADER), "medicines")
        self.assertIsNotNone(columns)
        self.assertEqual(columns["inn_name"], 2)
        self.assertEqual(columns["retail_price_uah"], 12)
        self.assertEqual(columns["who_daily_dose"], 13)
        self.assertEqual(columns["reimbursement_pack_uah"], 15)
        self.assertEqual(columns["patient_copay_uah"], 16)

    def test_dosage_form_not_confused_with_release_form(self) -> None:
        columns = builder.resolve_section_columns(tuple([None] + COMBINED_HEADER), "combined_medicines")
        self.assertIsNotNone(columns)
        self.assertEqual(columns["dosage_form"], 4)
        self.assertEqual(columns["release_form_or_primary_pack"], 13)
        self.assertNotEqual(columns["dosage_form"], columns["release_form_or_primary_pack"])

    def test_non_header_row_returns_none(self) -> None:
        self.assertIsNone(builder.resolve_section_columns((None, 1, "МНН", "ТОРГ"), "medicines"))

    def test_missing_column_fails_closed(self) -> None:
        broken = [value for value in MEDICINES_HEADER if not value.startswith("Роздрібна")]
        with self.assertRaises(RuntimeError) as ctx:
            builder.resolve_section_columns(tuple([None] + broken), "medicines")
        self.assertIn("does not match the expected structure", str(ctx.exception))

    def test_duplicated_column_fails_closed(self) -> None:
        duplicated = list(MEDICINES_HEADER)
        duplicated[12] = "Добова доза лікарського засобу"
        duplicated[13] = "Добова доза лікарського засобу"
        with self.assertRaises(RuntimeError):
            builder.resolve_section_columns(tuple([None] + duplicated), "medicines")


class WorkbookParsingTests(BuilderTestCase):
    def test_reordered_columns_are_followed_not_guessed(self) -> None:
        """Swapping two numeric headers must move the values with them, not corrupt them."""
        swapped = list(MEDICINES_HEADER)
        swapped[10], swapped[11] = swapped[11], swapped[10]
        path = write_workbook(self.root / "swapped.xlsx", medicines_header=swapped)
        rows, _ = builder.parse_xlsx(path)
        medicine = next(row for row in rows if row["section"] == "medicines")
        # Header order says column 11 is now retail and column 12 wholesale.
        self.assertEqual(medicine["retail_price_uah"], 101.0)
        self.assertEqual(medicine["wholesale_price_uah"], 134.0)

    def test_numbering_row_is_skipped(self) -> None:
        path = write_workbook(self.root / "book.xlsx", medicines=60, insulins=15, combined=30)
        rows, _ = builder.parse_xlsx(path)
        self.assertEqual(len(rows), 105)
        self.assertFalse([row for row in rows if row["inn_name"].isdigit()])

    def test_missing_register_date_fails(self) -> None:
        path = write_workbook(self.root / "nodate.xlsx", register_line=None)
        with self.assertRaises(RuntimeError) as ctx:
            builder.parse_xlsx(path)
        self.assertIn("register date is missing", str(ctx.exception))

    def test_missing_section_fails(self) -> None:
        path = write_workbook(self.root / "nosec.xlsx", combined=0)
        with self.assertRaises(RuntimeError):
            builder.parse_xlsx(path)


class InputSelectionTests(BuilderTestCase):
    def test_missing_directory(self) -> None:
        with self.assertRaises(RuntimeError) as ctx:
            builder.find_input_workbook(self.root / "absent")
        self.assertIn("directory is missing", str(ctx.exception))

    def test_zero_and_two_files(self) -> None:
        empty = self.root / "empty"
        empty.mkdir()
        with self.assertRaises(RuntimeError):
            builder.find_input_workbook(empty)

        two = self.root / "two"
        write_workbook(two / "a.xlsx")
        write_workbook(two / "b.xlsx")
        with self.assertRaises(RuntimeError):
            builder.find_input_workbook(two)

    def test_exactly_one(self) -> None:
        one = self.root / "one"
        write_workbook(one / "only.xlsx")
        self.assertEqual(builder.find_input_workbook(one).name, "only.xlsx")


# --------------------------------------------------------------------------- B2


class BaselinePreservationTests(BuilderTestCase):
    def test_missing_baseline_fails_closed(self) -> None:
        """B2/B5: no baseline must never silently produce a fresh, stripped database."""
        write_input_dir(self.input_dir)
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build()
        self.assertIn("Baseline manifest is required", str(ctx.exception))

    def test_manifest_without_zip_fails_closed(self) -> None:
        write_input_dir(self.input_dir)
        manifest, _ = make_baseline(self.baseline_dir)
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build("--previous-manifest", str(manifest))
        self.assertIn("Baseline database is required", str(ctx.exception))

    def test_corrupted_baseline_zip_fails_closed(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir)
        zip_path.write_bytes(b"not a zip at all")
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build("--previous-manifest", str(manifest), "--previous-zip", str(zip_path))
        self.assertIn("corrupted", str(ctx.exception))

    def test_extra_tables_and_views_survive(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, reference_prices=12, devices=5)
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        connection = self.built_database()
        self.assertEqual(
            connection.execute("SELECT COUNT(*) FROM reference_price_items").fetchone()[0], 12
        )
        self.assertEqual(
            connection.execute("SELECT COUNT(*) FROM reimbursement_medical_devices").fetchone()[0], 5
        )
        views = {
            row[0]
            for row in connection.execute("SELECT name FROM sqlite_master WHERE type='view'")
        }
        self.assertIn("v_free_items", views)
        summary = self.summary()
        self.assertEqual(summary["carried_over_tables"]["reference_price_items"], 12)
        self.assertIn("v_free_items", summary["carried_over_views"])

    def test_inventory_guard_detects_loss(self) -> None:
        previous = {
            "tables": {"reimbursement_items": 100, "metadata": 5, "reference_price_items": 467},
            "views": ["v_free_items"],
        }
        with self.assertRaises(RuntimeError) as ctx:
            builder.validate_inventory_preserved(
                previous,
                {"tables": {**previous["tables"], "reference_price_items": 0}, "views": ["v_free_items"]},
            )
        self.assertIn("lost all 467 rows", str(ctx.exception))

        with self.assertRaises(RuntimeError):
            builder.validate_inventory_preserved(
                previous, {"tables": previous["tables"], "views": []}
            )

        # Owned tables are allowed to change freely.
        builder.validate_inventory_preserved(
            previous,
            {"tables": {**previous["tables"], "reimbursement_items": 50, "metadata": 18}, "views": ["v_free_items"]},
        )


# --------------------------------------------------------------------------- B3


class VolumeCheckTests(BuilderTestCase):
    def test_volume_check_runs_on_first_manual_build(self) -> None:
        """B3: a cross-pipeline baseline must not disable the relative volume check."""
        write_input_dir(self.input_dir, medicines=60, insulins=15, combined=30)
        manifest, zip_path = make_baseline(
            self.baseline_dir,
            medicines=200, insulins=40, combined=90,
            pipeline_id=None, canonical_source_id=None,
        )
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build(
                "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
                "--allow-pipeline-transition",
            )
        self.assertIn("dropped unexpectedly", str(ctx.exception))

    def test_pipeline_transition_requires_explicit_override(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(
            self.baseline_dir, pipeline_id=None, canonical_source_id=None
        )
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build("--previous-manifest", str(manifest), "--previous-zip", str(zip_path))
        self.assertIn("different pipeline/source", str(ctx.exception))

    def test_same_pipeline_volume_drop_is_blocked(self) -> None:
        write_input_dir(self.input_dir, medicines=60, insulins=15, combined=30)
        manifest, zip_path = make_baseline(self.baseline_dir, medicines=200, insulins=40, combined=90)
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build(
                "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
                "--allow-same-source-date",
            )
        self.assertIn("dropped unexpectedly", str(ctx.exception))


# --------------------------------------------------------------------------- B4 / B5


class DateGateAndVersioningTests(BuilderTestCase):
    def test_same_date_skips_without_override(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, updated_at="2025-08-28")
        self.run_build("--previous-manifest", str(manifest), "--previous-zip", str(zip_path))
        summary = self.summary()
        self.assertFalse(summary["publish"])
        self.assertIn("allow-same-source-date", summary["reason"])
        # The extracted baseline must not be left behind as a build artifact.
        self.assertFalse((self.work_dir / builder.SQLITE_NAME).exists())

    def test_same_date_builds_with_override(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, updated_at="2025-08-28")
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        self.assertTrue((self.work_dir / builder.ZIP_NAME).exists())

    def test_version_must_exceed_existing_releases(self) -> None:
        """B5: never re-issue an already published version."""
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, database_version=2, updated_at="2025-01-01")
        with self.assertRaises(RuntimeError) as ctx:
            self.run_build(
                "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
                "--minimum-database-version", "9",
            )
        self.assertIn("already exists", str(ctx.exception))

    def test_expected_release_tag_matches_manifest(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, database_version=5, updated_at="2025-08-28")
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        summary = self.summary()
        produced = json.loads((self.work_dir / builder.MANIFEST_NAME).read_text(encoding="utf-8"))
        self.assertEqual(summary["database_version"], 6)
        self.assertEqual(summary["expected_release_tag"], "reimbursement-v6")
        self.assertIn("reimbursement-v6/", produced["downloadURL"])
        self.assertFalse(summary["release_created"])
        self.assertFalse(summary["publish"])


# --------------------------------------------------------------------------- B6


class Icd10RetentionTests(BuilderTestCase):
    def test_codes_are_carried_over(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, icd10_rows_per_section=40, updated_at="2025-08-28")
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        summary = self.summary()
        self.assertEqual(summary["icd10_codes_transferred"], 85)
        self.assertEqual(summary["icd10_codes_eligible"], 85)

    def test_retention_below_threshold_fails_closed(self) -> None:
        """B6: losing curated mappings must stop the build, not appear only in a log line."""
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, icd10_rows_per_section=40, updated_at="2025-08-28")
        # Break the stable key on most coded rows: the workbook can no longer match them.
        with zipfile.ZipFile(zip_path) as archive:
            archive.extract(builder.SQLITE_NAME, self.baseline_dir)
        database = self.baseline_dir / builder.SQLITE_NAME
        with sqlite3.connect(database) as connection:
            connection.execute(
                "UPDATE reimbursement_items SET registration_number = registration_number || '-X'"
                " WHERE COALESCE(TRIM(icd10_codes),'') <> '' AND ordinal_no > 4"
            )
        zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(database, builder.SQLITE_NAME)

        with self.assertRaises(RuntimeError) as ctx:
            self.run_build(
                "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
                "--allow-same-source-date",
            )
        message = str(ctx.exception)
        self.assertIn("icd10_codes retention dropped", message)
        self.assertIn("minimum=", message)

    def test_threshold_is_measured_against_eligible_codes(self) -> None:
        baseline = builder.BaselineProfile(
            has_items=True, section_counts={}, pipeline_id=None, canonical_source_id=None,
            icd10_codes_present=699, icd10_codes_eligible=645,
        )
        # 90% of 645 = 581 (not of 699), so duplicated keys are not charged as a regression.
        builder.validate_icd10_retention(transferred=581, baseline=baseline)
        with self.assertRaises(RuntimeError):
            builder.validate_icd10_retention(transferred=580, baseline=baseline)


# --------------------------------------------------------------------------- B7 metadata / schema


class MetadataAndSchemaTests(BuilderTestCase):
    def test_stale_provenance_is_removed_and_reference_prices_kept(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(
            self.baseline_dir,
            updated_at="2025-08-28",
            extra_metadata={
                "source_pdf": "old.pdf",
                "source_json": "old.json",
                "preliminary_rows_inserted_count": "296",
                "main_medicines_count_from_register": "621",
                "reference_prices_count": "467",
                "reference_prices_updated_at": "2026-06-11T00:00:00+03:00",
            },
        )
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        connection = self.built_database()
        keys = {row[0] for row in connection.execute("SELECT key FROM metadata")}
        for stale in ("source_pdf", "source_json", "preliminary_rows_inserted_count",
                      "main_medicines_count_from_register"):
            self.assertNotIn(stale, keys)
        self.assertIn("reference_prices_count", keys)
        self.assertIn("reference_prices_updated_at", keys)
        self.assertEqual(
            connection.execute("SELECT value FROM metadata WHERE key='pipeline_id'").fetchone()[0],
            builder.PIPELINE_ID,
        )
        self.assertIn("source_pdf", self.summary()["removed_stale_metadata_keys"])

    def test_added_columns_keep_declared_types(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, updated_at="2025-08-28")
        with zipfile.ZipFile(zip_path) as archive:
            archive.extract(builder.SQLITE_NAME, self.baseline_dir)
        database = self.baseline_dir / builder.SQLITE_NAME
        with sqlite3.connect(database) as connection:
            connection.execute("ALTER TABLE reimbursement_items DROP COLUMN who_daily_dose")
        zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(database, builder.SQLITE_NAME)

        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        connection = self.built_database()
        types = {
            row["name"]: row["type"]
            for row in connection.execute("SELECT name, type FROM pragma_table_info('reimbursement_items')")
        }
        self.assertEqual(types["who_daily_dose"], "REAL")

    def test_every_item_column_has_a_declared_type(self) -> None:
        self.assertEqual(set(builder.ITEM_COLUMNS), set(builder.ITEM_COLUMN_TYPES))


# --------------------------------------------------------------------------- client contract


class ClientContractTests(BuilderTestCase):
    def test_manifest_and_zip_match_the_swift_client(self) -> None:
        write_input_dir(self.input_dir)
        manifest, zip_path = make_baseline(self.baseline_dir, updated_at="2025-08-28")
        self.run_build(
            "--previous-manifest", str(manifest), "--previous-zip", str(zip_path),
            "--allow-same-source-date",
        )
        produced = json.loads((self.work_dir / builder.MANIFEST_NAME).read_text(encoding="utf-8"))
        self.assertEqual(
            set(produced),
            {"databaseVersion", "schemaVersion", "updatedAt", "minimumAppVersion", "downloadURL", "sha256"},
        )
        self.assertEqual(produced["schemaVersion"], builder.SCHEMA_VERSION)
        self.assertRegex(produced["updatedAt"], r"^\d{4}-\d{2}-\d{2}$")
        with zipfile.ZipFile(self.work_dir / builder.ZIP_NAME) as archive:
            self.assertEqual(archive.namelist(), [builder.SQLITE_NAME])

        import hashlib

        digest = hashlib.sha256((self.work_dir / builder.ZIP_NAME).read_bytes()).hexdigest()
        self.assertEqual(digest, produced["sha256"])

        connection = self.built_database()
        self.assertEqual(connection.execute("PRAGMA integrity_check").fetchone()[0], "ok")
        self.assertGreater(
            connection.execute("SELECT COUNT(*) FROM reimbursement_items").fetchone()[0], 0
        )


if __name__ == "__main__":
    unittest.main()
