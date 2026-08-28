# Manual reimbursement input

Place **exactly one** official reimbursement XLSX file in this directory before starting the
`Update reimbursement database` workflow (`workflow_dispatch`). The workflow never downloads the
source file: it reads only what is committed here.

## Requirements for the file

- Extension `.xlsx` (case-insensitive). Exactly one such file may be present — zero or two files
  fail the build.
- The workbook must contain its register date as `станом на <дата>` inside the first 20 rows of the
  first worksheet. The build derives the database date from that text, so a workbook without it
  fails.
- The first worksheet must contain the expected reimbursement columns and all three sections:
  `I. Лікарські засоби…`, `II. Препарати інсуліну`, `III. Комбіновані лікарські засоби`.

## What the build does with it

1. Downloads the previous GitHub Release as a **baseline only** — never as a source of new data.
2. Rewrites `reimbursement_items` from the XLSX and carries over everything else from the baseline
   (`reference_price_items`, `reimbursement_medical_devices`, views, curated `icd10_codes`).
3. Validates volume, ICD-10 retention, integrity, ZIP contents and checksum.
4. Uploads `reimbursement.sqlite`, `reimbursement.zip`, `reimbursement_manifest.json` and
   `build_summary.json` as workflow artifacts.

**No GitHub Release is created.** Publishing stays manual: use the tag printed as
`expected_release_tag` in `build_summary.json`, otherwise the `downloadURL` inside the manifest will
not match the release and the iOS client will fail to download the update.

## Overrides

All of these are off by default and must be enabled explicitly in the workflow inputs:

- `allow_pipeline_transition` — the previous release came from a different pipeline/source
  (for example the PDF-based `reimbursement-v4`). Volume checks still run against the previous
  database's real counts.
- `allow_same_source_date` — rebuild from a workbook whose register date is not newer than the
  published one (corrected file for the same date).
- `allow_missing_baseline` — there is genuinely no previous release yet. Never use this to work
  around a failed baseline download.
