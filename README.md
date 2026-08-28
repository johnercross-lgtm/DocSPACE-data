# DocSPACE-data
Public update data assets for DocSPACE

## Reimbursement database updates

The reimbursement release workflow builds the assets consumed by the DocSPACE iOS client:

- `reimbursement.zip`
- `reimbursement_manifest.json`

The ZIP contains exactly one file: `reimbursement.sqlite`. The manifest SHA-256 is calculated from
the ZIP bytes, matching the client-side verification flow.

The workflow is manual-input only. Place exactly one official XLSX file in
`input/reimbursement/` before starting `workflow_dispatch`. The workflow does not discover or
download source files from data.gov.ua or NSZU. It checks the workbook register date, expected
columns, expected medicine/insulin/combined sections, row counts, SQLite integrity, required tables,
required columns, ZIP contents, and ZIP checksum. The previous GitHub Release is used only as a
baseline for validation and curated `icd10_codes` transfer.

GitHub Release creation is currently disabled. Successful runs upload the generated SQLite, ZIP,
manifest, and build summary as GitHub Actions artifacts.

## Tests

The pipeline guards (baseline preservation, volume checks, versioning, ICD-10 retention, parser
and numeric-format guards) are covered by tests that need no extra dependencies beyond
`requirements.txt`:

```
python -m unittest discover -s tests -t .
```

The same command runs in the workflow before the build, so a broken guard fails the run instead of
producing an artifact.

The public NSZU page remains the official user-facing reference for the full list:

`https://nszu.gov.ua/gromadianam/dostupni-liki-povnii-perelik`
