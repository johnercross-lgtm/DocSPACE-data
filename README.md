# DocSPACE-data
Public update data assets for DocSPACE

## Reimbursement database updates

The reimbursement release workflow builds the assets consumed by the DocSPACE iOS client:

- `reimbursement.zip`
- `reimbursement_manifest.json`

The ZIP contains exactly one file: `reimbursement.sqlite`. The manifest SHA-256 is calculated from
the ZIP bytes, matching the client-side verification flow.

Source discovery starts from the Ministry of Health dataset on data.gov.ua:

`https://data.gov.ua/api/3/action/package_show?id=21a6930e-4346-461c-8d80-abeb6f9c0ae2`

The workflow does not treat the dataset page update date as a new database version. It only
publishes when the newest valid XLSX resource has a register date newer than the currently published
manifest. The builder checks the resource name/description, the workbook register date, expected
columns, expected medicine/insulin/combined sections, row counts, SQLite integrity, required tables,
required columns, ZIP contents, and ZIP checksum before a GitHub Release is created.

The public NSZU page remains the official user-facing reference for the full list:

`https://nszu.gov.ua/gromadianam/dostupni-liki-povnii-perelik`
