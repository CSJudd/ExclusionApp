# Client Ingest Guide

This guide explains how to configure per-client file ingest in `clients/<client>.yaml` without changing code.

## Supported Input Types

- `csv`
- `xls`
- `xlsx`
- `xlsm`

Each section (`staff`, `board`, `vendors`) can use:

- `file_type: auto | csv | excel`
- `header_row: auto | <row_index>`
- `skip_rows: <int>`
- `delimiter: "," | "|" | "\t" | auto` (CSV only)
- `true_header_tokens: [ ... ]` (for `header_row: auto`)

Row indexes are 0-based.

## Common Patterns

### 1) Clean CSV with header on first row

```yaml
staff:
  file_type: csv
  header_row: 0
  first_name: "First Name"
  last_name: "Last Name"
  dob: "DOB"
```

### 2) CSV with extra rows before header

```yaml
board:
  file_type: csv
  skip_rows: 2
  name_column: "NAME"
  dob: "DOB"
```

### 3) Noisy Vendor Excel (title/date rows before real header)

```yaml
vendors:
  file_type: auto
  header_row: auto
  true_header_tokens: ["NAME", "CITY", "STATE"]
  entity_name: "NAME"
  address: "ADDRESS1"
  address2: "ADDRESS2"
  city: "CITY"
  state: "STATE"
  zip: "ZIP"
  tax_id: "TAX_ID"
  vendor_id: "Vendor ID"
```

### 4) Pipe-delimited CSV

```yaml
vendors:
  file_type: csv
  delimiter: "|"
  header_row: 0
  entity_name: "Vendor"
```

### 5) Unknown delimiter CSV

```yaml
staff:
  file_type: csv
  delimiter: auto
  header_row: 0
```

### 6) Combined board location column (`CITY, STATE, ZIP`)

```yaml
board:
  file_type: csv
  skip_rows: 2
  name_column: "NAME"
  dob: "DOB"
  address: "ADDRESS"
  city_state_zip: "CITY, STATE, ZIP"
  ssn: "SSN"
```

The engine will parse combined location into separate `city/state/zip` values for matching and reporting.

## Recommended Workflow for New Clients

1. Start with `file_type: auto`.
2. If the header is not row 0, set either:
   - `skip_rows`, or
   - `header_row: auto` + `true_header_tokens`.
3. Run once and inspect `metadata.json` in the run folder.
4. Confirm resolved columns and adjust mappings in YAML.
5. Re-run and validate PDFs/audit output.

## Troubleshooting

- **Wrong columns selected**: set explicit mapping names in YAML.
- **Header still wrong**: provide `header_row: <int>` instead of `auto`.
- **CSV parse issues**: set `delimiter` explicitly.
- **Combined location repeated/odd**: set `city_state_zip` and keep separate `city/state/zip` unset or mapped correctly.
