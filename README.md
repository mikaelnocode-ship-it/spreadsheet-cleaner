# Spreadsheet cleaner

Turns a messy CSV or Excel export (CRM, shop, form, accounting tool) into a clean Excel file, and lists every value it could not fix so that a person can check it.

## Before → after

| Input (samples/messy-contacts.csv) | Output (samples/cleaned-example.xlsx) |
|---|---|
| ` Alice.Martin@Example.com ` | `alice.martin@example.com` |
| `00353 1 555 0199` | `+35315550199` |
| `14.03.2026`, `2026-02-14`, `31 Mar 2026` | real Excel dates |
| `£1,234.50`, `1 250,00 €`, `2,100` | `1234.50`, `1250.00`, `2100.00` |
| empty rows, duplicate rows | removed |
| `david.kowalski@example`, `soon`, `n/a` | kept as they are and listed in the **Issues** sheet |

Result on the sample: 7 rows read, 1 empty row and 1 duplicate removed, 5 rows kept, 4 values to check.

## What it does

- Reads CSV files (comma, semicolon, tab or pipe separators; UTF-8 or Windows encodings) and `.xlsx` files.
- Trims spaces and removes empty rows and exact duplicates.
- Detects column types from the headers (English and French): email, phone, date, amount.
- Writes three sheets:
  - **Clean**: the data as an Excel table, with date and number formats.
  - **Issues**: the source row, column, value and problem for everything it could not fix.
  - **Summary**: counts of rows read, removed and kept.
- It never deletes a doubtful value: it flags it.

## Use

```bash
pip install -r requirements.txt
python clean.py samples/messy-contacts.csv -o cleaned.xlsx
python clean.py export.xlsx --monthfirst   # for US dates such as 03/31/2026
```

## Tests

```bash
python -m unittest discover -s tests -v
```

## Need your own file cleaned?

Each export is different, so I adapt the rules (columns, formats, merges between files) to your data. Available on Upwork.

## License

MIT
