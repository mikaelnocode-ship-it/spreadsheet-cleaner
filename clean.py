"""Spreadsheet cleaner: turns a messy CSV or Excel export into a clean Excel file.

Usage:
    python clean.py samples/messy-contacts.csv -o cleaned.xlsx
    python clean.py export.xlsx --monthfirst      # dates written 03/31/2026

What it does:
    - trims spaces, removes empty rows and exact duplicate rows
    - normalises columns detected from their header:
        email  -> lower case, checked
        phone  -> digits and leading +, checked
        date   -> real Excel dates (many input formats)
        amount -> numbers (1 234,50 / 1,234.50 / £1,234.50 ...)
    - writes three sheets: Clean, Issues (every value it could not fix), Summary
"""
import argparse
import csv
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

EMAIL_RE = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$")
KINDS = {
    "email": ("email", "e-mail", "mail"),
    "phone": ("phone", "tel", "mobile", "téléphone", "telephone"),
    "date": ("date", "created", "joined", "due"),
    "amount": ("amount", "price", "total", "revenue", "montant", "prix"),
}
DATE_FORMATS_DAYFIRST = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%Y-%m-%d",
                         "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y"]
DATE_FORMATS_MONTHFIRST = ["%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y"] + DATE_FORMATS_DAYFIRST[4:]


# ---------- value normalisers (each returns (value, problem or None)) ----------

def clean_text(v):
    if v is None:
        return None
    if isinstance(v, (int, float, date)):
        return v
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def norm_email(v):
    if v is None:
        return None, None
    s = str(v).strip().lower().replace(" ", "")
    return s, None if EMAIL_RE.match(s) else "invalid email"


def norm_phone(v):
    if v is None:
        return None, None
    s = str(v).strip()
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    plus = s.startswith("+") or s.startswith("00")
    digits = re.sub(r"\D", "", s)
    if s.startswith("00"):
        digits = digits[2:]
    out = ("+" if plus else "") + digits
    return out, None if 8 <= len(digits) <= 15 else "phone number too short or too long"


def norm_date(v, monthfirst=False):
    if v is None:
        return None, None
    if isinstance(v, datetime):
        return v.date(), None
    if isinstance(v, date):
        return v, None
    s = str(v).strip()
    for fmt in (DATE_FORMATS_MONTHFIRST if monthfirst else DATE_FORMATS_DAYFIRST):
        try:
            return datetime.strptime(s, fmt).date(), None
        except ValueError:
            continue
    return s, "unrecognised date"


def norm_amount(v):
    if v is None:
        return None, None
    if isinstance(v, (int, float)):
        return float(v), None
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not re.search(r"\d", s):
        return str(v), "not a number"
    if "," in s and "." in s:            # the last separator is the decimal one
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:                        # 1,5 or 1,234
        head, _, tail = s.rpartition(",")
        s = s.replace(",", "") if len(tail) == 3 and head else s.replace(",", ".")
    try:
        return float(s), None
    except ValueError:
        return str(v), "not a number"


# ---------- reading ----------

def read_rows(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        ws = load_workbook(path, data_only=True).active
        return [list(r) for r in ws.iter_rows(values_only=True)]
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    return [row for row in csv.reader(text.splitlines(), dialect)]


def detect_kind(header):
    h = (header or "").lower()
    for kind, words in KINDS.items():
        if any(w in h for w in words):
            return kind
    return "text"


# ---------- cleaning ----------

def clean(rows, monthfirst=False):
    headers = [clean_text(h) or f"Column {i + 1}" for i, h in enumerate(rows[0])]
    kinds = [detect_kind(h) for h in headers]
    out, issues, seen = [], [], set()
    stats = {"rows read": len(rows) - 1, "empty rows removed": 0, "duplicates removed": 0}
    for n, row in enumerate(rows[1:], start=2):
        row = list(row) + [None] * (len(headers) - len(row))
        cells = [clean_text(v) for v in row[:len(headers)]]
        if all(c is None for c in cells):
            stats["empty rows removed"] += 1
            continue
        fixed = []
        for h, kind, v in zip(headers, kinds, cells):
            if kind == "email":
                v, prob = norm_email(v)
            elif kind == "phone":
                v, prob = norm_phone(v)
            elif kind == "date":
                v, prob = norm_date(v, monthfirst)
            elif kind == "amount":
                v, prob = norm_amount(v)
            else:
                prob = None
            if prob:
                issues.append((n, h, v, prob))
            fixed.append(v)
        key = tuple(str(x).lower() for x in fixed)
        if key in seen:
            stats["duplicates removed"] += 1
            continue
        seen.add(key)
        out.append(fixed)
    stats["rows kept"] = len(out)
    stats["values to check"] = len(issues)
    return headers, kinds, out, issues, stats


# ---------- writing ----------

def write_xlsx(headers, kinds, rows, issues, stats, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Clean"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for j, kind in enumerate(kinds, start=1):
        letter = ws.cell(row=1, column=j).column_letter
        ws.column_dimensions[letter].width = max(12, min(40, len(headers[j - 1]) + 6))
        fmt = {"date": "dd/mm/yyyy", "amount": "#,##0.00"}.get(kind)
        if fmt:
            for (c,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                if not isinstance(c.value, str):
                    c.number_format = fmt
    if rows:
        tab = Table(displayName="CleanData", ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tab)
    ws.freeze_panes = "A2"

    iss = wb.create_sheet("Issues")
    iss.append(["Source row", "Column", "Value", "Problem"])
    for i in issues:
        iss.append(list(i))
    red = PatternFill("solid", fgColor="FBEDED")
    for c in iss[1]:
        c.font = Font(bold=True)
    for row in iss.iter_rows(min_row=2):
        row[3].fill = red
    for letter, w in zip("ABCD", (11, 18, 30, 34)):
        iss.column_dimensions[letter].width = w

    summ = wb.create_sheet("Summary")
    for k, v in stats.items():
        summ.append([k.capitalize(), v])
    summ.column_dimensions["A"].width = 22
    wb.save(path)


def main(argv=None):
    p = argparse.ArgumentParser(description="Clean a messy CSV/Excel export.")
    p.add_argument("input")
    p.add_argument("-o", "--output", default="cleaned.xlsx")
    p.add_argument("--monthfirst", action="store_true", help="dates like 03/31/2026")
    a = p.parse_args(argv)
    headers, kinds, rows, issues, stats = clean(read_rows(a.input), a.monthfirst)
    write_xlsx(headers, kinds, rows, issues, stats, a.output)
    print(" | ".join(f"{k}: {v}" for k, v in stats.items()))
    return stats


if __name__ == "__main__":
    main()
