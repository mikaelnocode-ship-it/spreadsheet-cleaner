import sys
import unittest
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import clean  # noqa: E402


class Normalisers(unittest.TestCase):
    def test_email(self):
        self.assertEqual(clean.norm_email(" Bob@Ex.COM "), ("bob@ex.com", None))
        self.assertEqual(clean.norm_email("bob@ex")[1], "invalid email")

    def test_phone(self):
        self.assertEqual(clean.norm_phone("00353 1 555 0199"), ("+35315550199", None))
        self.assertEqual(clean.norm_phone("06 12 34 56 78"), ("0612345678", None))
        self.assertIsNotNone(clean.norm_phone("+48 22 123")[1])

    def test_date(self):
        self.assertEqual(clean.norm_date("03/02/2026")[0], date(2026, 2, 3))
        self.assertEqual(clean.norm_date("03/02/2026", monthfirst=True)[0], date(2026, 3, 2))
        self.assertEqual(clean.norm_date("31 Mar 2026")[0], date(2026, 3, 31))
        self.assertEqual(clean.norm_date("14.03.2026")[0], date(2026, 3, 14))
        self.assertEqual(clean.norm_date("soon")[1], "unrecognised date")

    def test_amount(self):
        self.assertEqual(clean.norm_amount("£1,234.50")[0], 1234.5)
        self.assertEqual(clean.norm_amount("1 250,00 €")[0], 1250.0)
        self.assertEqual(clean.norm_amount("1.234,56")[0], 1234.56)
        self.assertEqual(clean.norm_amount("2,100")[0], 2100.0)
        self.assertEqual(clean.norm_amount("12,5")[0], 12.5)
        self.assertEqual(clean.norm_amount("n/a")[1], "not a number")


class EndToEnd(unittest.TestCase):
    def test_sample(self):
        out = ROOT / "tests" / "_out.xlsx"
        stats = clean.main([str(ROOT / "samples" / "messy-contacts.csv"), "-o", str(out)])
        try:
            self.assertEqual(stats["rows read"], 7)
            self.assertEqual(stats["empty rows removed"], 1)
            self.assertEqual(stats["duplicates removed"], 1)
            self.assertEqual(stats["rows kept"], 5)
            wb = load_workbook(out)
            self.assertEqual(wb["Clean"]["A1"].value, "Name")
            self.assertEqual(wb["Clean"]["B2"].value, "alice.martin@example.com")
            problems = [r[3] for r in wb["Issues"].iter_rows(min_row=2, values_only=True)]
            self.assertIn("invalid email", problems)
            self.assertIn("unrecognised date", problems)
            self.assertIn("not a number", problems)
        finally:
            out.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
