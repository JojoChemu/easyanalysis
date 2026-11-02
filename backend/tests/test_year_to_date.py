import asyncio
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd
from openpyxl import load_workbook
from fastapi import UploadFile

from backend.services.year_to_date import generate_ro_year_to_date
from backend.services.processing import ProcessingError
from backend.routers.processing import ro_year_to_date_endpoint


class YearToDateTests(unittest.TestCase):
    def test_year_filter_annual_targets_percentages_colors_and_download(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            loans, targets, output = [root / n for n in ["loans.csv", "targets.xlsx", "output.xlsx"]]
            pd.DataFrame({
                "Loan Officer": ["DOE JANE", "JANE DOE", "JANE DOE", "OTHER", "ZERO", "TOTAL"],
                "Disbursement Date": ["2025-12-31", "2026-01-01", "2026-11-01", "2026-10-01", "2026-10-01", "2026-11-01"],
                "Total Loan Amount": [999999, 600, 600, 50, 10, 1260],
            }).to_csv(loans, index=False)
            pd.DataFrame({"Loan Officer": ["JANE DOE", "IDLE", "ZERO"], "Target": [100, 100, 0]}).to_excel(targets, index=False)
            original = targets.read_bytes()
            generate_ro_year_to_date(loans, targets, output)
            ws = load_workbook(output).active
            rows = {r[0]: r for r in list(ws.values)[5:]}
            self.assertEqual(rows["JANE DOE"][1:], (100, 1200, 1200, 0, 1, "Target achieved"))
            self.assertEqual(rows["IDLE"][3:], (0, 1200, 0, "Below target"))
            self.assertEqual(rows["ZERO"][5:], (None, "No target"))
            self.assertEqual(rows["OTHER"][3], 50)
            self.assertIsNone(rows["OTHER"][5])
            self.assertEqual(rows["TOTAL"][3], 1260)
            self.assertEqual(ws["A6"].fill.fgColor.rgb, "0086EFAC")
            self.assertEqual(ws["F6"].number_format, "0.00%")
            self.assertIn("01 Jan 2026", ws["A2"].value)
            self.assertEqual(targets.read_bytes(), original)
            def upload(path):
                return UploadFile(filename=path.name, file=BytesIO(path.read_bytes()))
            response = asyncio.run(ro_year_to_date_endpoint(upload(loans), upload(targets)))
            self.assertEqual(response.status_code, 200)
            self.assertTrue(response.body.startswith(b"PK"))
            self.assertIn("RO_Year_To_Date_vs_Target.xlsx", response.headers["content-disposition"])

    def test_invalid_dates_are_rejected(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            loans, targets = root / "loans.csv", root / "targets.xlsx"
            pd.DataFrame({"Loan Officer": ["JANE"], "Disbursement Date": ["bad"], "Total Loan Amount": [100]}).to_csv(loans, index=False)
            pd.DataFrame({"Loan Officer": ["JANE"], "Target": [100]}).to_excel(targets, index=False)
            with self.assertRaises(ProcessingError):
                generate_ro_year_to_date(loans, targets, root / "output.xlsx")
