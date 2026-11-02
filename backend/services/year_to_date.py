"""Annual target achievement from the beginning of the source report's year."""
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from rapidfuzz import fuzz, process

from backend.services.processing import (
    ProcessingError, clean_dashboard_text, find_column, normalize_column, normalize_name,
)
from backend.services.report_input import read_report


def generate_ro_year_to_date(loan_path: Path, target_path: Path, output_path: Path) -> None:
    loans = read_report(loan_path)
    targets = pd.read_excel(target_path)
    loans.columns = [normalize_column(c) for c in loans.columns]
    targets.columns = [normalize_column(c) for c in targets.columns]
    for frame, mapping in [(loans, {
        "LOAN OFFICER": ["LOAN OFFICER", "OFFICER", "LOANOFFICER", "LOAN OFFICER NAME"],
        "DATE": ["DISBURSEMENT DATE", "DISBURSMENT DATE", "BOOKING DATE"],
        "AMOUNT": ["TOTAL LOAN AMOUNT", "LOAN AMOUNT", "AMOUNT"],
    }), (targets, {
        "LOAN OFFICER": ["LOAN OFFICER", "OFFICER", "LOANOFFICER"],
        "MONTHLY TARGET": ["MONTHLY TARGET", "TARGETS", "TARGET"],
    })]:
        for destination, aliases in mapping.items():
            source = find_column(frame.columns, aliases)
            if source is None:
                raise ProcessingError(f"Missing required column: {destination}.")
            frame.rename(columns={source: destination}, inplace=True)
        frame["LOAN OFFICER"] = frame["LOAN OFFICER"].apply(clean_dashboard_text)
        frame.drop(frame[frame["LOAN OFFICER"].eq("") | frame["LOAN OFFICER"].str.fullmatch(r"(?:GRAND )?TOTALS?")].index, inplace=True)

    for frame, column in [(loans, "AMOUNT"), (targets, "MONTHLY TARGET")]:
        frame[column] = pd.to_numeric(frame[column].astype(str).str.upper().str.replace(r"KES|KSH|,", "", regex=True).str.strip(), errors="coerce")
        if frame[column].isna().any() or not frame[column].map(lambda x: float('-inf') < x < float('inf')).all():
            raise ProcessingError(f"Please correct missing or invalid values in {column}.")
    if targets.empty or targets["LOAN OFFICER"].duplicated().any() or targets["MONTHLY TARGET"].lt(0).any():
        raise ProcessingError("Targets must contain unique officers and non-negative monthly targets.")
    loans["DATE"] = pd.to_datetime(loans["DATE"], errors="coerce", format="mixed")
    if loans.empty or loans["DATE"].isna().any():
        raise ProcessingError("Please supply valid Disbursement Date values for every loan.")
    report_date = loans["DATE"].max()
    year_loans = loans[loans["DATE"] >= report_date.replace(month=1, day=1).normalize()].copy()
    lookup = {normalize_name(n): n for n in targets["LOAN OFFICER"]}
    def match_name(name):
        key = normalize_name(name)
        if key in lookup:
            return lookup[key]
        match = process.extractOne(key, list(lookup), scorer=fuzz.token_sort_ratio)
        return lookup[match[0]] if match and match[1] >= 80 else name
    year_loans["LOAN OFFICER"] = year_loans["LOAN OFFICER"].map(match_name)
    amounts = year_loans.groupby("LOAN OFFICER")["AMOUNT"].sum()
    report = targets[["LOAN OFFICER", "MONTHLY TARGET"]].copy()
    # Keep unmatched officers visible rather than silently dropping disbursements.
    unmatched = sorted(set(amounts.index) - set(report["LOAN OFFICER"]))
    report = pd.concat([report, pd.DataFrame({"LOAN OFFICER": unmatched})], ignore_index=True)
    report["ANNUAL TARGET"] = report["MONTHLY TARGET"] * 12
    report["YTD DISBURSEMENT"] = report["LOAN OFFICER"].map(amounts).fillna(0)
    report["REMAINING TO TARGET"] = (report["ANNUAL TARGET"] - report["YTD DISBURSEMENT"]).clip(lower=0)
    report["ACHIEVEMENT %"] = report["YTD DISBURSEMENT"] / report["ANNUAL TARGET"].where(report["ANNUAL TARGET"].gt(0))
    def status(value):
        if pd.isna(value): return "No target"
        return "Target achieved" if value >= 1 else "Close to target" if value >= .8 else "Needs progress" if value >= .5 else "Below target"
    report["STATUS"] = report["ACHIEVEMENT %"].map(status)
    report = report.sort_values("ACHIEVEMENT %", ascending=False, na_position="last")
    totals = {c: report[c].sum() for c in ["MONTHLY TARGET", "ANNUAL TARGET", "YTD DISBURSEMENT"]}
    totals.update({"LOAN OFFICER": "TOTAL", "REMAINING TO TARGET": max(0, totals["ANNUAL TARGET"] - totals["YTD DISBURSEMENT"]), "ACHIEVEMENT %": totals["YTD DISBURSEMENT"] / totals["ANNUAL TARGET"] if totals["ANNUAL TARGET"] > 0 else None, "STATUS": ""})
    report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
    report.to_excel(output_path, sheet_name="YTD vs Annual Target", index=False, startrow=4)
    wb = load_workbook(output_path)
    ws = wb.active
    ws["A1"] = "RO Year to Date vs Annual Target"
    ws["A2"] = f"01 Jan {report_date.year} – {report_date:%d %b %Y} | Annual target = monthly target × 12"
    ws["A3"] = "Achievement: red <50% | amber 50–79% | light green 80–99% | green ≥100% | grey: no target"
    ws["A4"] = "Upload complete year-to-date loan history. Unmatched officers are included without a target."
    colors = {"Target achieved": "86EFAC", "Close to target": "DCFCE7", "Needs progress": "FDE68A", "Below target": "FECACA", "No target": "E5E7EB"}
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row - 1):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=colors[row[6].value])
    for row in [ws[5], ws[ws.max_row]]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.font = Font(bold=True, color="FFFFFF")
    for row in ws.iter_rows(min_row=6):
        for cell in row[1:5]: cell.number_format = '#,##0.00'
        row[5].number_format = '0.00%'
    for col in "ABCDEFG": ws.column_dimensions[col].width = 26
    ws.freeze_panes = "B6"
    ws.auto_filter.ref = f"A5:G{ws.max_row - 1}"
    wb.save(output_path)
