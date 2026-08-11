"""File-path adapters for the four supplied report scripts.

The transformation logic below intentionally follows the original Colab scripts.
Only interactive upload/download calls are replaced with function arguments.
"""

from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process


EXPECTED_COLUMNS = ["Loan Officer", "Total Loan Amount", "Disbursement Date"]


class ProcessingError(Exception):
    """An expected, user-correctable report problem."""


def _require_columns(df: pd.DataFrame, required: list[str]) -> None:
    """Validate only the columns the original script accesses directly."""
    available = {str(column).strip().upper() for column in df.columns}
    if not {column.strip().upper() for column in required}.issubset(available):
        raise ProcessingError(
            "This report appears to be invalid.\n\nExpected columns:\n"
            + "\n".join(EXPECTED_COLUMNS)
        )


# ================================ ORIGINAL CLEANING HELPERS ================================

def clean_text(series):
    return (
        series.fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )


def clean_money(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    )


def clean_branch(series):
    return (
        clean_text(series)
        .str.replace(" BRANCH", "", regex=False)
        .str.replace("BRANCH", "", regex=False)
        .str.strip()
    )


def clean_account(series):
    return series.fillna("").astype(str).str.replace(".0", "", regex=False).str.strip()


def clean_date(series):
    s = series.astype(str).str.strip().replace(["", "nan", "NaN", "None"], pd.NA)
    dates = pd.to_datetime(s, errors="coerce", dayfirst=True)
    failed = dates.isna()
    if failed.any():
        dates.loc[failed] = pd.to_datetime(s.loc[failed], errors="coerce")
    return dates


def clean_all_loans(input_path: Path, output_path: Path) -> None:
    """Faithful file-based version of 01_clean_all_loans.py."""
    try:
        df = pd.read_csv(input_path)
        _require_columns(df, ["Account Number", "Branch", "Loan Officer", "Total Loan Amount", "Disbursement Date"])
        df["Account Number"] = df["Account Number"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        duplicates = df["Account Number"].duplicated().sum()
        print("Duplicate Account Numbers:", duplicates)
        df = df.dropna(how="all")
        columns_to_drop = [
            "Phone Number", "National Id Number", "Number Of Dependents",
            "IsPersonWithDisability", "ImpairmentType",
        ]
        df = df.drop(columns=columns_to_drop)
        money_columns = [
            "Total Loan Amount", "Total Interest Amount", "Outstanding Principal",
            "Outstanding Interest", "Outstanding Fees", "Outstanding Penalty",
            "Total Outstanding Amount", "Total In Arrears",
        ]
        for col in money_columns:
            if col in df.columns:
                df[col] = clean_money(df[col])
        text_columns = [
            "Client Name", "Branch", "Loan Officer", "Loan Product",
            "Loan Status", "Member Type", "Group",
        ]
        for col in text_columns:
            if col in df.columns:
                df[col] = clean_text(df[col])
        date_columns = ["Creation Date", "Disbursement Date", "Expiry Date"]
        for col in date_columns:
            if col in df.columns:
                df[col] = clean_date(df[col])
        df["Branch"] = clean_branch(df["Branch"])
        df = df[
            ~df.astype(str).apply(
                lambda row: row.str.contains("TOTAL", case=False, na=False).any(), axis=1
            )
        ]
        print("Rows:", len(df))
        print("Duplicate Accounts:", df["Account Number"].duplicated().sum())
        print("Missing Branch:", df["Branch"].isna().sum())
        print("Missing Loan Officer:", df["Loan Officer"].isna().sum())
        df.to_csv(output_path, index=False)
    except (KeyError, ValueError, TypeError) as error:
        raise ProcessingError(_friendly_error()) from error


def clean_portfolio_risk(input_path: Path, output_path: Path) -> None:
    """Faithful file-based version of 02_clean_portfolioriskreport.py."""
    try:
        df = pd.read_csv(input_path)
        _require_columns(df, ["Account Number", "Branch"])
        df["Account Number"] = clean_account(df["Account Number"])
        print("Blank Account Numbers:", df["Account Number"].eq("").sum())
        print("Duplicate Account Numbers:", df["Account Number"].duplicated().sum())
        df = df.dropna(how="all")
        df["Branch"] = clean_branch(df["Branch"])
        columns_to_drop = [
            "Number Of Dependents", "MembershipPolicy", "IsPersonWithDisability",
            "ImpairmentType", "Group",
        ]
        df = df.drop(columns=columns_to_drop)
        text_columns = ["Client Name", "Branch", "LoanOfficerName"]
        for col in text_columns:
            if col in df.columns:
                df[col] = clean_text(df[col])
        date_columns = [" LastRepaymentDate "]
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        money_columns = [
            "Loan Amount", "PortfolioAtRisk(PAR)", "Outstanding Principal",
            "PrincipalInArrears", "InterestInArrears", "LastRepaymentAmount",
            "Total In Arrears",
        ]
        for col in money_columns:
            if col in df.columns:
                df[col] = clean_money(df[col])
        df = df[
            ~df.astype(str).apply(
                lambda row: row.str.contains("TOTAL", case=False, na=False).any(), axis=1
            )
        ]
        print("Rows:", len(df))
        print("Missing Branches:", df["Branch"].eq("").sum())
        print("Missing Loan Officers:", df["LoanOfficerName"].eq("").sum())
        print("Missing PAR Values:", df["PortfolioAtRisk(PAR)"].isna().sum())
        df.to_csv(output_path, index=False)
    except (KeyError, ValueError, TypeError) as error:
        raise ProcessingError(_friendly_error()) from error


def clean_loan_arrears(input_path: Path, output_path: Path) -> None:
    """Faithful file-based version of 04_clean_loan_arrears.py."""
    try:
        df = pd.read_csv(input_path, dtype={"Account Number": str})
        df["Account Number"].head(10)
        pd.set_option("display.float_format", "{:.0f}".format)
        df["Account Number"].head(10)
        df.columns = df.columns.str.strip()
        _require_columns(df, ["Account Number", "LoanOfficerName", "Branch"])
        print(df.columns.tolist())
        df = df[df["LoanOfficerName"].astype(str).str.upper().ne("TOTAL")]
        df = df.dropna(how="all")
        df["Account Number"] = clean_account(df["Account Number"])
        columns_to_drop = [
            "ClientReferenceID", "Phone Number", "Number Of Dependents",
            "IsPersonWithDisability", "ImpairmentType", "National ID Number",
            "Group", "Outstanding Charges", "ChargeInArrears", "PenaltyOnArrears",
        ]
        df = df.drop(columns=columns_to_drop, errors="ignore")
        for col in ["LoanOfficerName", "Branch", "Client Name"]:
            if col in df.columns:
                df[col] = clean_text(df[col])
        for col in [
            "Loan Amount", "Outstanding Principal", "Outstanding Interest",
            "PrincipalInArrears", "InterestInArrears", "Total In Arrears",
        ]:
            if col in df.columns:
                df[col] = clean_money(df[col])
        for col in ["LoanDisbursmentDate", "Loan Expire Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        df["Branch"] = clean_branch(df["Branch"])
        print("Rows:", len(df))
        print("Duplicate Accounts:", df["Account Number"].duplicated().sum())
        print("Missing Branches:", df["Branch"].eq("").sum())
        print("Missing Loan Officers:", df["LoanOfficerName"].eq("").sum())
        print("Missing Total Arrears:", df["Total In Arrears"].isna().sum())
        df.to_csv(output_path, index=False)
    except (KeyError, ValueError, TypeError) as error:
        raise ProcessingError(_friendly_error()) from error


def normalize_column(col):
    col = str(col).strip().upper().replace("_", " ")
    return " ".join(col.split())


def find_column(columns, possibilities):
    for option in possibilities:
        if option in columns:
            return option
    return None


def clean_dashboard_text(text):
    if pd.isna(text):
        return ""
    text = str(text).upper().strip()
    text = re.sub(r"\s+", " ", text)
    return re.sub(r"[^\w\s]", "", text)


def normalize_name(name):
    name = clean_dashboard_text(name)
    words = name.split()
    cleaned = []
    for word in words:
        if len(word) == 1:
            continue
        if word in ["MR", "MRS", "MS", "MISS", "DR", "REV", "PASTOR"]:
            continue
        cleaned.append(word)
    return " ".join(sorted(cleaned))


def clean_repayment_report(input_path: Path, output_path: Path) -> None:
    """File-based version of Fyness_Clean_LoanRepaymentReport_for_PowerBI.ipynb."""
    try:
        if input_path.suffix.lower() == ".csv":
            df = pd.read_csv(input_path)
        else:
            df = pd.read_excel(input_path)

        df.columns = [re.sub(r"\s+", " ", str(column).strip()) for column in df.columns]
        required = [
            "Repayment Date", "Transaction Number", "Account Number", "Client Name",
            "Branch", "Loan Officer", "PrincipalAmountRepaid", "InterestAmountRepaid",
            "Charges", "TotalAmountRepaid",
        ]
        _require_columns(df, required)
        # Preserve the notebook's exact, case-sensitive field names after whitespace cleanup.
        missing = [column for column in required if column not in df.columns]
        if missing:
            raise ProcessingError("The following required columns are missing:\n" + "\n".join(missing))

        df = df.dropna(how="all").copy()
        df["Account Number"] = clean_account(df["Account Number"])
        df["Client Name"] = clean_text(df["Client Name"])
        df["Branch"] = clean_branch(df["Branch"])
        df["Loan Officer"] = clean_text(df["Loan Officer"])
        df["Transaction Number"] = df["Transaction Number"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        df["Repayment Date"] = pd.to_datetime(df["Repayment Date"], errors="coerce")

        money_columns = ["PrincipalAmountRepaid", "InterestAmountRepaid", "Charges", "TotalAmountRepaid"]
        for column in money_columns:
            df[column] = pd.to_numeric(
                df[column].fillna("").astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("KES", "", regex=False)
                .str.replace("KSH", "", regex=False)
                .str.strip(), errors="coerce"
            )

        df = df[df["Repayment Date"].notna()]
        df = df[df["Account Number"].ne("")]
        df = df[df["Branch"].ne("")]
        df = df[df["Loan Officer"].ne("")].copy()
        df["Year"] = df["Repayment Date"].dt.year
        df["Month Number"] = df["Repayment Date"].dt.month
        df["Month"] = df["Repayment Date"].dt.strftime("%B")
        df["Month Year"] = df["Repayment Date"].dt.strftime("%b %Y")
        df["Month Start"] = df["Repayment Date"].dt.to_period("M").dt.to_timestamp()
        df.to_excel(output_path, index=False)
        df.to_csv(output_path.with_suffix(".csv"), index=False)
        _format_repayment_report(output_path)
    except ProcessingError:
        raise
    except (KeyError, ValueError, TypeError, ImportError) as error:
        raise ProcessingError(_friendly_error()) from error


def _format_repayment_report(output_path: Path) -> None:
    workbook = load_workbook(output_path)
    worksheet = workbook.active
    worksheet.insert_rows(1, 2)
    worksheet["A1"] = "FYNESS NOTES CREDIT LIMITED"
    worksheet["A2"] = "Cleaned Loan Repayment Report - Power BI Source"
    last_column = worksheet.max_column
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    blue = PatternFill(fill_type="solid", start_color="1F4E78")
    header_fill = PatternFill(fill_type="solid", start_color="4472C4")
    worksheet["A1"].fill = blue
    worksheet["A1"].font = Font(bold=True, color="FFFFFF", size=18)
    worksheet["A1"].alignment = Alignment(horizontal="center")
    worksheet["A2"].font = Font(bold=True, italic=True, size=12)
    worksheet["A2"].alignment = Alignment(horizontal="center")
    for cell in worksheet[3]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if cell.value in ["PrincipalAmountRepaid", "InterestAmountRepaid", "Charges", "TotalAmountRepaid"]:
            for row in range(4, worksheet.max_row + 1):
                worksheet.cell(row=row, column=cell.column).number_format = "#,##0.00"
        if cell.value in ["Repayment Date", "Month Start"]:
            for row in range(4, worksheet.max_row + 1):
                worksheet.cell(row=row, column=cell.column).number_format = "dd-mmm-yyyy"
    for column in worksheet.columns:
        letter = get_column_letter(column[0].column)
        width = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
        worksheet.column_dimensions[letter].width = min(width + 3, 35)
    worksheet.freeze_panes = "A4"
    workbook.save(output_path)


def generate_ro_monthly_customer_mix(loan_path: Path, target_path: Path, output_path: Path) -> None:
    """File-based version of RO_Monthly_Performance_Customer_Mix.ipynb."""
    try:
        loans = pd.read_csv(loan_path)
        targets = pd.read_excel(target_path)
        loans.columns = [normalize_column(column) for column in loans.columns]
        targets.columns = [normalize_column(column) for column in targets.columns]
        loan_officer_col = find_column(loans.columns, ["LOAN OFFICER", "OFFICER", "LOANOFFICER", "LOAN OFFICER NAME"])
        date_col = find_column(loans.columns, ["DISBURSEMENT DATE", "DISBURSMENT DATE", "DATE", "BOOKING DATE"])
        amount_col = find_column(loans.columns, ["TOTAL LOAN AMOUNT", "LOAN AMOUNT", "AMOUNT"])
        client_ref_col = find_column(loans.columns, ["CLIENT REFERENCE ID", "CLIENT REF ID", "CLIENT REFERENCE"])
        target_officer_col = find_column(targets.columns, ["LOAN OFFICER", "OFFICER", "LOANOFFICER"])
        target_col = find_column(targets.columns, ["TARGETS", "TARGET", "MONTHLY TARGET"])
        ro_code_col = find_column(targets.columns, ["RO CODE", "CODE"])
        if not all([loan_officer_col, date_col, amount_col, client_ref_col, target_officer_col, target_col, ro_code_col]):
            raise ProcessingError(_friendly_error())
        loans = loans.rename(columns={loan_officer_col: "LOAN OFFICER", date_col: "DISBURSEMENT DATE", amount_col: "TOTAL LOAN AMOUNT", client_ref_col: "CLIENT REFERENCE ID"})
        targets = targets.rename(columns={target_officer_col: "LOAN OFFICER", target_col: "TARGETS", ro_code_col: "RO CODE"})
        loans["LOAN OFFICER"] = loans["LOAN OFFICER"].apply(clean_dashboard_text)
        targets["LOAN OFFICER"] = targets["LOAN OFFICER"].apply(clean_dashboard_text)
        loans["CLIENT REFERENCE ID"] = loans["CLIENT REFERENCE ID"].apply(lambda value: "" if pd.isna(value) else re.sub(r"\.0$", "", str(value).strip()))
        for column in ["TOTAL LOAN AMOUNT"]:
            loans[column] = pd.to_numeric(loans[column].astype(str).str.replace(",", "", regex=False).str.replace("KES", "", regex=False).str.replace("KSH", "", regex=False).str.strip(), errors="coerce")
        targets["TARGETS"] = pd.to_numeric(targets["TARGETS"].astype(str).str.replace(",", "", regex=False).str.replace("KES", "", regex=False).str.replace("KSH", "", regex=False).str.strip(), errors="coerce")
        loans["DISBURSEMENT DATE"] = pd.to_datetime(loans["DISBURSEMENT DATE"], errors="coerce")
        report_date = loans["DISBURSEMENT DATE"].max()
        if pd.isna(report_date):
            raise ProcessingError("No valid Disbursement Date values were found.")
        report_date = report_date.normalize()
        start_of_month = report_date.replace(day=1)
        targets = targets.drop_duplicates(subset="LOAN OFFICER").reset_index(drop=True)
        official_lookup = {normalize_name(officer): officer for officer in targets["LOAN OFFICER"]}
        changes = {}
        for loan_name in loans["LOAN OFFICER"].unique():
            normalized = normalize_name(loan_name)
            if normalized in official_lookup:
                changes[loan_name] = official_lookup[normalized]
                continue
            match = process.extractOne(normalized, list(official_lookup), scorer=fuzz.token_sort_ratio)
            if match and match[1] >= 80:
                changes[loan_name] = official_lookup[match[0]]
        loans["LOAN OFFICER"] = loans["LOAN OFFICER"].replace(changes)
        month_data = loans[(loans["DISBURSEMENT DATE"] >= start_of_month) & (loans["DISBURSEMENT DATE"] <= report_date)].copy()
        history = loans[loans["CLIENT REFERENCE ID"].ne("")].groupby("CLIENT REFERENCE ID")["DISBURSEMENT DATE"].min().rename("FIRST DISBURSEMENT DATE").reset_index()
        month_data = month_data.merge(history, on="CLIENT REFERENCE ID", how="left")
        month_data["CUSTOMER TYPE"] = np.where(month_data["FIRST DISBURSEMENT DATE"] < start_of_month, "REPEAT", "NEW")
        counts = month_data[month_data["CLIENT REFERENCE ID"].ne("")].groupby(["LOAN OFFICER", "CUSTOMER TYPE"])["CLIENT REFERENCE ID"].nunique().unstack(fill_value=0).reset_index()
        amounts = month_data.groupby(["LOAN OFFICER", "CUSTOMER TYPE"])["TOTAL LOAN AMOUNT"].sum().unstack(fill_value=0).reset_index()
        for frame, columns in [(counts, ["NEW", "REPEAT"]), (amounts, ["NEW", "REPEAT"])]:
            for column in columns:
                if column not in frame.columns:
                    frame[column] = 0
        counts = counts.rename(columns={"NEW": "NEW CUSTOMERS", "REPEAT": "REPEAT CUSTOMERS"})
        counts["TOTAL CUSTOMERS"] = counts["NEW CUSTOMERS"] + counts["REPEAT CUSTOMERS"]
        amounts = amounts.rename(columns={"NEW": "NEW CUSTOMER DISBURSEMENT", "REPEAT": "REPEAT CUSTOMER DISBURSEMENT"})
        mtd = month_data.groupby("LOAN OFFICER")["TOTAL LOAN AMOUNT"].sum().reset_index().rename(columns={"TOTAL LOAN AMOUNT": "MTD DISBURSEMENT"})
        numeric = ["MTD DISBURSEMENT", "NEW CUSTOMERS", "REPEAT CUSTOMERS", "TOTAL CUSTOMERS", "NEW CUSTOMER DISBURSEMENT", "REPEAT CUSTOMER DISBURSEMENT"]
        report = targets.copy().merge(mtd, on="LOAN OFFICER", how="left").merge(counts, on="LOAN OFFICER", how="left").merge(amounts, on="LOAN OFFICER", how="left")
        for column in numeric:
            report[column] = report[column].fillna(0)
        report["VARIANCE"] = report["MTD DISBURSEMENT"] - report["TARGETS"]
        report["ACHIEVEMENT %"] = np.where(report["TARGETS"] > 0, report["MTD DISBURSEMENT"] / report["TARGETS"], 0)
        report["STATUS"] = report["ACHIEVEMENT %"].apply(lambda value: "Excellent" if value >= 1 else "Good" if value >= .8 else "Average" if value >= .5 else "Needs Improvement")
        report = report.sort_values("ACHIEVEMENT %", ascending=False).reset_index(drop=True)
        report.insert(0, "RANK", range(1, len(report) + 1))
        totals = {column: report[column].sum() for column in ["TARGETS"] + numeric + ["VARIANCE"]}
        totals.update({"RANK": "", "RO CODE": "", "LOAN OFFICER": "TOTAL", "ACHIEVEMENT %": totals["MTD DISBURSEMENT"] / totals["TARGETS"] if totals["TARGETS"] else 0, "STATUS": ""})
        report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
        report = report[["RANK", "RO CODE", "LOAN OFFICER", "TARGETS", "MTD DISBURSEMENT", "VARIANCE", "ACHIEVEMENT %", "STATUS", "NEW CUSTOMERS", "REPEAT CUSTOMERS", "TOTAL CUSTOMERS", "NEW CUSTOMER DISBURSEMENT", "REPEAT CUSTOMER DISBURSEMENT"]]
        report.to_excel(output_path, index=False)
        _format_ro_monthly_mix(output_path, report_date)
    except ProcessingError:
        raise
    except (KeyError, ValueError, TypeError, ImportError) as error:
        raise ProcessingError(_friendly_error()) from error


def _format_ro_monthly_mix(output_path: Path, report_date) -> None:
    workbook = load_workbook(output_path)
    worksheet = workbook.active
    worksheet.insert_rows(1, 2)
    worksheet["A1"] = "FYNESS NOTES CREDIT LIMITED"
    worksheet["A2"] = f"RO Monthly Performance & Customer Mix | {report_date.strftime('%B %Y')} | MTD as at {report_date.strftime('%d %B %Y')}"
    last_column = worksheet.max_column
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    blue = PatternFill(fill_type="solid", start_color="1F4E78")
    header_fill = PatternFill(fill_type="solid", start_color="4472C4")
    worksheet["A1"].fill = blue
    worksheet["A1"].font = Font(bold=True, color="FFFFFF", size=18)
    worksheet["A1"].alignment = Alignment(horizontal="center")
    worksheet["A2"].font = Font(bold=True, italic=True, size=12)
    worksheet["A2"].alignment = Alignment(horizontal="center")
    for cell in worksheet[3]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    currency = ["TARGETS", "MTD DISBURSEMENT", "VARIANCE", "NEW CUSTOMER DISBURSEMENT", "REPEAT CUSTOMER DISBURSEMENT"]
    counts = ["NEW CUSTOMERS", "REPEAT CUSTOMERS", "TOTAL CUSTOMERS"]
    for cell in worksheet[3]:
        if cell.value in currency:
            for row in range(4, worksheet.max_row + 1): worksheet.cell(row=row, column=cell.column).number_format = "#,##0.00"
        if cell.value in counts:
            for row in range(4, worksheet.max_row + 1): worksheet.cell(row=row, column=cell.column).number_format = "#,##0"
        if cell.value == "ACHIEVEMENT %":
            for row in range(4, worksheet.max_row + 1): worksheet.cell(row=row, column=cell.column).number_format = "0.00%"
    for column in worksheet.columns:
        letter = get_column_letter(column[0].column)
        width = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
        worksheet.column_dimensions[letter].width = min(width + 3, 35)
    worksheet.freeze_panes = "D4"
    total_row = worksheet.max_row
    for cell in worksheet[total_row]:
        cell.fill = blue
        cell.font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows():
        for cell in row: cell.border = border
    workbook.save(output_path)


def _friendly_error():
    return "This report appears to be invalid.\n\nExpected columns:\n" + "\n".join(EXPECTED_COLUMNS)


def generate_ro_dashboard(loan_path: Path, target_path: Path, output_path: Path) -> None:
    """Faithful file-based version of ro_performance_dashboard_automation.py."""
    try:
        loans = pd.read_csv(loan_path)
        targets = pd.read_excel(target_path)
        loans.columns = [normalize_column(c) for c in loans.columns]
        targets.columns = [normalize_column(c) for c in targets.columns]
        loan_officer_col = find_column(loans.columns, ["LOAN OFFICER", "OFFICER", "LOANOFFICER", "LOAN OFFICER NAME"])
        disbursment_date_col = find_column(loans.columns, ["DISBURSEMENT DATE", "DATE", "BOOKING DATE"])
        amount_col = find_column(loans.columns, ["TOTAL LOAN AMOUNT", "LOAN AMOUNT", "AMOUNT"])
        target_officer_col = find_column(targets.columns, ["LOAN OFFICER"])
        target_col = find_column(targets.columns, ["TARGETS", "TARGET"])
        ro_code_col = find_column(targets.columns, ["RO CODE", "CODE"])
        if not all([loan_officer_col, disbursment_date_col, amount_col, target_officer_col, target_col, ro_code_col]):
            raise ProcessingError(_friendly_error())
        loans = loans.rename(columns={loan_officer_col: "LOAN OFFICER", disbursment_date_col: "DISBURSEMENT DATE", amount_col: "TOTAL LOAN AMOUNT"})
        targets = targets.rename(columns={target_officer_col: "LOAN OFFICER", target_col: "TARGETS", ro_code_col: "RO CODE"})
        loans["LOAN OFFICER"] = loans["LOAN OFFICER"].apply(clean_dashboard_text)
        targets["LOAN OFFICER"] = targets["LOAN OFFICER"].apply(clean_dashboard_text)
        loans["TOTAL LOAN AMOUNT"] = loans["TOTAL LOAN AMOUNT"].astype(str).str.replace(",", "", regex=False).str.replace("KES", "", regex=False).str.replace("KSH", "", regex=False).str.strip()
        loans["TOTAL LOAN AMOUNT"] = pd.to_numeric(loans["TOTAL LOAN AMOUNT"], errors="coerce")
        targets["TARGETS"] = targets["TARGETS"].astype(str).str.replace(",", "", regex=False).str.strip()
        targets["TARGETS"] = pd.to_numeric(targets["TARGETS"], errors="coerce")
        loans["DISBURSEMENT DATE"] = pd.to_datetime(loans["DISBURSEMENT DATE"], errors="coerce")
        targets = targets.drop_duplicates(subset="LOAN OFFICER").reset_index(drop=True)
        official_lookup = {normalize_name(officer): officer for officer in targets["LOAN OFFICER"]}
        official_keys = list(official_lookup.keys())
        name_changes = {}
        for loan_name in loans["LOAN OFFICER"].unique():
            normalized = normalize_name(loan_name)
            if normalized in official_lookup:
                name_changes[loan_name] = official_lookup[normalized]
                continue
            match = process.extractOne(normalized, official_keys, scorer=fuzz.token_sort_ratio)
            if match and match[1] >= 80:
                name_changes[loan_name] = official_lookup[match[0]]
        loans["LOAN OFFICER"] = loans["LOAN OFFICER"].replace(name_changes)
        missing = sorted(set(loans["LOAN OFFICER"]) - set(targets["LOAN OFFICER"]))
        print("Unmatched officers:", missing)
        report_date = loans["DISBURSEMENT DATE"].max()
        if pd.isna(report_date):
            raise ProcessingError(_friendly_error())
        report_date = report_date.normalize()
        start_of_month = report_date.replace(day=1)
        report = targets.copy()
        week_start = report_date - pd.Timedelta(days=report_date.weekday())
        working_days = pd.date_range(start=week_start, end=report_date, freq="B")
        day_columns = []
        for day in working_days:
            day_name = day.strftime("%a\n(%d-%b)")
            day_columns.append(day_name)
            daily = loans[loans["DISBURSEMENT DATE"].dt.normalize() == day].groupby("LOAN OFFICER")["TOTAL LOAN AMOUNT"].sum().reset_index()
            daily.rename(columns={"TOTAL LOAN AMOUNT": day_name}, inplace=True)
            report = report.merge(daily, on="LOAN OFFICER", how="left")
            report[day_name] = report[day_name].fillna(0)
        report["WEEK TOTAL"] = report[day_columns].sum(axis=1)
        month_data = loans[(loans["DISBURSEMENT DATE"] >= start_of_month) & (loans["DISBURSEMENT DATE"] <= report_date)]
        mtd = month_data.groupby("LOAN OFFICER")["TOTAL LOAN AMOUNT"].sum().reset_index()
        mtd.rename(columns={"TOTAL LOAN AMOUNT": "MTD"}, inplace=True)
        report = report.merge(mtd, on="LOAN OFFICER", how="left")
        report["MTD"] = report["MTD"].fillna(0)
        report["VARIANCE"] = report["MTD"] - report["TARGETS"]
        report["ACHIEVEMENT %"] = np.where(report["TARGETS"] > 0, report["MTD"] / report["TARGETS"], 0)
        def performance(value):
            if value >= 100:
                return "Excellent"
            elif value >= 80:
                return "Good"
            elif value >= 50:
                return "Average"
            return "Needs Improvement"
        report["STATUS"] = report["ACHIEVEMENT %"].apply(performance)
        report = report.sort_values("ACHIEVEMENT %", ascending=False)
        report.reset_index(drop=True, inplace=True)
        report.insert(0, "RANK", range(1, len(report) + 1))
        totals = {"RANK": "", "RO CODE": "", "LOAN OFFICER": "TOTAL"}
        for col in day_columns:
            totals[col] = report[col].sum()
        totals["WEEK TOTAL"] = report["WEEK TOTAL"].sum()
        totals["MTD"] = report["MTD"].sum()
        totals["TARGETS"] = report["TARGETS"].sum()
        totals["VARIANCE"] = totals["MTD"] - totals["TARGETS"]
        totals["ACHIEVEMENT %"] = 0 if totals["TARGETS"] == 0 else totals["MTD"] / totals["TARGETS"]
        totals["STATUS"] = ""
        report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
        ordered_columns = ["RANK", "RO CODE", "LOAN OFFICER", "TARGETS"]
        ordered_columns.extend(day_columns)
        ordered_columns.extend(["WEEK TOTAL", "MTD", "VARIANCE", "ACHIEVEMENT %", "STATUS"])
        report = report[ordered_columns]
        report.to_excel(output_path, index=False)
        _format_dashboard(output_path, report_date)
    except ProcessingError:
        raise
    except (KeyError, ValueError, TypeError) as error:
        raise ProcessingError(_friendly_error()) from error


def _format_dashboard(output_path: Path, report_date) -> None:
    """Original workbook formatting, kept unchanged from the supplied script."""
    workbook = load_workbook(output_path)
    worksheet = workbook.active
    worksheet.insert_rows(1, 2)
    worksheet["A1"] = "FYNESS NOTES CREDIT LIMITED"
    worksheet["A2"] = f"Week Ending {report_date.strftime('%d %B %Y')}"
    last_column = worksheet.max_column
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    blue = PatternFill(fill_type="solid", start_color="1F4E78")
    worksheet["A1"].fill = blue
    worksheet["A1"].font = Font(bold=True, color="FFFFFF", size=18)
    worksheet["A2"].font = Font(bold=True, italic=True, size=12)
    worksheet["A1"].alignment = Alignment(horizontal="center")
    worksheet["A2"].alignment = Alignment(horizontal="center")
    header_fill = PatternFill(fill_type="solid", start_color="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.freeze_panes = "D5"
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = max_length + 3
    for row in worksheet.iter_rows(min_row=4):
        for cell in row:
            cell.alignment = Alignment(horizontal="left" if cell.column == 3 else "center")
    currency_columns = ["WEEK TOTAL", "MTD", "TARGETS", "VARIANCE"]
    for cell in worksheet[3]:
        if cell.value in currency_columns:
            column_letter = get_column_letter(cell.column)
            for row in range(4, worksheet.max_row + 1):
                worksheet[f"{column_letter}{row}"].number_format = "#,##0.00"
    achievement_column = next(cell.column for cell in worksheet[3] if cell.value == "ACHIEVEMENT %")
    achievement_letter = get_column_letter(achievement_column)
    for row in range(4, worksheet.max_row + 1):
        worksheet[f"{achievement_letter}{row}"].number_format = "0.00%"
    green = PatternFill(fill_type="solid", start_color="92D050")
    yellow = PatternFill(fill_type="solid", start_color="FFD966")
    orange = PatternFill(fill_type="solid", start_color="F4B183")
    red = PatternFill(fill_type="solid", start_color="FF6666")
    for row in range(4, worksheet.max_row):
        cell = worksheet[f"{achievement_letter}{row}"]
        if cell.value is None:
            continue
        if cell.value >= 1:
            cell.fill = green
        elif cell.value >= 0.80:
            cell.fill = yellow
        elif cell.value >= 0.50:
            cell.fill = orange
        else:
            cell.fill = red
    variance_column = next(cell.column for cell in worksheet[3] if cell.value == "VARIANCE")
    variance_letter = get_column_letter(variance_column)
    for row in range(4, worksheet.max_row):
        cell = worksheet[f"{variance_letter}{row}"]
        if cell.value is not None:
            cell.font = Font(color="008000" if cell.value >= 0 else "FF0000", bold=True)
    medals = {4: PatternFill(fill_type="solid", start_color="FFD700"), 5: PatternFill(fill_type="solid", start_color="D9D9D9"), 6: PatternFill(fill_type="solid", start_color="CD7F32")}
    for row, fill in medals.items():
        if row < worksheet.max_row:
            worksheet.cell(row=row, column=1).fill = fill
    total_row = worksheet.max_row
    dark_blue = PatternFill(fill_type="solid", start_color="1F4E78")
    white_font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[total_row]:
        cell.fill = dark_blue
        cell.font = white_font
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = border
    worksheet.auto_filter.ref = None
    worksheet.row_dimensions[3].height = 35
    workbook.save(output_path)


def generate_par_performance(par_path: Path, target_path: Path, output_path: Path) -> None:
    """File-based version of parautomation.py; report logic is unchanged."""
    try:
        par = pd.read_csv(par_path)
        targets = pd.read_excel(target_path)

        def par_normalize_column(col):
            col = str(col).strip().upper().replace("_", " ")
            return " ".join(col.split())

        par.columns = [par_normalize_column(c) for c in par.columns]
        targets.columns = [par_normalize_column(c) for c in targets.columns]

        def par_find_column(columns, possibilities):
            for option in possibilities:
                if option in columns:
                    return option
            return None

        loan_officer_col = par_find_column(par.columns, ["LOANOFFICERNAME", "LOAN OFFICER", "LOAN OFFICER NAME", "OFFICER"])
        principal_col = par_find_column(par.columns, ["OUTSTANDING PRINCIPAL", "PRINCIPAL"])
        par_amount_col = par_find_column(par.columns, ["PORTFOLIOATRISK(PAR)", "PORTFOLIO AT RISK", "PAR"])
        target_officer_col = par_find_column(targets.columns, ["LOAN OFFICER"])
        ro_code_col = par_find_column(targets.columns, ["RO CODE", "CODE"])
        if not all([loan_officer_col, principal_col, par_amount_col, target_officer_col, ro_code_col]):
            raise ProcessingError(_friendly_error())

        par = par.rename(columns={loan_officer_col: "LOAN OFFICER", principal_col: "OUTSTANDING PRINCIPAL", par_amount_col: "PORTFOLIO AT RISK"})
        targets = targets.rename(columns={target_officer_col: "LOAN OFFICER", ro_code_col: "RO CODE"})

        def par_clean_text(text):
            if pd.isna(text):
                return ""
            text = str(text).upper().strip()
            text = re.sub(r"\s+", " ", text)
            return re.sub(r"[^\w\s]", "", text)

        def par_clean_money(series):
            return pd.to_numeric(
                series.astype(str).str.replace(",", "", regex=False).str.replace("KES", "", regex=False).str.replace("KSH", "", regex=False).str.strip(),
                errors="coerce",
            )

        par["LOAN OFFICER"] = par["LOAN OFFICER"].apply(par_clean_text)
        targets["LOAN OFFICER"] = targets["LOAN OFFICER"].apply(par_clean_text)
        par["OUTSTANDING PRINCIPAL"] = par_clean_money(par["OUTSTANDING PRINCIPAL"])
        par["PORTFOLIO AT RISK"] = par_clean_money(par["PORTFOLIO AT RISK"])
        par = par.dropna(how="all")
        par = par[par["LOAN OFFICER"] != ""]
        par = par[~par["LOAN OFFICER"].str.upper().str.contains("TOTAL", na=False)]
        targets = targets.drop_duplicates(subset="LOAN OFFICER").reset_index(drop=True)

        def par_normalize_name(name):
            words = par_clean_text(name).split()
            cleaned = []
            for word in words:
                if len(word) == 1:
                    continue
                if word in ["MR", "MRS", "MISS", "MS", "DR", "REV", "PASTOR"]:
                    continue
                cleaned.append(word)
            return " ".join(sorted(cleaned))

        official_lookup = {par_normalize_name(officer): officer for officer in targets["LOAN OFFICER"]}
        official_keys = list(official_lookup.keys())
        name_changes = {}
        for officer in par["LOAN OFFICER"].unique():
            normalized = par_normalize_name(officer)
            if normalized in official_lookup:
                name_changes[officer] = official_lookup[normalized]
                continue
            match = process.extractOne(normalized, official_keys, scorer=fuzz.token_sort_ratio)
            if match and match[1] >= 80:
                name_changes[officer] = official_lookup[match[0]]
        par["LOAN OFFICER"] = par["LOAN OFFICER"].replace(name_changes)
        missing = sorted(set(par["LOAN OFFICER"]) - set(targets["LOAN OFFICER"]))
        print("Unmatched officers:", missing)

        company_principal = par["OUTSTANDING PRINCIPAL"].sum()
        company_par = par["PORTFOLIO AT RISK"].sum()
        company_percentage = company_par / company_principal * 100 if company_principal > 0 else 0
        active_par = par[par["LOAN OFFICER"].isin(targets["LOAN OFFICER"])].copy()
        summary = active_par.groupby("LOAN OFFICER", as_index=False).agg({"OUTSTANDING PRINCIPAL": "sum", "PORTFOLIO AT RISK": "sum"})
        report = targets.merge(summary, on="LOAN OFFICER", how="left")
        report["HAS PORTFOLIO"] = report["OUTSTANDING PRINCIPAL"].notna()
        report["OUTSTANDING PRINCIPAL"] = report["OUTSTANDING PRINCIPAL"].fillna(0)
        report["PORTFOLIO AT RISK"] = report["PORTFOLIO AT RISK"].fillna(0)
        report["PAR %"] = np.where(report["OUTSTANDING PRINCIPAL"] > 0, (report["PORTFOLIO AT RISK"] / report["OUTSTANDING PRINCIPAL"]) * 100, 0)
        report = report.sort_values(by=["HAS PORTFOLIO", "PAR %"], ascending=[False, True])
        report.reset_index(drop=True, inplace=True)
        report.insert(0, "RANK", range(1, len(report) + 1))
        for col in ["OUTSTANDING PRINCIPAL", "PORTFOLIO AT RISK", "PAR %"]:
            report.loc[report["HAS PORTFOLIO"] == False, col] = "-"
        report = report[["RANK", "RO CODE", "LOAN OFFICER", "OUTSTANDING PRINCIPAL", "PORTFOLIO AT RISK", "PAR %"]]
        company_total = pd.DataFrame([{"RANK": "", "RO CODE": "", "LOAN OFFICER": "COMPANY TOTAL", "OUTSTANDING PRINCIPAL": company_principal, "PORTFOLIO AT RISK": company_par, "PAR %": company_percentage}])
        report = pd.concat([report, company_total], ignore_index=True)
        report.to_excel(output_path, index=False)
        _format_par_report(output_path)
    except ProcessingError:
        raise
    except (KeyError, ValueError, TypeError) as error:
        raise ProcessingError(_friendly_error()) from error


def _format_par_report(output_path: Path) -> None:
    """Original PAR workbook formatting, excluding only Colab download."""
    workbook = load_workbook(output_path)
    worksheet = workbook.active
    worksheet.insert_rows(1, 2)
    worksheet["A1"] = "FYNESS NOTES CREDIT LIMITED"
    worksheet["A2"] = "PAR PERFORMANCE REPORT"
    last_column = worksheet.max_column
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    blue = PatternFill(fill_type="solid", start_color="1F4E78")
    worksheet["A1"].fill = blue
    worksheet["A1"].font = Font(bold=True, color="FFFFFF", size=18)
    worksheet["A2"].font = Font(bold=True, italic=True, size=12)
    worksheet["A1"].alignment = Alignment(horizontal="center")
    worksheet["A2"].alignment = Alignment(horizontal="center")
    header_fill = PatternFill(fill_type="solid", start_color="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    par_column = next(cell.column for cell in worksheet[3] if cell.value == "PAR %")
    par_letter = get_column_letter(par_column)
    worksheet.conditional_formatting.add(f"{par_letter}4:{par_letter}{worksheet.max_row-1}", ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"))
    for cell in worksheet[3]:
        if cell.value in ["OUTSTANDING PRINCIPAL", "PORTFOLIO AT RISK"]:
            letter = get_column_letter(cell.column)
            for row in range(4, worksheet.max_row + 1):
                worksheet[f"{letter}{row}"].number_format = "#,##0.00"
        if cell.value == "PAR %":
            letter = get_column_letter(cell.column)
            for row in range(4, worksheet.max_row + 1):
                worksheet[f"{letter}{row}"].number_format = "0.00"
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = border
    for column in worksheet.columns:
        max_length = 0
        letter = get_column_letter(column[0].column)
        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[letter].width = max_length + 3
    worksheet.freeze_panes = "D4"
    total_row = worksheet.max_row
    dark_blue = PatternFill(fill_type="solid", start_color="1F4E78")
    white_font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[total_row]:
        cell.fill = dark_blue
        cell.font = white_font
    workbook.save(output_path)
