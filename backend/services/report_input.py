"""Detect an optional report title before any report-specific transformation."""

import csv
import re
from pathlib import Path

import pandas as pd


def _label(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


HEADER_LABELS = {
    "accountnumber", "clientname", "branch", "loanofficer", "loanofficername",
    "officer", "officername", "totalinarrears", "totalarrears", "arrears",
    "amountinarrears", "arrearsamount", "repaymentdate", "transactionnumber",
    "totalloanamount", "loanamount", "portfolioatriskpar", "disbursementdate",
}


def _has_title(rows):
    if len(rows) < 2:
        return False
    first = [str(value).strip() for value in rows[0] if pd.notna(value) and str(value).strip()]
    second = {_label(value) for value in rows[1] if pd.notna(value)}
    # Require both a single title cell and recognizable labels on the next row.
    # Never discard an existing header, an ambiguous row, or a data row blindly.
    return (
        len(first) == 1
        and _label(first[0]) not in HEADER_LABELS
        and len(second & HEADER_LABELS) >= 2
    )


def read_report(path: Path, **kwargs):
    """Read CSV/Excel with or without one title row; targets bypass this reader."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as source:
            reader = csv.reader(source)
            rows = [row for _, row in zip(range(2), reader)]
        return pd.read_csv(path, skiprows=int(_has_title(rows)), **kwargs)
    preview = pd.read_excel(path, header=None, nrows=2)
    return pd.read_excel(path, skiprows=int(_has_title(preview.values.tolist())), **kwargs)
