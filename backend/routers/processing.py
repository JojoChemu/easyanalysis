from pathlib import Path
from tempfile import TemporaryDirectory

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import Response

from backend.services.processing import (
    ProcessingError,
    clean_all_loans,
    clean_loan_arrears,
    clean_portfolio_risk,
    clean_repayment_report,
    generate_ro_dashboard,
    generate_ro_monthly_customer_mix,
    generate_par_performance,
)
from backend.utils.files import save_upload, validate_extension

router = APIRouter(tags=["processing"])


def _invalid_report(error: ProcessingError) -> HTTPException:
    return HTTPException(status_code=422, detail=str(error))


def _download(path: Path, media_type: str) -> Response:
    return Response(
        content=path.read_bytes(),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{path.name}"'},
    )


@router.post("/clean-all-loans")
async def clean_all_loans_endpoint(file: UploadFile = File(...)):
    validate_extension(file, {".csv"})
    with TemporaryDirectory(prefix="panta-") as directory:
        input_path = await save_upload(file, Path(directory))
        output_path = Path(directory) / "AllLoans_Clean.csv"
        try:
            clean_all_loans(input_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "text/csv")


@router.post("/clean-portfolio-risk")
async def clean_portfolio_risk_endpoint(file: UploadFile = File(...)):
    validate_extension(file, {".csv"})
    with TemporaryDirectory(prefix="panta-") as directory:
        input_path = await save_upload(file, Path(directory))
        output_path = Path(directory) / "PortfolioRiskReport_Clean.csv"
        try:
            clean_portfolio_risk(input_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "text/csv")


@router.post("/clean-loan-arrears")
async def clean_loan_arrears_endpoint(file: UploadFile = File(...)):
    validate_extension(file, {".csv"})
    with TemporaryDirectory(prefix="panta-") as directory:
        input_path = await save_upload(file, Path(directory))
        output_path = Path(directory) / "LoanArrears_Clean.csv"
        try:
            clean_loan_arrears(input_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "text/csv")


@router.post("/clean-repayment-report")
async def clean_repayment_report_endpoint(file: UploadFile = File(...)):
    validate_extension(file, {".csv", ".xlsx", ".xls"})
    with TemporaryDirectory(prefix="easyanalysis-") as directory:
        input_path = await save_upload(file, Path(directory))
        output_path = Path(directory) / "Cleaned_LoanRepaymentReport.xlsx"
        try:
            clean_repayment_report(input_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@router.post("/ro-performance-dashboard")
async def ro_performance_dashboard_endpoint(
    loan_file: UploadFile = File(...), targets_file: UploadFile = File(...)
):
    validate_extension(loan_file, {".csv"})
    validate_extension(targets_file, {".xlsx", ".xls"})
    with TemporaryDirectory(prefix="panta-") as directory:
        loan_path = await save_upload(loan_file, Path(directory))
        targets_path = await save_upload(targets_file, Path(directory))
        output_path = Path(directory) / "RO_Performance_Dashboard.xlsx"
        try:
            generate_ro_dashboard(loan_path, targets_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@router.post("/ro-monthly-customer-mix")
async def ro_monthly_customer_mix_endpoint(
    loan_file: UploadFile = File(...), targets_file: UploadFile = File(...)
):
    validate_extension(loan_file, {".csv"})
    validate_extension(targets_file, {".xlsx", ".xls"})
    with TemporaryDirectory(prefix="easyanalysis-") as directory:
        loan_path = await save_upload(loan_file, Path(directory))
        targets_path = await save_upload(targets_file, Path(directory))
        output_path = Path(directory) / "RO_Monthly_Performance_Customer_Mix.xlsx"
        try:
            generate_ro_monthly_customer_mix(loan_path, targets_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@router.post("/par-performance")
async def par_performance_endpoint(
    par_file: UploadFile = File(...), targets_file: UploadFile = File(...)
):
    validate_extension(par_file, {".csv"})
    validate_extension(targets_file, {".xlsx", ".xls"})
    with TemporaryDirectory(prefix="easyanalysis-") as directory:
        par_path = await save_upload(par_file, Path(directory))
        targets_path = await save_upload(targets_file, Path(directory))
        output_path = Path(directory) / "PAR_Performance_Report.xlsx"
        try:
            generate_par_performance(par_path, targets_path, output_path)
        except ProcessingError as error:
            raise _invalid_report(error) from error
        return _download(output_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
