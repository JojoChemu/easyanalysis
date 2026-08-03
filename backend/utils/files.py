from pathlib import Path
from uuid import uuid4

from fastapi import HTTPException, UploadFile


async def save_upload(upload: UploadFile, directory: Path) -> Path:
    safe_name = Path(upload.filename or "upload").name
    path = directory / f"{uuid4().hex}-{safe_name}"
    path.write_bytes(await upload.read())
    return path


def validate_extension(upload: UploadFile, allowed: set[str]) -> None:
    extension = Path(upload.filename or "").suffix.lower()
    if extension not in allowed:
        expected = ", ".join(sorted(allowed))
        raise HTTPException(status_code=415, detail=f"Expected file type: {expected}")
