from pathlib import Path
import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from backend.routers.processing import router as processing_router
from backend.routers.auth import router as auth_router


app = FastAPI(title="EasyAnalysis by Panta Development and AI", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://localhost:4173",
        os.getenv("FRONTEND_URL", "https://your-site.netlify.app"),
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(processing_router, prefix="/api")
app.include_router(auth_router, prefix="/api")

health_dir = Path(__file__).parent / "static"
health_dir.mkdir(exist_ok=True)


@app.get("/health")
def health_check() -> dict[str, str]:
    return {"status": "ok", "service": "easyanalysis-by-panta"}

