import base64
import hashlib
import hmac
import os
import secrets
import time
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, EmailStr
from backend.db import Base, engine, session_scope
from backend.models.user import Company, User

Base.metadata.create_all(bind=engine)
router = APIRouter(prefix="/auth", tags=["auth"])
SECRET = os.getenv("AUTH_SECRET", "local-development-secret-change-me")

class RegisterRequest(BaseModel):
    company_name: str
    full_name: str
    email: EmailStr
    password: str

class LoginRequest(BaseModel):
    email: EmailStr
    password: str

def _hash(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 150_000)
    return f"{base64.urlsafe_b64encode(salt).decode()}${base64.urlsafe_b64encode(digest).decode()}"

def _verify(password: str, encoded: str) -> bool:
    salt, expected = encoded.split("$", 1)
    actual = _hash(password, base64.urlsafe_b64decode(salt)).split("$", 1)[1]
    return hmac.compare_digest(actual, expected)

def _token(user: User) -> str:
    payload = f"{user.id}:{int(time.time()) + 86400}"
    signature = hmac.new(SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}:{signature}".encode()).decode()

@router.post("/register")
def register(request: RegisterRequest):
    if len(request.password) < 8:
        raise HTTPException(400, "Password must be at least 8 characters.")
    with session_scope() as session:
        if session.query(User).filter_by(email=request.email.lower()).first():
            raise HTTPException(409, "An account with this email already exists.")
        company = Company(name=request.company_name.strip())
        session.add(company); session.flush()
        user = User(company_id=company.id, full_name=request.full_name.strip(), email=request.email.lower(), password_hash=_hash(request.password))
        session.add(user); session.flush()
        return {"token": _token(user), "user": {"name": user.full_name, "email": user.email, "company": company.name}}

@router.post("/login")
def login(request: LoginRequest):
    with session_scope() as session:
        user = session.query(User).filter_by(email=request.email.lower()).first()
        if not user or not _verify(request.password, user.password_hash):
            raise HTTPException(401, "Email or password is incorrect.")
        company = session.get(Company, user.company_id)
        return {"token": _token(user), "user": {"name": user.full_name, "email": user.email, "company": company.name if company else ""}}
