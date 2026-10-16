import os
from contextlib import contextmanager
from threading import Lock
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./panta.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)
engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {"connect_timeout": 5},
    pool_pre_ping=True,
)
SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)
_schema_ready = False
_schema_lock = Lock()

class Base(DeclarativeBase):
    pass


def initialize_database():
    """Initialize on account access, without coupling report startup to Postgres.

    Failed attempts are not cached, so a restored database works on the next
    request without restarting the API.
    """
    global _schema_ready
    with _schema_lock:
        if not _schema_ready:
            Base.metadata.create_all(bind=engine)
            _schema_ready = True

@contextmanager
def session_scope():
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
