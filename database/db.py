from pathlib import Path

from sqlalchemy import create_engine, event
from sqlalchemy.orm import DeclarativeBase


# Project root directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Database storage location
DB_DIR = BASE_DIR / "data" / "processed"
DB_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DB_DIR / "decisiongraph.db"

# SQLite connection
DATABASE_URL = f"sqlite:///{DB_PATH.as_posix()}"

engine = create_engine(
    DATABASE_URL,
    echo=False
)


# Enable foreign key enforcement in SQLite
@event.listens_for(engine, "connect")
def enable_foreign_keys(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()


# Base class for all database models
class Base(DeclarativeBase):
    pass