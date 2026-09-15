import os
from urllib.parse import urlparse, unquote

from dotenv import load_dotenv

load_dotenv()

# Database configuration. For Railway MySQL, you can either set the five
# MYSQL* variables or paste MYSQL_PUBLIC_URL into the Vercel environment.
_db_url = os.environ.get("MYSQL_PUBLIC_URL") or os.environ.get("MYSQL_URL") or os.environ.get("DATABASE_URL")
_url_parts = urlparse(_db_url) if _db_url else None

DB_HOST = (
    os.environ.get("E_PROGRESS_CARD_DB_HOST")
    or os.environ.get("DB_HOST")
    or os.environ.get("MYSQLHOST")
    or (_url_parts.hostname if _url_parts else None)
    or "localhost"
)
DB_USER = (
    os.environ.get("E_PROGRESS_CARD_DB_USER")
    or os.environ.get("DB_USER")
    or os.environ.get("MYSQLUSER")
    or (unquote(_url_parts.username) if _url_parts and _url_parts.username else None)
    or "root"
)
DB_PASSWORD = (
    os.environ.get("E_PROGRESS_CARD_DB_PASSWORD")
    or os.environ.get("DB_PASSWORD")
    or os.environ.get("MYSQLPASSWORD")
    or (unquote(_url_parts.password) if _url_parts and _url_parts.password else None)
    or ""
)
DB_NAME = (
    os.environ.get("E_PROGRESS_CARD_DB_NAME")
    or os.environ.get("DB_NAME")
    or os.environ.get("MYSQLDATABASE")
    or ((_url_parts.path or "").lstrip("/") if _url_parts else None)
    or "e_progress_card"
)
DB_PORT = int(
    os.environ.get("E_PROGRESS_CARD_DB_PORT")
    or os.environ.get("DB_PORT")
    or os.environ.get("MYSQLPORT")
    or (_url_parts.port if _url_parts and _url_parts.port else 3306)
)

SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret-key-in-production")
