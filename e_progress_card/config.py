# =========================================================
# E-PROGRESS CARD - DATABASE CONFIGURATION
# =========================================================

import os

# Prefer explicit project-specific environment variables first, then the legacy
# local DB_* variables, then Railway's standard MYSQL* variables, and finally
# local development defaults. This keeps both Railway and local XAMPP setups
# working without hard-coding database credentials.
DB_HOST = (
    os.environ.get("E_PROGRESS_CARD_DB_HOST")
    or os.environ.get("DB_HOST")
    or os.environ.get("MYSQLHOST")
    or "localhost"
)
DB_USER = (
    os.environ.get("E_PROGRESS_CARD_DB_USER")
    or os.environ.get("DB_USER")
    or os.environ.get("MYSQLUSER")
    or "e_progress_app"
)
DB_PASSWORD = (
    os.environ.get("E_PROGRESS_CARD_DB_PASSWORD")
    or os.environ.get("DB_PASSWORD")
    or os.environ.get("MYSQLPASSWORD")
    or "EProgress_Local_2026_App"
)
DB_NAME = (
    os.environ.get("E_PROGRESS_CARD_DB_NAME")
    or os.environ.get("DB_NAME")
    or os.environ.get("MYSQLDATABASE")
    or "e_progress_card"
)
DB_PORT = int(
    os.environ.get("E_PROGRESS_CARD_DB_PORT")
    or os.environ.get("DB_PORT")
    or os.environ.get("MYSQLPORT")
    or "3307"
)

SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret-key-in-production")