from __future__ import annotations

import os
from pathlib import Path

# ===== Excel DB =====
# Allow override via environment variable for easy Windows setup.
DATA_FILE_PATH = Path(os.getenv("WHEELCHAIR_DATA_FILE", "Wheelchair_Data.xlsx"))

SHEET_TX = "WC_Transactions"
SHEET_WC = "Wheelchairs"
SHEET_LOC = "Locations"
SHEET_LISTS = "Lists"
SHEET_CURRENT = "WC_Current"

# ===== Mail (SMTP) =====
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_FROM = os.getenv("SMTP_FROM", SMTP_USER)
SMTP_TO_DEFAULT = os.getenv("SMTP_TO_DEFAULT", "")
USE_STARTTLS = os.getenv("USE_STARTTLS", "true").lower() in {"1", "true", "yes", "y"}
