from pathlib import Path

# ===== Excel DB =====
DATA_FILE_PATH = Path(r"C:\Users\avraham.shrem\Desktop\Wheelchair_App\Wheelchair_Data.xlsx")

SHEET_TX = "WC_Transactions"
SHEET_WC = "Wheelchairs"
SHEET_LOC = "Locations"
SHEET_LISTS = "Lists"
SHEET_CURRENT = "WC_Current"

# ===== Mail (SMTP) =====
SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = ""          # e.g. avraham.shrem@moh.gov.il
SMTP_PASSWORD = ""      # app password / org credential if allowed
SMTP_FROM = SMTP_USER
SMTP_TO_DEFAULT = "weelchairs.shmuelharofeh@gmail.com"
USE_STARTTLS = True
