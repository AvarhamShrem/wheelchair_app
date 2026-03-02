# Wheelchair Desktop App (Tkinter)

Python desktop application for wheelchair transactions, backed by **Excel only** (`Wheelchair_Data.xlsx`).

## Requirements

- Windows 10/11
- Python 3.10+ (no admin rights required)
- Excel workbook available at project root by default: `Wheelchair_Data.xlsx`

## Setup on Windows

Open **PowerShell** in the project folder and run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Configuration

The app reads config from environment variables (optional):

- `WHEELCHAIR_DATA_FILE` (default: `Wheelchair_Data.xlsx`)
- `SMTP_HOST` (default: `smtp.office365.com`)
- `SMTP_PORT` (default: `587`)
- `SMTP_USER`
- `SMTP_PASSWORD`
- `SMTP_FROM` (defaults to `SMTP_USER`)
- `SMTP_TO_DEFAULT`
- `USE_STARTTLS` (`true`/`false`, default `true`)

Example (PowerShell):

```powershell
$env:WHEELCHAIR_DATA_FILE = "C:\path\to\Wheelchair_Data.xlsx"
$env:SMTP_USER = "user@example.com"
$env:SMTP_PASSWORD = "<secret>"
$env:SMTP_TO_DEFAULT = "reports@example.com"
```

## Run

```powershell
python main.py
```

## Notes

- `WC_Current` is treated as read-only and never written to.
- New transactions are appended to `WC_Transactions` with `TxnID` and `SortKey` formula.
- New transaction email sends JSON payload as an attached `.json` file.
- Weekly report range is last 7 days ending yesterday.
