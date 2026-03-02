from __future__ import annotations

import smtplib
from datetime import date, timedelta
from email.message import EmailMessage
from typing import Dict

import excel_db
from config import SMTP_FROM, SMTP_HOST, SMTP_PASSWORD, SMTP_PORT, SMTP_USER, USE_STARTTLS


def _send_message(msg: EmailMessage) -> None:
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        if USE_STARTTLS:
            smtp.starttls()
        if SMTP_USER:
            smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)


def send_mail_with_attachment_json(
    to_list: str,
    subject: str,
    json_bytes: bytes,
    attachment_name: str,
    body: str = "",
) -> None:
    msg = EmailMessage()
    msg["From"] = SMTP_FROM or SMTP_USER
    msg["To"] = to_list
    msg["Subject"] = subject
    msg.set_content(body or "")
    msg.add_attachment(json_bytes, maintype="application", subtype="json", filename=attachment_name)
    _send_message(msg)


def send_weekly_report_email(to_list: str) -> None:
    end_date = date.today() - timedelta(days=1)
    start_date = end_date - timedelta(days=6)

    transfer_count = excel_db.count_txn_type_in_range("העברה", start_date, end_date)
    status_counts: Dict[str, int] = excel_db.status_aggregation_from_current()

    lines = [
        "Weekly Wheelchair Summary",
        f"Range: {start_date.isoformat()} to {end_date.isoformat()}",
        "",
        f'Transactions with TxnType="העברה": {transfer_count}',
        "",
        "Current status aggregation from WC_Current:",
    ]

    if status_counts:
        for status, count in sorted(status_counts.items()):
            lines.append(f"- {status}: {count}")
    else:
        lines.append("- No status data found")

    msg = EmailMessage()
    msg["From"] = SMTP_FROM or SMTP_USER
    msg["To"] = to_list
    msg["Subject"] = f"Weekly Wheelchair Report ({start_date.isoformat()} - {end_date.isoformat()})"
    msg.set_content("\n".join(lines))
    _send_message(msg)
