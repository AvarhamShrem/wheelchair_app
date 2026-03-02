from __future__ import annotations
import smtplib
from email.message import EmailMessage
from typing import Optional

from config import SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM, USE_STARTTLS


def send_mail_with_attachment_json(
    to_list: str,
    subject: str,
    json_bytes: bytes,
    attachment_name: str,
    body: str = ""
) -> None:
    msg = EmailMessage()
    msg["From"] = SMTP_FROM or SMTP_USER
    msg["To"] = to_list
    msg["Subject"] = subject
    msg.set_content(body or "")

    msg.add_attachment(json_bytes, maintype="application", subtype="json", filename=attachment_name)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        if USE_STARTTLS:
            s.starttls()
        if SMTP_USER:
            s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)
