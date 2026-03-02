from __future__ import annotations
import json
from datetime import datetime
from typing import Tuple

from db.models import Transaction
from db import excel_db


def validate_and_save(tx: Transaction) -> Tuple[int, bytes]:
    if not tx.created_by.strip():
        raise ValueError("חסר: עודכן ע״י")
    if not tx.wheelchair_id.strip():
        raise ValueError("חסר: כיסא גלגלים")
    if not tx.txn_type.strip():
        raise ValueError("חסר: פעולה")
    if not tx.to_location_id:
        raise ValueError("חסר: מיקום יעד")
    if tx.from_location_id == tx.to_location_id:
        raise ValueError("מיקום היעד חייב להיות שונה מהמיקום הנוכחי")
    if not tx.status_after.strip():
        raise ValueError("חסר: סטטוס חדש")
    if not tx.condition_after.strip():
        raise ValueError("חסר: מצב תקינות")

    txn_id = excel_db.append_transaction(tx)
    payload = json.dumps(tx.to_json_dict(), ensure_ascii=False).encode("utf-8")
    return txn_id, payload
