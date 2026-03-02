from __future__ import annotations

import json
from datetime import datetime
from typing import Tuple

import excel_db
from models import Transaction


def validate_timestamp(value: datetime) -> None:
    if not isinstance(value, datetime):
        raise ValueError("Invalid timestamp")


def validate_and_save(tx: Transaction) -> Tuple[int, bytes]:
    validate_timestamp(tx.timestamp)

    if not tx.created_by.strip():
        raise ValueError("חסר: CreatedBy")
    if not tx.wheelchair_id.strip():
        raise ValueError("חסר: WheelchairID")
    if not tx.txn_type.strip():
        raise ValueError("חסר: TxnType")
    if not tx.to_location_id:
        raise ValueError("חסר: ToLocation")
    if tx.from_location_id == tx.to_location_id:
        raise ValueError("ToLocation must be different from FromLocation")
    if not tx.status_after.strip():
        raise ValueError("חסר: StatusAfter")
    if not tx.condition_after.strip():
        raise ValueError("חסר: ConditionAfter")

    txn_id = excel_db.append_transaction(tx)
    payload = json.dumps(tx.to_json_dict(), ensure_ascii=False, indent=2).encode("utf-8")
    return txn_id, payload
