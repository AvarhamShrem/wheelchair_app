from dataclasses import dataclass
from datetime import datetime
from typing import Optional

@dataclass(frozen=True)
class Transaction:
    timestamp: datetime
    created_by: str
    wheelchair_id: str
    txn_type: str
    from_location_id: int
    to_location_id: int
    status_after: str
    condition_after: str
    patient_name: str = ""
    notes: str = ""

    def to_json_dict(self) -> dict:
        return {
            "WheelchairID": int(self.wheelchair_id),
            "Timestamp": self.timestamp.strftime("%Y-%m-%dT%H:%M:%S"),
            "CreatedBy": self.created_by,
            "TxnType": self.txn_type,
            "FromLocationID": int(self.from_location_id),
            "ToLocationID": int(self.to_location_id),
            "StatusAfter": self.status_after,
            "ConditionAfter": self.condition_after,
            "PatientName": self.patient_name,
            "Notes": self.notes,
        }
