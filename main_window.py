from __future__ import annotations

from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

import excel_db
from config import DATA_FILE_PATH, SMTP_TO_DEFAULT
from mail_service import send_mail_with_attachment_json, send_weekly_report_email
from models import Transaction
from transaction_service import validate_and_save


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Wheelchair App")
        self.geometry("900x520")

        self.locations = []
        self.location_display_to_id: dict[str, int] = {}

        self._build_ui()
        self._load_reference_data()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        self.timestamp_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.wheelchair_id_var = tk.StringVar()
        self.txn_type_var = tk.StringVar()
        self.from_location_var = tk.StringVar()
        self.to_location_var = tk.StringVar()
        self.status_after_var = tk.StringVar()
        self.condition_after_var = tk.StringVar()
        self.created_by_var = tk.StringVar()
        self.patient_name_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.email_to_var = tk.StringVar(value=SMTP_TO_DEFAULT)

        rows = [
            ("Timestamp", ttk.Entry(frame, textvariable=self.timestamp_var, state="readonly")),
            ("WheelchairID", ttk.Combobox(frame, textvariable=self.wheelchair_id_var, state="readonly")),
            ("TxnType", ttk.Combobox(frame, textvariable=self.txn_type_var, state="readonly")),
            ("FromLocation", ttk.Entry(frame, textvariable=self.from_location_var, state="readonly")),
            ("ToLocation", ttk.Combobox(frame, textvariable=self.to_location_var, state="readonly")),
            ("StatusAfter", ttk.Combobox(frame, textvariable=self.status_after_var, state="readonly")),
            ("ConditionAfter", ttk.Combobox(frame, textvariable=self.condition_after_var, state="readonly")),
            ("CreatedBy", ttk.Entry(frame, textvariable=self.created_by_var)),
            ("PatientName", ttk.Entry(frame, textvariable=self.patient_name_var)),
            ("Notes", ttk.Entry(frame, textvariable=self.notes_var)),
            ("EmailTo", ttk.Entry(frame, textvariable=self.email_to_var)),
        ]

        for idx, (label, widget) in enumerate(rows):
            ttk.Label(frame, text=label).grid(row=idx, column=0, padx=6, pady=5, sticky="e")
            widget.grid(row=idx, column=1, padx=6, pady=5, sticky="ew")

        frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="Save New Transaction", command=self.save_transaction).grid(row=12, column=0, pady=14)
        ttk.Button(frame, text="Send Weekly Report", command=self.send_weekly_report).grid(row=12, column=1, pady=14, sticky="w")

        self.wheelchair_combo = rows[1][1]
        self.txn_type_combo = rows[2][1]
        self.to_location_combo = rows[4][1]
        self.status_after_combo = rows[5][1]
        self.condition_after_combo = rows[6][1]

        self.wheelchair_combo.bind("<<ComboboxSelected>>", self._on_wheelchair_changed)

    def _load_reference_data(self) -> None:
        try:
            wheelchair_ids = excel_db.list_active_wheelchair_ids()
            locations = excel_db.list_locations()
            enums = excel_db.list_enums()
        except FileNotFoundError:
            messagebox.showerror("Missing Excel file", f"Excel file not found: {DATA_FILE_PATH}")
            return
        except Exception as exc:
            messagebox.showerror("Load error", f"Failed loading dropdown data: {exc}")
            return

        self.locations = locations
        self.location_display_to_id = {
            f'{item["LocationID"]} - {item["LocationName"]}': item["LocationID"] for item in locations
        }

        self.wheelchair_combo["values"] = wheelchair_ids
        self.txn_type_combo["values"] = enums.get("TxnType", [])
        self.to_location_combo["values"] = list(self.location_display_to_id.keys())
        self.status_after_combo["values"] = enums.get("StatusAfter", [])
        self.condition_after_combo["values"] = enums.get("ConditionAfter", [])

    def _on_wheelchair_changed(self, _event: object) -> None:
        wheelchair_id = self.wheelchair_id_var.get().strip()
        if not wheelchair_id:
            self.from_location_var.set("")
            return

        try:
            location_id = excel_db.get_current_location_id(wheelchair_id)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed reading current location: {exc}")
            return

        if location_id is None:
            self.from_location_var.set("")
            return

        name = next((i["LocationName"] for i in self.locations if i["LocationID"] == location_id), "")
        self.from_location_var.set(f"{location_id} - {name}".strip())

    def save_transaction(self) -> None:
        try:
            timestamp = datetime.strptime(self.timestamp_var.get().strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            messagebox.showerror("Validation", "Invalid timestamp")
            return

        from_location_text = self.from_location_var.get().strip()
        if not from_location_text:
            messagebox.showerror("Validation", "Missing FromLocation")
            return

        to_location_id = self.location_display_to_id.get(self.to_location_var.get().strip())
        if to_location_id is None:
            messagebox.showerror("Validation", "Missing ToLocation")
            return

        try:
            from_location_id = int(from_location_text.split("-")[0].strip())
        except Exception:
            messagebox.showerror("Validation", "Invalid FromLocation")
            return

        transaction = Transaction(
            timestamp=timestamp,
            wheelchair_id=self.wheelchair_id_var.get().strip(),
            txn_type=self.txn_type_var.get().strip(),
            from_location_id=from_location_id,
            to_location_id=to_location_id,
            status_after=self.status_after_var.get().strip(),
            condition_after=self.condition_after_var.get().strip(),
            created_by=self.created_by_var.get().strip(),
            patient_name=self.patient_name_var.get().strip(),
            notes=self.notes_var.get().strip(),
        )

        try:
            txn_id, payload = validate_and_save(transaction)
            to_email = self.email_to_var.get().strip()
            if to_email:
                send_mail_with_attachment_json(
                    to_list=to_email,
                    subject=f"Wheelchair transaction #{txn_id}",
                    json_bytes=payload,
                    attachment_name=f"transaction_{txn_id}.json",
                    body="Attached is the new transaction JSON payload.",
                )
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))
            return

        messagebox.showinfo("Success", f"Transaction #{txn_id} saved")
        self.timestamp_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def send_weekly_report(self) -> None:
        to_email = self.email_to_var.get().strip()
        if not to_email:
            messagebox.showerror("Validation", "EmailTo is required for weekly report")
            return

        try:
            send_weekly_report_email(to_email)
        except Exception as exc:
            messagebox.showerror("Mail failed", str(exc))
            return
        messagebox.showinfo("Success", "Weekly report email sent")


def run_app() -> None:
    app = App()
    app.mainloop()
