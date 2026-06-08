"""Track already-processed bills so re-runs don't create duplicates."""

from __future__ import annotations

import json
from pathlib import Path

from .models import BillRecord


def _key(rec: BillRecord) -> str:
    """A bill is unique by supplier + account + invoice (or total+date)."""
    invoice = rec.invoice_no or f"{rec.total_amount}-{rec.bill_date}"
    return f"{rec.supplier_code}|{rec.account_number}|{invoice}"


def load_state(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        return set(json.loads(path.read_text(encoding="utf-8")).get("processed", []))
    except (json.JSONDecodeError, OSError):
        return set()


def save_state(path: Path, processed: set[str]) -> None:
    path.write_text(json.dumps({"processed": sorted(processed)}, indent=2), encoding="utf-8")


def filter_new(records: list[BillRecord], processed: set[str]) -> tuple[list[BillRecord], set[str]]:
    """Return (new_records, updated_processed_set). Unknown billers always pass."""
    new: list[BillRecord] = []
    updated = set(processed)
    for rec in records:
        if rec.biller_name == "UNKNOWN":
            new.append(rec)            # let user see/fix unrecognised bills
            continue
        k = _key(rec)
        if k in updated:
            continue
        new.append(rec)
        updated.add(k)
    return new, updated
