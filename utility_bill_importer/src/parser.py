"""Turn raw PDF text into a normalised BillRecord using the biller registry."""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from .models import BillRecord

# Bahasa Malaysia + English month names mapped to month numbers.
_MONTHS = {
    "jan": 1, "januari": 1, "january": 1,
    "feb": 2, "februari": 2, "february": 2,
    "mac": 3, "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "mei": 5, "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "julai": 7, "july": 7,
    "ogos": 8, "aug": 8, "august": 8, "ogs": 8,
    "sep": 9, "sept": 9, "september": 9,
    "okt": 10, "oct": 10, "oktober": 10, "october": 10,
    "nov": 11, "november": 11,
    "dis": 12, "dec": 12, "disember": 12, "december": 12,
}


def identify_biller(text: str, billers: list[dict]) -> Optional[dict]:
    """Return the first biller whose any identifier appears in the text."""
    upper = text.upper()
    for biller in billers:
        for ident in biller.get("identifiers", []):
            if ident.upper() in upper:
                return biller
    return None


def _search(patterns: list[str], text: str) -> Optional[str]:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
        if m and m.groups():
            return m.group(1).strip()
    return None


def parse_amount(raw: Optional[str]) -> Optional[float]:
    if not raw:
        return None
    cleaned = re.sub(r"[^0-9.]", "", raw.replace(",", ""))
    try:
        return round(float(cleaned), 2)
    except (ValueError, TypeError):
        return None


def parse_date(raw: Optional[str]) -> Optional[date]:
    """Parse common Malaysian bill date formats (BM month names included)."""
    if not raw:
        return None
    raw = raw.strip()

    # Numeric dd/mm/yyyy or dd-mm-yyyy (Malaysian day-first convention).
    m = re.match(r"^(\d{1,2})[\s\-/.](\d{1,2})[\s\-/.](\d{2,4})$", raw)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        y = _normalise_year(y)
        return _safe_date(y, mo, d)

    # dd MMM yyyy with a textual month (e.g. "05 Mac 2026", "5-Apr-26").
    m = re.match(r"^(\d{1,2})[\s\-/.]([A-Za-z]{3,9})[\s\-/.](\d{2,4})$", raw)
    if m:
        d = int(m.group(1))
        mo = _MONTHS.get(m.group(2).lower())
        y = _normalise_year(int(m.group(3)))
        if mo:
            return _safe_date(y, mo, d)
    return None


def _normalise_year(y: int) -> int:
    return y + 2000 if y < 100 else y


def _safe_date(y: int, mo: int, d: int) -> Optional[date]:
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def parse_bill(text: str, billers: list[dict], defaults: dict, source_file: str) -> BillRecord:
    """Extract a BillRecord from PDF text. Unknown billers are flagged."""
    biller = identify_biller(text, billers)

    if biller is None:
        rec = BillRecord(
            biller_name="UNKNOWN",
            biller_category="Unknown",
            supplier_code="",
            source_file=Path(source_file).name,
            needs_review=True,
        )
        rec.review_notes.append("biller-not-recognised")
        return rec

    fields = biller.get("fields", {})

    rec = BillRecord(
        biller_name=biller["name"],
        biller_category=biller.get("category", ""),
        supplier_code=biller.get("supplier_code", ""),
        source_file=Path(source_file).name,
        account_number=_search(fields.get("account_number", []), text),
        invoice_no=_search(fields.get("invoice_no", []), text),
        bill_date=parse_date(_search(fields.get("bill_date", []), text)),
        due_date=parse_date(_search(fields.get("due_date", []), text)),
        billing_period=_search(fields.get("billing_period", []), text),
        amount_excl_tax=parse_amount(_search(fields.get("amount_excl_tax", []), text)),
        tax_amount=parse_amount(_search(fields.get("tax_amount", []), text)),
        total_amount=parse_amount(_search(fields.get("total_amount", []), text)),
        currency=defaults.get("currency", "MYR"),
        gl_account=biller.get("gl_account", defaults.get("gl_account", "")),
        tax_code=biller.get("tax_code", defaults.get("tax_code", "")),
        tax_inclusive=defaults.get("tax_inclusive", "T"),
    )
    return rec.finalise()
