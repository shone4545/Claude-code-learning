"""Data structures shared across the pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import date
from typing import Optional


# Fields that MUST be confidently extracted for a bill to import cleanly.
# If any is missing, the record is flagged needs_review.
REQUIRED_FIELDS = ("account_number", "bill_date", "total_amount")


@dataclass
class BillRecord:
    """One utility bill, normalised and ready for export.

    The attribute names here are exactly the `bill_field` values you can
    reference from config.yaml -> export.columns.
    """

    # --- identity / source ---
    biller_name: str                         # e.g. "Tenaga Nasional Berhad"
    biller_category: str                     # Electricity / Water / Telco
    supplier_code: str                       # SQL Account supplier code
    source_file: str                         # PDF filename it came from

    # --- extracted bill data ---
    account_number: Optional[str] = None
    invoice_no: Optional[str] = None
    bill_date: Optional[date] = None
    due_date: Optional[date] = None
    billing_period: Optional[str] = None
    amount_excl_tax: Optional[float] = None
    tax_amount: Optional[float] = None
    total_amount: Optional[float] = None

    # --- accounting metadata ---
    currency: str = "MYR"
    gl_account: str = ""
    tax_code: str = ""
    tax_inclusive: str = "T"
    description: str = ""                     # auto-built line description

    # --- review / audit ---
    needs_review: bool = False
    review_notes: list[str] = field(default_factory=list)

    def finalise(self) -> "BillRecord":
        """Derive amounts and set the review flag. Call after extraction."""
        # If we only have a total, treat it as tax-inclusive and back out tax.
        if self.total_amount is not None:
            if self.tax_amount is not None and self.amount_excl_tax is None:
                self.amount_excl_tax = round(self.total_amount - self.tax_amount, 2)
            elif self.amount_excl_tax is None:
                # No tax line found: assume whole amount is the net expense.
                self.amount_excl_tax = self.total_amount
                if self.tax_amount is None:
                    self.tax_amount = 0.0

        if not self.description:
            self.description = self._build_description()

        for f in REQUIRED_FIELDS:
            if getattr(self, f) in (None, ""):
                self.needs_review = True
                self.review_notes.append(f"missing:{f}")

        return self

    def _build_description(self) -> str:
        """Human-readable line description for the accounting entry."""
        parts = [f"{self.biller_category} bill" if self.biller_category else "Utility bill"]
        if self.account_number:
            parts.append(f"Acc {self.account_number}")
        if self.billing_period:
            parts.append(self.billing_period)
        return " - ".join(parts)

    def as_export_dict(self) -> dict:
        """Flat dict with display-friendly values for CSV/XLSX writers."""
        d = asdict(self)
        # Dates -> ISO strings; SQL Account import is happy with dd/mm/yyyy too,
        # but ISO avoids ambiguity. The exporter can reformat if needed.
        for k in ("bill_date", "due_date"):
            d[k] = self.__dict__[k].strftime("%d/%m/%Y") if self.__dict__[k] else ""
        d["tax_inclusive"] = self.tax_inclusive
        d["needs_review"] = "REVIEW" if self.needs_review else ""
        d["review_notes"] = "; ".join(self.review_notes)
        return d
