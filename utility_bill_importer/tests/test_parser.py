"""Unit tests for the parsing layer (no email/PDF/network needed)."""

import sys
from datetime import date
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.parser import parse_amount, parse_bill, parse_date  # noqa: E402

BASE = Path(__file__).resolve().parent.parent
BILLERS = yaml.safe_load((BASE / "billers.yaml").read_text())["billers"]
DEFAULTS = {"currency": "MYR", "gl_account": "910-000", "tax_code": "ZP", "tax_inclusive": "T"}


def test_parse_amount_strips_currency_and_commas():
    assert parse_amount("RM 1,358.40") == 1358.40
    assert parse_amount("358.40") == 358.40
    assert parse_amount(None) is None
    assert parse_amount("n/a") is None


def test_parse_date_numeric_day_first():
    assert parse_date("25/03/2026") == date(2026, 3, 25)
    assert parse_date("05-03-26") == date(2026, 3, 5)


def test_parse_date_bahasa_month():
    assert parse_date("05-Mac-2026") == date(2026, 3, 5)
    assert parse_date("28 Februari 2026") == date(2026, 2, 28)
    assert parse_date("01-Dis-2025") == date(2025, 12, 1)


def test_parse_tnb_bill_from_sample():
    text = (BASE / "sample_data" / "sample_tnb_text.txt").read_text()
    rec = parse_bill(text, BILLERS, DEFAULTS, "sample_tnb_text.pdf")

    assert rec.biller_name == "Tenaga Nasional Berhad"
    assert rec.supplier_code == "TNB"
    assert rec.account_number == "220012345678"
    assert rec.invoice_no == "1234567890"
    assert rec.bill_date == date(2026, 3, 5)
    assert rec.due_date == date(2026, 3, 25)
    assert rec.total_amount == 358.40
    assert rec.amount_excl_tax == 358.40   # no tax line -> net == total
    assert rec.tax_amount == 0.0
    assert rec.needs_review is False


def test_unknown_biller_is_flagged():
    rec = parse_bill("Some random receipt with no known biller", BILLERS, DEFAULTS, "x.pdf")
    assert rec.biller_name == "UNKNOWN"
    assert rec.needs_review is True


def test_export_dict_has_all_referenced_fields():
    """Every bill_field referenced in config.yaml columns must exist on the record."""
    cfg = yaml.safe_load((BASE / "config.yaml").read_text())
    text = (BASE / "sample_data" / "sample_tnb_text.txt").read_text()
    rec = parse_bill(text, BILLERS, DEFAULTS, "sample_tnb_text.pdf")
    export_dict = rec.as_export_dict()
    for bill_field in cfg["export"]["columns"].values():
        assert bill_field in export_dict, f"missing export field: {bill_field}"
