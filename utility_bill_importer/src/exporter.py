"""Write BillRecords to a SQL Accounting Purchase Invoice import file."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

from .models import BillRecord


def _rows(records: list[BillRecord], columns: dict[str, str]) -> tuple[list[str], list[list]]:
    """Build header + value rows from the export.columns mapping."""
    headers = list(columns.keys())
    rows: list[list] = []
    for rec in records:
        data = rec.as_export_dict()
        rows.append([data.get(bill_field, "") for bill_field in columns.values()])
    return headers, rows


def export(records: list[BillRecord], cfg: dict, out_dir: Path) -> list[Path]:
    """Write import file(s) per config.yaml -> export. Returns written paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    export_cfg = cfg.get("export", {})
    columns = export_cfg.get("columns", {})
    formats = [f.lower() for f in export_cfg.get("formats", ["csv"])]
    prefix = export_cfg.get("filename_prefix", "SQLAcc_PurchaseInvoice")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    headers, rows = _rows(records, columns)
    written: list[Path] = []

    if "csv" in formats:
        path = out_dir / f"{prefix}_{stamp}.csv"
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.writer(fh)
            writer.writerow(headers)
            writer.writerows(rows)
        written.append(path)

    if "xlsx" in formats:
        written.append(_write_xlsx(out_dir / f"{prefix}_{stamp}.xlsx", headers, rows, records))

    return written


def _write_xlsx(path: Path, headers: list[str], rows: list[list], records: list[BillRecord]) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "PurchaseInvoice"

    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    review_fill = PatternFill("solid", fgColor="FCE4D6")

    ws.append(headers)
    for cell in ws[1]:
        cell.font = hdr_font
        cell.fill = hdr_fill

    for rec, row in zip(records, rows):
        ws.append(row)
        if rec.needs_review:
            for cell in ws[ws.max_row]:
                cell.fill = review_fill

    for col_idx, header in enumerate(headers, start=1):
        width = max(len(str(header)), 12) + 2
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    ws.freeze_panes = "A2"
    wb.save(path)
    return path
