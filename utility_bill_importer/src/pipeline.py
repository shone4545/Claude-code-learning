"""Orchestrate: fetch -> extract -> parse -> dedupe -> export."""

from __future__ import annotations

from pathlib import Path

from .config import load_config, resolve_path
from .email_fetcher import fetch_bill_pdfs
from .exporter import export
from .parser import parse_bill
from .pdf_extractor import extract_text
from .state import filter_new, load_state, save_state


def run(fetch_email: bool = True, config_path: str | None = None) -> dict:
    """Run the full pipeline. Returns a summary dict."""
    bundle = load_config(config_path)
    cfg = bundle["config"]
    billers = bundle["billers"]
    base_dir = bundle["base_dir"]

    paths = cfg.get("paths", {})
    inbox = resolve_path(base_dir, paths.get("inbox_pdfs", "inbox_pdfs"))
    out_dir = resolve_path(base_dir, paths.get("output", "output"))
    state_path = resolve_path(base_dir, paths.get("state_file", "state.json"))
    inbox.mkdir(parents=True, exist_ok=True)

    # 1. Fetch new PDFs from email (optional — can also use --no-email).
    if fetch_email:
        fetch_bill_pdfs(bundle["email_creds"], cfg.get("email", {}), inbox)

    # 2 + 3. Extract text and parse every PDF currently in the inbox.
    defaults = cfg.get("defaults", {})
    ocr_cfg = cfg.get("ocr", {})
    records = []
    for pdf in sorted(inbox.glob("*.pdf")):
        text = extract_text(pdf, ocr_cfg)
        records.append(parse_bill(text, billers, defaults, str(pdf)))

    # 4. Dedupe against previously-processed bills.
    processed = load_state(state_path)
    new_records, updated = filter_new(records, processed)

    # 5. Export.
    written = export(new_records, cfg, out_dir) if new_records else []
    if new_records:
        save_state(state_path, updated)

    return {
        "pdfs_seen": len(records),
        "new_records": len(new_records),
        "needs_review": sum(1 for r in new_records if r.needs_review),
        "files_written": [str(p) for p in written],
        "records": new_records,
    }


def inspect_pdf(pdf_path: str, config_path: str | None = None) -> str:
    """Dump extracted text + parse result for tuning biller patterns."""
    bundle = load_config(config_path)
    cfg = bundle["config"]
    text = extract_text(Path(pdf_path), cfg.get("ocr", {}))
    rec = parse_bill(text, bundle["billers"], cfg.get("defaults", {}), pdf_path)

    lines = ["===== EXTRACTED TEXT =====", text, "", "===== PARSED RESULT ====="]
    for k, v in rec.as_export_dict().items():
        lines.append(f"  {k:16}: {v}")
    return "\n".join(lines)
