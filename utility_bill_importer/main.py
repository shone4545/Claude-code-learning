#!/usr/bin/env python3
"""Utility Bill Importer — CLI entry point.

Examples
--------
  # Full run: pull new bills from email, parse, export import file
  python main.py

  # Parse PDFs already sitting in inbox_pdfs/ without touching email
  python main.py --no-email

  # Inspect one PDF to tune regex patterns in billers.yaml
  python main.py --inspect inbox_pdfs/somebill.pdf
"""

from __future__ import annotations

import argparse
import sys

from src.pipeline import inspect_pdf, run


def main() -> int:
    ap = argparse.ArgumentParser(description="Email utility bills -> SQL Accounting import file.")
    ap.add_argument("--no-email", action="store_true",
                    help="Skip IMAP fetch; only process PDFs already in inbox_pdfs/.")
    ap.add_argument("--inspect", metavar="PDF",
                    help="Print extracted text + parse result for one PDF, then exit.")
    ap.add_argument("--config", metavar="PATH", help="Path to an alternate config.yaml.")
    args = ap.parse_args()

    if args.inspect:
        print(inspect_pdf(args.inspect, args.config))
        return 0

    try:
        summary = run(fetch_email=not args.no_email, config_path=args.config)
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"PDFs processed : {summary['pdfs_seen']}")
    print(f"New bills      : {summary['new_records']}")
    print(f"Need review    : {summary['needs_review']}")
    if summary["files_written"]:
        print("Import file(s) :")
        for f in summary["files_written"]:
            print(f"  - {f}")
    else:
        print("No new bills to export.")

    if summary["needs_review"]:
        print("\nReview rows flagged 'REVIEW' before importing into SQL Account:")
        for r in summary["records"]:
            if r.needs_review:
                print(f"  - {r.source_file}: {', '.join(r.review_notes)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
