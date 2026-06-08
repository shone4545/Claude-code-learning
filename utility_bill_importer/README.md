# Utility Bill Importer

Automatically pull utility bills (TNB / electricity, Telco, Water, and more) from
your email, extract the key figures from the PDF, and produce an import file
ready for **SQL Accounting (eStream SQL Acc)** as a Purchase Invoice.

```
  Email (IMAP)  ─▶  PDF attachments  ─▶  Text / OCR  ─▶  Parse (per biller)
                                                              │
        SQL Accounting  ◀──  CSV / XLSX import file  ◀──  Dedupe + review flag
```

## What it does

1. **Fetch** – logs into your mailbox over IMAP (e.g. Gmail with an app
   password) and downloads PDF bill attachments from a label/folder.
2. **Extract** – reads the PDF text; if a bill is a scanned image it falls back
   to OCR (Tesseract).
3. **Parse** – identifies the biller and pulls account no., bill/due date,
   billing period and amounts using configurable regex patterns
   (Bahasa Malaysia *and* English layouts).
4. **Dedupe** – remembers bills it has already exported so re-runs don't create
   duplicates.
5. **Export** – writes a `SQLAcc_PurchaseInvoice_*.csv` (and `.xlsx`) whose
   columns map to SQL Account's Purchase Invoice import. Low-confidence rows are
   flagged **REVIEW** so you check them before importing.

## Quick start

```bash
cd utility_bill_importer
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Configure email access
cp .env.example .env
#  -> edit .env: IMAP_USER, IMAP_PASSWORD (Gmail App Password), IMAP_FOLDER

# Run it
python main.py                 # fetch from email + parse + export
python main.py --no-email      # only parse PDFs already in inbox_pdfs/
python main.py --inspect inbox_pdfs/somebill.pdf   # tune patterns on one file
```

Output lands in `output/`. Open the `.xlsx` to eyeball it, fix any REVIEW rows,
then import the file in SQL Account.

### Gmail setup (one time)

1. Enable 2-Step Verification on your Google account.
2. Create an **App Password**: <https://myaccount.google.com/apppasswords> –
   put the 16-char code in `.env` as `IMAP_PASSWORD` (not your login password).
3. In Gmail, create a filter that labels utility bills (e.g. label `Bills`) and
   set `IMAP_FOLDER=Bills` in `.env`.

### OCR (only needed for scanned/image PDFs)

Install the system tools once:

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-msa poppler-utils   # Debian/Ubuntu
```

Disable it in `config.yaml` (`ocr.enabled: false`) if all your bills already
have selectable text.

## Configuring it

| File           | What you change |
|----------------|-----------------|
| `.env`         | IMAP login + which folder/label to scan |
| `config.yaml`  | Scan window, OCR, output format, **SQL Account column mapping**, default GL account / tax code |
| `billers.yaml` | The biller registry – add/adjust providers and their regex patterns |

### Mapping to your SQL Account import template

`config.yaml -> export.columns` maps each **output column header** to a bill
field. SQL Account's import wizard then lets you map those headers to its
internal fields. If you already have a SQL Account import template, just rename
the header keys here to match it. Available bill fields are documented in
`src/models.py` (`BillRecord`).

> **Tax codes / GL accounts** are placeholders (`910-xxx`, `ZP`, `SV`). Set them
> to your real SQL Account chart-of-accounts codes. Electricity/water are
> generally not SST-taxable; telco services carry 8% service tax – adjust per
> biller in `billers.yaml`.

### Adding a new biller

1. Run `python main.py --inspect <pdf>` to see the raw text.
2. Copy a block in `billers.yaml`, set `name`, `supplier_code`, `identifiers`.
3. Adjust the regex patterns so each `fields` entry captures the right value.
4. Re-run `--inspect` until the parsed result looks correct.

## Important notes

- **The regex patterns are starting points.** Real TNB/Telco/Water PDFs vary by
  layout and year. Tune them against your own bills with `--inspect`. The
  bundled unit test (`tests/test_parser.py`) runs against a sample text file and
  shows the expected behaviour.
- **Always review before importing.** This tool prepares entries; it does not
  post to your accounts. Treat the import file as a draft.
- **Secrets stay local.** `.env`, downloaded PDFs and generated output are
  git-ignored.

## Tests

```bash
pip install pytest
python -m pytest tests/ -q
```

## Project layout

```
utility_bill_importer/
├── main.py              CLI entry point
├── config.yaml          general settings + SQL Account column mapping
├── billers.yaml         biller registry (regex patterns + supplier codes)
├── .env.example         IMAP credentials template
├── requirements.txt
├── src/
│   ├── config.py        load YAML + .env
│   ├── email_fetcher.py IMAP fetch -> save PDF attachments
│   ├── pdf_extractor.py PDF text + OCR fallback
│   ├── parser.py        text -> BillRecord (biller match, regex, dates, amounts)
│   ├── models.py        BillRecord data model
│   ├── exporter.py      BillRecord[] -> CSV / XLSX
│   ├── state.py         dedupe tracking
│   └── pipeline.py      orchestration
├── tests/test_parser.py
└── sample_data/         sample bill text for testing
```
