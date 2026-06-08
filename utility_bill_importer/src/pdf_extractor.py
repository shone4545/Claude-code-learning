"""Extract plain text from a PDF, with OCR fallback for scanned bills."""

from __future__ import annotations

from pathlib import Path

# Minimum characters of embedded text before we trust it (else OCR fallback).
_MIN_TEXT_CHARS = 40


def extract_text(pdf_path: Path, ocr_cfg: dict | None = None) -> str:
    """Return all text in the PDF. Tries embedded text first, then OCR."""
    text = _extract_embedded_text(pdf_path)
    if len(text.strip()) >= _MIN_TEXT_CHARS:
        return text

    ocr_cfg = ocr_cfg or {}
    if ocr_cfg.get("enabled", True):
        ocr_text = _extract_ocr_text(pdf_path, ocr_cfg)
        if len(ocr_text.strip()) > len(text.strip()):
            return ocr_text
    return text


def _extract_embedded_text(pdf_path: Path) -> str:
    try:
        import pdfplumber
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pdfplumber not installed. Run: pip install -r requirements.txt") from exc

    chunks: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def _extract_ocr_text(pdf_path: Path, ocr_cfg: dict) -> str:
    """OCR scanned PDFs. Requires tesseract-ocr + poppler-utils installed."""
    try:
        import pytesseract
        from pdf2image import convert_from_path
    except ImportError:
        # OCR libs not available — silently skip; caller keeps embedded text.
        return ""

    try:
        images = convert_from_path(str(pdf_path), dpi=int(ocr_cfg.get("dpi", 300)))
    except Exception:
        return ""

    lang = ocr_cfg.get("languages", "eng")
    out: list[str] = []
    for img in images:
        try:
            out.append(pytesseract.image_to_string(img, lang=lang))
        except Exception:
            # Fall back to English-only if the requested lang pack is missing.
            out.append(pytesseract.image_to_string(img))
    return "\n".join(out)
