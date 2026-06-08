"""Load configuration from YAML + environment (.env)."""

from __future__ import annotations

import os
from pathlib import Path

import yaml
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent


def _load_yaml(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def load_config(config_path: str | None = None, billers_path: str | None = None) -> dict:
    """Return a single dict with `config`, `billers`, and `email_creds`."""
    load_dotenv(BASE_DIR / ".env")

    cfg = _load_yaml(Path(config_path) if config_path else BASE_DIR / "config.yaml")
    billers = _load_yaml(Path(billers_path) if billers_path else BASE_DIR / "billers.yaml")

    email_creds = {
        "host": os.getenv("IMAP_HOST", "imap.gmail.com"),
        "port": int(os.getenv("IMAP_PORT", "993")),
        "user": os.getenv("IMAP_USER", ""),
        "password": os.getenv("IMAP_PASSWORD", ""),
        "folder": os.getenv("IMAP_FOLDER", "INBOX"),
    }

    return {
        "config": cfg,
        "billers": billers.get("billers", []),
        "email_creds": email_creds,
        "base_dir": BASE_DIR,
    }


def resolve_path(base_dir: Path, rel: str) -> Path:
    """Resolve a possibly-relative path against the project base dir."""
    p = Path(rel)
    return p if p.is_absolute() else base_dir / p
