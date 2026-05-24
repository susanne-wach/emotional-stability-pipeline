"""
Lädt vertrauliche Konfigurations-Dateien (Fragen, Typen, Prompts) zur Laufzeit
aus dem privaten Netlify Blob Store 'config'.

Damit muss im (public) GitHub-Repo NICHTS vertrauliches liegen.
"""

from __future__ import annotations

import json
import os
import urllib.request
from pathlib import Path
from functools import lru_cache


NETLIFY_API_BASE = "https://api.netlify.com/api/v1"


def _fetch_blob(key: str) -> bytes:
    """Lädt eine Datei aus dem 'config'-Blob-Store."""
    site_id = os.environ.get("NETLIFY_SITE_ID")
    token = os.environ.get("NETLIFY_AUTH_TOKEN")
    if not site_id or not token:
        raise RuntimeError("NETLIFY_SITE_ID + NETLIFY_AUTH_TOKEN müssen gesetzt sein.")
    url = f"{NETLIFY_API_BASE}/blobs/{site_id}/site:config/{key}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


@lru_cache(maxsize=8)
def lade_fragen() -> list[dict]:
    """48 Tiefenfragen + Scoring-Gewichte."""
    raw = _fetch_blob("fragen.json")
    return json.loads(raw.decode("utf-8"))["fragen"]


@lru_cache(maxsize=8)
def lade_fragen_data() -> dict:
    """Komplettes fragen.json (mit meta, block_intros, intro_text)."""
    raw = _fetch_blob("fragen.json")
    return json.loads(raw.decode("utf-8"))


@lru_cache(maxsize=8)
def lade_typen_daten() -> dict[str, dict]:
    """5 Premium-Typen mit Beschreibungen."""
    raw = _fetch_blob("typen.json")
    return json.loads(raw.decode("utf-8"))["typen"]


@lru_cache(maxsize=8)
def lade_auswertungs_prompt() -> str:
    """Master-Prompt für die Premium-Auswertung."""
    return _fetch_blob("auswertung.md").decode("utf-8")


@lru_cache(maxsize=8)
def lade_auswertungs_leitlinien() -> str:
    """Tonalitäts-Leitlinien (für späteren Einbau in Prompt)."""
    return _fetch_blob("auswertungs-leitlinien.md").decode("utf-8")


def lade_fragen_lokal_temp() -> Path:
    """
    Schreibt fragen.json in eine temporäre Datei und gibt den Pfad zurück.
    Wird vom scoring.py gebraucht, das per Pfad arbeitet.
    """
    import tempfile
    raw = _fetch_blob("fragen.json")
    tmp = Path(tempfile.gettempdir()) / "yb_fragen.json"
    tmp.write_bytes(raw)
    return tmp


if __name__ == "__main__":
    # Self-Test
    print("→ Lade Fragen...")
    fragen = lade_fragen()
    print(f"  ✓ {len(fragen)} Fragen geladen")

    print("→ Lade Typen...")
    typen = lade_typen_daten()
    print(f"  ✓ {len(typen)} Typen geladen")

    print("→ Lade Auswertungs-Prompt...")
    prompt = lade_auswertungs_prompt()
    print(f"  ✓ {len(prompt)} Zeichen")
