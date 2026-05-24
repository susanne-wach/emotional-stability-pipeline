"""
process_submission.py — die Haupt-Pipeline.

Holt alle "pending" Submissions aus Netlify Blobs, verarbeitet sie und
sendet die Auswertung an die Käuferin.

Pro Submission:
  1. Antworten aus dem Blob laden
  2. Scoring (scoring.py)
  3. Claude API rufen für die Premium-Texte (auswertung.py)
  4. PDF im YB-Brand bauen (pdf_builder.py)
  5. Mail an Käuferin senden (send_mail.py)
  6. Blob auf "processed" markieren

Aufrufmodi:
  python3 src/process_submission.py                  → verarbeite alle pending
  python3 src/process_submission.py <blob_key>       → verarbeite genau diesen
  python3 src/process_submission.py --local <file>   → verarbeite eine lokale JSON
  python3 src/process_submission.py --dry-run        → kein Mail-Versand, kein Status-Update

Erforderliche Environment Variables:
  ANTHROPIC_API_KEY     → für Claude API
  NETLIFY_SITE_ID       → die Netlify-Site-ID (steht im Netlify-Dashboard → Site Settings → General)
  NETLIFY_AUTH_TOKEN    → Personal Access Token aus app.netlify.com/user/applications
  SMTP_HOST, SMTP_USER, SMTP_PASSWORD, SMTP_PORT, SMTP_USE_SSL  → für Mail-Versand

Optionale Environment Variables (für ActiveCampaign-Bridge-Sequenz):
  AC_API_URL            → z. B. https://yourbalance.api-us1.com
  AC_API_KEY            → API-Token aus AC → Einstellungen → Entwickler
  AC_TAG_TEST_ABGEGEBEN → numerische Tag-ID (Default: 164)
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import urllib.request
import urllib.error

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from scoring import score_antworten, format_score_report
from pdf_builder import render_auswertung_pdf
from send_mail import sende_auswertung_mail


# ============================================================
# NETLIFY BLOBS API
# ============================================================
# Die @netlify/blobs Funktion in der Cloud schreibt in einen Store namens "submissions".
# Wir lesen ihn hier über die Netlify-API.

NETLIFY_API_BASE = "https://api.netlify.com/api/v1"


def _netlify_blob_request(path: str, method: str = "GET", body: bytes | None = None) -> dict | bytes:
    """Authentifizierter Request gegen die Netlify Blobs-API (site-level store)."""
    site_id = os.environ.get("NETLIFY_SITE_ID")
    token = os.environ.get("NETLIFY_AUTH_TOKEN")
    if not site_id or not token:
        raise RuntimeError(
            "NETLIFY_SITE_ID und NETLIFY_AUTH_TOKEN müssen gesetzt sein."
        )
    url = f"{NETLIFY_API_BASE}/blobs/{site_id}{path}"
    req = urllib.request.Request(url, method=method, data=body)
    req.add_header("Authorization", f"Bearer {token}")
    if body is not None:
        req.add_header("Content-Type", "text/plain; charset=utf-8")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
            ct = resp.headers.get("Content-Type", "")
            if "application/json" in ct:
                return json.loads(raw.decode("utf-8"))
            return raw
    except urllib.error.HTTPError as e:
        raise RuntimeError(
            f"Netlify Blob-API {method} {path} failed: {e.code} {e.read().decode('utf-8', 'ignore')}"
        )


def list_pending_submissions() -> list[dict]:
    """Listet alle Blob-Einträge mit Status 'pending' im Store 'submissions'."""
    data = _netlify_blob_request("/site:submissions")
    items = data.get("blobs", []) if isinstance(data, dict) else []
    pending = []
    for b in items:
        key = b.get("key")
        if not key:
            continue
        try:
            sub = fetch_submission(key)
            status = sub.get("status") if isinstance(sub, dict) else None
            if status in (None, "pending"):
                pending.append(b)
        except Exception as e:
            print(f"  ⚠ Konnte Blob nicht laden: {key} ({e})")
    return pending


def fetch_submission(blob_key: str) -> dict:
    """Lädt eine einzelne Submission-JSON aus dem Blob-Store."""
    raw = _netlify_blob_request(f"/site:submissions/{blob_key}")
    if isinstance(raw, bytes):
        return json.loads(raw.decode("utf-8"))
    return raw


def _read_processed_keys_log() -> set[str]:
    """Liest die Liste der bereits verarbeiteten Blob-Keys aus _processed.log."""
    log_path = PROJECT_ROOT / "output" / "live" / "_processed.log"
    if not log_path.exists():
        return set()
    return {line.strip() for line in log_path.read_text(encoding="utf-8").splitlines() if line.strip()}


def _mark_processed_in_log(blob_key: str) -> None:
    """Hängt einen verarbeiteten Blob-Key an _processed.log an."""
    log_path = PROJECT_ROOT / "output" / "live" / "_processed.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(blob_key + "\n")


# ============================================================
# ACTIVECAMPAIGN — Tag setzen nach erfolgreicher PDF-Auswertung
# ============================================================
# Triggert die 4-teilige Bridge-Sequenz "From Heartbreak To Her" in AC.
# Wenn AC_API_URL oder AC_API_KEY fehlen, wird der Aufruf still übersprungen
# (Pipeline läuft auch ohne ActiveCampaign weiter).

def setze_ac_tag(email: str, tag_id: int | None = None) -> bool:
    """Setzt den TEST_ABGEGEBEN-Tag in ActiveCampaign für die angegebene E-Mail.

    Schritte:
      1. Kontakt anlegen oder aktualisieren (contact/sync)
      2. Tag dem Kontakt zuweisen (contactTags)

    Returns True bei Erfolg, False bei Fehler oder fehlender Konfiguration.
    Fehler werden geloggt, aber nicht weitergeworfen — Tag-Setzen ist optional.
    """
    api_url = os.environ.get("AC_API_URL")
    api_key = os.environ.get("AC_API_KEY")
    if not api_url or not api_key:
        print("  ⚠ AC_API_URL/AC_API_KEY nicht gesetzt — AC-Tag wird übersprungen")
        return False
    if not email:
        print("  ⚠ Keine E-Mail vorhanden — AC-Tag wird übersprungen")
        return False

    if tag_id is None:
        tag_id = int(os.environ.get("AC_TAG_TEST_ABGEGEBEN", "164"))

    api_url = api_url.rstrip("/")

    # 1. Kontakt anlegen oder updaten
    contact_payload = json.dumps({"contact": {"email": email}}).encode("utf-8")
    req = urllib.request.Request(
        f"{api_url}/api/3/contact/sync",
        data=contact_payload,
        headers={
            "Api-Token": api_key,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            contact_id = result.get("contact", {}).get("id")
            if not contact_id:
                print(f"  ⚠ AC contact/sync: keine contact_id zurückgegeben")
                return False
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        print(f"  ⚠ AC contact/sync HTTP {e.code}: {body[:200]}")
        return False
    except Exception as e:
        print(f"  ⚠ AC contact/sync Fehler: {e}")
        return False

    # 2. Tag zuweisen
    tag_payload = json.dumps({
        "contactTag": {"contact": int(contact_id), "tag": int(tag_id)}
    }).encode("utf-8")
    req = urllib.request.Request(
        f"{api_url}/api/3/contactTags",
        data=tag_payload,
        headers={
            "Api-Token": api_key,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            resp.read()  # Body verwerfen — Status 201 reicht
            print(f"  ✓ AC-Tag {tag_id} an {email} gesetzt (contact_id {contact_id})")
            return True
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        # 422 = Tag bereits gesetzt — kein Problem
        if e.code == 422 and "already" in body.lower():
            print(f"  ✓ AC-Tag {tag_id} war bereits an {email} gesetzt")
            return True
        print(f"  ⚠ AC contactTags HTTP {e.code}: {body[:200]}")
        return False
    except Exception as e:
        print(f"  ⚠ AC contactTags Fehler: {e}")
        return False


def mark_processed(blob_key: str, submission: dict) -> None:
    """
    Markiert eine Submission als verarbeitet:
    - Schreibt den Blob NEU mit upgedatetem Status (für already_submitted-Check)
    - Lokales Log als Backup
    """
    submission["status"] = "processed"
    submission["processed_at"] = datetime.utcnow().isoformat() + "Z"
    body = json.dumps(submission, ensure_ascii=False, indent=2).encode("utf-8")
    try:
        _netlify_blob_request(f"/site:submissions/{blob_key}", method="PUT", body=body)
    except Exception as e:
        print(f"  ⚠ Status-Update fehlgeschlagen: {e}")
    _mark_processed_in_log(blob_key)


# ============================================================
# AUSWERTUNGS-PIPELINE
# ============================================================

def process_one(submission: dict, blob_key: str | None = None, dry_run: bool = False) -> dict:
    """Verarbeitet eine einzelne Submission von A bis Z."""
    name = submission.get("name", "Liebe")
    email = submission.get("email")
    antworten = submission.get("antworten") or {}

    print(f"\n========== {name} <{email}> ==========")

    # 1. Scoring
    print("  → Scoring...")
    # Fragen + Typen kommen aus dem privaten Netlify Blob Store (nicht im public Repo)
    from config_loader import lade_fragen_lokal_temp
    fragen_tmp = lade_fragen_lokal_temp()
    auswertung = score_antworten(fragen_tmp, antworten)
    print(format_score_report(auswertung))

    # 2. Claude API rufen (auswertung.py)
    #    Wir importieren erst hier, damit das Skript auch ohne anthropic-Package
    #    starten kann (z. B. nur Scoring-Test).
    print("  → Claude API für Premium-Texte rufen...")
    from auswertung import (
        lade_master_prompt,
        lade_typen_daten,
        lade_fragen,
        build_user_prompt,
        call_claude,
        parse_json_output,
    )

    system_prompt, user_template = lade_master_prompt()
    typen_daten = lade_typen_daten()
    fragen = lade_fragen()
    user_prompt = build_user_prompt(user_template, name, auswertung, typen_daten, fragen)

    if dry_run:
        print("  ⚠ DRY-RUN: Claude wird NICHT gerufen.")
        # Wir nutzen die Test-Auswertung von Profil 1 als Stand-In
        test_auswertung_path = PROJECT_ROOT / "output" / "profil_1_warteschleifen_auswertung.json"
        auswertung_json = json.loads(test_auswertung_path.read_text(encoding="utf-8"))["auswertung"]
    else:
        raw = call_claude(system_prompt, user_prompt)
        auswertung_json = parse_json_output(raw)

    # 3. JSON-Output speichern (für Debug + ggf. spätere Web-Anzeige)
    output_dir = PROJECT_ROOT / "output" / "live"
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_name = "".join(c for c in name if c.isalnum() or c in ("-", "_")).rstrip()
    base = f"{timestamp}_{safe_name or 'kaeuferin'}"

    full_output = {
        "meta": {"name": name, "email": email, "timestamp": submission.get("timestamp")},
        "scores": auswertung.to_dict()["scores"],
        "typen": auswertung.to_dict()["typen"],
        "auswertung": auswertung_json,
    }
    json_path = output_dir / f"{base}_auswertung.json"
    json_path.write_text(json.dumps(full_output, ensure_ascii=False, indent=2), encoding="utf-8")

    # 4. PDF bauen
    print("  → PDF im YB-Brand bauen...")
    pdf_path = output_dir / f"Emotional-Stability-Analyse_{safe_name or 'Kaeuferin'}.docx"
    render_auswertung_pdf(json_path, pdf_path)

    # 5. Mail versenden
    print(f"  → Mail an {email} senden...")
    if not email:
        return {"ok": False, "info": "Keine E-Mail-Adresse in der Submission"}
    mail_res = sende_auswertung_mail(email, name, pdf_path, dry_run=dry_run)
    print(f"    {mail_res}")

    # 5b. AC-Tag setzen (Bridge-Sequenz "From Heartbreak To Her" triggern)
    #     Nur bei erfolgreichem Mail-Versand und nicht im dry-run.
    if not dry_run and mail_res.get("ok"):
        try:
            setze_ac_tag(email)
        except Exception as e:
            # AC-Fehler dürfen die Pipeline NICHT abbrechen — Auswertung ist raus.
            print(f"  ⚠ AC-Tag-Setzen fehlgeschlagen (nicht kritisch): {e}")

    # 6. Status updaten
    if not dry_run and blob_key:
        try:
            mark_processed(blob_key, submission)
            print("  ✓ Status auf 'processed' gesetzt")
        except Exception as e:
            print(f"  ⚠ Status konnte nicht gesetzt werden: {e}")

    return {"ok": mail_res["ok"], "info": mail_res["info"], "pdf": str(pdf_path)}


def process_local_file(path: Path, dry_run: bool = False) -> dict:
    """Verarbeitet eine lokale Antwort-JSON (z. B. zum Testen)."""
    if not path.exists():
        return {"ok": False, "info": f"Datei nicht gefunden: {path}"}
    data = json.loads(path.read_text(encoding="utf-8"))
    # Lokale Test-Profile haben Meta/Antworten unter unterschiedlichen Keys
    if "meta" in data and "antworten" in data:
        submission = {
            "name": data["meta"].get("name", "Test"),
            "email": data["meta"].get("email"),
            "antworten": data["antworten"],
            "timestamp": datetime.now().isoformat(),
        }
    else:
        submission = data
    return process_one(submission, blob_key=None, dry_run=dry_run)


def process_all_pending(dry_run: bool = False) -> int:
    """Verarbeitet alle noch nicht verarbeiteten Submissions im Netlify Blob Store."""
    print("→ Suche Submissions in Netlify Blobs...")
    all_blobs = list_pending_submissions()
    print(f"  Gefunden gesamt: {len(all_blobs)}")

    done = _read_processed_keys_log()
    pending = [b for b in all_blobs if b.get("key") and b["key"] not in done]
    print(f"  Davon noch nicht verarbeitet: {len(pending)}")
    if not pending:
        return 0

    erfolg = 0
    for blob in pending:
        key = blob.get("key")
        try:
            sub = fetch_submission(key)
            res = process_one(sub, blob_key=key, dry_run=dry_run)
            if res.get("ok"):
                erfolg += 1
        except Exception as e:
            print(f"  ✗ Fehler bei {key}: {e}")
    return erfolg


# ============================================================
# CLI
# ============================================================

def main():
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    args = [a for a in args if a != "--dry-run"]

    if not args:
        # Alle pending verarbeiten
        n = process_all_pending(dry_run=dry_run)
        print(f"\n✓ Fertig. {n} Submission(s) verarbeitet.")
        return

    if args[0] == "--local":
        if len(args) < 2:
            print("Usage: --local <pfad-zur-json>")
            sys.exit(1)
        res = process_local_file(Path(args[1]), dry_run=dry_run)
        print(f"\n{res}")
        return

    # Sonst: blob_key
    blob_key = args[0]
    print(f"→ Verarbeite Blob: {blob_key}")
    sub = fetch_submission(blob_key)
    res = process_one(sub, blob_key=blob_key, dry_run=dry_run)
    print(f"\n{res}")


if __name__ == "__main__":
    main()
