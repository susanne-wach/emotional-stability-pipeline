"""
E-Mail-Versand-Helper für die Emotional Stability Analyse.

Sendet eine HTML-Mail an die Käuferin — mit der DOCX-Auswertung als Anhang.
Absender: susanne@your-balance.at (laut Susannes Wunsch).

Konfiguration via .env oder Umgebungsvariablen:
  SMTP_HOST      z. B. smtp.world4you.com oder smtp.your-balance.at
  SMTP_PORT      z. B. 465 (SSL) oder 587 (STARTTLS)
  SMTP_USER      susanne@your-balance.at
  SMTP_PASSWORD  das Mailbox-Passwort
  SMTP_USE_SSL   "true" für SSL (Port 465), sonst STARTTLS
  REPLY_TO       optional, default = susanne@your-balance.at
"""

from __future__ import annotations

import os
import smtplib
import ssl
import subprocess
from email.message import EmailMessage
from pathlib import Path


def _which(cmd: str) -> str | None:
    """Sucht ein Kommando im PATH. Liefert None, wenn nicht gefunden."""
    import shutil
    return shutil.which(cmd)


def docx_to_pdf(docx_path: Path) -> Path:
    """
    Konvertiert DOCX → PDF.
    Strategie:
      1. macOS: Apple Pages via AppleScript
      2. Linux/Cloud (GitHub Actions): LibreOffice headless
      3. Sonst: DOCX-Pfad zurückgeben (Fallback)
    """
    import sys
    docx_path = Path(docx_path)
    pdf_path = docx_path.with_suffix(".pdf")

    # === Linux / GitHub Actions: LibreOffice ===
    soffice = _which("soffice") or _which("libreoffice")
    if soffice and sys.platform.startswith("linux"):
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf_path.parent.resolve()),
                 str(docx_path.resolve())],
                capture_output=True, text=True, timeout=120,
            )
            if result.returncode == 0 and pdf_path.exists():
                print(f"  ✓ PDF (LibreOffice): {pdf_path.name}")
                return pdf_path
            print(f"  ⚠ LibreOffice scheiterte: {result.stderr[:200]}")
        except Exception as e:
            print(f"  ⚠ LibreOffice Fehler: {e}")
        return docx_path

    # === macOS: Apple Pages ===
    script = Path(__file__).resolve().parent / "docx_to_pdf.applescript"
    if sys.platform == "darwin" and script.exists() and _which("osascript"):
        try:
            result = subprocess.run(
                ["osascript", str(script), str(docx_path.resolve()), str(pdf_path.resolve())],
                capture_output=True, text=True, timeout=60,
            )
            if result.returncode == 0 and pdf_path.exists():
                print(f"  ✓ PDF (Pages): {pdf_path.name}")
                return pdf_path
            print(f"  ⚠ Pages-Konvertierung scheiterte: {result.stderr[:200]}")
        except Exception as e:
            print(f"  ⚠ Pages Fehler: {e}")
        return docx_path

    print("  ⚠ Kein PDF-Konverter verfügbar, sende DOCX")
    return docx_path


MAIL_HTML_TEMPLATE = """<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="de">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Deine Emotional Stability Analyse</title>
  <!--[if mso]>
  <style type="text/css">
    body, table, td, p, a { font-family: Georgia, 'Times New Roman', serif !important; }
  </style>
  <![endif]-->
</head>
<body style="margin:0;padding:0;background:#FBF7F1;font-family:Georgia,'Times New Roman',serif;color:#5B2A4A;">
  <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background:#FBF7F1;">
    <tr><td align="center" style="padding:24px 16px;">

      <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="560" style="background:#ffffff;max-width:560px;width:100%;">
        <tr><td style="padding:32px;">

          <p style="margin:0 0 8px 0;text-transform:uppercase;letter-spacing:2px;font-size:11px;color:#CFAC5E;font-family:Georgia,serif;">DEINE PREMIUM-AUSWERTUNG</p>

          <h1 style="margin:0 0 24px 0;font-family:Georgia,'Times New Roman',serif;font-weight:normal;color:#DD3089;font-size:30px;line-height:1.2;">
            Liebe {NAME},
          </h1>

          <p style="margin:0 0 16px 0;line-height:1.7;font-size:15px;color:#5B2A4A;font-family:Georgia,serif;">
            hier kommt deine persönliche Emotional Stability Analyse™ — als PDF im Anhang.
          </p>

          <p style="margin:0 0 16px 0;line-height:1.7;font-size:15px;color:#5B2A4A;font-family:Georgia,serif;">
            Nimm dir Zeit dafür. Lies sie an einem ruhigen Moment. Du musst nichts sofort tun, nichts entscheiden, nichts beweisen. Diese Auswertung ist für dich — als Spiegel, nicht als Auftrag.
          </p>

          <p style="margin:0 0 24px 0;line-height:1.7;font-size:15px;color:#5B2A4A;font-family:Georgia,serif;">
            Wenn etwas in dir wackelt beim Lesen: Das ist okay. Es heisst nur, dass du gesehen wirst.
          </p>

          <p style="margin:24px 0;text-align:center;color:#CFAC5E;letter-spacing:8px;font-size:14px;">— ✦ —</p>

          <p style="margin:0 0 16px 0;line-height:1.7;font-size:15px;color:#5B2A4A;font-family:Georgia,serif;">
            Falls Du Fragen hast oder einfach kurz zurückschreiben möchtest: Diese E-Mail kommt direkt von mir. Antworte einfach.
          </p>

          <p style="margin:32px 0 0 0;font-family:Georgia,'Times New Roman',serif;font-style:italic;color:#5B2A4A;font-size:18px;line-height:1.4;">
            Von Herzen,<br>
            Susanne
          </p>

          <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top:32px;border-top:1px solid #E8E1D6;">
            <tr><td style="padding-top:16px;font-family:Georgia,serif;font-size:12px;color:#888888;line-height:1.6;">
              <a href="https://your-balance.at" style="color:#888888;text-decoration:none;">your-balance.at</a> ·
              <a href="mailto:susanne@your-balance.at" style="color:#888888;text-decoration:none;">susanne@your-balance.at</a><br>
              Diese E-Mail wurde an dich versendet, weil du die Emotional Stability Analyse™ gekauft hast.
            </td></tr>
          </table>

        </td></tr>
      </table>

    </td></tr>
  </table>
</body>
</html>
"""


def sende_auswertung_mail(
    empfaenger_email: str,
    empfaenger_name: str,
    pdf_pfad: str | Path,
    betreff: str | None = None,
    dry_run: bool = False,
) -> dict:
    """
    Sendet die Auswertung an die Käuferin.

    Args:
        empfaenger_email: an wen
        empfaenger_name: der Vorname (für die Anrede)
        pdf_pfad: Pfad zur DOCX-Datei (oder PDF, falls schon konvertiert)
        betreff: optional, default = "Deine Emotional Stability Analyse™"
        dry_run: wenn True, wird nichts versendet — nur die zusammengebaute Mail ausgegeben

    Returns:
        {"ok": bool, "info": "..."}
    """
    pdf_pfad = Path(pdf_pfad)
    if not pdf_pfad.exists():
        return {"ok": False, "info": f"PDF nicht gefunden: {pdf_pfad}"}

    # Wenn DOCX übergeben wurde: nach PDF konvertieren (via Apple Pages)
    if pdf_pfad.suffix.lower() == ".docx":
        pdf_pfad = docx_to_pdf(pdf_pfad)

    host = os.environ.get("SMTP_HOST")
    port = int(os.environ.get("SMTP_PORT", "465"))
    user = os.environ.get("SMTP_USER", "susanne@your-balance.at")
    password = os.environ.get("SMTP_PASSWORD")
    use_ssl = os.environ.get("SMTP_USE_SSL", "true").lower() == "true"
    reply_to = os.environ.get("REPLY_TO", "susanne@your-balance.at")

    if not host or not password:
        return {"ok": False, "info": "SMTP_HOST oder SMTP_PASSWORD nicht gesetzt"}

    msg = EmailMessage()
    msg["From"] = f"Susanne Wachter <{user}>"
    msg["To"] = empfaenger_email
    msg["Reply-To"] = reply_to
    msg["Subject"] = betreff or "Deine Emotional Stability Analyse™"

    # Plain-Text-Fallback
    msg.set_content(
        f"Liebe {empfaenger_name},\n\n"
        f"hier kommt deine persönliche Emotional Stability Analyse™ — als PDF im Anhang.\n\n"
        f"Nimm dir Zeit dafür. Du musst nichts sofort tun, nichts entscheiden, nichts beweisen.\n"
        f"Diese Auswertung ist für dich — als Spiegel, nicht als Auftrag.\n\n"
        f"Von Herzen,\nSusanne\n\n"
        f"—\nSusanne Wachter · Your Balance · your-balance.at"
    )

    # HTML-Version
    html = MAIL_HTML_TEMPLATE.replace("{NAME}", empfaenger_name)
    msg.add_alternative(html, subtype="html")

    # PDF/DOCX-Anhang
    suffix = pdf_pfad.suffix.lower()
    if suffix == ".pdf":
        maintype, subtype = "application", "pdf"
    elif suffix == ".docx":
        maintype, subtype = (
            "application",
            "vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        maintype, subtype = "application", "octet-stream"

    with open(pdf_pfad, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype=maintype,
            subtype=subtype,
            filename=pdf_pfad.name,
        )

    if dry_run:
        return {
            "ok": True,
            "info": "DRY-RUN: nichts versendet",
            "to": empfaenger_email,
            "subject": msg["Subject"],
        }

    # Versand
    try:
        context = ssl.create_default_context()
        if use_ssl:
            with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as srv:
                srv.login(user, password)
                srv.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=30) as srv:
                srv.starttls(context=context)
                srv.login(user, password)
                srv.send_message(msg)
        return {"ok": True, "info": f"Mail an {empfaenger_email} gesendet"}
    except Exception as e:
        return {"ok": False, "info": f"SMTP-Fehler: {e}"}


if __name__ == "__main__":
    # Self-Test (DRY-RUN)
    import sys

    if len(sys.argv) < 3:
        print("Usage: python send_mail.py <email> <name> [<pdf-pfad>]")
        sys.exit(1)
    email = sys.argv[1]
    name = sys.argv[2]
    pdf = sys.argv[3] if len(sys.argv) > 3 else (
        Path(__file__).resolve().parents[1]
        / "output"
        / "Emotional-Stability-Analyse_profil_1_warteschleifen.docx"
    )
    res = sende_auswertung_mail(email, name, pdf, dry_run=True)
    print(res)
