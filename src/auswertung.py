"""
Generiert per Claude API die personalisierte 7-teilige Premium-Auswertung.

Ablauf:
1. Antworten + Scores einlesen
2. Master-Prompt aus prompts/auswertung.md laden
3. System- und User-Prompt mit Käuferin-Daten füllen
4. Claude API rufen (claude-opus-4-7)
5. JSON-Output parsen + speichern

Voraussetzung: ANTHROPIC_API_KEY in der Umgebung.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any

# Anthropic SDK laden (anthropic ist in Susannes Setup für unternehmer-analyse bereits installiert)
try:
    import anthropic
except ImportError:
    print("✗ anthropic-Package fehlt. Installation:  pip install anthropic", file=sys.stderr)
    sys.exit(1)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from scoring import score_antworten, Auswertung, TYP_KEYS, SCORE_KEYS


CLAUDE_MODEL = "claude-opus-4-5"  # Susannes Standard für sorgfältige Generierung


def lade_master_prompt() -> tuple[str, str]:
    """Lädt Master-Prompt aus dem privaten Netlify Blob Store."""
    from config_loader import lade_auswertungs_prompt
    text = lade_auswertungs_prompt()

    m_sys = re.search(r"##\s*SYSTEM-PROMPT\s*\n(.*?)\n##\s*USER-PROMPT", text, re.DOTALL)
    system_prompt = m_sys.group(1).strip() if m_sys else ""

    m_usr = re.search(r"##\s*USER-PROMPT\s*\(template\)\s*\n(.*)", text, re.DOTALL)
    user_template = m_usr.group(1).strip() if m_usr else ""

    return system_prompt, user_template


def lade_typen_daten() -> dict[str, dict]:
    from config_loader import lade_typen_daten as _lade
    return _lade()


def lade_fragen() -> list[dict]:
    from config_loader import lade_fragen as _lade
    return _lade()


def format_typ_definition(typ_key: str, typen_daten: dict) -> str:
    """Verbale Beschreibung des Typs für den Prompt."""
    t = typen_daten.get(typ_key, {})
    lines = [
        f"Name: {t.get('name', typ_key)}",
        f"Innerer Satz: „{t.get('innerer_satz', '')}\"",
        "Hauptmuster:",
    ]
    for m in t.get("hauptmuster", []):
        lines.append(f"  - {m}")
    lines += [
        f"Kernproblem: {t.get('kernproblem', '')}",
        f"Hauptschmerz: {t.get('hauptschmerz', '')}",
        f"Blinder Fleck: {t.get('blinder_fleck', '')}",
        f"Validierung (Ton): {t.get('validierung', '')}",
    ]
    return "\n".join(lines)


def format_auswahl_antworten(
    antworten: dict[str, Any], fragen: list[dict], max_items: int = 18
) -> str:
    """
    Wählt die qualitativ interessantesten Antworten aus (Multi-Choice, Single-Choice
    mit String-Werten) — das sind die, die persönliche Sprache liefern.
    Skala-Antworten werden nur ergänzend mitgegeben (Highlights bei extremen Werten).
    """
    fragen_idx = {q["id"]: q for q in fragen}
    interessant: list[str] = []

    # Block-Reihenfolge beibehalten
    for q in fragen:
        if q["id"] not in antworten:
            continue
        antwort = antworten[q["id"]]

        if q["type"] == "scale":
            # Nur extreme Werte (≤3 oder ≥8) mitgeben
            try:
                v = int(antwort)
            except (TypeError, ValueError):
                continue
            if v <= 3 or v >= 8:
                interessant.append(
                    f"- {q['frage']}\n  → Skala {v}/10 ({q.get('label_low','')} ... {q.get('label_high','')})"
                )

        elif q["type"] == "single":
            opt = next((o for o in q["options"] if o["value"] == antwort), None)
            if opt is None:
                continue
            interessant.append(f"- {q['frage']}\n  → „{opt['label']}\"")

        elif q["type"] == "multi":
            chosen = antwort if isinstance(antwort, list) else []
            labels = []
            for v in chosen:
                opt = next((o for o in q["options"] if o["value"] == v), None)
                if opt:
                    labels.append(f"„{opt['label']}\"")
            if labels:
                interessant.append(
                    f"- {q['frage']}\n  → " + " · ".join(labels)
                )

    # Limitieren — die ersten interessantesten reichen meist (Fokus auf Spiegelung)
    # Aber: zu wenig ist auch schlecht. 18-25 Items ist ein guter Mittelweg.
    if len(interessant) > max_items:
        # Wir nehmen lieber die ersten (chronologisch, also Block 1-2-3-4-5-6)
        # Aber strecken über die Blöcke — etwas Auswahl statt nur Block 1
        # Simpler: alles, max 25
        interessant = interessant[: max_items + 7]

    return "\n".join(interessant)


def build_user_prompt(
    template: str,
    name: str,
    auswertung: Auswertung,
    typen_daten: dict,
    fragen: list[dict],
) -> str:
    s = auswertung.scores
    t = auswertung.typen

    haupttyp_def = format_typ_definition(t.haupttyp, typen_daten)

    if t.nebentyp:
        nebentyp_block = (
            f"{typen_daten.get(t.nebentyp, {}).get('name', t.nebentyp)} ({t.nebentyp_score}%)\n\n"
            + format_typ_definition(t.nebentyp, typen_daten)
        )
    else:
        nebentyp_block = "(Kein signifikantes Zweitmuster — Hauptmuster dominiert klar. Schreibe in teil_4_zweitmuster: \"vorhanden\": false und lasse die anderen Felder leer.)"

    auswahl = format_auswahl_antworten(auswertung.rohantworten, fragen)

    return (
        template.replace("{name}", name)
        .replace("{esi}", f"{s.emotional_stability_index:.1f}")
        .replace("{emotional_activation}", f"{s.emotional_activation:.0f}")
        .replace("{nervous_system_load}", f"{s.nervous_system_load:.0f}")
        .replace("{selbstverlust}", f"{s.selbstverlust:.0f}")
        .replace("{hoffnungsschleifen}", f"{s.hoffnungsschleifen:.0f}")
        .replace("{haupttyp_name}", typen_daten.get(t.haupttyp, {}).get("name", t.haupttyp))
        .replace("{haupttyp_score}", f"{t.haupttyp_score:.0f}")
        .replace("{haupttyp_definition}", haupttyp_def)
        .replace("{nebentyp_block}", nebentyp_block)
        .replace("{auswahl_antworten}", auswahl)
    )


def call_claude(system_prompt: str, user_prompt: str) -> str:
    """Ruft Claude API auf und gibt den Roh-Text zurück."""
    client = anthropic.Anthropic()  # Liest ANTHROPIC_API_KEY aus Umgebung
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return message.content[0].text


def parse_json_output(raw: str) -> dict:
    """Parsed das JSON aus Claudes Antwort. Robust gegen Code-Fences UND innere deutsche Anführungszeichen."""
    raw = raw.strip()
    # Code-Fences entfernen (falls Claude sie trotz Anweisung dazu setzt)
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

    # Innere ASCII-Quotes in deutschen Zitaten („...") durch typografische ersetzen
    out = []
    pending = False
    i = 0
    while i < len(raw):
        c = raw[i]
        if c == "\\" and i + 1 < len(raw):
            out.append(c)
            out.append(raw[i + 1])
            i += 2
            continue
        if c == "„":
            pending = True
            out.append(c)
        elif c == "“" and pending:
            pending = False
            out.append(c)
        elif c == '"' and pending:
            out.append("“")
            pending = False
        else:
            out.append(c)
        i += 1
    fixed = "".join(out)

    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        # Notfall: Roh-Output für Debug speichern
        from pathlib import Path
        dbg = Path(__file__).resolve().parents[1] / "output" / "live" / "_claude_raw.txt"
        dbg.parent.mkdir(parents=True, exist_ok=True)
        dbg.write_text(raw, encoding="utf-8")
        raise


def generiere_auswertung(profil_path: str | Path) -> dict:
    """
    Hauptfunktion: Profil-JSON laden → Scoring → Claude API → JSON-Auswertung speichern.
    """
    profil_path = Path(profil_path)
    profil = json.loads(profil_path.read_text(encoding="utf-8"))

    name = profil["meta"]["name"]
    print(f"  → Scoring für {name}...")
    auswertung = score_antworten(PROJECT_ROOT / "fragen.json", profil["antworten"])

    print(f"  → Lade Master-Prompt + Typen-Daten...")
    system_prompt, user_template = lade_master_prompt()
    typen_daten = lade_typen_daten()
    fragen = lade_fragen()

    user_prompt = build_user_prompt(user_template, name, auswertung, typen_daten, fragen)

    print(f"  → Rufe Claude API ({CLAUDE_MODEL}) — kann 30-60s dauern...")
    raw = call_claude(system_prompt, user_prompt)

    print(f"  → Parse JSON-Output...")
    auswertung_json = parse_json_output(raw)

    # Output speichern
    output_path = PROJECT_ROOT / "output" / f"{profil_path.stem}_auswertung.json"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(
            {
                "meta": profil["meta"],
                "scores": auswertung.to_dict()["scores"],
                "typen": auswertung.to_dict()["typen"],
                "auswertung": auswertung_json,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(f"  ✓ Gespeichert: {output_path}")
    return auswertung_json


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python auswertung.py <profil.json> [<profil2.json> ...]")
        sys.exit(1)

    for p in sys.argv[1:]:
        print(f"\n========== {p} ==========")
        generiere_auswertung(p)
