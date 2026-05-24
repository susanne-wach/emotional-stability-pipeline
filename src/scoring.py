"""
Scoring-Engine für die Emotional Stability Analyse.

Eingabe: Antworten einer Käuferin (Dict frage_id -> antwort)
Ausgabe: 4 Tiefenscores + Emotional Stability Index + Haupt-/Nebentyp + Kaufmotivation

Score-Logik:
- Jede Frage trägt mit Gewichten zu mehreren Scores bei.
- Skala-Fragen (1-10) werden auf 0-1 normalisiert.
- Single-Choice: Gewichte der gewählten Option.
- Multi-Choice: Summe der Gewichte aller gewählten Optionen.
- Final werden Scores auf 0-100% normalisiert (anhand des theoretischen Max).
- Emotional Stability Index = 100 - Durchschnitt der 4 Belastungsscores.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any


SCORE_KEYS = [
    "emotional_activation",
    "nervous_system_load",
    "selbstverlust",
    "hoffnungsschleifen",
]

TYP_KEYS = [
    "warteschleifen_frau",
    "tiefverbundene_feinfuehlige",
    "sicherheits_suchende",
    "emotionale_ueberfunktioniererin",
    "erschoepfte_herzfrau",
]


@dataclass
class Scores:
    """Die 4 Tiefenscores + Hauptindex + Kaufmotivation, alle 0-100."""

    emotional_activation: float = 0.0
    nervous_system_load: float = 0.0
    selbstverlust: float = 0.0
    hoffnungsschleifen: float = 0.0
    emotional_stability_index: float = 0.0
    kaufmotivation: float = 0.0


@dataclass
class TypProfil:
    """Haupt- + Nebentyp (Nebentyp nur, wenn signifikant)."""

    haupttyp: str
    haupttyp_score: float
    nebentyp: str | None
    nebentyp_score: float
    alle_typen: dict[str, float] = field(default_factory=dict)


@dataclass
class Auswertung:
    scores: Scores
    typen: TypProfil
    rohantworten: dict[str, Any]
    meta: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "scores": asdict(self.scores),
            "typen": asdict(self.typen),
            "rohantworten": self.rohantworten,
            "meta": self.meta,
        }


def _normalize_scale(value: int | float) -> float:
    """1-10 → 0-1. Werte außerhalb werden geklammert."""
    return max(0.0, min(1.0, (float(value) - 1) / 9.0))


def _add_dict(target: dict[str, float], source: dict[str, float], factor: float = 1.0):
    for k, v in source.items():
        target[k] = target.get(k, 0.0) + v * factor


def _option_for(question: dict, value: Any) -> dict | None:
    # Tolerantes Matching: Frontend kann numerische values als String liefern
    for opt in question.get("options", []):
        ov = opt.get("value")
        if ov == value:
            return opt
        # Cross-Type-Match: "5" == 5, 5 == "5"
        if isinstance(ov, (int, float)) and isinstance(value, str):
            try:
                if ov == int(value) or ov == float(value):
                    return opt
            except (TypeError, ValueError):
                pass
        if isinstance(ov, str) and isinstance(value, (int, float)):
            try:
                if int(ov) == value or float(ov) == value:
                    return opt
            except (TypeError, ValueError):
                pass
    return None


def score_antworten(
    fragen_path: str | Path,
    antworten: dict[str, Any],
    nebentyp_threshold: float = 0.5,
) -> Auswertung:
    """
    Berechnet aus den Antworten eine vollständige Auswertung.

    Args:
        fragen_path: Pfad zur fragen.json
        antworten: {frage_id: antwort}.
                   - scale: int 1-10
                   - single: str (option value)
                   - multi: list[str] (option values)
        nebentyp_threshold: Nebentyp wird nur ausgewiesen, wenn sein Score
                            mindestens diesen Anteil vom Haupttyp erreicht (default 50%).

    Returns:
        Auswertung mit Scores, Typprofil und Roh-Antworten.
    """
    fragen_path = Path(fragen_path)
    with open(fragen_path, "r", encoding="utf-8") as f:
        fragen_data = json.load(f)

    fragen_idx = {q["id"]: q for q in fragen_data["fragen"]}

    score_sum: dict[str, float] = {k: 0.0 for k in SCORE_KEYS}
    score_max: dict[str, float] = {k: 0.0 for k in SCORE_KEYS}
    typ_sum: dict[str, float] = {k: 0.0 for k in TYP_KEYS}
    typ_max: dict[str, float] = {k: 0.0 for k in TYP_KEYS}
    kauf_sum = 0.0
    kauf_max = 0.0

    fragen_beantwortet = 0
    fragen_uebersprungen: list[str] = []

    for q in fragen_data["fragen"]:
        qid = q["id"]
        if qid not in antworten:
            fragen_uebersprungen.append(qid)
            continue
        fragen_beantwortet += 1
        antwort = antworten[qid]

        if q["type"] == "scale":
            try:
                antwort = int(antwort) if not isinstance(antwort, (int, float)) else antwort
            except (TypeError, ValueError):
                continue
            wert = _normalize_scale(antwort)
            scoring = {k: v for k, v in q.get("scoring", {}).items() if k in SCORE_KEYS}
            typ_signal = {k: v for k, v in q.get("typ_signal", {}).items() if k in TYP_KEYS}

            _add_dict(score_sum, {k: wert * v for k, v in scoring.items()})
            _add_dict(score_max, scoring)  # max bei skala = 10 → faktor 1.0
            _add_dict(typ_sum, {k: wert * v for k, v in typ_signal.items()})
            _add_dict(typ_max, typ_signal)

            if "kaufmotivation" in q:
                kauf_sum += wert * q["kaufmotivation"]
                kauf_max += q["kaufmotivation"]

        elif q["type"] == "single":
            opt = _option_for(q, antwort)
            if opt is None:
                continue
            scoring = {k: v for k, v in opt.get("scoring", {}).items() if k in SCORE_KEYS}
            typ_signal = {k: v for k, v in opt.get("typ_signal", {}).items() if k in TYP_KEYS}

            _add_dict(score_sum, scoring)
            for k in SCORE_KEYS:
                score_max[k] += max((o.get("scoring", {}).get(k, 0) for o in q["options"]), default=0)
            _add_dict(typ_sum, typ_signal)
            for k in TYP_KEYS:
                typ_max[k] += max((o.get("typ_signal", {}).get(k, 0) for o in q["options"]), default=0)

            if "kaufmotivation" in opt:
                kauf_sum += opt["kaufmotivation"]
            kauf_max += max((o.get("kaufmotivation", 0) for o in q["options"]), default=0)

        elif q["type"] == "multi":
            chosen_values = antwort if isinstance(antwort, list) else []
            for value in chosen_values:
                opt = _option_for(q, value)
                if opt is None:
                    continue
                scoring = {k: v for k, v in opt.get("scoring", {}).items() if k in SCORE_KEYS}
                typ_signal = {k: v for k, v in opt.get("typ_signal", {}).items() if k in TYP_KEYS}
                _add_dict(score_sum, scoring)
                _add_dict(typ_sum, typ_signal)
                if "kaufmotivation" in opt:
                    kauf_sum += opt["kaufmotivation"]
            # max = alle Optionen gewählt
            for k in SCORE_KEYS:
                score_max[k] += sum(o.get("scoring", {}).get(k, 0) for o in q["options"])
            for k in TYP_KEYS:
                typ_max[k] += sum(o.get("typ_signal", {}).get(k, 0) for o in q["options"])
            kauf_max += sum(o.get("kaufmotivation", 0) for o in q["options"])

    # Normalize auf 0-100
    def pct(s: float, m: float) -> float:
        if m <= 0:
            return 0.0
        return max(0.0, min(100.0, (s / m) * 100.0))

    scores = Scores(
        emotional_activation=pct(score_sum["emotional_activation"], score_max["emotional_activation"]),
        nervous_system_load=pct(score_sum["nervous_system_load"], score_max["nervous_system_load"]),
        selbstverlust=pct(score_sum["selbstverlust"], score_max["selbstverlust"]),
        hoffnungsschleifen=pct(score_sum["hoffnungsschleifen"], score_max["hoffnungsschleifen"]),
        kaufmotivation=pct(kauf_sum, kauf_max),
    )

    # Emotional Stability Index = 100 - Durchschnitt der 4 Belastungsscores
    belastung_avg = (
        scores.emotional_activation
        + scores.nervous_system_load
        + scores.selbstverlust
        + scores.hoffnungsschleifen
    ) / 4.0
    scores.emotional_stability_index = round(100.0 - belastung_avg, 1)

    # Runden für hübschere Ausgabe
    scores.emotional_activation = round(scores.emotional_activation, 1)
    scores.nervous_system_load = round(scores.nervous_system_load, 1)
    scores.selbstverlust = round(scores.selbstverlust, 1)
    scores.hoffnungsschleifen = round(scores.hoffnungsschleifen, 1)
    scores.kaufmotivation = round(scores.kaufmotivation, 1)

    # Typen sortieren
    typ_pct = {k: round(pct(typ_sum[k], typ_max[k]), 1) for k in TYP_KEYS}
    sortiert = sorted(typ_pct.items(), key=lambda x: -x[1])

    haupttyp_name, haupttyp_score = sortiert[0]
    nebentyp_name: str | None = None
    nebentyp_score = 0.0
    if len(sortiert) > 1:
        n_name, n_score = sortiert[1]
        if haupttyp_score > 0 and (n_score / haupttyp_score) >= nebentyp_threshold:
            nebentyp_name = n_name
            nebentyp_score = n_score

    typprofil = TypProfil(
        haupttyp=haupttyp_name,
        haupttyp_score=haupttyp_score,
        nebentyp=nebentyp_name,
        nebentyp_score=nebentyp_score,
        alle_typen=typ_pct,
    )

    meta = {
        "fragen_beantwortet": fragen_beantwortet,
        "fragen_uebersprungen": fragen_uebersprungen,
        "fragen_gesamt": len(fragen_data["fragen"]),
    }

    return Auswertung(
        scores=scores,
        typen=typprofil,
        rohantworten=antworten,
        meta=meta,
    )


def format_score_report(auswertung: Auswertung, typen_path: str | Path | None = None) -> str:
    """Hübsche Konsolen-Ausgabe — für Debugging und Test-Runs."""
    typen_namen = {k: k.replace("_", " ").title() for k in TYP_KEYS}
    # Lade Typen-Namen aus Netlify Blob Store (oder lokaler Datei falls Pfad gegeben)
    if typen_path:
        try:
            with open(typen_path) as f:
                typen_data = json.load(f)
            for k, v in typen_data["typen"].items():
                typen_namen[k] = v["name"]
        except Exception:
            pass
    else:
        try:
            from config_loader import lade_typen_daten
            for k, v in lade_typen_daten().items():
                typen_namen[k] = v["name"]
        except Exception:
            pass

    s = auswertung.scores
    t = auswertung.typen
    m = auswertung.meta

    lines = [
        "═" * 60,
        "  EMOTIONAL STABILITY ANALYSE — SCORING",
        "═" * 60,
        "",
        f"  Beantwortet: {m['fragen_beantwortet']}/{m['fragen_gesamt']} Fragen",
        "",
        "  ─── EMOTIONAL STABILITY INDEX™ ───",
        f"    {s.emotional_stability_index:>5.1f} / 100",
        "",
        "  ─── TIEFENSCORES (Belastung) ───",
        f"    Emotionaler Aktivierungsgrad:  {s.emotional_activation:>5.1f}%",
        f"    Nervensystem-Belastung:        {s.nervous_system_load:>5.1f}%",
        f"    Selbstverlust-Level:           {s.selbstverlust:>5.1f}%",
        f"    Hoffnungsschleifen-Level:      {s.hoffnungsschleifen:>5.1f}%",
        "",
        "  ─── HAUPTMUSTER ───",
        f"    {typen_namen.get(t.haupttyp, t.haupttyp)}  ({t.haupttyp_score:.1f}%)",
    ]

    if t.nebentyp:
        lines += [
            "",
            "  ─── ZWEITMUSTER ───",
            f"    {typen_namen.get(t.nebentyp, t.nebentyp)}  ({t.nebentyp_score:.1f}%)",
        ]
    else:
        lines += [
            "",
            "  (Kein signifikantes Zweitmuster — Hauptmuster dominiert klar.)",
        ]

    lines += [
        "",
        "  ─── ALLE TYPEN ───",
    ]
    for typ, score in sorted(t.alle_typen.items(), key=lambda x: -x[1]):
        marker = "←" if typ == t.haupttyp else ("·" if typ == t.nebentyp else " ")
        lines.append(f"    {marker} {typen_namen.get(typ, typ):<45} {score:>5.1f}%")

    lines += [
        "",
        f"  Kaufmotivation: {s.kaufmotivation:.1f}%",
        "═" * 60,
    ]

    return "\n".join(lines)


if __name__ == "__main__":
    # Self-Test mit einem fiktiven Profil
    import sys

    ROOT = Path(__file__).resolve().parents[1]
    fragen_path = ROOT / "fragen.json"

    if len(sys.argv) > 1:
        antworten_path = Path(sys.argv[1])
        with open(antworten_path) as f:
            antworten = json.load(f)["antworten"]
        auswertung = score_antworten(fragen_path, antworten)
        print(format_score_report(auswertung, ROOT / "typen.json"))
    else:
        print("Usage: python scoring.py <antworten.json>")
        print("       (mit Datei aus test_data/)")
