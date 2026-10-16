#!/usr/bin/env python3
"""
KSC Arbeitsplan – Kernmodul (Version 7.0, Oktober 2026)
========================================================
Scheduler, Validierung und Excel-Export für den wöchentlichen Arbeitsplan
des KSC-Teams. Dieses Modul importiert KEIN Streamlit und hat KEINE
Nebenwirkungen (keine State-Datei): derselbe Montag + dieselbe Konfiguration
ergeben immer denselben Plan.

Alle Stammdaten (Personen, Pensen, Anwesenheiten, Rollen, Fixeinsätze,
Rotationen, Besetzungszahlen) kommen aus `team.json` – nichts davon ist im
Code fest verdrahtet. Personen kommen/gehen also ohne Code-Änderung.

Regelstand (Rückmeldungen des Teams, Oktober 2026):
- TEL und ABKL nur halbtags pro Person (VM TEL + NM ABKL ist ok, nie 2x dasselbe)
- ERF9 immer ganztags (Mo Brigitte, Di Dragi, Mi–Fr wechselnd)
- Tagesverantwortung: dieselbe TL VM+NM, an diesem Tag ohne Ämtli → "TL/Q"
- PHC-/18-Uhr-Liste: Ausschlüsse pro Person und pro Tag; PHC-Person scannt nicht
- Postöffnung: Aushilfe täglich PO/SCAN (VM) + scanning (NM); zusätzlich
  rotierende PO-Personen am Vormittag; niemand jeden Tag
- "PO/TEL" am Nachmittag wird nie automatisch vergeben (nur Notfall, manuell)
- Scanning nur halbtags (Ausnahme Aushilfe), max. N Halbtage pro Woche
- TAGES PA: niemand fix, 1 Person ganztags pro Tag, wechselnd
- Schalter: 1 Person ganztags pro Tag (ERF7/SCH), 2 Gruppen im Wochenwechsel
- HUB: 1 Person VM + 1 andere Person NM (ERF7/HUB)
- ERF8: 1 Person pro Tag, Nachmittag (läuft bis 17:00)
- Wochenaufgaben Direkt/ONB/BTM: je 1 Person, nie dieselbe doppelt,
  ganzwöchig Abwesende werden übersprungen
- BackUp Team (Stephi, SPP mit "ERP", KSC Spezial, KSV): nur manuell befüllt
"""

from __future__ import annotations

import argparse
import copy
import json
import random
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VERSION = "7.0"

# ============================================================
# KONSTANTEN
# ============================================================

DAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
DAY_KEYS = ["Mo", "Di", "Mi", "Do", "Fr"]
SLOTS = ["Vormittag", "Nachmittag"]
SLOT_KEYS = ["VM", "NM"]
HALBTAG_WERTE = ["ganz", "VM", "NM", "frei"]
PARITAET = ["gerade", "ungerade"]

# Rollen, die eine Person in team.json unter "kann" haben darf
ROLLEN: Dict[str, str] = {
    "HUB":      "ERF7/HUB (halbtags)",
    "RX Abo":   "Rezept-Abo",
    "VBZ":      "Vorbezüge (nur Info, nicht geplant)",
    "Schalter": "Schalter ganztags, 2 Gruppen im Wochenwechsel",
    "Schalter-Reserve": "Schalter nur, wenn aus der Gruppe niemand ganztags frei ist (seltener)",
    "Scanning": "scanning / PO/SCAN",
    "PO":       "Postöffnung am Vormittag",
    "TPA":      "TAGES PA (ganztags, wechselnd)",
    "PHC":      "18-Uhr-/PHC-Liste",
    "ERF8":     "ERF8 (Nachmittag)",
    "ERF9":     "ERF9-Rotation Mi–Fr (ganztags)",
    "ONB":      "Wochenaufgabe ONB",
    "Direkt":   "Wochenaufgabe Direktbestellung",
    "BTM":      "Wochenaufgabe BTM",
}

DEFAULT_REGELN = {
    "tel_montag_vm": 4,
    "tel_sonst": 3,
    "abkl_pro_halbtag": 2,
    "po_zusatz_vm": 2,
    "hub_pro_tag": 2,
    "rx_abo_pro_woche": 2,
    "scan_max_halbtage_pro_woche": 2,
    "scan_anteil": 0.35,
    "tl_tel_max_pro_woche": 2,
    "tl_tel_min_pro_woche": 1,
    "scan_nm_ersatz_wenn_aushilfe_fehlt": True,
    "schalter_code": "ERF7/SCH",
    "schalter_reserve_modus": "vor der anderen Gruppe",
    "tpa_tel_nm": "bei Bedarf",
}
SCHALTER_CODES = ["ERF7/SCH", "ERF4/SCH"]
SCHALTER_RESERVE_MODI = ["vor der anderen Gruppe", "nach der anderen Gruppe", "nie"]
TPA_TEL_MODI = ["bei Bedarf", "immer", "nie"]
# Codes, die der Generator nie automatisch vergibt (nur im Plan-Editor)
NUR_MANUELL = ["PO/ABKL", "PO/TEL", "ERF8/Q", "ERF8/HUB", "ERF9/Q", "VBZ/Q", "KSC Spez.", "KTG", "Access"]

DEFAULT_ROTATION = {
    "tv_reihenfolge": [],
    "tv_farben": {},
    "phc_montag_tl": [],
    "phc_fix": {},
    "phc_freitag_wechsel": [],
    "phc_tag_ausschluss": {},
    "labor_mi_nm": [],
    "po_dienstag_bevorzugt": "",
    "offset": {},
}
DEFAULT_OFFSETS = ["tv", "phc_montag", "phc_freitag", "schalter", "labor",
                   "Direkt", "ONB", "BTM"]

ABSENT_TASKS = ("KRANK", "FERIEN", "Krankheit")

# Farben für Aufgaben (Excel-Zellen und Web-Vorschau) – nach Referenzvorlage
TASK_COLORS = {
    "KTG":          "A9D18E",  # Kein Tagesgeschäft
    "KGS":          "A9D18E",
    "ABKL":         "FFC000",  # Abklärung
    "PO/ABKL":      "FFC000",
    "TEL":          "92D050",  # Telefon
    "PO/TEL":       "92D050",
    "TEL-16:30":    "92D050",
    "TEL 16:00":    "92D050",
    "ERF7":         "00B0F0",  # Erfassung
    "ERF7/Q":       "00B0F0",
    "ERF7/HUB":     "00B0F0",
    "ERF8":         "FF6600",  # Erfassung bis 17:00
    "ERF8/Q":       "FF6600",
    "ERF8/HUB":     "FF6600",
    "ERF9":         "FF0000",  # Auftragserfassung
    "ERF9/Q":       "FF0000",
    "ERF9/TEL":     "FF0000",
    "ERF7/SCH":     "C6EFCE",  # Schalter
    "ERF4/SCH":     "C6EFCE",
    "ERF4/Access":  "C6EFCE",
    "ERF5":         "F4B084",  # Labor
    "PO":           "D6DCE4",  # Postöffnung
    "PO/SCAN":      "A9D18E",  # Postöffnung & Scanning
    "scanning":     "E2EFDA",  # Scanning / Ablage
    "SCAN":         "E2EFDA",
    "Scan":         "E2EFDA",
    "HO":           "BDD7EE",  # Home Office
    "HO/Q":         "BDD7EE",
    "TL/Q":         "D9E2F3",  # TL mit Tagesverantwortung + Queue
    "RX Abo":       "D9B8FF",  # Rezept-Abo
    "VBZ":          "4FD1C5",  # Vorbezüge
    "VBZ/Q":        "4FD1C5",
    "KSC Spez.":    "FFFF00",
    "TAGES PA":     "00FFFF",
    "TAGES PA/TEL": "00FFFF",
    "TAGESPA":      "00FFFF",
    "KSV":          "00FFFF",
    "Access":       "FF00FF",
    "ERP":          "C6EFCE",  # SPP-Team: ERP Testen
    "Krankheit":    "FF9999",
    "KRANK":        "FF9999",
    "Feiertag":     "FF0000",
    "FERIEN":       "F2F2F2",
    "":             "FFFFFF",
}

# PHC-Liste = dunkelblau hinterlegt (Zellfarbe)
PHC_COLOR = "00008B"

# Fallback-Farben für die Tagesverantwortung, falls team.json keine liefert
TV_FALLBACK_COLORS = ["00B050", "FF6600", "FFFF00", "FF00FF", "7030A0", "00B0F0"]

# Legende im Excel: (Code, Erklärung) – Farben kommen aus TASK_COLORS
LEGEND_ROWS = [
    [("KTG", "Kein Tagesgeschäft"), ("ABKL", "Abklärung"), ("TEL", "Telefon"),
     ("ERF7", "Erfassung"), ("ERF8", "Erfassung bis 17:00"),
     ("ERF9", "Auftragserfassung (ganztags)"), ("PO", "Postöffnung"),
     ("PHC-Liste", "18-Uhr-Dienst"), ("TL/Q", "TL Tagesverantw. + Queue"),
     ("TAGES PA", "Tagespharma (ganztags)")],
    [("HO", "Home Office"), ("Krankheit", "Krankheit"), ("FERIEN", "Ferien"),
     ("scanning", "Scanning / Ablage"), ("PO/SCAN", "Postöffnung & Scanning"),
     ("ERF7/SCH", "Schalter (ganztags)"), ("ERF5", "Labor"),
     ("ERF7/HUB", "HUB"), ("RX Abo", "Rezept-Abo"), ("Feiertag", "Feiertag")],
    [("KSC Spez.", "KSC Spezial"), ("KSV", "KSV Aushilfe"), ("ERP", "ERP Testen"),
     ("KGS", "Kein Tagesgeschäft (Dipiga)"), ("Access", "Access"),
     ("VBZ/Q", "Vorbezüge + Queue"), ("HO/Q", "Home Office + Queue"),
     ("ERF9/TEL", "ERF9 + Telefon (NM)"), ("PO/TEL", "PÖ + Telefon (nur Notfall)"),
     ("", "")],
]

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "team.json"


# ============================================================
# KONFIGURATION (team.json)
# ============================================================

def load_config(path: Optional[Path | str] = None) -> dict:
    """Lädt team.json und ergänzt fehlende Felder mit Standardwerten."""
    p = Path(path) if path else DEFAULT_CONFIG_PATH
    if not p.exists():
        raise FileNotFoundError(
            f"Konfiguration {p} nicht gefunden. team.json muss neben arbeitskalender.py liegen."
        )
    with open(p, "r", encoding="utf-8") as fh:
        return normalize_config(json.load(fh))


def save_config(cfg: dict, path: Optional[Path | str] = None) -> Path:
    p = Path(path) if path else DEFAULT_CONFIG_PATH
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(config_to_json(cfg))
    return p


def config_to_json(cfg: dict) -> str:
    return json.dumps(cfg, ensure_ascii=False, indent=2) + "\n"


def normalize_config(cfg: dict) -> dict:
    """Füllt Standardwerte auf, damit alter/handgeschriebener JSON-Inhalt läuft."""
    cfg = copy.deepcopy(cfg or {})
    cfg.setdefault("version", VERSION)
    regeln = cfg.setdefault("regeln", {})
    for k, v in DEFAULT_REGELN.items():
        regeln.setdefault(k, v)
    rot = cfg.setdefault("rotation", {})
    for k, v in DEFAULT_ROTATION.items():
        rot.setdefault(k, copy.deepcopy(v))
    off = rot.setdefault("offset", {})
    for k in DEFAULT_OFFSETS:
        off.setdefault(k, 0)
    cfg.setdefault("backup_team", [])
    for b in cfg["backup_team"]:
        b.setdefault("name", b.get("short", ""))
        b.setdefault("standard", "")
    for m in cfg.setdefault("mitarbeiter", []):
        m.setdefault("name", m.get("short", ""))
        m.setdefault("pct", 100)
        m.setdefault("aktiv", True)
        m.setdefault("tl", False)
        m.setdefault("nur_tel", False)
        m.setdefault("aushilfe", False)
        m.setdefault("kann", [])
        m.setdefault("anwesend", {})
        m.setdefault("frei_alle_2_wochen", None)
        m.setdefault("fix", [])
        m.setdefault("bemerkung", "")
    return cfg


def validate_config(cfg: dict) -> List[str]:
    """Plausibilitätsprüfung der Konfiguration (für die UI). Leere Liste = ok."""
    return config_errors(cfg) + config_hints(cfg)


def config_errors(cfg: dict) -> List[str]:
    """Fehler, mit denen die Konfiguration nicht übernommen werden darf."""
    probs: List[str] = []
    shorts = [m.get("short", "") for m in cfg.get("mitarbeiter", [])]
    for s in shorts:
        if not s:
            probs.append("Person ohne Kurzname")
    dups = {s for s in shorts if shorts.count(s) > 1 and s}
    if dups:
        probs.append(f"Kurzname doppelt: {', '.join(sorted(dups))}")
    backup_shorts = [b.get("short", "") for b in cfg.get("backup_team", [])]
    overlap = set(shorts) & set(backup_shorts)
    if overlap:
        probs.append(f"Kurzname im Team UND im BackUp-Team: {', '.join(sorted(overlap))}")
    names = [(m.get("name") or m.get("short", "")).strip() for m in cfg.get("mitarbeiter", [])] + \
            [(b.get("name") or b.get("short", "")).strip() for b in cfg.get("backup_team", [])]
    dup_names = {n for n in names if n and names.count(n) > 1}
    if dup_names:
        probs.append(f"Name doppelt: {', '.join(sorted(dup_names))}")
    for m in cfg.get("mitarbeiter", []):
        who = m.get("name") or m.get("short")
        for r in m.get("kann", []):
            if r not in ROLLEN:
                probs.append(f"{who}: unbekannte Rolle '{r}'")
        for dk, val in m.get("anwesend", {}).items():
            if dk not in DAY_KEYS or val not in HALBTAG_WERTE:
                probs.append(f"{who}: Anwesenheit {dk}={val} ungültig")
        f2 = m.get("frei_alle_2_wochen")
        if f2:
            if f2.get("kw") not in PARITAET:
                probs.append(f"{who}: frei_alle_2_wochen.kw muss gerade/ungerade sein")
            for dk, val in f2.get("tage", {}).items():
                if dk not in DAY_KEYS or val not in HALBTAG_WERTE:
                    probs.append(f"{who}: frei_alle_2_wochen {dk}={val} ungültig")
        for fx in m.get("fix", []):
            if fx.get("tag") not in DAY_KEYS or fx.get("halbtag") not in HALBTAG_WERTE:
                probs.append(f"{who}: Fixeinsatz {fx} ungültig")
            if not fx.get("task"):
                probs.append(f"{who}: Fixeinsatz ohne Task")
            if fx.get("kw") not in (None, "", "gerade", "ungerade"):
                probs.append(f"{who}: Fixeinsatz kw muss leer/gerade/ungerade sein")
    for dk in cfg.get("rotation", {}).get("phc_fix", {}):
        if dk not in DAY_KEYS:
            probs.append(f"phc_fix: Tag '{dk}' ungültig")
    return probs


def config_hints(cfg: dict) -> List[str]:
    """Hinweise: der Generator läuft, überspringt aber z.B. inaktive Personen in Rotationen."""
    probs: List[str] = []
    members = {m["short"]: m for m in cfg.get("mitarbeiter", []) if m.get("short")}
    active = {s for s, m in members.items() if m.get("aktiv", True)}

    def why(n):
        return "ist inaktiv und wird übersprungen" if n in members else "ist unbekannt"

    rot = cfg.get("rotation", {})
    if not any(m.get("tl") for s, m in members.items() if s in active):
        probs.append("Keine aktive Teamleitung (tl=true) – Tagesverantwortung/Queue bleiben offen")
    for key, label in (("tv_reihenfolge", "Tagesverantwortung"), ("phc_montag_tl", "PHC Montag"),
                       ("phc_freitag_wechsel", "PHC Freitag"), ("labor_mi_nm", "Labor Mi NM")):
        for n in rot.get(key, []):
            if n not in active:
                probs.append(f"{label}: '{n}' {why(n)}")
    for dk, n in rot.get("phc_fix", {}).items():
        if n not in active:
            probs.append(f"PHC {dk} fix: '{n}' {why(n)}")
    return probs


# ── Team bearbeiten (für die App; ohne Streamlit testbar) ─────────────

FUNKTIONEN = ["Team", "Teamleitung", "nur Telefon", "Aushilfe"]
ROTATION_LISTEN = ["tv_reihenfolge", "phc_montag_tl", "phc_freitag_wechsel", "labor_mi_nm"]


def funktion(m: dict) -> str:
    if m.get("tl"):
        return "Teamleitung"
    if m.get("nur_tel"):
        return "nur Telefon"
    if m.get("aushilfe"):
        return "Aushilfe"
    return "Team"


def set_funktion(cfg: dict, short: str, value: str) -> None:
    """Setzt TL / nur TEL / Aushilfe. Neue TL kommen ans Ende der TV- und
    PHC-Montag-Rotation, ehemalige TL werden dort entfernt."""
    m = next(x for x in cfg["mitarbeiter"] if x["short"] == short)
    was_tl = bool(m.get("tl"))
    m["tl"], m["nur_tel"], m["aushilfe"] = value == "Teamleitung", value == "nur Telefon", value == "Aushilfe"
    rot = cfg["rotation"]
    for key in ("tv_reihenfolge", "phc_montag_tl"):
        lst = rot.get(key) or []
        if m["tl"] and not was_tl and lst and short not in lst:
            rot[key] = lst + [short]
        elif was_tl and not m["tl"]:
            rot[key] = [n for n in lst if n != short]


def new_short(name: str, cfg: dict) -> str:
    """Interner Schlüssel für eine neue Person: Vorname, bei Kollision der ganze Name."""
    name = name.strip()
    taken = {m["short"] for m in cfg.get("mitarbeiter", [])} | {b["short"] for b in cfg.get("backup_team", [])}
    first = name.split()[0] if name.split() else name
    for cand in (first, name):
        if cand and cand not in taken:
            return cand
    i = 2
    while f"{name} {i}" in taken:
        i += 1
    return f"{name} {i}"


def rotation_refs(cfg: dict, short: str) -> List[str]:
    """Wo steht die Person in Rotationen / Tagesregeln? (für Hinweise beim Entfernen)"""
    rot = cfg["rotation"]
    labels = {"tv_reihenfolge": "Tagesverantwortung", "phc_montag_tl": "18-Uhr-Liste Montag",
              "phc_freitag_wechsel": "18-Uhr-Liste Freitag", "labor_mi_nm": "Labor Mi NM"}
    refs = [lbl for key, lbl in labels.items() if short in rot.get(key, [])]
    refs += [f"18-Uhr-Liste fix {dk}" for dk, n in rot.get("phc_fix", {}).items() if n == short]
    if any(short in v for v in rot.get("phc_tag_ausschluss", {}).values()):
        refs.append("18-Uhr-Liste Ausschlüsse")
    if rot.get("po_dienstag_bevorzugt") == short:
        refs.append("Postöffnung Dienstag")
    return refs


def remove_member(cfg: dict, short: str) -> None:
    """Entfernt eine Person samt allen Verweisen in Rotationen."""
    cfg["mitarbeiter"] = [m for m in cfg["mitarbeiter"] if m["short"] != short]
    rot = cfg["rotation"]
    for key in ROTATION_LISTEN:
        rot[key] = [n for n in rot.get(key, []) if n != short]
    rot.get("tv_farben", {}).pop(short, None)
    rot["phc_fix"] = {dk: n for dk, n in rot.get("phc_fix", {}).items() if n != short}
    rot["phc_tag_ausschluss"] = {dk: [n for n in v if n != short]
                                 for dk, v in rot.get("phc_tag_ausschluss", {}).items()
                                 if [n for n in v if n != short]}
    if rot.get("po_dienstag_bevorzugt") == short:
        rot["po_dienstag_bevorzugt"] = ""


def recurring_rows(cfg: dict) -> List[dict]:
    """Fixeinsätze und «alle 2 Wochen frei» aller Personen als eine Liste:
    {short, tag, halbtag, task, kw}  (task "frei" = alle 2 Wochen frei)."""
    rows = []
    for m in cfg["mitarbeiter"]:
        for fx in m.get("fix", []):
            rows.append({"short": m["short"], "tag": fx.get("tag", "Mo"), "halbtag": fx.get("halbtag", "ganz"),
                         "task": fx.get("task", ""), "kw": fx.get("kw") or None})
        f2 = m.get("frei_alle_2_wochen") or {}
        for dk, val in f2.get("tage", {}).items():
            rows.append({"short": m["short"], "tag": dk, "halbtag": val, "task": "frei", "kw": f2.get("kw")})
    return rows


def apply_recurring_rows(cfg: dict, rows: List[dict]) -> List[str]:
    """Schreibt die Liste aus recurring_rows() zurück. Liefert Fehler (dann bleibt cfg unverändert)."""
    names = {m["short"]: m.get("name") or m["short"] for m in cfg["mitarbeiter"]}
    per = {s: {"fix": [], "frei": {}, "kw": None} for s in names}
    errors = []
    for r in rows:
        p = per.get(r["short"])
        if p is None or not r.get("task"):
            continue
        if r["task"] == "frei":
            if r.get("kw") not in PARITAET:
                errors.append(f"{names[r['short']]}: «frei» jede Woche bitte unter Personen → Anwesenheit "
                              f"({r['tag']}) eintragen; hier nur «gerade KW» oder «ungerade KW»")
            elif p["kw"] and p["kw"] != r["kw"]:
                errors.append(f"{names[r['short']]}: «frei» geht nur in geraden ODER in ungeraden KW, nicht gemischt")
            else:
                p["kw"] = r["kw"]
                p["frei"][r["tag"]] = r["halbtag"]
        else:
            fx = {"tag": r["tag"], "halbtag": r["halbtag"], "task": r["task"]}
            if r.get("kw") in PARITAET:
                fx["kw"] = r["kw"]
            p["fix"].append(fx)
    if errors:
        return errors
    for m in cfg["mitarbeiter"]:
        p = per[m["short"]]
        m["fix"] = p["fix"]
        m["frei_alle_2_wochen"] = {"kw": p["kw"], "tage": p["frei"]} if p["frei"] else None
    return []


REGEL_LABELS = {
    "tel_montag_vm": "TEL Montag Vormittag", "tel_sonst": "TEL übrige Halbtage",
    "abkl_pro_halbtag": "ABKL pro Halbtag", "po_zusatz_vm": "Zusatz-PO am Vormittag",
    "hub_pro_tag": "HUB-Halbtage pro Tag", "rx_abo_pro_woche": "RX Abo pro Woche",
    "scan_max_halbtage_pro_woche": "Scanning max. pro Person/Woche", "scan_anteil": "Scanning-Anteil",
    "tl_tel_min_pro_woche": "TL am Telefon mind.", "tl_tel_max_pro_woche": "TL am Telefon max.",
    "scan_nm_ersatz_wenn_aushilfe_fehlt": "NM-Scan-Ersatz ohne Aushilfe",
    "schalter_code": "Schalter-Code", "schalter_reserve_modus": "Schalter-Reserve",
    "tpa_tel_nm": "TAGES PA/TEL am Nachmittag",
}
ROTATION_LABELS = {
    "tv_reihenfolge": "Tagesverantwortung", "phc_montag_tl": "18-Uhr-Liste Montag",
    "phc_freitag_wechsel": "18-Uhr-Liste Freitag", "labor_mi_nm": "Labor Mi NM",
    "phc_fix": "18-Uhr-Liste fix pro Tag", "phc_tag_ausschluss": "18-Uhr-Liste Ausschlüsse",
    "po_dienstag_bevorzugt": "Postöffnung Dienstag", "tv_farben": "TV-Farben",
}
OFFSET_LABELS = {
    "tv": "Tagesverantwortung", "phc_montag": "18-Uhr-Liste Montag", "phc_freitag": "18-Uhr-Liste Freitag",
    "schalter": "Schalter-Gruppe", "labor": "Labor", "Direkt": "Direkt", "ONB": "ONB", "BTM": "BTM",
}


def describe_changes(old: dict, new: dict) -> List[str]:
    """Lesbare Liste der Unterschiede zwischen zwei Konfigurationen (für «ungespeichert»)."""
    old, new = normalize_config(old), normalize_config(new)
    out: List[str] = []
    om = {m["short"]: m for m in old["mitarbeiter"]}
    nm = {m["short"]: m for m in new["mitarbeiter"]}

    def name(s):
        return (nm.get(s) or om.get(s) or {}).get("name") or s

    def fx_str(fx):
        kw = f" ({fx['kw']} KW)" if fx.get("kw") else ""
        return f"{fx.get('tag')} {fx.get('halbtag', 'ganz')} {fx.get('task')}{kw}"

    out += [f"Neue Person: {name(s)}" for s in nm if s not in om]
    out += [f"Entfernt: {name(s)}" for s in om if s not in nm]
    for s, m in nm.items():
        o = om.get(s)
        if o is None:
            continue
        n = m.get("name") or s
        if o.get("name") != m.get("name"):
            out.append(f"{o.get('name')} heisst neu {m.get('name')}")
        if o.get("pct") != m.get("pct"):
            out.append(f"{n}: Pensum {o.get('pct')} → {m.get('pct')} %")
        if bool(o.get("aktiv", True)) != bool(m.get("aktiv", True)):
            out.append(f"{n}: {'wieder aktiv' if m.get('aktiv', True) else 'inaktiv (wird nicht eingeplant)'}")
        if funktion(o) != funktion(m):
            out.append(f"{n}: Funktion {funktion(o)} → {funktion(m)}")
        for dk in DAY_KEYS:
            a, b = o["anwesend"].get(dk, "ganz"), m["anwesend"].get(dk, "ganz")
            if a != b:
                out.append(f"{n}: Anwesenheit {dk} {a} → {b}")
        plus = [r for r in m["kann"] if r not in o["kann"]]
        minus = [r for r in o["kann"] if r not in m["kann"]]
        if plus:
            out.append(f"{n}: darf neu {', '.join(plus)}")
        if minus:
            out.append(f"{n}: darf nicht mehr {', '.join(minus)}")
        ofx, nfx = sorted(map(fx_str, o["fix"])), sorted(map(fx_str, m["fix"]))
        out += [f"{n}: fester Einsatz {x}" for x in nfx if x not in ofx]
        out += [f"{n}: fester Einsatz entfernt: {x}" for x in ofx if x not in nfx]
        if (o.get("frei_alle_2_wochen") or None) != (m.get("frei_alle_2_wochen") or None):
            f2 = m.get("frei_alle_2_wochen")
            out.append(f"{n}: alle 2 Wochen frei → " +
                       (f"{', '.join(f'{d} {v}' for d, v in f2['tage'].items())} ({f2['kw']} KW)" if f2 else "keine"))
        if (o.get("bemerkung") or "") != (m.get("bemerkung") or ""):
            out.append(f"{n}: Bemerkung geändert")
    common_old = [s for s in om if s in nm]
    common_new = [s for s in nm if s in om]
    if common_old != common_new:
        out.append("Reihenfolge im Plan geändert")
    for k, lbl in REGEL_LABELS.items():
        if old["regeln"].get(k) != new["regeln"].get(k):
            out.append(f"Regel «{lbl}»: {old['regeln'].get(k)} → {new['regeln'].get(k)}")
    for k, lbl in ROTATION_LABELS.items():
        if old["rotation"].get(k) != new["rotation"].get(k):
            out.append(f"{lbl} geändert")
    for k, lbl in OFFSET_LABELS.items():
        if old["rotation"]["offset"].get(k) != new["rotation"]["offset"].get(k):
            out.append(f"Rotation «{lbl}»: wer diese Woche dran ist, neu eingestellt")
    if old["backup_team"] != new["backup_team"]:
        out.append("BackUp-Team geändert")
    if not out and config_to_json(old) != config_to_json(new):
        out.append("Weitere Einstellungen geändert")
    return out


# ============================================================
# KALENDER-HILFEN
# ============================================================

def get_next_monday(ref: Optional[datetime] = None) -> datetime:
    if ref is None:
        ref = datetime.now()
    days = 0 - ref.weekday()
    if days <= 0:
        days += 7
    return ref + timedelta(days=days)


def week_index(monday) -> int:
    """Fortlaufende Wochennummer seit Mo 1.1.2024 (ISO-KW 1/2024). Läuft über
    Jahresgrenzen weiter – Grundlage für alle Rotationen (TV, PHC, Schalter …)."""
    if isinstance(monday, datetime):
        monday = monday.date()
    return (monday - date(2024, 1, 1)).days // 7


def kw_paritaet(week_number: int) -> str:
    return "gerade" if week_number % 2 == 0 else "ungerade"


def halbtag_slots(value: str) -> List[int]:
    return {"ganz": [0, 1], "VM": [0], "NM": [1], "frei": []}.get(value, [0, 1])


def is_scan_task(task: str) -> bool:
    return "scan" in (task or "").lower()


# ============================================================
# MITARBEITER
# ============================================================

@dataclass
class Employee:
    short: str
    name: str
    pct: int
    is_tl: bool = False
    nur_tel: bool = False
    aushilfe: bool = False
    kann: set = field(default_factory=set)
    available: Dict[Tuple[int, int], bool] = field(default_factory=dict)
    fix: List[Tuple[int, int, str]] = field(default_factory=list)
    bemerkung: str = ""

    def is_available(self, day: int, slot: int) -> bool:
        return self.available.get((day, slot), False)

    def can(self, role: str) -> bool:
        return role in self.kann

    @property
    def regulaer(self) -> bool:
        """Normales Teammitglied (keine Aushilfe, nicht nur TEL)."""
        return not self.nur_tel and not self.aushilfe


def create_employees(week_number: int, cfg: Optional[dict] = None) -> Dict[str, Employee]:
    """Erzeugt alle aktiven Personen für diese KW aus der Konfiguration."""
    cfg = normalize_config(cfg if cfg is not None else load_config())
    par = kw_paritaet(week_number)
    employees: Dict[str, Employee] = {}
    for m in cfg["mitarbeiter"]:
        if not m.get("aktiv", True):
            continue
        e = Employee(
            short=m["short"], name=m.get("name") or m["short"], pct=int(m.get("pct", 100)),
            is_tl=bool(m.get("tl")), nur_tel=bool(m.get("nur_tel")),
            aushilfe=bool(m.get("aushilfe")), kann=set(m.get("kann", [])),
            bemerkung=m.get("bemerkung", ""),
        )
        for d, dk in enumerate(DAY_KEYS):
            slots = halbtag_slots(m.get("anwesend", {}).get(dk, "ganz"))
            for s in range(2):
                e.available[(d, s)] = s in slots
        f2 = m.get("frei_alle_2_wochen")
        if f2 and f2.get("kw") == par:
            for dk, val in f2.get("tage", {}).items():
                if dk in DAY_KEYS:
                    for s in halbtag_slots(val):
                        e.available[(DAY_KEYS.index(dk), s)] = False
        for fx in m.get("fix", []):
            if fx.get("kw") in ("gerade", "ungerade") and fx["kw"] != par:
                continue
            if fx.get("tag") not in DAY_KEYS:
                continue
            d = DAY_KEYS.index(fx["tag"])
            for s in halbtag_slots(fx.get("halbtag", "ganz")):
                e.fix.append((d, s, fx.get("task", "")))
        employees[e.short] = e
    return employees


def two_week_notes(week_number: int, cfg: dict) -> List[str]:
    """Für die UI: wer hat in dieser KW seinen 2-Wochen-Freitag/-Mittwoch."""
    par = kw_paritaet(week_number)
    notes = []
    for m in normalize_config(cfg)["mitarbeiter"]:
        if not m.get("aktiv", True):
            continue
        f2 = m.get("frei_alle_2_wochen")
        if f2 and f2.get("kw") == par:
            tage = ", ".join(f"{dk} {val}" if val != "ganz" else dk for dk, val in f2.get("tage", {}).items())
            notes.append(f"{m['short']}: {tage} frei")
        for fx in m.get("fix", []):
            if fx.get("kw") == par:
                notes.append(f"{m['short']}: {fx.get('tag')} {fx.get('halbtag', 'ganz')} {fx.get('task')}")
    return notes


# ============================================================
# SCHEDULE
# ============================================================

class Schedule:
    def __init__(self, employees: Dict[str, Employee], week_number: int, week_start_date, cfg: dict):
        self.employees = employees
        self.week_number = week_number
        self.week_start_date = week_start_date
        self.cfg = cfg
        self.is_even_week = (week_number % 2 == 0)

        self.schedule: Dict[str, Dict[Tuple[int, int], str]] = {n: {} for n in employees}
        self.tel_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        self.abkl_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        self.erf8_assigned = {d: False for d in range(5)}

        self.tagesverantwortung: Dict[int, str] = {}
        self.phc_liste: Dict[int, str] = {}
        self.tagespa: Dict[int, str] = {}
        self.po_liste: Dict[int, str] = {}        # PO/SCAN-Person (Postöffnung des Tages)
        self.po_zusatz: Dict[int, List[str]] = {d: [] for d in range(5)}
        self.schalter: Dict[int, str] = {}
        self.erf9: Dict[int, str] = {}
        self.wochenaufgabe_direkt = ""
        self.wochenaufgabe_onb = ""
        self.wochenaufgabe_btm = ""
        self.notes: Dict[Tuple[str, int, int], str] = {}
        self.backup_tasks: Dict[Tuple[str, int, int], str] = {}
        self.conflicts: List[str] = []

    # --- Zuweisung -------------------------------------------------------
    def _count(self, day, slot, task):
        if "TEL" in task:
            self.tel_count[day][slot] += 1
        if "ABKL" in task:
            self.abkl_count[day][slot] += 1
        if "ERF8" in task:
            self.erf8_assigned[day] = True

    def assign(self, name, day, slot, task) -> bool:
        if name not in self.employees:
            return False
        if (day, slot) in self.schedule[name]:
            return False
        if not self.employees[name].is_available(day, slot):
            return False
        self.schedule[name][(day, slot)] = task
        self._count(day, slot, task)
        return True

    def force_assign(self, name, day, slot, task) -> None:
        if name not in self.employees:
            return
        self.schedule[name][(day, slot)] = task
        self._count(day, slot, task)

    def is_free(self, name, day, slot) -> bool:
        return (name in self.employees
                and (day, slot) not in self.schedule[name]
                and self.employees[name].is_available(day, slot))

    def is_free_day(self, name, day) -> bool:
        return self.is_free(name, day, 0) and self.is_free(name, day, 1)

    def get_task(self, name, day, slot) -> str:
        return self.schedule.get(name, {}).get((day, slot), "")

    def get_backup_task(self, name, day, slot) -> str:
        if (name, day, slot) in self.backup_tasks:
            return self.backup_tasks[(name, day, slot)]
        for b in self.cfg.get("backup_team", []):
            if b.get("short") == name:
                return b.get("standard", "")
        return ""

    def get_available(self, day, slot, exclude: Optional[Iterable[str]] = None) -> List[str]:
        ex = set(exclude or ())
        return [n for n in self.employees if n not in ex and self.is_free(n, day, slot)]

    def is_absent(self, name, day, slot) -> bool:
        return (self.get_task(name, day, slot) in ABSENT_TASKS
                or (name, day, slot) in self.notes)

    def is_present(self, name, day, slot) -> bool:
        return (name in self.employees
                and self.employees[name].is_available(day, slot)
                and not self.is_absent(name, day, slot))

    def is_present_day(self, name, day) -> bool:
        return self.is_present(name, day, 0) and self.is_present(name, day, 1)

    def present_anytime(self, name) -> bool:
        return any(self.is_present(name, d, s) for d in range(5) for s in range(2))

    def count_task(self, name, pred) -> int:
        return sum(1 for t in self.schedule.get(name, {}).values() if pred(t))


# ============================================================
# SCHEDULER
# ============================================================

def build_schedule(week_number: int, week_start_date, overrides: Optional[dict] = None,
                   cfg: Optional[dict] = None, seed: Optional[int] = None) -> Schedule:
    """Baut den Wochenplan.

    overrides: {(short, day, slot): task}  – "KRANK"/"FERIEN", "*Notiz" (Termin,
               Slot bleibt leer), oder ein beliebiger Task-Code (HO, …).
               Einträge für BackUp-Namen landen 1:1 in sched.backup_tasks.
    cfg:       Konfiguration (team.json); None = Datei neben diesem Modul laden.
    seed:      Zufalls-Seed; Standard = ISO-Jahr*100 + ISO-KW (deterministisch).
    """
    cfg = normalize_config(cfg if cfg is not None else load_config())
    employees = create_employees(week_number, cfg)
    sched = Schedule(employees, week_number, week_start_date, cfg)
    rules = cfg["regeln"]
    rot = cfg["rotation"]
    off = rot["offset"]
    order = [m["short"] for m in cfg["mitarbeiter"] if m["short"] in employees]
    backup_names = {b["short"] for b in cfg["backup_team"]}

    iso_year, iso_week, _ = week_start_date.isocalendar()
    rng = random.Random(seed if seed is not None else iso_year * 100 + iso_week)
    widx = week_index(week_start_date)

    def shuffled(items):
        items = list(items)
        rng.shuffle(items)
        return items

    def pick(pool, used_counter=None):
        """Wählt aus pool bevorzugt Personen mit der kleinsten Wochen-Zählung."""
        pool = list(pool)
        if not pool:
            return None
        if used_counter is None:
            return rng.choice(pool)
        best = min(used_counter.get(n, 0) for n in pool)
        return rng.choice([n for n in pool if used_counter.get(n, 0) == best])

    # ════════════════════════════════════════════════════════════
    # 0. OVERRIDES (Krank / Ferien / Termin / Sonderwunsch / BackUp)
    # ════════════════════════════════════════════════════════════
    for (name, day, slot), task in (overrides or {}).items():
        task = task or ""
        if name in employees:
            if task.startswith("*"):
                sched.notes[(name, day, slot)] = task[1:] or "Termin"
                sched.force_assign(name, day, slot, "")
            elif task in ABSENT_TASKS:
                sched.force_assign(name, day, slot, task)
            else:
                sched.assign(name, day, slot, task)
        elif name in backup_names:
            sched.backup_tasks[(name, day, slot)] = task.lstrip("*")

    # ════════════════════════════════════════════════════════════
    # 1. FIXE EINSÄTZE aus team.json (z.B. Brigitte Mo ERF9, Dipiga KGS,
    #    Maria B. Mo VM HO, Florence Mi VM HO, Linda Di HO/Q alle 2 Wochen)
    # ════════════════════════════════════════════════════════════
    for name in order:
        for d, s, task in employees[name].fix:
            if task:
                sched.assign(name, d, s, task)

    # 2. LABOR Mi NM – zwei Personen im Wochenwechsel
    labor = [n for n in rot["labor_mi_nm"] if n in employees]
    if labor:
        start = (widx + off["labor"]) % len(labor)
        for i in range(len(labor)):
            cand = labor[(start + i) % len(labor)]
            if sched.is_free(cand, 2, 1) and sched.assign(cand, 2, 1, "ERF5"):
                if i:
                    sched.conflicts.append(f"ℹ️ Labor Mi NM: {labor[start]} nicht frei → {cand}")
                break
        else:
            sched.conflicts.append("⚠️ Labor Mi NM (ERF5): niemand aus dem Labor-Pool frei")

    # 3. AUSHILFE: täglich PO/SCAN am VM (= Postöffnung des Tages) + scanning NM
    for name in order:
        if not employees[name].aushilfe:
            continue
        for day in range(5):
            if day not in sched.po_liste and sched.assign(name, day, 0, "PO/SCAN"):
                sched.po_liste[day] = name
            sched.assign(name, day, 1, "scanning")

    # 4. NUR-TEL-Personen: jeder anwesende Halbtag = TEL
    for name in order:
        if employees[name].nur_tel:
            for day in range(5):
                for slot in range(2):
                    sched.assign(name, day, slot, "TEL")

    # ════════════════════════════════════════════════════════════
    # 5. TAGESVERANTWORTUNG – dieselbe TL VM+NM, ohne Ämtli → "TL/Q"
    # ════════════════════════════════════════════════════════════
    tv_list = [n for n in rot["tv_reihenfolge"] if n in employees] or \
              [n for n in order if employees[n].is_tl]
    if tv_list:
        tv_start = (widx * 5 + off["tv"]) % len(tv_list)
        tv_count: Dict[str, int] = {n: 0 for n in tv_list}
        for day in range(5):
            idx = (tv_start + day) % len(tv_list)
            planned = tv_list[idx]
            chosen = None
            if sched.is_present_day(planned, day) and sched.is_free_day(planned, day):
                chosen = planned
            else:
                # Vertretung: verfügbare TL mit den wenigsten TV-Tagen diese Woche,
                # bei Gleichstand in Rotationsreihenfolge
                cands = [tv_list[(idx + i) % len(tv_list)] for i in range(1, len(tv_list))]
                cands = [c for c in cands if sched.is_present_day(c, day) and sched.is_free_day(c, day)]
                if cands:
                    chosen = min(cands, key=lambda c: tv_count[c])
            if chosen:
                tv_count[chosen] += 1
                sched.tagesverantwortung[day] = chosen
                sched.assign(chosen, day, 0, "TL/Q")
                sched.assign(chosen, day, 1, "TL/Q")
                if chosen != planned:
                    sched.conflicts.append(f"ℹ️ TV {DAYS[day]}: {planned} nicht verfügbar → {chosen} übernimmt")
            else:
                sched.tagesverantwortung[day] = "⚠️ OFFEN"
                sched.conflicts.append(f"❌ TV {DAYS[day]}: Keine TL verfügbar! Manuell zuweisen.")
    else:
        for day in range(5):
            sched.tagesverantwortung[day] = "⚠️ OFFEN"
        sched.conflicts.append("❌ Tagesverantwortung: keine Teamleitung in team.json (tl=true)")

    # ════════════════════════════════════════════════════════════
    # 6. PHC-/18-UHR-LISTE – vor der Verteilung, damit die PHC-Person
    #    kein Scanning bekommt (Arbeitszeiten)
    # ════════════════════════════════════════════════════════════
    tag_ausschluss = {DAY_KEYS.index(k): set(v) for k, v in rot["phc_tag_ausschluss"].items() if k in DAY_KEYS}

    def phc_ok(name, day):
        emp = employees.get(name)
        return (emp is not None and emp.can("PHC") and emp.regulaer
                and name not in tag_ausschluss.get(day, set())
                and sched.is_present(name, day, 1))

    phc_used: Dict[str, int] = {}

    def phc_set(day, name):
        sched.phc_liste[day] = name
        phc_used[name] = phc_used.get(name, 0) + 1

    def phc_pick(day):
        pool = [n for n in order if phc_ok(n, day)]
        name = pick(pool, phc_used)
        if name:
            phc_set(day, name)
            return name
        sched.phc_liste[day] = "⚠️ OFFEN"
        sched.conflicts.append(f"❌ PHC {DAYS[day]}: Niemand verfügbar!")
        return None

    phc_fix = {DAY_KEYS.index(k): v for k, v in rot["phc_fix"].items() if k in DAY_KEYS}
    phc_mo = [n for n in rot["phc_montag_tl"] if n in employees]
    phc_fr = [n for n in rot["phc_freitag_wechsel"] if n in employees]

    for day in range(5):
        if day == 0 and phc_mo:
            idx = (widx + off["phc_montag"]) % len(phc_mo)
            planned = phc_mo[idx]
            for i in range(len(phc_mo)):
                cand = phc_mo[(idx + i) % len(phc_mo)]
                if phc_ok(cand, 0):
                    phc_set(0, cand)
                    if cand != planned:
                        sched.conflicts.append(f"ℹ️ PHC Mo: {planned} nicht da → {cand}")
                    break
            else:
                repl = phc_pick(0)
                if repl:
                    sched.conflicts.append(f"⚠️ PHC Mo: keine TL verfügbar → {repl}")
        elif day in phc_fix:
            fixname = phc_fix[day]
            if phc_ok(fixname, day):
                phc_set(day, fixname)
            else:
                repl = phc_pick(day)
                if repl:
                    sched.conflicts.append(f"ℹ️ PHC {DAY_KEYS[day]}: {fixname} nicht da → {repl}")
        elif day == 4 and phc_fr:
            idx = (widx + off["phc_freitag"]) % len(phc_fr)
            planned = phc_fr[idx]
            for i in range(len(phc_fr)):
                cand = phc_fr[(idx + i) % len(phc_fr)]
                if phc_ok(cand, 4):
                    phc_set(4, cand)
                    if cand != planned:
                        sched.conflicts.append(f"ℹ️ PHC Fr: {planned} nicht da → {cand}")
                    break
            else:
                repl = phc_pick(4)
                if repl:
                    sched.conflicts.append(f"⚠️ PHC Fr: {' / '.join(phc_fr)} nicht da → {repl}")
        else:
            phc_pick(day)

    # ════════════════════════════════════════════════════════════
    # 6b. TL AM TELEFON – Minimum pro Woche vorab reservieren, bevor
    #     Schalter / ERF9 / RX Abo / ERF8 die TL-Halbtage belegen
    # ════════════════════════════════════════════════════════════
    tl_min = int(rules["tl_tel_min_pro_woche"])
    if tl_min > 0:
        for tl in [n for n in order if employees[n].is_tl]:
            have = sched.count_task(tl, lambda t: "TEL" in t)
            for _ in range(max(0, tl_min - have)):
                cands = [(d, s) for d in range(5) for s in range(2)
                         if sched.is_free(tl, d, s) and sched.is_present(tl, d, s)
                         and "TEL" not in sched.get_task(tl, d, 1 - s)]
                if not cands:
                    break
                d, s = min(cands, key=lambda ds: (sched.tel_count[ds[0]][ds[1]], rng.random()))
                sched.assign(tl, d, s, "TEL")

    # ════════════════════════════════════════════════════════════
    # 7. SCHALTER – 1 Person GANZTAGS pro Tag (ERF7/SCH), 2 Gruppen im
    #    Wochenwechsel, max. 1x pro Person und Woche
    # ════════════════════════════════════════════════════════════
    sch_pool = [n for n in order if employees[n].can("Schalter") and employees[n].regulaer]
    sch_reserve = [n for n in order if employees[n].can("Schalter-Reserve")
                   and not employees[n].can("Schalter") and employees[n].regulaer]
    sch_code = str(rules["schalter_code"] or "ERF7/SCH")
    if sch_pool or sch_reserve:
        mid = (len(sch_pool) + 1) // 2
        group = (widx + off["schalter"]) % 2
        active_grp = sch_pool[:mid] if group == 0 else sch_pool[mid:]
        other_grp = sch_pool[mid:] if group == 0 else sch_pool[:mid]
        modus = str(rules["schalter_reserve_modus"])
        stufen = [(active_grp, "")]
        if modus == "vor der anderen Gruppe":
            stufen += [(sch_reserve, "Reserve"), (other_grp, "Ersatz aus anderer Gruppe")]
        elif modus == "nach der anderen Gruppe":
            stufen += [(other_grp, "Ersatz aus anderer Gruppe"), (sch_reserve, "Reserve")]
        else:
            stufen += [(other_grp, "Ersatz aus anderer Gruppe")]
        sch_used: set = set()
        for day in range(5):
            chosen = None
            for grp, note in stufen:
                cands = [n for n in grp if n not in sch_used and sched.is_free_day(n, day)]
                if cands:
                    chosen = rng.choice(cands)
                    if note:
                        sched.conflicts.append(f"ℹ️ Schalter {DAYS[day]}: {note} ({chosen})")
                    break
            if chosen:
                sched.assign(chosen, day, 0, sch_code)
                sched.assign(chosen, day, 1, sch_code)
                sched.schalter[day] = chosen
                sch_used.add(chosen)
            else:
                sched.schalter[day] = "⚠️ OFFEN"
                sched.conflicts.append(f"⚠️ Schalter {DAYS[day]}: niemand ganztags frei")

    # ════════════════════════════════════════════════════════════
    # 8. TAGES PA – 1 Person GANZTAGS pro Tag, niemand fix, wechselnd
    # ════════════════════════════════════════════════════════════
    tpa_used: Dict[str, int] = {}
    for day in range(5):
        cands = [n for n in order if employees[n].can("TPA") and not employees[n].is_tl
                 and employees[n].regulaer and sched.is_free_day(n, day)]
        name = pick(cands, tpa_used)
        if name:
            sched.assign(name, day, 0, "TAGES PA")
            sched.assign(name, day, 1, "TAGES PA")
            sched.tagespa[day] = name
            tpa_used[name] = tpa_used.get(name, 0) + 1
        else:
            sched.tagespa[day] = "⚠️ OFFEN"
            sched.conflicts.append(f"❌ TAGES PA {DAYS[day]}: Niemand ganztags verfügbar!")

    # ════════════════════════════════════════════════════════════
    # 9. ERF9 Mi–Fr – GANZTAGS, wechselnd (Mo/Di sind Fixeinsätze)
    # ════════════════════════════════════════════════════════════
    erf9_used: Dict[str, int] = {}
    for day in range(5):
        if any("ERF9" in sched.get_task(n, day, 0) for n in employees):
            holder = next(n for n in employees if "ERF9" in sched.get_task(n, day, 0))
            sched.erf9[day] = holder
            continue
        cands = [n for n in order if employees[n].can("ERF9") and employees[n].regulaer
                 and sched.is_free_day(n, day)]
        name = pick(cands, erf9_used)
        if name:
            sched.assign(name, day, 0, "ERF9")
            sched.assign(name, day, 1, "ERF9")
            sched.erf9[day] = name
            erf9_used[name] = erf9_used.get(name, 0) + 1
        else:
            sched.erf9[day] = "⚠️ OFFEN"
            sched.conflicts.append(f"❌ ERF9 {DAYS[day]}: Niemand ganztags frei")

    # ════════════════════════════════════════════════════════════
    # 10. RX ABO – N Halbtage pro Woche, max. 1 pro Tag, Pool "RX Abo"
    # ════════════════════════════════════════════════════════════
    rx_pool = [n for n in order if employees[n].can("RX Abo") and employees[n].regulaer]
    rx_target = int(rules["rx_abo_pro_woche"])
    if rx_pool and rx_target > 0:
        rx_days: set = set()
        rx_used: Dict[str, int] = {}
        done = 0
        for (d, s) in shuffled((d, s) for d in range(5) for s in range(2)):
            if done >= rx_target:
                break
            if d in rx_days:
                continue
            name = pick([n for n in rx_pool if sched.is_free(n, d, s)], rx_used)
            if name and sched.assign(name, d, s, "RX Abo"):
                rx_days.add(d)
                rx_used[name] = rx_used.get(name, 0) + 1
                done += 1
        if done < rx_target:
            sched.conflicts.append(f"⚠️ RX Abo: nur {done}/{rx_target} Halbtage besetzt")

    # ════════════════════════════════════════════════════════════
    # 11. POSTÖFFNUNG (Vormittag)
    #     - fehlt die Aushilfe: 1 Person PO/SCAN (scanning-fähig, nicht PHC)
    #     - zusätzlich N Personen "PO", rotierend (niemand jeden Tag)
    #     - Nachmittag: fehlt die Aushilfe, 1 Person "SCAN" (halbtags)
    #     - Dienstag bevorzugt die konfigurierte Person (po_dienstag_bevorzugt)
    #     - TL, nur-TEL und Aushilfe sind ausgenommen
    # ════════════════════════════════════════════════════════════
    po_used: Dict[str, int] = {}
    scan_used: Dict[str, int] = {}
    po_pref = rot.get("po_dienstag_bevorzugt") or ""

    def po_ok(name, day):
        emp = employees[name]
        return emp.can("PO") and not emp.is_tl and emp.regulaer and sched.is_free(name, day, 0)

    for day in range(5):
        if day not in sched.po_liste:
            cands = [n for n in order if po_ok(n, day) and employees[n].can("Scanning")
                     and sched.phc_liste.get(day) != n]
            name = None
            if day == 1 and po_pref in cands:
                name = po_pref
            name = name or pick(cands, po_used)
            if name:
                sched.assign(name, day, 0, "PO/SCAN")
                sched.po_liste[day] = name
                po_used[name] = po_used.get(name, 0) + 1
                scan_used[name] = scan_used.get(name, 0) + 1
            else:
                sched.po_liste[day] = "⚠️ OFFEN"
                sched.conflicts.append(f"❌ Postöffnung {DAYS[day]} VM: Niemand verfügbar!")
        for _ in range(int(rules["po_zusatz_vm"])):
            cands = [n for n in order if po_ok(n, day)]
            name = None
            if day == 1 and po_pref in cands:
                name = po_pref
            name = name or pick(cands, po_used)
            if not name:
                sched.conflicts.append(f"⚠️ PO {DAYS[day]} VM: zu wenig Personen für die Zusatz-PO")
                break
            sched.assign(name, day, 0, "PO")
            sched.po_zusatz[day].append(name)
            po_used[name] = po_used.get(name, 0) + 1
        if rules["scan_nm_ersatz_wenn_aushilfe_fehlt"] and not any(
                employees[n].aushilfe and is_scan_task(sched.get_task(n, day, 1)) for n in employees):
            cands = [n for n in order if employees[n].can("Scanning") and employees[n].regulaer
                     and not employees[n].is_tl and sched.is_free(n, day, 1)
                     and sched.phc_liste.get(day) != n
                     and not is_scan_task(sched.get_task(n, day, 0))
                     and scan_used.get(n, 0) < int(rules["scan_max_halbtage_pro_woche"])]
            name = pick(cands, scan_used)
            if name:
                sched.assign(name, day, 1, "SCAN")
                scan_used[name] = scan_used.get(name, 0) + 1

    # ════════════════════════════════════════════════════════════
    # 12. ERF8 – 1 Person pro Tag am Nachmittag (läuft bis 17:00)
    # ════════════════════════════════════════════════════════════
    erf8_used: Dict[str, int] = {}
    for day in range(5):
        if sched.erf8_assigned[day]:
            continue
        cands = [n for n in order if employees[n].can("ERF8") and employees[n].regulaer
                 and sched.is_free(n, day, 1)]
        name = pick(cands, erf8_used)
        if name:
            sched.assign(name, day, 1, "ERF8")
            erf8_used[name] = erf8_used.get(name, 0) + 1
        else:
            sched.conflicts.append(f"❌ ERF8 {DAYS[day]}: Niemand am Nachmittag frei")

    # ════════════════════════════════════════════════════════════
    # 13. TEL (Mo VM 4, sonst 3) – nur halbtags pro Person,
    #     TL bevorzugt (max. N pro Woche), sonst gleichmässig verteilt;
    #     Nachmittag notfalls TAGES PA → "TAGES PA/TEL"
    # ════════════════════════════════════════════════════════════
    tl_names = [n for n in order if employees[n].is_tl]
    tl_max = int(rules["tl_tel_max_pro_woche"])
    tel_week: Dict[str, int] = {n: sched.count_task(n, lambda t: "TEL" in t) for n in order}

    def tel_allowed(name, day, slot):
        if employees[name].nur_tel:
            return True
        return "TEL" not in sched.get_task(name, day, 1 - slot)

    tpa_mode = str(rules["tpa_tel_nm"])

    def tpa_mit_tel(day):
        tpa = sched.tagespa.get(day)
        if tpa in employees and sched.get_task(tpa, day, 1) == "TAGES PA":
            sched.schedule[tpa][(day, 1)] = "TAGES PA/TEL"
            sched.tel_count[day][1] += 1
            tel_week[tpa] += 1

    for day in range(5):
        for slot in range(2):
            target = int(rules["tel_montag_vm"]) if (day == 0 and slot == 0) else int(rules["tel_sonst"])
            if slot == 1 and tpa_mode == "immer":
                tpa_mit_tel(day)
            needed = target - sched.tel_count[day][slot]
            if needed > 0:
                avail = [n for n in sched.get_available(day, slot) if tel_allowed(n, day, slot)]
                tls = [n for n in avail if n in tl_names and tel_week[n] < tl_max]
                others = sorted((n for n in avail if n not in tl_names),
                                key=lambda n: (tel_week[n], rng.random()))
                rng.shuffle(tls)
                count = 0
                for name in tls + others:
                    if count >= needed:
                        break
                    if sched.assign(name, day, slot, "TEL"):
                        tel_week[name] += 1
                        count += 1
            if slot == 1 and tpa_mode == "bei Bedarf" and sched.tel_count[day][1] < target:
                tpa_mit_tel(day)


    # ════════════════════════════════════════════════════════════
    # 14. ABKL (2 pro Halbtag) – nur halbtags pro Person
    # ════════════════════════════════════════════════════════════
    abkl_week: Dict[str, int] = {n: sched.count_task(n, lambda t: "ABKL" in t) for n in order}
    for day in range(5):
        for slot in range(2):
            needed = int(rules["abkl_pro_halbtag"]) - sched.abkl_count[day][slot]
            if needed <= 0:
                continue
            avail = [n for n in sched.get_available(day, slot)
                     if employees[n].regulaer and "ABKL" not in sched.get_task(n, day, 1 - slot)]
            avail.sort(key=lambda n: (abkl_week[n], rng.random()))
            count = 0
            for name in avail:
                if count >= needed:
                    break
                if sched.assign(name, day, slot, "ABKL"):
                    abkl_week[name] += 1
                    count += 1

    # ════════════════════════════════════════════════════════════
    # 15. HUB – ERF7/HUB halbtags: 1 Person VM + 1 andere Person NM
    # ════════════════════════════════════════════════════════════
    hub_pool = [n for n in order if employees[n].can("HUB") and employees[n].regulaer]
    hub_used: Dict[str, int] = {}
    hub_per_day = int(rules["hub_pro_tag"])
    for day in range(5):
        today: set = set()
        for slot in range(min(2, hub_per_day)):
            cands = [n for n in hub_pool if n not in today and sched.is_free(n, day, slot)]
            name = pick(cands, hub_used)
            if name and sched.assign(name, day, slot, "ERF7/HUB"):
                today.add(name)
                hub_used[name] = hub_used.get(name, 0) + 1

    # ════════════════════════════════════════════════════════════
    # 16. RESTLICHE SLOTS: TL → ERF7/Q, sonst ERF7; scanning-fähige
    #     Personen ab und zu ein HALBER Tag "scanning" (max. 1 pro Tag,
    #     max. N pro Woche, nie die PHC-Person des Tages)
    # ════════════════════════════════════════════════════════════
    scan_max = int(rules["scan_max_halbtage_pro_woche"])
    scan_p = float(rules["scan_anteil"])
    for name in order:
        emp = employees[name]
        week_scans = sched.count_task(name, is_scan_task)
        for day in range(5):
            scanned_today = any(is_scan_task(sched.get_task(name, day, s)) for s in range(2))
            for slot in range(2):
                if not sched.is_free(name, day, slot):
                    continue
                if emp.is_tl:
                    sched.assign(name, day, slot, "ERF7/Q")
                elif (emp.can("Scanning") and emp.regulaer and not scanned_today
                      and week_scans < scan_max and sched.phc_liste.get(day) != name
                      and rng.random() < scan_p):
                    sched.assign(name, day, slot, "scanning")
                    scanned_today = True
                    week_scans += 1
                elif emp.regulaer:
                    sched.assign(name, day, slot, "ERF7")
                elif emp.nur_tel:
                    sched.assign(name, day, slot, "TEL")
                else:
                    sched.assign(name, day, slot, "scanning")

    # ════════════════════════════════════════════════════════════
    # 17. WOCHENAUFGABEN – je 1 Person, keine Person doppelt,
    #     ganzwöchig Abwesende werden übersprungen
    # ════════════════════════════════════════════════════════════
    used_weekly: set = set()

    def pick_week_person(role, label):
        pool = [n for n in order if employees[n].can(role) and employees[n].regulaer]
        if not pool:
            return "⚠️ OFFEN"
        idx = (widx + off.get(label, 0)) % len(pool)
        planned = pool[idx]
        for i in range(len(pool)):
            cand = pool[(idx + i) % len(pool)]
            if cand in used_weekly:
                continue
            if sched.present_anytime(cand):
                if cand != planned:
                    why = "ganze Woche abwesend" if not sched.present_anytime(planned) else "hat schon eine Wochenaufgabe"
                    sched.conflicts.append(f"ℹ️ {label}: {planned} {why} → {cand} übernimmt")
                used_weekly.add(cand)
                return cand
        used_weekly.add(planned)
        return planned

    sched.wochenaufgabe_direkt = pick_week_person("Direkt", "Direkt")
    sched.wochenaufgabe_onb = pick_week_person("ONB", "ONB")
    sched.wochenaufgabe_btm = pick_week_person("BTM", "BTM")

    return sched


# ============================================================
# VALIDIERUNG
# ============================================================

def validate_schedule(sched: Schedule) -> List[str]:
    """Prüft alle harten Regeln nach dem Planen (auch manuelle Einträge)."""
    issues: List[str] = list(sched.conflicts)
    emps = sched.employees
    rules = sched.cfg["regeln"]
    rot = sched.cfg["rotation"]
    tag_ausschluss = {DAY_KEYS.index(k): set(v) for k, v in rot["phc_tag_ausschluss"].items() if k in DAY_KEYS}

    # TEL / ABKL Mindestbesetzung
    for d in range(5):
        for s in range(2):
            target = int(rules["tel_montag_vm"]) if (d == 0 and s == 0) else int(rules["tel_sonst"])
            if sched.tel_count[d][s] < target:
                issues.append(f"❌ {DAYS[d]} {SLOTS[s]} TEL: {sched.tel_count[d][s]}/{target}")
            if sched.abkl_count[d][s] < int(rules["abkl_pro_halbtag"]):
                issues.append(f"❌ {DAYS[d]} {SLOTS[s]} ABKL: {sched.abkl_count[d][s]}/{rules['abkl_pro_halbtag']}")

    # ERF8, PO, TAGES PA, ERF9 pro Tag
    for d in range(5):
        if not sched.erf8_assigned[d]:
            issues.append(f"❌ {DAYS[d]} ERF8: Nicht besetzt")
        po = sched.po_liste.get(d, "")
        if not po or po.startswith("⚠"):
            issues.append(f"❌ {DAYS[d]} VM Postöffnung (PO/SCAN): Nicht besetzt")
        tpa = sched.tagespa.get(d, "")
        if not tpa or tpa.startswith("⚠"):
            issues.append(f"❌ {DAYS[d]} TAGES PA: Nicht besetzt")
        if not any("ERF9" in sched.get_task(n, d, 0) for n in emps):
            issues.append(f"❌ {DAYS[d]} ERF9: Nicht besetzt")

    # Halbtags-Regeln & Ganztags-Regeln pro Person
    for name, emp in emps.items():
        for d in range(5):
            vm, nm = sched.get_task(name, d, 0), sched.get_task(name, d, 1)
            if "TEL" in vm and "TEL" in nm and not emp.nur_tel:
                issues.append(f"❌ {name} {DAYS[d]}: TEL ganztags – TEL nur halbtags erlaubt")
            if "ABKL" in vm and "ABKL" in nm:
                issues.append(f"❌ {name} {DAYS[d]}: ABKL ganztags – ABKL nur halbtags erlaubt")
            if is_scan_task(vm) and is_scan_task(nm) and not emp.aushilfe:
                issues.append(f"❌ {name} {DAYS[d]}: Scanning ganztags – nur halbtags (Ausnahme Aushilfe)")
            if ("ERF9" in vm) != ("ERF9" in nm) and sched.is_present(name, d, 1 if "ERF9" in vm else 0):
                issues.append(f"⚠️ {name} {DAYS[d]}: ERF9 nur halbtags – ERF9 ist ganztags")
            if vm.startswith("TAGES PA") != nm.startswith("TAGES PA") and sched.is_present(name, d, 1 if vm.startswith("TAGES PA") else 0):
                issues.append(f"⚠️ {name} {DAYS[d]}: TAGES PA nur halbtags – TAGES PA ist ganztags")
            if ("SCH" in vm) != ("SCH" in nm) and sched.is_present(name, d, 1 if "SCH" in vm else 0):
                issues.append(f"⚠️ {name} {DAYS[d]}: Schalter nur halbtags – Schalter ist ganztags")
            if nm.startswith("PO/TEL"):
                issues.append(f"⚠️ {name} {DAYS[d]} NM: PO/TEL – nur im Notfall (muss früher vom Telefon weg)")
        scans = sched.count_task(name, is_scan_task)
        if not emp.aushilfe and scans > int(rules["scan_max_halbtage_pro_woche"]):
            issues.append(f"⚠️ {name}: {scans} Halbtage Scanning – max. {rules['scan_max_halbtage_pro_woche']} pro Woche")

    # PHC-Person: kein Scanning, nur Berechtigte, nicht an Ausschluss-Tagen
    for d in range(5):
        phc = sched.phc_liste.get(d, "")
        if phc not in emps:
            continue
        if any(is_scan_task(sched.get_task(phc, d, s)) for s in range(2)):
            issues.append(f"❌ {DAYS[d]}: {phc} hat PHC-Liste und Scanning – geht nicht (Arbeitszeiten)")
        emp = emps[phc]
        if not emp.can("PHC") or not emp.regulaer or phc in tag_ausschluss.get(d, set()):
            issues.append(f"❌ {DAYS[d]}: {phc} darf keine PHC-Liste machen")
        if not sched.is_present(phc, d, 1):
            issues.append(f"❌ {DAYS[d]}: PHC-Person {phc} ist am Nachmittag nicht da")

    # TL mit Tagesverantwortung: keine Ämtli
    for d in range(5):
        tv = sched.tagesverantwortung.get(d, "")
        if tv in emps:
            for s in range(2):
                task = sched.get_task(tv, d, s)
                if task and task != "TL/Q":
                    issues.append(f"⚠️ {DAYS[d]} {SLOTS[s]}: TV {tv} hat noch {task} – sollte TL/Q sein")

    # Rollen-Berechtigungen (greifen auch bei manuellen Einträgen)
    for name, emp in emps.items():
        for d in range(5):
            for s in range(2):
                task = sched.get_task(name, d, s)
                if not task:
                    continue
                where = f"{name} {DAYS[d]} {SLOTS[s]}"
                if "HUB" in task and not emp.can("HUB"):
                    issues.append(f"❌ {where}: HUB nicht erlaubt")
                if task.startswith("PO") and not emp.can("PO"):
                    issues.append(f"❌ {where}: Postöffnung nicht erlaubt")
                if task.startswith("TAGES PA") and not emp.can("TPA"):
                    issues.append(f"❌ {where}: TAGES PA nicht erlaubt")
                if is_scan_task(task) and not emp.can("Scanning"):
                    issues.append(f"❌ {where}: Scanning nicht erlaubt")
                if "SCH" in task and not (emp.can("Schalter") or emp.can("Schalter-Reserve")):
                    issues.append(f"❌ {where}: Schalter nicht erlaubt")
                if "RX Abo" in task and not emp.can("RX Abo"):
                    issues.append(f"❌ {where}: RX Abo nicht erlaubt")
                if emp.nur_tel and task not in ("TEL",) and task not in ABSENT_TASKS and not task.startswith("*") and task != "HO":
                    issues.append(f"⚠️ {where}: {task} – Person ist nur für TEL eingeplant")
                if task and not emp.is_available(d, s):
                    issues.append(f"❌ {where}: {task} eingetragen, aber laut Stammdaten nicht anwesend")

    # TL 1–2× pro Woche am Telefon (Soll)
    tl_min, tl_max = int(rules["tl_tel_min_pro_woche"]), int(rules["tl_tel_max_pro_woche"])
    for name, emp in emps.items():
        if not emp.is_tl or not sched.present_anytime(name):
            continue
        c = sched.count_task(name, lambda t: "TEL" in t)
        if tl_min and c < tl_min:
            issues.append(f"⚠️ TL {name}: {c}× TEL diese Woche – Soll mind. {tl_min}×")
        if c > tl_max:
            issues.append(f"⚠️ TL {name}: {c}× TEL diese Woche – Soll max. {tl_max}×")

    # Wochenaufgaben
    wa = [sched.wochenaufgabe_direkt, sched.wochenaufgabe_onb, sched.wochenaufgabe_btm]
    for label, who in zip(("Direkt", "ONB", "BTM"), wa):
        if not who or who.startswith("⚠"):
            issues.append(f"❌ Wochenaufgabe {label}: niemand berechtigt")
        elif who in emps and not emps[who].can(label):
            issues.append(f"❌ Wochenaufgabe {label}: {who} nicht berechtigt")
    real = [w for w in wa if w in emps]
    if len(set(real)) != len(real):
        issues.append("⚠️ Wochenaufgaben: eine Person hat mehrere Wochenaufgaben")

    return issues


def coverage_table(sched: Schedule) -> List[dict]:
    rules = sched.cfg["regeln"]
    out = []
    for d in range(5):
        for s in range(2):
            target = int(rules["tel_montag_vm"]) if (d == 0 and s == 0) else int(rules["tel_sonst"])
            out.append({
                "day": DAYS[d], "slot": SLOTS[s],
                "tel": sched.tel_count[d][s], "tel_target": target,
                "abkl": sched.abkl_count[d][s], "abkl_target": int(rules["abkl_pro_halbtag"]),
                "tel_ok": sched.tel_count[d][s] >= target,
                "abkl_ok": sched.abkl_count[d][s] >= int(rules["abkl_pro_halbtag"]),
            })
    return out


def check_rule_impact(absences: List[Tuple[str, int, int]], employees: Dict[str, Employee],
                      cfg: dict) -> List[str]:
    """Warnungen VOR dem Planen: betroffene Fixeinsätze und knappe Besetzung."""
    warnings: List[str] = []
    rules = normalize_config(cfg)["regeln"]
    for name, day, slot in absences:
        emp = employees.get(name)
        if not emp:
            continue
        tag = f"{name} {DAYS[day]} {SLOTS[slot]}"
        for d, s, task in emp.fix:
            if d == day and s == slot:
                warnings.append(f"{tag}: Fixeinsatz {task} fällt aus")
        if emp.aushilfe and slot == 0:
            warnings.append(f"{tag}: Aushilfe fehlt – Postöffnung (PO/SCAN) wird an jemand anderen vergeben")
    for day in range(5):
        for slot in range(2):
            absent = sum(1 for n, d, s in absences if d == day and s == slot)
            available = sum(1 for n in employees if employees[n].is_available(day, slot)) - absent
            target_tel = int(rules["tel_montag_vm"]) if (day == 0 and slot == 0) else int(rules["tel_sonst"])
            min_needed = target_tel + int(rules["abkl_pro_halbtag"]) + 1 + (1 + int(rules["po_zusatz_vm"]) if slot == 0 else 1)
            if available < min_needed:
                warnings.append(f"KRITISCH {DAYS[day]} {SLOTS[slot]}: nur {available} verfügbar – mind. {min_needed} benötigt")
            elif available < min_needed + 3:
                warnings.append(f"Knapp: {DAYS[day]} {SLOTS[slot]} – nur {available} Personen")
    return warnings


# ============================================================
# EXCEL
# ============================================================

def get_task_color(task: str) -> Optional[str]:
    if not task:
        return None
    if task in TASK_COLORS:
        return TASK_COLORS[task]
    for key in sorted((k for k in TASK_COLORS if k), key=len, reverse=True):
        if task.startswith(key):
            return TASK_COLORS[key]
    return None


def legend_rows(cfg: dict) -> list:
    """Legende mit dem konfigurierten Schalter-Code."""
    code = str(normalize_config(cfg)["regeln"].get("schalter_code") or "ERF7/SCH")
    return [[(code if c == "ERF7/SCH" else c, e) for c, e in row] for row in LEGEND_ROWS]


def tv_colors(cfg: dict) -> Dict[str, str]:
    cfg = normalize_config(cfg)
    colors = dict(cfg["rotation"].get("tv_farben", {}))
    tls = cfg["rotation"].get("tv_reihenfolge") or [m["short"] for m in cfg["mitarbeiter"] if m.get("tl")]
    i = 0
    for n in tls:
        if n not in colors:
            colors[n] = TV_FALLBACK_COLORS[i % len(TV_FALLBACK_COLORS)]
            i += 1
    return colors


def _is_dark(hex_color: str) -> bool:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.55


def write_excel(sched: Schedule, output) -> None:
    """Schreibt den Plan nach `output` (Dateipfad oder BytesIO)."""
    cfg = sched.cfg
    wb = Workbook()
    ws = wb.active
    ws.title = "Arbeitsplan"

    start = sched.week_start_date
    end = start + timedelta(days=4)
    date_str = f"KW {sched.week_number}: {start.strftime('%d.%m.%Y')} - {end.strftime('%d.%m.%Y')}"

    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    hdr_font = Font(bold=True, size=10, name="Arial")
    cell_font = Font(size=9, name="Arial")
    leg_font = Font(size=8, name="Arial")

    # ── Legende (Zeilen 1–6) ───────────────────────────────────────
    for i, row_items in enumerate(legend_rows(cfg)):
        r_code, r_expl = 1 + 2 * i, 2 + 2 * i
        ws.cell(row=r_code, column=1, value="Legende").font = hdr_font
        ws.cell(row=r_expl, column=1, value="Erklärung").font = leg_font
        for j, (code, expl) in enumerate(row_items):
            c = ws.cell(row=r_code, column=3 + j, value=code)
            c.font = leg_font
            c.alignment = Alignment(horizontal="center")
            color = PHC_COLOR if code == "PHC-Liste" else get_task_color(code)
            if code and color:
                c.fill = PatternFill("solid", fgColor=color)
                if _is_dark(color):
                    c.font = Font(size=8, name="Arial", color="FFFFFF")
            ws.cell(row=r_expl, column=3 + j, value=expl).font = leg_font

    # TV-Legende rechts
    ws.cell(row=1, column=14, value="Tagesverantwortung").font = hdr_font
    for i, (tl, color) in enumerate(tv_colors(cfg).items()):
        c = ws.cell(row=2 + i, column=14, value=tl)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(size=8, name="Arial", color="FFFFFF" if _is_dark(color) else "000000")

    # ── Datum (Zeile 8) ────────────────────────────────────────────
    ws.merge_cells("A8:L8")
    ws["A8"] = date_str
    ws["A8"].font = Font(bold=True, size=12, name="Arial")
    ws["A8"].alignment = Alignment(horizontal="center")

    # ── Header (Zeilen 10–11) ──────────────────────────────────────
    row = 10
    headers = ["Name", "%"]
    for day in DAYS:
        headers.extend([day, ""])
    headers.extend(["Wochen Aufgabe", "Bemerkungen"])
    for col, val in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.border = thin
    for i in range(5):
        ws.merge_cells(start_row=row, start_column=3 + i * 2, end_row=row, end_column=4 + i * 2)

    row = 11
    ws.cell(row=row, column=1, value="").border = thin
    ws.cell(row=row, column=2, value="").border = thin
    tvc = tv_colors(cfg)
    for day in range(5):
        tv_name = sched.tagesverantwortung.get(day, "")
        color = tvc.get(tv_name, "D9E2F3")
        for slot in range(2):
            c = ws.cell(row=row, column=3 + day * 2 + slot, value=SLOTS[slot])
            c.font = Font(bold=True, size=8, color="FFFFFF" if _is_dark(color) else "000000")
            c.alignment = Alignment(horizontal="center")
            c.border = thin
            c.fill = PatternFill("solid", fgColor=color)
    ws.cell(row=row, column=13, value="").border = thin
    ws.cell(row=row, column=14, value="").border = thin

    # ── Team ───────────────────────────────────────────────────────
    row = 12
    for m in cfg["mitarbeiter"]:
        name = m["short"]
        if name not in sched.employees:
            continue
        emp = sched.employees[name]
        ws.cell(row=row, column=1, value=emp.name).font = cell_font
        ws.cell(row=row, column=1).border = thin
        ws.cell(row=row, column=2, value=emp.pct).font = cell_font
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = thin
        for day in range(5):
            for slot in range(2):
                task = sched.get_task(name, day, slot)
                c = ws.cell(row=row, column=3 + day * 2 + slot, value=task or None)
                c.font = cell_font
                c.alignment = Alignment(horizontal="center")
                c.border = thin
                color = get_task_color(task)
                if color:
                    c.fill = PatternFill("solid", fgColor=color)
                if slot == 1 and sched.phc_liste.get(day) == name:
                    c.fill = PatternFill("solid", fgColor=PHC_COLOR)
                    c.font = Font(bold=True, size=9, name="Arial", color="FFFFFF")
        wa_parts = [lbl for lbl, who in (("Direkt", sched.wochenaufgabe_direkt),
                                          ("ONB", sched.wochenaufgabe_onb),
                                          ("BTM", sched.wochenaufgabe_btm)) if who == name]
        wa_cell = ws.cell(row=row, column=13)
        wa_cell.border = thin
        if wa_parts:
            wa_cell.value = "\n".join(wa_parts)
            wa_cell.font = Font(bold=True, size=9, name="Arial")
            wa_cell.alignment = Alignment(horizontal="center", wrap_text=True)
        bem_parts = [f"{DAY_KEYS[d]} {SLOT_KEYS[s]}: {sched.notes[(name, d, s)]}"
                     for d in range(5) for s in range(2) if (name, d, s) in sched.notes]
        bem_cell = ws.cell(row=row, column=14)
        bem_cell.border = thin
        if bem_parts:
            bem_cell.value = "\n".join(bem_parts)
            bem_cell.font = cell_font
            bem_cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    # ── BackUp Team ────────────────────────────────────────────────
    if cfg["backup_team"]:
        for col in range(1, 15):
            c = ws.cell(row=row, column=col, value="BackUp Team" if col == 1 else None)
            c.font = Font(bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="4472C4")
            c.border = thin
        row += 1
        for b in cfg["backup_team"]:
            ws.cell(row=row, column=1, value=b.get("name") or b["short"]).font = cell_font
            ws.cell(row=row, column=1).border = thin
            ws.cell(row=row, column=2).border = thin
            for day in range(5):
                for slot in range(2):
                    task = sched.get_backup_task(b["short"], day, slot)
                    c = ws.cell(row=row, column=3 + day * 2 + slot, value=task or None)
                    c.font = cell_font
                    c.alignment = Alignment(horizontal="center")
                    c.border = thin
                    color = get_task_color(task)
                    if color:
                        c.fill = PatternFill("solid", fgColor=color)
            ws.cell(row=row, column=13).border = thin
            ws.cell(row=row, column=14).border = thin
            row += 1

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 5
    for i in range(10):
        ws.column_dimensions[get_column_letter(3 + i)].width = 12
    ws.column_dimensions[get_column_letter(13)].width = 12
    ws.column_dimensions[get_column_letter(14)].width = 22
    ws.freeze_panes = "C12"
    wb.save(output)


# ============================================================
# KOMMANDOZEILE (für Tests / Automatisierung)
# ============================================================

def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="KSC Arbeitsplan – Plan für eine KW erzeugen")
    ap.add_argument("--montag", help="Montag der Zielwoche, z.B. 2026-10-19 (Standard: nächste Woche)")
    ap.add_argument("--config", help="Pfad zu team.json (Standard: neben diesem Modul)")
    ap.add_argument("--out", help="Excel-Datei (Standard: Arbeitsplan_KWxx_YYYYMMDD.xlsx)")
    args = ap.parse_args(argv)

    monday = datetime.strptime(args.montag, "%Y-%m-%d") if args.montag else get_next_monday()
    kw = monday.isocalendar()[1]
    cfg = load_config(args.config)
    sched = build_schedule(kw, monday, cfg=cfg)
    issues = validate_schedule(sched)
    out = args.out or f"Arbeitsplan_KW{kw}_{monday.strftime('%Y%m%d')}.xlsx"
    write_excel(sched, out)

    print(f"KW {kw}  {monday.strftime('%d.%m.%Y')} – {(monday + timedelta(days=4)).strftime('%d.%m.%Y')}")
    print(f"Datei: {out}")
    print("Tagesverantwortung:", ", ".join(f"{DAY_KEYS[d]} {sched.tagesverantwortung.get(d, '?')}" for d in range(5)))
    print("PHC-Liste:         ", ", ".join(f"{DAY_KEYS[d]} {sched.phc_liste.get(d, '?')}" for d in range(5)))
    print("TAGES PA:          ", ", ".join(f"{DAY_KEYS[d]} {sched.tagespa.get(d, '?')}" for d in range(5)))
    print(f"Wochenaufgaben: Direkt {sched.wochenaufgabe_direkt}, ONB {sched.wochenaufgabe_onb}, BTM {sched.wochenaufgabe_btm}")
    if issues:
        print(f"\n{len(issues)} Hinweis(e):")
        for i in issues:
            print("  ", i)
    else:
        print("\n✅ Alle Regeln eingehalten.")
    return 0 if not any(i.startswith("❌") for i in issues) else 1


if __name__ == "__main__":
    raise SystemExit(main())
