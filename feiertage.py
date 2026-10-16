#!/usr/bin/env python3
"""
Feiertage (Schweiz) für Jahresplan und Wochenplan
=================================================
Welche Feiertage ein Team frei hat, steht zusammen mit den Ferien in
abwesenheiten.json / abwesenheiten_ksv.json (Schlüssel «feiertage») und ist in
der App unter «📆 Jahresplan» änderbar:

    {"standard": {"Weihnachten": "ganz", "Heiligabend": "NM", …},
     "eigene":   [{"datum": "2026-12-24", "name": "Betriebsapéro", "frei": "NM"}]}

frei: "ganz" = ganzer Tag frei, "VM" / "NM" = nur dieser Halbtag ist frei.
Die Daten der Standard-Feiertage werden jedes Jahr neu berechnet (Ostern nach
dem gregorianischen Kalender) – es muss nichts nachgetragen werden.

An Feiertagen planen die Scheduler nichts, alle haben «Feiertag», und die
Besetzung wird nicht geprüft (`build_schedule(…, feiertage=woche(fe, montag))`).
Das Modul hat keine Abhängigkeiten.
"""

from __future__ import annotations

import copy
from datetime import date, datetime, timedelta
from typing import Callable, Dict, List, Optional, Tuple

FREI_WERTE = ["ganz", "VM", "NM"]
FREI_SLOTS = {"ganz": (0, 1), "VM": (0,), "NM": (1,)}
FREI_LABEL = {"ganz": "ganzer Tag", "VM": "Vormittag", "NM": "Nachmittag"}


def ostersonntag(jahr: int) -> date:
    """Ostersonntag (gregorianisch, Algorithmus nach Meeus/Jones/Butcher)."""
    a = jahr % 19
    b, c = divmod(jahr, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    monat, tag = divmod(h + l - 7 * m + 114, 31)
    return date(jahr, monat, tag + 1)


def _ostern(tage: int) -> Callable[[int], date]:
    return lambda j: ostersonntag(j) + timedelta(days=tage)


def _fix(monat: int, tag: int) -> Callable[[int], date]:
    return lambda j: date(j, monat, tag)


# Auswahl in der App – nicht jeder Kanton kennt jeden dieser Tage
KATALOG: List[Tuple[str, Callable[[int], date]]] = [
    ("Neujahr", _fix(1, 1)),
    ("Berchtoldstag", _fix(1, 2)),
    ("Karfreitag", _ostern(-2)),
    ("Ostermontag", _ostern(1)),
    ("Tag der Arbeit", _fix(5, 1)),
    ("Auffahrt", _ostern(39)),
    ("Pfingstmontag", _ostern(50)),
    ("Fronleichnam", _ostern(60)),
    ("Bundesfeier", _fix(8, 1)),
    ("Mariä Himmelfahrt", _fix(8, 15)),
    ("Allerheiligen", _fix(11, 1)),
    ("Mariä Empfängnis", _fix(12, 8)),
    ("Heiligabend", _fix(12, 24)),
    ("Weihnachten", _fix(12, 25)),
    ("Stephanstag", _fix(12, 26)),
    ("Silvester", _fix(12, 31)),
]
KATALOG_NAMEN = [n for n, _ in KATALOG]

# Voreinstellung: die in der Deutschschweiz verbreiteten Feiertage – im KSV-Plan 2026
# waren Karfreitag, Ostermontag, Auffahrt und Pfingstmontag frei, der 1. Mai nicht.
STANDARD = {n: "ganz" for n in ("Neujahr", "Berchtoldstag", "Karfreitag", "Ostermontag", "Auffahrt",
                                "Pfingstmontag", "Bundesfeier", "Weihnachten", "Stephanstag")}


def default() -> dict:
    return {"standard": dict(STANDARD), "eigene": []}


def _as_date(x) -> Optional[date]:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    try:
        return date.fromisoformat(str(x)[:10])
    except (TypeError, ValueError):
        return None


def normalize(fe: Optional[dict]) -> dict:
    """Ergänzt/bereinigt die Einstellung; None → Voreinstellung."""
    if not isinstance(fe, dict):
        return default()
    std = fe.get("standard")
    std = {n: v for n, v in (std.items() if isinstance(std, dict) else ()) if n in KATALOG_NAMEN and v in FREI_WERTE}
    eigene = []
    for e in fe.get("eigene") or []:
        d = _as_date((e or {}).get("datum"))
        if d is None:
            continue
        frei = e.get("frei") if e.get("frei") in FREI_WERTE else "ganz"
        eigene.append({"datum": d.isoformat(), "name": str(e.get("name") or "").strip() or "freier Tag", "frei": frei})
    eigene.sort(key=lambda e: e["datum"])
    return {"standard": {n: std[n] for n in KATALOG_NAMEN if n in std}, "eigene": eigene}


def katalog(jahr: int) -> List[Tuple[str, date]]:
    """Alle wählbaren Feiertage mit Datum im Jahr."""
    return [(n, fn(jahr)) for n, fn in KATALOG]


def tage(fe: dict, von, bis) -> Dict[date, Tuple[str, Tuple[int, ...]]]:
    """Freie Tage zwischen von und bis (inkl.): Datum → (Name, freie Halbtage).
    Wochenenden sind nicht enthalten; zwei Einträge am selben Tag werden zusammengelegt."""
    von, bis = _as_date(von), _as_date(bis)
    fe = normalize(fe)
    found: List[Tuple[date, str, str]] = []
    for jahr in range(von.year, bis.year + 1):
        for n, d in katalog(jahr):
            if n in fe["standard"]:
                found.append((d, n, fe["standard"][n]))
    found += [(_as_date(e["datum"]), e["name"], e["frei"]) for e in fe["eigene"]]
    out: Dict[date, Tuple[str, Tuple[int, ...]]] = {}
    for d, n, frei in sorted(found, key=lambda x: x[0]):
        if not (von <= d <= bis) or d.weekday() > 4:
            continue
        if d in out:
            name, slots = out[d]
            out[d] = (name if n in name else f"{name} / {n}", tuple(sorted(set(slots) | set(FREI_SLOTS[frei]))))
        else:
            out[d] = (n, FREI_SLOTS[frei])
    return out


def woche(fe: dict, montag) -> Dict[Tuple[int, int], str]:
    """Für build_schedule: {(Tag 0–4, Halbtag 0/1): Name} der Woche ab `montag`."""
    montag = _as_date(montag)
    out: Dict[Tuple[int, int], str] = {}
    for d, (name, slots) in tage(fe, montag, montag + timedelta(days=4)).items():
        for s in slots:
            out[((d - montag).days, s)] = name
    return out


def tag_text(name: str, slots) -> str:
    """'Weihnachten' bzw. 'Heiligabend (NM)' – für Kopfzeilen."""
    slots = tuple(slots)
    return name if len(slots) == 2 else f"{name} ({'VM' if slots == (0,) else 'NM'})"


def pro_tag(woche_fe: Dict[Tuple[int, int], str]) -> Dict[int, str]:
    """{(Tag, Halbtag): Name} → {Tag: 'Weihnachten' / 'Heiligabend (NM)'}."""
    out: Dict[int, str] = {}
    for d in sorted({d for d, _ in woche_fe}):
        slots = tuple(s for s in (0, 1) if (d, s) in woche_fe)
        out[d] = tag_text(woche_fe[(d, slots[0])], slots)
    return out


def copy_of(fe: dict) -> dict:
    return copy.deepcopy(normalize(fe))
