#!/usr/bin/env python3
"""
Feiertage: Berechnung (feiertage.py) und Wirkung in beiden Planern
(arbeitskalender.py für KSC, teamplaner.py für KSV).

    pytest tests/test_feiertage.py
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import arbeitskalender as ak  # noqa: E402
import feiertage as ft  # noqa: E402
import teamplaner as tp  # noqa: E402

TEAMS = [(ak, "team.json"), (tp, "team_ksv.json")]
WEIHNACHTEN = {"standard": {"Heiligabend": "NM", "Weihnachten": "ganz"}, "eigene": []}


def test_ostern_und_standard_feiertage():
    assert ft.ostersonntag(2025) == date(2025, 4, 20)
    assert ft.ostersonntag(2026) == date(2026, 4, 5)
    assert ft.ostersonntag(2027) == date(2027, 3, 28)
    tage = ft.tage(ft.default(), date(2026, 1, 1), date(2026, 12, 31))
    # wie im KSV-Plan 2026: Karfreitag, Ostermontag, Auffahrt und Pfingstmontag frei – der 1. Mai nicht
    for d in (date(2026, 4, 3), date(2026, 4, 6), date(2026, 5, 14), date(2026, 5, 25), date(2026, 12, 25)):
        assert d in tage
    assert date(2026, 5, 1) not in tage
    assert date(2026, 12, 26) not in tage   # Stephanstag 2026 ist ein Samstag


def test_woche_mit_halben_und_eigenen_freien_tagen():
    fe = {"standard": {"Heiligabend": "NM", "Weihnachten": "ganz"},
          "eigene": [{"datum": "2026-12-21", "name": "Brückentag", "frei": "VM"}]}
    w = ft.woche(fe, date(2026, 12, 21))
    assert w == {(0, 0): "Brückentag", (3, 1): "Heiligabend", (4, 0): "Weihnachten", (4, 1): "Weihnachten"}
    assert ft.pro_tag(w) == {0: "Brückentag (VM)", 3: "Heiligabend (NM)", 4: "Weihnachten"}


def test_einstellung_wird_bereinigt():
    assert ft.normalize(None) == ft.default()
    n = ft.normalize({"standard": {"Weihnachten": "ganz", "Gibtsnicht": "ganz", "Neujahr": "falsch"},
                      "eigene": [{"datum": "kaputt"}, {"datum": "2026-12-24"}]})
    assert n == {"standard": {"Weihnachten": "ganz"},
                 "eigene": [{"datum": "2026-12-24", "name": "freier Tag", "frei": "ganz"}]}


@pytest.mark.parametrize("mod, cfgfile", TEAMS)
def test_feiertag_im_plan(mod, cfgfile):
    """Alle Anwesenden haben «Feiertag», keine Pflicht-Regel meldet sich, der offene Halbtag wird geplant."""
    cfg = mod.load_config(ROOT / cfgfile)
    monday = date(2026, 12, 21)
    fe = ft.woche(WEIHNACHTEN, monday)
    sched = mod.build_schedule(52, monday, cfg=cfg, feiertage=fe)
    issues = mod.validate_schedule(sched)
    assert not [i for i in issues if i.startswith("❌")], issues
    for n, e in sched.employees.items():
        for d, s in fe:
            assert (sched.get_task(n, d, s) == "Feiertag") == e.is_available(d, s), (n, d, s)
    assert any(sched.get_task(n, 3, 0) not in ("", "Feiertag", "Frei") for n in sched.employees)
    assert all(c.get("feiertag") for c in mod.coverage_table(sched) if c["day"] == "Freitag")
    # Feiertag geht einer Abwesenheit vor
    name = next(n for n, e in sched.employees.items() if e.is_available(4, 0))
    krank = "KRANK"
    sched2 = mod.build_schedule(52, monday, {(name, 4, 0): krank}, cfg=cfg, feiertage=fe)
    assert sched2.get_task(name, 4, 0) == "Feiertag"


@pytest.mark.parametrize("mod, cfgfile", TEAMS)
def test_ohne_feiertage_unveraendert(mod, cfgfile):
    cfg = mod.load_config(ROOT / cfgfile)
    for monday in (date(2026, 11, 2), date(2027, 2, 1)):
        kw = monday.isocalendar()[1]
        a = mod.build_schedule(kw, monday, cfg=cfg)
        b = mod.build_schedule(kw, monday, cfg=cfg, feiertage={})
        assert a.schedule == b.schedule
        assert mod.validate_schedule(a) == mod.validate_schedule(b)


def test_ksc_rollen_am_halben_feiertag():
    cfg = ak.load_config(ROOT / "team.json")
    monday = date(2026, 12, 21)
    sched = ak.build_schedule(52, monday, cfg=cfg, feiertage=ft.woche(WEIHNACHTEN, monday))
    tv = sched.tagesverantwortung[3]
    assert tv in sched.employees
    assert (sched.get_task(tv, 3, 0), sched.get_task(tv, 3, 1)) == ("TL/Q", "Feiertag")
    assert sched.phc_liste[3] == "Feiertag"          # 18-Uhr-Liste nur mit offenem Nachmittag
    assert sched.tagespa[3] in sched.employees       # TAGES PA am offenen Vormittag
    assert sched.po_liste[3] in sched.employees      # Postöffnung am Vormittag
    for rolle in (sched.tagesverantwortung, sched.phc_liste, sched.tagespa, sched.schalter, sched.erf9):
        assert rolle[4] == "Feiertag"


def test_excel_zeigt_feiertag():
    from io import BytesIO
    import openpyxl
    monday = date(2026, 12, 21)
    for mod, cfgfile in TEAMS:
        sched = mod.build_schedule(52, monday, cfg=mod.load_config(ROOT / cfgfile),
                                   feiertage=ft.woche(WEIHNACHTEN, monday))
        buf = BytesIO()
        mod.write_excel(sched, buf)
        ws = openpyxl.load_workbook(BytesIO(buf.getvalue())).active
        values = [str(c.value) for row in ws.iter_rows() for c in row if c.value]
        assert any("Weihnachten" in v for v in values), mod.__name__
        assert values.count("Feiertag") >= 5, mod.__name__
