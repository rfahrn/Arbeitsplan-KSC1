#!/usr/bin/env python3
"""
Regel-Validator: baut mehrere aufeinanderfolgende Wochen mit team.json und
prüft alle harten Regeln. Läuft ohne pytest:

    python tests/test_regeln.py            # 8 Wochen ab nächster Woche
    python tests/test_regeln.py 2026-10-19 12

… und auch mit pytest (`pytest tests/`).
"""
from __future__ import annotations

import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import arbeitskalender as ak  # noqa: E402


def wochen(start: datetime, n: int):
    for i in range(n):
        monday = start + timedelta(weeks=i)
        yield monday.isocalendar()[1], monday


def pruefe_woche(sched: "ak.Schedule") -> list[str]:
    """Harte Regeln (❌) aus validate_schedule plus strukturelle Prüfungen."""
    fehler = [i for i in ak.validate_schedule(sched) if i.startswith("❌")]
    emps = sched.employees
    cfg = sched.cfg
    # Tagesverantwortung: dieselbe TL VM+NM, als TL/Q eingetragen
    for d in range(5):
        tv = sched.tagesverantwortung.get(d, "")
        if tv in emps:
            if not emps[tv].is_tl:
                fehler.append(f"TV {ak.DAYS[d]}: {tv} ist keine TL")
            if sched.get_task(tv, d, 0) != "TL/Q" or sched.get_task(tv, d, 1) != "TL/Q":
                fehler.append(f"TV {ak.DAYS[d]}: {tv} nicht als TL/Q eingetragen")
        else:
            fehler.append(f"TV {ak.DAYS[d]}: offen")
    # Aushilfe: PO/SCAN VM + scanning NM an jedem anwesenden Tag
    for n, e in emps.items():
        if e.aushilfe:
            for d in range(5):
                if sched.is_present(n, d, 0) and sched.get_task(n, d, 0) != "PO/SCAN":
                    fehler.append(f"Aushilfe {n} {ak.DAYS[d]} VM: {sched.get_task(n, d, 0)!r} statt PO/SCAN")
                if sched.is_present(n, d, 1) and sched.get_task(n, d, 1) != "scanning":
                    fehler.append(f"Aushilfe {n} {ak.DAYS[d]} NM: {sched.get_task(n, d, 1)!r} statt scanning")
        if e.nur_tel:
            for d in range(5):
                for s in range(2):
                    if sched.is_present(n, d, s) and sched.get_task(n, d, s) != "TEL":
                        fehler.append(f"nur-TEL {n} {ak.DAYS[d]} {ak.SLOTS[s]}: {sched.get_task(n, d, s)!r}")
    # Fixeinsätze stehen im Plan (wenn anwesend)
    for n, e in emps.items():
        for d, s, task in e.fix:
            if sched.is_present(n, d, s) and sched.get_task(n, d, s) != task:
                fehler.append(f"Fixeinsatz {n} {ak.DAYS[d]} {ak.SLOTS[s]}: {sched.get_task(n, d, s)!r} statt {task}")
    # Niemand arbeitet, wenn laut Stammdaten nicht anwesend
    for n, e in emps.items():
        for d in range(5):
            for s in range(2):
                if sched.get_task(n, d, s) and not e.is_available(d, s):
                    fehler.append(f"{n} {ak.DAYS[d]} {ak.SLOTS[s]}: eingeplant, aber nicht anwesend")
    # Schalter: ganztags, gleiche Person, max. 1x pro Person
    sch_people = []
    for d in range(5):
        who = [n for n in emps if "SCH" in sched.get_task(n, d, 0)]
        who_nm = [n for n in emps if "SCH" in sched.get_task(n, d, 1)]
        if len(who) > 1 or who != who_nm:
            fehler.append(f"Schalter {ak.DAYS[d]}: VM {who} / NM {who_nm}")
        sch_people += who
    if len(sch_people) != len(set(sch_people)):
        fehler.append(f"Schalter: Person mehrfach in einer Woche {sch_people}")
    # HUB: höchstens eine Person pro Halbtag
    for d in range(5):
        for s in range(2):
            hub = [n for n in emps if "HUB" in sched.get_task(n, d, s)]
            if len(hub) > 1:
                fehler.append(f"HUB {ak.DAYS[d]} {ak.SLOTS[s]}: {hub}")
    # Wochenaufgaben: drei verschiedene, berechtigte, anwesende Personen
    wa = [sched.wochenaufgabe_direkt, sched.wochenaufgabe_onb, sched.wochenaufgabe_btm]
    if len(set(wa)) != 3:
        fehler.append(f"Wochenaufgaben doppelt: {wa}")
    for label, who in zip(("Direkt", "ONB", "BTM"), wa):
        if who not in emps or not emps[who].can(label) or not sched.present_anytime(who):
            fehler.append(f"Wochenaufgabe {label}: {who} ungültig")
    return fehler


def test_regeln_mehrere_wochen():
    cfg = ak.load_config()
    assert not ak.validate_config(cfg), ak.validate_config(cfg)
    start = datetime(2026, 10, 19)
    for kw, monday in wochen(start, 8):
        sched = ak.build_schedule(kw, monday, cfg=cfg)
        fehler = pruefe_woche(sched)
        assert not fehler, f"KW {kw}: {fehler}"


def test_determinismus():
    cfg = ak.load_config()
    monday = datetime(2026, 10, 19)
    a = ak.build_schedule(43, monday, cfg=cfg)
    b = ak.build_schedule(43, monday, cfg=cfg)
    assert a.schedule == b.schedule and a.phc_liste == b.phc_liste


def test_abwesenheit_aushilfe_und_tl():
    """Aushilfe krank → PO/SCAN wird ersetzt; TL krank → andere TL übernimmt TV."""
    cfg = ak.load_config()
    monday = datetime(2026, 10, 19)
    aushilfe = next(m["short"] for m in cfg["mitarbeiter"] if m.get("aushilfe"))
    tl = cfg["rotation"]["tv_reihenfolge"][0]
    overrides = {(aushilfe, d, s): "KRANK" for d in range(5) for s in range(2)}
    overrides.update({(tl, d, s): "FERIEN" for d in range(5) for s in range(2)})
    sched = ak.build_schedule(43, monday, overrides, cfg=cfg)
    for d in range(5):
        assert sched.po_liste.get(d) not in (None, aushilfe, "⚠️ OFFEN"), sched.po_liste
        assert sched.tagesverantwortung.get(d) not in (tl, "⚠️ OFFEN")
        assert sched.get_task(aushilfe, d, 0) == "KRANK"
    fehler = [f for f in pruefe_woche(sched) if "ERF9" not in f]
    assert not fehler, fehler


def test_person_deaktivieren_und_neu():
    """Konfiguration ändern ohne Code: Person deaktivieren, neue Person anlegen."""
    cfg = ak.load_config()
    for m in cfg["mitarbeiter"]:
        if m["short"] == "Alessia":
            m["aktiv"] = False
    cfg["mitarbeiter"].append({
        "short": "Neu", "name": "Neu N.", "pct": 100, "aktiv": True,
        "kann": ["HUB", "Schalter", "Scanning", "PO", "TPA", "PHC", "ERF8", "ERF9", "ONB"],
    })
    cfg = ak.normalize_config(cfg)
    assert not ak.validate_config(cfg)
    sched = ak.build_schedule(44, datetime(2026, 10, 26), cfg=cfg)
    assert "Alessia" not in sched.employees and "Neu" in sched.employees
    assert not pruefe_woche(sched)


def test_entscheidungen():
    """Tab «Entscheidungen»: Schalter-Reserve, Schalter-Code, TL-Minimum, TAGES PA/TEL."""
    base = ak.load_config()
    reserve = {m["short"] for m in base["mitarbeiter"] if "Schalter-Reserve" in m.get("kann", [])}
    assert reserve, "team.json: niemand als Schalter-Reserve"
    start = datetime(2026, 10, 19)

    def run(cfg, n=12):
        return [ak.build_schedule(m.isocalendar()[1], m, cfg=cfg) for _, m in wochen(start, n)]

    # Reserve «nie» → nie Schalter; «vor der anderen Gruppe» → kommt vor (Gruppe B hat nur 4 Personen)
    cfg = ak.normalize_config(base); cfg["regeln"]["schalter_reserve_modus"] = "nie"
    assert not any(s.schalter.get(d) in reserve for s in run(cfg) for d in range(5))
    cfg = ak.normalize_config(base); cfg["regeln"]["schalter_reserve_modus"] = "vor der anderen Gruppe"
    scheds = run(cfg)
    assert any(s.schalter.get(d) in reserve for s in scheds for d in range(5))
    for s in scheds:
        assert not pruefe_woche(s), pruefe_woche(s)
    # Schalter-Code
    cfg = ak.normalize_config(base); cfg["regeln"]["schalter_code"] = "ERF4/SCH"
    s = run(cfg, 1)[0]
    assert any(s.get_task(n, 0, 0) == "ERF4/SCH" for n in s.employees)
    assert not pruefe_woche(s)
    # TL-Minimum 1: jede anwesende TL hat mind. 1 TEL
    cfg = ak.normalize_config(base); cfg["regeln"]["tl_tel_min_pro_woche"] = 1
    for s in run(cfg):
        for n, e in s.employees.items():
            if e.is_tl:
                assert s.count_task(n, lambda t: "TEL" in t) >= 1, (s.week_number, n)
        assert not pruefe_woche(s)
    # TAGES PA/TEL immer / nie
    cfg = ak.normalize_config(base); cfg["regeln"]["tpa_tel_nm"] = "immer"
    s = run(cfg, 1)[0]
    assert all(s.get_task(s.tagespa[d], d, 1) == "TAGES PA/TEL" for d in range(5))
    assert not pruefe_woche(s)
    cfg = ak.normalize_config(base); cfg["regeln"]["tpa_tel_nm"] = "nie"
    s = run(cfg, 1)[0]
    assert not any(s.get_task(n, d, 1) == "TAGES PA/TEL" for n in s.employees for d in range(5))


def test_team_bearbeiten():
    """Helfer der Seite «Team & Regeln»: Kurzname automatisch, TL-Rotation, Entfernen räumt auf."""
    cfg = ak.load_config()
    # Kurzname: Vorname, bei Kollision der ganze Name
    assert ak.new_short("Sandra K.", cfg) == "Sandra"
    assert ak.new_short("Lara K.", cfg) == "Lara K."
    # Neue TL kommt ans Ende von TV- und PHC-Montag-Rotation, ehemalige TL fällt heraus
    ak.set_funktion(cfg, "Amra", "Teamleitung")
    assert cfg["rotation"]["tv_reihenfolge"][-1] == "Amra" and cfg["rotation"]["phc_montag_tl"][-1] == "Amra"
    ak.set_funktion(cfg, "Amra", "Team")
    assert "Amra" not in cfg["rotation"]["tv_reihenfolge"] + cfg["rotation"]["phc_montag_tl"]
    # Entfernen: keine Verweise mehr, Konfiguration bleibt gültig und planbar
    assert ak.rotation_refs(cfg, "Dipiga") == ["Labor Mi NM", "18-Uhr-Liste fix Di"]
    ak.remove_member(cfg, "Dipiga")
    assert not ak.rotation_refs(cfg, "Dipiga") and "Dipiga" not in cfg["rotation"]["phc_fix"].values()
    cfg = ak.normalize_config(cfg)
    assert not ak.validate_config(cfg), ak.validate_config(cfg)
    assert not pruefe_woche(ak.build_schedule(44, datetime(2026, 10, 26), cfg=cfg))


def test_inaktiv_in_rotation_ist_nur_hinweis():
    cfg = ak.load_config()
    for m in cfg["mitarbeiter"]:
        if m["short"] == "Jesika":
            m["aktiv"] = False
    assert not ak.config_errors(cfg)
    assert any("Jesika" in h and "übersprungen" in h for h in ak.config_hints(cfg))
    sched = ak.build_schedule(44, datetime(2026, 10, 26), cfg=cfg)
    assert "Jesika" not in sched.employees and not pruefe_woche(sched)


def test_name_doppelt_ist_fehler():
    cfg = ak.load_config()
    cfg["mitarbeiter"][1]["name"] = cfg["mitarbeiter"][0]["name"]
    assert any(e.startswith("Name doppelt") for e in ak.config_errors(cfg))


def test_feste_einsaetze_als_liste():
    """Fixeinsätze + «alle 2 Wochen frei» als eine Liste – unverändert zurückschreiben ergibt dasselbe."""
    base = ak.load_config()
    cfg = ak.normalize_config(base)
    rows = ak.recurring_rows(cfg)
    assert {"short": "Linda", "tag": "Fr", "halbtag": "ganz", "task": "frei", "kw": "gerade"} in rows
    assert not ak.apply_recurring_rows(cfg, rows)
    assert ak.config_to_json(cfg) == ak.config_to_json(ak.normalize_config(base))
    assert not ak.describe_changes(base, cfg)
    # neue Zeile + Fehlerfälle
    rows.append({"short": "Amra", "tag": "Mo", "halbtag": "VM", "task": "HO", "kw": None})
    assert not ak.apply_recurring_rows(cfg, rows)
    assert {"tag": "Mo", "halbtag": "VM", "task": "HO"} in next(m for m in cfg["mitarbeiter"] if m["short"] == "Amra")["fix"]
    assert ak.describe_changes(base, cfg) == ["Amra I.: fester Einsatz Mo VM HO"]
    assert ak.apply_recurring_rows(cfg, rows + [{"short": "Amra", "tag": "Fr", "halbtag": "ganz", "task": "frei", "kw": None}])
    assert ak.apply_recurring_rows(cfg, rows + [{"short": "Linda", "tag": "Mo", "halbtag": "ganz", "task": "frei", "kw": "ungerade"}])


def test_aenderungen_lesbar():
    base = ak.load_config()
    cfg = ak.normalize_config(base)
    lara = next(m for m in cfg["mitarbeiter"] if m["short"] == "Lara")
    lara["pct"] = 80
    lara["kann"] = [r for r in lara["kann"] if r != "HUB"]
    cfg["regeln"]["abkl_pro_halbtag"] = 3
    assert ak.describe_changes(base, cfg) == [
        "Lara I.: Pensum 100 → 80 %", "Lara I.: darf nicht mehr HUB", "Regel «ABKL pro Halbtag»: 2 → 3"]


if __name__ == "__main__":
    start = datetime.strptime(sys.argv[1], "%Y-%m-%d") if len(sys.argv) > 1 else ak.get_next_monday()
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 8
    cfg = ak.load_config()
    probleme = ak.validate_config(cfg)
    print("team.json:", "ok" if not probleme else probleme)
    total = 0
    for kw, monday in wochen(start, n):
        sched = ak.build_schedule(kw, monday, cfg=cfg)
        fehler = pruefe_woche(sched)
        hinweise = [i for i in ak.validate_schedule(sched) if not i.startswith("❌")]
        total += len(fehler)
        status = "✅" if not fehler else "❌"
        print(f"{status} KW {kw:2d} ({monday:%d.%m.%Y}): {len(fehler)} Fehler, {len(hinweise)} Hinweise"
              f" | TV {[sched.tagesverantwortung[d] for d in range(5)]}")
        for f in fehler:
            print("     ", f)
    for name in ("test_determinismus", "test_abwesenheit_aushilfe_und_tl", "test_person_deaktivieren_und_neu",
                 "test_entscheidungen"):
        try:
            globals()[name]()
            print(f"✅ {name}")
        except AssertionError as exc:
            total += 1
            print(f"❌ {name}: {exc}")
    print("\nERGEBNIS:", "alle Regeln eingehalten 🎉" if total == 0 else f"{total} Fehler")
    sys.exit(0 if total == 0 else 1)
