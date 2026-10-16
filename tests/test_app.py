#!/usr/bin/env python3
"""
Oberflächen-Test (Streamlit AppTest, ohne Browser): beide Teams laden, Plan
generieren, alle Bereiche von «Team & Regeln» öffnen, Eintrag erfassen,
Team wechseln, Jahresplan (Ferien von–bis, freie Tage). Arbeitet auf Kopien von
team.json / team_ksv.json – die Dateien im Repo bleiben unverändert.

    pytest tests/test_app.py
"""
from __future__ import annotations

import json
import shutil
from datetime import date, timedelta
from pathlib import Path

import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
APP = str(ROOT / "streamlit_app.py")
SECTIONS = {
    "KSC": ["👤 Personen", "🎯 Wer darf was", "📌 Feste Einsätze", "🔄 Rotationen", "⚖️ Regeln", "🗂️ BackUp-Team"],
    "KSV": ["👤 Personen", "🎯 Wer darf was", "📋 Aufgaben & Besetzung", "📌 Feste Einsätze", "🗂️ BackUp-Team"],
}


@pytest.fixture(autouse=True)
def daten(tmp_path, monkeypatch):
    for name in ("team.json", "team_ksv.json"):
        shutil.copy(ROOT / name, tmp_path / name)
    monkeypatch.setenv("KSC_TEAM_CONFIG", str(tmp_path / "team.json"))
    monkeypatch.setenv("KSV_TEAM_CONFIG", str(tmp_path / "team_ksv.json"))
    monkeypatch.setenv("KSC_ABSENCES", str(tmp_path / "abwesenheiten.json"))
    monkeypatch.setenv("KSV_ABSENCES", str(tmp_path / "abwesenheiten_ksv.json"))
    monkeypatch.delenv("ARBEITSPLAN_TEAM", raising=False)
    return tmp_path


def start(team: str | None = None) -> AppTest:
    at = AppTest.from_file(APP, default_timeout=60)
    if team:
        at.query_params["team"] = team
    at.run()
    assert not at.exception, [e.value for e in at.exception]
    return at


@pytest.mark.parametrize("team", ["KSC", "KSV"])
def test_team_ganz_durchklicken(team):
    at = start(team)
    assert at.session_state["team"] == team
    assert any(f"Arbeitsplan <span class=\"team\">{team}</span>" in m.value for m in at.markdown)
    at.button(key="cta_generate").click().run()
    assert not at.exception, [e.value for e in at.exception]
    res = at.session_state["teams"][team]["result"]
    assert res and res["team"] == team and res["excel_bytes"]
    assert not [i for i in res["issues"] if i.startswith("❌")], res["issues"]
    at.button(key="step4").click().run()                  # 📆 Jahresplan
    assert not at.exception, [e.value for e in at.exception]
    assert any('class="jp"' in m.value for m in at.markdown)
    at.button(key="step3").click().run()
    for sec in SECTIONS[team]:
        at.session_state[f"{team}·team_section"] = sec
        at.run()
        assert not at.exception, (sec, [e.value for e in at.exception])
    at.button(key="step1").click().run()
    assert not at.exception


def _umschalter(at: AppTest):
    return next(s for s in at.segmented_control if s.key in ("team", "team_pick"))


def test_standard_ist_ksc_und_wechsel_zu_ksv():
    at = start()
    assert at.session_state["team"] == "KSC"
    _umschalter(at).set_value("KSV").run()
    assert not at.exception, [e.value for e in at.exception]
    assert at.session_state["team"] == "KSV"
    assert any("Arbeitsplan <span class=\"team\">KSV</span>" in m.value for m in at.markdown)
    # KSC-Daten bleiben erhalten, wenn man zurückwechselt
    _umschalter(at).set_value("KSC").run()
    assert at.session_state["team"] == "KSC" and set(at.session_state["teams"]) == {"KSC", "KSV"}


def test_woche_und_team_wechseln_im_ergebnis():
    at = start("KSV")
    at.button(key="cta_generate").click().run()
    monday = at.session_state["teams"]["KSV"]["result"]["monday"]
    at.button(key="wk_next").click().run()                 # Plan der nächsten KW wird sofort neu erstellt
    assert not at.exception, [e.value for e in at.exception]
    assert at.session_state["step"] == 2
    assert at.session_state["teams"]["KSV"]["result"]["kw"] == (monday + timedelta(weeks=1)).isocalendar()[1]
    _umschalter(at).set_value("KSC").run()                  # KSC hat noch keinen Plan → Schritt 1
    assert not at.exception and at.session_state["step"] == 1
    assert at.button(key="step2").disabled


def test_beliebige_woche_waehlen():
    """Wochenauswahl oben: laufende Woche bis ein Jahr im Voraus, Plan für die gewählte KW."""
    at = start("KSC")
    wahl = at.selectbox(key="week_offset")
    assert wahl.options[0].endswith("diese Woche") and wahl.options[1].endswith("nächste Woche")
    assert len(wahl.options) == 54
    wahl.set_value(4).run()
    at.button(key="cta_generate").click().run()
    assert not at.exception, [e.value for e in at.exception]
    res = at.session_state["teams"]["KSC"]["result"]
    assert wahl.options[5].startswith(f"KW {res['kw']} ·")
    at.button(key="wk_prev").click().run()                 # Pfeile bleiben nutzbar
    assert at.session_state["week_offset"] == 3 and at.session_state["teams"]["KSC"]["result"]["kw"] != res["kw"]


def test_eintrag_erfassen_ksv(daten):
    at = start("KSV")
    at.selectbox(key="KSV·f_employee").set_value("Joy B.")
    at.segmented_control(key="KSV·f_type").set_value("Ferien")
    at.segmented_control(key="KSV·f_days").set_value(["Mo", "Di"])
    at.button(key="KSV·add_entry").click().run()
    assert not at.exception, [e.value for e in at.exception]
    entries = at.session_state["teams"]["KSV"]["entries"]
    assert len(entries) == 1 and entries[0]["name"] == "Joy" and entries[0]["days"] == [0, 1]
    assert (daten / "abwesenheiten_ksv.json").exists()
    at.button(key="cta_generate").click().run()
    sched = at.session_state["teams"]["KSV"]["sched"]
    assert sched.get_task("Joy", 0, 0) == "Ferien" and sched.get_task("Joy", 1, 1) == "Ferien"
    assert sched.get_task("Joy", 2, 0) != "Ferien"


def test_aufgabe_setzen_ksc_backup():
    """KSC: «📌 Aufgabe setzen» für eine BackUp-Zeile und «Termin» mit Notiz für eine Person."""
    at = start("KSC")
    at.selectbox(key="KSC·f_employee").set_value("Stephi R. (BackUp)")
    at.segmented_control(key="KSC·f_type").set_value("Custom").run()
    at.selectbox(key="KSC·f_task").set_value("KSV")
    at.segmented_control(key="KSC·f_days").set_value(["Mi"])
    at.segmented_control(key="KSC·f_slot").set_value("Vormittag")
    at.button(key="KSC·add_entry").click().run()
    at.selectbox(key="KSC·f_employee").set_value("Lara I.")
    at.segmented_control(key="KSC·f_type").set_value("Termin")
    at.segmented_control(key="KSC·f_days").set_value(["Fr"])
    at.segmented_control(key="KSC·f_slot").set_value("Nachmittag")
    at.text_input(key="KSC·f_note").input("Arzt 14:30")
    at.button(key="KSC·add_entry").click().run()
    assert not at.exception, [e.value for e in at.exception]
    at.button(key="cta_generate").click().run()
    sched = at.session_state["teams"]["KSC"]["sched"]
    assert sched.get_backup_task("Stephi", 2, 0) == "KSV" and sched.get_backup_task("Stephi", 2, 1) == ""
    assert sched.get_task("Lara", 4, 1) == "" and sched.notes[("Lara", 4, 1)] == "Arzt 14:30"


def test_fest_eingestelltes_team(monkeypatch):
    monkeypatch.setenv("ARBEITSPLAN_TEAM", "KSV")
    at = start("KSC")   # Link wird ignoriert
    assert at.session_state["team"] == "KSV"
    assert not [s for s in at.segmented_control if s.key in ("team", "team_pick")]


# ── 📆 Jahresplan ─────────────────────────────────────────────────────

def _montag(offset: int) -> date:
    """Montag der Woche, die oben mit week_offset=offset gewählt ist."""
    import arbeitskalender as ak
    return (ak.get_next_monday() + timedelta(weeks=offset)).date()


def _ohne_feiertage(daten, datei: str, eigene: list | None = None) -> None:
    """Testwoche unabhängig vom Datum: keine Standard-Feiertage, nur die angegebenen."""
    (daten / datei).write_text(json.dumps({"entries": [], "feiertage": {"standard": {}, "eigene": eigene or []}}),
                               encoding="utf-8")


def test_jahresplan_ferien_ueber_zwei_wochen(daten):
    _ohne_feiertage(daten, "abwesenheiten_ksv.json")
    at = start("KSV")
    at.button(key="step4").click().run()
    mo = _montag(3)
    at.selectbox(key="KSV·j_employee").set_value("Joy B.")
    at.segmented_control(key="KSV·j_type").set_value("Ferien")
    at.date_input(key="KSV·j_von").set_value(mo + timedelta(days=2))     # Mittwoch
    at.date_input(key="KSV·j_bis").set_value(mo + timedelta(days=8))     # Dienstag der Folgewoche
    at.button(key="KSV·j_add").click().run()
    assert not at.exception, [e.value for e in at.exception]
    e = at.session_state["teams"]["KSV"]["entries"][-1]
    assert (e["name"], e["type"], e["von"], e["bis"]) == ("Joy", "Ferien", (mo + timedelta(days=2)).isoformat(),
                                                         (mo + timedelta(days=8)).isoformat())
    gespeichert = json.loads((daten / "abwesenheiten_ksv.json").read_text(encoding="utf-8"))
    assert gespeichert["entries"][-1]["von"] == e["von"] and "feiertage" in gespeichert
    # Wochenplan: Ferien Mi–Fr, in der Folgewoche Mo–Di
    at.selectbox(key="week_offset").set_value(3).run()
    at.button(key="cta_generate").click().run()
    sched = at.session_state["teams"]["KSV"]["sched"]
    assert [sched.get_task("Joy", d, 0) == "Ferien" for d in range(5)] == [False, False, True, True, True]
    at.button(key="wk_next").click().run()
    sched = at.session_state["teams"]["KSV"]["sched"]
    assert [sched.get_task("Joy", d, 1) == "Ferien" for d in range(5)] == [True, True, False, False, False]


def test_jahresplan_freier_tag_gilt_im_wochenplan(daten):
    _ohne_feiertage(daten, "abwesenheiten.json")
    at = start("KSC")
    at.button(key="step4").click().run()
    freitag = _montag(2) + timedelta(days=4)
    next(d for d in at.date_input if d.label == "Datum").set_value(freitag)
    next(t for t in at.text_input if t.label == "Bezeichnung").input("Teamtag")
    next(b for b in at.button if "Freien Tag hinzufügen" in b.label).click().run()
    assert not at.exception, [e.value for e in at.exception]
    assert at.session_state["teams"]["KSC"]["feiertage"]["eigene"] == [
        {"datum": freitag.isoformat(), "name": "Teamtag", "frei": "ganz"}]
    at.selectbox(key="week_offset").set_value(2).run()
    at.button(key="cta_generate").click().run()
    res, sched = at.session_state["teams"]["KSC"]["result"], at.session_state["teams"]["KSC"]["sched"]
    assert res["feiertage"] == {4: "Teamtag"}
    assert sched.tagesverantwortung[4] == "Feiertag"
    assert not [i for i in res["issues"] if i.startswith("❌") and "Freitag" in i], res["issues"]
    assert all(sched.get_task(n, 4, s) in ("Feiertag", "") for n in sched.employees for s in (0, 1))


def test_plan_wird_nach_neuem_eintrag_aktualisiert(daten):
    _ohne_feiertage(daten, "abwesenheiten_ksv.json")
    at = start("KSV")
    at.button(key="cta_generate").click().run()               # Plan der nächsten Woche
    assert at.session_state["teams"]["KSV"]["sched"].get_task("Maria", 0, 0) != "Ferien"
    mo = _montag(0)
    at.button(key="step4").click().run()
    at.selectbox(key="KSV·j_employee").set_value("Maria G.")
    at.date_input(key="KSV·j_von").set_value(mo)
    at.date_input(key="KSV·j_bis").set_value(mo + timedelta(days=4))
    at.button(key="KSV·j_add").click().run()
    at.button(key="step2").click().run()                      # zurück zum Plan: wird neu erstellt
    assert not at.exception, [e.value for e in at.exception]
    sched = at.session_state["teams"]["KSV"]["sched"]
    assert all(sched.get_task("Maria", d, s) == "Ferien" for d in range(5) for s in (0, 1))
