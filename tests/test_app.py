#!/usr/bin/env python3
"""
Oberflächen-Test (Streamlit AppTest, ohne Browser): beide Teams laden, Plan
generieren, alle Bereiche von «Team & Regeln» öffnen, Eintrag erfassen,
Team wechseln. Arbeitet auf Kopien von team.json / team_ksv.json – die Dateien
im Repo bleiben unverändert.

    pytest tests/test_app.py
"""
from __future__ import annotations

import shutil
from datetime import timedelta
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
