#!/usr/bin/env python3
"""
KSC Arbeitsplan — Streamlit-Variante v2 (Test-Build)
=====================================================
Identisch zur Hauptvariante, plus:
- Live-Regenerierung: Klick auf ‹ / › auf der Ergebnis-Seite plant die
  neue Kalenderwoche sofort neu mit den bestehenden Step-1-Eingaben
  (Abwesenheiten, Team, Pensum). Step 1 dient nur noch dem Anpassen der
  Eingaben; die Resultate für KW+1, KW+2, … werden automatisch erzeugt.

Start:
    pip install streamlit openpyxl
    streamlit run streamlit_app2.py
"""

from __future__ import annotations

import os
import random
import sys
from datetime import timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Kern-Logik importieren ────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
APP_DIR = BASE_DIR / "ksc_arbeitsplan" / "app"
for p in (str(BASE_DIR), str(APP_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

import arbeitskalender as _ak  # noqa: E402
from arbeitskalender import (  # noqa: E402
    DAYS,
    SLOTS,
    Employee,
    build_schedule,
    create_employees,
    get_next_monday,
    validate_schedule,
    write_excel,
)

# ── Monkey-patch create_employees so user-edited team is honoured ─────
_ORIG_CREATE_EMPLOYEES = _ak.create_employees

_SLOT_ORDER = [(d, s) for d in range(5) for s in range(2)]


def _adjust_availability_to_pct(emp: Employee, pct: int) -> None:
    """Match the count of available (day, slot) pairs to the given pensum.

    10 weekly slots total (5 days × VM/NM); each 10 % pensum equals one slot.
    Slots are added in calendar order (Mo VM → Fr NM) and removed in reverse,
    so increasing Brigitte from 70 % to 80 % unblocks Mi NM first instead of
    leaving her at her hardcoded 7-slot availability.
    """
    target = max(0, min(10, round(pct / 10)))
    for slot in _SLOT_ORDER:
        emp.available.setdefault(slot, False)
    current = sum(1 for slot in _SLOT_ORDER if emp.available[slot])

    if current < target:
        for slot in _SLOT_ORDER:
            if current >= target:
                break
            if not emp.available[slot]:
                emp.available[slot] = True
                current += 1
    elif current > target:
        for slot in reversed(_SLOT_ORDER):
            if current <= target:
                break
            if emp.available[slot]:
                emp.available[slot] = False
                current -= 1


def _patched_create_employees(week_number: int):
    base = _ORIG_CREATE_EMPLOYEES(week_number)
    team = st.session_state.get("team")
    if not team:
        return base
    out: dict = {}
    for entry in team:
        if not entry.get("active", True):
            continue
        short = entry["short"]
        pct = int(entry["pct"])
        if entry.get("is_default") and short in base:
            emp = base[short]
            if emp.pct != pct:
                _adjust_availability_to_pct(emp, pct)
                emp.pct = pct
            out[short] = emp
        elif not entry.get("is_default"):
            new_emp = Employee(
                name=entry["name"],
                short=short,
                pct=pct,
                can_schalter=True,
                can_scanning=True,
                can_onb=True,
            )
            for d in range(5):
                for s in range(2):
                    new_emp.available[(d, s)] = True
            _adjust_availability_to_pct(new_emp, pct)
            out[short] = new_emp
    return out


_ak.create_employees = _patched_create_employees

# ── Konstanten ────────────────────────────────────────────────────────

EMPLOYEE_LIST = [
    "Silvana", "Linda", "Lara", "Andrea G.", "Dipiga",
    "Isaura", "Martina", "Alessia", "Amra", "Nina",
    "Jesika", "Brigitte", "Florence", "Corinne", "Saskia",
    "Dragi", "Maria B.", "Andrea A.",
]

ABSENCE_TYPES = {
    "Krank": "KRANK",
    "Ferien": "FERIEN",
    "Termin": "TERMIN",
    "Home Office": "HO",
    "Custom": "CUSTOM",
}

DAY_SHORTS = ["Mo", "Di", "Mi", "Do", "Fr"]

# Pastellfarben für die Pill-Zellen — bg + dunklere Schrift
TASK_PILL = {
    "TEL":      ("#D6F0DE", "#1E6F3D"),
    "ABKL":     ("#FFE9B0", "#7A4F00"),
    "KTG":      ("#D6F0DE", "#1E6F3D"),
    "KGS":      ("#D6F0DE", "#1E6F3D"),
    "ERF7":     ("#DCEAF7", "#1F4E79"),
    "ERF7/Q":   ("#DCEAF7", "#1F4E79"),
    "ERF7/HUB": ("#FFE0CC", "#9C5400"),
    "ERF8":     ("#FFD9D2", "#8C2A1C"),
    "ERF8/Q":   ("#FFD9D2", "#8C2A1C"),
    "ERF8/HUB": ("#FFE0CC", "#9C5400"),
    "ERF9":     ("#FFD0D0", "#9B1C1C"),
    "ERF9/Q":   ("#FFD0D0", "#9B1C1C"),
    "ERF9/TEL": ("#FFD0D0", "#9B1C1C"),
    "ERF4/SCH": ("#E2EFDA", "#386641"),
    "ERF5":     ("#FCE4D6", "#9C4A1A"),
    "PO":       ("#E7E6E6", "#3F3F3F"),
    "PO/SCAN":  ("#E7E6E6", "#3F3F3F"),
    "PO/ABKL":  ("#FFE9B0", "#7A4F00"),
    "PO/TEL":   ("#D6F0DE", "#1E6F3D"),
    "HO":       ("#DDEBF7", "#1F4E79"),
    "HO/Q":     ("#DDEBF7", "#1F4E79"),
    "VBZ/Q":    ("#E2EFDA", "#386641"),
    "KSC Spez.":("#FFF2CC", "#806000"),
    "TAGES PA": ("#CFEFFA", "#0E6E8C"),
    "TAGESPA":  ("#CFEFFA", "#0E6E8C"),
    "RX Abo":   ("#E7E6E6", "#3F3F3F"),
    "KRANK":    ("#FFC7CE", "#9C0006"),
    "FERIEN":   ("#F2F2F2", "#595959"),
    "scanning": ("#E2EFDA", "#386641"),
}

# Drop-down options for the editable Wochenplan. Order: leer, Pflichtrollen,
# Erfassungs-Kombis, PO-Kombis, Schalter / Labor, HO, Spezialrollen,
# Abwesenheiten. The empty string clears the cell.
TASK_OPTIONS = [
    "",
    "TEL", "ABKL",
    "ERF7", "ERF7/Q", "ERF7/HUB",
    "ERF8", "ERF8/Q", "ERF8/HUB",
    "ERF9", "ERF9/Q", "ERF9/TEL",
    "ERF4/SCH", "ERF5",
    "PO", "PO/Q", "PO/SCAN", "PO/ABKL", "PO/TEL",
    "HO", "HO/Q",
    "KGS", "KTG",
    "KSC Spez.", "TAGES PA",
    "RX Abo", "VBZ/Q",
    "KRANK", "FERIEN",
]

OUTPUT_DIR = Path(os.environ.get("KSC_OUTPUT_DIR", BASE_DIR / "output"))
STATE_FILE = Path(
    os.environ.get("KSC_STATE_FILE", OUTPUT_DIR / "scheduler_state.json")
)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ── Streamlit-Setup ────────────────────────────────────────────────────

st.set_page_config(
    page_title="KSC Arbeitsplan",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CUSTOM_CSS = """
<style>
:root {
  --bg:        #f4f6fb;
  --surface:   #ffffff;
  --surface-2: #f8fafc;
  --border:    #e3e8f0;
  --border-2:  #cbd5e1;
  --text:      #0f172a;
  --text-dim:  #475569;
  --text-mute: #94a3b8;
  --heading:   #0b1020;
  --primary:   #0f172a;
  --primary-2: #1e293b;
  --accent:    #2563eb;
  --accent-soft: #dbeafe;
  --ok:        #16a34a;
  --ok-soft:   #dcfce7;
  --warn:      #d97706;
  --warn-soft: #fef3c7;
  --err:       #dc2626;
  --err-soft:  #fee2e2;
  --phc:       #1e3a8a;
}

html, body, [data-testid="stAppViewContainer"] {
  background: var(--bg);
  color: var(--text);
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stSidebar"] { display: none; }
[data-testid="stSidebarCollapsedControl"] { display: none; }

.block-container {
  padding-top: 1.2rem !important;
  padding-bottom: 2rem !important;
  max-width: 1500px;
}

/* ── Top Header Bar ─────────────────────────────────────────── */

.topbar {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px 16px;
  margin-bottom: 14px;
  box-shadow: 0 1px 2px rgba(15,23,42,.04);
}

.brand {
  display: flex; align-items: center; gap: 12px;
}
.brand-icon {
  width: 38px; height: 38px; border-radius: 9px;
  background: var(--primary); color: white;
  display: grid; place-items: center; font-size: 18px;
}
.brand-title { font-size: 15px; font-weight: 700; color: var(--heading); line-height: 1.1; }
.brand-sub   { font-size: 10.5px; color: var(--text-mute); letter-spacing: .8px; text-transform: uppercase; }

/* Stepper */
.stepper { display: flex; align-items: center; gap: 8px; }
.step {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 7px 14px; border-radius: 999px;
  background: transparent; color: var(--text-dim);
  font-size: 13px; font-weight: 500;
  cursor: pointer; user-select: none;
  border: 1px solid transparent;
  transition: all .15s;
}
.step:hover { background: var(--surface-2); }
.step .num {
  width: 22px; height: 22px; border-radius: 50%;
  background: var(--surface-2); color: var(--text-dim);
  display: grid; place-items: center;
  font-size: 11.5px; font-weight: 700;
  border: 1px solid var(--border);
}
.step.active {
  background: var(--accent-soft);
  color: var(--accent);
  border-color: rgba(37,99,235,.18);
}
.step.active .num { background: var(--accent); color: white; border-color: var(--accent); }
.step-sep { color: var(--text-mute); padding: 0 2px; font-weight: 300; }

/* KW Pill */
.kw-pill {
  display: flex; align-items: center; gap: 4px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 4px;
}
.kw-center {
  padding: 4px 14px; min-width: 170px; text-align: center;
}
.kw-big   { display:block; font-size: 14px; font-weight: 700; color: var(--heading); }
.kw-range { display:block; font-size: 11px; color: var(--text-mute); margin-top: 2px; }

.badge-kw {
  background: #f1f5f9; color: var(--text-dim);
  border: 1px solid var(--border);
  padding: 6px 12px; border-radius: 999px;
  font-size: 11px; font-weight: 700; letter-spacing: .5px;
  text-transform: uppercase;
}

/* CTA Button (mache native primary buttons dunkel) */
div.stButton > button[kind="primary"] {
  background: var(--primary) !important;
  border: 1px solid var(--primary) !important;
  color: white !important;
  font-weight: 600 !important;
  border-radius: 9px !important;
  padding: 8px 18px !important;
  box-shadow: 0 2px 6px rgba(15,23,42,.18) !important;
}
div.stButton > button[kind="primary"]:hover {
  background: var(--primary-2) !important;
  border-color: var(--primary-2) !important;
  transform: translateY(-1px);
}

/* Secondary buttons (light) */
div.stButton > button[kind="secondary"] {
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  color: var(--text) !important;
  border-radius: 8px !important;
  font-weight: 500 !important;
}
div.stButton > button[kind="secondary"]:hover {
  background: var(--surface-2) !important;
  border-color: var(--border-2) !important;
}

/* Prevent vertical text wrapping in narrow button columns */
div.stButton > button,
div.stButton > button > div,
div.stButton > button p {
  white-space: nowrap !important;
}
div.stButton > button { padding-left: 6px !important; padding-right: 6px !important; }

div.stDownloadButton > button {
  background: var(--ok) !important;
  border: 1px solid var(--ok) !important;
  color: white !important;
  border-radius: 9px !important;
  font-weight: 600 !important;
}
div.stDownloadButton > button:hover { filter: brightness(1.08); }

/* Inputs */
div[data-baseweb="select"] > div,
.stTextInput input,
.stMultiSelect div[data-baseweb="select"] > div {
  background: var(--surface) !important;
  border-color: var(--border) !important;
  border-radius: 8px !important;
}
.stTextInput input { color: var(--text) !important; }

/* ── Stats Strip ────────────────────────────────────────────── */

.stats-strip {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 18px;
  margin-bottom: 16px;
  display: flex; gap: 36px; align-items: center;
  flex-wrap: wrap;
  box-shadow: 0 1px 2px rgba(15,23,42,.04);
}
.stat-item .lbl {
  font-size: 10.5px; color: var(--text-mute);
  text-transform: uppercase; letter-spacing: .8px;
  font-weight: 600; margin-bottom: 4px;
}
.stat-item .val {
  font-size: 18px; font-weight: 700; color: var(--heading);
  font-feature-settings: "tnum";
}
.stat-item .val small { font-size: 11.5px; font-weight: 500; color: var(--text-dim); margin-left: 4px; }
.stat-divider {
  height: 32px; width: 1px; background: var(--border);
}
.info-pill {
  margin-left: auto;
  display: flex; align-items: center; gap: 8px;
  background: #eff6ff;
  border: 1px solid #bfdbfe;
  color: #1e3a8a;
  padding: 7px 14px;
  border-radius: 999px;
  font-size: 12.5px;
}
.info-pill .ico {
  width: 16px; height: 16px; border-radius: 50%;
  background: #2563eb; color: white;
  display: grid; place-items: center;
  font-size: 10px; font-weight: 700;
}

/* ── Cards ──────────────────────────────────────────────────── */

.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 18px 22px;
  margin-bottom: 16px;
  box-shadow: 0 1px 2px rgba(15,23,42,.04);
}
.card h2 {
  font-size: 15.5px;
  font-weight: 700;
  color: var(--heading);
  margin: 0 0 4px;
  letter-spacing: -.1px;
}
.card .sub {
  font-size: 12.5px; color: var(--text-mute);
  margin-bottom: 14px;
}
.card-head {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 12px;
}

.field-label {
  font-size: 10.5px; color: var(--text-mute);
  text-transform: uppercase; letter-spacing: .8px;
  font-weight: 700; margin: 0 0 6px 2px;
}

/* Segmented control (Halbtag) */
.seg-row { display: flex; gap: 0; }
.seg {
  flex: 1; padding: 8px 14px;
  text-align: center;
  border: 1px solid var(--border);
  background: var(--surface);
  font-size: 13px; color: var(--text-dim);
  cursor: pointer; user-select: none;
  font-weight: 500;
}
.seg:first-child  { border-radius: 8px 0 0 8px; }
.seg:last-child   { border-radius: 0 8px 8px 0; }
.seg + .seg { border-left: none; }
.seg.active {
  background: var(--surface-2);
  color: var(--heading);
  font-weight: 600;
  box-shadow: inset 0 0 0 1px var(--border-2);
}

/* Day chips */
.day-chip-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.day-chip {
  padding: 7px 14px;
  border-radius: 8px;
  border: 1px solid var(--border);
  background: var(--surface);
  font-size: 13px; font-weight: 600; color: var(--text-dim);
  cursor: pointer; user-select: none;
  min-width: 44px; text-align: center;
}
.day-chip.active {
  background: var(--accent-soft);
  border-color: rgba(37,99,235,.35);
  color: var(--accent);
}
.day-link {
  color: var(--text-mute); font-size: 12.5px;
  background: transparent; border: none;
  padding: 4px 6px; cursor: pointer;
}
.day-link:hover { color: var(--text-dim); }

/* Empty state */
.empty-state {
  padding: 40px 18px; text-align: center;
  color: var(--text-mute); font-size: 13px;
}
.empty-state .icon {
  width: 44px; height: 44px; margin: 0 auto 10px;
  border-radius: 12px; background: var(--surface-2);
  border: 1px solid var(--border);
  display: grid; place-items: center;
  color: var(--text-mute); font-size: 20px;
}
.empty-state .ttl { color: var(--text); font-weight: 600; margin-bottom: 4px; font-size: 13px; }

/* Entries list (right card) */
.entries-list { display: flex; flex-direction: column; gap: 8px; }
.entry-row {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 9px;
  padding: 9px 12px;
  display: flex; align-items: center; gap: 10px;
  font-size: 13px;
}
.entry-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.entry-dot.KRANK  { background: var(--err); }
.entry-dot.FERIEN { background: var(--warn); }
.entry-dot.TERMIN { background: #06b6d4; }
.entry-dot.HO     { background: var(--accent); }
.entry-dot.CUSTOM { background: var(--text-mute); }
.entry-name { font-weight: 600; color: var(--text); }
.entry-meta { color: var(--text-dim); font-size: 12.5px; margin-left: 4px; }
.entry-note { color: #b45309; font-style: italic; margin-left: 6px; font-size: 12px; }

/* ── Banner ─────────────────────────────────────────────────── */

.banner {
  display: flex; align-items: flex-start; gap: 10px;
  padding: 12px 16px; border-radius: 10px;
  margin-bottom: 14px;
  border: 1px solid;
  font-size: 13.5px;
}
.banner .check {
  width: 22px; height: 22px; border-radius: 50%;
  display: grid; place-items: center; flex-shrink: 0;
  font-weight: 700; font-size: 13px;
}
.banner-ok    { background: var(--ok-soft);  border-color: rgba(22,163,74,.3); color: #14532d; }
.banner-ok    .check { background: var(--ok); color: white; }
.banner-warn  { background: var(--warn-soft); border-color: rgba(217,119,6,.3); color: #78350f; }
.banner-warn  .check { background: var(--warn); color: white; }
.banner-error { background: var(--err-soft); border-color: rgba(220,38,38,.3); color: #7f1d1d; }
.banner-error .check { background: var(--err); color: white; }
.banner h4 { margin: 0 0 4px; font-size: 13px; font-weight: 700; }
.banner ul { margin: 4px 0 0; padding-left: 18px; font-size: 12.5px; }

/* ── Wochenplan Tabelle ─────────────────────────────────────── */

.plan-wrap {
  overflow-x: auto;
  border-radius: 10px;
  border: 1px solid var(--border);
  background: var(--surface);
}
table.plan {
  width: 100%; border-collapse: separate; border-spacing: 0;
  font-size: 11.5px;
  font-family: -apple-system, "Segoe UI", "Inter", sans-serif;
  min-width: 1100px;
  background: var(--surface);
}
table.plan th, table.plan td {
  padding: 6px 6px;
  border-right: 1px solid var(--border);
  border-bottom: 1px solid var(--border);
  vertical-align: middle;
}
table.plan th {
  background: var(--surface-2);
  color: var(--text-dim);
  font-weight: 700; font-size: 10.5px;
  letter-spacing: .6px; text-transform: uppercase;
  text-align: center;
}
table.plan th.team-head {
  text-align: left; padding: 10px 14px;
  vertical-align: top;
  background: var(--surface);
  border-right: 1px solid var(--border);
}
table.plan th.team-head .ttl { color: var(--text-dim); font-size: 10.5px; font-weight: 700; letter-spacing:.6px; }
table.plan th.team-head .legend {
  margin-top: 4px; color: var(--text-mute);
  font-size: 10.5px; font-weight: 500;
  text-transform: none; letter-spacing: 0;
}
table.plan th.day-head {
  background: var(--surface);
  padding: 10px 8px;
  border-bottom: 1px solid var(--border);
}
table.plan th.day-head .day-name {
  display: block;
  color: var(--text-dim); font-weight: 700; font-size: 11px;
  letter-spacing: .8px; text-transform: uppercase;
}
.tv-pill {
  display: inline-flex; align-items: center; gap: 6px;
  margin-top: 6px;
  padding: 3px 9px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 999px;
  font-size: 11px; font-weight: 600;
  color: var(--text);
  text-transform: none; letter-spacing: 0;
}
.tv-tag {
  background: #e0e7ff; color: #3730a3;
  font-size: 9.5px; font-weight: 800;
  padding: 1px 5px; border-radius: 4px;
  letter-spacing: .4px;
}
.tv-pill.open .tv-name { color: var(--warn); }

table.plan th.slot-head {
  background: var(--surface-2);
  font-size: 9.5px; padding: 5px 4px;
  color: var(--text-mute); font-weight: 600;
}

table.plan td.name-cell {
  background: var(--surface);
  padding: 6px 12px;
  text-align: left;
  font-size: 12px;
  color: var(--text); font-weight: 500;
  white-space: nowrap;
  position: sticky; left: 0;
}
table.plan td.pct-cell {
  background: var(--surface);
  text-align: right; padding-right: 14px;
  color: var(--text-mute);
  font-size: 11px; font-weight: 600;
}

table.plan tbody tr:nth-child(even) td.name-cell,
table.plan tbody tr:nth-child(even) td.pct-cell {
  background: var(--surface-2);
}

td.task-cell {
  text-align: center;
  padding: 5px 4px;
  background: var(--surface);
}
table.plan tbody tr:nth-child(even) td.task-cell {
  background: var(--surface-2);
}

.pill {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 11px; font-weight: 600;
  font-feature-settings: "tnum";
  letter-spacing: .2px;
  min-width: 64px;
  text-align: center;
}
.pill-phc {
  background: var(--phc) !important;
  color: white !important;
  font-weight: 700;
}
.pill-empty { color: transparent; }
.pill-termin { background: #fef3c7; color: #78350f; }

/* ── Wochenaufgaben Strip ────────────────────────────────────── */

.wa-strip {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
  margin-top: 14px;
}
.wa-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px 16px;
  display: flex; align-items: center; gap: 12px;
}
.wa-icon {
  width: 32px; height: 32px; border-radius: 8px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  display: grid; place-items: center;
  color: var(--text-dim); font-size: 14px;
}
.wa-meta .lbl {
  font-size: 10px; color: var(--text-mute);
  text-transform: uppercase; letter-spacing: .8px;
  font-weight: 700;
}
.wa-meta .val {
  font-size: 14px; font-weight: 700; color: var(--heading);
  margin-top: 2px;
}

/* Streamlit utility tweaks */
.stMultiSelect [data-baseweb="tag"] {
  background: var(--accent-soft) !important;
  color: var(--accent) !important;
}

/* Hide the form submit hint */
.stForm > [data-testid="stFormSubmitButton"] { margin-top: 6px; }

hr { margin: 10px 0; border-color: var(--border); }

/* Section titles inside cards */
.sec-title {
  font-size: 12px; font-weight: 700; color: var(--text-dim);
  letter-spacing: .8px; text-transform: uppercase;
  margin: 0 0 10px;
}

/* Generated content compactness */
.element-container { margin-bottom: 0.4rem; }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ── Hilfen ─────────────────────────────────────────────────────────────

def get_week_info(offset: int = 0) -> dict:
    monday = get_next_monday() + timedelta(weeks=offset)
    kw = monday.isocalendar()[1]
    friday = monday + timedelta(days=4)
    is_even = kw % 2 == 0
    return {
        "monday": monday,
        "friday": friday,
        "kw": kw,
        "is_even": is_even,
        "monday_str": monday.strftime("%d.%m.%Y"),
        "friday_str": friday.strftime("%d.%m.%Y"),
        "two_week_notes": (
            ["Linda: Freitag frei", "Isaura: Mittwoch frei", "Corinne: Freitag frei"]
            if is_even else
            ["Linda, Isaura, Corinne normal verfügbar"]
        ),
    }


def pill_for_task(task: str) -> tuple[str, str]:
    if not task:
        return "transparent", "transparent"
    if task.startswith("*"):
        return "#fef3c7", "#78350f"
    if task in TASK_PILL:
        return TASK_PILL[task]
    for key in sorted(TASK_PILL.keys(), key=len, reverse=True):
        if task.startswith(key):
            return TASK_PILL[key]
    return "#f1f5f9", "#334155"


def check_rule_impact(absences, employees):
    warnings = []
    for name, day, slot in absences:
        tag = f"{name} {DAYS[day]} {SLOTS[slot]}"
        if name == "Brigitte" and day == 0:
            warnings.append(f"{tag}: Brigitte hat normalerweise ERF9 am Montag")
        if name == "Dragi" and day == 1:
            warnings.append(f"{tag}: Dragi hat normalerweise ERF9 am Dienstag")
        if name == "Dipiga" and day == 1 and slot == 0:
            warnings.append(f"{tag}: Dipiga hat TAGES PA am Dienstag VM")
        if name == "Dipiga" and day == 3 and slot == 1:
            warnings.append(f"{tag}: Dipiga hat KGS am Donnerstag NM")
        if name == "Martina" and day in [1, 3] and slot == 0:
            warnings.append(f"{tag}: Martina hat KSC Spez. am {DAYS[day]} VM")
        if name == "Florence" and day == 1:
            warnings.append(f"{tag}: Florence macht normalerweise PO am Dienstag")
        if name == "Maria B." and day == 0 and slot == 0:
            warnings.append(f"{tag}: Maria B. hat normalerweise HO am Montag VM")
        if name in ["Jesika", "Dipiga"] and day == 2 and slot == 1:
            warnings.append(f"{tag}: {name} könnte ERF5 (Labor) am Mi NM haben")

    for day in range(5):
        for slot in range(2):
            absent_count = sum(1 for n, d, s in absences if d == day and s == slot)
            available_count = sum(
                1 for n in employees if employees[n].is_available(day, slot)
            ) - absent_count
            target_tel = 4 if (day == 0 and slot == 0) else 3
            min_needed = target_tel + 2
            if available_count < min_needed:
                warnings.append(
                    f"KRITISCH {DAYS[day]} {SLOTS[slot]}: nur {available_count} verfügbar "
                    f"– mind. {min_needed} benötigt"
                )
            elif available_count < min_needed + 3:
                warnings.append(
                    f"Knapp: {DAYS[day]} {SLOTS[slot]} – nur {available_count} Personen"
                )
    return warnings


def type_to_key(t: str) -> str:
    return ABSENCE_TYPES.get(t, "CUSTOM")


def html_escape(s: str) -> str:
    if s is None:
        return ""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# ── Session-State ──────────────────────────────────────────────────────

st.session_state.setdefault("week_offset", 0)
st.session_state.setdefault("entries", [])
st.session_state.setdefault("result", None)
st.session_state.setdefault("step", 1)              # 1 = Planung, 2 = Ergebnis
st.session_state.setdefault("sel_days", set())
st.session_state.setdefault("sel_slot", "Ganzer Tag")
st.session_state.setdefault("f_type", "Krank")
st.session_state.setdefault("f_note", "")
st.session_state.setdefault("f_weeks", 1)


def _init_team() -> None:
    """Seed the editable team list from the scheduler's defaults."""
    if "team" in st.session_state and st.session_state.team:
        return
    defaults = _ORIG_CREATE_EMPLOYEES(1)
    seed = []
    for short in EMPLOYEE_LIST:
        if short not in defaults:
            continue
        emp = defaults[short]
        seed.append({
            "short": short,
            "name": emp.name,
            "pct": emp.pct,
            "active": True,
            "is_default": True,
        })
    st.session_state.team = seed


_init_team()


# ── Apply pending widget clears (must happen BEFORE widgets render) ───
if st.session_state.pop("_pending_clear_note", False):
    st.session_state["f_note"] = ""
if st.session_state.pop("_pending_clear_new_person", False):
    for k in ("new_name", "new_short"):
        st.session_state[k] = ""


def active_shorts() -> list[str]:
    return [e["short"] for e in st.session_state.team if e.get("active", True)]


# Default for the Mitarbeiterin dropdown — first active member
if "f_employee" not in st.session_state:
    a = active_shorts()
    st.session_state["f_employee"] = a[0] if a else EMPLOYEE_LIST[0]
elif st.session_state["f_employee"] not in active_shorts():
    a = active_shorts()
    st.session_state["f_employee"] = a[0] if a else EMPLOYEE_LIST[0]


# ── Plan-Generierung ───────────────────────────────────────────────────

def run_generation() -> None:
    cur_info = get_week_info(st.session_state.week_offset)
    kw = cur_info["kw"]
    monday = cur_info["monday"]

    employees = create_employees(kw)
    overrides: dict = {}
    extra_notes: dict = {}
    absences: list = []

    for entry in st.session_state.entries:
        if entry["name"] not in employees:
            continue
        # Only apply the entry to the configured consecutive week range.
        # Legacy entries without kw_start fall back to "current week only".
        kw_start = entry.get("kw_start", kw)
        weeks = max(1, int(entry.get("weeks", 1)))
        if not (kw_start <= kw < kw_start + weeks):
            continue
        task_code = ABSENCE_TYPES.get(entry["type"], "CUSTOM")
        note = (entry.get("note") or "").strip()

        if task_code == "TERMIN":
            label = f"*{note}" if note else "*Termin"
        elif task_code == "CUSTOM":
            label = (entry.get("task") or "").strip() or "HO"
        else:
            label = task_code

        for day in entry["days"]:
            for slot in entry["slots"]:
                if employees[entry["name"]].is_available(day, slot):
                    overrides[(entry["name"], day, slot)] = label
                    if note and task_code in ("KRANK", "FERIEN", "HO"):
                        extra_notes[(entry["name"], day, slot)] = note
                    if task_code in ("KRANK", "FERIEN") or label.startswith("*"):
                        absences.append((entry["name"], day, slot))

    warnings = check_rule_impact(absences, employees) if absences else []

    random.seed(kw)
    sched = build_schedule(kw, monday, overrides, state_file=str(STATE_FILE))
    for key, n in extra_notes.items():
        sched.notes[key] = n

    issues = validate_schedule(sched)

    filename = f"Arbeitsplan_KW{kw}_{monday.strftime('%Y%m%d')}.xlsx"
    output_path = OUTPUT_DIR / filename
    write_excel(sched, str(output_path))
    with open(output_path, "rb") as fh:
        excel_bytes = fh.read()

    coverage = []
    for day in range(5):
        for slot in range(2):
            target_tel = 4 if (day == 0 and slot == 0) else 3
            coverage.append({
                "day": DAYS[day],
                "slot": SLOTS[slot],
                "tel": sched.tel_count[day][slot],
                "tel_target": target_tel,
                "abkl": sched.abkl_count[day][slot],
                "abkl_target": 2,
                "tel_ok": sched.tel_count[day][slot] >= target_tel,
                "abkl_ok": sched.abkl_count[day][slot] >= 2,
            })

    tagesverantwortung = [
        sched.tagesverantwortung.get(d, "?") for d in range(5)
    ]
    phc_liste = [sched.phc_liste.get(d, "?") for d in range(5)]

    rows = []
    team_order = [e["short"] for e in st.session_state.team if e.get("active", True)]
    for short in team_order:
        if short not in sched.employees:
            continue
        emp = sched.employees[short]
        cells = []
        for day in range(5):
            for slot in range(2):
                task = sched.get_task(short, day, slot)
                is_phc = slot == 1 and sched.phc_liste.get(day) == short
                cells.append({"task": task, "is_phc": is_phc})
        rows.append({
            "full_name": emp.name,
            "short": short,
            "pct": emp.pct,
            "cells": cells,
        })

    st.session_state.result = {
        "filename": filename,
        "excel_bytes": excel_bytes,
        "kw": kw,
        "warnings": warnings,
        "issues": issues,
        "coverage": coverage,
        "tagesverantwortung": tagesverantwortung,
        "phc_liste": phc_liste,
        "wochenaufgaben": {
            "direkt": sched.wochenaufgabe_direkt,
            "onb": sched.wochenaufgabe_onb,
            "btm": sched.wochenaufgabe_btm,
        },
        "rows": rows,
    }
    # Keep the live Schedule + output path so the manual editor can mutate
    # cells in place and rewrite the Excel without regenerating.
    st.session_state.sched = sched
    st.session_state.output_path = str(output_path)
    st.session_state.step = 2


# ════════════════════════════════════════════════════════════════════════
#  TOP HEADER
# ════════════════════════════════════════════════════════════════════════

info = get_week_info(st.session_state.week_offset)

st.markdown('<div class="topbar">', unsafe_allow_html=True)
hcol = st.columns([2.6, 2.4, 2.6, 1.0, 1.6])

# Brand
with hcol[0]:
    st.markdown(
        """
        <div class="brand">
          <div class="brand-icon">📅</div>
          <div>
            <div class="brand-title">Arbeitsplan</div>
            <div class="brand-sub">KSC · Team-Scheduling</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Stepper (clickable)
with hcol[1]:
    s1, s2 = st.columns([1, 1])
    with s1:
        if st.button(
            ("● 1   Planung" if st.session_state.step == 1 else "1   Planung"),
            key="step1", use_container_width=True,
            type=("primary" if st.session_state.step == 1 else "secondary"),
        ):
            st.session_state.step = 1
            st.rerun()
    with s2:
        result_avail = st.session_state.result is not None
        label2 = ("● 2   Ergebnis" if st.session_state.step == 2 else "2   Ergebnis")
        if st.button(
            label2, key="step2", use_container_width=True,
            type=("primary" if st.session_state.step == 2 else "secondary"),
            disabled=not result_avail,
        ):
            st.session_state.step = 2
            st.rerun()

# Week navigation
with hcol[2]:
    w1, w2, w3 = st.columns([1, 4, 1])
    with w1:
        if st.button("‹", key="wk_prev", help="Woche zurück", use_container_width=True):
            if st.session_state.week_offset > 0:
                st.session_state.week_offset -= 1
                if st.session_state.step == 2 and st.session_state.result is not None:
                    try:
                        run_generation()
                    except Exception as exc:
                        st.error(f"Fehler beim Generieren: {exc}")
                st.rerun()
    with w2:
        st.markdown(
            f"""
            <div class="kw-pill" style="justify-content:center;">
              <div class="kw-center">
                <span class="kw-big">KW {info['kw']}</span>
                <span class="kw-range">{info['monday_str']} – {info['friday_str']}</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with w3:
        if st.button("›", key="wk_next", help="Woche vor", use_container_width=True):
            st.session_state.week_offset += 1
            if st.session_state.step == 2 and st.session_state.result is not None:
                try:
                    run_generation()
                except Exception as exc:
                    st.error(f"Fehler beim Generieren: {exc}")
            st.rerun()

# KW badge
with hcol[3]:
    badge = "GERADE KW" if info["is_even"] else "UNGERADE KW"
    st.markdown(
        f'<div style="text-align:center; padding-top:10px;">'
        f'<span class="badge-kw">{badge}</span></div>',
        unsafe_allow_html=True,
    )

# Primary CTA — generate
with hcol[4]:
    cta_label = ("⬇ Excel" if st.session_state.step == 2 and st.session_state.result
                 else "→ Plan generieren")
    if st.session_state.step == 2 and st.session_state.result:
        st.download_button(
            "⬇  Excel herunterladen",
            data=st.session_state.result["excel_bytes"],
            file_name=st.session_state.result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="cta_download",
        )
    else:
        if st.button("→  Plan generieren", key="cta_generate",
                     type="primary", use_container_width=True):
            with st.spinner("Plan wird generiert …"):
                try:
                    run_generation()
                    st.rerun()
                except Exception as exc:
                    st.error(f"Fehler beim Generieren: {exc}")

st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════
#  STATS STRIP
# ════════════════════════════════════════════════════════════════════════

mo = info["monday_str"].replace(".", "."*1)
notes_str = " · ".join(info["two_week_notes"])
st.markdown(
    f"""
    <div class="stats-strip">
      <div class="stat-item">
        <div class="lbl">Kalenderwoche</div>
        <div class="val">{info['kw']}</div>
      </div>
      <div class="stat-divider"></div>
      <div class="stat-item">
        <div class="lbl">Montag</div>
        <div class="val">{info['monday_str']}</div>
      </div>
      <div class="stat-divider"></div>
      <div class="stat-item">
        <div class="lbl">Freitag</div>
        <div class="val">{info['friday_str']}</div>
      </div>
      <div class="stat-divider"></div>
      <div class="stat-item">
        <div class="lbl">Team</div>
        <div class="val">18 <small>Personen</small></div>
      </div>
      <div class="info-pill"><span class="ico">i</span><span>{html_escape(notes_str)}</span></div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ════════════════════════════════════════════════════════════════════════
#  STEP 1 — PLANUNG
# ════════════════════════════════════════════════════════════════════════

def render_team_management() -> None:
    """Editable team list — pension, active flag, add/remove persons."""
    n_active = sum(1 for e in st.session_state.team if e.get("active", True))
    n_total = len(st.session_state.team)
    label = f"👥  Team verwalten · {n_active}/{n_total} aktiv  (Pensum & Personen anpassen)"
    with st.expander(label, expanded=False):
        st.caption(
            "Pensum pro Person via Dropdown anpassen, Personen aktivieren / "
            "deaktivieren, oder neue Personen hinzufügen. Änderungen wirken "
            "auf die nächste Plan-Generierung."
        )

        # Header row
        h = st.columns([3, 1.5, 1.2, 0.8, 0.6])
        h[0].markdown("**Name**")
        h[1].markdown("**Kurzform**")
        h[2].markdown("**Pensum**")
        h[3].markdown("**Aktiv**")
        h[4].markdown("&nbsp;", unsafe_allow_html=True)

        pct_options = list(range(10, 110, 10))
        for idx, entry in enumerate(st.session_state.team):
            r = st.columns([3, 1.5, 1.2, 0.8, 0.6])
            tag = "" if entry["is_default"] else "  ·  ✨ neu"
            r[0].markdown(f"{entry['name']}{tag}")
            r[1].markdown(f"`{entry['short']}`")
            cur = entry["pct"]
            try:
                idx_pct = pct_options.index(cur)
            except ValueError:
                idx_pct = pct_options.index(min(pct_options, key=lambda v: abs(v - cur)))
            new_pct = r[2].selectbox(
                "Pensum", pct_options, index=idx_pct,
                key=f"team_pct_{idx}", label_visibility="collapsed",
                format_func=lambda v: f"{v}%",
            )
            new_active = r[3].checkbox(
                "Aktiv", value=entry["active"],
                key=f"team_act_{idx}", label_visibility="collapsed",
            )
            entry["pct"] = new_pct
            entry["active"] = new_active

            if not entry["is_default"]:
                if r[4].button("✕", key=f"team_rm_{idx}", help="Person entfernen"):
                    st.session_state.team.pop(idx)
                    st.rerun()
            else:
                r[4].markdown("&nbsp;", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**Neue Person hinzufügen**")
        nc = st.columns([3, 1.5, 1.2, 1.5])
        nc[0].text_input(
            "Voller Name", key="new_name",
            placeholder="z.B. Mustermann Anna",
            label_visibility="collapsed",
        )
        nc[1].text_input(
            "Kurzform", key="new_short",
            placeholder="Anna",
            label_visibility="collapsed",
        )
        nc[2].selectbox(
            "Pensum", pct_options, index=9,  # 100%
            key="new_pct", label_visibility="collapsed",
            format_func=lambda v: f"{v}%",
        )
        if nc[3].button("＋  Person hinzufügen", key="new_add",
                        type="primary", use_container_width=True):
            name = (st.session_state.get("new_name") or "").strip()
            short = (st.session_state.get("new_short") or "").strip()
            if not name or not short:
                st.warning("Bitte Voller Name und Kurzform angeben.")
            elif any(e["short"] == short for e in st.session_state.team):
                st.warning(f"Kurzform „{short}\" ist bereits vergeben.")
            else:
                st.session_state.team.append({
                    "short": short,
                    "name": name,
                    "pct": int(st.session_state.get("new_pct", 100)),
                    "active": True,
                    "is_default": False,
                })
                st.session_state["_pending_clear_new_person"] = True
                st.rerun()

        st.markdown("---")
        rcol1, rcol2 = st.columns([1, 4])
        with rcol1:
            if st.button("↺  Auf Standard zurücksetzen", key="team_reset"):
                st.session_state.pop("team", None)
                _init_team()
                st.rerun()
        with rcol2:
            st.caption(
                "Hinweis: Neue Personen erhalten standardmässig volle "
                "Wochenverfügbarkeit und alle Basis-Skills (Schalter, "
                "Scanning, ONB). Die fest verdrahteten Spezial-Tasks "
                "(z.B. Brigitte = ERF9 Mo, Dipiga = TAGES PA Di) bleiben "
                "an die Originalnamen gebunden."
            )


def render_planung() -> None:
    render_team_management()
    left, right = st.columns([1.25, 1])

    # ─── Left card: Form ─────────────────────────────────────────
    with left:
        st.markdown(
            """
            <div class="card">
              <h2>Abwesenheit oder Sonderwunsch</h2>
              <div class="sub">Krank · Ferien · Termin · Home Office · Eigene Regel</div>
            """,
            unsafe_allow_html=True,
        )

        f1, f2 = st.columns(2)
        with f1:
            st.markdown('<div class="field-label">Mitarbeiterin</div>',
                        unsafe_allow_html=True)
            st.selectbox(
                "Mitarbeiterin", active_shorts(),
                key="f_employee", label_visibility="collapsed",
            )
        with f2:
            st.markdown('<div class="field-label">Grund</div>',
                        unsafe_allow_html=True)
            st.selectbox(
                "Grund", list(ABSENCE_TYPES.keys()),
                key="f_type", label_visibility="collapsed",
            )

        # Halbtag — segmented control
        st.markdown('<div class="field-label" style="margin-top:8px;">Halbtag</div>',
                    unsafe_allow_html=True)
        seg_cols = st.columns(3)
        seg_options = ["Ganzer Tag", "Vormittag", "Nachmittag"]
        for i, lbl in enumerate(seg_options):
            with seg_cols[i]:
                is_active = st.session_state.sel_slot == lbl
                if st.button(
                    lbl, key=f"seg_{i}",
                    type=("primary" if is_active else "secondary"),
                    use_container_width=True,
                ):
                    st.session_state.sel_slot = lbl
                    st.rerun()

        st.markdown(
            '<div class="field-label" style="margin-top:10px;">Notiz '
            '<span style="font-weight:500; text-transform:none; letter-spacing:0; '
            'color:var(--text-mute);">optional</span></div>',
            unsafe_allow_html=True,
        )
        st.text_input(
            "Notiz", key="f_note",
            placeholder="z.B. Arzttermin 14:30, oder Task-Code für Custom",
            label_visibility="collapsed",
        )

        # Tage chips — day buttons on row 1
        st.markdown('<div class="field-label" style="margin-top:10px;">Tage</div>',
                    unsafe_allow_html=True)
        d_cols = st.columns(5)
        for i, lbl in enumerate(DAY_SHORTS):
            with d_cols[i]:
                is_active = i in st.session_state.sel_days
                if st.button(
                    lbl, key=f"day_{i}",
                    type=("primary" if is_active else "secondary"),
                    use_container_width=True,
                ):
                    if i in st.session_state.sel_days:
                        st.session_state.sel_days.remove(i)
                    else:
                        st.session_state.sel_days.add(i)
                    st.rerun()

        # Anzahl Wochen (consecutive) — defaults to 1
        st.markdown(
            '<div class="field-label" style="margin-top:10px;">Anzahl '
            'Wochen <span style="font-weight:500; text-transform:none; '
            'letter-spacing:0; color:var(--text-mute);">aufeinander '
            'folgend, ab dieser KW</span></div>',
            unsafe_allow_html=True,
        )
        st.number_input(
            "Anzahl Wochen", min_value=1, max_value=52, step=1,
            key="f_weeks", label_visibility="collapsed",
        )

        # Alle / Keine / Hinzufügen on row 2
        a_cols = st.columns([1, 1, 4, 2])
        with a_cols[0]:
            if st.button("Alle", key="day_all", type="secondary",
                         use_container_width=True):
                st.session_state.sel_days = {0, 1, 2, 3, 4}
                st.rerun()
        with a_cols[1]:
            if st.button("Keine", key="day_none", type="secondary",
                         use_container_width=True):
                st.session_state.sel_days = set()
                st.rerun()
        with a_cols[3]:
            if st.button("＋  Hinzufügen", key="add_entry",
                         type="primary", use_container_width=True):
                if not st.session_state.sel_days:
                    st.warning("Bitte mindestens einen Tag auswählen.")
                else:
                    days = sorted(st.session_state.sel_days)
                    sl = st.session_state.sel_slot
                    slots = [0, 1] if sl == "Ganzer Tag" else [0] if sl == "Vormittag" else [1]
                    note = (st.session_state.f_note or "").strip()
                    f_type = st.session_state.f_type
                    task = note if (f_type == "Custom" and note) else (
                        "HO" if f_type == "Custom" else ""
                    )
                    cur_kw = get_week_info(st.session_state.week_offset)["kw"]
                    weeks_n = max(1, int(st.session_state.f_weeks or 1))
                    st.session_state.entries.append({
                        "name": st.session_state.f_employee,
                        "type": f_type,
                        "days": days,
                        "slots": slots,
                        "note": note,
                        "task": task,
                        "kw_start": cur_kw,
                        "weeks": weeks_n,
                    })
                    # Pending-clear pattern: Streamlit forbids mutating a
                    # widget-bound key after it was instantiated this run.
                    st.session_state.sel_days = set()
                    st.session_state["_pending_clear_note"] = True
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

    # ─── Right card: Entries ─────────────────────────────────────
    with right:
        n = len(st.session_state.entries)
        st.markdown(
            f"""
            <div class="card">
              <div class="card-head">
                <div>
                  <h2 style="margin-bottom:0;">Einträge für diese Woche</h2>
                  <div class="sub" style="margin-bottom:0;">{n} {"Eintrag" if n == 1 else "Einträge"}</div>
                </div>
              </div>
            """,
            unsafe_allow_html=True,
        )

        if n == 0:
            st.markdown(
                """
                <div class="empty-state">
                  <div class="icon">📅</div>
                  <div class="ttl">Keine Abwesenheiten erfasst.</div>
                  <div>Diese Woche läuft wie geplant.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            cur_kw = info["kw"]
            for idx, e in enumerate(st.session_state.entries):
                type_key = type_to_key(e["type"])
                days_str = ", ".join(DAY_SHORTS[d] for d in e["days"])
                slot_str = (
                    "ganzer Tag" if len(e["slots"]) == 2
                    else "Vormittag" if e["slots"][0] == 0
                    else "Nachmittag"
                )
                note_html = (
                    f'<span class="entry-note">· „{html_escape(e["note"])}"</span>'
                    if e["note"] else ""
                )
                kw_start = e.get("kw_start", cur_kw)
                weeks_n = max(1, int(e.get("weeks", 1)))
                kw_end = kw_start + weeks_n - 1
                if weeks_n == 1:
                    kw_str = f"KW {kw_start}"
                else:
                    kw_str = f"KW {kw_start}–{kw_end}"
                is_active = kw_start <= cur_kw <= kw_end
                opacity = "" if is_active else 'style="opacity:0.45;"'
                erow_l, erow_r = st.columns([12, 1])
                with erow_l:
                    st.markdown(
                        f"""
                        <div class="entry-row" {opacity}>
                          <span class="entry-dot {type_key}"></span>
                          <span class="entry-name">{html_escape(e['name'])}</span>
                          <span class="entry-meta">· {html_escape(e['type'])} · {days_str} ({slot_str}) · {kw_str}</span>
                          {note_html}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with erow_r:
                    if st.button("✕", key=f"rm_{idx}", help="Entfernen",
                                 use_container_width=True):
                        st.session_state.entries.pop(idx)
                        st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)


def render_plan_editor(res: dict) -> None:
    """Editable Wochenplan: per-cell Dropdowns + freier Text via Custom-Eingabe.

    Mutates st.session_state.sched in place when the user applies edits and
    re-writes the Excel file so the download stays in sync. The visible
    HTML table above is rebuilt from res['rows'] on the next rerun.
    """
    sched = st.session_state.get("sched")
    if sched is None:
        return

    slot_cols = [
        f"{DAY_SHORTS[d]} {'VM' if s == 0 else 'NM'}"
        for d in range(5) for s in range(2)
    ]

    with st.expander("✏️  Plan bearbeiten · Felder per Dropdown ändern", expanded=False):
        st.caption(
            "Pro Halbtag eine Tätigkeit auswählen. Leer = Slot frei. "
            "Änderungen werden nach „Übernehmen“ in den Plan und in die "
            "Excel-Datei geschrieben."
        )

        df_rows = []
        for r in res["rows"]:
            row = {"Name": r["full_name"], "%": f"{r['pct']}%"}
            for i, col in enumerate(slot_cols):
                task = r["cells"][i]["task"] or ""
                # Stripped Termin marker („*Arzt“) wird zur reinen Notiz
                if task.startswith("*"):
                    task = task[1:]
                row[col] = task
            df_rows.append(row)
        df = pd.DataFrame(df_rows, columns=["Name", "%"] + slot_cols)

        col_config: dict = {
            "Name": st.column_config.TextColumn("Name", disabled=True, width="medium"),
            "%": st.column_config.TextColumn("%", disabled=True, width="small"),
        }
        for col in slot_cols:
            col_config[col] = st.column_config.SelectboxColumn(
                col,
                options=TASK_OPTIONS,
                required=False,
                width="small",
                help="Tätigkeit für diesen Halbtag",
            )

        edited = st.data_editor(
            df,
            column_config=col_config,
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            key=f"plan_editor_{res['kw']}",
        )

        c1, c2 = st.columns([1, 4])
        with c1:
            apply = st.button(
                "✓  Änderungen übernehmen",
                key=f"apply_edits_{res['kw']}",
                type="primary",
                use_container_width=True,
            )
        with c2:
            st.caption(
                "Nach „Übernehmen“: Tabelle oben + Excel-Download werden "
                "aktualisiert. Wochenwechsel via ‹ / › regeneriert komplett neu."
            )

        if not apply:
            return

        # Apply edits: mutate sched.schedule and res['rows'] in lock-step.
        changes = 0
        for ridx, r in enumerate(res["rows"]):
            short = r["short"]
            if short not in sched.employees:
                continue
            for cidx, col in enumerate(slot_cols):
                new_val = edited.at[ridx, col]
                if new_val is None or (isinstance(new_val, float) and pd.isna(new_val)):
                    new_val = ""
                new_val = str(new_val)
                old_val = r["cells"][cidx]["task"] or ""
                # Original sternstart-Notiz wieder ausblenden
                if old_val.startswith("*"):
                    old_val = old_val[1:]
                if new_val == old_val:
                    continue
                d, s = divmod(cidx, 2)
                sched.schedule[short][(d, s)] = new_val
                r["cells"][cidx]["task"] = new_val
                changes += 1

        if changes == 0:
            st.info("Keine Änderungen erkannt.")
            return

        # Rewrite Excel from the mutated Schedule
        output_path = st.session_state.get("output_path")
        if output_path:
            try:
                write_excel(sched, output_path)
                with open(output_path, "rb") as fh:
                    st.session_state.result["excel_bytes"] = fh.read()
            except Exception as exc:
                st.error(f"Excel-Update fehlgeschlagen: {exc}")
                return

        st.success(f"{changes} Änderung(en) übernommen.")
        st.rerun()


# ════════════════════════════════════════════════════════════════════════
#  STEP 2 — ERGEBNIS
# ════════════════════════════════════════════════════════════════════════

def render_ergebnis() -> None:
    res = st.session_state.result
    if not res:
        st.info("Noch kein Plan generiert. Klicke auf **Plan generieren**.")
        return

    # Banner
    if res["issues"]:
        items = "".join(f"<li>{html_escape(i)}</li>" for i in res["issues"])
        st.markdown(
            f'<div class="banner banner-error"><div class="check">!</div>'
            f'<div><h4>Konflikte / Regeln nicht vollständig einhaltbar</h4>'
            f'<ul>{items}</ul></div></div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="banner banner-ok"><div class="check">✓</div>'
            '<div><h4>Alle harten Regeln eingehalten.</h4></div></div>',
            unsafe_allow_html=True,
        )

    if res["warnings"]:
        items = "".join(f"<li>{html_escape(w)}</li>" for w in res["warnings"])
        st.markdown(
            f'<div class="banner banner-warn"><div class="check">!</div>'
            f'<div><h4>Warnungen vor dem Generieren (Regel-Impact)</h4>'
            f'<ul>{items}</ul></div></div>',
            unsafe_allow_html=True,
        )

    # Wochenplan card
    st.markdown(
        """
        <div class="card">
          <div class="card-head">
            <h2 style="margin-bottom:0;">Wochenplan</h2>
          </div>
        """,
        unsafe_allow_html=True,
    )

    parts = ['<div class="plan-wrap"><table class="plan">']

    # ── HEAD ROW 1: Team header + day-heads with TV pills ───────
    parts.append("<thead><tr>")
    parts.append(
        '<th class="team-head" colspan="2" rowspan="2">'
        '<div class="ttl">TEAM</div>'
        '<div class="legend">TV · Tagesverantwortung</div>'
        '</th>'
    )
    for d in range(5):
        tv = res["tagesverantwortung"][d]
        is_open = "OFFEN" in str(tv)
        tv_short = next(
            (r["short"] for r in res["rows"] if r["short"] == tv or r["full_name"] == tv),
            tv,
        )
        # use short name in the TV pill (fits the small space)
        cls = " open" if is_open else ""
        parts.append(
            f'<th class="day-head" colspan="2">'
            f'<span class="day-name">{DAYS[d].upper()}</span>'
            f'<span class="tv-pill{cls}"><span class="tv-tag">TV</span>'
            f'<span class="tv-name">{html_escape(tv_short)}</span></span>'
            f'</th>'
        )
    parts.append("</tr>")

    # ── HEAD ROW 2: VM/NM ──────────────────────────────────────
    parts.append("<tr>")
    for _ in range(5):
        parts.append('<th class="slot-head">VM</th><th class="slot-head">NM</th>')
    parts.append("</tr></thead>")

    # ── BODY ──────────────────────────────────────────────────
    parts.append("<tbody>")
    for r in res["rows"]:
        parts.append("<tr>")
        parts.append(f'<td class="name-cell">{html_escape(r["full_name"])}</td>')
        parts.append(f'<td class="pct-cell">{r["pct"]}%</td>')
        for c in r["cells"]:
            task = c["task"] or ""
            if not task:
                parts.append('<td class="task-cell"></td>')
                continue
            if c["is_phc"]:
                parts.append(
                    f'<td class="task-cell"><span class="pill pill-phc">'
                    f'{html_escape(task)}</span></td>'
                )
            elif task.startswith("*"):
                parts.append(
                    f'<td class="task-cell"><span class="pill pill-termin">'
                    f'{html_escape(task[1:])}</span></td>'
                )
            else:
                bg, fg = pill_for_task(task)
                parts.append(
                    f'<td class="task-cell"><span class="pill" '
                    f'style="background:{bg}; color:{fg};">'
                    f'{html_escape(task)}</span></td>'
                )
        parts.append("</tr>")
    parts.append("</tbody></table></div>")

    # ── Wochenaufgaben strip ──────────────────────────────────
    wa = res["wochenaufgaben"]
    parts.append(
        f"""
        <div class="wa-strip">
          <div class="wa-card">
            <div class="wa-icon">📦</div>
            <div class="wa-meta">
              <div class="lbl">Direkt</div>
              <div class="val">{html_escape(wa['direkt'] or '–')}</div>
            </div>
          </div>
          <div class="wa-card">
            <div class="wa-icon">🆕</div>
            <div class="wa-meta">
              <div class="lbl">ONB</div>
              <div class="val">{html_escape(wa['onb'] or '–')}</div>
            </div>
          </div>
          <div class="wa-card">
            <div class="wa-icon">💊</div>
            <div class="wa-meta">
              <div class="lbl">BTM</div>
              <div class="val">{html_escape(wa['btm'] or '–')}</div>
            </div>
          </div>
        </div>
        """
    )

    st.markdown("".join(parts), unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Editable plan ────────────────────────────────────────────
    render_plan_editor(res)

    # Coverage details (collapsed)
    with st.expander("Besetzung (Detailansicht) · TEL und ABKL pro Halbtag"):
        cov_per_row = 5
        for i in range(0, len(res["coverage"]), cov_per_row):
            row = res["coverage"][i:i + cov_per_row]
            cols = st.columns(len(row))
            for col, c in zip(cols, row):
                tel_color = "#16a34a" if c["tel_ok"] else "#dc2626"
                abkl_color = "#16a34a" if c["abkl_ok"] else "#dc2626"
                col.markdown(
                    f"""
                    <div style="background:#f8fafc; border:1px solid #e3e8f0;
                                border-radius:8px; padding:10px 12px; font-size:12px;">
                      <div style="display:flex; justify-content:space-between; color:#475569;">
                        <strong style="color:#0b1020;">{c['day']}</strong>
                        <span>{c['slot']}</span>
                      </div>
                      <div style="display:flex; justify-content:space-between; margin-top:4px;">
                        <span>TEL</span>
                        <span style="color:{tel_color}; font-weight:700;">{c['tel']}/{c['tel_target']}</span>
                      </div>
                      <div style="display:flex; justify-content:space-between;">
                        <span>ABKL</span>
                        <span style="color:{abkl_color}; font-weight:700;">{c['abkl']}/{c['abkl_target']}</span>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    # PHC-Liste
    with st.expander("PHC-Liste (18 Uhr)"):
        for d, p in enumerate(res["phc_liste"]):
            is_open = "OFFEN" in str(p)
            color = "#d97706" if is_open else "#0b1020"
            st.markdown(
                f'<div style="display:flex; justify-content:space-between; '
                f'padding:6px 12px; border-bottom:1px solid #e3e8f0;">'
                f'<span style="color:#475569;">{DAYS[d]}</span>'
                f'<span style="color:{color}; font-weight:600;">{html_escape(p)}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ── Render aktive Step ─────────────────────────────────────────────────

if st.session_state.step == 1 or st.session_state.result is None:
    render_planung()
else:
    render_ergebnis()
