#!/usr/bin/env python3
"""
KSC Arbeitsplan — Streamlit-Variante
====================================
Spiegelt die FastAPI/HTML-Oberfläche (siehe ksc_arbeitsplan/app/) als
Streamlit-App, damit die Anwendung ohne Docker zum Testen geteilt werden
kann.

Start:
    pip install streamlit openpyxl
    streamlit run streamlit_app.py
"""

from __future__ import annotations

import os
import random
import sys
from datetime import timedelta
from pathlib import Path

import streamlit as st

# ── Kern-Logik importieren ────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
APP_DIR = BASE_DIR / "ksc_arbeitsplan" / "app"
for p in (str(BASE_DIR), str(APP_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

from arbeitskalender import (  # noqa: E402
    DAYS,
    SLOTS,
    build_schedule,
    create_employees,
    get_next_monday,
    validate_schedule,
    write_excel,
)

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

TASK_STYLE = {
    "KTG": "#A9D18E", "KGS": "#A9D18E",
    "ABKL": "#FFC000",
    "TEL": "#92D050",
    "ERF7": "#00B0F0", "ERF7/Q": "#00B0F0", "ERF7/HUB": "#00B0F0",
    "ERF8": "#FF6600", "ERF8/Q": "#FF6600", "ERF8/HUB": "#FF6600",
    "ERF9": "#FF0000", "ERF9/Q": "#FF0000", "ERF9/TEL": "#FF0000",
    "ERF4/SCH": "#C6EFCE",
    "ERF5": "#F4B084",
    "PO": "#D6DCE4",
    "PO/ABKL": "#FFC000",
    "PO/TEL": "#92D050",
    "PO/SCAN": "#A9D18E",
    "HO": "#BDD7EE", "HO/Q": "#BDD7EE",
    "VBZ/Q": "#C5E0B4",
    "KSC Spez.": "#FFFF00",
    "TAGES PA": "#00FFFF", "TAGESPA": "#00FFFF",
    "RX Abo": "#D9D9D9",
    "KRANK": "#FF9999",
    "FERIEN": "#F2F2F2",
    "scanning": "#E2EFDA",
}

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
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
:root {
  --bg:#0b1020; --panel:#141b2d; --panel-2:#1a2338; --elev:#1f2940;
  --border:#273149; --border-2:#324061;
  --text:#e6ebf5; --text-dim:#94a3b8; --text-mute:#64748b; --heading:#ffffff;
  --accent:#6366f1; --accent-2:#8b5cf6;
  --ok:#22c55e; --warn:#f59e0b; --err:#ef4444; --info:#38bdf8;
}
html, body, [data-testid="stAppViewContainer"] {
  background:
    radial-gradient(1200px 600px at 10% -10%, rgba(99,102,241,.08), transparent 60%),
    radial-gradient(900px 500px at 110% 10%, rgba(139,92,246,.06), transparent 60%),
    var(--bg);
  color: var(--text);
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stSidebar"] {
  background: var(--panel);
  border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] * { color: var(--text); }
.brand {
  display:flex; align-items:center; gap:12px;
  padding: 4px 4px 16px; border-bottom:1px solid var(--border);
  margin-bottom:14px;
}
.brand-mark {
  width:42px; height:42px; border-radius:10px;
  background: linear-gradient(135deg, var(--accent), var(--accent-2));
  color:white; font-weight:700; font-size:15px;
  display:grid; place-items:center; letter-spacing:.5px;
  box-shadow: 0 6px 20px rgba(99,102,241,.35);
}
.brand-title { font-size:16px; font-weight:700; color:var(--heading); line-height:1.1; }
.brand-sub { font-size:11px; color:var(--text-dim); letter-spacing:.4px; text-transform:uppercase; }

.page-title {
  font-size:24px; font-weight:700; color:var(--heading);
  margin: 0 0 4px 0; letter-spacing:-.3px;
}
.page-sub { color:var(--text-dim); margin:0; font-size:14px; }

.kw-pill {
  display:inline-flex; flex-direction:column; align-items:center;
  background:var(--panel); border:1px solid var(--border);
  padding:8px 18px; border-radius:10px; min-width:200px;
}
.kw-big { font-size:15px; font-weight:700; color:var(--heading); }
.kw-range { font-size:11px; color:var(--text-dim); }

.card {
  background: var(--panel);
  border:1px solid var(--border);
  border-radius:14px; padding: 18px 22px;
  margin-bottom: 18px;
  box-shadow: 0 1px 2px rgba(0,0,0,.2);
}
.card-title {
  font-size:16px; font-weight:600; color:var(--heading);
  margin:0 0 12px 0; letter-spacing:-.2px;
}
.badge {
  display:inline-block;
  background: rgba(99,102,241,.16);
  color: var(--accent);
  padding: 4px 10px; border-radius: 999px;
  font-size:11.5px; font-weight:600;
}
.stat {
  background: var(--panel-2);
  border:1px solid var(--border);
  border-radius:10px; padding:14px 16px;
}
.stat-label {
  font-size:11px; color:var(--text-mute);
  letter-spacing:.5px; text-transform:uppercase;
  font-weight:600; margin-bottom:6px;
}
.stat-value { font-size:20px; font-weight:700; color:var(--heading); }

.info-strip {
  display:flex; gap:12px; align-items:center;
  background: rgba(56,189,248,.14);
  border:1px solid rgba(56,189,248,.3);
  padding:12px 16px; border-radius:10px;
  margin-top: 12px;
}
.info-strip-icon {
  width:22px; height:22px; border-radius:50%;
  background:var(--info); color:#0b1020;
  display:grid; place-items:center; font-weight:700; font-size:12px;
}
.info-strip-text { font-size:13px; color:var(--text); }

.entry-item {
  background: var(--panel-2);
  border:1px solid var(--border);
  border-radius:8px; padding:10px 14px;
  display:flex; align-items:center; gap:10px;
  font-size:13px; margin-bottom:6px;
}
.entry-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.entry-dot.KRANK  { background:#ef4444; }
.entry-dot.FERIEN { background:#f59e0b; }
.entry-dot.TERMIN { background:#38bdf8; }
.entry-dot.HO     { background:#6366f1; }
.entry-dot.CUSTOM { background:#64748b; }
.entry-name { font-weight:600; color:var(--text); }
.entry-meta { color:var(--text-dim); margin-left:6px; }
.entry-note { color:#fbbf24; font-style:italic; margin-left:6px; }

.entries-empty {
  padding:22px; text-align:center; color:var(--text-mute);
  font-size:13px; background:var(--panel-2);
  border:1px dashed var(--border); border-radius:10px;
}

.banner {
  padding:14px 18px; border-radius:10px;
  margin-bottom:14px; font-size:13.5px; line-height:1.55;
  border:1px solid;
}
.banner h4 { margin:0 0 8px; font-size:13px; color:var(--heading); }
.banner ul { margin:0; padding-left:20px; }
.banner-ok    { background:rgba(34,197,94,.14);  border-color:rgba(34,197,94,.3);  color:#22c55e; font-weight:500; }
.banner-warn  { background:rgba(245,158,11,.14); border-color:rgba(245,158,11,.3); color:#fbbf24; }
.banner-error { background:rgba(239,68,68,.14);  border-color:rgba(239,68,68,.3);  color:#fca5a5; }

.cov-card {
  background:var(--panel-2); border:1px solid var(--border);
  border-radius:8px; padding:12px 14px;
  display:flex; flex-direction:column; gap:6px;
  font-size:12.5px;
}
.cov-head {
  display:flex; justify-content:space-between;
  font-size:12.5px; color:var(--text-dim);
}
.cov-day { color:var(--text); font-weight:600; }
.cov-row { display:flex; justify-content:space-between; align-items:center; }
.cov-badge { padding:2px 8px; border-radius:6px; font-size:11px; font-weight:700; }
.cov-ok { background: rgba(34,197,94,.14); color:#22c55e; }
.cov-bad { background: rgba(239,68,68,.14); color:#ef4444; }

.pill {
  background:var(--panel-2); border:1px solid var(--border);
  border-radius:8px; padding:10px 14px;
  display:flex; justify-content:space-between; align-items:center;
  font-size:13px; margin-bottom:6px;
}
.pill-day { color:var(--text-dim); font-weight:500; }
.pill-val { color:var(--heading); font-weight:600; }
.pill-val.open { color: var(--warn); }

.wa-card {
  background:var(--panel-2);
  border:1px solid var(--border);
  border-left:3px solid var(--accent);
  border-radius:8px; padding:12px 16px;
}
.wa-label {
  font-size:11px; color:var(--text-mute);
  text-transform:uppercase; letter-spacing:.5px;
  font-weight:600; margin-bottom:4px;
}
.wa-val { color:var(--heading); font-weight:600; font-size:14px; }

.schedule-wrap {
  overflow-x:auto;
  border-radius:10px;
  border:1px solid var(--border);
  background:var(--panel-2);
}
table.schedule {
  width:100%; border-collapse:collapse;
  font-size:11.5px; font-family: "JetBrains Mono", Consolas, monospace;
  min-width: 900px;
}
table.schedule th, table.schedule td {
  padding:6px 8px; text-align:center;
  border-right:1px solid var(--border);
  border-bottom:1px solid var(--border);
  white-space:nowrap; color:#0b1020;
}
table.schedule th {
  background:#1e293b; color:var(--text);
  font-weight:600; font-family:"Inter", sans-serif;
  font-size:11px; padding:8px 6px;
}
table.schedule th.day-col {
  background: linear-gradient(135deg, #334155, #1e293b);
  color:white;
}
table.schedule td.name-col {
  text-align:left; background:var(--panel);
  color:var(--text); font-family:"Inter", sans-serif;
  font-size:12px; font-weight:500; padding-left:14px;
}
table.schedule td.pct-col {
  background:var(--panel-2); color:var(--text-dim);
  font-family:"Inter", sans-serif; font-size:11.5px; width:44px;
}
table.schedule td.cell-phc {
  background:#0070C0 !important; color:white !important; font-weight:700;
}
.tiny-hint { color:var(--text-mute); font-size:11.5px; margin-top:8px; }

.section-title {
  font-size:13px; font-weight:600; color:var(--heading);
  letter-spacing:.3px; text-transform:uppercase;
  margin: 18px 0 10px;
}

div.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--accent), var(--accent-2));
  border:none; color:white; font-weight:600;
  box-shadow: 0 8px 22px rgba(99,102,241,.35);
}
div.stButton > button[kind="primary"]:hover {
  filter: brightness(1.08);
  transform: translateY(-1px);
}
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


def style_for_task(task: str) -> tuple[str, str]:
    if not task:
        return "#ffffff", "#cbd5e1"
    if task.startswith("*"):
        return "#facc15", "#1e293b"
    if task in TASK_STYLE:
        return TASK_STYLE[task], "#0b1020"
    for key in sorted(TASK_STYLE.keys(), key=len, reverse=True):
        if task.startswith(key):
            return TASK_STYLE[key], "#0b1020"
    return "#ffffff", "#0b1020"


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

if "week_offset" not in st.session_state:
    st.session_state.week_offset = 0
if "entries" not in st.session_state:
    st.session_state.entries = []
if "result" not in st.session_state:
    st.session_state.result = None


# ── SIDEBAR ────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        """
        <div class="brand">
          <div class="brand-mark">KSC</div>
          <div>
            <div class="brand-title">Arbeitsplan</div>
            <div class="brand-sub">Team-Scheduling</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("**Navigation**")
    st.markdown(
        "- Woche\n- Abwesenheiten\n- Ergebnis",
    )
    st.markdown("---")
    st.caption(f"Version 6.0 · 18 Personen")


# ── TOPBAR ─────────────────────────────────────────────────────────────

info = get_week_info(st.session_state.week_offset)

top_l, top_r = st.columns([3, 2])
with top_l:
    st.markdown(
        '<h1 class="page-title">Wochenplan generieren</h1>'
        '<p class="page-sub">Arbeitsplan für das KSC-Team — Montag bis Freitag</p>',
        unsafe_allow_html=True,
    )

with top_r:
    c1, c2, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("‹", key="wk_prev", help="Woche zurück"):
            if st.session_state.week_offset > 0:
                st.session_state.week_offset -= 1
                st.rerun()
    with c2:
        st.markdown(
            f"""
            <div class="kw-pill">
              <span class="kw-big">KW {info['kw']}</span>
              <span class="kw-range">{info['monday_str']} – {info['friday_str']}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        if st.button("›", key="wk_next", help="Woche vor"):
            st.session_state.week_offset += 1
            st.rerun()


# ── SECTION: WOCHEN-INFO ───────────────────────────────────────────────

st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown(
    f"""
    <div style="display:flex; justify-content:space-between; align-items:center;">
      <h2 class="card-title">Wochen-Kontext</h2>
      <span class="badge">{'gerade KW' if info['is_even'] else 'ungerade KW'}</span>
    </div>
    """,
    unsafe_allow_html=True,
)
sc = st.columns(4)
stats = [
    ("Kalenderwoche", str(info["kw"])),
    ("Montag", info["monday_str"]),
    ("Freitag", info["friday_str"]),
    ("Team", "18"),
]
for col, (lbl, val) in zip(sc, stats):
    col.markdown(
        f'<div class="stat"><div class="stat-label">{lbl}</div>'
        f'<div class="stat-value">{val}</div></div>',
        unsafe_allow_html=True,
    )

notes_str = " · ".join(info["two_week_notes"])
st.markdown(
    f"""
    <div class="info-strip">
      <div class="info-strip-icon">i</div>
      <div class="info-strip-text">{html_escape(notes_str)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("</div>", unsafe_allow_html=True)


# ── SECTION: EINGABE ───────────────────────────────────────────────────

st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown(
    '<h2 class="card-title">Abwesenheit oder Sonderwunsch hinzufügen</h2>'
    '<p style="color:#94a3b8; margin: -8px 0 14px; font-size:12.5px;">'
    'Krankmeldung · Ferien · Termin · HO · eigene Regel</p>',
    unsafe_allow_html=True,
)

with st.form("entry_form", clear_on_submit=True):
    fc1, fc2, fc3, fc4 = st.columns([1, 1, 1, 2])
    with fc1:
        f_employee = st.selectbox("Mitarbeiterin", EMPLOYEE_LIST, key="f_employee")
    with fc2:
        f_type = st.selectbox("Grund", list(ABSENCE_TYPES.keys()), key="f_type")
    with fc3:
        f_slot_label = st.selectbox(
            "Halbtag",
            ["Ganzer Tag", "Vormittag", "Nachmittag"],
            key="f_slot",
        )
    with fc4:
        f_note = st.text_input(
            "Notiz (optional)",
            placeholder="z.B. Arzttermin 14:30, oder Task-Code für Custom",
            key="f_note",
        )

    f_days = st.multiselect(
        "Tage",
        options=[0, 1, 2, 3, 4],
        format_func=lambda d: DAY_SHORTS[d],
        default=[],
        key="f_days",
    )

    bc1, bc2, bc3 = st.columns([1, 1, 5])
    with bc1:
        all_days = st.form_submit_button("Alle Tage")
    with bc2:
        no_days = st.form_submit_button("Keine")
    with bc3:
        submit = st.form_submit_button("＋ Hinzufügen", type="primary")

    if all_days:
        st.session_state.f_days = [0, 1, 2, 3, 4]
        st.rerun()
    if no_days:
        st.session_state.f_days = []
        st.rerun()

    if submit:
        if not f_days:
            st.warning("Bitte mindestens einen Tag auswählen.")
        else:
            slots = (
                [0, 1] if f_slot_label == "Ganzer Tag"
                else [0] if f_slot_label == "Vormittag"
                else [1]
            )
            note = (f_note or "").strip()
            task = note if (f_type == "Custom" and note) else (
                "HO" if f_type == "Custom" else ""
            )
            st.session_state.entries.append({
                "name": f_employee,
                "type": f_type,
                "days": sorted(f_days),
                "slots": slots,
                "note": note,
                "task": task,
            })
            st.rerun()

# ── Eingetragene Einträge ──────────────────────────────────────────────

st.markdown(
    f'<div style="display:flex; justify-content:space-between; align-items:center; margin-top:14px;">'
    f'<h3 style="font-size:13px; margin:0; color:#e6ebf5;">Eingetragene Einträge</h3>'
    f'<span style="font-size:12px; color:#64748b; background:#1a2338; padding:3px 10px; border-radius:999px;">'
    f'{len(st.session_state.entries)} '
    f'{"Eintrag" if len(st.session_state.entries) == 1 else "Einträge"}</span></div>',
    unsafe_allow_html=True,
)

if not st.session_state.entries:
    st.markdown(
        '<div class="entries-empty">Keine Abwesenheiten — alles normal diese Woche.</div>',
        unsafe_allow_html=True,
    )
else:
    for idx, e in enumerate(st.session_state.entries):
        col_e, col_x = st.columns([20, 1])
        with col_e:
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
            st.markdown(
                f"""
                <div class="entry-item">
                  <span class="entry-dot {type_key}"></span>
                  <span class="entry-name">{html_escape(e['name'])}</span>
                  <span class="entry-meta">· {html_escape(e['type'])} · {days_str} ({slot_str})</span>
                  {note_html}
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_x:
            if st.button("✕", key=f"rm_{idx}", help="Entfernen"):
                st.session_state.entries.pop(idx)
                st.rerun()

st.markdown("</div>", unsafe_allow_html=True)


# ── ACTION BAR ─────────────────────────────────────────────────────────

ac1, ac2 = st.columns([1, 3])
with ac1:
    generate = st.button(
        "▶  Arbeitsplan generieren",
        type="primary",
        use_container_width=True,
    )
status_slot = ac2.empty()

if generate:
    with st.spinner("Plan wird generiert · Regelprüfung · Zuteilung · Excel-Export …"):
        try:
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

            random.seed(kw + st.session_state.week_offset)
            sched = build_schedule(kw, monday, overrides, state_file=str(STATE_FILE))
            for key, n in extra_notes.items():
                sched.notes[key] = n

            issues = validate_schedule(sched)

            filename = f"Arbeitsplan_KW{kw}_{monday.strftime('%Y%m%d')}.xlsx"
            output_path = OUTPUT_DIR / filename
            write_excel(sched, str(output_path))

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
                {"day": DAYS[d], "person": sched.tagesverantwortung.get(d, "?")}
                for d in range(5)
            ]
            phc_liste = [
                {"day": DAYS[d], "person": sched.phc_liste.get(d, "?")}
                for d in range(5)
            ]

            rows = []
            for name in EMPLOYEE_LIST:
                if name not in sched.employees:
                    continue
                emp = sched.employees[name]
                cells = []
                for day in range(5):
                    for slot in range(2):
                        task = sched.get_task(name, day, slot)
                        is_phc = slot == 1 and sched.phc_liste.get(day) == name
                        cells.append({"task": task, "is_phc": is_phc})
                rows.append({
                    "name": emp.name,
                    "pct": emp.pct,
                    "cells": cells,
                })

            with open(output_path, "rb") as fh:
                excel_bytes = fh.read()

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
            status_slot.success(f"✓ {filename}")
        except Exception as exc:  # pragma: no cover
            status_slot.error(f"Fehler beim Generieren: {exc}")
            st.exception(exc)


# ── SECTION: ERGEBNIS ──────────────────────────────────────────────────

res = st.session_state.result
if res:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    head_l, head_r = st.columns([3, 1])
    head_l.markdown('<h2 class="card-title">Ergebnis</h2>', unsafe_allow_html=True)
    with head_r:
        st.download_button(
            "⬇  Excel herunterladen",
            data=res["excel_bytes"],
            file_name=res["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if res["warnings"]:
        items = "".join(f"<li>{html_escape(w)}</li>" for w in res["warnings"])
        st.markdown(
            f'<div class="banner banner-warn"><h4>Warnungen vor dem Generieren '
            f'(Regel-Impact)</h4><ul>{items}</ul></div>',
            unsafe_allow_html=True,
        )

    if res["issues"]:
        items = "".join(f"<li>{html_escape(i)}</li>" for i in res["issues"])
        st.markdown(
            f'<div class="banner banner-error"><h4>Konflikte / Regeln nicht '
            f'vollständig einhaltbar</h4><ul>{items}</ul></div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="banner banner-ok">Alle harten Regeln eingehalten.</div>',
            unsafe_allow_html=True,
        )

    # ── Coverage ─────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Besetzung</div>', unsafe_allow_html=True)
    cov_cols_per_row = 5
    cov = res["coverage"]
    for i in range(0, len(cov), cov_cols_per_row):
        row = cov[i:i + cov_cols_per_row]
        cols = st.columns(len(row))
        for col, c in zip(cols, row):
            tel_cls = "cov-ok" if c["tel_ok"] else "cov-bad"
            abkl_cls = "cov-ok" if c["abkl_ok"] else "cov-bad"
            col.markdown(
                f"""
                <div class="cov-card">
                  <div class="cov-head">
                    <span class="cov-day">{c['day']}</span>
                    <span>{c['slot']}</span>
                  </div>
                  <div class="cov-row"><span>TEL</span>
                    <span class="cov-badge {tel_cls}">{c['tel']}/{c['tel_target']}</span>
                  </div>
                  <div class="cov-row"><span>ABKL</span>
                    <span class="cov-badge {abkl_cls}">{c['abkl']}/{c['abkl_target']}</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ── Tagesverantwortung & PHC ─────────────────────────────────────
    tv_col, phc_col = st.columns(2)
    with tv_col:
        st.markdown('<div class="section-title">Tagesverantwortung</div>', unsafe_allow_html=True)
        for item in res["tagesverantwortung"]:
            is_open = "OFFEN" in item["person"]
            cls = "pill-val open" if is_open else "pill-val"
            st.markdown(
                f'<div class="pill"><span class="pill-day">{item["day"]}</span>'
                f'<span class="{cls}">{html_escape(item["person"])}</span></div>',
                unsafe_allow_html=True,
            )
    with phc_col:
        st.markdown('<div class="section-title">PHC-Liste (18 Uhr)</div>', unsafe_allow_html=True)
        for item in res["phc_liste"]:
            is_open = "OFFEN" in item["person"]
            cls = "pill-val open" if is_open else "pill-val"
            st.markdown(
                f'<div class="pill"><span class="pill-day">{item["day"]}</span>'
                f'<span class="{cls}">{html_escape(item["person"])}</span></div>',
                unsafe_allow_html=True,
            )

    # ── Wochenaufgaben ───────────────────────────────────────────────
    st.markdown('<div class="section-title">Wochenaufgaben</div>', unsafe_allow_html=True)
    wa = res["wochenaufgaben"]
    wa_entries = [
        ("Direktbestellung", wa["direkt"]),
        ("ONB", wa["onb"]),
        ("BTM", wa["btm"]),
    ]
    cols = st.columns(len(wa_entries))
    for col, (lbl, val) in zip(cols, wa_entries):
        col.markdown(
            f'<div class="wa-card"><div class="wa-label">{lbl}</div>'
            f'<div class="wa-val">{html_escape(val or "–")}</div></div>',
            unsafe_allow_html=True,
        )

    # ── Schedule-Tabelle ─────────────────────────────────────────────
    st.markdown('<div class="section-title">Vorschau</div>', unsafe_allow_html=True)
    parts = ['<div class="schedule-wrap"><table class="schedule">']
    parts.append("<thead><tr>")
    parts.append('<th rowspan="2" style="text-align:left; padding-left:14px;">Name</th>')
    parts.append('<th rowspan="2">%</th>')
    for d in DAYS:
        parts.append(f'<th class="day-col" colspan="2">{d}</th>')
    parts.append("</tr><tr>")
    for _ in DAYS:
        parts.append("<th>VM</th><th>NM</th>")
    parts.append("</tr></thead><tbody>")

    for r in res["rows"]:
        parts.append("<tr>")
        parts.append(f'<td class="name-col">{html_escape(r["name"])}</td>')
        parts.append(f'<td class="pct-col">{r["pct"]}%</td>')
        for c in r["cells"]:
            task = c["task"] or ""
            if c["is_phc"]:
                parts.append(f'<td class="cell-phc">{html_escape(task)}</td>')
            else:
                bg, fg = style_for_task(task)
                parts.append(
                    f'<td style="background:{bg}; color:{fg};">{html_escape(task)}</td>'
                )
        parts.append("</tr>")
    parts.append("</tbody></table></div>")
    parts.append(
        '<p class="tiny-hint">Farbige Markierungen entsprechen den Task-Typen '
        'im Excel. Blaue Zellen = PHC/18-Uhr-Liste.</p>'
    )
    st.markdown("".join(parts), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ── FOOTER ─────────────────────────────────────────────────────────────

st.markdown(
    '<div style="display:flex; justify-content:space-between; padding:16px 6px 0; '
    'font-size:12px; color:#64748b; border-top:1px solid #273149; margin-top:18px;">'
    '<span>KSC Arbeitsplan v6.0 · Streamlit-Variante</span>'
    '<span style="color:#22c55e;">● online</span></div>',
    unsafe_allow_html=True,
)
