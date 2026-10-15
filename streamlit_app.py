#!/usr/bin/env python3
"""
KSC Arbeitsplan – Streamlit-App (Version 7.0, Oktober 2026)
============================================================
Oberfläche für https://arbeitsplan-ksc.streamlit.app

- Schritt 1 «Planung»: Abwesenheiten/Sonderwünsche erfassen (auch über
  mehrere Wochen) und das Team verwalten – Personen, Pensen, Anwesenheit,
  Rollen, Fixeinsätze, Rotationen, Besetzungsregeln, BackUp-Team.
  Alles ohne Code-Änderung, gespeichert in team.json.
- Schritt 2 «Ergebnis»: Wochenplan mit Tagesrollen, Konflikten,
  Excel-Download und Plan-Editor pro Zelle.
- Speichern: team.json / abwesenheiten.json lokal; optional direkt ins
  GitHub-Repo (Secrets [github] token/repo/branch) – damit überleben
  Änderungen auch Neustarts von Streamlit Cloud.

Start lokal:
    pip install -r requirements.txt
    streamlit run streamlit_app.py
"""

from __future__ import annotations

import base64
import copy
import json
import os
import sys
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import arbeitskalender as ak  # noqa: E402
from arbeitskalender import (  # noqa: E402
    DAYS, DAY_KEYS, HALBTAG_WERTE, NUR_MANUELL, PARITAET, ROLLEN, SCHALTER_CODES,
    SCHALTER_RESERVE_MODI, SLOTS, SLOT_KEYS, TASK_COLORS, TPA_TEL_MODI,
)

CONFIG_PATH = Path(os.environ.get("KSC_TEAM_CONFIG", BASE_DIR / "team.json"))
ABSENCE_PATH = Path(os.environ.get("KSC_ABSENCES", BASE_DIR / "abwesenheiten.json"))

ABSENCE_TYPES = {
    "Krank": "KRANK",
    "Ferien": "FERIEN",
    "Termin": "TERMIN",
    "Home Office": "HO",
    "Custom": "CUSTOM",
}

# Auswahl im Plan-Editor (leer = Zelle frei)
TASK_OPTIONS = [
    "", "TEL", "ABKL",
    "ERF7", "ERF7/Q", "ERF7/HUB", "ERF7/SCH",
    "ERF8", "ERF8/Q", "ERF8/HUB",
    "ERF9", "ERF9/Q", "ERF9/TEL",
    "ERF4/SCH", "ERF5",
    "PO", "PO/SCAN", "PO/ABKL", "PO/TEL", "scanning", "SCAN",
    "HO", "HO/Q", "TL/Q",
    "KGS", "KTG", "KSC Spez.", "TAGES PA", "TAGES PA/TEL",
    "RX Abo", "VBZ/Q", "ERP", "KSV",
    "KRANK", "FERIEN",
]


# ══════════════════════════════════════════════════════════════════════
#  Streamlit-Setup & CSS
# ══════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="KSC Arbeitsplan", page_icon="📅", layout="wide",
                   initial_sidebar_state="collapsed")

CUSTOM_CSS = """
<style>
:root {
  --bg: #f4f6fb; --surface: #ffffff; --surface-2: #f8fafc;
  --border: #e3e8f0; --border-2: #cbd5e1;
  --text: #0f172a; --text-dim: #475569; --text-mute: #94a3b8; --heading: #0b1020;
  --primary: #0f172a; --primary-2: #1e293b;
  --accent: #2563eb; --accent-soft: #dbeafe;
  --ok: #16a34a; --ok-soft: #dcfce7; --warn: #d97706; --warn-soft: #fef3c7;
  --err: #dc2626; --err-soft: #fee2e2; --info-soft: #eff6ff; --phc: #00008B;
}
html, body, [data-testid="stAppViewContainer"] { background: var(--bg); color: var(--text); }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display: none; }
.block-container { padding-top: 1.2rem !important; padding-bottom: 2rem !important; max-width: 1500px; }

.topbar { background: var(--surface); border: 1px solid var(--border); border-radius: 12px;
  padding: 12px 16px; margin-bottom: 14px; box-shadow: 0 1px 2px rgba(15,23,42,.04); }
.brand { display: flex; align-items: center; gap: 12px; }
.brand-icon { width: 38px; height: 38px; border-radius: 9px; background: var(--primary); color: white;
  display: grid; place-items: center; font-size: 18px; }
.brand-title { font-size: 15px; font-weight: 700; color: var(--heading); line-height: 1.1; }
.brand-sub { font-size: 10.5px; color: var(--text-mute); letter-spacing: .8px; text-transform: uppercase; }

.kw-pill { display: flex; align-items: center; gap: 4px; background: var(--surface-2);
  border: 1px solid var(--border); border-radius: 10px; padding: 4px; }
.kw-center { padding: 4px 14px; min-width: 170px; text-align: center; }
.kw-big { display:block; font-size: 14px; font-weight: 700; color: var(--heading); }
.kw-range { display:block; font-size: 11px; color: var(--text-mute); margin-top: 2px; }
.badge-kw { background: #f1f5f9; color: var(--text-dim); border: 1px solid var(--border);
  padding: 6px 12px; border-radius: 999px; font-size: 11px; font-weight: 700; letter-spacing: .5px; text-transform: uppercase; }

div.stButton > button[kind="primary"] { background: var(--primary) !important; border: 1px solid var(--primary) !important;
  color: white !important; font-weight: 600 !important; border-radius: 9px !important; }
div.stButton > button[kind="primary"]:hover { background: var(--primary-2) !important; border-color: var(--primary-2) !important; }
div.stButton > button[kind="secondary"] { background: var(--surface) !important; border: 1px solid var(--border) !important;
  color: var(--text) !important; border-radius: 8px !important; font-weight: 500 !important; }
div.stButton > button[kind="secondary"]:hover { background: var(--surface-2) !important; border-color: var(--border-2) !important; }
div.stButton > button, div.stButton > button > div, div.stButton > button p { white-space: nowrap !important; }
div.stButton > button { padding-left: 6px !important; padding-right: 6px !important; }
div.stDownloadButton > button { background: var(--ok) !important; border: 1px solid var(--ok) !important;
  color: white !important; border-radius: 9px !important; font-weight: 600 !important; }
div.stDownloadButton > button:hover { filter: brightness(1.08); }

.stats-strip { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 14px 18px;
  margin-bottom: 16px; display: flex; gap: 36px; align-items: center; flex-wrap: wrap; box-shadow: 0 1px 2px rgba(15,23,42,.04); }
.stat-item .lbl { font-size: 10.5px; color: var(--text-mute); text-transform: uppercase; letter-spacing: .8px; font-weight: 600; margin-bottom: 4px; }
.stat-item .val { font-size: 18px; font-weight: 700; color: var(--heading); font-feature-settings: "tnum"; }
.stat-item .val small { font-size: 11.5px; font-weight: 500; color: var(--text-dim); margin-left: 4px; }
.stat-divider { height: 32px; width: 1px; background: var(--border); }
.info-pill { margin-left: auto; display: flex; align-items: center; gap: 8px; background: var(--info-soft);
  border: 1px solid #bfdbfe; color: #1e3a8a; padding: 7px 14px; border-radius: 999px; font-size: 12.5px; }
.info-pill .ico { width: 16px; height: 16px; border-radius: 50%; background: #2563eb; color: white;
  display: grid; place-items: center; font-size: 10px; font-weight: 700; }

.card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 18px 22px;
  margin-bottom: 16px; box-shadow: 0 1px 2px rgba(15,23,42,.04); }
.card h2 { font-size: 15.5px; font-weight: 700; color: var(--heading); margin: 0 0 4px; }
.card .sub { font-size: 12.5px; color: var(--text-mute); margin-bottom: 14px; }
.card-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
.field-label { font-size: 10.5px; color: var(--text-mute); text-transform: uppercase; letter-spacing: .8px; font-weight: 700; margin: 0 0 6px 2px; }

.empty-state { padding: 40px 18px; text-align: center; color: var(--text-mute); font-size: 13px; }
.empty-state .icon { width: 44px; height: 44px; margin: 0 auto 10px; border-radius: 12px; background: var(--surface-2);
  border: 1px solid var(--border); display: grid; place-items: center; color: var(--text-mute); font-size: 20px; }
.empty-state .ttl { color: var(--text); font-weight: 600; margin-bottom: 4px; font-size: 13px; }

.entry-row { background: var(--surface); border: 1px solid var(--border); border-radius: 9px; padding: 9px 12px;
  display: flex; align-items: center; gap: 10px; font-size: 13px; flex-wrap: wrap; }
.entry-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.entry-dot.KRANK { background: var(--err); } .entry-dot.FERIEN { background: var(--warn); }
.entry-dot.TERMIN { background: #06b6d4; } .entry-dot.HO { background: var(--accent); } .entry-dot.CUSTOM { background: var(--text-mute); }
.entry-name { font-weight: 600; color: var(--text); }
.entry-meta { color: var(--text-dim); font-size: 12.5px; margin-left: 4px; }
.entry-note { color: #b45309; font-style: italic; margin-left: 6px; font-size: 12px; }

.banner { display: flex; align-items: flex-start; gap: 10px; padding: 12px 16px; border-radius: 10px; margin-bottom: 14px; border: 1px solid; font-size: 13.5px; }
.banner .check { width: 22px; height: 22px; border-radius: 50%; display: grid; place-items: center; flex-shrink: 0; font-weight: 700; font-size: 13px; }
.banner-ok { background: var(--ok-soft); border-color: rgba(22,163,74,.3); color: #14532d; } .banner-ok .check { background: var(--ok); color: white; }
.banner-warn { background: var(--warn-soft); border-color: rgba(217,119,6,.3); color: #78350f; } .banner-warn .check { background: var(--warn); color: white; }
.banner-error { background: var(--err-soft); border-color: rgba(220,38,38,.3); color: #7f1d1d; } .banner-error .check { background: var(--err); color: white; }
.banner-info { background: var(--info-soft); border-color: #bfdbfe; color: #1e3a8a; } .banner-info .check { background: #2563eb; color: white; }
.banner h4 { margin: 0 0 4px; font-size: 13px; font-weight: 700; }
.banner ul { margin: 4px 0 0; padding-left: 18px; font-size: 12.5px; }

.plan-wrap { overflow-x: auto; border-radius: 10px; border: 1px solid var(--border); background: var(--surface); }
table.plan { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 11.5px;
  font-family: -apple-system, "Segoe UI", "Inter", sans-serif; min-width: 1100px; background: var(--surface); }
table.plan th, table.plan td { padding: 6px 6px; border-right: 1px solid var(--border); border-bottom: 1px solid var(--border); vertical-align: middle; }
table.plan th { background: var(--surface-2); color: var(--text-dim); font-weight: 700; font-size: 10.5px; letter-spacing: .6px; text-transform: uppercase; text-align: center; }
table.plan th.team-head { text-align: left; padding: 10px 14px; vertical-align: top; background: var(--surface); }
table.plan th.team-head .legend { margin-top: 4px; color: var(--text-mute); font-size: 10.5px; font-weight: 500; text-transform: none; letter-spacing: 0; }
table.plan th.day-head { background: var(--surface); padding: 10px 8px; }
table.plan th.day-head .day-name { display: block; color: var(--text-dim); font-weight: 700; font-size: 11px; letter-spacing: .8px; text-transform: uppercase; }
.tv-pill { display: inline-flex; align-items: center; gap: 6px; margin-top: 6px; padding: 3px 9px; background: var(--surface-2);
  border: 1px solid var(--border); border-radius: 999px; font-size: 11px; font-weight: 600; color: var(--text); text-transform: none; letter-spacing: 0; }
.tv-tag { font-size: 9.5px; font-weight: 800; padding: 1px 5px; border-radius: 4px; letter-spacing: .4px; }
.tv-pill.open .tv-name { color: var(--warn); }
table.plan th.slot-head { background: var(--surface-2); font-size: 9.5px; padding: 5px 4px; color: var(--text-mute); font-weight: 600; }
table.plan td.name-cell { background: var(--surface); padding: 6px 12px; text-align: left; font-size: 12px; color: var(--text); font-weight: 500; white-space: nowrap; position: sticky; left: 0; }
table.plan td.pct-cell { background: var(--surface); text-align: right; padding-right: 14px; color: var(--text-mute); font-size: 11px; font-weight: 600; }
table.plan tbody tr:nth-child(even) td.name-cell, table.plan tbody tr:nth-child(even) td.pct-cell, table.plan tbody tr:nth-child(even) td.task-cell { background: var(--surface-2); }
td.task-cell { text-align: center; padding: 5px 4px; background: var(--surface); }
table.plan td.sep-cell { background: #e2e8f0 !important; color: var(--text-dim); font-weight: 700; font-size: 10.5px; letter-spacing: .6px; text-transform: uppercase; padding: 6px 12px; text-align: left; }
.pill { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; letter-spacing: .2px; min-width: 64px; text-align: center; }
.pill-phc { background: var(--phc) !important; color: white !important; font-weight: 700; }
.pill-termin { background: #fef3c7; color: #78350f; }
.pill-note { display:block; font-size: 9.5px; color: var(--text-mute); margin-top: 2px; }

table.roles { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 14px; }
table.roles th, table.roles td { border: 1px solid var(--border); padding: 6px 10px; text-align: center; }
table.roles th { background: var(--surface-2); color: var(--text-dim); font-size: 10.5px; text-transform: uppercase; letter-spacing: .6px; }
table.roles td.lbl { text-align: left; font-weight: 600; color: var(--text); background: var(--surface-2); white-space: nowrap; }
table.roles td.open { color: var(--warn); font-weight: 700; }

.wa-strip { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 14px; }
.wa-card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 12px 16px; display: flex; align-items: center; gap: 12px; }
.wa-icon { width: 32px; height: 32px; border-radius: 8px; background: var(--surface-2); border: 1px solid var(--border); display: grid; place-items: center; font-size: 14px; }
.wa-meta .lbl { font-size: 10px; color: var(--text-mute); text-transform: uppercase; letter-spacing: .8px; font-weight: 700; }
.wa-meta .val { font-size: 14px; font-weight: 700; color: var(--heading); margin-top: 2px; }
.legend-wrap { display: flex; flex-wrap: wrap; gap: 8px; }
.legend-item { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--text-dim); }
.element-container { margin-bottom: 0.4rem; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
#  Hilfsfunktionen
# ══════════════════════════════════════════════════════════════════════

def html_escape(s) -> str:
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _fg_for(bg_hex: str) -> str:
    h = bg_hex.lstrip("#")
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return "#1F1F1F"
    lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#1F1F1F" if lum > 0.6 else "#FFFFFF"


def pill_for_task(task: str) -> tuple[str, str]:
    if not task:
        return "transparent", "transparent"
    if task.startswith("*"):
        return "#fef3c7", "#78350f"
    color = ak.get_task_color(task)
    if color:
        return f"#{color}", _fg_for(color)
    return "#f1f5f9", "#334155"


def person_label(short: str) -> str:
    cfg = st.session_state.cfg
    for m in cfg["mitarbeiter"]:
        if m["short"] == short:
            return short if m.get("name") in (short, "", None) else f"{short} · {m['name']}"
    for b in cfg["backup_team"]:
        if b["short"] == short:
            return f"{short} (BackUp)"
    return short


def active_members() -> list[dict]:
    return [m for m in st.session_state.cfg["mitarbeiter"] if m.get("aktiv", True)]


def active_shorts() -> list[str]:
    return [m["short"] for m in active_members()]


def backup_shorts() -> list[str]:
    return [b["short"] for b in st.session_state.cfg["backup_team"]]


def employee_options() -> dict[str, str]:
    """Anzeige-Label → Kurzname (aktives Team + BackUp-Team)."""
    return {person_label(s): s for s in active_shorts() + backup_shorts()}


def role_options() -> dict[str, str]:
    return {f"{r} – {ROLLEN[r]}": r for r in ROLLEN}


def get_week_info(offset: int = 0) -> dict:
    monday = ak.get_next_monday() + timedelta(weeks=offset)
    iso_year, kw, _ = monday.isocalendar()
    friday = monday + timedelta(days=4)
    return {
        "monday": monday, "friday": friday, "kw": kw, "iso_year": iso_year,
        "widx": ak.week_index(monday), "is_even": kw % 2 == 0,
        "monday_str": monday.strftime("%d.%m.%Y"), "friday_str": friday.strftime("%d.%m.%Y"),
        "two_week_notes": ak.two_week_notes(kw, st.session_state.cfg),
    }


def entry_widx_start(entry: dict) -> int:
    jahr = int(entry.get("jahr") or date.today().isocalendar()[0])
    kw = int(entry.get("kw_start") or 1)
    try:
        return ak.week_index(date.fromisocalendar(jahr, kw, 1))
    except ValueError:
        return ak.week_index(date.fromisocalendar(jahr, 1, 1))


def entry_active(entry: dict, widx: int) -> bool:
    start = entry_widx_start(entry)
    weeks = max(1, int(entry.get("weeks", 1)))
    return start <= widx < start + weeks


def entry_kw_label(entry: dict) -> str:
    kw = int(entry.get("kw_start") or 0)
    weeks = max(1, int(entry.get("weeks", 1)))
    if weeks == 1:
        return f"KW {kw}"
    end = date.fromisocalendar(int(entry.get("jahr") or date.today().year), kw, 1) + timedelta(weeks=weeks - 1)
    return f"KW {kw}–{end.isocalendar()[1]}"


# ══════════════════════════════════════════════════════════════════════
#  Persistenz (lokal + optional GitHub)
# ══════════════════════════════════════════════════════════════════════

def github_conf() -> dict | None:
    try:
        g = st.secrets["github"]
        token, repo = g["token"], g["repo"]
    except Exception:
        return None
    if not token or not repo:
        return None
    try:
        branch = g.get("branch", "main") or "main"
    except Exception:
        branch = "main"
    return {"token": token, "repo": repo, "branch": branch}


def github_put(rel_path: str, content: str, message: str) -> tuple[bool, str]:
    conf = github_conf()
    if not conf:
        return False, "GitHub nicht konfiguriert (Secrets [github] token/repo fehlen)"
    try:
        import requests
    except ImportError:
        return False, "Paket 'requests' fehlt"
    url = f"https://api.github.com/repos/{conf['repo']}/contents/{rel_path}"
    headers = {"Authorization": f"Bearer {conf['token']}", "Accept": "application/vnd.github+json"}
    try:
        r = requests.get(url, headers=headers, params={"ref": conf["branch"]}, timeout=20)
        sha = r.json().get("sha") if r.status_code == 200 else None
        body = {"message": message, "branch": conf["branch"],
                "content": base64.b64encode(content.encode("utf-8")).decode("ascii")}
        if sha:
            body["sha"] = sha
        r = requests.put(url, headers=headers, json=body, timeout=20)
        if r.status_code in (200, 201):
            return True, f"{rel_path} nach GitHub ({conf['repo']}@{conf['branch']}) gespeichert"
        return False, f"GitHub-Fehler {r.status_code}: {r.text[:200]}"
    except Exception as exc:  # Netzwerk etc.
        return False, f"GitHub nicht erreichbar: {exc}"


def write_local(path: Path, content: str) -> tuple[bool, str]:
    try:
        path.write_text(content, encoding="utf-8")
        return True, f"{path.name} lokal gespeichert"
    except Exception as exc:
        return False, f"{path.name} konnte nicht geschrieben werden: {exc}"


def entries_json() -> str:
    return json.dumps({"_hinweis": "Abwesenheiten/Sonderwünsche der App (automatisch gepflegt)",
                       "entries": st.session_state.entries}, ensure_ascii=False, indent=2) + "\n"


def save_everything(push_github: bool) -> list[tuple[bool, str]]:
    msgs = []
    cfg_json = ak.config_to_json(st.session_state.cfg)
    msgs.append(write_local(CONFIG_PATH, cfg_json))
    msgs.append(write_local(ABSENCE_PATH, entries_json()))
    if push_github:
        msgs.append(github_put(CONFIG_PATH.name, cfg_json, "team.json über die App aktualisiert"))
        msgs.append(github_put(ABSENCE_PATH.name, entries_json(), "Abwesenheiten über die App aktualisiert"))
    st.session_state.cfg_saved_json = cfg_json
    st.session_state.entries_saved_json = entries_json()
    return msgs


def load_entries_from_disk() -> list[dict]:
    if ABSENCE_PATH.exists():
        try:
            data = json.loads(ABSENCE_PATH.read_text(encoding="utf-8"))
            return list(data.get("entries", []))
        except Exception:
            return []
    return []


def replace_cfg(new_cfg: dict) -> None:
    st.session_state.cfg = ak.normalize_config(new_cfg)
    st.session_state.cfg_version = st.session_state.get("cfg_version", 0) + 1
    st.session_state.result = None


# ══════════════════════════════════════════════════════════════════════
#  Session-State
# ══════════════════════════════════════════════════════════════════════

if "cfg" not in st.session_state:
    try:
        replace_cfg(ak.load_config(CONFIG_PATH))
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()
    st.session_state.cfg_saved_json = ak.config_to_json(st.session_state.cfg)
if "entries" not in st.session_state:
    st.session_state.entries = load_entries_from_disk()
    st.session_state.entries_saved_json = entries_json()

st.session_state.setdefault("week_offset", 0)
st.session_state.setdefault("result", None)
st.session_state.setdefault("step", 1)
st.session_state.setdefault("sel_days", set())
st.session_state.setdefault("sel_slot", "Ganzer Tag")
st.session_state.setdefault("f_type", "Krank")
st.session_state.setdefault("f_note", "")
st.session_state.setdefault("f_weeks", 1)
st.session_state.setdefault("save_msgs", [])

# Pending-Clears (Widget-Keys dürfen nur VOR dem Rendern geändert werden)
if st.session_state.pop("_pending_clear_note", False):
    st.session_state["f_note"] = ""

_options = employee_options()
if st.session_state.get("f_employee") not in _options:
    st.session_state["f_employee"] = next(iter(_options), "")

cfg_dirty = ak.config_to_json(st.session_state.cfg) != st.session_state.cfg_saved_json
entries_dirty = entries_json() != st.session_state.entries_saved_json


# ══════════════════════════════════════════════════════════════════════
#  Plan-Generierung
# ══════════════════════════════════════════════════════════════════════

def build_rows(sched: "ak.Schedule") -> tuple[list[dict], list[dict]]:
    cfg = st.session_state.cfg
    rows, backup_rows = [], []
    for m in cfg["mitarbeiter"]:
        short = m["short"]
        if short not in sched.employees:
            continue
        emp = sched.employees[short]
        cells = [{"task": sched.get_task(short, d, s),
                  "is_phc": s == 1 and sched.phc_liste.get(d) == short,
                  "note": sched.notes.get((short, d, s), "")}
                 for d in range(5) for s in range(2)]
        rows.append({"full_name": emp.name, "short": short, "pct": emp.pct, "cells": cells})
    for b in cfg["backup_team"]:
        cells = [{"task": sched.get_backup_task(b["short"], d, s), "is_phc": False, "note": ""}
                 for d in range(5) for s in range(2)]
        backup_rows.append({"full_name": b.get("name") or b["short"], "short": b["short"], "pct": None, "cells": cells})
    return rows, backup_rows


def excel_bytes(sched: "ak.Schedule") -> bytes:
    buf = BytesIO()
    ak.write_excel(sched, buf)
    return buf.getvalue()


def refresh_result_from_sched(sched: "ak.Schedule") -> None:
    res = st.session_state.result
    res["issues"] = ak.validate_schedule(sched)
    res["coverage"] = ak.coverage_table(sched)
    res["rows"], res["backup_rows"] = build_rows(sched)
    res["excel_bytes"] = excel_bytes(sched)


def run_generation() -> None:
    cfg = st.session_state.cfg
    info = get_week_info(st.session_state.week_offset)
    kw, monday, widx = info["kw"], info["monday"], info["widx"]

    employees = ak.create_employees(kw, cfg)
    backups = set(backup_shorts())
    overrides, extra_notes, absences = {}, {}, []

    for entry in st.session_state.entries:
        if not entry_active(entry, widx):
            continue
        name = entry["name"]
        is_backup = name in backups
        if name not in employees and not is_backup:
            continue
        code = ABSENCE_TYPES.get(entry["type"], "CUSTOM")
        note = (entry.get("note") or "").strip()
        if code == "TERMIN":
            label = f"*{note}" if note else "*Termin"
        elif code == "CUSTOM":
            label = (entry.get("task") or "").strip() or "HO"
        else:
            label = code
        for day in entry["days"]:
            for slot in entry["slots"]:
                if is_backup:
                    overrides[(name, day, slot)] = label
                    continue
                if employees[name].is_available(day, slot):
                    overrides[(name, day, slot)] = label
                    if note and code in ("KRANK", "FERIEN", "HO"):
                        extra_notes[(name, day, slot)] = note
                    if code in ("KRANK", "FERIEN") or label.startswith("*"):
                        absences.append((name, day, slot))

    warnings = ak.check_rule_impact(absences, employees, cfg) if absences else []
    sched = ak.build_schedule(kw, monday, overrides, cfg=cfg)
    for key, n in extra_notes.items():
        sched.notes[key] = n

    rows, backup_rows = build_rows(sched)
    st.session_state.result = {
        "filename": f"Arbeitsplan_KW{kw}_{monday.strftime('%Y%m%d')}.xlsx",
        "excel_bytes": excel_bytes(sched),
        "kw": kw, "iso_year": info["iso_year"],
        "monday_str": info["monday_str"], "friday_str": info["friday_str"],
        "warnings": warnings,
        "issues": ak.validate_schedule(sched),
        "coverage": ak.coverage_table(sched),
        "tagesverantwortung": [sched.tagesverantwortung.get(d, "?") for d in range(5)],
        "phc_liste": [sched.phc_liste.get(d, "?") for d in range(5)],
        "tagespa": [sched.tagespa.get(d, "?") for d in range(5)],
        "po_liste": [sched.po_liste.get(d, "?") for d in range(5)],
        "po_zusatz": [", ".join(sched.po_zusatz.get(d, [])) or "–" for d in range(5)],
        "schalter": [sched.schalter.get(d, "–") for d in range(5)],
        "erf9": [sched.erf9.get(d, "–") for d in range(5)],
        "wochenaufgaben": {"direkt": sched.wochenaufgabe_direkt, "onb": sched.wochenaufgabe_onb,
                           "btm": sched.wochenaufgabe_btm},
        "rows": rows, "backup_rows": backup_rows,
        "config_problems": ak.validate_config(cfg),
    }
    st.session_state.sched = sched
    st.session_state.step = 2


# ══════════════════════════════════════════════════════════════════════
#  Kopfzeile
# ══════════════════════════════════════════════════════════════════════

info = get_week_info(st.session_state.week_offset)

st.markdown('<div class="topbar">', unsafe_allow_html=True)
hcol = st.columns([2.6, 2.4, 2.6, 1.0, 1.6])

with hcol[0]:
    st.markdown(
        '<div class="brand"><div class="brand-icon">📅</div><div>'
        '<div class="brand-title">Arbeitsplan</div>'
        '<div class="brand-sub">KSC · Team-Scheduling</div></div></div>',
        unsafe_allow_html=True,
    )

with hcol[1]:
    s1, s2 = st.columns(2)
    with s1:
        if st.button("● 1   Planung" if st.session_state.step == 1 else "1   Planung", key="step1",
                     width="stretch", type="primary" if st.session_state.step == 1 else "secondary"):
            st.session_state.step = 1
            st.rerun()
    with s2:
        if st.button("● 2   Ergebnis" if st.session_state.step == 2 else "2   Ergebnis", key="step2",
                     width="stretch", type="primary" if st.session_state.step == 2 else "secondary",
                     disabled=st.session_state.result is None):
            st.session_state.step = 2
            st.rerun()

with hcol[2]:
    w1, w2, w3 = st.columns([1, 4, 1])
    with w1:
        if st.button("‹", key="wk_prev", help="Woche zurück", width="stretch"):
            if st.session_state.week_offset > -1:
                st.session_state.week_offset -= 1
                if st.session_state.step == 2 and st.session_state.result is not None:
                    run_generation()
                st.rerun()
    with w2:
        st.markdown(
            f'<div class="kw-pill" style="justify-content:center;"><div class="kw-center">'
            f'<span class="kw-big">KW {info["kw"]}</span>'
            f'<span class="kw-range">{info["monday_str"]} – {info["friday_str"]}</span></div></div>',
            unsafe_allow_html=True,
        )
    with w3:
        if st.button("›", key="wk_next", help="Woche vor", width="stretch"):
            st.session_state.week_offset += 1
            if st.session_state.step == 2 and st.session_state.result is not None:
                run_generation()
            st.rerun()

with hcol[3]:
    st.markdown(
        f'<div style="text-align:center; padding-top:10px;"><span class="badge-kw">'
        f'{"GERADE KW" if info["is_even"] else "UNGERADE KW"}</span></div>',
        unsafe_allow_html=True,
    )

with hcol[4]:
    if st.session_state.step == 2 and st.session_state.result:
        st.download_button("⬇  Excel herunterladen", data=st.session_state.result["excel_bytes"],
                           file_name=st.session_state.result["filename"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch", key="cta_download")
    else:
        if st.button("→  Plan generieren", key="cta_generate", type="primary", width="stretch"):
            with st.spinner("Plan wird generiert …"):
                try:
                    run_generation()
                except Exception as exc:
                    st.error(f"Fehler beim Generieren: {exc}")
                else:
                    st.rerun()

st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
#  Stats-Leiste
# ══════════════════════════════════════════════════════════════════════

n_active = len(active_shorts())
notes_str = " · ".join(info["two_week_notes"]) or "keine 2-Wochen-Besonderheiten"
storage = []
if cfg_dirty:
    storage.append("Team: ungespeicherte Änderungen")
if entries_dirty:
    storage.append("Abwesenheiten: ungespeichert")
storage_html = (f'<div class="info-pill" style="margin-left:12px; background:var(--warn-soft); '
                f'border-color:rgba(217,119,6,.3); color:#78350f;"><span class="ico" style="background:var(--warn);">!</span>'
                f'<span>{html_escape(" · ".join(storage))}</span></div>') if storage else ""
st.markdown(
    f"""
    <div class="stats-strip">
      <div class="stat-item"><div class="lbl">Kalenderwoche</div><div class="val">{info['kw']}</div></div>
      <div class="stat-divider"></div>
      <div class="stat-item"><div class="lbl">Montag</div><div class="val">{info['monday_str']}</div></div>
      <div class="stat-divider"></div>
      <div class="stat-item"><div class="lbl">Freitag</div><div class="val">{info['friday_str']}</div></div>
      <div class="stat-divider"></div>
      <div class="stat-item"><div class="lbl">Team</div><div class="val">{n_active} <small>aktiv</small></div></div>
      <div class="info-pill"><span class="ico">i</span><span>{html_escape(notes_str)}</span></div>
      {storage_html}
    </div>
    """,
    unsafe_allow_html=True,
)


# ══════════════════════════════════════════════════════════════════════
#  Team verwalten
# ══════════════════════════════════════════════════════════════════════

def _b(v) -> bool:
    return bool(v) and not (isinstance(v, float) and pd.isna(v))


def _s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _apply_people_table(edited: pd.DataFrame) -> list[str]:
    """Übernimmt die Personen-Tabelle in die Konfiguration. Liefert Probleme."""
    cfg = copy.deepcopy(st.session_state.cfg)
    by_short = {m["short"]: m for m in cfg["mitarbeiter"]}
    new_list, seen = [], set()
    for _, r in edited.iterrows():
        short = _s(r.get("Kurzname"))
        if not short:
            continue
        if short in seen:
            return [f"Kurzname doppelt: {short}"]
        seen.add(short)
        m = by_short.get(short) or {"short": short, "kann": [], "fix": [], "frei_alle_2_wochen": None}
        m["name"] = _s(r.get("Name")) or short
        try:
            m["pct"] = int(r.get("%") if not pd.isna(r.get("%")) else 100)
        except (TypeError, ValueError):
            m["pct"] = 100
        m["aktiv"] = _b(r.get("Aktiv"))
        m["tl"] = _b(r.get("TL"))
        m["nur_tel"] = _b(r.get("nur TEL"))
        m["aushilfe"] = _b(r.get("Aushilfe"))
        anw = {}
        for dk in DAY_KEYS:
            val = _s(r.get(dk)) or "ganz"
            if val != "ganz":
                anw[dk] = val
        m["anwesend"] = anw
        m["bemerkung"] = _s(r.get("Bemerkung"))
        new_list.append(m)
    cfg["mitarbeiter"] = new_list
    cfg = ak.normalize_config(cfg)
    probs = ak.validate_config(cfg)
    if not probs:
        replace_cfg(cfg)
    return probs


def render_team_people(ver: int) -> None:
    cfg = st.session_state.cfg
    st.caption("Zeile hinzufügen = neue Person (unten «+»). Zeile löschen = Person entfernen "
               "(Zeile markieren, Entf). Deaktivieren statt löschen, wenn jemand nur vorübergehend fehlt. "
               "Anwesenheit pro Tag: ganz · VM · NM · frei. Danach «Übernehmen» klicken.")
    rows = []
    for m in cfg["mitarbeiter"]:
        row = {"Kurzname": m["short"], "Name": m.get("name", ""), "%": int(m.get("pct", 100)),
               "Aktiv": bool(m.get("aktiv", True)), "TL": bool(m.get("tl")),
               "nur TEL": bool(m.get("nur_tel")), "Aushilfe": bool(m.get("aushilfe"))}
        for dk in DAY_KEYS:
            row[dk] = m.get("anwesend", {}).get(dk, "ganz")
        row["Bemerkung"] = m.get("bemerkung", "")
        rows.append(row)
    cols = ["Kurzname", "Name", "%", "Aktiv", "TL", "nur TEL", "Aushilfe"] + DAY_KEYS + ["Bemerkung"]
    df = pd.DataFrame(rows, columns=cols)
    col_config = {
        "Kurzname": st.column_config.TextColumn("Kurzname", help="Eindeutig, wird im Plan und in Rotationen verwendet", width="small"),
        "Name": st.column_config.TextColumn("Name (Anzeige)", help="z.B. Vorname + Initial: «Christina B.»", width="medium"),
        "%": st.column_config.NumberColumn("%", min_value=0, max_value=100, step=10, width="small"),
        "Aktiv": st.column_config.CheckboxColumn("Aktiv", width="small"),
        "TL": st.column_config.CheckboxColumn("TL", help="Teamleitung: Tagesverantwortung, Queue, PHC Montag", width="small"),
        "nur TEL": st.column_config.CheckboxColumn("nur TEL", help="Wird ausschliesslich mit TEL eingeplant", width="small"),
        "Aushilfe": st.column_config.CheckboxColumn("Aushilfe", help="Täglich PO/SCAN (VM) + scanning (NM), keine anderen Rollen", width="small"),
        "Bemerkung": st.column_config.TextColumn("Bemerkung", width="large"),
    }
    for dk in DAY_KEYS:
        col_config[dk] = st.column_config.SelectboxColumn(dk, options=HALBTAG_WERTE, required=True, width="small")
    edited = st.data_editor(df, column_config=col_config, hide_index=True, num_rows="dynamic",
                            width="stretch", key=f"team_editor_{ver}")
    c1, c2 = st.columns([1, 4])
    with c1:
        if st.button("✓  Übernehmen", key="team_apply", type="primary", width="stretch"):
            probs = _apply_people_table(edited)
            if probs:
                st.error("Nicht übernommen: " + "; ".join(probs))
            else:
                st.success("Personen übernommen – unter «Speichern / Laden» dauerhaft sichern.")
                st.rerun()
    with c2:
        st.caption("Neue Personen bekommen noch keine Rollen – diese im Tab «Rollen & Fixeinsätze» setzen.")


def render_team_roles(ver: int) -> None:
    cfg = st.session_state.cfg
    shorts = [m["short"] for m in cfg["mitarbeiter"]]
    if not shorts:
        st.info("Keine Personen vorhanden.")
        return
    labels = {person_label(s): s for s in shorts}
    who = labels[st.selectbox("Person", list(labels.keys()), key=f"roles_person_{ver}")]
    m = next(x for x in cfg["mitarbeiter"] if x["short"] == who)

    st.markdown('<div class="field-label" style="margin-top:8px;">Darf / macht</div>', unsafe_allow_html=True)
    ropts = role_options()
    kann_labels = st.multiselect("Rollen", list(ropts.keys()),
                                 default=[lbl for lbl, r in ropts.items() if r in m.get("kann", [])],
                                 key=f"roles_kann_{ver}_{who}", label_visibility="collapsed")
    kann = [ropts[lbl] for lbl in kann_labels]

    st.markdown('<div class="field-label" style="margin-top:8px;">Alle 2 Wochen frei</div>', unsafe_allow_html=True)
    f2 = m.get("frei_alle_2_wochen") or {}
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        par_opts = ["nie"] + PARITAET
        par = st.selectbox("In KW", par_opts, index=par_opts.index(f2["kw"]) if f2.get("kw") in PARITAET else 0,
                           key=f"roles_par_{ver}_{who}")
    with c2:
        f2_days = st.multiselect("Tage", DAY_KEYS, default=[d for d in f2.get("tage", {}) if d in DAY_KEYS],
                                 key=f"roles_f2days_{ver}_{who}")
    with c3:
        vals = [v for v in f2.get("tage", {}).values()]
        f2_val = st.selectbox("Halbtag", ["ganz", "VM", "NM"], index=["ganz", "VM", "NM"].index(vals[0]) if vals and vals[0] in ("ganz", "VM", "NM") else 0,
                              key=f"roles_f2val_{ver}_{who}")

    st.markdown('<div class="field-label" style="margin-top:8px;">Fixe Einsätze (jede Woche, oder nur in geraden/ungeraden KW)</div>',
                unsafe_allow_html=True)
    fix_df = pd.DataFrame([{"Tag": f.get("tag", "Mo"), "Halbtag": f.get("halbtag", "ganz"),
                            "Task": f.get("task", ""), "KW": f.get("kw") or "jede"} for f in m.get("fix", [])],
                          columns=["Tag", "Halbtag", "Task", "KW"])
    fix_edit = st.data_editor(
        fix_df, hide_index=True, num_rows="dynamic", width="stretch", key=f"roles_fix_{ver}_{who}",
        column_config={
            "Tag": st.column_config.SelectboxColumn("Tag", options=DAY_KEYS, required=True),
            "Halbtag": st.column_config.SelectboxColumn("Halbtag", options=["ganz", "VM", "NM"], required=True),
            "Task": st.column_config.SelectboxColumn("Task", options=[t for t in TASK_OPTIONS if t], required=True),
            "KW": st.column_config.SelectboxColumn("KW", options=["jede"] + PARITAET, required=True),
        })

    if st.button("✓  Übernehmen", key=f"roles_apply_{who}", type="primary"):
        new_cfg = copy.deepcopy(cfg)
        mm = next(x for x in new_cfg["mitarbeiter"] if x["short"] == who)
        mm["kann"] = list(kann)
        mm["frei_alle_2_wochen"] = ({"kw": par, "tage": {d: f2_val for d in f2_days}}
                                    if par in PARITAET and f2_days else None)
        fixes = []
        for _, r in fix_edit.iterrows():
            if not _s(r.get("Task")):
                continue
            fx = {"tag": _s(r.get("Tag")) or "Mo", "halbtag": _s(r.get("Halbtag")) or "ganz", "task": _s(r.get("Task"))}
            if _s(r.get("KW")) in PARITAET:
                fx["kw"] = _s(r.get("KW"))
            fixes.append(fx)
        mm["fix"] = fixes
        probs = ak.validate_config(ak.normalize_config(new_cfg))
        if probs:
            st.error("; ".join(probs))
        else:
            replace_cfg(new_cfg)
            st.success(f"{who} übernommen.")
            st.rerun()


def _rotation_now(lst: list[str], offset: int, widx: int, mult: int = 1) -> int | None:
    if not lst:
        return None
    return (widx * mult + offset) % len(lst)


def render_team_rotation(ver: int) -> None:
    cfg = st.session_state.cfg
    rot = cfg["rotation"]
    rules = cfg["regeln"]
    act = active_shorts()
    tls = [m["short"] for m in active_members() if m.get("tl")]
    widx = info["widx"]
    st.caption(f"Rotationen werden aus der Kalenderwoche berechnet (keine Zustandsdatei). "
               f"Mit «diese Woche (KW {info['kw']})» stellst du ein, wer in der angezeigten Woche dran ist – "
               f"die Folgewochen ergeben sich daraus.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="field-label">Tagesverantwortung – Reihenfolge der TL</div>', unsafe_allow_html=True)
        tv_order = st.multiselect("TV Reihenfolge", tls, default=[n for n in rot["tv_reihenfolge"] if n in tls],
                                  key=f"rot_tv_{ver}", label_visibility="collapsed",
                                  help="Reihenfolge = Rotationsreihenfolge Tag für Tag")
        tv_list = tv_order or tls
        tv_idx = _rotation_now(tv_list, rot["offset"]["tv"], widx, mult=5)
        tv_now = st.selectbox(f"TV am Montag diese Woche (KW {info['kw']})", tv_list,
                              index=tv_idx or 0, key=f"rot_tvnow_{ver}") if tv_list else None

        st.markdown('<div class="field-label" style="margin-top:8px;">PHC-Liste Montag – TL im Wechsel</div>', unsafe_allow_html=True)
        phc_mo = st.multiselect("PHC Montag", tls, default=[n for n in rot["phc_montag_tl"] if n in tls],
                                key=f"rot_phcmo_{ver}", label_visibility="collapsed")
        phc_mo_idx = _rotation_now(phc_mo, rot["offset"]["phc_montag"], widx)
        phc_mo_now = st.selectbox(f"PHC Montag diese Woche (KW {info['kw']})", phc_mo, index=phc_mo_idx or 0,
                                  key=f"rot_phcmonow_{ver}") if phc_mo else None

        st.markdown('<div class="field-label" style="margin-top:8px;">PHC-Liste Freitag – im Wechsel</div>', unsafe_allow_html=True)
        phc_fr = st.multiselect("PHC Freitag", act, default=[n for n in rot["phc_freitag_wechsel"] if n in act],
                                key=f"rot_phcfr_{ver}", label_visibility="collapsed")
        phc_fr_idx = _rotation_now(phc_fr, rot["offset"]["phc_freitag"], widx)
        phc_fr_now = st.selectbox(f"PHC Freitag diese Woche (KW {info['kw']})", phc_fr, index=phc_fr_idx or 0,
                                  key=f"rot_phcfrnow_{ver}") if phc_fr else None

        st.markdown('<div class="field-label" style="margin-top:8px;">Labor Mi NM (ERF5) – im Wechsel</div>', unsafe_allow_html=True)
        labor = st.multiselect("Labor", act, default=[n for n in rot["labor_mi_nm"] if n in act],
                               key=f"rot_labor_{ver}", label_visibility="collapsed")
        labor_idx = _rotation_now(labor, rot["offset"]["labor"], widx)
        labor_now = st.selectbox(f"Labor diese Woche (KW {info['kw']})", labor, index=labor_idx or 0,
                                 key=f"rot_labornow_{ver}") if labor else None

    with c2:
        st.markdown('<div class="field-label">PHC-Liste fix pro Tag (– = aus dem Pool)</div>', unsafe_allow_html=True)
        phc_fix = {}
        pcols = st.columns(5)
        for i, dk in enumerate(DAY_KEYS):
            cur = rot["phc_fix"].get(dk, "–")
            opts = ["–"] + act
            val = pcols[i].selectbox(dk, opts, index=opts.index(cur) if cur in opts else 0, key=f"rot_phcfix_{ver}_{dk}")
            if val != "–":
                phc_fix[dk] = val

        st.markdown('<div class="field-label" style="margin-top:8px;">PHC-Liste: an diesen Tagen ausgeschlossen</div>', unsafe_allow_html=True)
        phc_ex = {}
        for dk in DAY_KEYS:
            sel = st.multiselect(dk, act, default=[n for n in rot["phc_tag_ausschluss"].get(dk, []) if n in act],
                                 key=f"rot_phcex_{ver}_{dk}")
            if sel:
                phc_ex[dk] = sel

        st.markdown('<div class="field-label" style="margin-top:8px;">Postöffnung Dienstag bevorzugt</div>', unsafe_allow_html=True)
        po_opts = ["–"] + act
        cur = rot.get("po_dienstag_bevorzugt") or "–"
        po_pref = st.selectbox("PO Di", po_opts, index=po_opts.index(cur) if cur in po_opts else 0,
                               key=f"rot_popref_{ver}", label_visibility="collapsed")

        st.markdown('<div class="field-label" style="margin-top:8px;">Schalter-Gruppe diese Woche</div>', unsafe_allow_html=True)
        sch_pool = [m["short"] for m in active_members() if "Schalter" in m.get("kann", [])
                    and not m.get("nur_tel") and not m.get("aushilfe")]
        mid = (len(sch_pool) + 1) // 2
        grp_now = (widx + rot["offset"]["schalter"]) % 2
        grp_labels = [f"Gruppe A: {', '.join(sch_pool[:mid]) or '–'}", f"Gruppe B: {', '.join(sch_pool[mid:]) or '–'}"]
        sch_now = st.selectbox("Schalter", grp_labels, index=grp_now, key=f"rot_schnow_{ver}", label_visibility="collapsed")

        st.markdown('<div class="field-label" style="margin-top:8px;">Wochenaufgaben diese Woche</div>', unsafe_allow_html=True)
        wa_now = {}
        for label in ("Direkt", "ONB", "BTM"):
            pool = [m["short"] for m in active_members() if label in m.get("kann", [])
                    and not m.get("nur_tel") and not m.get("aushilfe")]
            if pool:
                idx = (widx + rot["offset"].get(label, 0)) % len(pool)
                wa_now[label] = (pool, st.selectbox(label, pool, index=idx, key=f"rot_wa_{ver}_{label}"))
            else:
                st.caption(f"{label}: niemand berechtigt")

    st.markdown('<div class="field-label" style="margin-top:12px;">Besetzungsregeln</div>', unsafe_allow_html=True)
    r1, r2, r3, r4, r5 = st.columns(5)
    tel_mo = r1.number_input("TEL Mo VM", 0, 10, int(rules["tel_montag_vm"]), key=f"rule_telmo_{ver}")
    tel_s = r2.number_input("TEL sonst", 0, 10, int(rules["tel_sonst"]), key=f"rule_tels_{ver}")
    abkl = r3.number_input("ABKL / Halbtag", 0, 10, int(rules["abkl_pro_halbtag"]), key=f"rule_abkl_{ver}")
    po_z = r4.number_input("Zusatz-PO am VM", 0, 5, int(rules["po_zusatz_vm"]), key=f"rule_po_{ver}",
                           help="Zusätzlich zur PO/SCAN-Person (Aushilfe)")
    hub = r5.number_input("HUB / Tag", 0, 2, int(rules["hub_pro_tag"]), key=f"rule_hub_{ver}")
    r6, r7, r8, r9 = st.columns(4)
    rx = r6.number_input("RX Abo / Woche", 0, 10, int(rules["rx_abo_pro_woche"]), key=f"rule_rx_{ver}")
    scan_max = r7.number_input("Scanning max. Halbtage/Woche", 0, 10, int(rules["scan_max_halbtage_pro_woche"]), key=f"rule_scanmax_{ver}")
    scan_p = r8.slider("Scanning-Anteil", 0.0, 1.0, float(rules["scan_anteil"]), 0.05, key=f"rule_scanp_{ver}",
                       help="Anteil der freien Halbtage, die bei Scanning-Berechtigten zu «scanning» werden")
    scan_nm = r9.checkbox("NM-Scan-Ersatz, wenn Aushilfe fehlt", bool(rules["scan_nm_ersatz_wenn_aushilfe_fehlt"]), key=f"rule_scannm_{ver}")
    st.caption("Schalter-Reserve, Schalter-Code, TL am Telefon und TAGES PA mit Telefon: Tab «Entscheidungen».")

    if st.button("✓  Übernehmen", key="rot_apply", type="primary"):
        new_cfg = copy.deepcopy(cfg)
        nr = new_cfg["rotation"]
        nr["tv_reihenfolge"] = list(tv_order)
        nr["phc_montag_tl"] = list(phc_mo)
        nr["phc_freitag_wechsel"] = list(phc_fr)
        nr["labor_mi_nm"] = list(labor)
        nr["phc_fix"] = phc_fix
        nr["phc_tag_ausschluss"] = phc_ex
        nr["po_dienstag_bevorzugt"] = "" if po_pref == "–" else po_pref
        off = nr["offset"]
        if tv_list and tv_now in tv_list:
            off["tv"] = (tv_list.index(tv_now) - widx * 5) % len(tv_list)
        if phc_mo and phc_mo_now in phc_mo:
            off["phc_montag"] = (phc_mo.index(phc_mo_now) - widx) % len(phc_mo)
        if phc_fr and phc_fr_now in phc_fr:
            off["phc_freitag"] = (phc_fr.index(phc_fr_now) - widx) % len(phc_fr)
        if labor and labor_now in labor:
            off["labor"] = (labor.index(labor_now) - widx) % len(labor)
        off["schalter"] = (grp_labels.index(sch_now) - widx) % 2
        for label, (pool, chosen) in wa_now.items():
            off[label] = (pool.index(chosen) - widx) % len(pool)
        new_cfg["regeln"].update({
            "tel_montag_vm": int(tel_mo), "tel_sonst": int(tel_s), "abkl_pro_halbtag": int(abkl),
            "po_zusatz_vm": int(po_z), "hub_pro_tag": int(hub), "rx_abo_pro_woche": int(rx),
            "scan_max_halbtage_pro_woche": int(scan_max), "scan_anteil": float(scan_p),
            "scan_nm_ersatz_wenn_aushilfe_fehlt": bool(scan_nm),
        })
        probs = ak.validate_config(ak.normalize_config(new_cfg))
        if probs:
            st.error("; ".join(probs))
        else:
            replace_cfg(new_cfg)
            st.success("Rotation & Regeln übernommen.")
            st.rerun()


def render_team_decisions(ver: int) -> None:
    """Punkte, die das Team nicht eindeutig geregelt hat – hier entscheidet die Planerin."""
    cfg = st.session_state.cfg
    rules = cfg["regeln"]
    act = active_shorts()
    st.markdown("Diese Punkte sind im Team **nicht eindeutig geregelt**. Hier legst du fest, wie der "
                "Generator damit umgeht. Die Einstellung wirkt ab der nächsten Generierung; "
                "«Speichern / Laden» macht sie dauerhaft.")

    st.markdown("**1 · Schalter: «Brigitte, Saskia, Dragi nicht alle 2 Wochen»**")
    st.caption("Personen mit Rolle «Schalter» bilden zwei Gruppen, die sich wöchentlich abwechseln "
               "(jede Person etwa alle 2 Wochen einen ganzen Tag). Reserve-Personen kommen nur dran, "
               "wenn aus der aktiven Gruppe niemand ganztags frei ist – also seltener. "
               "Rolle «Schalter» wird bei Reserve-Personen entfernt.")
    c1, c2, c3 = st.columns([2, 1.4, 1])
    with c1:
        reserve = st.multiselect("Schalter-Reserve", act,
                                 default=[m["short"] for m in active_members() if "Schalter-Reserve" in m.get("kann", [])],
                                 key=f"dec_reserve_{ver}")
    with c2:
        cur = rules.get("schalter_reserve_modus", SCHALTER_RESERVE_MODI[0])
        modus = st.selectbox("Reserve einsetzen …", SCHALTER_RESERVE_MODI,
                             index=SCHALTER_RESERVE_MODI.index(cur) if cur in SCHALTER_RESERVE_MODI else 0,
                             key=f"dec_modus_{ver}",
                             help="«vor der anderen Gruppe»: Reserve kommt vor der Gruppe dran, die diese Woche nicht aktiv ist. "
                                  "«nie»: Reserve-Personen machen keinen Schalter.")
    with c3:
        cur = rules.get("schalter_code", SCHALTER_CODES[0])
        code = st.selectbox("Code im Plan", SCHALTER_CODES, index=SCHALTER_CODES.index(cur) if cur in SCHALTER_CODES else 0,
                            key=f"dec_code_{ver}", help="April-Liste: ERF4/SCH · seit Mai im Plan: ERF7/SCH")

    st.markdown("**2 · TL 1 bis 2× pro Woche am Telefon**")
    st.caption("Der Generator reserviert für jede TL vorab so viele Telefon-Halbtage wie das Minimum verlangt "
               "(bevor Schalter, ERF9, RX Abo und ERF8 verteilt werden) und gibt TL danach den Vorzug, bis das "
               "Maximum erreicht ist. Unter- oder Überschreitungen stehen im Ergebnis als Hinweis. 0 = kein Minimum.")
    c4, c5, _ = st.columns([1, 1, 3])
    tl_min = c4.number_input("TL mind. TEL/Woche", 0, 10, int(rules.get("tl_tel_min_pro_woche", 1)), key=f"dec_tlmin_{ver}")
    tl_max = c5.number_input("TL max. TEL/Woche", 0, 10, int(rules.get("tl_tel_max_pro_woche", 2)), key=f"dec_tlmax_{ver}")

    st.markdown("**3 · TAGES PA am Nachmittag mit Telefon (TAGES PA/TEL)**")
    st.caption("«bei Bedarf»: nur wenn die TEL-Besetzung am Nachmittag sonst nicht reicht · "
               "«immer»: jeden Nachmittag · «nie».")
    cur = rules.get("tpa_tel_nm", TPA_TEL_MODI[0])
    tpa_mode = st.selectbox("TAGES PA/TEL", TPA_TEL_MODI, index=TPA_TEL_MODI.index(cur) if cur in TPA_TEL_MODI else 0,
                            key=f"dec_tpa_{ver}", label_visibility="collapsed")

    st.markdown("**4 · Nur manuell**")
    st.caption("Diese Codes vergibt der Generator nie, sie stehen aber im Plan-Editor (Schritt 2) und bei "
               "«Custom»-Einträgen zur Auswahl: " + ", ".join(NUR_MANUELL) + ". "
               "PO/TEL am Nachmittag erzeugt im Ergebnis den Notfall-Hinweis (muss früher vom Telefon weg). "
               "Vorbezüge (VBZ) sind als Rolle hinterlegt, werden aber nicht geplant.")

    if st.button("✓  Übernehmen", key="decide_apply", type="primary"):
        new_cfg = copy.deepcopy(cfg)
        for m in new_cfg["mitarbeiter"]:
            kann = [r for r in m.get("kann", []) if r != "Schalter-Reserve"]
            if m["short"] in reserve:
                kann = [r for r in kann if r != "Schalter"] + ["Schalter-Reserve"]
            m["kann"] = kann
        new_cfg["regeln"].update({
            "schalter_reserve_modus": modus, "schalter_code": code,
            "tl_tel_min_pro_woche": int(tl_min), "tl_tel_max_pro_woche": int(tl_max),
            "tpa_tel_nm": tpa_mode,
        })
        probs = ak.validate_config(ak.normalize_config(new_cfg))
        if probs:
            st.error("; ".join(probs))
        else:
            replace_cfg(new_cfg)
            st.success("Entscheidungen übernommen.")
            st.rerun()


def render_team_backup(ver: int) -> None:
    cfg = st.session_state.cfg
    st.caption("Diese Zeilen erscheinen unterhalb des Teams und werden nie automatisch geplant. "
               "«Standard» steht in jeder Zelle (z.B. ERP beim SPP-Team); manuelle Einträge über "
               "Abwesenheit/Sonderwunsch mit Grund «Custom».")
    df = pd.DataFrame([{"Kurzname": b["short"], "Name": b.get("name", ""), "Standard": b.get("standard", "")}
                       for b in cfg["backup_team"]], columns=["Kurzname", "Name", "Standard"])
    edited = st.data_editor(df, hide_index=True, num_rows="dynamic", width="stretch", key=f"backup_editor_{ver}")
    if st.button("✓  Übernehmen", key="backup_apply", type="primary"):
        new_cfg = copy.deepcopy(cfg)
        new_cfg["backup_team"] = [
            {"short": _s(r.get("Kurzname")), "name": _s(r.get("Name")) or _s(r.get("Kurzname")), "standard": _s(r.get("Standard"))}
            for _, r in edited.iterrows() if _s(r.get("Kurzname"))
        ]
        probs = ak.validate_config(ak.normalize_config(new_cfg))
        if probs:
            st.error("; ".join(probs))
        else:
            replace_cfg(new_cfg)
            st.success("BackUp-Team übernommen.")
            st.rerun()


def render_team_save(ver: int) -> None:
    gh = github_conf()
    if gh:
        st.markdown(f"🔗 GitHub-Speicherung aktiv: `{gh['repo']}` @ `{gh['branch']}` – "
                    "Speichern schreibt team.json und abwesenheiten.json ins Repo (Streamlit Cloud lädt danach neu).")
    else:
        st.markdown("ℹ️ Keine GitHub-Speicherung konfiguriert: Änderungen werden lokal gespeichert "
                    "(auf Streamlit Cloud nur bis zum nächsten Neustart). Für dauerhafte Speicherung "
                    "`team.json` herunterladen und ins Repo committen – oder in den App-Secrets "
                    "`[github] token / repo / branch` hinterlegen (siehe README).")
    c1, c2, c3, c4 = st.columns([1.3, 1.3, 1.3, 1.3])
    with c1:
        if st.button("💾  Speichern", key="save_all", type="primary", width="stretch",
                     disabled=not (cfg_dirty or entries_dirty)):
            st.session_state.save_msgs = save_everything(push_github=bool(gh))
            st.rerun()
    with c2:
        st.download_button("⬇  team.json", data=ak.config_to_json(st.session_state.cfg), file_name="team.json",
                           mime="application/json", width="stretch", key=f"dl_cfg_{ver}")
    with c3:
        st.download_button("⬇  abwesenheiten.json", data=entries_json(), file_name="abwesenheiten.json",
                           mime="application/json", width="stretch", key=f"dl_abs_{ver}")
    with c4:
        if st.button("↺  Gespeicherte Version laden", key="reload_cfg", width="stretch"):
            try:
                replace_cfg(ak.load_config(CONFIG_PATH))
                st.session_state.cfg_saved_json = ak.config_to_json(st.session_state.cfg)
                st.session_state.entries = load_entries_from_disk()
                st.session_state.entries_saved_json = entries_json()
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
    for ok, msg in st.session_state.get("save_msgs", []):
        (st.success if ok else st.warning)(msg)
    if st.session_state.get("save_msgs"):
        st.session_state.save_msgs = []

    st.markdown('<div class="field-label" style="margin-top:12px;">team.json hochladen (ersetzt die aktuelle Konfiguration)</div>',
                unsafe_allow_html=True)
    up = st.file_uploader("team.json", type=["json"], key=f"upload_cfg_{ver}", label_visibility="collapsed")
    if up is not None:
        try:
            data = json.loads(up.getvalue().decode("utf-8"))
            cfg_new = ak.normalize_config(data)
            probs = ak.validate_config(cfg_new)
            if probs:
                st.error("Datei nicht übernommen: " + "; ".join(probs))
            else:
                replace_cfg(cfg_new)
                st.success("Konfiguration geladen – jetzt «Speichern», damit sie bleibt.")
                st.rerun()
        except Exception as exc:
            st.error(f"Ungültige Datei: {exc}")
    probs = ak.validate_config(st.session_state.cfg)
    if probs:
        st.warning("Konfigurationsprobleme: " + "; ".join(probs))


def render_team_management() -> None:
    n_total = len(st.session_state.cfg["mitarbeiter"])
    label = f"👥  Team verwalten · {n_active}/{n_total} aktiv · Personen, Anwesenheit, Rollen, Rotation, Regeln"
    if cfg_dirty:
        label += "  ·  ⚠️ ungespeichert"
    with st.expander(label, expanded=False):
        ver = st.session_state.cfg_version
        tabs = st.tabs(["Personen", "Rollen & Fixeinsätze", "Rotation & Regeln", "Entscheidungen",
                        "BackUp Team", "Speichern / Laden"])
        with tabs[0]:
            render_team_people(ver)
        with tabs[1]:
            render_team_roles(ver)
        with tabs[2]:
            render_team_rotation(ver)
        with tabs[3]:
            render_team_decisions(ver)
        with tabs[4]:
            render_team_backup(ver)
        with tabs[5]:
            render_team_save(ver)


# ══════════════════════════════════════════════════════════════════════
#  Schritt 1 – Planung
# ══════════════════════════════════════════════════════════════════════

def render_planung() -> None:
    render_team_management()
    left, right = st.columns([1.25, 1])

    with left:
        st.markdown('<div class="card"><h2>Abwesenheit oder Sonderwunsch</h2>'
                    '<div class="sub">Krank · Ferien · Termin · Home Office · Custom (Task-Code) – auch für das BackUp-Team</div>',
                    unsafe_allow_html=True)
        f1, f2 = st.columns(2)
        with f1:
            st.markdown('<div class="field-label">Mitarbeiterin</div>', unsafe_allow_html=True)
            st.selectbox("Mitarbeiterin", list(employee_options().keys()), key="f_employee",
                         label_visibility="collapsed")
        with f2:
            st.markdown('<div class="field-label">Grund</div>', unsafe_allow_html=True)
            st.selectbox("Grund", list(ABSENCE_TYPES.keys()), key="f_type", label_visibility="collapsed")

        st.markdown('<div class="field-label" style="margin-top:8px;">Halbtag</div>', unsafe_allow_html=True)
        seg_cols = st.columns(3)
        for i, lbl in enumerate(["Ganzer Tag", "Vormittag", "Nachmittag"]):
            with seg_cols[i]:
                if st.button(lbl, key=f"seg_{i}", width="stretch",
                             type="primary" if st.session_state.sel_slot == lbl else "secondary"):
                    st.session_state.sel_slot = lbl
                    st.rerun()

        st.markdown('<div class="field-label" style="margin-top:10px;">Notiz <span style="font-weight:500; '
                    'text-transform:none; letter-spacing:0; color:var(--text-mute);">optional · bei Custom = Task-Code</span></div>',
                    unsafe_allow_html=True)
        st.text_input("Notiz", key="f_note", placeholder="z.B. Arzttermin 14:30 – oder Task-Code (ERP, HO, TEL …) bei Custom",
                      label_visibility="collapsed")

        st.markdown('<div class="field-label" style="margin-top:10px;">Tage</div>', unsafe_allow_html=True)
        d_cols = st.columns(5)
        for i, lbl in enumerate(DAY_KEYS):
            with d_cols[i]:
                if st.button(lbl, key=f"day_{i}", width="stretch",
                             type="primary" if i in st.session_state.sel_days else "secondary"):
                    if i in st.session_state.sel_days:
                        st.session_state.sel_days.remove(i)
                    else:
                        st.session_state.sel_days.add(i)
                    st.rerun()

        st.markdown('<div class="field-label" style="margin-top:10px;">Anzahl Wochen <span style="font-weight:500; '
                    'text-transform:none; letter-spacing:0; color:var(--text-mute);">aufeinander folgend, ab der angezeigten KW</span></div>',
                    unsafe_allow_html=True)
        st.number_input("Anzahl Wochen", min_value=1, max_value=52, step=1, key="f_weeks", label_visibility="collapsed")

        a_cols = st.columns([1, 1, 4, 2])
        with a_cols[0]:
            if st.button("Alle", key="day_all", width="stretch"):
                st.session_state.sel_days = {0, 1, 2, 3, 4}
                st.rerun()
        with a_cols[1]:
            if st.button("Keine", key="day_none", width="stretch"):
                st.session_state.sel_days = set()
                st.rerun()
        with a_cols[3]:
            if st.button("＋  Hinzufügen", key="add_entry", type="primary", width="stretch"):
                if not st.session_state.sel_days:
                    st.warning("Bitte mindestens einen Tag auswählen.")
                elif not st.session_state.f_employee:
                    st.warning("Keine Person ausgewählt.")
                else:
                    sl = st.session_state.sel_slot
                    note = (st.session_state.f_note or "").strip()
                    f_type = st.session_state.f_type
                    st.session_state.entries.append({
                        "name": employee_options().get(st.session_state.f_employee, st.session_state.f_employee),
                        "type": f_type,
                        "days": sorted(st.session_state.sel_days),
                        "slots": [0, 1] if sl == "Ganzer Tag" else [0] if sl == "Vormittag" else [1],
                        "note": note,
                        "task": note if (f_type == "Custom" and note) else ("HO" if f_type == "Custom" else ""),
                        "jahr": info["iso_year"], "kw_start": info["kw"],
                        "weeks": max(1, int(st.session_state.f_weeks or 1)),
                    })
                    write_local(ABSENCE_PATH, entries_json())
                    st.session_state.sel_days = set()
                    st.session_state["_pending_clear_note"] = True
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        entries = st.session_state.entries
        n = len(entries)
        n_week = sum(1 for e in entries if entry_active(e, info["widx"]))
        st.markdown(
            f'<div class="card"><div class="card-head"><div><h2 style="margin-bottom:0;">Einträge</h2>'
            f'<div class="sub" style="margin-bottom:0;">{n_week} für KW {info["kw"]} · {n} insgesamt</div></div></div>',
            unsafe_allow_html=True,
        )
        if n == 0:
            st.markdown('<div class="empty-state"><div class="icon">📅</div><div class="ttl">Keine Abwesenheiten erfasst.</div>'
                        '<div>Diese Woche läuft wie geplant.</div></div>', unsafe_allow_html=True)
        else:
            for idx, e in enumerate(entries):
                type_key = ABSENCE_TYPES.get(e["type"], "CUSTOM")
                days_str = ", ".join(DAY_KEYS[d] for d in e["days"])
                slot_str = "ganzer Tag" if len(e["slots"]) == 2 else ("Vormittag" if e["slots"][0] == 0 else "Nachmittag")
                note_html = f'<span class="entry-note">· „{html_escape(e["note"])}"</span>' if e.get("note") else ""
                opacity = "" if entry_active(e, info["widx"]) else 'style="opacity:0.45;"'
                l, r = st.columns([12, 1])
                with l:
                    st.markdown(
                        f'<div class="entry-row" {opacity}><span class="entry-dot {type_key}"></span>'
                        f'<span class="entry-name">{html_escape(e["name"])}</span>'
                        f'<span class="entry-meta">· {html_escape(e["type"])} · {days_str} ({slot_str}) · {entry_kw_label(e)}</span>'
                        f'{note_html}</div>',
                        unsafe_allow_html=True,
                    )
                with r:
                    if st.button("✕", key=f"rm_{idx}", help="Entfernen", width="stretch"):
                        st.session_state.entries.pop(idx)
                        write_local(ABSENCE_PATH, entries_json())
                        st.rerun()
            b1, b2 = st.columns([1, 1])
            with b1:
                if st.button("🗑  Vergangene Einträge aufräumen", key="purge_old", width="stretch"):
                    st.session_state.entries = [e for e in entries if entry_widx_start(e) + max(1, int(e.get("weeks", 1))) > info["widx"]]
                    write_local(ABSENCE_PATH, entries_json())
                    st.rerun()
            with b2:
                if st.button("💾  Einträge speichern", key="save_entries", width="stretch", disabled=not entries_dirty,
                             help="Lokal und – falls konfiguriert – ins GitHub-Repo"):
                    st.session_state.save_msgs = save_everything(push_github=bool(github_conf()))
                    st.rerun()
            for ok, msg in st.session_state.get("save_msgs", []):
                (st.success if ok else st.warning)(msg)
            if st.session_state.get("save_msgs"):
                st.session_state.save_msgs = []
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
#  Schritt 2 – Ergebnis
# ══════════════════════════════════════════════════════════════════════

def render_banner(kind: str, title: str, items: list[str]) -> None:
    lis = "".join(f"<li>{html_escape(i)}</li>" for i in items)
    mark = {"ok": "✓", "warn": "!", "error": "!", "info": "i"}[kind]
    st.markdown(f'<div class="banner banner-{kind}"><div class="check">{mark}</div><div><h4>{html_escape(title)}</h4>'
                f'{"<ul>" + lis + "</ul>" if items else ""}</div></div>', unsafe_allow_html=True)


def render_plan_editor(res: dict) -> None:
    sched = st.session_state.get("sched")
    if sched is None:
        return
    slot_cols = [f"{DAY_KEYS[d]} {SLOT_KEYS[s]}" for d in range(5) for s in range(2)]
    with st.expander("✏️  Plan bearbeiten · Zellen per Dropdown ändern (Team und BackUp-Team)", expanded=False):
        st.caption("Pro Halbtag eine Tätigkeit wählen, leer = frei. Nach «Übernehmen» werden Tabelle, "
                   "Regelprüfung und Excel aktualisiert. Wochenwechsel via ‹ › generiert komplett neu.")
        all_rows = res["rows"] + res["backup_rows"]
        df_rows = []
        for r in all_rows:
            row = {"Name": r["full_name"], "%": f"{r['pct']}%" if r["pct"] is not None else "BackUp"}
            for i, col in enumerate(slot_cols):
                t = r["cells"][i]["task"] or ""
                row[col] = t[1:] if t.startswith("*") else t
            df_rows.append(row)
        df = pd.DataFrame(df_rows, columns=["Name", "%"] + slot_cols)
        col_config = {"Name": st.column_config.TextColumn("Name", disabled=True, width="medium"),
                      "%": st.column_config.TextColumn("%", disabled=True, width="small")}
        for col in slot_cols:
            col_config[col] = st.column_config.SelectboxColumn(col, options=TASK_OPTIONS, required=False, width="small")
        edited = st.data_editor(df, column_config=col_config, hide_index=True, num_rows="fixed", width="stretch",
                                key=f"plan_editor_{res['iso_year']}_{res['kw']}")
        if not st.button("✓  Änderungen übernehmen", key=f"apply_edits_{res['kw']}", type="primary"):
            return
        changes = 0
        backups = set(backup_shorts())
        for ridx, r in enumerate(all_rows):
            short = r["short"]
            for cidx, col in enumerate(slot_cols):
                new_val = _s(edited.at[ridx, col])
                old_val = r["cells"][cidx]["task"] or ""
                if old_val.startswith("*"):
                    old_val = old_val[1:]
                if new_val == old_val:
                    continue
                d, s = divmod(cidx, 2)
                if short in backups:
                    sched.backup_tasks[(short, d, s)] = new_val
                elif short in sched.schedule:
                    sched.schedule[short][(d, s)] = new_val
                changes += 1
        if changes == 0:
            st.info("Keine Änderungen erkannt.")
            return
        # Zähler neu aufbauen, damit die Regelprüfung stimmt
        sched.tel_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        sched.abkl_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        sched.erf8_assigned = {d: False for d in range(5)}
        for name, cells in sched.schedule.items():
            for (d, s), t in cells.items():
                sched._count(d, s, t)
        refresh_result_from_sched(sched)
        st.success(f"{changes} Änderung(en) übernommen.")
        st.rerun()


def render_ergebnis() -> None:
    res = st.session_state.result
    if not res:
        st.info("Noch kein Plan generiert. Klicke auf **Plan generieren**.")
        return

    hard = [i for i in res["issues"] if i.startswith("❌")]
    soft = [i for i in res["issues"] if i.startswith("⚠️")]
    infos = [i for i in res["issues"] if not i.startswith("❌") and not i.startswith("⚠️")]
    if res.get("config_problems"):
        render_banner("warn", "Konfiguration prüfen (Team verwalten)", res["config_problems"])
    if hard:
        render_banner("error", "Harte Regeln nicht vollständig einhaltbar", hard)
    else:
        render_banner("ok", "Alle harten Regeln eingehalten.", [])
    if soft:
        render_banner("warn", "Hinweise", soft)
    if res["warnings"]:
        render_banner("warn", "Auswirkungen der erfassten Abwesenheiten", res["warnings"])
    if infos:
        render_banner("info", "Vertretungen und Rotation", infos)

    st.markdown('<div class="card"><div class="card-head"><h2 style="margin-bottom:0;">Wochenplan</h2></div>', unsafe_allow_html=True)
    tvc = ak.tv_colors(st.session_state.cfg)
    parts = ['<div class="plan-wrap"><table class="plan"><thead><tr>',
             '<th class="team-head" colspan="2" rowspan="2"><div class="ttl">TEAM</div>'
             '<div class="legend">TV · Tagesverantwortung · dunkelblau = PHC-Liste</div></th>']
    for d in range(5):
        tv = str(res["tagesverantwortung"][d])
        is_open = "OFFEN" in tv
        color = tvc.get(tv)
        style = f'style="background:#{color}; color:{_fg_for(color)};"' if color else ""
        parts.append(f'<th class="day-head" colspan="2"><span class="day-name">{DAYS[d].upper()}</span>'
                     f'<span class="tv-pill{" open" if is_open else ""}"><span class="tv-tag" {style}>TV</span>'
                     f'<span class="tv-name">{html_escape(tv)}</span></span></th>')
    parts.append("</tr><tr>")
    for _ in range(5):
        parts.append('<th class="slot-head">VM</th><th class="slot-head">NM</th>')
    parts.append("</tr></thead><tbody>")

    def row_html(r):
        out = [f'<tr><td class="name-cell">{html_escape(r["full_name"])}</td>',
               f'<td class="pct-cell">{str(r["pct"]) + "%" if r["pct"] is not None else ""}</td>']
        for c in r["cells"]:
            task = c["task"] or ""
            note = f'<span class="pill-note">{html_escape(c["note"])}</span>' if c.get("note") else ""
            if not task and not c.get("note"):
                out.append('<td class="task-cell"></td>')
            elif c["is_phc"]:
                out.append(f'<td class="task-cell"><span class="pill pill-phc">{html_escape(task)}</span>{note}</td>')
            elif task.startswith("*") or (not task and c.get("note")):
                out.append(f'<td class="task-cell"><span class="pill pill-termin">{html_escape(task[1:] or c["note"])}</span></td>')
            else:
                bg, fg = pill_for_task(task)
                out.append(f'<td class="task-cell"><span class="pill" style="background:{bg}; color:{fg};">{html_escape(task)}</span>{note}</td>')
        out.append("</tr>")
        return "".join(out)

    for r in res["rows"]:
        parts.append(row_html(r))
    if res["backup_rows"]:
        parts.append('<tr><td class="sep-cell" colspan="12">BackUp Team · nur manuell</td></tr>')
        for r in res["backup_rows"]:
            parts.append(row_html(r))
    parts.append("</tbody></table></div>")

    # Tagesrollen
    parts.append('<table class="roles"><thead><tr><th></th>' + "".join(f"<th>{DAY_KEYS[d]}</th>" for d in range(5)) + "</tr></thead><tbody>")
    for lbl, key in (("Tagesverantwortung", "tagesverantwortung"), ("PHC-Liste (18 Uhr)", "phc_liste"),
                     ("TAGES PA", "tagespa"), ("Postöffnung (PO/SCAN)", "po_liste"), ("Zusatz-PO", "po_zusatz"),
                     ("Schalter", "schalter"), ("ERF9", "erf9")):
        cells = "".join(f'<td class="{"open" if "OFFEN" in str(v) else ""}">{html_escape(v)}</td>' for v in res[key])
        parts.append(f'<tr><td class="lbl">{lbl}</td>{cells}</tr>')
    parts.append("</tbody></table>")

    wa = res["wochenaufgaben"]
    parts.append('<div class="wa-strip">' + "".join(
        f'<div class="wa-card"><div class="wa-icon">{ico}</div><div class="wa-meta"><div class="lbl">{lbl}</div>'
        f'<div class="val">{html_escape(wa[k] or "–")}</div></div></div>'
        for ico, lbl, k in (("📦", "Direkt", "direkt"), ("🆕", "ONB", "onb"), ("💊", "BTM", "btm"))) + "</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    render_plan_editor(res)

    with st.expander("Besetzung · TEL und ABKL pro Halbtag"):
        for i in range(0, len(res["coverage"]), 5):
            cols = st.columns(5)
            for col, c in zip(cols, res["coverage"][i:i + 5]):
                tel_color = "#16a34a" if c["tel_ok"] else "#dc2626"
                abkl_color = "#16a34a" if c["abkl_ok"] else "#dc2626"
                col.markdown(
                    f'<div style="background:#f8fafc; border:1px solid #e3e8f0; border-radius:8px; padding:10px 12px; font-size:12px;">'
                    f'<div style="display:flex; justify-content:space-between; color:#475569;"><strong style="color:#0b1020;">{c["day"]}</strong><span>{c["slot"]}</span></div>'
                    f'<div style="display:flex; justify-content:space-between; margin-top:4px;"><span>TEL</span><span style="color:{tel_color}; font-weight:700;">{c["tel"]}/{c["tel_target"]}</span></div>'
                    f'<div style="display:flex; justify-content:space-between;"><span>ABKL</span><span style="color:{abkl_color}; font-weight:700;">{c["abkl"]}/{c["abkl_target"]}</span></div></div>',
                    unsafe_allow_html=True,
                )

    with st.expander("Legende · Farben wie im Excel"):
        items = []
        for row in ak.legend_rows(st.session_state.cfg):
            for code, expl in row:
                if not code:
                    continue
                bg = f"#{ak.PHC_COLOR}" if code == "PHC-Liste" else pill_for_task(code)[0]
                fg = "#FFFFFF" if code == "PHC-Liste" else pill_for_task(code)[1]
                items.append(f'<span class="legend-item"><span class="pill" style="background:{bg}; color:{fg};">{html_escape(code)}</span>{html_escape(expl)}</span>')
        st.markdown('<div class="legend-wrap">' + "".join(items) + "</div>", unsafe_allow_html=True)


if st.session_state.step == 1 or st.session_state.result is None:
    render_planung()
else:
    render_ergebnis()
