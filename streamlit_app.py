#!/usr/bin/env python3
"""
KSC Arbeitsplan – Streamlit-App (Version 7.0, Oktober 2026)
============================================================
Oberfläche für https://arbeitsplan-ksc.streamlit.app

- Schritt 1 «Planung»: Abwesenheiten/Sonderwünsche erfassen (auch über
  mehrere Wochen).
- Seite «Team & Regeln»: Personen, Pensen, Anwesenheit,
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
    DAYS, DAY_KEYS, HALBTAG_WERTE, NUR_MANUELL, ROLLEN, SCHALTER_CODES,
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

div.stButton button[kind="primary"], div.stFormSubmitButton button[kind="primaryFormSubmit"] { background: var(--primary) !important; border: 1px solid var(--primary) !important;
  color: white !important; font-weight: 600 !important; border-radius: 9px !important; }
div.stButton button[kind="primary"]:hover, div.stFormSubmitButton button[kind="primaryFormSubmit"]:hover { background: var(--primary-2) !important; border-color: var(--primary-2) !important; }
div.stButton button[kind="secondary"] { background: var(--surface) !important; border: 1px solid var(--border) !important;
  color: var(--text) !important; border-radius: 8px !important; font-weight: 500 !important; }
div.stButton button[kind="secondary"]:hover { background: var(--surface-2) !important; border-color: var(--border-2) !important; }
div.stButton > button, div.stButton > button > div, div.stButton > button p { white-space: nowrap !important; }
div.stButton button:disabled { opacity: .45; cursor: not-allowed; }
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

.page-head h2 { font-size: 18px; font-weight: 700; color: var(--heading); margin: 0 0 2px; padding: 0; }
.page-head .sub { font-size: 12.5px; color: var(--text-mute); margin-bottom: 10px; }
.st-key-savebar_ok { background: var(--ok-soft); border-color: rgba(22,163,74,.3) !important; color: #14532d; }
.st-key-savebar_dirty { background: var(--warn-soft); border-color: rgba(217,119,6,.35) !important; color: #78350f; }
.st-key-savebar_ok p, .st-key-savebar_dirty p { margin-bottom: 0; font-size: 13.5px; }
.team-hint { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 9px 14px;
  font-size: 13px; color: var(--text-dim); }
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
            return m.get("name") or short
    for b in cfg["backup_team"]:
        if b["short"] == short:
            return f"{b.get('name') or short} (BackUp)"
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
hcol = st.columns([2.0, 3.6, 2.4, 1.0, 1.6])

with hcol[0]:
    st.markdown(
        '<div class="brand"><div class="brand-icon">📅</div><div>'
        '<div class="brand-title">Arbeitsplan</div>'
        '<div class="brand-sub">KSC · Team-Scheduling</div></div></div>',
        unsafe_allow_html=True,
    )

with hcol[1]:
    s1, s2, s3 = st.columns([1, 1, 1.25])
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
    with s3:
        if st.button("⚙  Team & Regeln" + ("  •" if cfg_dirty else ""), key="step3", width="stretch",
                     type="primary" if st.session_state.step == 3 else "secondary",
                     help="Personen, Rollen, feste Einsätze, Rotationen und Regeln" +
                          (" – ungespeicherte Änderungen" if cfg_dirty else "")):
            st.session_state.step = 3
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
#  Team & Regeln (eigene Seite)
# ══════════════════════════════════════════════════════════════════════
#  Jede Änderung wird sofort in die Sitzung übernommen (Callback → commit)
#  und gilt für den nächsten Plan; «Speichern» in der Leiste oben macht sie
#  dauerhaft. Intern bleibt der Kurzname ("short") der Schlüssel jeder
#  Person – angezeigt wird überall nur der Name.

TEAM_SECTIONS = ["👤 Personen", "🎯 Wer darf was", "📌 Feste Einsätze", "🔄 Rotationen",
                 "⚖️ Regeln", "🗂️ BackUp-Team"]
WOCHEN = {"jede Woche": None, "gerade KW": "gerade", "ungerade KW": "ungerade"}
WOCHEN_LABEL = {v: k for k, v in WOCHEN.items()}
FREI = "frei (nicht da)"
# Rollen-Matrix: Rolle in team.json → Spaltentitel (Schalter hat eine eigene Auswahl-Spalte)
ROLLEN_SPALTEN = {
    "HUB": "HUB", "ERF8": "ERF8", "ERF9": "ERF9 Mi–Fr", "TPA": "TAGES PA", "PHC": "18-Uhr-Liste",
    "PO": "Postöffnung", "Scanning": "Scanning", "RX Abo": "RX Abo",
    "Direkt": "Direkt", "ONB": "ONB", "BTM": "BTM", "VBZ": "VBZ",
}
SCHALTER_WAHL = ["–", "ja", "Reserve"]
HALBTAG_HILFE = "ganz = ganzer Tag · VM = nur Vormittag · NM = nur Nachmittag · frei = nicht da"

HILFE = [
    ("eine neue Person aufnehmen",
     "👤 Personen → «➕ Neue Person». Rollen lassen sich von einer ähnlichen Person übernehmen – danach in 🎯 Wer darf was prüfen."),
    ("jemanden vorübergehend nicht einplanen (krank, Mutterschutz)",
     "👤 Personen → Häkchen «Aktiv» entfernen. Rollen und feste Einsätze bleiben erhalten."),
    ("Pensum, Arbeitstage oder Funktion ändern", "👤 Personen → direkt in der Tabelle."),
    ("jemandem eine Aufgabe erlauben oder wegnehmen", "🎯 Wer darf was → Häkchen setzen oder entfernen."),
    ("einen festen Einsatz ändern (z. B. Mo ERF9 an eine andere Person)",
     "📌 Feste Einsätze → bei der alten Zeile «entfernen» anhaken, oben die neue Zeile hinzufügen."),
    ("«jede 2. Woche frei» eintragen", "📌 Feste Einsätze → Einsatz «frei», Wochen «gerade KW» oder «ungerade KW»."),
    ("korrigieren, wer diese Woche TV, 18-Uhr-Liste, Labor oder eine Wochenaufgabe hat",
     "🔄 Rotationen → bei «diese Woche» die richtige Person wählen; die Folgewochen ergeben sich daraus."),
    ("die Anzahl Personen für TEL, ABKL, PO, HUB, RX Abo … ändern", "⚖️ Regeln"),
    ("eine einmalige Abwesenheit oder einen Sonderwunsch erfassen", "Nicht hier – in Schritt 1 «Planung»."),
]

# Regeln, die im Programm fest eingebaut sind (Änderung = Code-Anpassung)
FESTE_REGELN = [
    "TEL und ABKL: pro Person höchstens ein Halbtag pro Tag (nur-Telefon-Personen ausgenommen).",
    "Tagesverantwortung: dieselbe Teamleitung ganztags, an dem Tag ohne andere Aufgaben (TL/Q).",
    "TAGES PA, Schalter und ERF9 immer ganztags; ERF9 Mi–Fr wechselnd (Mo/Di über feste Einsätze).",
    "ERF8: eine Person pro Nachmittag.",
    "Teamleitungen bekommen nie TAGES PA oder Postöffnung; freie Halbtage werden ERF7/Q.",
    "18-Uhr-Liste: die Person scannt an diesem Tag nicht; nie nur-Telefon- oder Aushilfe-Personen.",
    "Aushilfe: täglich PO/SCAN am Vormittag und scanning am Nachmittag.",
    "Wochenaufgaben Direkt / ONB / BTM: drei verschiedene Personen; wer die ganze Woche fehlt, wird übersprungen.",
    "PO/TEL am Nachmittag wird nie automatisch vergeben (nur Notfall, manuell im Plan).",
]


def _b(v) -> bool:
    return bool(v) and not (isinstance(v, float) and pd.isna(v))


def _s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def disp(short: str) -> str:
    """Anzeigename zum internen Kurznamen."""
    cfg = st.session_state.cfg
    for p in cfg["mitarbeiter"] + cfg["backup_team"]:
        if p["short"] == short:
            return p.get("name") or short
    return short


def _member(cfg: dict, short: str) -> dict:
    return next(m for m in cfg["mitarbeiter"] if m["short"] == short)


def _cfg_copy() -> dict:
    return copy.deepcopy(st.session_state.cfg)


def _editor_height(n_rows: int) -> int:
    """Höhe, bei der eine Tabelle ganz ohne eigene Scrollleiste sichtbar ist."""
    return 35 * (n_rows + 1) + 3


def try_commit(new_cfg: dict, msg: str = "", remount: bool = False) -> list[str]:
    """Übernimmt eine geänderte Konfiguration sofort in die Sitzung.
    Liefert Fehler – dann bleibt alles beim Alten."""
    new_cfg = ak.normalize_config(new_cfg)
    errors = ak.config_errors(new_cfg)
    if errors:
        return errors
    st.session_state.cfg = new_cfg
    st.session_state.result = None
    if remount:  # Tabellen frisch aufbauen: st.data_editor verwirft seinen Bearbeitungsstand,
        st.session_state.cfg_version += 1  # sobald sich die Daten ändern – sonst gehen Folge-Edits verloren
    if msg:
        st.toast(msg, icon="✅")
    return []


def commit_cb(new_cfg: dict, msg: str = "", remount: bool = False) -> None:
    """Für Widget-Callbacks: Fehler oben anzeigen und Widgets auf den gültigen Stand zurücksetzen."""
    errors = try_commit(new_cfg, msg, remount)
    if errors:
        _fail(errors)


def _fail(errors: list[str]) -> None:
    st.session_state.team_error = errors
    st.session_state.cfg_version += 1


# ── Speicherleiste ────────────────────────────────────────────────────

def _save_cb() -> None:
    st.session_state.save_msgs = save_everything(push_github=bool(github_conf()))


def _discard_cb() -> None:
    replace_cfg(json.loads(st.session_state.cfg_saved_json))
    st.toast("Änderungen verworfen – wieder der zuletzt gespeicherte Stand.", icon="↩️")


def render_save_bar() -> None:
    changes = ak.describe_changes(json.loads(st.session_state.cfg_saved_json), st.session_state.cfg)
    if entries_dirty:
        changes.append("Abwesenheiten / Sonderwünsche aus Schritt 1")
    # Gleiche Höhe in beiden Zuständen, damit die Seite beim ersten Ändern nicht springt
    with st.container(border=True, key="savebar_dirty" if changes else "savebar_ok"):
        c1, c2, c3, c4 = st.columns([5.2, 1.6, 1.2, 1.2], vertical_alignment="center")
        n = len(changes)
        c1.markdown(f"**{n} ungespeicherte Änderung{'en' if n != 1 else ''}** – gilt schon für den nächsten Plan, "
                    "geht aber verloren, wenn die App neu startet." if changes else
                    "✓ **Alles gespeichert** – der Generator arbeitet mit diesem Stand.")
        with c2.popover("Was wurde geändert?", width="stretch", disabled=not changes):
            st.markdown("\n".join(f"- {c}" for c in changes))
        c3.button("💾  Speichern", key="savebar_save", type="primary", width="stretch", on_click=_save_cb,
                  disabled=not changes)
        c4.button("↶  Verwerfen", key="savebar_discard", width="stretch", on_click=_discard_cb,
                  disabled=not cfg_dirty, help="Team & Regeln auf den zuletzt gespeicherten Stand zurücksetzen")
    for ok, msg in st.session_state.pop("save_msgs", []):
        (st.success if ok else st.warning)(msg)


# ── 👤 Personen ───────────────────────────────────────────────────────

def _on_people_edit(key: str, shorts: list[str]) -> None:
    cfg = _cfg_copy()
    moves, errors = [], []
    for ridx, changes in st.session_state[key].get("edited_rows", {}).items():
        short = shorts[int(ridx)]
        m = _member(cfg, short)
        for col, val in changes.items():
            if col == "Name":
                if _s(val):
                    m["name"] = _s(val)
                else:
                    errors.append("Name darf nicht leer sein")
            elif col == "%" and _s(val):
                m["pct"] = int(float(val))
            elif col == "Aktiv":
                m["aktiv"] = _b(val)
            elif col == "Funktion":
                ak.set_funktion(cfg, short, _s(val) or "Team")
            elif col in DAY_KEYS:
                anw = dict(m.get("anwesend", {}))
                anw[col] = _s(val) or "ganz"
                m["anwesend"] = {dk: anw[dk] for dk in DAY_KEYS if anw.get(dk, "ganz") != "ganz"}
            elif col == "Bemerkung":
                m["bemerkung"] = _s(val)
            elif col == "Nr." and _s(val):
                moves.append((short, int(float(val))))
    if errors:
        _fail(errors)
        return
    for short, pos in moves:
        lst = cfg["mitarbeiter"]
        m = _member(cfg, short)
        lst.remove(m)
        lst.insert(max(0, min(len(lst), pos - 1)), m)
    commit_cb(cfg, msg="Reihenfolge geändert" if moves else "", remount=True)


@st.dialog("Neue Person", width="large")
def dialog_add_person() -> None:
    cfg = st.session_state.cfg
    shorts = [m["short"] for m in cfg["mitarbeiter"]]
    name = st.text_input("Name, wie er im Plan erscheint", placeholder="z. B. Sandra K.")
    c1, c2 = st.columns(2)
    pct = c1.number_input("Pensum %", min_value=10, max_value=100, value=100, step=10)
    fkt = c2.selectbox("Funktion", ak.FUNKTIONEN,
                       help="Teamleitung: Tagesverantwortung, Queue, 18-Uhr-Liste Montag · nur Telefon: immer TEL · "
                            "Aushilfe: täglich PO/SCAN + scanning")
    st.markdown(f"**Anwesenheit** <span style='color:var(--text-mute); font-size:12px;'>{HALBTAG_HILFE}</span>",
                unsafe_allow_html=True)
    dcols = st.columns(5)
    anw = {dk: dcols[i].selectbox(dk, HALBTAG_WERTE, key=f"np_anw_{dk}") for i, dk in enumerate(DAY_KEYS)}
    c3, c4 = st.columns(2)
    regs = [m["short"] for m in cfg["mitarbeiter"] if ak.funktion(m) in ("Team", "Teamleitung")]
    vorlage = c3.selectbox("Rollen übernehmen von (optional)", ["–"] + regs,
                           format_func=lambda s: "– keine –" if s == "–" else disp(s),
                           help="Die neue Person darf dann dieselben Aufgaben – anpassen unter «Wer darf was».")
    nach = c4.selectbox("Position im Plan", ["–"] + shorts,
                        format_func=lambda s: "ganz unten" if s == "–" else f"nach {disp(s)}")
    if st.button("➕  Hinzufügen", type="primary"):
        name = name.strip()
        if not name:
            st.error("Bitte einen Namen eingeben.")
            return
        new = _cfg_copy()
        short = ak.new_short(name, new)
        m = {"short": short, "name": name, "pct": int(pct), "aktiv": True, "tl": False,
             "nur_tel": False, "aushilfe": False,
             "kann": list(_member(new, vorlage).get("kann", [])) if vorlage != "–" else [],
             "anwesend": {dk: v for dk, v in anw.items() if v != "ganz"},
             "frei_alle_2_wochen": None, "fix": [], "bemerkung": ""}
        lst = new["mitarbeiter"]
        lst.insert(len(lst) if nach == "–" else [x["short"] for x in lst].index(nach) + 1, m)
        ak.set_funktion(new, short, fkt)
        errors = try_commit(new, msg=f"{name} hinzugefügt – Rollen unter «Wer darf was» prüfen", remount=True)
        if errors:
            st.error(" · ".join(errors))
            return
        st.rerun()


@st.dialog("Person entfernen")
def dialog_remove_person() -> None:
    cfg = st.session_state.cfg
    shorts = [m["short"] for m in cfg["mitarbeiter"]]
    who = st.selectbox("Wen?", shorts, format_func=lambda s: disp(s) + ("" if _member(cfg, s).get("aktiv", True) else " (inaktiv)"))
    aktiv = _member(cfg, who).get("aktiv", True)
    st.warning(f"**{disp(who)}** wird mit allen Rollen und festen Einsätzen gelöscht.\n\n"
               "Nur vorübergehend weg (krank, Mutterschutz, lange Ferien)? Dann besser **deaktivieren** – "
               "so bleibt alles erhalten, und später genügt das Häkchen «Aktiv».")
    refs = ak.rotation_refs(cfg, who)
    if refs:
        st.info("Wird beim Entfernen auch hier herausgenommen: " + ", ".join(refs))
    c1, c2, c3 = st.columns(3)
    if c1.button("Deaktivieren", width="stretch", disabled=not aktiv):
        new = _cfg_copy()
        _member(new, who)["aktiv"] = False
        try_commit(new, msg=f"{disp(who)} deaktiviert", remount=True)
        st.rerun()
    if c2.button("Endgültig entfernen", type="primary", width="stretch"):
        name = disp(who)
        new = _cfg_copy()
        ak.remove_member(new, who)
        try_commit(new, msg=f"{name} entfernt", remount=True)
        st.rerun()
    if c3.button("Abbrechen", width="stretch"):
        st.rerun()


def section_people(ver: int) -> None:
    cfg = st.session_state.cfg
    members = cfg["mitarbeiter"]
    c1, c2, c3 = st.columns([1.2, 1.5, 6], vertical_alignment="center")
    if c1.button("➕  Neue Person", key=f"btn_add_person_{ver}", type="primary", width="stretch"):
        dialog_add_person()
    if c2.button("🗑️  Person entfernen …", key=f"btn_rm_person_{ver}", width="stretch"):
        dialog_remove_person()
    c3.caption("Direkt in der Tabelle ändern – jede Änderung gilt sofort. **Aktiv** ausschalten, wenn jemand "
               "vorübergehend nicht eingeplant werden soll; Rollen und feste Einsätze bleiben dann erhalten. "
               "**Nr.** ändern verschiebt die Person im Plan.")
    rows = [{"Nr.": i + 1, "Name": m.get("name") or m["short"], "%": int(m.get("pct", 100)),
             "Aktiv": bool(m.get("aktiv", True)), "Funktion": ak.funktion(m),
             **{dk: m.get("anwesend", {}).get(dk, "ganz") for dk in DAY_KEYS},
             "Bemerkung": m.get("bemerkung", "")} for i, m in enumerate(members)]
    df = pd.DataFrame(rows, columns=["Nr.", "Name", "%", "Aktiv", "Funktion"] + DAY_KEYS + ["Bemerkung"])
    col_config = {
        "Nr.": st.column_config.NumberColumn("Nr.", min_value=1, max_value=max(1, len(rows)), step=1, width=48,
                                             help="Reihenfolge im Plan – Zahl ändern, um die Person zu verschieben"),
        "Name": st.column_config.TextColumn("Name", width=140, pinned=True,
                                            help="So erscheint die Person im Plan, z. B. «Sandra K.»"),
        "%": st.column_config.NumberColumn("%", min_value=0, max_value=100, step=5, width=56, help="Pensum in Prozent"),
        "Aktiv": st.column_config.CheckboxColumn("Aktiv", width=60,
                                                 help="Aus = wird vorübergehend nicht eingeplant"),
        "Funktion": st.column_config.SelectboxColumn(
            "Funktion", options=ak.FUNKTIONEN, required=True, width=120,
            help="Teamleitung: Tagesverantwortung, Queue, 18-Uhr-Liste Montag · nur Telefon: immer TEL · "
                 "Aushilfe: täglich PO/SCAN + scanning"),
        "Bemerkung": st.column_config.TextColumn("Bemerkung", width="large"),
    }
    for dk in DAY_KEYS:
        col_config[dk] = st.column_config.SelectboxColumn(dk, options=HALBTAG_WERTE, required=True, width=64,
                                                          help=f"Anwesenheit {dk}: {HALBTAG_HILFE}")
    key = f"people_{ver}"
    st.data_editor(df, key=key, column_config=col_config, hide_index=True, num_rows="fixed", width="stretch",
                   height=_editor_height(len(df)), on_change=_on_people_edit, args=(key, [m["short"] for m in members]))


# ── 🎯 Wer darf was ───────────────────────────────────────────────────

def _on_roles_edit(key: str, shorts: list[str]) -> None:
    cfg = _cfg_copy()
    by_label = {lbl: r for r, lbl in ROLLEN_SPALTEN.items()}
    for ridx, changes in st.session_state[key].get("edited_rows", {}).items():
        m = _member(cfg, shorts[int(ridx)])
        kann = list(m.get("kann", []))
        for col, val in changes.items():
            if col == "Schalter":
                kann = [r for r in kann if r not in ("Schalter", "Schalter-Reserve")]
                if val == "ja":
                    kann.append("Schalter")
                elif val == "Reserve":
                    kann.append("Schalter-Reserve")
            elif col in by_label:
                role = by_label[col]
                kann = [r for r in kann if r != role] + ([role] if _b(val) else [])
        m["kann"] = kann
    commit_cb(cfg, remount=True)


def section_roles(ver: int) -> None:
    cfg = st.session_state.cfg
    regs = [m for m in cfg["mitarbeiter"] if not m.get("nur_tel") and not m.get("aushilfe")]
    st.caption("Häkchen = darf diese Aufgabe übernehmen – der Generator verteilt Aufgaben nur an Personen mit Häkchen. "
               "**Schalter**: «ja» = regulär (zwei Gruppen im Wochenwechsel), «Reserve» = nur, wenn aus der Gruppe "
               "niemand frei ist. Teamleitungen (TL) bekommen nie TAGES PA oder Postöffnung. "
               "Personen mit «nur Telefon» oder «Aushilfe» fehlen hier – sie haben feste Aufgaben.")
    rows = []
    for m in regs:
        kann = set(m.get("kann", []))
        label = (m.get("name") or m["short"]) + (" · TL" if m.get("tl") else "") + \
                ("" if m.get("aktiv", True) else " (inaktiv)")
        row = {"Name": label,
               "Schalter": "ja" if "Schalter" in kann else "Reserve" if "Schalter-Reserve" in kann else "–"}
        row.update({lbl: r in kann for r, lbl in ROLLEN_SPALTEN.items()})
        rows.append(row)
    df = pd.DataFrame(rows, columns=["Name", "Schalter"] + list(ROLLEN_SPALTEN.values()))
    col_config = {
        "Name": st.column_config.TextColumn("Name", disabled=True, pinned=True, width=170),
        "Schalter": st.column_config.SelectboxColumn("Schalter", options=SCHALTER_WAHL, required=True, width=90,
                                                     help="ja = regulär · Reserve = nur wenn aus der Gruppe niemand frei ist"),
    }
    for r, lbl in ROLLEN_SPALTEN.items():
        col_config[lbl] = st.column_config.CheckboxColumn(lbl, help=ROLLEN[r], width=max(62, int(7.5 * len(lbl)) + 26))
    key = f"roles_{ver}"
    st.data_editor(df, key=key, column_config=col_config, hide_index=True, num_rows="fixed", width="stretch",
                   height=_editor_height(len(df)), on_change=_on_roles_edit, args=(key, [m["short"] for m in regs]))
    act = [m for m in regs if m.get("aktiv", True)]
    counts = [("Schalter", sum("Schalter" in m.get("kann", []) for m in act))] + \
             [(lbl, sum(r in m.get("kann", []) for m in act)) for r, lbl in ROLLEN_SPALTEN.items() if r != "VBZ"]
    st.caption("Aktive Personen pro Aufgabe: " +
               " · ".join(f"{lbl} **{n}**" + (" ⚠️" if n <= 1 else "") for lbl, n in counts))


# ── 📌 Feste Einsätze ────────────────────────────────────────────────

def _fix_options(cfg: dict) -> list[str]:
    opts = [FREI] + [t for t in TASK_OPTIONS if t and t not in ("KRANK", "FERIEN")]
    for r in ak.recurring_rows(cfg):
        if r["task"] != "frei" and r["task"] not in opts:
            opts.append(r["task"])
    return opts


def _commit_recurring(rows: list[dict], msg: str = "", remount: bool = False) -> list[str]:
    cfg = _cfg_copy()
    errors = ak.apply_recurring_rows(cfg, rows)
    return errors or try_commit(cfg, msg=msg, remount=remount)


def _on_fixed_edit(key: str, base: list[dict]) -> None:
    rows = [dict(r) for r in base]
    removed = set()
    for ridx, changes in st.session_state[key].get("edited_rows", {}).items():
        i = int(ridx)
        r = rows[i]
        for col, val in changes.items():
            if col == "entfernen" and _b(val):
                removed.add(i)
            elif col == "Tag":
                r["tag"] = _s(val) or r["tag"]
            elif col == "Halbtag":
                r["halbtag"] = _s(val) or r["halbtag"]
            elif col == "Einsatz" and _s(val):
                r["task"] = "frei" if _s(val) == FREI else _s(val)
            elif col == "Wochen":
                r["kw"] = WOCHEN.get(_s(val))
    errors = _commit_recurring([r for i, r in enumerate(rows) if i not in removed],
                               msg="Fester Einsatz entfernt" if removed else "", remount=True)
    if errors:
        _fail(errors)


def section_fixed(ver: int) -> None:
    cfg = st.session_state.cfg
    shorts = [m["short"] for m in cfg["mitarbeiter"]]
    st.caption("Was jede Woche – oder jede zweite Woche – fix ist, z. B. «Mo ganz ERF9», «Di VM KGS» oder "
               "«alle 2 Wochen Fr frei». Der Generator plant diese Einsätze vor allem anderen. "
               f"Die angezeigte **KW {info['kw']}** ist eine **{'gerade' if info['is_even'] else 'ungerade'}** Woche.")
    opts = _fix_options(cfg)
    with st.form(f"add_fix_{ver}", border=True):
        st.markdown("**Neuer fester Einsatz**")
        c = st.columns([2, 0.9, 1, 1.6, 1.3, 1.3], vertical_alignment="bottom")
        who = c[0].selectbox("Person", shorts, format_func=disp)
        tag = c[1].selectbox("Tag", DAY_KEYS)
        halb = c[2].selectbox("Halbtag", ["ganz", "VM", "NM"])
        task = c[3].selectbox("Einsatz", opts, index=opts.index("HO") if "HO" in opts else 0,
                              help="«frei» = an diesem Tag nicht da (nur mit «gerade KW» / «ungerade KW»)")
        wochen = c[4].selectbox("Wochen", list(WOCHEN))
        if c[5].form_submit_button("➕  Hinzufügen", type="primary", width="stretch"):
            new_row = {"short": who, "tag": tag, "halbtag": halb, "task": "frei" if task == FREI else task,
                       "kw": WOCHEN[wochen]}
            errors = _commit_recurring(ak.recurring_rows(cfg) + [new_row],
                                       msg=f"{disp(who)}: {tag} {halb} {task} hinzugefügt", remount=True)
            if errors:
                st.error(" · ".join(errors))
            else:
                st.rerun()

    base = ak.recurring_rows(cfg)
    if not base:
        st.info("Noch keine festen Einsätze.")
        return
    inaktiv = {m["short"] for m in cfg["mitarbeiter"] if not m.get("aktiv", True)}
    df = pd.DataFrame([{"Person": disp(r["short"]) + (" (inaktiv)" if r["short"] in inaktiv else ""),
                        "Tag": r["tag"], "Halbtag": r["halbtag"],
                        "Einsatz": FREI if r["task"] == "frei" else r["task"],
                        "Wochen": WOCHEN_LABEL.get(r["kw"], "jede Woche"), "entfernen": False} for r in base],
                      columns=["Person", "Tag", "Halbtag", "Einsatz", "Wochen", "entfernen"])
    key = f"fixed_{ver}"
    st.data_editor(
        df, key=key, hide_index=True, num_rows="fixed", width="stretch", height=_editor_height(len(df)),
        disabled=["Person"], on_change=_on_fixed_edit, args=(key, base),
        column_config={
            "Person": st.column_config.TextColumn("Person", width="medium"),
            "Tag": st.column_config.SelectboxColumn("Tag", options=DAY_KEYS, required=True, width="small"),
            "Halbtag": st.column_config.SelectboxColumn("Halbtag", options=["ganz", "VM", "NM"], required=True, width="small"),
            "Einsatz": st.column_config.SelectboxColumn("Einsatz", options=opts, required=True, width="medium"),
            "Wochen": st.column_config.SelectboxColumn("Wochen", options=list(WOCHEN), required=True, width="small"),
            "entfernen": st.column_config.CheckboxColumn("entfernen", width="small", help="Anhaken = Zeile löschen"),
        })


# ── 🔄 Rotationen ────────────────────────────────────────────────────

def _on_rot_list(key: str, field: str, hidden: list[str]) -> None:
    cfg = _cfg_copy()
    cfg["rotation"][field] = list(st.session_state[key]) + [n for n in hidden if n not in st.session_state[key]]
    commit_cb(cfg)


def _on_rot_now(key: str, off_key: str, pool: list[str], widx: int, mult: int) -> None:
    cfg = _cfg_copy()
    cfg["rotation"]["offset"][off_key] = (pool.index(st.session_state[key]) - widx * mult) % len(pool)
    commit_cb(cfg)


def _on_rot_value(key: str, field: str) -> None:
    cfg = _cfg_copy()
    v = st.session_state[key]
    cfg["rotation"][field] = "" if v == "–" else v
    commit_cb(cfg)


def _on_phc_fix(key: str, dk: str) -> None:
    cfg = _cfg_copy()
    fix = dict(cfg["rotation"]["phc_fix"])
    if st.session_state[key] == "–":
        fix.pop(dk, None)
    else:
        fix[dk] = st.session_state[key]
    cfg["rotation"]["phc_fix"] = {d: fix[d] for d in DAY_KEYS if d in fix}
    commit_cb(cfg)


def _on_phc_excl(key: str, shorts: list[str]) -> None:
    cfg = _cfg_copy()
    ex = {dk: list(v) for dk, v in cfg["rotation"]["phc_tag_ausschluss"].items()}
    for ridx, changes in st.session_state[key].get("edited_rows", {}).items():
        s = shorts[int(ridx)]
        for dk, val in changes.items():
            lst = [n for n in ex.get(dk, []) if n != s]
            ex[dk] = lst + [s] if _b(val) else lst
    cfg["rotation"]["phc_tag_ausschluss"] = {dk: ex[dk] for dk in DAY_KEYS if ex.get(dk)}
    commit_cb(cfg, remount=True)


def _rotation_block(ver: int, title: str, caption: str, field: str, off_key: str, options: list[str],
                    now_label: str, mult: int = 1) -> list[str]:
    """Liste «wer, in welcher Reihenfolge» + «wer ist in der angezeigten Woche dran»."""
    rot = st.session_state.cfg["rotation"]
    widx = info["widx"]
    st.markdown(f"**{title}**")
    st.caption(caption)
    current = [n for n in rot[field] if n in options]
    hidden = [n for n in rot[field] if n not in options]
    key = f"rot_{field}_{ver}"
    pool = st.multiselect(title, options, default=current, format_func=disp, key=key, label_visibility="collapsed",
                          placeholder="Personen wählen – Reihenfolge = Reihenfolge der Auswahl",
                          on_change=_on_rot_list, args=(key, field, hidden))
    if hidden:
        st.caption("Inaktiv, wird übersprungen: " + ", ".join(disp(n) for n in hidden))
    if pool:
        idx = (widx * mult + rot["offset"].get(off_key, 0)) % len(pool)
        nkey = f"rot_{off_key}_now_{ver}_{'|'.join(pool)}"
        st.selectbox(now_label, pool, index=idx, format_func=disp, key=nkey,
                     on_change=_on_rot_now, args=(nkey, off_key, list(pool), widx, mult))
    return pool


def _rotation_preview(cfg: dict, n_weeks: int = 6) -> pd.DataFrame:
    """Wer ist in den nächsten Wochen laut Rotation dran (ohne Abwesenheiten)."""
    rot, off = cfg["rotation"], cfg["rotation"]["offset"]
    act = [m for m in cfg["mitarbeiter"] if m.get("aktiv", True)]
    act_set = {m["short"] for m in act}
    reg = [m for m in act if not m.get("nur_tel") and not m.get("aushilfe")]
    tv = [n for n in rot["tv_reihenfolge"] if n in act_set] or [m["short"] for m in act if m.get("tl")]
    lists = {k: [n for n in rot[k] if n in act_set] for k in ("phc_montag_tl", "phc_freitag_wechsel", "labor_mi_nm")}
    wa_pools = {lbl: [m["short"] for m in reg if lbl in m.get("kann", [])] for lbl in ("Direkt", "ONB", "BTM")}

    def at(lst, w, off_key, mult=1):
        return disp(lst[(w * mult + off.get(off_key, 0)) % len(lst)]) if lst else "–"

    rows = []
    for i in range(n_weeks):
        w = info["widx"] + i
        monday = info["monday"] + timedelta(weeks=i)
        used, wa = set(), {}
        for lbl, pool in wa_pools.items():
            if not pool:
                wa[lbl] = "–"
                continue
            idx = (w + off.get(lbl, 0)) % len(pool)
            cand = next((pool[(idx + j) % len(pool)] for j in range(len(pool))
                         if pool[(idx + j) % len(pool)] not in used), pool[idx])
            used.add(cand)
            wa[lbl] = disp(cand)
        rows.append({
            "KW": f"KW {monday.isocalendar()[1]}", "ab": monday.strftime("%d.%m."),
            "TV Montag": at(tv, w, "tv", 5), "18-Uhr Mo": at(lists["phc_montag_tl"], w, "phc_montag"),
            "18-Uhr Fr": at(lists["phc_freitag_wechsel"], w, "phc_freitag"),
            "Labor Mi": at(lists["labor_mi_nm"], w, "labor"),
            "Schalter": "Gruppe A" if (w + off.get("schalter", 0)) % 2 == 0 else "Gruppe B", **wa,
        })
    return pd.DataFrame(rows)


def section_rotation(ver: int) -> None:
    cfg = st.session_state.cfg
    rot = cfg["rotation"]
    widx = info["widx"]
    kw = info["kw"]
    act = [m for m in cfg["mitarbeiter"] if m.get("aktiv", True)]
    reg = [m for m in act if not m.get("nur_tel") and not m.get("aushilfe")]
    tls = [m["short"] for m in act if m.get("tl")]
    phc_ok = [m["short"] for m in reg if "PHC" in m.get("kann", [])]
    st.caption(f"Rotationen ergeben sich aus der Kalenderwoche. Mit «diese Woche» stellst du ein, wer in der "
               f"angezeigten **KW {kw}** dran ist – die Folgewochen laufen von dort aus weiter (Vorschau unten). "
               f"Die Reihenfolge entspricht der Reihenfolge der Auswahl: zum Umsortieren Person entfernen (×) und neu wählen.")

    left, right = st.columns(2, gap="large")
    with left:
        with st.container(border=True):
            tv = _rotation_block(ver, "Tagesverantwortung (TV)",
                                 "Die Teamleitungen wechseln sich Tag für Tag ab. Leer = alle Teamleitungen.",
                                 "tv_reihenfolge", "tv", tls, f"TV am Montag, KW {kw}", mult=5)
            tv_list = tv or tls
            if tv_list:
                start = (widx * 5 + rot["offset"]["tv"]) % len(tv_list)
                st.caption("Diese Woche: " + " · ".join(f"{DAY_KEYS[d]} {disp(tv_list[(start + d) % len(tv_list)])}"
                                                         for d in range(5)))
        with st.container(border=True):
            _rotation_block(ver, "Labor Mi Nachmittag (ERF5)", "Wöchentlicher Wechsel.",
                            "labor_mi_nm", "labor", [m["short"] for m in reg], f"Labor in KW {kw}")
        with st.container(border=True):
            st.markdown("**Schalter-Gruppe**")
            sch_pool = [m["short"] for m in reg if "Schalter" in m.get("kann", [])]
            mid = (len(sch_pool) + 1) // 2
            groups = [sch_pool[:mid], sch_pool[mid:]]
            st.caption("Alle mit Schalter «ja» (unter «Wer darf was») werden in Plan-Reihenfolge in zwei Gruppen "
                       "geteilt, die sich wöchentlich abwechseln.")
            key = f"rot_schalter_{ver}_{'|'.join(sch_pool)}"
            st.selectbox(f"Aktive Gruppe in KW {kw}", [0, 1], index=(widx + rot["offset"]["schalter"]) % 2, key=key,
                         format_func=lambda g: f"Gruppe {'AB'[g]}: {', '.join(disp(n) for n in groups[g]) or '–'}",
                         on_change=_on_rot_now, args=(key, "schalter", [0, 1], widx, 1))
        with st.container(border=True):
            st.markdown("**Wochenaufgaben**")
            st.caption("Je eine Person pro Woche, nie dieselbe doppelt. Wer infrage kommt, steht unter «Wer darf was».")
            wcols = st.columns(3)
            for col, label in zip(wcols, ("Direkt", "ONB", "BTM")):
                pool = [m["short"] for m in reg if label in m.get("kann", [])]
                if not pool:
                    col.caption(f"{label}: niemand berechtigt")
                    continue
                key = f"rot_wa_{label}_{ver}_{'|'.join(pool)}"
                col.selectbox(f"{label} in KW {kw}", pool, index=(widx + rot["offset"].get(label, 0)) % len(pool),
                              format_func=disp, key=key, on_change=_on_rot_now, args=(key, label, pool, widx, 1))

    with right:
        with st.container(border=True):
            st.markdown("**18-Uhr-Liste (PHC)**")
            _rotation_block(ver, "Montag – Teamleitungen im Wechsel", "Wöchentlicher Wechsel.",
                            "phc_montag_tl", "phc_montag", tls, f"Montag in KW {kw}")
            _rotation_block(ver, "Freitag – im Wechsel", "Wöchentlicher Wechsel. Leer = aus dem Pool.",
                            "phc_freitag_wechsel", "phc_freitag", phc_ok, f"Freitag in KW {kw}")
            st.markdown("**Fix an einem Wochentag**")
            st.caption("Diese Person macht die 18-Uhr-Liste an diesem Tag immer (wenn sie da ist). "
                       "Montag und Freitag laufen über die Wechsel oben, sofern dort jemand eingetragen ist. "
                       "– = jemand aus dem Pool.")
            fcols = st.columns(5)
            for i, dk in enumerate(DAY_KEYS):
                cur = rot["phc_fix"].get(dk, "–")
                opts = ["–"] + phc_ok + ([cur] if cur not in phc_ok and cur != "–" else [])
                key = f"rot_phcfix_{dk}_{ver}"
                fcols[i].selectbox(dk, opts, index=opts.index(cur), key=key, format_func=lambda s: s if s == "–" else disp(s),
                                   on_change=_on_phc_fix, args=(key, dk))
            st.markdown("**Nie an diesem Tag**")
            st.caption("Häkchen = diese Person kommt an diesem Wochentag nicht auf die 18-Uhr-Liste.")
            ex = rot["phc_tag_ausschluss"]
            df = pd.DataFrame([{"Name": disp(s), **{dk: s in ex.get(dk, []) for dk in DAY_KEYS}} for s in phc_ok],
                              columns=["Name"] + DAY_KEYS)
            key = f"rot_phcex_{ver}"
            st.data_editor(df, key=key, hide_index=True, num_rows="fixed", width="stretch",
                           height=_editor_height(len(df)), disabled=["Name"], on_change=_on_phc_excl, args=(key, phc_ok),
                           column_config={"Name": st.column_config.TextColumn("Name", width="medium"),
                                          **{dk: st.column_config.CheckboxColumn(dk, width="small") for dk in DAY_KEYS}})

    st.markdown(f"**Vorschau – wer ist laut Rotation dran** "
                f"<span style='color:var(--text-mute); font-size:12px;'>ohne Abwesenheiten; ab KW {kw}</span>",
                unsafe_allow_html=True)
    st.dataframe(_rotation_preview(cfg), hide_index=True, width="stretch")


# ── ⚖️ Regeln ─────────────────────────────────────────────────────────

def _on_rule(key: str, field: str, kind: type) -> None:
    cfg = _cfg_copy()
    cfg["regeln"][field] = kind(st.session_state[key])
    commit_cb(cfg)


def _rule_number(container, label: str, field: str, lo: int, hi: int, ver: int, help: str | None = None) -> int:
    key = f"rule_{field}_{ver}"
    return container.number_input(label, lo, hi, int(st.session_state.cfg["regeln"][field]), step=1, key=key, help=help,
                                  on_change=_on_rule, args=(key, field, int))


def _rule_select(container, label: str, field: str, options: list[str], ver: int, help: str | None = None) -> None:
    cur = st.session_state.cfg["regeln"].get(field, options[0])
    key = f"rule_{field}_{ver}"
    container.selectbox(label, options, index=options.index(cur) if cur in options else 0, key=key, help=help,
                        on_change=_on_rule, args=(key, field, str))


def section_rules(ver: int) -> None:
    cfg = st.session_state.cfg
    rules = cfg["regeln"]
    st.caption("Wie viele Personen braucht es wofür? Jede Änderung gilt sofort für den nächsten Plan. "
               "Wer welche Aufgabe übernehmen darf, steht unter «Wer darf was».")
    left, right = st.columns(2, gap="large")
    with left:
        with st.container(border=True):
            st.markdown("**📞 Telefon (TEL)**")
            c1, c2 = st.columns(2)
            _rule_number(c1, "Montag Vormittag – Personen", "tel_montag_vm", 0, 10, ver)
            _rule_number(c2, "übrige Halbtage – Personen", "tel_sonst", 0, 10, ver)
            c3, c4 = st.columns(2)
            tl_min = _rule_number(c3, "Teamleitung am Telefon – mind. Halbtage/Woche", "tl_tel_min_pro_woche", 0, 10, ver,
                                  help="Wird für jede TL vorab reserviert. 0 = kein Minimum.")
            tl_max = _rule_number(c4, "… höchstens Halbtage/Woche", "tl_tel_max_pro_woche", 0, 10, ver,
                                  help="Bis zu dieser Zahl bekommen TL beim Telefon den Vorzug.")
            if tl_min > tl_max:
                st.warning("Minimum ist grösser als Maximum – das Minimum gewinnt.")
            _rule_select(st, "TAGES PA am Nachmittag zusätzlich am Telefon (TAGES PA/TEL)", "tpa_tel_nm", TPA_TEL_MODI, ver,
                         help="bei Bedarf: nur wenn die TEL-Besetzung sonst nicht reicht · immer · nie")
        with st.container(border=True):
            st.markdown("**🔎 Abklärung (ABKL)**")
            _rule_number(st, "Personen pro Halbtag", "abkl_pro_halbtag", 0, 10, ver)
        with st.container(border=True):
            st.markdown("**📬 Postöffnung (PO)**")
            _rule_number(st, "Zusätzliche PO-Personen am Vormittag (zur Aushilfe mit PO/SCAN)", "po_zusatz_vm", 0, 5, ver)
            po_opts = ["–"] + [m["short"] for m in cfg["mitarbeiter"]
                               if m.get("aktiv", True) and "PO" in m.get("kann", []) and not m.get("tl")]
            cur = cfg["rotation"].get("po_dienstag_bevorzugt") or "–"
            if cur not in po_opts:
                po_opts.append(cur)
            key = f"rot_popref_{ver}"
            st.selectbox("Dienstag bevorzugt", po_opts, index=po_opts.index(cur), key=key,
                         format_func=lambda s: "– niemand –" if s == "–" else disp(s),
                         on_change=_on_rot_value, args=(key, "po_dienstag_bevorzugt"))
            key = f"rule_scan_nm_{ver}"
            st.checkbox("Fehlt die Aushilfe: Nachmittags-Scan durch jemand anderen ersetzen",
                        bool(rules["scan_nm_ersatz_wenn_aushilfe_fehlt"]), key=key,
                        on_change=_on_rule, args=(key, "scan_nm_ersatz_wenn_aushilfe_fehlt", bool))
    with right:
        with st.container(border=True):
            st.markdown("**🏬 Schalter**")
            st.caption("Eine Person pro Tag, ganztags. Wer Schalter oder Reserve macht: «Wer darf was».")
            c1, c2 = st.columns(2)
            _rule_select(c1, "Reserve einsetzen …", "schalter_reserve_modus", SCHALTER_RESERVE_MODI, ver,
                         help="vor / nach der anderen Gruppe, wenn aus der aktiven Gruppe niemand frei ist · nie")
            _rule_select(c2, "Code im Plan", "schalter_code", SCHALTER_CODES, ver,
                         help="April-Liste: ERF4/SCH · seit Mai im Plan: ERF7/SCH")
        with st.container(border=True):
            st.markdown("**🧾 HUB und RX Abo**")
            c1, c2 = st.columns(2)
            _rule_number(c1, "HUB-Halbtage pro Tag", "hub_pro_tag", 0, 2, ver,
                         help="2 = eine Person am Vormittag + eine andere am Nachmittag · 1 = nur Vormittag · 0 = aus")
            _rule_number(c2, "RX Abo – Halbtage pro Woche", "rx_abo_pro_woche", 0, 10, ver,
                         help="höchstens einer pro Tag · 0 = aus")
        with st.container(border=True):
            st.markdown("**🗂️ Scanning (Stammteam)**")
            _rule_number(st, "höchstens Halbtage pro Person und Woche", "scan_max_halbtage_pro_woche", 0, 10, ver)
            key = f"rule_scan_anteil_{ver}"
            p = st.slider("Wie oft wird ein freier Halbtag zu «scanning»?", 0.0, 1.0, float(rules["scan_anteil"]), 0.05,
                          key=key, on_change=_on_rule, args=(key, "scan_anteil", float))
            st.caption("≈ nie" if p == 0 else f"≈ jeder {max(1, round(1 / p))}. freie Halbtag bei Personen mit Scanning")
    with st.expander("🔒 Fest im Programm eingebaut (Änderung braucht eine Anpassung im Code)"):
        st.markdown("\n".join(f"- {r}" for r in FESTE_REGELN))
        st.markdown("**Nur manuell** (nie automatisch, aber im Plan-Editor wählbar): " + ", ".join(NUR_MANUELL) +
                    ". Vorbezüge (VBZ) sind als Rolle hinterlegt, werden aber nicht geplant.")


# ── 🗂️ BackUp-Team ────────────────────────────────────────────────────

def _on_backup_edit(key: str, base: list[dict]) -> None:
    rows = [dict(b) for b in base]
    removed, errors = set(), []
    for ridx, changes in st.session_state[key].get("edited_rows", {}).items():
        i = int(ridx)
        for col, val in changes.items():
            if col == "entfernen" and _b(val):
                removed.add(i)
            elif col == "Name":
                if _s(val):
                    rows[i]["name"] = _s(val)
                else:
                    errors.append("Name darf nicht leer sein")
            elif col == "Standard":
                rows[i]["standard"] = _s(val)
    if errors:
        _fail(errors)
        return
    cfg = _cfg_copy()
    cfg["backup_team"] = [b for i, b in enumerate(rows) if i not in removed]
    commit_cb(cfg, msg="BackUp-Zeile entfernt" if removed else "", remount=True)


def section_backup(ver: int) -> None:
    cfg = st.session_state.cfg
    st.caption("Zeilen unterhalb des Teams im Plan – werden nie automatisch eingeplant. «Standard» steht in jeder "
               "Zelle (z. B. ERP beim SPP-Team); einzelne Einträge über Schritt 1 mit Grund «Custom».")
    with st.form(f"add_backup_{ver}", border=True):
        c = st.columns([3, 2, 1.3], vertical_alignment="bottom")
        name = c[0].text_input("Name", placeholder="z. B. Petra M.")
        standard = c[1].text_input("Standard-Eintrag (optional)", placeholder="z. B. ERP")
        if c[2].form_submit_button("➕  Hinzufügen", type="primary", width="stretch"):
            if not name.strip():
                st.error("Bitte einen Namen eingeben.")
            else:
                new = _cfg_copy()
                new["backup_team"].append({"short": ak.new_short(name, new), "name": name.strip(),
                                           "standard": standard.strip()})
                errors = try_commit(new, msg=f"{name.strip()} hinzugefügt", remount=True)
                if errors:
                    st.error(" · ".join(errors))
                else:
                    st.rerun()
    base = [dict(b) for b in cfg["backup_team"]]
    df = pd.DataFrame([{"Name": b.get("name") or b["short"], "Standard": b.get("standard", ""), "entfernen": False}
                       for b in base], columns=["Name", "Standard", "entfernen"])
    key = f"backup_{ver}"
    st.data_editor(df, key=key, hide_index=True, num_rows="fixed", width="stretch", height=_editor_height(len(df)),
                   on_change=_on_backup_edit, args=(key, base),
                   column_config={"Name": st.column_config.TextColumn("Name", width="medium"),
                                  "Standard": st.column_config.TextColumn("Standard-Eintrag", width="medium"),
                                  "entfernen": st.column_config.CheckboxColumn("entfernen", width="small")})


# ── Seite ─────────────────────────────────────────────────────────────

def _keep_section() -> None:
    """Segment-Auswahl lässt sich abwählen – dann beim letzten Bereich bleiben."""
    if not st.session_state.get("team_section"):
        st.session_state.team_section = st.session_state.get("team_section_last", TEAM_SECTIONS[0])
    st.session_state.team_section_last = st.session_state.team_section


def render_team_page() -> None:
    ver = st.session_state.cfg_version
    n_total = len(st.session_state.cfg["mitarbeiter"])
    st.markdown(f'<div class="page-head"><h2>Team & Regeln</h2><div class="sub">{n_active} von {n_total} Personen aktiv · '
                'alles, was der Generator über das Team weiss. Änderungen gelten sofort für den nächsten Plan.</div></div>',
                unsafe_allow_html=True)
    render_save_bar()
    errors = st.session_state.pop("team_error", None)
    if errors:
        st.error("Nicht übernommen: " + " · ".join(errors))
    hints = ak.config_hints(st.session_state.cfg)
    if hints:
        st.warning("Hinweis: " + " · ".join(hints))

    if st.session_state.get("team_section") not in TEAM_SECTIONS:
        st.session_state.team_section = TEAM_SECTIONS[0]
    section = st.segmented_control("Bereich", TEAM_SECTIONS, key="team_section", on_change=_keep_section,
                                   label_visibility="collapsed") or TEAM_SECTIONS[0]
    with st.expander("❓  Wie ändere ich …?"):
        st.markdown("| Ich möchte … | So geht's |\n|---|---|\n" +
                    "\n".join(f"| {was} | {wie} |" for was, wie in HILFE))
    {
        "👤 Personen": section_people, "🎯 Wer darf was": section_roles, "📌 Feste Einsätze": section_fixed,
        "🔄 Rotationen": section_rotation, "⚖️ Regeln": section_rules, "🗂️ BackUp-Team": section_backup,
    }[section](ver)


# ══════════════════════════════════════════════════════════════════════
#  Schritt 1 – Planung
# ══════════════════════════════════════════════════════════════════════

def _goto_team() -> None:
    st.session_state.step = 3


def render_planung() -> None:
    n_total = len(st.session_state.cfg["mitarbeiter"])
    t1, t2 = st.columns([6, 1.4], vertical_alignment="center")
    t1.markdown(f'<div class="team-hint">👥 <b>{n_active}</b> von {n_total} Personen aktiv'
                f'{" · <b>ungespeicherte Team-Änderungen</b>" if cfg_dirty else ""} · '
                'Personen, Rollen, feste Einsätze, Rotationen und Regeln ändern unter «Team & Regeln».</div>',
                unsafe_allow_html=True)
    t2.button("⚙  Team & Regeln", key="goto_team", width="stretch", on_click=_goto_team)
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
                        f'<span class="entry-name">{html_escape(disp(e["name"]))}</span>'
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
        render_banner("warn", "Konfiguration prüfen (Team & Regeln)", res["config_problems"])
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


if st.session_state.step == 3:
    render_team_page()
elif st.session_state.step == 1 or st.session_state.result is None:
    render_planung()
else:
    render_ergebnis()
