#!/usr/bin/env python3
"""
KSC Arbeitsplan – Web-Backend (FastAPI)
========================================
Stellt das bestehende arbeitskalender.py als moderne Web-App bereit.
Läuft in Docker, plattformunabhängig, Zugriff über Browser.
"""

from __future__ import annotations

import os
import random
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field

# Kern-Logik aus dem bestehenden Modul (unverändert)
from arbeitskalender import (
    DAYS,
    SLOTS,
    build_schedule,
    create_employees,
    get_next_monday,
    validate_schedule,
    write_excel,
)

# ──────────────────────────────────────────────────────────────
# Pfade
# ──────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
OUTPUT_DIR = Path(os.environ.get("KSC_OUTPUT_DIR", BASE_DIR.parent / "output"))
STATE_FILE = Path(os.environ.get("KSC_STATE_FILE", OUTPUT_DIR / "scheduler_state.json"))

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ──────────────────────────────────────────────────────────────
# App
# ──────────────────────────────────────────────────────────────

app = FastAPI(
    title="KSC Arbeitsplan",
    description="Wöchentlicher Arbeitsplan-Generator für das KSC-Team",
    version="6.0.0",
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


# ──────────────────────────────────────────────────────────────
# Datenmodelle
# ──────────────────────────────────────────────────────────────

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


class AbsenceEntry(BaseModel):
    """Eine einzelne Abwesenheit / ein Sonderwunsch."""
    name: str = Field(..., description="Mitarbeitername (Kurzform)")
    type: str = Field(..., description="Krank | Ferien | Termin | Home Office | Custom")
    days: list[int] = Field(..., description="0=Mo … 4=Fr")
    slots: list[int] = Field(..., description="0=VM, 1=NM")
    note: str = Field("", description="Optionale Bemerkung")
    task: str = Field("", description="Bei 'Custom' der gewünschte Task-Code")


class GenerateRequest(BaseModel):
    week_offset: int = Field(0, description="0 = nächste Woche, 1 = übernächste, …")
    entries: list[AbsenceEntry] = Field(default_factory=list)


# ──────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────

def get_week_info(offset: int = 0) -> dict[str, Any]:
    """Liefert Meta-Informationen zur Zielwoche."""
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


def check_rule_impact(absences: list[tuple[str, int, int]], employees: dict) -> list[str]:
    """Warnungen bei Regelkonflikten – analog zur bestehenden GUI-Logik."""
    warnings: list[str] = []

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

    # Mindestbesetzung
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
                    f"KRITISCH {DAYS[day]} {SLOTS[slot]}: nur {available_count} verfügbar – mind. {min_needed} benötigt"
                )
            elif available_count < min_needed + 3:
                warnings.append(
                    f"Knapp: {DAYS[day]} {SLOTS[slot]} – nur {available_count} Personen"
                )

    return warnings


# ──────────────────────────────────────────────────────────────
# Routen
# ──────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    info = get_week_info(0)
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "employees": EMPLOYEE_LIST,
            "absence_types": list(ABSENCE_TYPES.keys()),
            "week_info": info,
        },
    )


@app.get("/api/week-info")
async def api_week_info(offset: int = 0) -> dict[str, Any]:
    info = get_week_info(offset)
    return {
        "kw": info["kw"],
        "is_even": info["is_even"],
        "monday_str": info["monday_str"],
        "friday_str": info["friday_str"],
        "two_week_notes": info["two_week_notes"],
    }


@app.post("/api/generate")
async def api_generate(req: GenerateRequest) -> JSONResponse:
    """Generiert den Arbeitsplan und liefert Ergebnis-Daten + Excel-Pfad."""
    info = get_week_info(req.week_offset)
    monday = info["monday"]
    kw = info["kw"]

    employees = create_employees(kw)
    overrides: dict[tuple[str, int, int], str] = {}
    extra_notes: dict[tuple[str, int, int], str] = {}
    absences: list[tuple[str, int, int]] = []

    for entry in req.entries:
        if entry.name not in employees:
            continue
        task_code = ABSENCE_TYPES.get(entry.type, "CUSTOM")
        note = entry.note.strip()

        if task_code == "TERMIN":
            label = f"*{note}" if note else "*Termin"
        elif task_code == "CUSTOM":
            label = entry.task.strip() or "HO"
        else:
            label = task_code

        for day in entry.days:
            for slot in entry.slots:
                if employees[entry.name].is_available(day, slot):
                    overrides[(entry.name, day, slot)] = label
                    if note and task_code in ("KRANK", "FERIEN", "HO"):
                        extra_notes[(entry.name, day, slot)] = note
                    if task_code in ("KRANK", "FERIEN") or label.startswith("*"):
                        absences.append((entry.name, day, slot))

    # Warnungen vor dem Planen
    warnings = check_rule_impact(absences, employees) if absences else []

    # Plan bauen
    random.seed(kw)
    sched = build_schedule(
        kw, monday, overrides, state_file=str(STATE_FILE)
    )
    for key, note in extra_notes.items():
        sched.notes[key] = note

    issues = validate_schedule(sched)

    # Excel schreiben
    filename = f"Arbeitsplan_KW{kw}_{monday.strftime('%Y%m%d')}.xlsx"
    output_path = OUTPUT_DIR / filename
    write_excel(sched, str(output_path))

    # Ergebnis aufbereiten
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

    # Plan-Matrix für die Vorschau-Tabelle
    order = [
        "Silvana", "Linda", "Lara", "Andrea G.", "Dipiga",
        "Isaura", "Martina", "Alessia", "Amra", "Nina",
        "Jesika", "Brigitte", "Florence", "Corinne", "Saskia",
        "Dragi", "Maria B.", "Andrea A.",
    ]
    rows = []
    for name in order:
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
            "short": name,
            "pct": emp.pct,
            "cells": cells,
        })

    return JSONResponse({
        "ok": True,
        "filename": filename,
        "download_url": f"/download/{filename}",
        "kw": kw,
        "monday_str": info["monday_str"],
        "friday_str": info["friday_str"],
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
        "days": DAYS,
        "slots": SLOTS,
    })


@app.get("/download/{filename}")
async def download(filename: str) -> FileResponse:
    # Sicherheit: nur Dateien aus OUTPUT_DIR
    safe_name = os.path.basename(filename)
    path = OUTPUT_DIR / safe_name
    if not path.exists():
        raise HTTPException(404, f"Datei {safe_name} nicht gefunden")
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=safe_name,
    )


@app.get("/api/health")
async def health() -> dict[str, str]:
    return {"status": "ok", "version": "6.0.0"}


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("KSC_PORT", 8000))
    host = os.environ.get("KSC_HOST", "0.0.0.0")
    uvicorn.run("server:app", host=host, port=port, reload=False)
