#!/usr/bin/env python3
"""
Arbeitskalender-Generator für KSC-Team
Erstellt wöchentliche Arbeitspläne als Excel-Datei mit allen Regeln.

OFFENE PUNKTE (vor weiterer Automatisierung zu klären):
  1. Ist HUB nur mit ERF7 kombinierbar oder auch mit ERF8?
     (Screenshot zeigt ERF8/HUB, aktuelle Regel sagt nur ERF7)
  2. Was genau bedeutet KGS? Im Code wird KGT verwendet (= Kein Tagesgeschäft).
     KGS könnte eine Schreibvariante oder ein eigenes Kürzel sein.
  3. Was bedeutet ONB genau? Es gibt Ausschlussregeln, aber keine Definition.
  4. Was bedeutet TPA genau? Bisher nur "NM möglich mit TEL".
  5. Welche Personen sind TL? Aktuell definiert: Silvana, Linda, Lara.
  6. Wie genau rotiert die Tagesverantwortung?
     Aktuell: Lara/Silvana wechseln wöchentlich, Linda fix Di-NM + Mi-VM + Fr-NM.
  7. 2-Wochen-Logik: Startwochen-Referenz für gerade/ungerade KW
     (Linda Fr frei, Isaura Mi frei, Corinne Fr frei, Linda HO Di, Schalter).
  8. ERF9 Mo/Di: Wird als Ganztags-Aufgabe behandelt (VM + NM).
  9. Florence Di PO: Aktuell als Muss-Regel implementiert. Muss oder Soll?
 10. Stephi (Rämi Stephanie) erscheint im Screenshot, ist als KSC-Leiterin
     mit permanentem KGT implementiert. Stammdaten ggf. ergänzen.
 11. VBZ/Q: Im Screenshot sichtbar, Bedeutung teilweise unklar.
     Aktuell: Vorbezüge + Queue, nur Lara und Nina.
"""

import json
import random
import os
from datetime import datetime, timedelta
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# KONFIGURATION
# ============================================================

DAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
SLOTS = ["Vormittag", "Nachmittag"]
DAY_SLOTS = [(d, s) for d in DAYS for s in SLOTS]

# Farben für Aufgaben (RGB hex ohne #)
COLORS = {
    "TEL":       "92D050",  # Grün
    "ABKL":      "FFC000",  # Orange
    "ERF7":      "00B0F0",  # Blau (bis 16:00)
    "ERF7/Q":    "00B0F0",  # ERF7 + Queue
    "ERF7/HUB":  "00B0F0",  # ERF7 + Hub
    "ERF8":      "FF6600",  # Dunkelorange (bis 17:00)
    "ERF8/Q":    "FF6600",  # ERF8 + Queue
    "ERF8/HUB":  "FF6600",  # ERF8 + Hub
    "ERF9":      "FF0000",  # Rot (Auftragserfassung)
    "ERF9/TEL":  "FF0000",  # ERF9 + Telefon
    "ERF4/SCH":  "C6EFCE",  # Schalter - hellgrün
    "ERF5":      "F4B084",  # ERF5 (Hellorange)
    "PO":        "D9E1F2",  # Postöffnung
    "PO/ABKL":   "FFC000",  # Post + Abklärung
    "PO/TEL":    "92D050",  # Post + Telefon
    "PO/SCAN":   "E2EFDA",  # Post/Scanning
    "HO":        "BDD7EE",  # Home Office
    "HO/Q":      "BDD7EE",  # Home Office + Queue
    "VBZ/Q":     "E2EFDA",  # Vorbezüge + Queue
    "KSC Spez.": "FFFF00",  # Kundenservicecenter Spezial
    "KGT":       "FFFF00",  # Kein Tagesgeschäft
    "TAGESPA":   "00FFFF",  # Tagespharma
    "RX Abo":    "D9D9D9",  # Rezept Abo
    "Direkt":    "B4C6E7",  # Direktbestellung
    "ONB":       "D9E1F2",
    "BTM":       "F8CBAD",
    "KTG":       "A9D18E",  # Kranktaggeld
    "KRANK":     "FF9999",  # Krank
    "FERIEN":    "F2F2F2",  # Ferien
    "":          "FFFFFF",
}

# Farben für Tagesverantwortung (Name-Zelle wird eingefärbt)
TV_COLORS = {
    "Silvana":    "FF6600",  # Orange
    "Linda":      "FF00FF",  # Pink
    "Lara":       "00B050",  # Grün
    "Andrea G.":  "00B0F0",  # Blau
    "Dipiga":     "FFFF00",  # Gelb
    "Isaura":     "BDD7EE",  # Hellblau
}

# 18 Uhr / PHC-Liste = blau hinterlegt
PHC_COLOR = "0070C0"  # Dunkelblau (Text weiss)

# ============================================================
# MITARBEITER-DEFINITIONEN
# ============================================================

@dataclass
class Employee:
    name: str
    short: str
    pct: int
    # Verfügbarkeit: dict von (day_idx, slot_idx) -> True/False
    available: dict = field(default_factory=dict)
    # Fähigkeiten/Rollen
    can_hub: bool = False
    can_rx_abo: bool = False
    can_schalter: bool = True
    can_scanning: bool = True
    can_onb: bool = True
    can_direkt: bool = False
    can_vbz: bool = False
    is_tl: bool = False
    # Tracking
    schalter_last_week: bool = False
    tel_count: int = 0
    abkl_count: int = 0

    def is_available(self, day_idx, slot_idx):
        return self.available.get((day_idx, slot_idx), False)


def create_employees(week_number):
    """Erstellt alle Mitarbeiter mit Verfügbarkeit basierend auf Wochennummer."""
    is_even_week = (week_number % 2 == 0)

    employees = {}

    # --- Silvana Grossenbacher - 100% (TL) ---
    e = Employee("Grossenbacher Silvana", "Silvana", 100, can_direkt=True, is_tl=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Silvana"] = e

    # --- Rexhaj Majlinda (Linda) - 90%, jeden 2. Freitag frei (TL) ---
    e = Employee("Rexhaj Majlinda", "Linda", 90, can_rx_abo=True, is_tl=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    if is_even_week:
        e.available[(4, 0)] = False
        e.available[(4, 1)] = False
    employees["Linda"] = e

    # --- Lara Ierano - 100% (TL) ---
    e = Employee("Lara Ierano", "Lara", 100, can_rx_abo=True, can_vbz=True, is_tl=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    e.can_hub = True
    employees["Lara"] = e

    # --- Gygax Andrea - 100% ---
    e = Employee("Gygax Andrea", "Andrea G.", 100, can_direkt=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Andrea G."] = e

    # --- Jeyalingam Dipiga - 100% ---
    e = Employee("Jeyalingam Dipiga", "Dipiga", 100, can_hub=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Dipiga"] = e

    # --- Bohnenblust Isaura - 90%, jeden 2. Mittwoch frei ---
    e = Employee("Bohnenblust Isaura", "Isaura", 90, can_rx_abo=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    if is_even_week:
        e.available[(2, 0)] = False
        e.available[(2, 1)] = False
    employees["Isaura"] = e

    # --- Martina Pizzi - 80%, jeden Montag frei ---
    e = Employee("Martina Pizzi", "Martina", 80, can_rx_abo=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    e.available[(0, 0)] = False
    e.available[(0, 1)] = False
    employees["Martina"] = e

    # --- Giombanco Alessia - 100% ---
    e = Employee("Giombanco Alessia", "Alessia", 100, can_hub=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Alessia"] = e

    # --- Imsirovic Amra - 100% ---
    e = Employee("Imsirovic Amra", "Amra", 100, can_hub=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Amra"] = e

    # --- Hänni Nina - 100% ---
    e = Employee("Hänni Nina", "Nina", 100, can_hub=True, can_vbz=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Nina"] = e

    # --- Bushaj Jesika - 100% ---
    e = Employee("Bushaj Jesika", "Jesika", 100, can_hub=True)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    employees["Jesika"] = e

    # --- Siegrist Brigitte - 70%, Mi/Do/Fr NM frei ---
    e = Employee("Siegrist Brigitte", "Brigitte", 70,
                 can_schalter=False, can_scanning=False, can_onb=False)
    for d in range(5):
        e.available[(d, 0)] = True
        e.available[(d, 1)] = True
    e.available[(2, 1)] = False  # Mi NM
    e.available[(3, 1)] = False  # Do NM
    e.available[(4, 1)] = False  # Fr NM
    employees["Brigitte"] = e

    # --- Florence Dornbierer - 70%, Mi NM + Fr frei ---
    e = Employee("Florence Dornbierer", "Florence", 70,
                 can_schalter=False, can_onb=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    e.available[(2, 1)] = False  # Mi NM
    e.available[(4, 0)] = False  # Fr
    e.available[(4, 1)] = False
    employees["Florence"] = e

    # --- Eggimann Corinne - 90%, jeden 2. Freitag frei ---
    e = Employee("Eggimann Corinne", "Corinne", 90,
                 can_hub=True, can_schalter=False, can_scanning=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = True
    if is_even_week:
        e.available[(4, 0)] = False
        e.available[(4, 1)] = False
    employees["Corinne"] = e

    # --- Schöni Saskia - 40%, nur Mi und Fr ---
    e = Employee("Schöni Saskia", "Saskia", 40,
                 can_schalter=False, can_scanning=False, can_onb=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = False
    e.available[(2, 0)] = True  # Mi VM
    e.available[(2, 1)] = True  # Mi NM
    e.available[(4, 0)] = True  # Fr VM
    e.available[(4, 1)] = True  # Fr NM
    employees["Saskia"] = e

    # --- Milenkovic Dragana - 60%, Di/Do/Fr hier ---
    e = Employee("Milenkovic Dragana", "Dragi", 60,
                 can_hub=True, can_schalter=False, can_scanning=False, can_onb=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = False
    e.available[(1, 0)] = True  # Di
    e.available[(1, 1)] = True
    e.available[(3, 0)] = True  # Do
    e.available[(3, 1)] = True
    e.available[(4, 0)] = True  # Fr
    e.available[(4, 1)] = True
    employees["Dragi"] = e

    # --- Bruzzese Maria - 30%, Mo und Do VM ---
    e = Employee("Bruzzese Maria", "Maria B.", 30,
                 can_schalter=False, can_scanning=False, can_onb=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = False
    e.available[(0, 0)] = True  # Mo VM
    e.available[(3, 0)] = True  # Do VM
    employees["Maria B."] = e

    # --- Ackermann Andrea - 60%, Mo/Di/Mi hier ---
    e = Employee("Ackermann Andrea", "Andrea A.", 60,
                 can_direkt=True, can_schalter=False, can_onb=False)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = False
    e.available[(0, 0)] = True  # Mo
    e.available[(0, 1)] = True
    e.available[(1, 0)] = True  # Di
    e.available[(1, 1)] = True
    e.available[(2, 0)] = True  # Mi
    e.available[(2, 1)] = True
    employees["Andrea A."] = e

    # --- Stucki Stephanie (Stephi) - KSC-Leiterin, nicht im Tagesgeschäft ---
    e = Employee("Stucki Stephanie", "Stephi", 100)
    for d in range(5):
        for s in range(2):
            e.available[(d, s)] = False  # Nicht im Tagesgeschäft
    employees["Stephi"] = e

    return employees


# ============================================================
# SCHEDULE-KLASSE
# ============================================================

class Schedule:
    def __init__(self, employees, week_number, week_start_date):
        self.employees = employees
        self.week_number = week_number
        self.week_start_date = week_start_date
        self.is_even_week = (week_number % 2 == 0)
        # schedule[name][(day_idx, slot_idx)] = task_string
        self.schedule = {name: {} for name in employees}
        # Tracking für Constraints
        self.tel_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        self.abkl_count = {d: {s: 0 for s in range(2)} for d in range(5)}
        self.erf8_assigned = {d: False for d in range(5)}
        # Notizen / Sonderwünsche
        self.notes = {}  # (name, day, slot) -> note string (e.g. "Arzttermin 14.45")
        # Zeitangaben rechts
        self.time_notes = {}  # name -> string
        # Tagesverantwortung: day -> name
        self.tagesverantwortung = {}
        # 18 Uhr / PHC-Liste: day -> name (blau hinterlegt)
        self.phc_liste = {}

    def assign(self, name, day, slot, task):
        """Weist eine Aufgabe zu, wenn der Slot noch frei ist."""
        if (day, slot) in self.schedule[name]:
            return False
        if not self.employees[name].is_available(day, slot):
            return False
        self.schedule[name][(day, slot)] = task
        if "TEL" in task:
            self.tel_count[day][slot] += 1
        if "ABKL" in task:
            self.abkl_count[day][slot] += 1
        return True

    def is_free(self, name, day, slot):
        return ((day, slot) not in self.schedule[name]
                and self.employees[name].is_available(day, slot))

    def get_task(self, name, day, slot):
        return self.schedule[name].get((day, slot), "")

    def get_assigned_names(self, day, slot, task_contains=""):
        """Gibt Namen zurück, die an diesem Slot eine bestimmte Aufgabe haben."""
        result = []
        for name in self.employees:
            t = self.get_task(name, day, slot)
            if task_contains and task_contains in t:
                result.append(name)
            elif not task_contains and t:
                result.append(name)
        return result

    def count_task(self, day, slot, task):
        return len(self.get_assigned_names(day, slot, task))

    def get_available_for_task(self, day, slot, exclude=None):
        """Gibt verfügbare Mitarbeiter für einen Slot zurück."""
        result = []
        for name, emp in self.employees.items():
            if exclude and name in exclude:
                continue
            if self.is_free(name, day, slot):
                result.append(name)
        return result


# ============================================================
# SCHEDULER
# ============================================================

def build_schedule(week_number, week_start_date, overrides=None, state_file="scheduler_state.json"):
    """
    Erstellt den Wochenplan.
    overrides: dict von (name, day_idx, slot_idx) -> task_string für Sonderwünsche
    """
    employees = create_employees(week_number)
    sched = Schedule(employees, week_number, week_start_date)

    # Lade gespeicherten State (für Wechsel-Tracking)
    state = load_state(state_file)

    # ========================================
    # WOCHE-GUARD: Nur einmal pro Woche toggeln
    # Wenn erneut für gleiche Woche, State nicht ändern
    # ========================================
    last_week = state.get("last_generated_week", 0)
    is_new_week = (last_week != week_number)
    if is_new_week:
        state["last_generated_week"] = week_number
        # Alle Wechsel-Toggles drehen sich nur bei neuer Woche
        state["linda_ho_week"] = not state.get("linda_ho_week", False)
        state["labor_jesika"] = not state.get("labor_jesika", True)
        state["schalter_group"] = 1 - state.get("schalter_group", 0)
        state["friday_tagesv_corinne"] = not state.get("friday_tagesv_corinne", True)
        state["tl_monday_idx"] = state.get("tl_monday_idx", 0) + 1
        state["tv_lara_week"] = not state.get("tv_lara_week", False)
    # else: Re-Run für gleiche Woche → State bleibt identisch

    # ========================================
    # 0. SONDERWÜNSCHE / OVERRIDES eintragen
    # ========================================
    if overrides:
        for (name, day, slot), task in overrides.items():
            if name in employees:
                sched.assign(name, day, slot, task)
                if task.startswith("*"):
                    sched.notes[(name, day, slot)] = task

    # ========================================
    # 1. FESTE ZUWEISUNGEN
    # ========================================

    # Stephi (KSC-Leiterin) — nicht im Tagesgeschäft, kein Eintrag im Plan

    # Maria B. Montagmorgen HO
    if "Maria B." in employees:
        sched.assign("Maria B.", 0, 0, "HO")

    # Brigitte jeden Montag ERF9 (ganztags)
    sched.assign("Brigitte", 0, 0, "ERF9")
    sched.assign("Brigitte", 0, 1, "ERF9")

    # Dragi jeden Dienstag ERF9
    if employees["Dragi"].is_available(1, 0):
        sched.assign("Dragi", 1, 0, "ERF9")
    if employees["Dragi"].is_available(1, 1):
        sched.assign("Dragi", 1, 1, "ERF9")

    # Florence PÖ wenn am Dienstag
    if employees["Florence"].is_available(1, 0):
        sched.assign("Florence", 1, 0, "PO")

    # Dipiga KGT: Di-morgens und Do-nachmittags
    sched.assign("Dipiga", 1, 0, "KGT")
    sched.assign("Dipiga", 3, 1, "KGT")

    # Martina KSC Spez.: Di und Do morgens
    if employees["Martina"].is_available(1, 0):
        sched.assign("Martina", 1, 0, "KSC Spez.")
    if employees["Martina"].is_available(3, 0):
        sched.assign("Martina", 3, 0, "KSC Spez.")

    # Linda alle 2 Wochen Dienstag HO (State-basiert)
    if state.get("linda_ho_week", False):
        sched.assign("Linda", 1, 0, "HO")
        sched.assign("Linda", 1, 1, "HO")

    # ERF5: Jesika und Dipiga im Wochenwechsel, Mi NM (State-basiert)
    if state.get("labor_jesika", True):
        sched.assign("Jesika", 2, 1, "ERF5")
    else:
        sched.assign("Dipiga", 2, 1, "ERF5")

    # ========================================
    # 2. TEL - HÖCHSTE PRIORITÄT
    #    4 am Montagmorgen, sonst 3
    #    Soll-Regel: Jede TL soll 1-2x pro Woche TEL machen
    # ========================================
    tl_names = [n for n, emp in employees.items() if emp.is_tl]
    tl_tel_count = {n: 0 for n in tl_names}

    for day in range(5):
        for slot in range(2):
            target = 4 if (day == 0 and slot == 0) else 3
            current = sched.tel_count[day][slot]
            needed = target - current
            if needed <= 0:
                continue
            available = sched.get_available_for_task(day, slot)
            # TL bevorzugt einplanen, wenn sie noch < 2x TEL haben
            tl_available = [n for n in available
                            if n in tl_tel_count and tl_tel_count[n] < 2]
            other_available = [n for n in available if n not in tl_names]
            random.shuffle(tl_available)
            random.shuffle(other_available)
            # Zuerst TL, dann Rest
            prioritized = tl_available + other_available
            count = 0
            for name in prioritized:
                if count >= needed:
                    break
                if sched.assign(name, day, slot, "TEL"):
                    count += 1
                    if name in tl_tel_count:
                        tl_tel_count[name] += 1

    # ========================================
    # 3. ABKL - 2 pro Halbtag
    # ========================================
    for day in range(5):
        for slot in range(2):
            needed = 2 - sched.abkl_count[day][slot]
            available = sched.get_available_for_task(day, slot)
            random.shuffle(available)
            count = 0
            for name in available:
                if count >= needed:
                    break
                if sched.assign(name, day, slot, "ABKL"):
                    count += 1

    # ========================================
    # 4. ERF7/HUB ZUWEISUNGEN (halbtags)
    # ========================================
    hub_eligible = ["Jesika", "Dragi", "Lara", "Corinne", "Dipiga", "Nina", "Amra", "Alessia"]
    hub_assigned_today = {d: 0 for d in range(5)}

    for day in range(5):
        random.shuffle(hub_eligible)
        for name in hub_eligible:
            if hub_assigned_today[day] >= 2:
                break
            for slot in [0, 1]:
                if sched.is_free(name, day, slot):
                    sched.assign(name, day, slot, "ERF7/HUB")
                    hub_assigned_today[day] += 1
                    break

    # ========================================
    # 5. ERF9 Mi-Fr (undefiniert, 1 Person zuweisen)
    # ========================================
    erf9_candidates = [n for n in employees
                       if n not in ["Brigitte", "Dragi"]]
    random.shuffle(erf9_candidates)
    for day in [2, 3, 4]:  # Mi, Do, Fr
        for name in erf9_candidates:
            if sched.is_free(name, day, 0):
                sched.assign(name, day, 0, "ERF9")
                break

    # ========================================
    # 6. ERF8 - 1 Person pro Tag
    # ========================================
    erf8_candidates = [n for n in employees if n not in ["Maria B.", "Saskia"]]
    random.shuffle(erf8_candidates)
    for day in range(5):
        for name in erf8_candidates:
            if sched.is_free(name, day, 0) or sched.is_free(name, day, 1):
                slot = 0 if sched.is_free(name, day, 0) else 1
                sched.assign(name, day, slot, "ERF8")
                sched.erf8_assigned[day] = True
                break

    # ========================================
    # 7. SCHALTER (ERF4/SCH) - alle 2 Wochen
    # ========================================
    # Feste, stabile Gruppen für Schalter-Rotation (nicht shufflen!)
    schalter_eligible = sorted(
        [name for name, emp in employees.items()
         if emp.can_schalter and name not in ["Linda", "Florence", "Corinne",
                                               "Maria B.", "Andrea A.",
                                               "Brigitte", "Saskia", "Dragi",
                                               "Stephi"]],
    )
    schalter_this_week = state.get("schalter_group", 0)
    mid = len(schalter_eligible) // 2
    if schalter_this_week == 0:
        schalter_active = schalter_eligible[:mid]
    else:
        schalter_active = schalter_eligible[mid:]

    for name in schalter_active:
        assigned = False
        for day in range(5):
            if assigned:
                break
            for slot in [0, 1]:
                if sched.is_free(name, day, slot):
                    sched.assign(name, day, slot, "ERF4/SCH")
                    assigned = True
                    break

    # ========================================
    # 8. RX ABO: Lara, Linda, Isaura, Martina
    # ========================================
    rx_abo_people = ["Lara", "Linda", "Isaura", "Martina"]
    for name in rx_abo_people:
        assigned = False
        for day in range(5):
            if assigned:
                break
            for slot in [0, 1]:
                if sched.is_free(name, day, slot):
                    sched.assign(name, day, slot, "RX Abo")
                    assigned = True
                    break

    # ========================================
    # 9. DIREKTBESTELLUNG: Andrea A., Andrea G., Silvana
    # ========================================
    direkt_people = ["Andrea A.", "Andrea G.", "Silvana"]
    for name in direkt_people:
        assigned = False
        for day in range(5):
            if assigned:
                break
            for slot in [1, 0]:  # NM bevorzugt
                if sched.is_free(name, day, slot):
                    sched.assign(name, day, slot, "Direkt")
                    assigned = True
                    break

    # ========================================
    # 10. VORBEZÜGE: Lara, Nina
    # ========================================
    for name in ["Lara", "Nina"]:
        for day in range(5):
            if sched.is_free(name, day, 0):
                sched.assign(name, day, 0, "VBZ/Q")
                break

    # ========================================
    # 11. PO/SCANNING für übrige (wo erlaubt, max 2 pro Tag)
    # ========================================
    no_scanning = ["Brigitte", "Corinne", "Maria B.", "Saskia", "Dragi"]
    po_scan_today = {d: 0 for d in range(5)}
    scanning_eligible = [n for n in employees if n not in no_scanning]
    random.shuffle(scanning_eligible)
    for day in range(5):
        for name in scanning_eligible:
            if po_scan_today[day] >= 2:
                break
            for slot in range(2):
                if sched.is_free(name, day, slot) and po_scan_today[day] < 2:
                    sched.assign(name, day, slot, "PO/SCAN")
                    po_scan_today[day] += 1

    # ========================================
    # 12. Restliche Slots mit ERF7 füllen (TL bekommen ERF7/Q)
    # ========================================
    for name in employees:
        for day in range(5):
            for slot in range(2):
                if sched.is_free(name, day, slot):
                    if employees[name].is_tl:
                        sched.assign(name, day, slot, "ERF7/Q")
                    else:
                        sched.assign(name, day, slot, "ERF7")

    # ========================================
    # 14. 18 UHR / PHC-LISTE (blau, 1 Pers/Tag)
    # ========================================
    # Mo: TL im Wechsel, Di: Dipiga, Mi/Do: undefiniert, Fr: Corinne/Andrea G. im Wechsel
    phc = {}
    # Montag: TL im Wechsel
    tl_idx = state.get("tl_monday_idx", 0) % 2
    phc[0] = f"TL ({tl_idx + 1})"
    # Dienstag: Dipiga
    phc[1] = "Dipiga"
    # Mittwoch: undefiniert → zufällig aus verfügbaren
    # Donnerstag: undefiniert → zufällig aus verfügbaren
    available_for_phc = [n for n in employees
                         if n not in ["Saskia", "Maria B."]]
    random.shuffle(available_for_phc)
    for day in [2, 3]:
        for name in available_for_phc:
            if employees[name].is_available(day, 1):  # NM muss da sein
                if name not in phc.values():
                    phc[day] = name
                    break
    # Freitag: Corinne/Andrea G. im Wechsel
    if state.get("friday_tagesv_corinne", True):
        phc[4] = "Corinne"
    else:
        phc[4] = "Andrea G."

    sched.phc_liste = phc

    # ========================================
    # 15. TAGESVERANTWORTUNG (farbiger Name)
    # ========================================
    # Woche A (tv_lara_week=True): Lara 3 Halbtage + Linda 2 Halbtage
    # Woche B (tv_lara_week=False): Silvana 3 Halbtage + Linda 2 Halbtage
    # Verteilung: Mo-VM, Mo-NM, Di-VM = Hauptperson (3 Halbtage)
    #             Di-NM, Mi-VM = Linda (2 Halbtage)
    #             Mi-NM bis Fr-NM = verbleibende TL rotierend
    tv = {}
    if state.get("tv_lara_week", False):
        haupt_tl = "Lara"
        rest_tl = "Silvana"
    else:
        haupt_tl = "Silvana"
        rest_tl = "Lara"
    # Erste 3 Halbtage: Haupt-TL
    tv[(0, 0)] = haupt_tl  # Mo VM
    tv[(0, 1)] = haupt_tl  # Mo NM
    tv[(1, 0)] = haupt_tl  # Di VM
    # Nächste 2 Halbtage: Linda
    tv[(1, 1)] = "Linda"   # Di NM
    tv[(2, 0)] = "Linda"   # Mi VM
    # Restliche 5 Halbtage: Rotation zwischen rest_tl und haupt_tl
    tv[(2, 1)] = rest_tl   # Mi NM
    tv[(3, 0)] = rest_tl   # Do VM
    tv[(3, 1)] = rest_tl   # Do NM
    tv[(4, 0)] = haupt_tl  # Fr VM
    tv[(4, 1)] = "Linda"   # Fr NM
    sched.tagesverantwortung = tv

    # State speichern
    save_state(state, state_file)

    return sched


# ============================================================
# STATE PERSISTENCE (für Wechsel über Wochen)
# ============================================================

def load_state(filepath):
    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            return json.load(f)
    return {}

def save_state(state, filepath):
    with open(filepath, 'w') as f:
        json.dump(state, f, indent=2)


# ============================================================
# EXCEL-AUSGABE
# ============================================================

def _write_tv_merged_row(ws, row, tv_dict, names_for_row, thin_border):
    """Schreibt eine TV-Zeile mit gemerged Blöcken für die angegebenen TL-Namen."""
    halbtage = [(d, s) for d in range(5) for s in range(2)]

    # Zusammenhängende Runs gleicher Namen finden
    runs = []  # (start_col, end_col, name)
    current_name = None
    start_col = None
    end_col = None
    for i, (day, slot) in enumerate(halbtage):
        col = 3 + i
        tv_name = tv_dict.get((day, slot), "")
        show = tv_name if tv_name in names_for_row else ""
        if show == current_name and show != "":
            end_col = col
        else:
            if current_name and current_name != "":
                runs.append((start_col, end_col, current_name))
            current_name = show
            start_col = col
            end_col = col
    if current_name and current_name != "":
        runs.append((start_col, end_col, current_name))

    # Alle Zellen erst leer + Border
    for col in range(3, 13):
        c = ws.cell(row=row, column=col, value="")
        c.border = thin_border

    # Runs schreiben + mergen
    for s_col, e_col, name in runs:
        color = TV_COLORS.get(name, 'FFFFFF')
        for col in range(s_col, e_col + 1):
            c = ws.cell(row=row, column=col)
            c.fill = PatternFill('solid', fgColor=color)
            c.border = thin_border
        # Text nur in der ersten Zelle des Runs
        c = ws.cell(row=row, column=s_col, value=name)
        c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
        c.alignment = Alignment(horizontal='center')
        if e_col > s_col:
            ws.merge_cells(start_row=row, start_column=s_col,
                           end_row=row, end_column=e_col)


def write_excel(sched, output_path):
    """Schreibt den Arbeitsplan als formatierte Excel-Datei."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Arbeitsplan"

    # Datumstring
    start = sched.week_start_date
    end = start + timedelta(days=4)
    date_str = f"{start.strftime('%d.%m.%Y')}-{end.strftime('%d.%m.%Y')}"

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_font = Font(bold=True, size=10, name='Arial')
    cell_font = Font(size=9, name='Arial')
    title_font = Font(bold=True, size=12, name='Arial')

    # Titel
    ws.merge_cells('A1:L1')
    ws['A1'] = date_str
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center')

    # Header Zeile 2 + 3
    row = 3
    headers_row1 = ["Name", "%"]
    for day in DAYS:
        headers_row1.extend([day, ""])
    # Zeitnotizen
    headers_row1.append("Zeiten")

    headers_row2 = ["", ""]
    for _ in DAYS:
        headers_row2.extend(["Vormittag", "Nachmittag"])
    headers_row2.append("")

    for col, val in enumerate(headers_row1, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor='4472C4')
        cell.font = Font(bold=True, size=10, name='Arial', color='FFFFFF')
        cell.border = thin_border

    row = 4
    for col, val in enumerate(headers_row2, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(bold=True, size=8, name='Arial')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor='D9E2F3')
        cell.border = thin_border

    # Merge Tag-Header über 2 Spalten
    for i, day in enumerate(DAYS):
        start_col = 3 + i * 2
        ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=start_col + 1)

    # Spaltenbreiten
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 5
    for i in range(5):
        ws.column_dimensions[get_column_letter(3 + i * 2)].width = 12
        ws.column_dimensions[get_column_letter(4 + i * 2)].width = 12
    ws.column_dimensions[get_column_letter(13)].width = 18

    # Mitarbeiter-Reihenfolge (wie im Bild)
    order = [
        "Silvana", "Linda", "Lara", "Andrea G.", "Dipiga",
        "Isaura", "Martina", "Alessia", "Amra", "Nina",
        "Jesika", "Brigitte", "Florence", "Corinne", "Saskia",
        "Dragi", "Maria B.", "Andrea A."
    ]

    row = 5
    for name in order:
        if name not in sched.employees:
            continue
        emp = sched.employees[name]

        # Name-Zelle
        cell = ws.cell(row=row, column=1, value=emp.name)
        cell.font = cell_font
        cell.border = thin_border

        # Prozent
        cell = ws.cell(row=row, column=2, value=emp.pct)
        cell.font = cell_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

        # Aufgaben pro Tag/Slot
        for day in range(5):
            for slot in range(2):
                col = 3 + day * 2 + slot
                task = sched.get_task(name, day, slot)

                if not emp.is_available(day, slot) and not task:
                    task = ""

                cell = ws.cell(row=row, column=col, value=task)
                cell.font = cell_font
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border

                # Farbe setzen
                color = get_color(task)
                if color:
                    cell.fill = PatternFill('solid', fgColor=color)

        # Zeitnotiz
        if name in sched.time_notes:
            cell = ws.cell(row=row, column=13, value=sched.time_notes[name])
            cell.font = cell_font
            cell.border = thin_border

        # Sonderwunsch-Notizen (z.B. *Arzttermin 14.45) in Zeitspalte
        for day in range(5):
            for slot in range(2):
                note_key = (name, day, slot)
                if note_key in sched.notes:
                    note_text = sched.notes[note_key].lstrip("*")
                    cell = ws.cell(row=row, column=13, value=note_text)
                    cell.font = cell_font
                    cell.border = thin_border

        row += 1

    # ========================================
    # Tagesverantwortung (2 Zeilen, farbig gemerged)
    # ========================================
    if sched.tagesverantwortung:
        row += 1  # Leerzeile
        tv = sched.tagesverantwortung

        # TL-Namen in 2 Zeilen aufteilen: Zeile 1 = haupt/rest_tl, Zeile 2 = Linda
        all_tv_names = set(tv.values())
        linda_names = {n for n in all_tv_names if n == "Linda"}
        other_names = all_tv_names - linda_names

        # Zeile 1: haupt_tl + rest_tl
        _write_tv_merged_row(ws, row, tv, other_names, thin_border)
        names_sorted = sorted(other_names)
        if len(names_sorted) >= 1:
            c = ws.cell(row=row, column=1, value=names_sorted[0])
            c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
            c.fill = PatternFill('solid', fgColor=TV_COLORS.get(names_sorted[0], 'FFFFFF'))
            c.border = thin_border
        if len(names_sorted) >= 2:
            c = ws.cell(row=row, column=2, value=names_sorted[1])
            c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
            c.fill = PatternFill('solid', fgColor=TV_COLORS.get(names_sorted[1], 'FFFFFF'))
            c.border = thin_border
        else:
            ws.cell(row=row, column=2, value="").border = thin_border

        # Zeile 2: Linda
        row += 1
        _write_tv_merged_row(ws, row, tv, linda_names, thin_border)
        c = ws.cell(row=row, column=1, value="Silvana" if "Silvana" not in other_names else "")
        c.border = thin_border
        if linda_names:
            c = ws.cell(row=row, column=2, value="Linda")
            c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
            c.fill = PatternFill('solid', fgColor=TV_COLORS.get("Linda", 'FF00FF'))
            c.border = thin_border
        else:
            ws.cell(row=row, column=2, value="").border = thin_border

    # ========================================
    # PHC-Liste / 18 Uhr Zeile (blau)
    # ========================================
    if sched.phc_liste:
        row += 1
        cell = ws.cell(row=row, column=1, value="18 Uhr / PHC-Liste")
        cell.font = Font(bold=True, size=9, name='Arial')
        cell.border = thin_border
        cell = ws.cell(row=row, column=2, value="")
        cell.border = thin_border

        for day in range(5):
            phc_name = sched.phc_liste.get(day, "")
            # PHC spans both VM and NM columns for the day
            for slot in range(2):
                col = 3 + day * 2 + slot
                cell = ws.cell(row=row, column=col, value=phc_name if slot == 0 else "")
                cell.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border
                cell.fill = PatternFill('solid', fgColor=PHC_COLOR)
            # Merge the two cells for the day
            start_col = 3 + day * 2
            ws.merge_cells(start_row=row, start_column=start_col,
                           end_row=row, end_column=start_col + 1)

    wb.save(output_path)
    return output_path


def get_color(task):
    if not task:
        return None
    for key in COLORS:
        if task == key or task.startswith(key):
            return COLORS[key]
    return None


# ============================================================
# SONDERWÜNSCHE EINLESEN
# ============================================================

def load_overrides(filepath="sonderwuensche.json"):
    """
    Lädt Sonderwünsche aus einer JSON-Datei.
    Format:
    {
        "overrides": [
            {"name": "Martina", "day": 4, "slot": 1, "task": "*Arzttermin 14.45"},
            {"name": "Isaura", "day": 2, "slot": 1, "task": "ONB"}
        ]
    }
    day: 0=Mo, 1=Di, 2=Mi, 3=Do, 4=Fr
    slot: 0=Vormittag, 1=Nachmittag
    task mit * = wird als Notiz angezeigt (z.B. Arzttermin)
    """
    if not os.path.exists(filepath):
        return {}
    with open(filepath, 'r') as f:
        data = json.load(f)
    result = {}
    for item in data.get("overrides", []):
        key = (item["name"], item["day"], item["slot"])
        result[key] = item["task"]
    return result


# ============================================================
# HAUPTPROGRAMM MIT INTERAKTIVER EINGABE
# ============================================================

EMPLOYEE_NAMES = [
    "Stephi",
    "Silvana", "Linda", "Lara", "Andrea G.", "Dipiga",
    "Isaura", "Martina", "Alessia", "Amra", "Nina",
    "Jesika", "Brigitte", "Florence", "Corinne", "Saskia",
    "Dragi", "Maria B.", "Andrea A."
]

DAY_NAMES_SHORT = ["Mo", "Di", "Mi", "Do", "Fr"]

def print_employee_list():
    print("\n  Mitarbeiterinnen:")
    for i, name in enumerate(EMPLOYEE_NAMES, 1):
        print(f"    {i:2d}. {name}")

def pick_employee():
    """Lässt den User eine Mitarbeiterin wählen."""
    print_employee_list()
    while True:
        inp = input("\n  Nummer oder Name (Enter = fertig): ").strip()
        if not inp:
            return None
        if inp.isdigit():
            idx = int(inp) - 1
            if 0 <= idx < len(EMPLOYEE_NAMES):
                return EMPLOYEE_NAMES[idx]
        # Suche nach Name (case-insensitive, Teilmatch)
        matches = [n for n in EMPLOYEE_NAMES if inp.lower() in n.lower()]
        if len(matches) == 1:
            return matches[0]
        elif len(matches) > 1:
            print(f"    Mehrere Treffer: {', '.join(matches)}")
            print(f"    Bitte genauer eingeben.")
        else:
            print(f"    '{inp}' nicht gefunden. Nochmal versuchen.")

def pick_days():
    """Lässt den User Tage wählen."""
    print(f"\n  Welche Tage?")
    print(f"    1=Mo  2=Di  3=Mi  4=Do  5=Fr")
    print(f"    Mehrere mit Komma (z.B. '1,3,5') oder 'alle'")
    while True:
        inp = input("  Tage: ").strip().lower()
        if inp in ["alle", "all", "a"]:
            return [0, 1, 2, 3, 4]
        try:
            days = [int(x.strip()) - 1 for x in inp.split(",")]
            if all(0 <= d <= 4 for d in days):
                return days
        except ValueError:
            pass
        print("    Ungültige Eingabe. Beispiel: '1,3' für Mo+Mi")

def pick_slots():
    """Lässt den User Halbtage wählen."""
    print(f"\n  Welche Halbtage?")
    print(f"    1=Vormittag  2=Nachmittag  3=Ganzer Tag")
    while True:
        inp = input("  Halbtag: ").strip()
        if inp == "1":
            return [0]
        elif inp == "2":
            return [1]
        elif inp in ["3", "beide", "ganzer tag", "ganz"]:
            return [0, 1]
        print("    Bitte 1, 2 oder 3 eingeben.")

def pick_absence_type():
    """Lässt den User den Abwesenheitstyp wählen."""
    print(f"\n  Was ist der Grund?")
    print(f"    1 = Krank (ganzer Tag/Tage)")
    print(f"    2 = Ferien / Abwesend (ganzer Tag/Tage)")
    print(f"    3 = Arzttermin / fixer Termin (Halbtag)")
    print(f"    4 = Anderer Sonderwunsch (z.B. HO, bestimmte Aufgabe)")
    while True:
        inp = input("  Typ (1-4): ").strip()
        if inp in ["1", "2", "3", "4"]:
            return int(inp)
        print("    Bitte 1-4 eingeben.")


def get_next_monday(ref_date=None):
    """Gibt den nächsten Montag zurück."""
    if ref_date is None:
        ref_date = datetime.now()
    days_ahead = 0 - ref_date.weekday()
    if days_ahead <= 0:
        days_ahead += 7
    return ref_date + timedelta(days=days_ahead)


def check_rule_impact(absences, employees):
    """
    Prüft welche Regeln durch Abwesenheiten verletzt werden könnten.
    absences: list von (name, day, slot)
    """
    warnings = []
    absent_set = set(absences)

    for name, day, slot in absences:
        day_name = DAYS[day]
        slot_name = SLOTS[slot]
        tag = f"{name} {day_name} {slot_name}"

        # Brigitte Mo ERF9
        if name == "Brigitte" and day == 0:
            warnings.append(f"⚠️  {tag}: Brigitte hat normalerweise ERF9 am Montag! → ERF9 Mo muss neu besetzt werden.")

        # Dragi Di ERF9
        if name == "Dragi" and day == 1:
            warnings.append(f"⚠️  {tag}: Dragi hat normalerweise ERF9 am Dienstag! → ERF9 Di muss neu besetzt werden.")

        # Dipiga KGT
        if name == "Dipiga" and day == 1 and slot == 0:
            warnings.append(f"⚠️  {tag}: Dipiga hat KGT am Dienstag Vormittag! → KGT fällt aus oder muss vertreten werden.")
        if name == "Dipiga" and day == 3 and slot == 1:
            warnings.append(f"⚠️  {tag}: Dipiga hat KGT am Donnerstag Nachmittag! → KGT fällt aus oder muss vertreten werden.")

        # Martina KSC Spez.
        if name == "Martina" and day in [1, 3] and slot == 0:
            warnings.append(f"⚠️  {tag}: Martina hat KSC Spez. am {day_name} VM! → KSC Spez. fällt aus.")

        # Florence PO Di
        if name == "Florence" and day == 1:
            warnings.append(f"⚠️  {tag}: Florence hat normalerweise PO am Dienstag.")

        # Maria B. HO Mo
        if name == "Maria B." and day == 0 and slot == 0:
            warnings.append(f"⚠️  {tag}: Maria B. hat normalerweise HO am Montag VM.")

        # ERF5 Mi NM
        if name in ["Jesika", "Dipiga"] and day == 2 and slot == 1:
            warnings.append(f"⚠️  {tag}: {name} könnte ERF5 am Mi NM haben! → ERF5 evtl. nicht besetzt.")

    # Zähle verfügbare Personen pro Slot um TEL/ABKL Engpässe zu finden
    for day in range(5):
        for slot in range(2):
            absent_count = sum(1 for n, d, s in absences if d == day and s == slot)
            available_count = sum(1 for n in employees
                                 if employees[n].is_available(day, slot)) - absent_count
            target_tel = 4 if (day == 0 and slot == 0) else 3
            min_needed = target_tel + 2  # TEL + ABKL
            if available_count < min_needed:
                warnings.append(
                    f"🔴 {DAYS[day]} {SLOTS[slot]}: Nur noch {available_count} Personen verfügbar! "
                    f"Brauche mind. {min_needed} (TEL:{target_tel} + ABKL:2). "
                    f"REGELN KÖNNEN NICHT EINGEHALTEN WERDEN!"
                )
            elif available_count < min_needed + 3:
                warnings.append(
                    f"🟡 {DAYS[day]} {SLOTS[slot]}: Nur {available_count} Personen verfügbar. "
                    f"Eng, aber machbar."
                )

    return warnings


def interactive_input(week_number):
    """Interaktive Eingabe von Abwesenheiten und Sonderwünschen."""
    employees = create_employees(week_number)
    overrides = {}
    absences = []

    print("\n" + "─" * 60)
    print("  SONDERWÜNSCHE / ABWESENHEITEN EINGEBEN")
    print("─" * 60)
    print("  Hier kannst du Krankmeldungen, Termine, Ferien etc. eingeben.")
    print("  Am Ende prüft das Programm welche Regeln betroffen sind.")

    entry_count = 0

    while True:
        print("\n" + "─" * 40)
        if entry_count == 0:
            prompt = "  Gibt es eine Abwesenheit oder einen Sonderwunsch? (j/n): "
        else:
            prompt = "  Gibt es eine weitere Abwesenheit oder einen Sonderwunsch? (j/n): "
        ans = input(prompt).strip().lower()
        if ans not in ["j", "ja", "y", "yes"]:
            break

        # Mitarbeiterin wählen
        name = pick_employee()
        if not name:
            break

        # Typ wählen
        typ = pick_absence_type()

        if typ in [1, 2]:
            # Krank / Ferien → ganze Tage
            label = "KRANK" if typ == 1 else "FERIEN"
            days = pick_days()
            for day in days:
                for slot in [0, 1]:
                    if employees[name].is_available(day, slot):
                        overrides[(name, day, slot)] = label
                        absences.append((name, day, slot))
            day_str = ", ".join(DAYS[d] for d in days)
            print(f"\n  ✓ {name} → {label} am {day_str}")
            entry_count += 1

        elif typ == 3:
            # Arzttermin → bestimmter Halbtag
            days = pick_days()
            slots = pick_slots()
            notiz = input("  Notiz (z.B. 'Arzttermin 14:30', Enter=leer): ").strip()
            if not notiz:
                notiz = "Termin"
            for day in days:
                for slot in slots:
                    if employees[name].is_available(day, slot):
                        overrides[(name, day, slot)] = f"*{notiz}"
                        absences.append((name, day, slot))
            day_str = ", ".join(DAYS[d] for d in days)
            slot_str = "+".join(SLOTS[s] for s in slots)
            print(f"\n  ✓ {name} → {notiz} am {day_str} ({slot_str})")
            entry_count += 1

        elif typ == 4:
            # Sonderwunsch → bestimmte Aufgabe
            days = pick_days()
            slots = pick_slots()
            print(f"\n  Welche Aufgabe? (z.B. HO, TEL, ERF7, ABKL, etc.)")
            task = input("  Aufgabe: ").strip()
            if not task:
                task = "HO"
            for day in days:
                for slot in slots:
                    if employees[name].is_available(day, slot):
                        overrides[(name, day, slot)] = task
            day_str = ", ".join(DAYS[d] for d in days)
            slot_str = "+".join(SLOTS[s] for s in slots)
            print(f"\n  ✓ {name} → {task} am {day_str} ({slot_str})")
            entry_count += 1

    # Zusammenfassung
    if overrides:
        print("\n" + "═" * 60)
        print("  EINGEGEBENE SONDERWÜNSCHE:")
        print("═" * 60)
        for (name, day, slot), task in overrides.items():
            print(f"  • {name:15s}  {DAYS[day]:12s} {SLOTS[slot]:12s}  →  {task}")

        # Regel-Impact prüfen
        if absences:
            warnings = check_rule_impact(absences, employees)
            if warnings:
                print("\n" + "═" * 60)
                print("  AUSWIRKUNGEN AUF REGELN:")
                print("═" * 60)
                for w in warnings:
                    print(f"  {w}")
                print()
                ans = input("  Trotzdem weiterfahren? (j/n): ").strip().lower()
                if ans not in ["j", "ja", "y", "yes"]:
                    print("  Abgebrochen. Bitte Eingaben anpassen und erneut starten.")
                    return None
            else:
                print("\n  ✅ Keine kritischen Regel-Auswirkungen erkannt.")
    else:
        print("\n  Keine Sonderwünsche eingegeben → normaler Plan wird generiert.")

    return overrides


def validate_result(sched):
    """Prüft die wichtigsten Regeln und zeigt Warnungen."""
    issues = []
    warnings = []
    employees = sched.employees

    # TEL und ABKL Mindestbesetzung
    for day in range(5):
        for slot in range(2):
            target_tel = 4 if (day == 0 and slot == 0) else 3
            tel_c = sched.tel_count[day][slot]
            abkl_c = sched.abkl_count[day][slot]
            if tel_c < target_tel:
                issues.append(f"❌ {DAYS[day]} {SLOTS[slot]}: TEL nur {tel_c}/{target_tel}")
            if abkl_c < 2:
                issues.append(f"❌ {DAYS[day]} {SLOTS[slot]}: ABKL nur {abkl_c}/2")

    # ERF8 pro Tag
    for day in range(5):
        has_erf8 = any("ERF8" in sched.get_task(n, day, s)
                       for n in employees for s in range(2))
        if not has_erf8:
            issues.append(f"❌ {DAYS[day]}: Kein ERF8 besetzt!")

    # ONB-Ausschluss prüfen (via Overrides zugewiesen)
    no_onb = [n for n, emp in employees.items() if not emp.can_onb]
    for name in no_onb:
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "ONB" in t:
                    issues.append(f"❌ {name}: ONB zugewiesen an {DAYS[day]} {SLOTS[slot]}, aber nicht erlaubt!")

    # Queue nur bei TL prüfen
    for name, emp in employees.items():
        if emp.is_tl:
            continue
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "/Q" in t and "PO" not in t and "SCAN" not in t:
                    warnings.append(f"⚠️  {name}: hat Queue-Aufgabe ({t}) an {DAYS[day]} {SLOTS[slot]}, ist aber kein TL")

    # TL-TEL Soft-Regel: Jede TL soll 1-2x pro Woche TEL haben
    for name, emp in employees.items():
        if not emp.is_tl:
            continue
        tel_count = sum(1 for day in range(5) for slot in range(2)
                        if "TEL" in sched.get_task(name, day, slot))
        if tel_count == 0:
            warnings.append(f"⚠️  TL {name}: hat 0x TEL diese Woche (Soll: 1-2x)")
        elif tel_count > 2:
            warnings.append(f"⚠️  TL {name}: hat {tel_count}x TEL diese Woche (Soll: 1-2x)")

    return issues + warnings


def main():
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║         ARBEITSKALENDER-GENERATOR                      ║")
    print("║         KSC Team - Wöchentlicher Arbeitsplan            ║")
    print("╚══════════════════════════════════════════════════════════╝")

    # Datum bestimmen
    next_monday = get_next_monday()
    week_number = next_monday.isocalendar()[1]
    end_friday = next_monday + timedelta(days=4)

    print(f"\n  Woche:  KW {week_number}")
    print(f"  Von:    {next_monday.strftime('%d.%m.%Y')} (Montag)")
    print(f"  Bis:    {end_friday.strftime('%d.%m.%Y')} (Freitag)")

    # Interaktive Eingabe
    overrides = interactive_input(week_number)
    if overrides is None:
        return

    # Schedule generieren
    print("\n  Generiere Arbeitsplan...")
    random.seed(week_number)

    sched = build_schedule(week_number, next_monday, overrides)

    # Regel-Validierung nach Generierung
    issues = validate_result(sched)

    # Excel schreiben
    filename = f"Arbeitsplan_KW{week_number}_{next_monday.strftime('%Y%m%d')}.xlsx"
    output_path = write_excel(sched, filename)

    # Ergebnis
    print(f"\n{'═' * 60}")
    print(f"  ERGEBNIS")
    print(f"{'═' * 60}")

    if issues:
        print(f"\n  ⚠️  ACHTUNG - Folgende Regeln konnten nicht eingehalten werden:")
        for issue in issues:
            print(f"    {issue}")
        print(f"\n  Grund: Vermutlich zu wenige Personen verfügbar.")
        print(f"  → Prüfe ob zusätzliche Abwesenheiten angepasst werden können.")
    else:
        print(f"\n  ✅ Alle Regeln eingehalten!")

    print(f"\n  📄 Datei erstellt: {filename}")

    # Zusammenfassung
    print(f"\n{'─' * 60}")
    print(f"  BESETZUNG:")
    print(f"{'─' * 60}")
    for day in range(5):
        for slot in range(2):
            tel_c = sched.tel_count[day][slot]
            abkl_c = sched.abkl_count[day][slot]
            target_tel = 4 if (day == 0 and slot == 0) else 3
            tel_sym = "✅" if tel_c >= target_tel else "❌"
            abkl_sym = "✅" if abkl_c >= 2 else "❌"
            print(f"  {DAYS[day]:12s} {SLOTS[slot]:12s}  "
                  f"TEL: {tel_c}/{target_tel} {tel_sym}  "
                  f"ABKL: {abkl_c}/2 {abkl_sym}")

    print(f"\n{'═' * 60}")
    print(f"  Fertig!")
    print(f"{'═' * 60}\n")

    return filename


if __name__ == "__main__":
    main()
