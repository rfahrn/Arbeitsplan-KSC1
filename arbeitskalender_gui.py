#!/usr/bin/env python3
"""
Arbeitskalender-Generator — GUI v5.1
Grafische Oberfläche für den wöchentlichen Arbeitsplan.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys
import json
import random
import subprocess
import platform
from datetime import datetime, timedelta

# Import aus dem Hauptmodul
from arbeitskalender import (
    build_schedule, create_employees, write_excel, validate_schedule,
    DAYS, SLOTS, get_next_monday, load_state,
    ONB_ERLAUBT, DIREKT_ERLAUBT, BTM_ERLAUBT
)

# ============================================================
# FARBEN & STYLES (Modernes Dunkles Design)
# ============================================================

BG           = "#1e1e2e"
BG_CARD      = "#2a2a3d"
BG_INPUT     = "#363650"
BG_LIST      = "#2f2f45"
FG           = "#e0e0e0"
FG_DIM       = "#8888aa"
FG_HEAD      = "#ffffff"
ACCENT       = "#7c6ff7"
ACCENT_HOVER = "#9580ff"
GREEN        = "#50c878"
YELLOW       = "#ffc857"
RED          = "#ff6b6b"
ORANGE       = "#f0a050"
BORDER       = "#3d3d5c"
BLUE         = "#5dabf0"

FONT         = ("Segoe UI", 10)
FONT_BOLD    = ("Segoe UI", 10, "bold")
FONT_HEAD    = ("Segoe UI", 16, "bold")
FONT_SUB     = ("Segoe UI", 11, "bold")
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)

# Mitarbeiter-Liste
EMPLOYEE_LIST = [
    "Silvana", "Linda", "Lara", "Andrea G.", "Dipiga",
    "Isaura", "Martina", "Alessia", "Amra", "Nina",
    "Jesika", "Brigitte", "Florence", "Corinne", "Saskia",
    "Dragi", "Maria B.", "Andrea A."
]

ABSENCE_TYPES = {
    "Krank":                "KRANK",
    "Ferien / Abwesend":    "FERIEN",
    "Arzttermin / Termin":  "TERMIN",
    "Home Office":          "HO",
    "Anderer Wunsch":       "CUSTOM",
}

DAY_OPTIONS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
SLOT_OPTIONS = ["Ganzer Tag", "Vormittag", "Nachmittag"]


# ============================================================
# REGEL-IMPACT PRÜFUNG
# ============================================================

def check_rule_impact(absences, employees):
    """Prüft welche Regeln durch Abwesenheiten verletzt werden könnten."""
    warnings = []
    
    for name, day, slot in absences:
        day_name = DAYS[day]
        slot_name = SLOTS[slot]
        tag = f"{name} {day_name} {slot_name}"
        
        if name == "Brigitte" and day == 0:
            warnings.append(f"⚠️  {tag}: Brigitte hat normalerweise ERF9 am Montag!")
        if name == "Dragi" and day == 1:
            warnings.append(f"⚠️  {tag}: Dragi hat normalerweise ERF9 am Dienstag!")
        if name == "Dipiga" and day == 1 and slot == 0:
            warnings.append(f"⚠️  {tag}: Dipiga hat TAGES PA am Dienstag VM!")
        if name == "Dipiga" and day == 3 and slot == 1:
            warnings.append(f"⚠️  {tag}: Dipiga hat KGS am Donnerstag NM!")
        if name == "Martina" and day in [1, 3] and slot == 0:
            warnings.append(f"⚠️  {tag}: Martina hat KSC Spez. am {day_name} VM!")
        if name == "Florence" and day == 1:
            warnings.append(f"⚠️  {tag}: Florence hat normalerweise PO am Dienstag!")
        if name == "Maria B." and day == 0 and slot == 0:
            warnings.append(f"⚠️  {tag}: Maria B. hat normalerweise HO am Montag VM!")
        if name in ["Jesika", "Dipiga"] and day == 2 and slot == 1:
            warnings.append(f"⚠️  {tag}: {name} könnte ERF5 (Labor) am Mi NM haben!")
    
    # Prüfe Mindestbesetzung
    for day in range(5):
        for slot in range(2):
            absent_count = sum(1 for n, d, s in absences if d == day and s == slot)
            available_count = sum(1 for n in employees 
                                 if employees[n].is_available(day, slot)) - absent_count
            target_tel = 4 if (day == 0 and slot == 0) else 3
            min_needed = target_tel + 2
            if available_count < min_needed:
                warnings.append(
                    f"🔴 {DAYS[day]} {SLOTS[slot]}: Nur {available_count} Personen! "
                    f"Brauche mind. {min_needed}!"
                )
            elif available_count < min_needed + 3:
                warnings.append(
                    f"🟡 {DAYS[day]} {SLOTS[slot]}: Nur {available_count} Personen - eng!"
                )
    
    return warnings


# ============================================================
# GUI APP
# ============================================================

class ArbeitskalenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KSC Arbeitsplan-Generator")
        self.root.configure(bg=BG)
        self.root.geometry("900x850")
        self.root.minsize(800, 700)

        # State
        self.entries = []
        self.next_monday = get_next_monday()
        self.week_number = self.next_monday.isocalendar()[1]
        self.end_friday = self.next_monday + timedelta(days=4)
        self.is_even_week = (self.week_number % 2 == 0)

        self._build_ui()
        self._apply_styles()

    # ────────────────────────────────────────
    # UI AUFBAU
    # ────────────────────────────────────────

    def _build_ui(self):
        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=20, pady=12)

        # ── HEADER ──
        hdr = tk.Frame(container, bg=BG)
        hdr.pack(fill="x", pady=(0, 12))

        tk.Label(hdr, text="📅  KSC Arbeitsplan-Generator",
                 font=FONT_HEAD, fg=FG_HEAD, bg=BG).pack(side="left")

        # Wochen-Info rechts
        week_frame = tk.Frame(hdr, bg=BG)
        week_frame.pack(side="right")
        
        week_text = f"KW {self.week_number}"
        tk.Label(week_frame, text=week_text, font=FONT_HEAD, fg=ACCENT, bg=BG).pack(side="left", padx=(0, 8))
        
        date_text = f"{self.next_monday.strftime('%d.%m.%Y')} – {self.end_friday.strftime('%d.%m.%Y')}"
        tk.Label(week_frame, text=date_text, font=FONT, fg=FG_DIM, bg=BG).pack(side="left")

        self._sep(container)

        # ── 2-WOCHEN-INFO ──
        info_frame = tk.Frame(container, bg=BG_CARD, padx=12, pady=8)
        info_frame.pack(fill="x", pady=(0, 10))
        
        week_type = "gerade" if self.is_even_week else "ungerade"
        tk.Label(info_frame, text=f"ℹ️  KW {self.week_number} ist {week_type}:", 
                 font=FONT_BOLD, fg=BLUE, bg=BG_CARD).pack(side="left")
        
        if self.is_even_week:
            info_text = "Linda Fr frei • Isaura Mi frei • Corinne Fr frei"
        else:
            info_text = "Linda, Isaura, Corinne normal verfügbar"
        tk.Label(info_frame, text=info_text, font=FONT, fg=FG_DIM, bg=BG_CARD).pack(side="left", padx=12)

        # ── EINGABE-KARTE ──
        card = self._card(container, "➕  Abwesenheit / Sonderwunsch hinzufügen")

        # Zeile 1: Mitarbeiterin + Typ
        row1 = tk.Frame(card, bg=BG_CARD)
        row1.pack(fill="x", pady=(0, 10))

        self._label(row1, "Mitarbeiterin:", side="left")
        self.emp_var = tk.StringVar(value=EMPLOYEE_LIST[0])
        self.emp_combo = ttk.Combobox(row1, textvariable=self.emp_var,
                                       values=EMPLOYEE_LIST, width=18,
                                       state="readonly", font=FONT)
        self.emp_combo.pack(side="left", padx=(4, 20))

        self._label(row1, "Grund:", side="left")
        self.type_var = tk.StringVar(value="Krank")
        self.type_combo = ttk.Combobox(row1, textvariable=self.type_var,
                                        values=list(ABSENCE_TYPES.keys()),
                                        width=20, state="readonly", font=FONT)
        self.type_combo.pack(side="left", padx=4)
        self.type_combo.bind("<<ComboboxSelected>>", self._on_type_change)

        # Zeile 2: Tage
        row2 = tk.Frame(card, bg=BG_CARD)
        row2.pack(fill="x", pady=(0, 10))

        self._label(row2, "Tage:", side="left")

        self.day_frame = tk.Frame(row2, bg=BG_CARD)
        self.day_frame.pack(side="left", padx=(4, 20))
        self.day_vars = {}
        for d in DAY_OPTIONS:
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(self.day_frame, text=d[:2], variable=var,
                                bg=BG_CARD, fg=FG, selectcolor=BG_INPUT,
                                activebackground=BG_CARD, activeforeground=FG,
                                font=FONT_BOLD, indicatoron=True, pady=2)
            cb.pack(side="left", padx=4)
            self.day_vars[d] = var

        # Schnellauswahl-Buttons
        tk.Button(row2, text="Alle", font=FONT_SMALL, bg=BG_INPUT, fg=FG,
                  relief="flat", padx=6, command=self._select_all_days).pack(side="left", padx=2)
        tk.Button(row2, text="Keine", font=FONT_SMALL, bg=BG_INPUT, fg=FG,
                  relief="flat", padx=6, command=self._deselect_all_days).pack(side="left", padx=2)

        # Zeile 3: Halbtag + Notiz
        row3 = tk.Frame(card, bg=BG_CARD)
        row3.pack(fill="x", pady=(0, 10))

        self._label(row3, "Halbtag:", side="left")
        self.slot_var = tk.StringVar(value="Ganzer Tag")
        self.slot_combo = ttk.Combobox(row3, textvariable=self.slot_var,
                                        values=SLOT_OPTIONS, width=12,
                                        state="readonly", font=FONT)
        self.slot_combo.pack(side="left", padx=(4, 20))

        self._label(row3, "Notiz:", side="left")
        self.note_var = tk.StringVar()
        self.note_entry = tk.Entry(row3, textvariable=self.note_var, width=25,
                                    font=FONT, bg=BG_INPUT, fg=FG,
                                    insertbackground=FG, relief="flat",
                                    highlightthickness=1, highlightcolor=ACCENT,
                                    highlightbackground=BORDER)
        self.note_entry.pack(side="left", padx=(4, 20), ipady=4)

        self.add_btn = tk.Button(row3, text="＋  Hinzufügen", font=FONT_BOLD,
                                  bg=ACCENT, fg="#fff", relief="flat",
                                  activebackground=ACCENT_HOVER,
                                  cursor="hand2", padx=18, pady=6,
                                  command=self._add_entry)
        self.add_btn.pack(side="right")

        self._on_type_change()

        # ── LISTE ──
        list_card = self._card(container, "📋  Eingetragene Abwesenheiten / Wünsche")

        # Scrollbarer Bereich für die Liste
        list_container = tk.Frame(list_card, bg=BG_CARD)
        list_container.pack(fill="both", expand=True)
        
        self.list_canvas = tk.Canvas(list_container, bg=BG_CARD, highlightthickness=0, height=120)
        self.list_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.list_canvas.yview)
        self.list_frame = tk.Frame(self.list_canvas, bg=BG_CARD)
        
        self.list_frame.bind("<Configure>", lambda e: self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.list_canvas.configure(yscrollcommand=self.list_scrollbar.set)
        
        self.list_canvas.pack(side="left", fill="both", expand=True)
        self.list_scrollbar.pack(side="right", fill="y")

        self._refresh_list()

        # ── GENERATE BUTTON ──
        btn_frame = tk.Frame(container, bg=BG)
        btn_frame.pack(fill="x", pady=(14, 8))

        self.gen_btn = tk.Button(btn_frame,
                                  text="📄  Arbeitsplan generieren",
                                  font=("Segoe UI", 13, "bold"),
                                  bg=GREEN, fg="#000", relief="flat",
                                  activebackground="#3ddb6e",
                                  cursor="hand2", padx=28, pady=10,
                                  command=self._generate)
        self.gen_btn.pack(side="left")

        self.status_label = tk.Label(btn_frame, text="", font=FONT,
                                      fg=FG_DIM, bg=BG)
        self.status_label.pack(side="left", padx=20)

        # ── ERGEBNIS ──
        result_card = self._card(container, "📊  Ergebnis")
        self.result_text = scrolledtext.ScrolledText(
            result_card, height=14, font=FONT_MONO,
            bg=BG_LIST, fg=FG, relief="flat", wrap="word",
            insertbackground=FG, padx=10, pady=8
        )
        self.result_text.pack(fill="both", expand=True)
        self._show_initial_result()

    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
                         fieldbackground=BG_INPUT, background=BG_INPUT,
                         foreground=FG, arrowcolor=FG,
                         selectbackground=ACCENT, selectforeground="#fff",
                         borderwidth=0)
        style.map("TCombobox", fieldbackground=[("readonly", BG_INPUT)])
        style.configure("Vertical.TScrollbar", background=BG_INPUT, 
                        troughcolor=BG_CARD, borderwidth=0)

    def _show_initial_result(self):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("end", "Noch kein Plan generiert.\n\n")
        self.result_text.insert("end", "1. Trage Abwesenheiten ein (falls nötig)\n")
        self.result_text.insert("end", "2. Klicke auf «Arbeitsplan generieren»\n")
        self.result_text.insert("end", "3. Die Excel-Datei wird automatisch erstellt\n")
        self.result_text.config(state="disabled")

    # ────────────────────────────────────────
    # HILFSFUNKTIONEN
    # ────────────────────────────────────────

    def _card(self, parent, title):
        outer = tk.Frame(parent, bg=BORDER, bd=0)
        outer.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(outer, bg=BG_CARD, padx=16, pady=12)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, font=FONT_SUB, fg=ACCENT, bg=BG_CARD).pack(anchor="w", pady=(0, 10))
        return inner

    def _label(self, parent, text, side="left"):
        tk.Label(parent, text=text, font=FONT, fg=FG_DIM, bg=BG_CARD).pack(side=side, padx=(0, 4))

    def _sep(self, parent):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=6)

    def _select_all_days(self):
        for v in self.day_vars.values():
            v.set(True)

    def _deselect_all_days(self):
        for v in self.day_vars.values():
            v.set(False)

    def _on_type_change(self, event=None):
        typ = self.type_var.get()
        # Halbtag ist IMMER wählbar
        self.slot_combo.config(state="readonly")
        
        # Notiz ist IMMER verfügbar (für zusätzliche Infos)
        self.note_entry.config(state="normal")
        
        # Nur Default-Werte setzen, aber nicht sperren
        if typ in ["Krank", "Ferien / Abwesend"]:
            self.slot_var.set("Ganzer Tag")  # Default, aber änderbar
            self.note_var.set("")
        elif typ == "Arzttermin / Termin":
            self.slot_var.set("Nachmittag")  # Typischerweise NM
            self.note_var.set("")
        elif typ == "Home Office":
            self.slot_var.set("Ganzer Tag")
            self.note_var.set("")
        else:  # Anderer Wunsch
            self.slot_var.set("Ganzer Tag")
            self.note_var.set("")

    # ────────────────────────────────────────
    # EINTRAG HINZUFÜGEN / ENTFERNEN
    # ────────────────────────────────────────

    def _add_entry(self):
        name = self.emp_var.get()
        typ = self.type_var.get()
        slot_choice = self.slot_var.get()
        note = self.note_var.get().strip()

        selected_days = [d for d in DAY_OPTIONS if self.day_vars[d].get()]
        if not selected_days:
            messagebox.showwarning("Keine Tage", "Bitte mindestens einen Tag auswählen.")
            return

        if slot_choice == "Ganzer Tag":
            slots = [0, 1]
            slot_display = "ganzer Tag"
        elif slot_choice == "Vormittag":
            slots = [0]
            slot_display = "VM"
        else:
            slots = [1]
            slot_display = "NM"

        task_code = ABSENCE_TYPES[typ]
        
        # Task-Label je nach Typ
        if task_code == "TERMIN":
            task_label = f"*{note}" if note else "*Termin"
        elif task_code == "CUSTOM":
            task_label = note if note else "HO"
        elif task_code in ["KRANK", "FERIEN", "HO"]:
            # Bei Krank/Ferien/HO: Task bleibt, Notiz wird separat angezeigt
            task_label = task_code
        else:
            task_label = task_code

        day_short = ", ".join(d[:2] for d in selected_days)

        # Display: Notiz wird IMMER angezeigt wenn vorhanden
        display_text = f"{name}  ·  {typ}  ·  {day_short} ({slot_display})"
        if note:
            display_text += f"  ·  «{note}»"

        entry = {
            "name": name,
            "type_label": typ,
            "days": [DAY_OPTIONS.index(d) for d in selected_days],
            "slots": slots,
            "task": task_label,
            "note": note,  # Notiz separat speichern
            "display": display_text,
        }
        self.entries.append(entry)

        # Reset
        self._deselect_all_days()
        self.note_var.set("")

        self._refresh_list()

    def _remove_entry(self, idx):
        if 0 <= idx < len(self.entries):
            self.entries.pop(idx)
            self._refresh_list()

    def _refresh_list(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        if not self.entries:
            empty = tk.Label(self.list_frame,
                             text="Noch keine Einträge — alles normal diese Woche? 😊",
                             font=FONT, fg=FG_DIM, bg=BG_CARD, pady=14)
            empty.pack()
            return

        for i, entry in enumerate(self.entries):
            row = tk.Frame(self.list_frame, bg=BG_LIST, pady=6, padx=10)
            row.pack(fill="x", pady=3, padx=4)

            typ = entry["type_label"]
            color = RED if "Krank" in typ else (ORANGE if "Ferien" in typ
                    else YELLOW if "Termin" in typ else (BLUE if "Home" in typ else FG_DIM))

            tk.Label(row, text="●", font=FONT, fg=color, bg=BG_LIST).pack(side="left", padx=(0, 8))
            tk.Label(row, text=entry["display"], font=FONT, fg=FG, bg=BG_LIST).pack(side="left")

            rm_btn = tk.Button(row, text="✕", font=FONT_SMALL,
                                bg=BG_LIST, fg=RED, relief="flat",
                                activebackground=BG_INPUT,
                                cursor="hand2", padx=6,
                                command=lambda idx=i: self._remove_entry(idx))
            rm_btn.pack(side="right")

        count = len(self.entries)
        suffix = "Eintrag" if count == 1 else "Einträge"
        tk.Label(self.list_frame,
                 text=f"  {count} {suffix}  ·  Weitere hinzufügen oder → Generieren",
                 font=FONT_SMALL, fg=FG_DIM, bg=BG_CARD, pady=6).pack(anchor="w")

    # ────────────────────────────────────────
    # PLAN GENERIEREN
    # ────────────────────────────────────────

    def _generate(self):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")

        employees = create_employees(self.week_number)

        # Build overrides und separate Notizen
        overrides = {}
        extra_notes = {}  # Für Notizen bei KRANK/FERIEN/HO
        absences = []
        
        for entry in self.entries:
            name = entry["name"]
            note = entry.get("note", "")
            
            for day in entry["days"]:
                for slot in entry["slots"]:
                    if name in employees and employees[name].is_available(day, slot):
                        overrides[(name, day, slot)] = entry["task"]
                        
                        # Notiz separat speichern wenn vorhanden
                        if note and entry["task"] in ["KRANK", "FERIEN", "HO"]:
                            extra_notes[(name, day, slot)] = note
                        
                        if entry["task"] in ["KRANK", "FERIEN"] or entry["task"].startswith("*"):
                            absences.append((name, day, slot))

        # Pre-check
        if absences:
            warnings = check_rule_impact(absences, employees)
            if warnings:
                self._log("⚠️  AUSWIRKUNGEN AUF REGELN:\n")
                self._log("─" * 50 + "\n")
                for w in warnings:
                    self._log(f"  {w}\n")
                self._log("\n")

        self._log("Generiere Arbeitsplan für KW " + str(self.week_number) + "...\n\n")
        random.seed(self.week_number)

        sched = build_schedule(self.week_number, self.next_monday, overrides)
        
        # Extra-Notizen hinzufügen (für KRANK/FERIEN/HO mit Bemerkung)
        for key, note in extra_notes.items():
            sched.notes[key] = note
        
        issues = validate_schedule(sched)

        filename = f"Arbeitsplan_KW{self.week_number}_{self.next_monday.strftime('%Y%m%d')}.xlsx"
        write_excel(sched, filename)

        # Result
        self._log("═" * 50 + "\n")
        if issues:
            self._log("⚠️  ACHTUNG — Einige Regeln nicht vollständig einhaltbar:\n\n")
            for issue in issues:
                self._log(f"  {issue}\n")
            self._log("\n")
        else:
            self._log("✅  Alle Regeln eingehalten!\n\n")

        self._log(f"📄  Datei erstellt: {filename}\n\n")
        self._log("═" * 50 + "\n")
        self._log("BESETZUNG:\n\n")

        for day in range(5):
            for slot in range(2):
                tel_c = sched.tel_count[day][slot]
                abkl_c = sched.abkl_count[day][slot]
                target_tel = 4 if (day == 0 and slot == 0) else 3
                tel_sym = "✅" if tel_c >= target_tel else "❌"
                abkl_sym = "✅" if abkl_c >= 2 else "❌"
                self._log(f"  {DAYS[day]:12s} {SLOTS[slot]:12s}  "
                          f"TEL: {tel_c}/{target_tel} {tel_sym}  "
                          f"ABKL: {abkl_c}/2 {abkl_sym}\n")

        self._log("\n" + "─" * 50 + "\n")
        self._log("WOCHENAUFGABEN:\n\n")
        self._log(f"  Direkt:  {sched.wochenaufgabe_direkt}\n")
        self._log(f"  ONB:     {sched.wochenaufgabe_onb}\n")
        self._log(f"  BTM:     {sched.wochenaufgabe_btm}\n")

        self._log("\n" + "─" * 50 + "\n")
        self._log("TAGESVERANTWORTUNG:\n\n")
        for d in range(5):
            self._log(f"  {DAYS[d]:12s}  {sched.tagesverantwortung.get(d, '?')}\n")

        self.result_text.config(state="disabled")

        # Status
        self.status_label.config(
            text=f"✅ {filename} erstellt!" if not issues
            else f"⚠️ {filename} — mit Einschränkungen",
            fg=GREEN if not issues else YELLOW
        )

        # Fragen ob öffnen
        if messagebox.askyesno("Fertig!",
                                f"Arbeitsplan KW {self.week_number} wurde erstellt.\n\n"
                                f"Datei: {filename}\n\n"
                                f"Excel jetzt öffnen?"):
            self._open_file(filename)

    def _log(self, text):
        self.result_text.insert("end", text)

    def _open_file(self, path):
        abs_path = os.path.abspath(path)
        try:
            if platform.system() == "Windows":
                os.startfile(abs_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", abs_path])
            else:
                subprocess.Popen(["xdg-open", abs_path])
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht öffnen:\n{e}")


# ============================================================
# START
# ============================================================

def main():
    root = tk.Tk()
    
    # Icon setzen (falls vorhanden)
    try:
        if platform.system() == "Windows":
            root.iconbitmap(default='')
    except:
        pass
    
    app = ArbeitskalenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
