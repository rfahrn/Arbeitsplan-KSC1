#!/usr/bin/env python3
"""
Arbeitskalender-Generator — GUI
Grafische Oberfläche für den wöchentlichen Arbeitsplan.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os, sys, json, random, subprocess, platform
from datetime import datetime, timedelta

from arbeitskalender import (
    build_schedule, create_employees, write_excel,
    DAYS, SLOTS, EMPLOYEE_NAMES, check_rule_impact, validate_result,
    get_next_monday
)

# ============================================================
# FARBEN & STYLES
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

FONT         = ("Segoe UI", 10)
FONT_BOLD    = ("Segoe UI", 10, "bold")
FONT_HEAD    = ("Segoe UI", 14, "bold")
FONT_SUB     = ("Segoe UI", 11, "bold")
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)

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
# GUI APP
# ============================================================

class ArbeitskalenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Arbeitskalender-Generator — KSC Team")
        self.root.configure(bg=BG)
        self.root.geometry("820x780")
        self.root.minsize(700, 650)

        # State
        self.entries = []  # list of dicts
        self.next_monday = get_next_monday()
        self.week_number = self.next_monday.isocalendar()[1]
        self.end_friday = self.next_monday + timedelta(days=4)

        self._build_ui()

    # ────────────────────────────────────────
    # UI AUFBAU
    # ────────────────────────────────────────

    def _build_ui(self):
        # Scrollbare Hauptfläche
        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=16, pady=10)

        # ── HEADER ──
        hdr = tk.Frame(container, bg=BG)
        hdr.pack(fill="x", pady=(0, 10))

        tk.Label(hdr, text="📅  Arbeitskalender-Generator",
                 font=FONT_HEAD, fg=FG_HEAD, bg=BG).pack(side="left")

        week_text = (f"KW {self.week_number}  ·  "
                     f"{self.next_monday.strftime('%d.%m.%Y')} – "
                     f"{self.end_friday.strftime('%d.%m.%Y')}")
        tk.Label(hdr, text=week_text, font=FONT, fg=ACCENT, bg=BG).pack(side="right")

        self._sep(container)

        # ── EINGABE-KARTE ──
        card = self._card(container, "Abwesenheit / Sonderwunsch hinzufügen")

        # Zeile 1: Mitarbeiterin + Typ
        row1 = tk.Frame(card, bg=BG_CARD)
        row1.pack(fill="x", pady=(0, 8))

        self._label(row1, "Mitarbeiterin:", side="left")
        self.emp_var = tk.StringVar(value=EMPLOYEE_NAMES[0])
        self.emp_combo = ttk.Combobox(row1, textvariable=self.emp_var,
                                       values=EMPLOYEE_NAMES, width=18,
                                       state="readonly", font=FONT)
        self.emp_combo.pack(side="left", padx=(4, 16))

        self._label(row1, "Grund:", side="left")
        self.type_var = tk.StringVar(value="Krank")
        self.type_combo = ttk.Combobox(row1, textvariable=self.type_var,
                                        values=list(ABSENCE_TYPES.keys()),
                                        width=20, state="readonly", font=FONT)
        self.type_combo.pack(side="left", padx=4)
        self.type_combo.bind("<<ComboboxSelected>>", self._on_type_change)

        # Zeile 2: Tage + Halbtag
        row2 = tk.Frame(card, bg=BG_CARD)
        row2.pack(fill="x", pady=(0, 8))

        self._label(row2, "Tage:", side="left")

        self.day_frame = tk.Frame(row2, bg=BG_CARD)
        self.day_frame.pack(side="left", padx=(4, 16))
        self.day_vars = {}
        for d in DAY_OPTIONS:
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(self.day_frame, text=d[:2], variable=var,
                                bg=BG_CARD, fg=FG, selectcolor=BG_INPUT,
                                activebackground=BG_CARD, activeforeground=FG,
                                font=FONT_BOLD, indicatoron=True)
            cb.pack(side="left", padx=2)
            self.day_vars[d] = var

        self._label(row2, "Halbtag:", side="left")
        self.slot_var = tk.StringVar(value="Ganzer Tag")
        self.slot_combo = ttk.Combobox(row2, textvariable=self.slot_var,
                                        values=SLOT_OPTIONS, width=12,
                                        state="readonly", font=FONT)
        self.slot_combo.pack(side="left", padx=4)

        # Zeile 3: Notiz + Button
        row3 = tk.Frame(card, bg=BG_CARD)
        row3.pack(fill="x", pady=(0, 4))

        self._label(row3, "Notiz:", side="left")
        self.note_var = tk.StringVar()
        self.note_entry = tk.Entry(row3, textvariable=self.note_var, width=30,
                                    font=FONT, bg=BG_INPUT, fg=FG,
                                    insertbackground=FG, relief="flat",
                                    highlightthickness=1, highlightcolor=ACCENT)
        self.note_entry.pack(side="left", padx=(4, 16), ipady=3)

        self.add_btn = tk.Button(row3, text="＋  Hinzufügen", font=FONT_BOLD,
                                  bg=ACCENT, fg="#fff", relief="flat",
                                  activebackground=ACCENT_HOVER,
                                  cursor="hand2", padx=16, pady=4,
                                  command=self._add_entry)
        self.add_btn.pack(side="right")

        self._on_type_change()

        # ── LISTE ──
        list_card = self._card(container, "Eingetragene Abwesenheiten / Wünsche")

        self.list_frame = tk.Frame(list_card, bg=BG_CARD)
        self.list_frame.pack(fill="both", expand=True)

        self.empty_label = tk.Label(self.list_frame,
                                     text="Noch keine Einträge — alles normal diese Woche? 😊",
                                     font=FONT, fg=FG_DIM, bg=BG_CARD, pady=12)
        self.empty_label.pack()

        # ── GENERATE BUTTON ──
        btn_frame = tk.Frame(container, bg=BG)
        btn_frame.pack(fill="x", pady=(12, 6))

        self.gen_btn = tk.Button(btn_frame,
                                  text="📄  Arbeitsplan generieren",
                                  font=("Segoe UI", 12, "bold"),
                                  bg=GREEN, fg="#000", relief="flat",
                                  activebackground="#3ddb6e",
                                  cursor="hand2", padx=24, pady=8,
                                  command=self._generate)
        self.gen_btn.pack(side="left")

        self.status_label = tk.Label(btn_frame, text="", font=FONT,
                                      fg=FG_DIM, bg=BG)
        self.status_label.pack(side="left", padx=16)

        # ── ERGEBNIS ──
        self.result_card = self._card(container, "Ergebnis")
        self.result_text = scrolledtext.ScrolledText(
            self.result_card, height=10, font=FONT_MONO,
            bg=BG_LIST, fg=FG, relief="flat", wrap="word",
            insertbackground=FG
        )
        self.result_text.pack(fill="both", expand=True)
        self.result_text.insert("end", "Noch kein Plan generiert.\n"
                                       "Trage Abwesenheiten ein (falls nötig) und klicke auf «Generieren».")
        self.result_text.config(state="disabled")

        # Style für ttk Combobox
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
                         fieldbackground=BG_INPUT, background=BG_INPUT,
                         foreground=FG, arrowcolor=FG,
                         selectbackground=ACCENT, selectforeground="#fff")

    # ────────────────────────────────────────
    # HILFSFUNKTIONEN FÜR UI
    # ────────────────────────────────────────

    def _card(self, parent, title):
        outer = tk.Frame(parent, bg=BORDER, bd=0)
        outer.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(outer, bg=BG_CARD, padx=14, pady=10)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, font=FONT_SUB, fg=ACCENT, bg=BG_CARD).pack(anchor="w", pady=(0, 8))
        return inner

    def _label(self, parent, text, side="left"):
        tk.Label(parent, text=text, font=FONT, fg=FG_DIM, bg=BG_CARD).pack(side=side, padx=(0, 2))

    def _sep(self, parent):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=4)

    def _on_type_change(self, event=None):
        typ = self.type_var.get()
        if typ in ["Krank", "Ferien / Abwesend"]:
            self.slot_var.set("Ganzer Tag")
            self.slot_combo.config(state="disabled")
            self.note_entry.config(state="disabled")
            self.note_var.set("")
        elif typ == "Arzttermin / Termin":
            self.slot_combo.config(state="readonly")
            self.note_entry.config(state="normal")
            self.note_var.set("")
        elif typ == "Home Office":
            self.slot_var.set("Ganzer Tag")
            self.slot_combo.config(state="readonly")
            self.note_entry.config(state="disabled")
            self.note_var.set("")
        else:
            self.slot_combo.config(state="readonly")
            self.note_entry.config(state="normal")
            self.note_var.set("")

    # ────────────────────────────────────────
    # EINTRAG HINZUFÜGEN
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

        # Build slots
        if slot_choice == "Ganzer Tag":
            slots = [0, 1]
            slot_display = "ganzer Tag"
        elif slot_choice == "Vormittag":
            slots = [0]
            slot_display = "VM"
        else:
            slots = [1]
            slot_display = "NM"

        # Build task label
        task_code = ABSENCE_TYPES[typ]
        if task_code == "TERMIN":
            task_label = f"*{note}" if note else "*Termin"
        elif task_code == "CUSTOM":
            task_label = note if note else "HO"
        else:
            task_label = task_code

        day_short = ", ".join(d[:2] for d in selected_days)

        entry = {
            "name": name,
            "type_label": typ,
            "days": [DAY_OPTIONS.index(d) for d in selected_days],
            "slots": slots,
            "task": task_label,
            "display": f"{name}  ·  {typ}  ·  {day_short} ({slot_display})"
                       + (f"  ·  {note}" if note and task_code not in ["KRANK", "FERIEN", "HO"] else ""),
        }
        self.entries.append(entry)

        # Reset Inputs
        for v in self.day_vars.values():
            v.set(False)
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
            self.empty_label = tk.Label(self.list_frame,
                                         text="Noch keine Einträge — alles normal diese Woche? 😊",
                                         font=FONT, fg=FG_DIM, bg=BG_CARD, pady=12)
            self.empty_label.pack()
            return

        for i, entry in enumerate(self.entries):
            row = tk.Frame(self.list_frame, bg=BG_LIST, pady=4, padx=8)
            row.pack(fill="x", pady=2)

            # Farbiger Punkt je nach Typ
            typ = entry["type_label"]
            color = RED if "Krank" in typ else (ORANGE if "Ferien" in typ
                    else YELLOW if "Termin" in typ else FG_DIM)

            tk.Label(row, text="●", font=FONT, fg=color, bg=BG_LIST).pack(side="left", padx=(0, 6))
            tk.Label(row, text=entry["display"], font=FONT, fg=FG, bg=BG_LIST).pack(side="left")

            rm_btn = tk.Button(row, text="✕", font=FONT_SMALL,
                                bg=BG_LIST, fg=RED, relief="flat",
                                activebackground=BG_INPUT,
                                cursor="hand2", padx=4,
                                command=lambda idx=i: self._remove_entry(idx))
            rm_btn.pack(side="right")

        count = len(self.entries)
        suffix = "Eintrag" if count == 1 else "Einträge"
        hint = tk.Label(self.list_frame,
                         text=f"  Gibt es eine weitere Abwesenheit oder einen Sonderwunsch?"
                              f"  Oben eintragen oder → Generieren",
                         font=FONT_SMALL, fg=FG_DIM, bg=BG_CARD, pady=4)
        hint.pack(anchor="w")

    # ────────────────────────────────────────
    # PLAN GENERIEREN
    # ────────────────────────────────────────

    def _generate(self):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")

        employees = create_employees(self.week_number)

        # Build overrides
        overrides = {}
        absences = []
        for entry in self.entries:
            name = entry["name"]
            for day in entry["days"]:
                for slot in entry["slots"]:
                    if name in employees and employees[name].is_available(day, slot):
                        overrides[(name, day, slot)] = entry["task"]
                        if entry["task"] in ["KRANK", "FERIEN"] or entry["task"].startswith("*"):
                            absences.append((name, day, slot))

        # Pre-check warnings
        if absences:
            warnings = check_rule_impact(absences, employees)
            if warnings:
                self._log("⚠️  AUSWIRKUNGEN AUF REGELN:\n")
                for w in warnings:
                    self._log(f"  {w}\n")
                self._log("\n")

        # Generate
        self._log("Generiere Arbeitsplan...\n\n")
        random.seed(self.week_number)

        sched = build_schedule(self.week_number, self.next_monday, overrides)

        # Post-check
        issues = validate_result(sched)

        filename = f"Arbeitsplan_KW{self.week_number}_{self.next_monday.strftime('%Y%m%d')}.xlsx"
        write_excel(sched, filename)

        # Result
        self._log("━" * 50 + "\n")
        if issues:
            self._log("⚠️  ACHTUNG — Regeln nicht vollständig einhaltbar:\n\n")
            for issue in issues:
                self._log(f"  {issue}\n")
            self._log("\nGrund: Zu wenige Personen verfügbar.\n")
        else:
            self._log("✅  Alle Regeln eingehalten!\n")

        self._log(f"\n📄  Datei: {filename}\n")
        self._log("\n" + "━" * 50 + "\n")
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

        self.result_text.config(state="disabled")

        # Status
        self.status_label.config(
            text=f"✅ {filename} erstellt!" if not issues
            else f"⚠️ {filename} — mit Einschränkungen",
            fg=GREEN if not issues else YELLOW
        )

        # Ask to open
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
    root.iconname("Arbeitskalender")
    app = ArbeitskalenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()