#!/usr/bin/env python3
"""
Regel-Validator: Prüft ALLE definierten Regeln über mehrere Wochen.
"""
import json, os
from datetime import datetime, timedelta
from arbeitskalender import build_schedule, create_employees, DAYS, SLOTS

PASS = "✅"
FAIL = "❌"
WARN = "⚠️"
total_pass = 0
total_fail = 0
total_warn = 0

def check(rule_name, condition, detail=""):
    global total_pass, total_fail
    if condition:
        total_pass += 1
        print(f"  {PASS} {rule_name}")
    else:
        total_fail += 1
        print(f"  {FAIL} {rule_name}  →  {detail}")

def warn(rule_name, detail=""):
    global total_warn
    total_warn += 1
    print(f"  {WARN} {rule_name}  →  {detail}")

def validate_week(sched, kw, employees):
    print(f"\n{'='*70}")
    print(f"  VALIDIERUNG KW {kw} ({sched.week_start_date.strftime('%d.%m.%Y')})")
    print(f"{'='*70}")

    # ──────────────────────────────────────────
    # REGEL 1: Mo VM immer 4 Pers am TEL
    # ──────────────────────────────────────────
    tel_mo_vm = sched.tel_count[0][0]
    check("Mo VM: 4 Pers am TEL", tel_mo_vm >= 4, f"nur {tel_mo_vm}")

    # ──────────────────────────────────────────
    # REGEL 2: Sonst jeweils 3 Pers am TEL
    # ──────────────────────────────────────────
    for day in range(5):
        for slot in range(2):
            if day == 0 and slot == 0:
                continue
            c = sched.tel_count[day][slot]
            check(f"{DAYS[day]} {SLOTS[slot]}: 3 Pers am TEL", c >= 3,
                  f"nur {c}")

    # ──────────────────────────────────────────
    # REGEL 3: Täglich VM und NM je 2 am ABKL
    # ──────────────────────────────────────────
    for day in range(5):
        for slot in range(2):
            c = sched.abkl_count[day][slot]
            check(f"{DAYS[day]} {SLOTS[slot]}: 2 Pers ABKL", c >= 2,
                  f"nur {c}")

    # ──────────────────────────────────────────
    # REGEL 4: Brigitte jeden Montag ERF9
    # ──────────────────────────────────────────
    brig_mo = sched.get_task("Brigitte", 0, 0)
    check("Brigitte Mo VM = ERF9", "ERF9" in brig_mo, f"hat: {brig_mo}")

    # ──────────────────────────────────────────
    # REGEL 5: Dragi jeden Dienstag ERF9
    # ──────────────────────────────────────────
    if employees["Dragi"].is_available(1, 0):
        dragi_di = sched.get_task("Dragi", 1, 0)
        check("Dragi Di VM = ERF9", "ERF9" in dragi_di, f"hat: {dragi_di}")

    # ──────────────────────────────────────────
    # REGEL 6: 1 Person pro Tag ERF8
    # ──────────────────────────────────────────
    for day in range(5):
        erf8_found = False
        for name in employees:
            for slot in range(2):
                if "ERF8" in sched.get_task(name, day, slot):
                    erf8_found = True
        check(f"{DAYS[day]}: mind. 1 Pers ERF8", erf8_found)

    # ──────────────────────────────────────────
    # REGEL 7: Dipiga KGS Di-VM und Do-NM
    # ──────────────────────────────────────────
    check("Dipiga Di VM = KGT",
          "KGT" in sched.get_task("Dipiga", 1, 0),
          f"hat: {sched.get_task('Dipiga', 1, 0)}")
    check("Dipiga Do NM = KGT",
          "KGT" in sched.get_task("Dipiga", 3, 1),
          f"hat: {sched.get_task('Dipiga', 3, 1)}")

    # ──────────────────────────────────────────
    # REGEL 8: Martina KSC Spez. Di und Do VM
    # ──────────────────────────────────────────
    if employees["Martina"].is_available(1, 0):
        check("Martina Di VM = KSC Spez.",
              "KSC" in sched.get_task("Martina", 1, 0),
              f"hat: {sched.get_task('Martina', 1, 0)}")
    if employees["Martina"].is_available(3, 0):
        check("Martina Do VM = KSC Spez.",
              "KSC" in sched.get_task("Martina", 3, 0),
              f"hat: {sched.get_task('Martina', 3, 0)}")

    # ──────────────────────────────────────────
    # REGEL 9: Florence PO wenn am Dienstag
    # ──────────────────────────────────────────
    if employees["Florence"].is_available(1, 0):
        flor_di = sched.get_task("Florence", 1, 0)
        check("Florence Di VM = PO", "PO" in flor_di, f"hat: {flor_di}")

    # ──────────────────────────────────────────
    # REGEL 10: Maria B. Mo VM = HO
    # ──────────────────────────────────────────
    if employees["Maria B."].is_available(0, 0):
        mb_mo = sched.get_task("Maria B.", 0, 0)
        check("Maria B. Mo VM = HO", "HO" in mb_mo, f"hat: {mb_mo}")

    # ──────────────────────────────────────────
    # REGEL 11: Kein Scanning für bestimmte Personen
    # ──────────────────────────────────────────
    no_scan = ["Brigitte", "Corinne", "Maria B.", "Saskia", "Dragi"]
    for name in no_scan:
        has_scan = False
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "SCAN" in t:
                    has_scan = True
        check(f"{name}: kein Scanning", not has_scan,
              "hat Scanning zugewiesen!")

    # ──────────────────────────────────────────
    # REGEL 12: Stephi hat immer KGT
    # ──────────────────────────────────────────
    if "Stephi" in employees:
        stephi_ok = True
        for day in range(5):
            for slot in range(2):
                t = sched.get_task("Stephi", day, slot)
                if t != "KGT":
                    stephi_ok = False
        check("Stephi: immer KGT", stephi_ok,
              "Stephi hat nicht überall KGT!")

    # ──────────────────────────────────────────
    # REGEL 13: HUB nur für berechtigte Personen
    # ──────────────────────────────────────────
    hub_allowed = {"Jesika", "Dragi", "Lara", "Corinne", "Dipiga", "Nina", "Amra", "Alessia"}
    for name in employees:
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "HUB" in t and name not in hub_allowed:
                    check(f"{name}: kein HUB erlaubt", False,
                          f"hat HUB an {DAYS[day]} {SLOTS[slot]}")
    check("HUB nur für berechtigte Personen",
          all(name in hub_allowed
              for name in employees
              for day in range(5)
              for slot in range(2)
              if "HUB" in sched.get_task(name, day, slot)))

    # ──────────────────────────────────────────
    # REGEL 14: Direktbestellung nur Andrea A., Andrea G., Silvana
    # ──────────────────────────────────────────
    direkt_allowed = {"Andrea A.", "Andrea G.", "Silvana"}
    for name in employees:
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "Direkt" in t and name not in direkt_allowed:
                    check(f"{name}: kein Direkt erlaubt", False,
                          f"hat Direkt an {DAYS[day]} {SLOTS[slot]}")
    check("Direkt nur für berechtigte Personen",
          all(name in direkt_allowed
              for name in employees
              for day in range(5)
              for slot in range(2)
              if "Direkt" in sched.get_task(name, day, slot)))

    # ──────────────────────────────────────────
    # REGEL 15: RX/Fx Abo nur Lara, Linda, Isaura, Martina
    # ──────────────────────────────────────────
    abo_allowed = {"Lara", "Linda", "Isaura", "Martina"}
    for name in employees:
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "Abo" in t and name not in abo_allowed:
                    check(f"{name}: kein Abo erlaubt", False,
                          f"hat Abo an {DAYS[day]} {SLOTS[slot]}")
    check("RX Abo nur für berechtigte Personen",
          all(name in abo_allowed
              for name in employees
              for day in range(5)
              for slot in range(2)
              if "Abo" in sched.get_task(name, day, slot)))

    # ──────────────────────────────────────────
    # REGEL 16: Schalter NICHT für ausgeschlossene
    # ──────────────────────────────────────────
    no_schalter = ["Linda", "Florence", "Corinne", "Maria B.", "Andrea A.",
                   "Brigitte", "Saskia", "Dragi"]
    for name in no_schalter:
        has_sch = False
        for day in range(5):
            for slot in range(2):
                if "SCH" in sched.get_task(name, day, slot):
                    has_sch = True
        check(f"{name}: kein Schalter", not has_sch,
              "hat Schalter zugewiesen!")

    # ──────────────────────────────────────────
    # REGEL 17: Niemand arbeitet wenn nicht verfügbar
    # ──────────────────────────────────────────
    availability_ok = True
    for name, emp in employees.items():
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if t and not emp.is_available(day, slot):
                    check(f"{name}: nicht verfügbar {DAYS[day]} {SLOTS[slot]}", False,
                          f"aber zugewiesen: {t}")
                    availability_ok = False
    if availability_ok:
        check("Verfügbarkeit: niemand arbeitet wenn frei", True)

    # ──────────────────────────────────────────
    # REGEL 18: ONB-Ausschluss
    # ──────────────────────────────────────────
    no_onb = [n for n in employees if not employees[n].can_onb]
    for name in no_onb:
        has_onb = False
        for day in range(5):
            for slot in range(2):
                t = sched.get_task(name, day, slot)
                if "ONB" in t:
                    has_onb = True
        check(f"{name}: kein ONB", not has_onb,
              "hat ONB zugewiesen!")

    # ──────────────────────────────────────────
    # REGEL 19: Brigitte Mo NM = ERF9
    # ──────────────────────────────────────────
    brig_mo_nm = sched.get_task("Brigitte", 0, 1)
    check("Brigitte Mo NM = ERF9", "ERF9" in brig_mo_nm, f"hat: {brig_mo_nm}")

    # ──────────────────────────────────────────
    # SOFT-CHECK: TL 1-2x/Wo TEL
    # ──────────────────────────────────────────
    for name in employees:
        if not employees[name].is_tl:
            continue
        tel_c = sum(1 for d in range(5) for s in range(2)
                    if "TEL" in sched.get_task(name, d, s))
        if tel_c == 0:
            warn(f"TL {name}: 0x TEL diese Woche", "Soll: 1-2x")
        elif tel_c > 2:
            warn(f"TL {name}: {tel_c}x TEL diese Woche", "Soll: 1-2x")

    # ──────────────────────────────────────────
    # HINWEISE (undefinierte Regeln)
    # ──────────────────────────────────────────
    print(f"\n  Hinweise (teilweise undefinierte Regeln):")
    warn("TL Tagesverantwortung Wechsel", "Nicht genau definiert")
    warn("TPA undefiniert NM möglich", "Nicht implementiert (undefiniert)")
    warn("Mi-Fr ERF9 Person", "Undefiniert - wird zufällig zugeteilt")


# ──────────────────────────────────────────────────
# MULTI-WOCHEN-WECHSEL-PRÜFUNG
# ──────────────────────────────────────────────────

def validate_alternation(weeks_data):
    print(f"\n{'='*70}")
    print(f"  WECHSEL-PRÜFUNG über {len(weeks_data)} Wochen")
    print(f"{'='*70}")

    labor_pattern = []
    linda_ho_pattern = []
    schalter_pattern = []
    friday_tv_pattern = []

    for kw, sched, state in weeks_data:
        # Labor
        jes = sched.get_task("Jesika", 2, 1)
        dip = sched.get_task("Dipiga", 2, 1)
        if "ERF5" in jes:
            labor_pattern.append("Jesika")
        elif "ERF5" in dip:
            labor_pattern.append("Dipiga")
        else:
            labor_pattern.append("?")

        # Linda HO
        linda_di = sched.get_task("Linda", 1, 0)
        linda_ho_pattern.append("HO" if "HO" in linda_di else "-")

        # Schalter Gruppe
        schalter_pattern.append(state.get("schalter_group", "?"))

        # Fr TV
        friday_tv_pattern.append(
            "Corinne" if state.get("friday_tagesv_corinne") else "Andrea G.")

    print(f"  Labor Mi NM:      {' → '.join(labor_pattern)}")
    alternates = all(labor_pattern[i] != labor_pattern[i+1]
                     for i in range(len(labor_pattern)-1))
    check("Labor wechselt jede Woche", alternates,
          f"Pattern: {labor_pattern}")

    print(f"  Linda HO Di:      {' → '.join(linda_ho_pattern)}")
    alternates = all(linda_ho_pattern[i] != linda_ho_pattern[i+1]
                     for i in range(len(linda_ho_pattern)-1))
    check("Linda HO wechselt alle 2 Wo", alternates,
          f"Pattern: {linda_ho_pattern}")

    print(f"  Schalter-Gruppe:  {' → '.join(str(s) for s in schalter_pattern)}")
    alternates = all(schalter_pattern[i] != schalter_pattern[i+1]
                     for i in range(len(schalter_pattern)-1))
    check("Schalter-Gruppe wechselt", alternates,
          f"Pattern: {schalter_pattern}")

    print(f"  Fr 18h-Liste:     {' → '.join(friday_tv_pattern)}")
    alternates = all(friday_tv_pattern[i] != friday_tv_pattern[i+1]
                     for i in range(len(friday_tv_pattern)-1))
    check("Fr 18h-Liste wechselt", alternates,
          f"Pattern: {friday_tv_pattern}")

    # Schalter-Gruppen-Stabilität: Personen mit SCH in Gruppe 0/1 sollen stabil sein
    schalter_people_per_week = []
    for kw, sched, state in weeks_data:
        people_with_sch = set()
        for name in sched.employees:
            for day in range(5):
                for slot in range(2):
                    if "SCH" in sched.get_task(name, day, slot):
                        people_with_sch.add(name)
        schalter_people_per_week.append((state.get("schalter_group", "?"), people_with_sch))

    # Prüfe: Gleiche Gruppe-Nummer sollte gleiche Personen haben
    group_members = {}
    stable = True
    for group, people in schalter_people_per_week:
        if people:  # nur wenn jemand Schalter hat
            if group in group_members:
                if group_members[group] != people:
                    stable = False
            else:
                group_members[group] = people
    check("Schalter-Gruppen stabil über Wochen", stable,
          f"Gruppen-Mitglieder ändern sich zwischen gleicher Gruppenummer")


# ──────────────────────────────────────────────────
# HAUPTTEST
# ──────────────────────────────────────────────────

if __name__ == "__main__":
    STATE_FILE = "test_state.json"
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)

    print("╔══════════════════════════════════════════════════════════════╗")
    print("║     ARBEITSKALENDER - VOLLSTÄNDIGE REGELVALIDIERUNG        ║")
    print("║     Test über 6 aufeinanderfolgende Wochen                 ║")
    print("╚══════════════════════════════════════════════════════════════╝")

    start = datetime(2026, 4, 13)
    weeks_data = []

    for i in range(6):
        monday = start + timedelta(weeks=i)
        kw = monday.isocalendar()[1]
        employees = create_employees(kw)
        sched = build_schedule(kw, monday, state_file=STATE_FILE)
        state = json.load(open(STATE_FILE))
        validate_week(sched, kw, employees)
        weeks_data.append((kw, sched, state))

    validate_alternation(weeks_data)

    os.remove(STATE_FILE)

    print(f"\n{'='*70}")
    print(f"  ENDERGEBNIS")
    print(f"{'='*70}")
    print(f"  {PASS} Bestanden:  {total_pass}")
    print(f"  {FAIL} Fehlgeschlagen: {total_fail}")
    print(f"  {WARN} Hinweise:   {total_warn} (undefinierte Regeln)")
    print(f"{'='*70}")
    if total_fail == 0:
        print(f"  🎉 ALLE DEFINIERTEN REGELN WERDEN EINGEHALTEN!")
    else:
        print(f"  ⚠️  {total_fail} Regel(n) verletzt - bitte prüfen!")
    print(f"{'='*70}")
