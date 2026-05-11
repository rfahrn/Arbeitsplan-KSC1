# Regeln Arbeitsplan KSC1

> Diese Datei ist die Single-Source-of-Truth für die Scheduling-Logik.
> `arbeitskalender.py` implementiert diese Regeln; `REGELN.txt` ist die
> Plaintext-Variante derselben Information.

---

## 1) Tägliche Aufgaben — Übersicht

| Aufgabe | Code | Anzahl / Tag | VM | NM | Berechtigte Personen | Rhythmus |
|---|---|---|---|---|---|---|
| Telefon | `TEL` | **Mo VM = 4**, sonst **3 / Halbtag** | ✅ | ✅ | alle verfügbaren; TLs bevorzugt, max. 2× / Woche / TL | täglich, jeder Halbtag |
| Abklärung | `ABKL` | **2 / Halbtag** | ✅ | ✅ | alle verfügbaren | täglich, VM + NM |
| Tagespharm | `TAGES PA` | **1 / Tag**, ganzer Tag | ✅ | ✅ (gleiche Person) | nicht TLs, nicht Brigitte / Saskia / Dragi / Maria B. / Andrea A. / Florence | täglich, max. 1× / Person / Woche |
| Schalter | `ERF7/SCH` | **1 / Tag**, ganzer Tag | ✅ | ✅ (gleiche Person) | siehe §3 (Schalter-Pool) | täglich; alle 2 Wochen pro Person |
| HUB | `ERF7/HUB` | **2 / Tag** (halbtags) | ✅ (1 Person) | ✅ (andere Person) | Jesika, Dragi, Lara, Corinne, Dipiga, Nina, Amra, Alessia | täglich, VM + NM unterschiedliche Person |
| ERF8 | `ERF8` | **1 / Tag** | – | ✅ **strikt NM** | alle ausser Maria B., Saskia | täglich, nur Nachmittag (läuft bis 17:00); `ERF8/Q` & `ERF8/HUB` können manuell auch VM gesetzt werden |
| ERF9 Montag | `ERF9` | 1 ganztags | ✅ | ✅ | **Brigitte** (fix) | jeden Montag |
| ERF9 Dienstag | `ERF9` / `ERF9/TEL` | VM = `ERF9`, NM = `ERF9/TEL` | ✅ | ✅ | **Dragi** (fix) | jeden Dienstag |
| ERF9 Mi–Fr | `ERF9` | 1 (VM) | ✅ | – | alle ausser Brigitte, Dragi | Mi / Do / Fr, undefiniert |
| Postöffnung (PÖ + Scan) | `PO/SCAN` | **1 / Vormittag** | ✅ | – | nicht TLs, nicht Brigitte / Corinne / Maria B. / Saskia / Dragi | jeden Werktag VM |
| Postöffnung (nur PÖ) | `PO` | **2 / Vormittag** | ✅ | – | nicht TLs | jeden Werktag VM |
| Scanning Nachmittag | `SCAN` | **1 / Nachmittag** | – | ✅ | nicht TLs, nicht Brigitte / Corinne / Maria B. / Saskia / Dragi | jeden Werktag NM |
| Queue | `ERF7/Q` | **1 / Tag** (implizit über TL-Auffüllung) | ✅ | ✅ | TLs (Silvana, Linda, Lara) | TLs decken Queue ab |
| Tagesverantwortung | (Header) | 1 / Tag (VM + NM gleiche Person) | ✅ | ✅ | TLs (Silvana, Linda, Lara) im Wechsel | täglich, gleiche Person ganzer Tag |

---

## 2) Wöchentliche / spezielle Aufgaben

| Aufgabe | Code | Frequenz | VM/NM | Berechtigte Personen | Wann |
|---|---|---|---|---|---|
| KGS | `KGS` | 2 × / Woche | VM (Di), NM (Do) | **Dipiga** (fix) | Di VM + Do NM |
| KSC Spezial | `KSC Spez.` | 2 × / Woche | VM | **Martina** (fix) | Di VM + Do VM |
| Labor | `ERF5` | 1 × / Woche | NM | **Jesika** und **Dipiga** im Wechsel | Mittwoch NM |
| RX Abo | `RX Abo` | **2 × / Woche**, halbtags | VM oder NM | Lara, Linda, Isa (Isaura), Martina | 2 Slots / Woche, max. 1 / Tag |
| HO (Maria) | `HO` | 1 × / Woche | VM | **Maria B.** (fix) | Montag VM |
| HO/Q (Linda) | `HO/Q` | alle 2 Wochen | VM + NM | **Linda** (fix) | Dienstag ganztags, biweekly |
| Direktbestellung | (Wochenaufgabe) | 1 × / Woche | – | Andrea A., Andrea G., Silvana im Wechsel | rotierend |
| ONB | (Wochenaufgabe) | 1 × / Woche | – | Silvana, Linda, Lara, Andrea G., Dipiga, Isaura, Martina, Alessia, Amra, Nina, Jesika, Corinne | rotierend |
| BTM | (Wochenaufgabe) | 1 × / Woche | – | Jesika, Silvana, Linda, Lara, Dipiga, Isaura | rotierend |
| Vorbezüge | `VBZ/Q` | – | – | Lara, Nina (berechtigt, nicht fix eingeplant) | bei Bedarf |

---

## 3) Schalter-Pool

| Status | Personen |
|---|---|
| **Im Schalter-Pool (alle 2 Wochen dran)** | Silvana, Lara, Andrea G., Dipiga, Isaura, Martina, Alessia, Amra, Nina, Jesika |
| **Nie Schalter** | Linda, Florence, Corinne, Maria B., Andrea A. |
| **Nicht strikt alle 2 Wochen (seltener im Pool)** | Brigitte, Saskia, Dragi |

> Rotation: 10 Pool-Personen werden in zwei Gruppen à 5 geteilt; jede Gruppe ist
> in einer Woche aktiv (= 1 Person ganztags pro Werktag).

---

## 4) Ausschlüsse

| Liste | Personen | Bedeutung |
|---|---|---|
| Kein Scanning | Brigitte, Corinne, Maria B., Saskia, Dragi | Bekommen weder `PO/SCAN` noch `SCAN` zugewiesen |
| Keine ONB | Brigitte, Florence, Saskia, Dragi, Maria B., Andrea A. | Nicht für die ONB-Wochenaufgabe |
| Nicht TL | (alle ausser Silvana, Linda, Lara) | Bekommen kein `ERF7/Q` als Default-Auffüllung |

---

## 5) Pensen & Verfügbarkeit

| # | Person | Pensum | Wochenverfügbarkeit |
|---|---|---:|---|
| 1 | Grossenbacher Silvana (`Silvana`) | 100 % | Mo–Fr ganztags · **TL** |
| 2 | Rexhaj Majlinda (`Linda`) | 90 % | Mo–Fr ganztags, jeden 2. Freitag frei · **TL** |
| 3 | Lara Ierano (`Lara`) | 100 % | Mo–Fr ganztags · **TL** |
| 4 | Gygax Andrea (`Andrea G.`) | 100 % | Mo–Fr ganztags |
| 5 | Jeyalingam Dipiga (`Dipiga`) | 100 % | Mo–Fr ganztags |
| 6 | Bohnenblust Isaura (`Isaura`) | 90 % | Mo–Fr ganztags, jeden 2. Mittwoch frei |
| 7 | Martina Pizzi (`Martina`) | 80 % | Di–Fr ganztags, Montag frei |
| 8 | Giombanco Alessia (`Alessia`) | 100 % | Mo–Fr ganztags |
| 9 | Imsirovic Amra (`Amra`) | 100 % | Mo–Fr ganztags |
| 10 | Hänni Nina (`Nina`) | 100 % | Mo–Fr ganztags |
| 11 | Bushaj Jesika (`Jesika`) | 100 % | Mo–Fr ganztags |
| 12 | Siegrist Brigitte (`Brigitte`) | 70 % | Mo–Fr VM + Mo+Di NM (Mi/Do/Fr NM frei) |
| 13 | Florence Dornbierer (`Florence`) | 70 % | Mo–Do (Mi NM frei) + Freitag frei |
| 14 | Eggimann Corinne (`Corinne`) | 90 % | Mo–Fr ganztags, jeden 2. Freitag frei |
| 15 | Schöni Saskia (`Saskia`) | 40 % | nur Mi + Fr (ganztags) |
| 16 | Milenkovic Dragana (`Dragi`) | 60 % | nur Di, Do, Fr (ganztags) |
| 17 | Bruzzese Maria (`Maria B.`) | **30 % (effektiv 20 %)** | nur Mo VM + Do VM ⚠ |
| 18 | Ackermann Andrea (`Andrea A.`) | 60 % | nur Mo, Di, Mi (ganztags) |

> **⚠ Hinweis Maria B.:** Pensum 30 % impliziert 3 Halbtage, aber nur Mo VM
> und Do VM (= 2 Halbtage = 20 %) sind eingetragen. Inkonsistenz noch nicht
> aufgelöst – siehe Mail-Abstimmung.

---

## 6) Kombinations-Matrix

| Aufgabe | mit `Q` | mit `TEL` | mit `SCH` | mit `HUB` | mit `ABKL` | mit `SCAN` |
|---|:---:|:---:|:---:|:---:|:---:|:---:|
| `ERF7` | ✅ | – | ✅ | ✅ | – | – |
| `ERF8` | ✅ | – | – | (✅) | – | – |
| `ERF9` | ✅ | ✅ (Di NM) | – | – | – | – |
| `PO` | ✅ | ✅ | – | – | ✅ | – |
| `TPA` (`TAGES PA`) | ✅ | ✅ (NM) | – | – | – | – |
| `Schalter` (`ERF7/SCH`) | – | – | – | – | – | ❌ |
| `Schalter` (Person ganztags) | – | – | – | – | – | ❌ kein PÖ/Scan |
| `HUB` (`ERF7/HUB`) | – | – | – | – | – | – |

Legende: ✅ = erlaubt · ❌ = verboten · – = nicht relevant / nicht definiert · (✅) = möglich aber selten.

---

## 7) Prioritäts-Reihenfolge im Generator

Der Scheduler arbeitet in dieser Reihenfolge — Aufgaben weiter oben sichern sich Plätze zuerst:

```
1.  Overrides (Krank, Ferien, Termin) aus Schritt-1-Eingaben
        ↓
2.  Fixzuweisungen:
    – Brigitte Mo ERF9 ganztags
    – Dragi Di ERF9 VM + ERF9/TEL NM
    – Maria B. Mo VM HO
    – Linda Di HO/Q ganztags (alle 2 Wochen)
    – Dipiga Di VM KGS + Do NM KGS
    – Martina Di VM + Do VM KSC Spez.
    – Labor Mi NM (Jesika ↔ Dipiga im Wechsel)
        ↓
3.  Schalter ganztags (ERF7/SCH) – aktiver 5er-Pool dieser Woche
        ↓
4.  RX Abo (2 Slots / Woche)
        ↓
5.  Tagespharm ganztags (TAGES PA) – 1 Person / Tag, max. 1× / Person / Woche
        ↓
6.  HUB halbtags (ERF7/HUB) – 1 VM + 1 (andere) NM pro Werktag
        ↓
7.  PÖ Vormittag (1 PO/SCAN + 2 PO) und 1 NM Scan
        ↓
8.  ERF8 (1 / Tag)
        ↓
9.  ERF9 Mi–Fr (1 VM)
        ↓
10. TEL füllen (Target: 4 Mo VM, sonst 3)
        ↓
11. ABKL füllen (2 / Halbtag)
        ↓
12. Restliche freie Slots:
    – TL → ERF7/Q
    – Andere VM → PO
    – Andere NM → ERF7
        ↓
13. Tagesverantwortung & PHC-Liste (TL-Auswahl pro Tag)
        ↓
14. Wochenaufgaben rotieren (Direkt, ONB, BTM)
```

---

## 8) Abwesenheiten (Schritt 1)

| Eingabe-Typ | Task-Code | Belegt Slots | Wirkung |
|---|---|---|---|
| Krank | `KRANK` | gewählte Tage/Halbtage | Person fällt aus, blockiert Slot |
| Ferien | `FERIEN` | gewählte Tage/Halbtage | Person fällt aus, blockiert Slot |
| Termin | `*<Notiz>` | gewählter Halbtag | Slot wird leer; Notiz erscheint |
| Home Office | `HO` | gewählte Halbtage | Person macht HO statt Plan-Task |
| Custom | freier Text | gewählte Halbtage | Beliebige Bezeichnung |

> **Wochen-Range:** Pro Eintrag wird `Anzahl Wochen` (Default 1) eingegeben.
> Eintrag gilt nur für `[KW_eingabe … KW_eingabe + N − 1]`. Es gibt keine
> automatische Übernahme auf alle Folgewochen.
