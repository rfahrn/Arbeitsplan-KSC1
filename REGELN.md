# Regeln Arbeitsplan KSC (Version 8.2, Oktober 2026)

> `arbeitskalender.py` setzt diese Regeln um; wer was darf und alle Fixeinsätze
> stehen in `team.json` (bearbeitbar in der App unter «⚙ Team & Regeln»).
> Rotationen werden aus der Kalenderwoche berechnet – keine Zustandsdatei.

---

## 1) Tägliche Aufgaben

| Aufgabe | Code | Anzahl | VM | NM | Wer | Regel |
|---|---|---|:-:|:-:|---|---|
| Telefon | `TEL` | **Mo VM 4**, sonst **3 pro Halbtag** | ✅ | ✅ | alle anwesenden; TL mind. 1× und max. 2× pro Woche (einstellbar); nur-TEL-Personen immer | **nur halbtags** pro Person (VM TEL + NM ABKL ok, nie 2× TEL) |
| Abklärung | `ABKL` | **2 pro Halbtag** | ✅ | ✅ | alle regulären | **nur halbtags** pro Person |
| Tagespharma | `TAGES PA` | 1 pro Tag, **ganztags** | ✅ | ✅ | Rolle TPA; keine TL, nicht Corinne / Saskia / Maria B. / Andrea A. / Florence / Brigitte / Dragi | niemand fix, wechselt täglich; NM `TAGES PA/TEL` bei Bedarf / immer / nie (einstellbar) |
| Schalter | `ERF7/SCH` (oder `ERF4/SCH`, einstellbar) | 1 pro Tag, **ganztags** | ✅ | ✅ | Rolle Schalter (Silvana, Lara, Andrea G., Dipiga, Isaura, Alessia, Amra, Nina, Jesika); Reserve Brigitte, Saskia, Dragi nur, wenn aus der Gruppe niemand frei ist | 2 Gruppen im Wochenwechsel, max. 1× pro Person und Woche |
| HUB | `ERF7/HUB` | 2 Halbtage pro Tag | ✅ (1 Person) | ✅ (andere Person) | Rolle HUB (Lara, Dipiga, Alessia, Amra, Nina, Jesika, Dragi) | nur mit ERF7 |
| ERF8 | `ERF8` | 1 pro Tag | – | ✅ | Rolle ERF8 (alle ausser Saskia, Maria B., nur-TEL, Aushilfe) | Nachmittag (bis 17:00) |
| ERF9 Mo | `ERF9` | ganztags | ✅ | ✅ | **Brigitte** (Fixeinsatz) | |
| ERF9 Di | `ERF9` / `ERF9/TEL` | ganztags | ✅ | ✅ | **Dragi** (Fixeinsatz, NM mit TEL) | |
| ERF9 Mi–Fr | `ERF9` | 1 pro Tag, **ganztags** | ✅ | ✅ | Rolle ERF9 (nicht Brigitte, Dragi) | wechselt täglich |
| Postöffnung + Scan | `PO/SCAN` | 1 pro Vormittag | ✅ | – | **Aushilfe Sarah** täglich; fehlt sie: Rolle PO + Scanning, Di bevorzugt Florence | nie die PHC-Person des Tages |
| Zusatz-Postöffnung | `PO` | 2 pro Vormittag (einstellbar) | ✅ | – | Rolle PO (nicht TL, nicht Brigitte / Corinne / Saskia / Dragi / Maria B.) | rotierend – niemand jeden Tag |
| Scanning | `scanning` | Aushilfe täglich NM; Stammteam gelegentlich | ✅ | ✅ | Rolle Scanning (nicht Brigitte / Corinne / Saskia / Dragi / Maria B.) | **nur halbtags** (Ausnahme Aushilfe), max. 2 Halbtage pro Woche, nie die PHC-Person |
| Queue | `ERF7/Q` | Auffüllung | ✅ | ✅ | TL | |
| Tagesverantwortung | `TL/Q` (Header-Farbe) | 1 pro Tag, **gleiche TL VM + NM** | ✅ | ✅ | TL im Wechsel (Lara → Silvana → Linda) | TL mit TV hat **keine Ämtli** |
| PHC-/18-Uhr-Liste | (dunkelblau, NM) | 1 pro Tag | – | ✅ | Mo TL im Wechsel · Di Dipiga · Mi/Do Pool · Fr Corinne / Andrea G. im Wechsel | nie Brigitte, Saskia, Maria B., nur-TEL, Aushilfe; Corinne + Andrea G. nicht Mi/Do; PHC-Person scannt nicht |

`PO/TEL` am Nachmittag: wird **nicht** automatisch vergeben (nur Notfall); manuell gesetzt → Hinweis.

## 2) Wöchentliche / spezielle Einsätze

| Einsatz | Code | Wann | Wer |
|---|---|---|---|
| KGS | `KGS` | Di VM + Do NM | Dipiga (Fixeinsatz) |
| Labor | `ERF5` | Mi NM | Jesika / Dipiga im Wechsel |
| Home Office | `HO` | Mo VM | Maria B. (Mo NM vor Ort) |
| Home Office | `HO` | Mi VM | Florence (Mi NM frei, Fr anwesend) |
| Home Office + Queue | `HO/Q` | Di ganztags, ungerade KW | Linda |
| KSC Spezial | `KSC Spez.` | Di VM + Do VM | Martina (zurzeit inaktiv) |
| RX Abo | `RX Abo` | 2 Halbtage pro Woche, max. 1 pro Tag | Lara, Linda, Isaura (Martina inaktiv) |
| Direktbestellung | Wochenaufgabe | 1 Person pro Woche | Silvana, Andrea G., Andrea A. |
| ONB | Wochenaufgabe | 1 Person pro Woche | Silvana, Linda, Lara, Andrea G., Dipiga, Isaura, Alessia, Amra, Nina, Jesika, Corinne |
| BTM | Wochenaufgabe | 1 Person pro Woche | Silvana, Linda, Lara, Dipiga, Isaura, Jesika |
| Vorbezüge | `VBZ` | nicht automatisch | Lara, Nina |

Wochenaufgaben: nie dieselbe Person doppelt; wer die ganze Woche fehlt, wird übersprungen.

## 3) Personen (Stand Oktober 2026, aus `team.json`)

| Kurzname | Anzeige | % | Anwesenheit | Besonderes |
|---|---|---:|---|---|
| Silvana | Silvana G. | 100 | Mo–Fr | TL, Direkt |
| Linda | Linda R. | 90 | Mo–Fr, gerade KW Fr frei | TL, RX Abo, Di HO/Q (ungerade KW), nie Schalter |
| Lara | Lara I. | 100 | Mo–Fr | TL, RX Abo, HUB, VBZ |
| Andrea G. | Andrea G. | 100 | Mo–Fr | Direkt, PHC nicht Mi/Do, Fr PHC im Wechsel |
| Dipiga | Dipiga J. | 100 | Mo–Fr | HUB, KGS Di VM + Do NM, PHC Di, Labor-Wechsel |
| Isaura | Isaura B. | 90 | Mo–Fr, gerade KW Mi frei | RX Abo |
| Martina | Martina P. | 80 | Mo frei | **inaktiv** (krank, dann Mutterschutz) |
| Alessia | Alessia G. | 100 | Mo–Fr | HUB |
| Amra | Amra I. | 100 | Mo–Fr | HUB |
| Nina | Nina H. | 100 | Mo–Fr | HUB, VBZ |
| Jesika | Jesika B. | 100 | Mo–Fr | HUB, Labor-Wechsel |
| Brigitte | Brigitte S. | 70 | Mi/Do/Fr nur VM | Mo ERF9; Schalter nur als Reserve; kein Scanning/ONB/PO/PHC/TPA |
| Florence | Florence D. | 90* | Mi nur VM (HO) | Di PO bevorzugt; keine ONB, kein Schalter, kein TPA |
| Corinne | Corinne E. | 90 | Mo–Fr, gerade KW Fr frei | keine PO/TPA/HUB/Scanning/Schalter; PHC nicht Mi/Do |
| Saskia | Saskia S. | 40 | nur Mi + Fr | TEL/ABKL/ERF7/ERF9; Schalter nur als Reserve |
| Dragi | Dragi M. | 60 | nur Di/Do/Fr | Di ERF9 + ERF9/TEL, HUB; Schalter nur als Reserve; kein Scanning/ONB/PO/TPA |
| Maria B. | Maria B. | 30 | Mo ganz, Do VM | Mo VM HO; keine Rollen |
| Andrea A. | Andrea A. | 60 | nur Mo/Di/Mi | Direkt; keine ONB, kein Schalter, kein TPA |
| Manuela | Manuela I. | 80* | Mi + Fr nur VM | **nur TEL** |
| Maria F. | Maria F. | 80* | Mo frei | **nur TEL** |
| Sarah | Sarah E. | 100 | Mo–Fr | **Aushilfe**: PO/SCAN VM + scanning NM |

\* Pensum angenommen (noch vom Team zu bestätigen).

BackUp Team (nur manuell): Stephi R. · Sabrina D. (ERP) · Brigitte F. (ERP) · Dominique L. (ERP) · KSC Spezial · KSV

## 4) Reihenfolge im Generator

```
 0. Abwesenheiten / Sonderwünsche (Krank, Ferien, Termin, HO, Custom; BackUp-Einträge)
 1. Fixeinsätze aus team.json (auch «nur gerade/ungerade KW»)
 2. Labor Mi NM (Wechsel)
 3. Aushilfe PO/SCAN + scanning · nur-TEL-Personen TEL
 4. Tagesverantwortung → TL/Q ganztags
 5. PHC-Liste
 5b. Telefon-Minimum der TL reservieren
 6. Schalter ganztags (aktive Gruppe, dann Reserve, dann andere Gruppe)
 7. TAGES PA ganztags
 8. ERF9 Mi–Fr ganztags
 9. RX Abo
10. Postöffnung (Ersatz-PO/SCAN, Zusatz-PO, NM-Scan-Ersatz)
11. ERF8 (NM)
12. TEL (halbtags, TL bevorzugt, gleichmässig)
13. ABKL (halbtags, gleichmässig)
14. HUB (VM + NM, verschiedene Personen)
15. Rest: ERF7/Q (TL), scanning (gelegentlich), ERF7
16. Wochenaufgaben Direkt / ONB / BTM
```

## 5) Abwesenheiten (Schritt 1 der App)

| Grund | Eintrag im Plan | Wirkung |
|---|---|---|
| Krank | `KRANK` | Halbtag blockiert, Person zählt als abwesend |
| Ferien | `FERIEN` | wie Krank |
| Termin | leer + Notiz in «Bemerkungen» | Halbtag blockiert |
| Home Office | `HO` | Person arbeitet im HO, bekommt keine andere Aufgabe |
| Custom | beliebiger Task-Code | z. B. `ERP`, `KSV` für BackUp-Zeilen |

Jeder Eintrag gilt für «Anzahl Wochen» ab der angezeigten KW. Ferien über mehrere
Wochen (von–bis) trägt man im **📆 Jahresplan** ein – sie gelten automatisch für jede
betroffene Woche.

## 6) Regelprüfung nach dem Planen

❌ = harte Regel verletzt (Mindestbesetzung TEL/ABKL, ERF8/ERF9/PO/TAGES PA pro Tag,
Halbtags-Regeln, PHC-Ausschlüsse, Rollen-Berechtigungen – auch bei manuellen Einträgen),
⚠️ = Soll-Regel (ERF9/TPA/Schalter halbtags, PO/TEL NM, Scanning-Limit, TL mit Ämtli, TL ohne/zu viel TEL),
ℹ️ = Vertretung / Rotation (TL abwesend, Schalter-Ersatz aus anderer Gruppe, Wochenaufgabe übersprungen).

## 7) Feiertage (📆 Jahresplan → Feiertage)

An einem Feiertag hat das ganze Team frei: alle, die laut Stammdaten arbeiten würden,
haben `Feiertag` (rot wie in der Vorlage), die BackUp-Zeilen ebenfalls. Es wird nichts
geplant und keine Besetzung geprüft; Abwesenheiten zählen an Feiertagen nicht.

| Feiertag | Wirkung |
|---|---|
| ganzer Tag | keine Tagesverantwortung, 18-Uhr-Liste, TAGES PA, Schalter, ERF9, PO, ERF8, TEL/ABKL |
| nur Nachmittag (z. B. Heiligabend) | Vormittag normal (TV als TL/Q, PO, TEL/ABKL); keine 18-Uhr-Liste, kein ERF8, kein Labor; TAGES PA / Schalter / ERF9 nur am Vormittag |
| nur Vormittag | Nachmittag normal; keine Postöffnung |

Voreingestellt: Neujahr, Berchtoldstag, Karfreitag, Ostermontag, Auffahrt,
Pfingstmontag, 1. August, Weihnachten, Stephanstag – einstellbar, dazu eigene freie Tage.
Ohne Feiertag in der Woche plant der Generator genau wie vorher.
