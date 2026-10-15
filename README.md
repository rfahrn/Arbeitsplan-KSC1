# KSC Arbeitsplan – Streamlit-App

Wöchentlicher Arbeitsplan-Generator für das KSC-Team, online unter
**https://arbeitsplan-ksc.streamlit.app/**

Erzeugt pro Kalenderwoche einen farbcodierten Wochenplan (Vorschau + Excel nach
der bestehenden Vorlage), hält alle harten Regeln ein und zeigt Konflikte
transparent an. **Version 7.0 (Oktober 2026)**: alle Rückmeldungen des Teams
eingearbeitet, Nachnamen anonymisiert (nur noch «Vorname N.»), und das Team ist
komplett **ohne Code-Änderung** pflegbar – Personen kommen und gehen, Pensen,
Anwesenheiten, Rollen, Fixeinsätze und Rotationen werden in der App bearbeitet.

---

## Dateien

| Datei | Zweck |
|---|---|
| `streamlit_app.py` | Die App (Hauptdatei für Streamlit Cloud) |
| `arbeitskalender.py` | Kernmodul: Scheduler, Regelprüfung, Excel-Export – ohne Streamlit nutzbar |
| `team.json` | **Stammdaten**: Personen, Pensen, Anwesenheit, Rollen, Fixeinsätze, Rotationen, Besetzungsregeln, BackUp-Team |
| `abwesenheiten.json` | Erfasste Abwesenheiten/Sonderwünsche (wird von der App gepflegt, nicht im Git) |
| `REGELN.md` | Regelübersicht in Tabellenform |
| `tests/test_regeln.py` | Regel-Validator über mehrere Wochen (`python tests/test_regeln.py`) |
| `requirements.txt`, `.streamlit/config.toml`, `.devcontainer/` | Deployment / Codespaces |

Es gibt **keine Zustandsdatei** mehr: Rotationen (Tagesverantwortung, PHC-Liste,
Schalter-Gruppen, Labor, Wochenaufgaben) werden aus der Kalenderwoche berechnet.
Dieselbe KW ergibt immer denselben Plan – egal wann und wie oft sie generiert wird.

---

## Schnellstart

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py          # http://localhost:8501
python tests/test_regeln.py             # Regeln über 8 Wochen prüfen
python arbeitskalender.py --montag 2026-10-19 --out plan.xlsx   # ohne UI
```

**Streamlit Cloud:** Repo `rfahrn/Arbeitsplan-KSC1`, Branch `main`, Hauptdatei
`streamlit_app.py`. Jeder Push auf `main` deployt neu.

---

## Bedienung

### Schritt 1 – Planung

1. **Kalenderwoche** oben wählen (‹ › – Standard ist die nächste Woche).
2. **Abwesenheit oder Sonderwunsch** erfassen: Person (auch BackUp-Team),
   Grund (Krank / Ferien / Termin / Home Office / Custom), Halbtag, Tage,
   optional Notiz, **Anzahl Wochen** (z. B. Ferien über 3 Wochen). Einträge
   gelten ab der angezeigten KW und bleiben gespeichert, bis sie entfernt werden
   («Vergangene Einträge aufräumen» löscht abgelaufene).
   - *Termin*: Halbtag bleibt leer, Notiz erscheint in der Bemerkungsspalte.
   - *Custom*: Notiz = Task-Code (z. B. `ERP`, `KSV`, `TEL`) – so werden auch die
     BackUp-Zeilen (Stephi, SPP-Team, KSC Spezial, KSV) befüllt.
3. **→ Plan generieren**.

### Schritt 2 – Ergebnis

- Wochenplan mit Tagesverantwortung im Tagesheader, PHC-Liste dunkelblau,
  BackUp-Team unterhalb.
- Tabelle «Tagesrollen»: Tagesverantwortung, PHC, TAGES PA, Postöffnung,
  Zusatz-PO, Schalter, ERF9 pro Tag; Wochenaufgaben Direkt / ONB / BTM.
- Banner: ❌ harte Regel verletzt (auch «⚠️ OFFEN» im Plan), ⚠️ Hinweise,
  ℹ️ Vertretungen (z. B. TL abwesend → andere TL übernimmt).
- **Plan bearbeiten**: jede Zelle per Dropdown ändern (auch BackUp-Zeilen);
  Regelprüfung und Excel werden nachgezogen.
- **⬇ Excel herunterladen** (`Arbeitsplan_KWxx_YYYYMMDD.xlsx`).

### Team verwalten (ohne Code)

Im Aufklapper **👥 Team verwalten** (Schritt 1):

| Tab | Was |
|---|---|
| **Personen** | Tabelle: Kurzname, Anzeigename («Vorname N.»), Pensum, Aktiv, TL, nur TEL, Aushilfe, Anwesenheit Mo–Fr (ganz / VM / NM / frei), Bemerkung. Zeile hinzufügen = neue Person, Zeile löschen = Person entfernen. **Deaktivieren** statt löschen, wenn jemand vorübergehend fehlt (z. B. Martina). |
| **Rollen & Fixeinsätze** | Pro Person: was sie darf (HUB, RX Abo, Schalter, Scanning, PO, TPA, PHC, ERF8, ERF9-Rotation, ONB, Direkt, BTM, VBZ), «alle 2 Wochen frei» (gerade/ungerade KW), fixe Einsätze (z. B. Mo ganz ERF9, Di VM KGS, Di ganz HO/Q nur in ungeraden KW). |
| **Rotation & Regeln** | Reihenfolge der TL für die Tagesverantwortung, PHC Montag/Freitag-Wechsel, PHC fix pro Tag, PHC-Ausschlüsse pro Tag, Labor-Wechsel, PO-Dienstag-Präferenz, Schalter-Gruppe und Wochenaufgaben **«diese Woche»** (damit stellst du die Rotation auf die Realität ein), sowie die Besetzungszahlen (TEL, ABKL, Zusatz-PO, HUB, RX Abo, Scanning). |
| **Entscheidungen** | Punkte, die das Team nicht eindeutig geregelt hat – hier entscheidet die Planerin: **Schalter-Reserve** (Brigitte, Saskia, Dragi «nicht alle 2 Wochen»: kommen nur dran, wenn aus der Gruppe niemand frei ist; Einsatzregel vor/nach der anderen Gruppe/nie), **Schalter-Code** im Plan (ERF7/SCH oder ERF4/SCH), **TL am Telefon** (Minimum und Maximum pro Woche; Abweichungen erscheinen als Hinweis), **TAGES PA/TEL** am Nachmittag (bei Bedarf / immer / nie), und die Liste der Codes, die nur manuell vergeben werden. |
| **BackUp Team** | Zeilen unterhalb des Teams mit Standard-Eintrag (SPP-Team: `ERP`). |
| **Speichern / Laden** | Speichern, `team.json` herunter-/hochladen, gespeicherte Version neu laden. |

Typische Fälle:

- **Neue Person**: Tab Personen → «+» → Kurzname, Name, %, Anwesenheit → Übernehmen →
  Tab Rollen → Rollen setzen → Übernehmen → Speichern.
- **Person verlässt das Team**: Aktiv-Häkchen entfernen (oder Zeile löschen) → Speichern.
  Fixeinsätze und Rotationen dieser Person entfallen automatisch; steht sie in einer
  Rotation (TV, PHC, Labor), zeigt die App einen Hinweis.
- **Pensum / Arbeitstage ändern**: Tab Personen → % und Anwesenheit → Übernehmen → Speichern.
- **Fixe Aufgabe wechselt** (z. B. Mo-ERF9 an eine andere Person): Tab Rollen,
  Fixeinsatz bei der einen Person löschen, bei der anderen hinzufügen.

### Speichern – wichtig für Streamlit Cloud

Streamlit Cloud hat **kein dauerhaftes Dateisystem**: lokal gespeicherte
Änderungen gehen beim nächsten Neustart der App verloren. Zwei Wege:

1. **Empfohlen – direkt ins GitHub-Repo speichern.** In den App-Secrets
   (Streamlit Cloud → App → Settings → Secrets) eintragen:

   ```toml
   [github]
   token  = "github_pat_…"        # Fine-grained PAT, Repo Arbeitsplan-KSC1, Permission «Contents: Read and write»
   repo   = "rfahrn/Arbeitsplan-KSC1"
   branch = "main"
   ```

   «💾 Speichern» schreibt dann `team.json` und `abwesenheiten.json` ins Repo
   (ein Commit pro Datei); Streamlit Cloud lädt die App danach automatisch neu.
   So funktioniert die Pflege auch, wenn niemand mit Git-Zugang da ist.
2. **Ohne Token**: «⬇ team.json» herunterladen und die Datei im Repo ersetzen
   (GitHub-Weboberfläche: Datei öffnen → Stift → Inhalt einfügen → Commit).

Lokal (eigener Rechner / Docker) wird immer in die Dateien neben der App geschrieben.
Pfade lassen sich mit `KSC_TEAM_CONFIG` und `KSC_ABSENCES` übersteuern.

---

## Regeln

Die vollständige Übersicht steht in **REGELN.md**. Kurzfassung der Planungsreihenfolge:

1. Abwesenheiten / Sonderwünsche
2. Fixeinsätze aus `team.json` (z. B. Brigitte Mo ERF9, Dragi Di ERF9 + ERF9/TEL,
   Dipiga KGS Di VM + Do NM, Maria B. Mo VM HO, Florence Mi VM HO, Linda Di HO/Q alle 2 Wochen)
3. Labor Mi NM (Jesika / Dipiga im Wechsel)
4. Aushilfe: täglich PO/SCAN (VM) + scanning (NM) · nur-TEL-Personen: jeder Halbtag TEL
5. Tagesverantwortung: dieselbe TL VM + NM, ohne Ämtli → **TL/Q**
6. PHC-/18-Uhr-Liste (Mo TL-Wechsel, Di Dipiga, Mi/Do Pool ohne Corinne/Andrea G.,
   Fr Corinne/Andrea G. im Wechsel; nie Brigitte, Saskia, Maria B.; PHC-Person scannt nicht)
6b. Telefon-Minimum der TL vorab reservieren (Standard 1 Halbtag pro Woche)
7. Schalter ganztags (ERF7/SCH), 1 Person pro Tag, 2 Gruppen im Wochenwechsel;
   Reserve-Personen (Brigitte, Saskia, Dragi) nur, wenn aus der Gruppe niemand frei ist
8. TAGES PA ganztags, 1 Person pro Tag, wechselnd (niemand fix, keine TL, nicht Corinne)
9. ERF9 Mi–Fr ganztags, wechselnd
10. RX Abo (2 Halbtage pro Woche, Pool Lara / Linda / Isaura)
11. Postöffnung: ersatzweise PO/SCAN, wenn die Aushilfe fehlt (Di bevorzugt Florence),
    plus 2 rotierende «PO»-Personen am Vormittag – niemand jeden Tag
12. ERF8 (1 Person, Nachmittag)
13. TEL (Mo VM 4, sonst 3) – nur halbtags pro Person, TL bevorzugt (max. 2/Woche),
    Nachmittag TAGES PA/TEL je nach Entscheidung (bei Bedarf / immer / nie)
14. ABKL (2 pro Halbtag) – nur halbtags pro Person
15. HUB: 1 Person VM + 1 andere NM (ERF7/HUB)
16. Rest: TL → ERF7/Q, sonst ERF7; gelegentlich ein halber Tag scanning (max. 2 pro Woche)
17. Wochenaufgaben Direkt / ONB / BTM: je 1 Person, nie dieselbe doppelt,
    ganzwöchig Abwesende übersprungen

«PO/TEL» am Nachmittag wird **nie automatisch** vergeben (Team: nur im Notfall,
weil man früher vom Telefon weg muss); manuell gesetzt erscheint ein Hinweis.

---

## Was wurde gegenüber der bisherigen Streamlit-Version (Mai 2026) geändert?

- Rückmeldungen Oktober 2026 (Confluence-Kommentare + Nachtrag) komplett umgesetzt:
  Martina vorerst inaktiv; Manuela I. und Maria F. (nur TEL); Sarah E. als Aushilfe
  (PO/SCAN + scanning täglich); BackUp-Team (Stephi, SPP mit ERP, KSC Spezial, KSV);
  Florence Mi VM Homeoffice + Fr anwesend; Maria B. Mo NM vor Ort, keine PO;
  Corinne keine PO/TPA/HUB; TEL/ABKL/Scanning nur halbtags; ERF9 ganztags;
  TL mit Tagesverantwortung = TL/Q; PHC-Ausschlüsse (Brigitte nie, Corinne/Andrea G.
  nicht Mi/Do), PHC-Person scannt nicht; TAGES PA niemand fix (Dipiga nur KGS);
  Postöffnung rotiert (vorher «Andrea A. jeden Tag PÖ»).
- Nachnamen anonymisiert: Plan und App zeigen nur «Vorname N.».
- Stammdaten in `team.json` + Team-Verwaltung in der App (vorher im Code).
- Keine Zustandsdatei mehr; Rotationen folgen der Kalenderwoche (vorher kippte die
  Rotation je nach Reihenfolge der Generierung).
- Wochenaufgaben wieder ohne Doppelbelegung (war im letzten Commit vom Mai verloren gegangen).
- Aufgeräumt: Tkinter-GUI, Windows-Batch, Docker-Kopie, zweite App-Variante, alte
  Testskripte entfernt.

---

## Offene Punkte – Annahmen, die das Team bestätigen sollte

| # | Thema | Annahme in Version 7.0 | Wo anpassen |
|---|---|---|---|
| 1 | Pensen Florence (neu Fr anwesend), Manuela, Maria F. | 90 / 80 / 80 % (aus den Halbtagen abgeleitet) | Team verwalten → Personen |
| 2 | Postöffnung am Vormittag | Aushilfe PO/SCAN **plus 2** rotierende PO-Personen (wie seit Mai) | Rotation & Regeln → «Zusatz-PO am VM» |
| 3 | «PÖ und TEL am Nachmittag nur Notfall» | gelesen als: die Kombination **PO/TEL** am Nachmittag wird nie automatisch geplant | – |
| 4 | RX Abo | 2 Halbtage pro Woche, max. 1 pro Tag, Pool Lara / Linda / Isaura | Rotation & Regeln → «RX Abo / Woche» (0 = aus) |
| 5 | Scanning Stammteam | ca. jeder dritte freie Halbtag bei Scanning-Berechtigten, max. 2 pro Woche | «Scanning-Anteil» / «max. Halbtage» |
| 6 | TAGES PA + TEL | TAGES PA/TEL nur, wenn die TEL-Besetzung am NM sonst nicht reicht | Entscheidungen |
| 7 | HUB | nur ERF7/HUB (nicht ERF8/HUB), 1 VM + 1 NM | «HUB / Tag» |
| 8 | Schalter | ganztags, 1 Person pro Tag, Pool in 2 Gruppen; Brigitte, Saskia, Dragi als **Reserve** (nur wenn aus der Gruppe niemand frei ist); Linda, Florence, Corinne, Maria B., Andrea A. nie | Entscheidungen / Rollen |
| 8b | Schalter-Code | Plan zeigt ERF7/SCH (seit Mai); April-Liste nannte ERF4/SCH | Entscheidungen |
| 8c | TL am Telefon | mind. 1, max. 2 Halbtage pro Woche; Abweichung = Hinweis | Entscheidungen |
| 9 | Bezeichnung TL/Q | für die TL mit Tagesverantwortung | Farbe/Code in `arbeitskalender.py` |
| 10 | 2-Wochen-Rhythmus | gerade KW = Linda / Isaura / Corinne frei; Linda HO/Q in ungeraden KW. 2026 hat 53 KW → Parität kippt an Neujahr; dann einmal «alle 2 Wochen frei» in team.json umstellen | Rollen → «Alle 2 Wochen frei» |
| 11 | VBZ (Vorbezüge) | wird nicht automatisch geplant (nur als Rolle hinterlegt) | – |

---

## Anonymisierung

Alle Anzeigenamen sind «Vorname N.» (`name` in `team.json`); Kurznamen sind die im
Team üblichen Vornamen. Die Git-Historie vor Oktober 2026 enthält noch alte Dateien
mit vollen Namen – wer das Repo öffentlich macht, sollte die Historie bereinigen
(z. B. `git filter-repo`) oder ein frisches Repo anlegen.
