# Arbeitsplan KSC · KSV – Streamlit-App

Wöchentlicher Arbeitsplan-Generator für das **KSC-** und das **KSV-Team**, online unter
**https://arbeitsplan-ksc.streamlit.app/** (KSV direkt: **…/?team=KSV**)

Erzeugt pro Kalenderwoche einen farbcodierten Wochenplan (Vorschau + Excel), hält
die Regeln des Teams ein und zeigt Konflikte transparent an. **Version 8.0
(Oktober 2026)**: eine App für beide Teams – Team oben links umschalten, gleiche
Funktionen, getrennte Daten. Das KSV-Team hat einen frei konfigurierbaren
Generator: Aufgaben, Anzahl Personen pro Tag, wer was darf und feste Einsätze
pflegt die KSV-Planerin selbst in der App – ohne Code.

---

## Ein Tool oder zwei?

**Ein Tool mit Team-Umschalter** – aus diesen Gründen:

- Dieselben Funktionen (Abwesenheiten über mehrere Wochen, Plan-Editor, Regelprüfung,
  Excel, Speichern ins GitHub-Repo, Team-Verwaltung) gibt es nur einmal; jede
  Verbesserung kommt beiden Teams zugute.
- Die Daten sind trotzdem strikt getrennt (`team.json` / `team_ksv.json`,
  `abwesenheiten.json` / `abwesenheiten_ksv.json`) – Änderungen im KSV-Team können
  den KSC-Plan nicht beeinflussen und umgekehrt.
- Die Regeln sind verschieden und bleiben es: KSC behält seinen bewährten
  Spezial-Generator (`arbeitskalender.py`, unverändert), KSV bekommt einen
  konfigurierbaren (`teamplaner.py`).
- **Eigener Zugang für die KSV-Planerin:** Link `…/?team=KSV` als Lesezeichen –
  die App öffnet direkt im KSV-Team. Wer eine komplett eigene App möchte, deployt
  dasselbe Repo ein zweites Mal auf Streamlit Cloud mit dem Secret
  `ARBEITSPLAN_TEAM = "KSV"` (oberste Ebene) – dann ist die App fest auf KSV
  eingestellt und der Umschalter verschwindet.

---

## Dateien

| Datei | Zweck |
|---|---|
| `streamlit_app.py` | Die App für beide Teams (Hauptdatei für Streamlit Cloud) |
| `arbeitskalender.py` | KSC-Generator: Scheduler mit Spezialregeln, Regelprüfung, Excel nach KSC-Vorlage |
| `teamplaner.py` | KSV-Generator: Aufgaben/Besetzung frei konfigurierbar, Regelprüfung, Excel im Jahresplan-Format |
| `team.json` / `team_ksv.json` | **Stammdaten** KSC / KSV: Personen, Anwesenheit, wer darf was, feste Einsätze, Regeln |
| `abwesenheiten.json` / `abwesenheiten_ksv.json` | Erfasste Abwesenheiten/Sonderwünsche (von der App gepflegt, nicht im Git) |
| `REGELN.md` / `REGELN_KSV.md` | Regelübersicht KSC / KSV in Tabellenform (KSV inkl. offene Punkte) |
| `tests/` | `test_regeln.py` (KSC), `test_ksv.py` (KSV), `test_app.py` (Oberfläche, beide Teams) |
| `requirements.txt`, `.streamlit/config.toml`, `.devcontainer/` | Deployment / Codespaces |

Es gibt **keine Zustandsdatei**: Rotationen und Verteilung ergeben sich aus der
Kalenderwoche. Dieselbe KW ergibt immer denselben Plan – egal wann und wie oft sie
generiert wird.

---

## Schnellstart

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py                 # http://localhost:8501  (KSV: http://localhost:8501/?team=KSV)
pytest tests/                                  # alle Tests (KSC, KSV, Oberfläche)
python tests/test_regeln.py                    # KSC-Regeln über 8 Wochen
python tests/test_ksv.py                       # KSV-Regeln über 12 Wochen
python arbeitskalender.py --montag 2026-10-19 --out plan.xlsx      # KSC ohne UI
python teamplaner.py --montag 2026-10-19 --out plan_ksv.xlsx       # KSV ohne UI
```

**Streamlit Cloud:** Repo `rfahrn/Arbeitsplan-KSC1`, Branch `main`, Hauptdatei
`streamlit_app.py`. Jeder Push auf `main` deployt neu.

---

## Bedienung

Oben links das **Team** wählen (KSC / KSV) – die Adresse ändert sich mit
(`?team=KSV`), so lässt sich jedes Team als Lesezeichen speichern. Der Titel
(«Arbeitsplan KSC» / «Arbeitsplan KSV», auch im Browser-Tab) folgt dem Team; der
Wortlaut steht als `APP_TITEL` oben in `streamlit_app.py`.

### Schritt 1 – Planung

1. **Woche** oben wählen: Auswahlliste mit allen Wochen von der laufenden bis ein Jahr im
   Voraus (oder ‹ › für eine Woche zurück/vor). Standard ist die nächste Woche.
2. **Abwesenheit oder Sonderwunsch** erfassen: Wer? · Was? (🤒 Krank, 🏖️ Ferien,
   🌙 Frei *(KSV)*, 🕑 Termin, 🏠 Home Office, 📌 Aufgabe setzen) · Tage · Halbtag ·
   wie viele Wochen · Notiz. Die Zeile «→ …» zeigt vor dem Hinzufügen, was
   eingetragen wird. Einträge gelten ab der angezeigten KW und bleiben gespeichert,
   bis sie entfernt werden.
   - *Termin*: Halbtag bleibt leer, die Notiz steht im Plan.
   - *Aufgabe setzen*: eine Aufgabe fest eintragen (Auswahl oder eigener Code, z. B.
     `ERP`, `Schulung`, `Einführung`) – so werden auch die BackUp-Zeilen befüllt.
3. **→ Plan generieren**.

### Schritt 2 – Ergebnis

- Ampel oben: ✓ alles eingehalten · ❌ Regel verletzt · ⚠️ Hinweise · ℹ️ Vertretungen.
- Wochenplan mit Datum pro Tag (KSC zusätzlich Tagesverantwortung, PHC-Liste dunkelblau).
- Darunter: **KSC** Tagesrollen + Wochenaufgaben; **KSV** «Wer hat was – pro Tag» und
  «Verteilung pro Person» (sofort sichtbar, ob die Arbeit fair verteilt ist).
- **Plan bearbeiten**: jede Zelle per Auswahl ändern; Schalter **Freitext** erlaubt
  eigenen Text wie `KSV2 10.00`. Regelprüfung und Excel werden nachgezogen.
- **⬇ Excel herunterladen** (oben rechts und beim Plan). KSV: Spalten A–L sind genau
  ein Wochenblock wie im Jahresplan und lassen sich dort direkt einfügen.

### Team & Regeln (ohne Code)

Eigene Seite **⚙ Team & Regeln**. Jede Änderung gilt **sofort** für den nächsten
Plan. Die Leiste oben zeigt, wie viele Änderungen noch ungespeichert sind («Was
wurde geändert?» listet sie lesbar auf), **💾 Speichern** macht sie dauerhaft,
**↶ Verwerfen** kehrt zum gespeicherten Stand zurück. «❓ Wie ändere ich …?» erklärt
die typischen Fälle – pro Team.

| Bereich | KSC | KSV |
|---|---|---|
| **👤 Personen** | Name, Nr. (Reihenfolge), Pensum, Aktiv, Funktion (Team / Teamleitung / nur Telefon / Aushilfe), Anwesenheit Mo–Fr, Bemerkung; ➕ Neue Person, 🗑️ Entfernen / Deaktivieren | dito, statt Funktion die **Zeilenfarbe** im Excel |
| **🎯 Wer darf was** | Häkchen-Matrix (HUB, ERF8, ERF9, TAGES PA, 18-Uhr-Liste, PO, Scanning, RX Abo, Direkt, ONB, BTM, VBZ), Schalter ja / Reserve | pro Aufgabe **✅ ja · 🟨 Reserve · — nie** |
| **📋 Aufgaben & Besetzung** | – | Aufgaben mit Code, Bedeutung, Farbe, Dauer (ganzer Tag / Halbtag / nur VM / nur NM / nur manuell) und **Personen pro Wochentag**; Nr. = Wichtigkeit; «Auffüllen mit» (KSV2); ➕ Neue Aufgabe; 🎨 eigene Farbe |
| **📌 Feste Einsätze** | fixe Einsätze und «alle 2 Wochen frei» | dito, plus **Zusatz** (z. B. Arbeitszeit «12-16» im Wochenwechsel) |
| **🔄 Rotationen** | TV, 18-Uhr-Liste, Labor, Schalter-Gruppe, Wochenaufgaben, Vorschau 6 Wochen | – (Verteilung automatisch fair) |
| **⚖️ Regeln** | Besetzungszahlen und Entscheidungen (TEL, ABKL, PO, Schalter, HUB, RX Abo, Scanning) | – (steckt in «Aufgaben & Besetzung») |
| **🗂️ BackUp-Team** | Zeilen unterhalb des Teams mit Standard-Eintrag (SPP-Team: `ERP`) | dito (zurzeit leer) |

Typische Fälle KSV:

- **Neue Aufgabe** (so wie Redcare im August): Aufgaben & Besetzung → «➕ Neue Aufgabe»
  → Code, Farbe, ganzer Tag / Halbtag, Personen pro Tag → Speichern.
- **Montag zu zweit contact**: Aufgaben & Besetzung → Spalte «Mo» auf 2.
- **Wer macht was**: Wer darf was → ja / Reserve / nie.
- **Teilzeit**: Personen → Mo–Fr (ganz / VM / NM / frei).

Kurznamen sind nur ein interner Schlüssel (Rotationen, Abwesenheiten): Die App zeigt
überall den Namen und vergibt den Kurznamen für neue Personen automatisch.

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

   «💾 Speichern» schreibt dann die Dateien des **aktuellen Teams** ins Repo
   (KSC: `team.json` + `abwesenheiten.json`, KSV: `team_ksv.json` +
   `abwesenheiten_ksv.json`); Streamlit Cloud lädt die App danach neu.
2. **Ohne Token**: «💾 Speichern» schreibt nur auf den Server – Änderungen bleiben
   bis zum nächsten Neustart der App. Dauerhaft nur über das Repo.

Lokal (eigener Rechner / Docker) wird immer in die Dateien neben der App geschrieben.
Pfade lassen sich übersteuern: `KSC_TEAM_CONFIG`, `KSC_ABSENCES`, `KSV_TEAM_CONFIG`,
`KSV_ABSENCES`; `ARBEITSPLAN_TEAM=KSV` stellt die App fest auf ein Team ein.

---

## Regeln

- **KSC**: vollständige Übersicht in **REGELN.md** (Version 7.0, unverändert).
- **KSV**: **REGELN_KSV.md** – abgeleitet aus dem Excel-Plan 2026 (KW 4–44):
  pro Tag 1× `contact` ganztags (Montag 2×), 1× `KSV1` ganztags, 1× `Redcare` ganztags,
  pro Halbtag 1× `KD Komm`, Rest `KSV2`; Pia `Quality` Di + Do VM; Kim nie KSV1;
  Melanie contact nur als Reserve; Nicht-Arbeitstage «Frei».

Kurzfassung KSC (Planungsreihenfolge):

1. Abwesenheiten / Sonderwünsche
2. Fixeinsätze aus `team.json` (z. B. Brigitte Mo ERF9, Dragi Di ERF9 + ERF9/TEL,
   Dipiga KGS Di VM + Do NM, Maria B. Mo VM HO, Florence Mi VM HO, Linda Di HO/Q alle 2 Wochen)
3. Labor Mi NM (Jesika / Dipiga im Wechsel)
4. Aushilfe: täglich PO/SCAN (VM) + scanning (NM) · nur-TEL-Personen: jeder Halbtag TEL
5. Tagesverantwortung: dieselbe TL VM + NM, ohne Ämtli → **TL/Q**
6. PHC-/18-Uhr-Liste (Mo TL-Wechsel, Di Dipiga, Mi/Do Pool ohne Corinne/Andrea G.,
   Fr Corinne/Andrea G. im Wechsel; nie Brigitte, Saskia, Maria B.; PHC-Person scannt nicht)
7. Telefon-Minimum der TL vorab reservieren (Standard 1 Halbtag pro Woche)
8. Schalter ganztags (ERF7/SCH), 1 Person pro Tag, 2 Gruppen im Wochenwechsel;
   Reserve-Personen (Brigitte, Saskia, Dragi) nur, wenn aus der Gruppe niemand frei ist
9. TAGES PA ganztags, 1 Person pro Tag, wechselnd (niemand fix, keine TL, nicht Corinne)
10. ERF9 Mi–Fr ganztags, wechselnd
11. RX Abo (2 Halbtage pro Woche, Pool Lara / Linda / Isaura)
12. Postöffnung: ersatzweise PO/SCAN, wenn die Aushilfe fehlt (Di bevorzugt Florence),
    plus 2 rotierende «PO»-Personen am Vormittag – niemand jeden Tag
13. ERF8 (1 Person, Nachmittag)
14. TEL (Mo VM 4, sonst 3) – nur halbtags pro Person, TL bevorzugt (max. 2/Woche),
    Nachmittag TAGES PA/TEL je nach Entscheidung (bei Bedarf / immer / nie)
15. ABKL (2 pro Halbtag) – nur halbtags pro Person
16. HUB: 1 Person VM + 1 andere NM (ERF7/HUB)
17. Rest: TL → ERF7/Q, sonst ERF7; gelegentlich ein halber Tag scanning (max. 2 pro Woche)
18. Wochenaufgaben Direkt / ONB / BTM: je 1 Person, nie dieselbe doppelt,
    ganzwöchig Abwesende übersprungen

«PO/TEL» am Nachmittag wird **nie automatisch** vergeben (Team: nur im Notfall,
weil man früher vom Telefon weg muss); manuell gesetzt erscheint ein Hinweis.

---

## Was ist neu in Version 8.0 (Oktober 2026)?

- **KSV-Team** in derselben App: Team-Umschalter oben links, Direktlink `?team=KSV`,
  getrennte Daten, optional fest auf ein Team (`ARBEITSPLAN_TEAM`).
- **Konfigurierbarer KSV-Generator** (`teamplaner.py`) mit Regeln aus dem Plan 2026,
  fairer Verteilung (Teilzeit kommt an ihren Tagen zuerst dran), Reserve-Stufe,
  Zusatztexten (Arbeitszeiten) und Excel im Jahresplan-Format.
- **Titel** «Arbeitsplan KSC» / «Arbeitsplan KSV» (Browser-Tab und Kopfzeile)
  statt «KSC Arbeitsplan».
- **Oberfläche**: Kopfzeile in Teamfarbe und zwei Zeilen (oben Team, Woche, «Plan generieren»;
  darunter die Schritte) – passt auch auf Laptops; Wochenauswahl als Liste (laufende Woche bis
  ein Jahr im Voraus); «So geht's»-Leiste in Schritt 1;
  Formular mit sichtbaren Optionen statt Dropdowns und Vorschau «→ wird eingetragen»;
  «📌 Aufgabe setzen» mit Aufgabenliste statt Task-Code in der Notiz; Einträge nach
  «diese Woche» / «andere Wochen» gruppiert; Datum im Plankopf; «Neu generieren»
  und Download direkt beim Plan; Plan-Editor mit Freitext.
- Tests für KSV und für die Oberfläche (beide Teams).

---

## Offene Punkte – Annahmen, die die Teams bestätigen sollten

**KSV**: siehe Tabelle «Offene Punkte» in **REGELN_KSV.md** (u. a. Pensen, Redcare
ganztags oder halbtags, GH, Melanie/Medine als Reserve).

**KSC** (unverändert seit Version 7.0):

| # | Thema | Annahme in Version 7.0 | Wo anpassen |
|---|---|---|---|
| 1 | Pensen Florence (neu Fr anwesend), Manuela, Maria F. | 90 / 80 / 80 % (aus den Halbtagen abgeleitet) | Team & Regeln → Personen |
| 2 | Postöffnung am Vormittag | Aushilfe PO/SCAN **plus 2** rotierende PO-Personen (wie seit Mai) | Regeln → «Zusatz-PO am VM» |
| 3 | «PÖ und TEL am Nachmittag nur Notfall» | gelesen als: die Kombination **PO/TEL** am Nachmittag wird nie automatisch geplant | – |
| 4 | RX Abo | 2 Halbtage pro Woche, max. 1 pro Tag, Pool Lara / Linda / Isaura | Regeln → «RX Abo / Woche» (0 = aus) |
| 5 | Scanning Stammteam | ca. jeder dritte freie Halbtag bei Scanning-Berechtigten, max. 2 pro Woche | «Scanning-Anteil» / «max. Halbtage» |
| 6 | TAGES PA + TEL | TAGES PA/TEL nur, wenn die TEL-Besetzung am NM sonst nicht reicht | Regeln |
| 7 | HUB | nur ERF7/HUB (nicht ERF8/HUB), 1 VM + 1 NM | «HUB / Tag» |
| 8 | Schalter | ganztags, 1 Person pro Tag, Pool in 2 Gruppen; Brigitte, Saskia, Dragi als **Reserve**; Linda, Florence, Corinne, Maria B., Andrea A. nie | Regeln / Wer darf was |
| 8b | Schalter-Code | Plan zeigt ERF7/SCH (seit Mai); April-Liste nannte ERF4/SCH | Regeln |
| 8c | TL am Telefon | mind. 1, max. 2 Halbtage pro Woche; Abweichung = Hinweis | Regeln |
| 9 | Bezeichnung TL/Q | für die TL mit Tagesverantwortung | Farbe/Code in `arbeitskalender.py` |
| 10 | 2-Wochen-Rhythmus | gerade KW = Linda / Isaura / Corinne frei; Linda HO/Q in ungeraden KW. 2026 hat 53 KW → Parität kippt an Neujahr; dann einmal «alle 2 Wochen frei» umstellen | Feste Einsätze |
| 11 | VBZ (Vorbezüge) | wird nicht automatisch geplant (nur als Rolle hinterlegt) | – |

Hinweis zum 2-Wochen-Rhythmus: Das gilt auch für KSV (Medines Arbeitszeit am Mi NM
wechselt nach gerader/ungerader KW) – nach Neujahr 2027 einmal prüfen.

---

## Anonymisierung

Alle Anzeigenamen sind «Vorname N.» (`name` in `team.json` / `team_ksv.json`).
Die KSV-Vorlage `Arbeitsplan_neu 2026_anonym.xlsx` ist per `.gitignore`
ausgeschlossen. Die Git-Historie vor Oktober 2026 enthält noch alte Dateien mit
vollen Namen – wer das Repo öffentlich macht, sollte die Historie bereinigen
(z. B. `git filter-repo`) oder ein frisches Repo anlegen.
