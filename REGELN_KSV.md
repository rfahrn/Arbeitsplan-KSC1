# Regeln Arbeitsplan KSV (Version 1.1, Oktober 2026)

> Abgeleitet aus dem bisherigen Excel-Plan «Arbeitsplan_neu 2026_anonym.xlsx»
> (KW 4–44 ausgefüllt, 40 Wochen). `teamplaner.py` setzt die Regeln um, die
> Werte stehen in `team_ksv.json` und sind in der App unter **⚙ Team & Regeln**
> (Team **KSV**) ohne Code änderbar. Rotationen und Verteilung ergeben sich aus
> der Kalenderwoche – dieselbe KW ergibt immer denselben Plan.

---

## 1) Aufgaben pro Tag (automatisch geplant)

Reihenfolge = Wichtigkeit: fehlen Leute, wird Nr. 1 zuerst besetzt.

| Nr. | Code | Dauer | Mo | Di | Mi | Do | Fr | Beobachtung im Plan 2026 |
|---:|---|---|:-:|:-:|:-:|:-:|:-:|---|
| 1 | `contact` | ganzer Tag | **2** | 1 | 1 | 1 | 1 | fast immer genau 1 Person pro Halbtag (85–98 %), 196 von 199 contact-Tagen ganztags; seit KW 42 montags zu zweit («contact 1» / «contact 2») |
| 2 | `KSV1` | ganzer Tag | 1 | 1 | 1 | 1 | 1 | meist genau 1 Person pro Halbtag (72–95 %), 174 von 202 Tagen ganztags |
| 3 | `Redcare` | ganzer Tag | 1 | 1 | 1 | 1 | 1 | seit KW 32; KW 42–44 fast immer 1 Person pro Tag, ganztags |
| 4 | `KD Komm` | Halbtag | 1 | 1 | 1 | 1 | 1 | genau 1 Person pro Halbtag (82–98 %); VM und NM oft verschiedene Personen |
| – | `KSV2` | Auffüllen | | | | | | alle übrigen Halbtage (1–3 Personen pro Halbtag) |

- **ganzer Tag**: dieselbe Person VM + NM. Ist niemand ganztags frei, übernehmen
  zwei Personen je einen Halbtag (z. B. Mi: Kim nur VM, Medine nur NM).
- **Halbtag**: Vormittag und Nachmittag werden einzeln besetzt; bevorzugt Personen,
  deren anderer Halbtag schon belegt ist (so bleiben ganze Tage frei).
- Mehrere Personen am selben Halbtag werden nummeriert: `contact 1`, `contact 2`.

## 2) Personen und «Wer darf was»

| Person | Pensum | Anwesenheit | ✅ ja | 🟨 Reserve | Fest |
|---|---:|---|---|---|---|
| Medine A. | 50 %* | Mo + Di ganz, Mi NM | contact, KSV1, Redcare | KD Komm | Mi NM Arbeitszeit «12-16» (ungerade KW) / «13-17» (gerade KW) |
| Joy B. | 100 % | Mo–Fr | contact, KSV1, Redcare, KD Komm | | |
| Jana | 100 % | Mo–Fr | contact, KSV1, Redcare, KD Komm | | neu seit KW 27 |
| Maria G. | 100 % | Mo–Fr | contact, KSV1, Redcare, KD Komm | | |
| Kim K. | 80 %* | Mi + Fr nur VM | contact, Redcare, KD Komm | | **nie KSV1** (im Plan 2026 nie) |
| Melanie R. | 100 % | Mo–Fr | KSV1, Redcare, KD Komm | contact | |
| Pia S. | 100 % | Mo–Fr | contact, KSV1, Redcare, KD Komm | | **Quality Di VM + Do VM** |

\* Pensum aus der Anwesenheit abgeleitet. **Reserve** = wird nur eingeplant, wenn
niemand mit «ja» frei ist (im Ergebnis steht dann ein ℹ️-Hinweis).

## 3) Weitere Codes (nur manuell) und Abwesenheiten

| Code | Farbe im Excel | Wie in den Plan |
|---|---|---|
| `Quality` | violett | fester Einsatz (Pia) |
| `KTG` (Kein Tagesgeschäft), `Schulung`, `Sitzung`, `HO` | violett | Schritt 1 «📌 Aufgabe setzen» / «🏠 Home Office» oder Plan bearbeiten |
| `Tel`, `Pö KSV` | orange / gelb | Plan bearbeiten (im Plan 2026 nur bis KW 8 bzw. vereinzelt) |
| `Einführung` | hellgelb | Schritt 1 «📌 Aufgabe setzen» (neue Mitarbeitende) |
| `Frei`, `Ferien` | blau | Schritt 1 (🌙 Frei, 🏖️ Ferien) oder 📆 Jahresplan (von–bis); Nicht-Arbeitstage laut Anwesenheit stehen automatisch als «Frei» im Plan |
| `Feiertag` | hellgrau | automatisch an Feiertagen (📆 Jahresplan → Feiertage) – ganzes Team frei, keine Besetzung |
| `KRANK` | grau | Schritt 1 (🤒 Krank) |

Zusätze wie Arbeitszeiten (`KSV2 10.00`) oder Kürzel (`GH`, `RET`, `RDC`, `ID`)
gehen über **Plan bearbeiten → Freitext**; regelmässige Zusätze über
**Feste Einsätze → Zusatz** (so steht Medines Mi-NM-Zeit jede Woche automatisch im Plan).

## 4) Verteilung (fest eingebaut)

Pro Aufgabe und Tag kommt zuerst dran, wer …
1. die Aufgabe diese Woche am wenigsten hatte,
2. sie gestern nicht hatte,
3. seltener da ist (Teilzeit kommt an ihren Tagen zuerst dran – so hat z. B. auch Medine jede Woche einen contact-Tag),
4. insgesamt am wenigsten Aufgaben hat;
5. bei Gleichstand entscheidet ein pro KW fester Zufall.

Ergebnis über 30 Testwochen: jede Person mit «ja» hat genau **einen contact-Tag pro Woche**.

## 5) Regelprüfung nach dem Planen

❌ = eine Aufgabe ist an einem Halbtag nicht voll besetzt oder jemand hat eine
Aufgabe ohne «ja»/«Reserve» · ⚠️ = Ganztags-Aufgabe aufgeteilt, obwohl die Person
den ganzen Tag da ist; Eintrag an einem Halbtag, an dem die Person laut Anwesenheit
frei hat; anwesender Halbtag ohne Aufgabe · ℹ️ = Reserve eingesetzt.

## 6) Excel

Spalten A–L entsprechen genau einem Wochenblock im Jahresplan
(`Woche 43 | 19.10.2026 – 23.10.2026`, Name | Vorname | Mo VM … Fr NM, Zeilenfarbe pro
Person, Farben der Aufgaben) – der Block lässt sich direkt in die Jahresdatei
einfügen. Rechts davon stehen Bemerkungen (Termine), darunter eine Legende.

## 7) Feiertage

Im Plan 2026 waren an **Karfreitag, Ostermontag, Auffahrt und Pfingstmontag** alle
Zellen leer, am **1. Mai** wurde gearbeitet. Voreingestellt sind deshalb die neun
verbreiteten Deutschschweizer Feiertage (Neujahr, Berchtoldstag, Karfreitag,
Ostermontag, Auffahrt, Pfingstmontag, 1. August, Weihnachten, Stephanstag) – einstellbar
im **📆 Jahresplan**, auch halbe Tage und eigene freie Tage. An Feiertagen steht
`Feiertag` im Plan, es wird nichts geplant und keine Besetzung geprüft.

## 8) Offene Punkte – vom KSV-Team zu bestätigen

| # | Thema | Annahme in Version 1.0 | Wo anpassen |
|---|---|---|---|
| 1 | Pensen | Medine 50 %, Kim 80 %, alle anderen 100 % (aus der Anwesenheit abgeleitet) | Personen |
| 2 | Redcare | 1 Person pro Tag, **ganztags** (wie KW 42–44); KW 32–41 oft nur halbtags | Aufgaben & Besetzung → Dauer «Halbtag» |
| 3 | contact am Montag | **2 Personen** (seit KW 42), übrige Tage 1 | Aufgaben & Besetzung → Spalte Mo |
| 4 | GH | bis KW 41 hatte fast täglich eine Person «… GH» (meist Melanie, Joy, Kim, Pia); seit KW 42 nicht mehr im Plan → **nicht automatisch** | bei Bedarf: neue Aufgabe oder Freitext |
| 5 | Kürzel RET / RDC / RBK / ID / TC / Ret/GS | nur als Freitext | Plan bearbeiten → Freitext |
| 6 | Melanie contact | nur **Reserve** (im Plan 2026 nur 9 contact-Tage) | Wer darf was |
| 7 | Medine KD Komm | nur **Reserve** (im Plan 2026 fast nie) | Wer darf was |
| 8 | Medine Arbeitszeit Mi NM | «12-16» ungerade KW, «13-17» gerade KW (wie in KW 35–53 eingetragen) | Feste Einsätze → Zusatz |
| 9 | Jana | Nachname-Initial fehlt in der Vorlage | Personen → Name |
| 10 | Tiziana S. | seit KW 9 nicht mehr im Plan → nicht übernommen | Personen → ➕ Neue Person |
| 11 | Bedeutung von contact / KSV1 / KSV2 | nur die Kürzel übernommen | Aufgaben & Besetzung → Bedeutung |
| 12 | «Frei» an Nicht-Arbeitstagen | wird eingetragen (wie meist im Plan 2026) | Aufgaben & Besetzung → Häkchen |
| 13 | Heiligabend / Silvester | normale Arbeitstage (im Plan 2026 noch nicht ausgefüllt) – falls nachmittags frei: «Nachmittag» wählen | 📆 Jahresplan → Feiertage |
| 14 | Fronleichnam | Arbeitstag; 2026 waren die meisten ab 12 Uhr frei, zwei Personen bis 14 Uhr | 📆 Jahresplan → Feiertage («Nachmittag») oder pro Person «Frei» |
