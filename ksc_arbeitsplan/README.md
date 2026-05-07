# KSC Arbeitsplan – Web-App

Wöchentlicher Arbeitsplan-Generator für das KSC-Team (18 Personen).
Erzeugt farbcodierte Excel-Dateien nach dem Muster der bestehenden Vorlage,
hält alle harten Regeln ein und zeigt Konflikte transparent an.

**Neu in v6.0:** Web-Oberfläche statt Desktop-GUI, vollständig containerisiert,
läuft auf jedem Rechner mit Docker.

---

## Schnellstart

### 1. Mit Docker Compose (empfohlen)

```bash
docker compose up --build -d
```

Dann im Browser öffnen: **http://localhost:8000**

Fertig. Die generierten Excel-Dateien landen im Ordner `./data/output/`.

### 2. Container stoppen

```bash
docker compose down
```

---

## Voraussetzungen

Nur **eine** Sache muss installiert sein:

- **Docker Desktop** (Windows / macOS) oder **Docker Engine** (Linux)
  Download: https://www.docker.com/products/docker-desktop/

Kein Python, kein pip, keine Bibliotheken, kein Tkinter — alles steckt im Container.

---

## Ordner­struktur

```
ksc_arbeitsplan/
├── app/
│   ├── arbeitskalender.py      ← Kernlogik (Scheduler + Excel-Export)
│   ├── server.py               ← FastAPI-Web-Backend
│   ├── templates/
│   │   └── index.html          ← UI
│   └── static/
│       ├── styles.css
│       └── app.js
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── .dockerignore
├── README.md
└── data/                       ← wird beim ersten Start automatisch erzeugt
    ├── output/                   └─ enthält die Excel-Dateien
    └── scheduler_state.json        └─ State (2-Wochen-Rhythmen etc.)
```

---

## Ohne Docker Compose (nur mit Docker)

Falls du lieber ohne Compose arbeitest:

```bash
# 1. Image bauen
docker build -t ksc-arbeitsplan .

# 2. Container starten
docker run -d \
  --name ksc-arbeitsplan \
  -p 8000:8000 \
  -v "$(pwd)/data:/data" \
  --restart unless-stopped \
  ksc-arbeitsplan

# 3. Logs anschauen
docker logs -f ksc-arbeitsplan

# 4. Stoppen
docker stop ksc-arbeitsplan
docker rm ksc-arbeitsplan
```

### Auf Windows (PowerShell)

```powershell
docker build -t ksc-arbeitsplan .

docker run -d `
  --name ksc-arbeitsplan `
  -p 8000:8000 `
  -v "${PWD}\data:/data" `
  --restart unless-stopped `
  ksc-arbeitsplan
```

---

## Image auf anderen Rechner übertragen

Du kannst das fertige Image exportieren und z.B. per USB-Stick übertragen:

```bash
# Auf dem Entwickler-Rechner:
docker save ksc-arbeitsplan:latest -o ksc-arbeitsplan.tar

# Auf dem Ziel-Rechner:
docker load -i ksc-arbeitsplan.tar
docker run -d -p 8000:8000 -v "$(pwd)/data:/data" ksc-arbeitsplan
```

Oder in ein Registry pushen (Docker Hub, GitHub Container Registry):

```bash
docker tag ksc-arbeitsplan:latest <deinuser>/ksc-arbeitsplan:latest
docker push <deinuser>/ksc-arbeitsplan:latest
```

Auf anderem Rechner:

```bash
docker pull <deinuser>/ksc-arbeitsplan:latest
docker run -d -p 8000:8000 -v "$(pwd)/data:/data" <deinuser>/ksc-arbeitsplan
```

---

## Konfiguration (Umgebungsvariablen)

| Variable           | Default                           | Zweck                                    |
|--------------------|-----------------------------------|------------------------------------------|
| `KSC_HOST`         | `0.0.0.0`                         | Bind-Adresse                             |
| `KSC_PORT`         | `8000`                            | Port im Container                        |
| `KSC_OUTPUT_DIR`   | `/data/output`                    | wo Excel-Dateien abgelegt werden         |
| `KSC_STATE_FILE`   | `/data/scheduler_state.json`      | State für 2-Wochen-Rhythmen              |
| `TZ`               | `Europe/Zurich`                   | Zeitzone                                 |

Beispiel: Port 9090 statt 8000

```bash
docker run -d -p 9090:8000 -v "$(pwd)/data:/data" ksc-arbeitsplan
```

---

## Nutzung der Web-App

1. **Kalenderwoche wählen** – Pfeile oben rechts, oder direkt „nächste Woche"
2. **Abwesenheiten erfassen** (optional)
   - Mitarbeiterin + Grund (Krank / Ferien / Termin / HO / Custom)
   - Tage anklicken
   - Halbtag wählen
   - Notiz optional
   - „Hinzufügen"
3. **„Arbeitsplan generieren"** drücken
4. Ergebnis prüfen:
   - ✅ Besetzung (TEL/ABKL) pro Halbtag
   - Tagesverantwortung, PHC-Liste
   - Wochenaufgaben (Direkt / ONB / BTM)
   - Konflikte werden klar mit „⚠️ OFFEN" markiert
5. **„Excel herunterladen"** – die Datei liegt auch unter `./data/output/`

---

## Logs & Debug

```bash
# Live-Logs
docker compose logs -f

# Nur die letzten 50 Zeilen
docker compose logs --tail 50

# In den laufenden Container reinspringen
docker compose exec ksc-arbeitsplan bash
```

---

## Fehlerdiagnose

| Problem                                      | Lösung                                                                            |
|---------------------------------------------|-----------------------------------------------------------------------------------|
| `Port 8000 already in use`                  | Anderen Port mappen: `-p 8080:8000`                                               |
| Excel erscheint nicht im `./data/output/`   | Volume-Mount prüfen, Rechte am Ordner `./data` prüfen                             |
| `Permission denied` unter Linux             | `sudo chown -R 1001:1001 ./data` (Container-User ist UID 1001)                    |
| Container startet nicht                     | `docker compose logs` zeigt die Ursache                                           |
| Seite lädt nicht                            | Firewall prüfen, `http://localhost:8000` (nicht https)                            |

---

## Was wurde gegenüber v5.1 verbessert?

- **Tkinter-GUI entfernt** → Ersatz durch professionelle Web-Oberfläche (FastAPI + HTML/CSS/JS)
- **Plattform-unabhängig** → läuft auf Windows, macOS, Linux, sogar Raspberry Pi
- **Docker-ready** → eine einzige Build-Umgebung, identisches Verhalten überall
- **Bessere UX** → sichtbare Vorschau der Wochenplans direkt im Browser
- **Mehrbenutzer-fähig** → mehrere Personen können gleichzeitig aus dem Netzwerk zugreifen
- **State persistent** → 2-Wochen-Rhythmen überleben Container-Neustarts

Die Kernlogik aus `arbeitskalender.py` wurde **nicht verändert** – alle bewährten
Regeln, Rotationen und der Excel-Output arbeiten exakt wie vorher.
