# ═══════════════════════════════════════════════════════════
#  KSC Arbeitsplan — Dockerfile
#  Multi-Stage-Build: kompakt und production-ready
# ═══════════════════════════════════════════════════════════

# ── Stage 1: Build — Abhängigkeiten vorbereiten
FROM python:3.12-slim AS builder

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /build

# Nur requirements.txt kopieren für besseres Layer-Caching
COPY requirements.txt .
RUN pip install --user --no-cache-dir -r requirements.txt


# ── Stage 2: Runtime — schlankes Final-Image
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/home/appuser/.local/bin:$PATH" \
    KSC_HOST=0.0.0.0 \
    KSC_PORT=8000 \
    KSC_OUTPUT_DIR=/data/output \
    KSC_STATE_FILE=/data/scheduler_state.json \
    TZ=Europe/Zurich

# Non-root user für Sicherheit
RUN groupadd --system --gid 1001 appgroup \
    && useradd --system --uid 1001 --gid 1001 --create-home --home-dir /home/appuser appuser

# Installierte Python-Pakete aus dem Build-Stage übernehmen
COPY --from=builder --chown=appuser:appgroup /root/.local /home/appuser/.local

WORKDIR /app

# Anwendungscode
COPY --chown=appuser:appgroup app/ ./

# Datenverzeichnis (Excel-Output + State) — als Volume mountbar
RUN mkdir -p /data/output && chown -R appuser:appgroup /data

USER appuser

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/api/health').read()" || exit 1

CMD ["python", "-m", "uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
