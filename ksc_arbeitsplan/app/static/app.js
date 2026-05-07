/* ═══════════════════════════════════════════════════════════
   KSC Arbeitsplan — Frontend-Logik
   ═══════════════════════════════════════════════════════════ */

const state = {
  weekOffset: 0,
  selectedDays: new Set(),
  entries: [],
};

const DAY_LABELS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"];
const DAY_SHORTS = ["Mo", "Di", "Mi", "Do", "Fr"];

// Task → CSS-Variable für Einfärbung im Schedule-Preview
const TASK_STYLE = {
  "KTG": "var(--t-ktg)",
  "KGS": "var(--t-ktg)",
  "ABKL": "var(--t-abkl)",
  "TEL": "var(--t-tel)",
  "ERF7": "var(--t-erf7)",
  "ERF7/Q": "var(--t-erf7)",
  "ERF7/HUB": "var(--t-erf7)",
  "ERF8": "var(--t-erf8)",
  "ERF8/Q": "var(--t-erf8)",
  "ERF8/HUB": "var(--t-erf8)",
  "ERF9": "var(--t-erf9)",
  "ERF9/Q": "var(--t-erf9)",
  "ERF9/TEL": "var(--t-erf9)",
  "ERF4/SCH": "var(--t-sch)",
  "ERF5": "var(--t-erf5)",
  "PO": "var(--t-po)",
  "PO/ABKL": "var(--t-abkl)",
  "PO/TEL": "var(--t-tel)",
  "PO/SCAN": "var(--t-ktg)",
  "HO": "var(--t-ho)",
  "HO/Q": "var(--t-ho)",
  "VBZ/Q": "#C5E0B4",
  "KSC Spez.": "var(--t-ksc)",
  "TAGES PA": "var(--t-pa)",
  "TAGESPA": "var(--t-pa)",
  "RX Abo": "var(--t-rx)",
  "KRANK": "var(--t-krank)",
  "FERIEN": "#F2F2F2",
  "scanning": "var(--t-scan)",
};

function styleForTask(task) {
  if (!task) return { bg: "#ffffff", color: "#cbd5e1" };
  if (task.startsWith("*")) return { bg: "#facc15", color: "#1e293b" };
  const bg = TASK_STYLE[task] || (() => {
    for (const key of Object.keys(TASK_STYLE).sort((a, b) => b.length - a.length)) {
      if (task.startsWith(key)) return TASK_STYLE[key];
    }
    return "#ffffff";
  })();
  return { bg, color: "#0b1020" };
}

/* ─── DOM-Helfer ─── */
const $ = (id) => document.getElementById(id);

/* ═════════════ INIT ═════════════ */

document.addEventListener("DOMContentLoaded", () => {
  initDayChips();
  initButtons();
  initNav();
  renderEntries();
});

function initDayChips() {
  document.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const d = parseInt(chip.dataset.day, 10);
      if (state.selectedDays.has(d)) {
        state.selectedDays.delete(d);
        chip.classList.remove("active");
      } else {
        state.selectedDays.add(d);
        chip.classList.add("active");
      }
    });
  });

  $("btn-all-days").addEventListener("click", () => {
    [0, 1, 2, 3, 4].forEach((d) => state.selectedDays.add(d));
    document.querySelectorAll(".chip").forEach((c) => c.classList.add("active"));
  });

  $("btn-no-days").addEventListener("click", () => {
    state.selectedDays.clear();
    document.querySelectorAll(".chip").forEach((c) => c.classList.remove("active"));
  });
}

function initButtons() {
  $("btn-add").addEventListener("click", addEntry);
  $("btn-generate").addEventListener("click", generatePlan);
  $("btn-week-prev").addEventListener("click", () => changeWeek(-1));
  $("btn-week-next").addEventListener("click", () => changeWeek(+1));
}

function initNav() {
  const navItems = document.querySelectorAll(".nav-item");
  navItems.forEach((item) => {
    item.addEventListener("click", (e) => {
      e.preventDefault();
      navItems.forEach((n) => n.classList.remove("active"));
      item.classList.add("active");
      const target = document.querySelector(item.getAttribute("href"));
      if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

/* ═════════════ WOCHEN-NAVIGATION ═════════════ */

async function changeWeek(delta) {
  state.weekOffset += delta;
  if (state.weekOffset < 0) state.weekOffset = 0;

  try {
    const resp = await fetch(`/api/week-info?offset=${state.weekOffset}`);
    const data = await resp.json();
    $("kw-label").textContent = `KW ${data.kw}`;
    $("kw-range").textContent = `${data.monday_str} – ${data.friday_str}`;
    $("stat-kw").textContent = data.kw;
    $("stat-mo").textContent = data.monday_str;
    $("stat-fr").textContent = data.friday_str;
    $("week-type-badge").textContent = data.is_even ? "gerade KW" : "ungerade KW";
    $("info-strip-text").textContent = data.two_week_notes.join(" · ");
  } catch (err) {
    console.error(err);
  }
}

/* ═════════════ EINTRÄGE VERWALTEN ═════════════ */

function addEntry() {
  const name = $("f-employee").value;
  const type = $("f-type").value;
  const slot = $("f-slot").value;
  const note = $("f-note").value.trim();

  if (state.selectedDays.size === 0) {
    flashStatus("Bitte mindestens einen Tag auswählen.", "warn");
    return;
  }

  const days = Array.from(state.selectedDays).sort();
  const slots = slot === "both" ? [0, 1] : [parseInt(slot, 10)];
  const task = type === "Custom" ? (note || "HO") : "";

  state.entries.push({
    name, type, days, slots, note, task,
  });

  // reset form
  state.selectedDays.clear();
  document.querySelectorAll(".chip").forEach((c) => c.classList.remove("active"));
  $("f-note").value = "";

  renderEntries();
}

function removeEntry(idx) {
  state.entries.splice(idx, 1);
  renderEntries();
}

function renderEntries() {
  const list = $("entries-list");
  const empty = $("entries-empty");
  const count = $("entries-count");

  list.innerHTML = "";

  if (state.entries.length === 0) {
    empty.classList.remove("hidden");
    count.textContent = "0 Einträge";
    return;
  }

  empty.classList.add("hidden");
  count.textContent = state.entries.length === 1 ? "1 Eintrag" : `${state.entries.length} Einträge`;

  state.entries.forEach((e, idx) => {
    const div = document.createElement("div");
    div.className = "entry-item";

    const typeKey = typeToKey(e.type);
    const daysStr = e.days.map((d) => DAY_SHORTS[d]).join(", ");
    const slotStr = e.slots.length === 2 ? "ganzer Tag"
                  : e.slots[0] === 0 ? "Vormittag" : "Nachmittag";

    div.innerHTML = `
      <span class="entry-dot ${typeKey}"></span>
      <span class="entry-text">
        <span class="entry-name">${escapeHtml(e.name)}</span>
        <span class="entry-meta">· ${escapeHtml(e.type)} · ${daysStr} (${slotStr})</span>
        ${e.note ? `<span class="entry-note">· „${escapeHtml(e.note)}"</span>` : ""}
      </span>
      <button class="entry-remove" title="Entfernen">×</button>
    `;
    div.querySelector(".entry-remove").addEventListener("click", () => removeEntry(idx));
    list.appendChild(div);
  });
}

function typeToKey(type) {
  return { "Krank": "KRANK", "Ferien": "FERIEN", "Termin": "TERMIN",
           "Home Office": "HO", "Custom": "CUSTOM" }[type] || "CUSTOM";
}

/* ═════════════ PLAN GENERIEREN ═════════════ */

async function generatePlan() {
  showOverlay(true);
  setStatus("Generiere …", "info");

  const payload = {
    week_offset: state.weekOffset,
    entries: state.entries,
  };

  try {
    const resp = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!resp.ok) {
      throw new Error(`HTTP ${resp.status}`);
    }

    const data = await resp.json();
    renderResult(data);
    setStatus(`✓ ${data.filename}`, "ok");
  } catch (err) {
    console.error(err);
    setStatus("Fehler beim Generieren: " + err.message, "err");
  } finally {
    showOverlay(false);
  }
}

function renderResult(data) {
  const card = $("section-result");
  card.classList.remove("hidden");

  // Download-Link
  const dl = $("download-link");
  dl.href = data.download_url;
  dl.download = data.filename;
  dl.classList.remove("hidden");

  // Warnungen
  renderBanner("result-warnings", data.warnings,
    "Warnungen vor dem Generieren (Regel-Impact)");

  // Issues
  renderBanner("result-issues", data.issues,
    "Konflikte / Regeln nicht vollständig einhaltbar");

  // Success
  if (data.issues.length === 0) {
    $("result-success").classList.remove("hidden");
  } else {
    $("result-success").classList.add("hidden");
  }

  // Coverage
  const covGrid = $("coverage-grid");
  covGrid.innerHTML = "";
  data.coverage.forEach((c) => {
    const card = document.createElement("div");
    card.className = "cov-card";
    card.innerHTML = `
      <div class="cov-head">
        <span class="cov-day">${c.day}</span>
        <span>${c.slot}</span>
      </div>
      <div class="cov-row">
        <span>TEL</span>
        <span class="cov-badge ${c.tel_ok ? 'cov-ok' : 'cov-bad'}">${c.tel}/${c.tel_target}</span>
      </div>
      <div class="cov-row">
        <span>ABKL</span>
        <span class="cov-badge ${c.abkl_ok ? 'cov-ok' : 'cov-bad'}">${c.abkl}/${c.abkl_target}</span>
      </div>
    `;
    covGrid.appendChild(card);
  });

  // Tagesverantwortung
  renderPills("tv-list", data.tagesverantwortung);
  renderPills("phc-list", data.phc_liste);

  // Wochenaufgaben
  const wa = $("wa-grid");
  wa.innerHTML = "";
  const waEntries = [
    ["Direktbestellung", data.wochenaufgaben.direkt],
    ["ONB", data.wochenaufgaben.onb],
    ["BTM", data.wochenaufgaben.btm],
  ];
  waEntries.forEach(([lbl, val]) => {
    const card = document.createElement("div");
    card.className = "wa-card";
    card.innerHTML = `<div class="wa-label">${lbl}</div><div class="wa-val">${escapeHtml(val || "–")}</div>`;
    wa.appendChild(card);
  });

  // Schedule-Tabelle
  renderScheduleTable(data);

  card.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderPills(id, items) {
  const el = $(id);
  el.innerHTML = "";
  items.forEach((item) => {
    const pill = document.createElement("div");
    pill.className = "pill";
    const isOpen = item.person.includes("OFFEN");
    pill.innerHTML = `
      <span class="pill-day">${item.day}</span>
      <span class="pill-val ${isOpen ? 'open' : ''}">${escapeHtml(item.person)}</span>
    `;
    el.appendChild(pill);
  });
}

function renderBanner(id, items, title) {
  const el = $(id);
  if (!items || items.length === 0) {
    el.classList.add("hidden");
    return;
  }
  el.classList.remove("hidden");
  el.innerHTML = `
    <h4>${title}</h4>
    <ul>${items.map((t) => `<li>${escapeHtml(t)}</li>`).join("")}</ul>
  `;
}

function renderScheduleTable(data) {
  const t = $("schedule-table");
  t.innerHTML = "";

  // Header-Row 1: Tage
  const thead = document.createElement("thead");
  const tr1 = document.createElement("tr");
  tr1.innerHTML = `<th class="name-col" rowspan="2">Name</th><th class="pct-col" rowspan="2">%</th>`;
  data.days.forEach((d) => {
    const th = document.createElement("th");
    th.className = "day-col";
    th.colSpan = 2;
    th.textContent = d;
    tr1.appendChild(th);
  });
  thead.appendChild(tr1);

  // Header-Row 2: VM/NM
  const tr2 = document.createElement("tr");
  data.days.forEach(() => {
    tr2.innerHTML += `<th>VM</th><th>NM</th>`;
  });
  thead.appendChild(tr2);
  t.appendChild(thead);

  // Body
  const tbody = document.createElement("tbody");
  data.rows.forEach((r) => {
    const tr = document.createElement("tr");
    const nameTd = document.createElement("td");
    nameTd.className = "name-col";
    nameTd.textContent = r.name;
    tr.appendChild(nameTd);

    const pctTd = document.createElement("td");
    pctTd.className = "pct-col";
    pctTd.textContent = r.pct + "%";
    tr.appendChild(pctTd);

    r.cells.forEach((c) => {
      const td = document.createElement("td");
      td.textContent = c.task || "";
      if (c.is_phc) {
        td.className = "cell-phc";
      } else {
        const s = styleForTask(c.task);
        td.style.background = s.bg;
        td.style.color = s.color;
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  t.appendChild(tbody);
}

/* ═════════════ UTILITIES ═════════════ */

function showOverlay(show) {
  $("overlay").classList.toggle("hidden", !show);
}

function setStatus(text, level) {
  const el = $("action-status");
  el.textContent = text;
  const colorMap = {
    "ok":   "var(--ok)",
    "warn": "var(--warn)",
    "err":  "var(--err)",
    "info": "var(--text-dim)",
  };
  el.style.color = colorMap[level] || "var(--text-dim)";
}

function flashStatus(text, level) {
  setStatus(text, level);
  setTimeout(() => setStatus("", "info"), 3500);
}

function escapeHtml(str) {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
