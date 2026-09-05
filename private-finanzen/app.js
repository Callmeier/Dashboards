const $ = (id) => document.getElementById(id);

const DB_NAME = "finance-cockpit-local";
const DB_VERSION = 1;
const STORE_NAME = "workbooks";
const WORKBOOK_KEY = "current";

const money = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
  maximumFractionDigits: 0,
});
const number = new Intl.NumberFormat("de-DE", { maximumFractionDigits: 1 });
const percent = new Intl.NumberFormat("de-DE", { style: "percent", maximumFractionDigits: 1 });

let model = null;
let selectedMonthIndex = 0;

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveLocalWorkbook(buffer, meta) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put({ buffer, ...meta }, WORKBOOK_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

async function loadLocalWorkbook() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get(WORKBOOK_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function forgetLocalWorkbook() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).delete(WORKBOOK_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

function asNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value).replace(/\./g, "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : 0;
}

function sheetRows(workbook, name) {
  const ws = workbook.Sheets[name];
  if (!ws) throw new Error(`Das Tabellenblatt „${name}“ fehlt.`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

function mapParameters(rows) {
  const out = new Map();
  rows.slice(1).forEach((row) => {
    if (row[0] !== null && row[0] !== undefined) out.set(String(row[0]).trim(), row[1]);
  });
  return out;
}

function mapMetrics(rows) {
  const months = (rows[0] || []).slice(1, 13).map((m, i) => m || `Monat ${i + 1}`);
  const metrics = new Map();
  rows.slice(1).forEach((row) => {
    if (row[0]) metrics.set(String(row[0]).trim(), row.slice(1, 13).map(asNumber));
  });
  return { months, metrics };
}

function metric(metrics, name, monthIndex) {
  const row = metrics.get(name);
  return row ? asNumber(row[monthIndex]) : 0;
}

function parsePayments(rows) {
  const headerIndex = rows.findIndex((row) => row && row[0] === "ID");
  if (headerIndex < 0) return { activeCount: 0, monthlyEquivalent: 0, wealthBuild: 0 };
  const headers = rows[headerIndex].map((h) => String(h || ""));
  const idx = (name) => headers.indexOf(name);
  const activeIdx = idx("Aktiv");
  const equivalentIdx = idx("Monatsäquivalent EUR");
  const wealthIdx = idx("Vermögensaufbau mtl.");
  const activeRows = rows.slice(headerIndex + 1).filter((row) => row && row[activeIdx] === "Ja");
  return {
    activeCount: activeRows.length,
    monthlyEquivalent: activeRows.reduce((sum, row) => sum + asNumber(row[equivalentIdx]), 0),
    wealthBuild: activeRows.reduce((sum, row) => sum + asNumber(row[wealthIdx]), 0),
  };
}

function buildModel(workbook, meta) {
  const evaluation = mapMetrics(sheetRows(workbook, "04_Monatsauswertung"));
  const params = mapParameters(sheetRows(workbook, "05_Ziele_Parameter"));
  let payments = { activeCount: 0, monthlyEquivalent: 0, wealthBuild: 0 };
  if (workbook.Sheets["11_Regelmaessige_Zahlungen"]) {
    payments = parsePayments(sheetRows(workbook, "11_Regelmaessige_Zahlungen"));
  }
  return { ...evaluation, params, payments, meta };
}

function formatMoney(value) {
  return money.format(asNumber(value));
}

function formatSignedMoney(value) {
  const v = asNumber(value);
  const sign = v > 0 ? "+" : "";
  return `${sign}${money.format(v)}`;
}

function setSignedClass(element, value) {
  element.classList.remove("positive", "negative");
  if (value > 0) element.classList.add("positive");
  if (value < 0) element.classList.add("negative");
}

function getParameter(name, fallback = 0) {
  if (!model) return fallback;
  const raw = model.params.get(name);
  return raw === undefined || raw === null ? fallback : raw;
}

function statusForScore(score) {
  const green = asNumber(getParameter("Min. Finanzscore Grün", 80));
  const yellow = asNumber(getParameter("Min. Finanzscore Gelb", 60));
  if (score >= green) return { label: "stabil", className: "good" };
  if (score >= yellow) return { label: "beobachten", className: "warn" };
  return { label: "prüfen", className: "bad" };
}

function render() {
  if (!model) return;
  const i = selectedMonthIndex;
  const m = model.metrics;

  const assets = metric(m, "Aktiva Gesamt", i);
  const liabilities = metric(m, "Passiva Gesamt", i);
  const netWorth = metric(m, "Nettovermögen", i);
  const mom = metric(m, "Veränderung ggü. Vormonat", i);
  const ytd = metric(m, "YTD-Veränderung", i);
  const score = metric(m, "Finanzscore Gesamt", i);
  const missing = metric(m, "Fehlende Monatsstände", i);
  const liquid = metric(m, "Liquidität frei", i) + metric(m, "Rücklagen", i);
  const reserve = metric(m, "Liquiditätsreserve (Monate)", i);
  const investments = metric(m, "Wertpapiere / Depot", i) + metric(m, "Bitcoin / Krypto", i) + metric(m, "Edelmetalle", i);

  $("netWorth").textContent = formatMoney(netWorth);
  const targetNetWorth = asNumber(getParameter("Ziel-Nettovermögen", 0));
  $("netWorthSub").textContent = targetNetWorth > 0
    ? `${percent.format(netWorth / targetNetWorth)} des langfristigen Ziels`
    : "Nettovermögen";

  $("momChange").textContent = formatSignedMoney(mom);
  setSignedClass($("momChange"), mom);
  const previous = i > 0 ? metric(m, "Nettovermögen", i - 1) : 0;
  const momPct = previous !== 0 ? mom / Math.abs(previous) : null;
  $("momChangePct").textContent = momPct === null ? "kein Vergleichswert" : `${momPct > 0 ? "+" : ""}${percent.format(momPct)}`;
  setSignedClass($("momChangePct"), mom);

  $("ytdChange").textContent = formatSignedMoney(ytd);
  setSignedClass($("ytdChange"), ytd);

  $("financeScore").textContent = `${number.format(score)}/100`;
  const scoreStatus = statusForScore(score);
  $("financeStatus").textContent = scoreStatus.label;
  $("financeStatus").className = `pill ${scoreStatus.className}`;
  $("qualityText").textContent = missing === 0 ? "Monatsdaten vollständig" : `${number.format(missing)} Monatsstände offen`;

  $("liquidity").textContent = formatMoney(liquid);
  $("reserveMonths").textContent = `${number.format(reserve)} Monate`;
  $("investments").textContent = formatMoney(investments);
  $("liabilities").textContent = formatMoney(liabilities);

  renderAlert(missing, assets);
  renderTrend();
  renderAllocation(assets);
  renderHints(mom, missing);
  renderPayments();
}

function renderAlert(missing, assets) {
  const box = $("alertBox");
  if (missing > 0) {
    box.textContent = `Für ${model.months[selectedMonthIndex]} fehlen noch ${number.format(missing)} Monatsstände. Das Dashboard zeigt deshalb möglicherweise unvollständige Werte.`;
    box.classList.remove("hidden");
  } else if (assets === 0) {
    box.textContent = "Für diesen Monat sind noch keine Vermögenswerte vorhanden.";
    box.classList.remove("hidden");
  } else {
    box.classList.add("hidden");
  }
}

function renderAllocation(totalAssets) {
  const container = $("allocation");
  container.replaceChildren();
  const i = selectedMonthIndex;
  const m = model.metrics;
  const items = [
    ["Liquidität + Rücklagen", metric(m, "Liquidität frei", i) + metric(m, "Rücklagen", i)],
    ["Wertpapiere / Depot", metric(m, "Wertpapiere / Depot", i)],
    ["Bitcoin / Krypto", metric(m, "Bitcoin / Krypto", i)],
    ["Edelmetalle", metric(m, "Edelmetalle", i)],
    ["Immobilien", metric(m, "Immobilien", i)],
    ["Business + Sonstige", metric(m, "Business-Liquidität", i) + metric(m, "Sonstige Assets", i)],
  ];

  items.forEach(([label, value]) => {
    const share = totalAssets > 0 ? value / totalAssets : 0;
    const row = document.createElement("div");
    row.className = "alloc-row";

    const head = document.createElement("div");
    head.className = "alloc-head";
    const name = document.createElement("span");
    name.textContent = label;
    const val = document.createElement("strong");
    val.textContent = formatMoney(value);
    head.append(name, val);

    const track = document.createElement("div");
    track.className = "alloc-track";
    const fill = document.createElement("div");
    fill.className = "alloc-fill";
    fill.style.width = `${Math.max(0, Math.min(100, share * 100))}%`;
    track.append(fill);

    const meta = document.createElement("div");
    meta.className = "alloc-meta";
    meta.textContent = percent.format(share);

    row.append(head, track, meta);
    container.append(row);
  });
}

function renderHints(mom, missing) {
  const container = $("assistantHints");
  container.replaceChildren();
  const i = selectedMonthIndex;
  const m = model.metrics;
  const reserve = metric(m, "Liquiditätsreserve (Monate)", i);
  const bitcoinShare = metric(m, "Bitcoin-Anteil an Aktiva", i);
  const propertyShare = metric(m, "Immobilien-Anteil an Aktiva", i);
  const reserveTarget = asNumber(getParameter("Ziel-Notgroschen", 6));
  const maxBitcoin = asNumber(getParameter("Max. Bitcoin-Anteil", 0.1));
  const maxProperty = asNumber(getParameter("Max. Immobilien-Anteil", 0.7));

  const hints = [
    reserve >= reserveTarget
      ? ["good", "Liquiditätsreserve erreicht oder über Ziel."]
      : ["warn", `Liquiditätsreserve unter dem Ziel von ${number.format(reserveTarget)} Monaten.`],
    bitcoinShare <= maxBitcoin
      ? ["good", "Bitcoin-Anteil liegt innerhalb der definierten Schwelle."]
      : ["warn", `Bitcoin-Anteil liegt über der Schwelle von ${percent.format(maxBitcoin)}.`],
    propertyShare <= maxProperty
      ? ["good", "Immobilien-Anteil liegt innerhalb der definierten Schwelle."]
      : ["warn", `Immobilien-Anteil liegt über der Schwelle von ${percent.format(maxProperty)}.`],
    mom > 0
      ? ["good", "Nettovermögen ist gegenüber dem Vormonat gestiegen."]
      : ["warn", "Nettovermögen ist gegenüber dem Vormonat nicht gestiegen oder es fehlt ein Vergleichswert."],
    missing === 0
      ? ["good", "Alle erwarteten Monatsstände sind gepflegt."]
      : ["bad", `${number.format(missing)} erwartete Monatsstände fehlen noch.`],
  ];

  hints.forEach(([type, text]) => {
    const item = document.createElement("div");
    item.className = `hint ${type}`;
    const dot = document.createElement("span");
    dot.className = "hint-dot";
    const p = document.createElement("p");
    p.textContent = text;
    item.append(dot, p);
    container.append(item);
  });
}

function renderPayments() {
  $("activePayments").textContent = number.format(model.payments.activeCount);
  $("monthlyEquivalent").textContent = formatMoney(model.payments.monthlyEquivalent);
  $("monthlyWealthBuild").textContent = formatMoney(model.payments.wealthBuild);
}

function escapeXml(text) {
  return String(text).replace(/[<>&'\"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" }[c]));
}

function renderTrend() {
  const container = $("trendChart");
  const values = model.metrics.get("Nettovermögen") || [];
  const points = values.map(asNumber);
  if (!points.some((v) => v !== 0)) {
    container.innerHTML = '<div class="chart-empty">Noch keine Monatsstände für den Jahresverlauf vorhanden.</div>';
    return;
  }

  const W = 720, H = 250, L = 44, R = 14, T = 14, B = 32;
  const plotW = W - L - R, plotH = H - T - B;
  let min = Math.min(...points), max = Math.max(...points);
  if (min === max) { min -= 1; max += 1; }
  const pad = (max - min) * 0.12;
  min -= pad; max += pad;
  const x = (idx) => L + (idx / Math.max(1, points.length - 1)) * plotW;
  const y = (val) => T + ((max - val) / (max - min)) * plotH;
  const line = points.map((v, idx) => `${idx === 0 ? "M" : "L"}${x(idx).toFixed(1)},${y(v).toFixed(1)}`).join(" ");
  const area = `${line} L${x(points.length - 1).toFixed(1)},${(H - B).toFixed(1)} L${x(0).toFixed(1)},${(H - B).toFixed(1)} Z`;

  const grids = Array.from({ length: 4 }, (_, g) => {
    const yy = T + (g / 3) * plotH;
    const value = max - (g / 3) * (max - min);
    return `<line class="grid-line" x1="${L}" y1="${yy}" x2="${W - R}" y2="${yy}"/><text class="axis-label" x="0" y="${yy + 3}">${escapeXml(money.format(value).replace(/,00/g, ""))}</text>`;
  }).join("");

  const labels = model.months.map((label, idx) => {
    if (idx % 2 !== 0 && idx !== model.months.length - 1) return "";
    return `<text class="axis-label" text-anchor="middle" x="${x(idx)}" y="${H - 8}">${escapeXml(String(label).slice(0, 3))}</text>`;
  }).join("");

  const dots = points.map((v, idx) => idx === selectedMonthIndex
    ? `<circle class="trend-dot" cx="${x(idx)}" cy="${y(v)}" r="5"/>`
    : "").join("");

  container.innerHTML = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true">
    <defs><linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#111827" stop-opacity=".14"/><stop offset="100%" stop-color="#111827" stop-opacity="0"/></linearGradient></defs>
    ${grids}<path class="trend-area" d="${area}"/><path class="trend-line" d="${line}"/>${dots}${labels}
  </svg>`;
}

function fillMonthSelector() {
  const select = $("monthSelect");
  select.replaceChildren();
  model.months.forEach((label, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = label;
    select.append(option);
  });
  select.value = String(selectedMonthIndex);
}

function showDashboard() {
  $("welcome").classList.add("hidden");
  $("dashboard").classList.remove("hidden");
  $("fileName").textContent = model.meta.name || "Excel-Datei";
  const d = model.meta.importedAt ? new Date(model.meta.importedAt) : new Date();
  $("fileMeta").textContent = `importiert ${d.toLocaleDateString("de-DE")} · lokal gespeichert`;
  fillMonthSelector();
  render();
}

async function parseAndShow(buffer, meta, persist = true) {
  if (!window.XLSX) throw new Error("Die Excel-Bibliothek konnte nicht geladen werden. Bitte die Seite einmal mit Internetverbindung öffnen.");
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  model = buildModel(workbook, meta);

  const preferred = Number(localStorage.getItem("financeSelectedMonth"));
  const fromExcel = Math.max(0, Math.min(11, asNumber(model.params.get("Auswertungsmonat")) - 1));
  selectedMonthIndex = Number.isInteger(preferred) && preferred >= 0 && preferred < model.months.length ? preferred : fromExcel;

  if (persist) await saveLocalWorkbook(buffer, meta);
  showDashboard();
}

async function importFile(file) {
  if (!file) return;
  const buffer = await file.arrayBuffer();
  const meta = { name: file.name, importedAt: new Date().toISOString(), lastModified: file.lastModified || null };
  await parseAndShow(buffer, meta, true);
}

function showError(error) {
  console.error(error);
  const box = $("alertBox");
  box.textContent = error?.message || "Die Datei konnte nicht gelesen werden.";
  box.classList.remove("hidden");
  $("welcome").classList.add("hidden");
  $("dashboard").classList.remove("hidden");
}

function wireEvents() {
  const openPicker = () => $("fileInput").click();
  $("importButton").addEventListener("click", openPicker);
  $("welcomeImportButton").addEventListener("click", openPicker);
  $("fileInput").addEventListener("change", async (event) => {
    try { await importFile(event.target.files?.[0]); }
    catch (error) { showError(error); }
    finally { event.target.value = ""; }
  });
  $("monthSelect").addEventListener("change", (event) => {
    selectedMonthIndex = Number(event.target.value);
    localStorage.setItem("financeSelectedMonth", String(selectedMonthIndex));
    render();
  });
  $("forgetButton").addEventListener("click", async () => {
    await forgetLocalWorkbook();
    localStorage.removeItem("financeSelectedMonth");
    model = null;
    $("dashboard").classList.add("hidden");
    $("welcome").classList.remove("hidden");
  });

  document.addEventListener("dragover", (event) => event.preventDefault());
  document.addEventListener("drop", async (event) => {
    event.preventDefault();
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    try { await importFile(file); } catch (error) { showError(error); }
  });
}

async function init() {
  wireEvents();
  try {
    const cached = await loadLocalWorkbook();
    if (cached?.buffer) await parseAndShow(cached.buffer, cached, false);
  } catch (error) {
    console.warn("Lokale Datei konnte nicht geladen werden", error);
  }

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch((error) => console.warn("Service Worker", error));
  }
}

init();
