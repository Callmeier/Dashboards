(() => {
  "use strict";

  const CLIENT_ID_KEY = "financeMicrosoftClientId";
  const FILE_KEY = "financeOneDriveFile";
  const SHEET_NAME = "03_Monatsstaende";
  const RANGE_ADDRESS = "A1:V25";
  const SCOPES = ["Files.ReadWrite"];

  const state = {
    msalInstance: null,
    account: null,
    selectedFile: null,
    rows: null,
    monthIndex: 0,
  };

  const $ = (id) => document.getElementById(id);

  function setStatus(message, type = "info") {
    const el = $("syncStatus");
    if (!el) return;
    el.textContent = message;
    el.className = `sync-status ${type}`;
  }

  function currentRedirectUri() {
    return `${window.location.origin}${window.location.pathname}`;
  }

  function getClientId() {
    return (localStorage.getItem(CLIENT_ID_KEY) || "").trim();
  }

  function getSavedFile() {
    try {
      return JSON.parse(localStorage.getItem(FILE_KEY) || "null");
    } catch (_) {
      return null;
    }
  }

  function saveSelectedFile(file) {
    state.selectedFile = file;
    localStorage.setItem(FILE_KEY, JSON.stringify(file));
    renderSelectedFile();
  }

  function renderSelectedFile() {
    const el = $("selectedOneDriveFile");
    if (!el) return;
    el.textContent = state.selectedFile
      ? `Ausgewählt: ${state.selectedFile.name}`
      : "Noch keine OneDrive-Datei ausgewählt.";
  }

  function ensureConfigured() {
    const clientId = getClientId();
    if (!clientId) {
      setStatus("Bitte zuerst die Microsoft App-ID eintragen.", "warn");
      throw new Error("Microsoft App-ID fehlt.");
    }
    if (!window.msal) {
      setStatus("Microsoft-Anmeldung konnte nicht geladen werden.", "bad");
      throw new Error("MSAL wurde nicht geladen.");
    }
    return clientId;
  }

  function ensureMsal() {
    const clientId = ensureConfigured();
    if (state.msalInstance && state.msalInstance.getConfiguration().auth.clientId === clientId) {
      return state.msalInstance;
    }

    state.msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: currentRedirectUri(),
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    });

    const accounts = state.msalInstance.getAllAccounts();
    state.account = accounts[0] || null;
    renderAuthState();
    return state.msalInstance;
  }

  function renderAuthState() {
    const loginButton = $("microsoftLogin");
    const logoutButton = $("microsoftLogout");
    const label = $("microsoftAccount");
    if (!loginButton || !logoutButton || !label) return;

    if (state.account) {
      label.textContent = state.account.username || state.account.name || "Microsoft verbunden";
      loginButton.textContent = "Neu verbinden";
      logoutButton.classList.remove("hidden");
    } else {
      label.textContent = "Nicht verbunden";
      loginButton.textContent = "Microsoft verbinden";
      logoutButton.classList.add("hidden");
    }
  }

  async function login() {
    const instance = ensureMsal();
    const result = await instance.loginPopup({ scopes: SCOPES, prompt: "select_account" });
    state.account = result.account;
    renderAuthState();
    setStatus("Microsoft-Konto verbunden. Jetzt kannst du deine Excel-Datei suchen.", "good");
    return result.account;
  }

  async function getAccessToken() {
    const instance = ensureMsal();
    const accounts = instance.getAllAccounts();
    state.account = state.account || accounts[0] || null;
    if (!state.account) await login();

    try {
      const token = await instance.acquireTokenSilent({ scopes: SCOPES, account: state.account });
      return token.accessToken;
    } catch (_) {
      const token = await instance.acquireTokenPopup({ scopes: SCOPES, account: state.account });
      state.account = token.account || state.account;
      renderAuthState();
      return token.accessToken;
    }
  }

  async function graphFetch(path, options = {}) {
    const token = await getAccessToken();
    const headers = new Headers(options.headers || {});
    headers.set("Authorization", `Bearer ${token}`);
    if (options.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");

    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers,
    });

    if (!response.ok) {
      let details = "";
      try {
        const body = await response.json();
        details = body?.error?.message || "";
      } catch (_) {
        details = await response.text().catch(() => "");
      }
      throw new Error(`Microsoft Graph: ${response.status} ${details}`.trim());
    }
    return response;
  }

  function odataLiteral(value) {
    return String(value).replace(/'/g, "''");
  }

  async function searchFiles() {
    const query = ($("onedriveSearch")?.value || "").trim();
    if (!query) {
      setStatus("Bitte einen Teil des Excel-Dateinamens eingeben.", "warn");
      return;
    }

    setStatus("OneDrive wird durchsucht …");
    const encodedQuery = encodeURIComponent(odataLiteral(query));
    const response = await graphFetch(`/me/drive/root/search(q='${encodedQuery}')?$select=id,name,file,parentReference,lastModifiedDateTime,webUrl&$top=50`);
    const json = await response.json();
    const files = (json.value || []).filter((item) => item.file && /\.xls(x|m)$/i.test(item.name || ""));

    const select = $("onedriveResults");
    select.replaceChildren();
    files.forEach((file) => {
      const option = document.createElement("option");
      option.value = file.id;
      option.textContent = file.name;
      option.dataset.file = JSON.stringify({ id: file.id, name: file.name, webUrl: file.webUrl || null });
      select.append(option);
    });

    if (files.length) {
      select.classList.remove("hidden");
      $("useOneDriveFile").classList.remove("hidden");
      setStatus(`${files.length} passende Excel-Datei${files.length === 1 ? "" : "en"} gefunden.`, "good");
    } else {
      select.classList.add("hidden");
      $("useOneDriveFile").classList.add("hidden");
      setStatus("Keine passende .xlsx- oder .xlsm-Datei gefunden.", "warn");
    }
  }

  function selectSearchResult() {
    const select = $("onedriveResults");
    const option = select?.selectedOptions?.[0];
    if (!option) return;
    saveSelectedFile(JSON.parse(option.dataset.file));
    setStatus("OneDrive-Datei ausgewählt. Jetzt Monatsstände laden.", "good");
  }

  function workbookRangePath(itemId, sheetName, address) {
    const sheet = encodeURIComponent(sheetName);
    const addr = String(address).replace(/'/g, "''");
    return `/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${sheet}/range(address='${addr}')`;
  }

  async function loadMonthEntries() {
    if (!state.selectedFile) {
      setStatus("Bitte zuerst eine Excel-Datei aus OneDrive auswählen.", "warn");
      return;
    }

    state.monthIndex = Number($("entryMonth")?.value || 0);
    setStatus("Monatsstände werden direkt aus Excel geladen …");
    const response = await graphFetch(`${workbookRangePath(state.selectedFile.id, SHEET_NAME, RANGE_ADDRESS)}?$select=values`);
    const json = await response.json();
    const rows = json.values || [];
    if (!rows.length || !rows[0]?.length) throw new Error("Das Tabellenblatt 03_Monatsstaende enthält keine Daten.");

    state.rows = rows;
    renderEntryRows();
    setStatus(`Monatsstände für ${rows[0][10 + state.monthIndex] || "den gewählten Monat"} geladen.`, "good");
  }

  function euroInputValue(value) {
    if (value === null || value === undefined || value === "") return "";
    const number = Number(value);
    return Number.isFinite(number) ? String(number) : "";
  }

  function renderEntryRows() {
    const container = $("entryRows");
    container.replaceChildren();
    const monthCol = 10 + state.monthIndex;
    const accountRows = state.rows.slice(1).filter((row) => row && row[0]);

    accountRows.forEach((row, index) => {
      const excelRow = index + 2;
      const active = String(row[8] || "").trim();
      if (active !== "Ja" && active !== "Optional") return;

      const item = document.createElement("div");
      item.className = "entry-row";

      const info = document.createElement("div");
      info.className = "entry-info";
      const name = document.createElement("strong");
      name.textContent = row[1] || row[0];
      const meta = document.createElement("span");
      meta.textContent = [row[2], row[5], active === "Optional" ? "optional" : null].filter(Boolean).join(" · ");
      const hint = document.createElement("small");
      hint.textContent = row[9] || "Wert in EUR";
      info.append(name, meta, hint);

      const inputWrap = document.createElement("label");
      inputWrap.className = "entry-input-wrap";
      const currency = document.createElement("span");
      currency.textContent = "€";
      const input = document.createElement("input");
      input.type = "number";
      input.inputMode = "decimal";
      input.step = "0.01";
      input.min = "0";
      input.value = euroInputValue(row[monthCol]);
      input.dataset.excelRow = String(excelRow);
      input.dataset.original = input.value;
      input.setAttribute("aria-label", `${row[1] || row[0]} in Euro`);
      inputWrap.append(currency, input);

      item.append(info, inputWrap);
      container.append(item);
    });

    $("saveEntries").classList.remove("hidden");
    $("entrySummary").textContent = `${container.children.length} Positionen für diesen Monat`;
  }

  function columnName(index1Based) {
    let n = index1Based;
    let result = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  }

  async function saveMonthEntries() {
    if (!state.selectedFile || !state.rows) {
      setStatus("Bitte zuerst Monatsstände laden.", "warn");
      return;
    }

    const inputs = Array.from($("entryRows").querySelectorAll("input[data-excel-row]"));
    const changed = inputs.filter((input) => input.value !== input.dataset.original);
    if (!changed.length) {
      setStatus("Es wurden keine Werte geändert.", "info");
      return;
    }

    setStatus(`${changed.length} Änderung${changed.length === 1 ? "" : "en"} werden in Excel gespeichert …`);
    const excelColumn = columnName(11 + state.monthIndex);

    for (const input of changed) {
      const row = Number(input.dataset.excelRow);
      const raw = input.value.trim();
      if (raw === "") continue;
      const value = Number(raw.replace(",", "."));
      if (!Number.isFinite(value) || value < 0) throw new Error(`Ungültiger Wert in Zeile ${row}.`);
      const address = `${excelColumn}${row}`;
      await graphFetch(workbookRangePath(state.selectedFile.id, SHEET_NAME, address), {
        method: "PATCH",
        body: JSON.stringify({ values: [[value]] }),
      });
      input.dataset.original = input.value;
    }

    setStatus("In OneDrive gespeichert. Excel berechnet die Auswertungen neu …", "good");
    await refreshLocalWorkbookFromOneDrive();
  }

  async function refreshLocalWorkbookFromOneDrive() {
    try {
      await new Promise((resolve) => setTimeout(resolve, 700));
      const response = await graphFetch(`/me/drive/items/${encodeURIComponent(state.selectedFile.id)}/content`);
      const buffer = await response.arrayBuffer();
      const db = await openFinanceDb();
      await new Promise((resolve, reject) => {
        const tx = db.transaction("workbooks", "readwrite");
        tx.objectStore("workbooks").put({
          buffer,
          name: state.selectedFile.name,
          importedAt: new Date().toISOString(),
          source: "onedrive",
        }, "current");
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
      setStatus("Gespeichert und lokale Dashboard-Daten aktualisiert. Seite wird neu geladen …", "good");
      setTimeout(() => window.location.reload(), 450);
    } catch (error) {
      console.warn("Lokaler Cache konnte nach OneDrive-Speicherung nicht aktualisiert werden", error);
      setStatus("Excel wurde in OneDrive gespeichert. Die lokale Dashboard-Ansicht konnte noch nicht automatisch aktualisiert werden.", "warn");
    }
  }

  function openFinanceDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open("finance-cockpit-local", 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains("workbooks")) db.createObjectStore("workbooks");
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function logout() {
    if (!state.msalInstance) ensureMsal();
    const account = state.account || state.msalInstance.getAllAccounts()[0];
    if (account) {
      await state.msalInstance.logoutPopup({ account, postLogoutRedirectUri: currentRedirectUri() });
    }
    state.account = null;
    renderAuthState();
    setStatus("Microsoft-Verbindung getrennt.");
  }

  function fillMonths() {
    const select = $("entryMonth");
    if (!select || select.options.length) return;
    const months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    months.forEach((label, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = label;
      select.append(option);
    });
    const dashboardMonth = Number(localStorage.getItem("financeSelectedMonth"));
    select.value = Number.isInteger(dashboardMonth) && dashboardMonth >= 0 && dashboardMonth < 12
      ? String(dashboardMonth)
      : String(new Date().getMonth());
  }

  function wire() {
    $("entryToggle")?.addEventListener("click", () => {
      $("syncPanel")?.classList.toggle("hidden");
      $("syncPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    $("syncClose")?.addEventListener("click", () => $("syncPanel")?.classList.add("hidden"));

    $("saveClientId")?.addEventListener("click", () => {
      const value = ($("clientIdInput")?.value || "").trim();
      if (!value) {
        setStatus("Bitte eine Microsoft App-ID eintragen.", "warn");
        return;
      }
      localStorage.setItem(CLIENT_ID_KEY, value);
      state.msalInstance = null;
      state.account = null;
      try { ensureMsal(); } catch (_) {}
      setStatus("App-ID lokal gespeichert. Sie wird nicht an GitHub übertragen.", "good");
    });

    $("microsoftLogin")?.addEventListener("click", () => login().catch((error) => setStatus(error.message, "bad")));
    $("microsoftLogout")?.addEventListener("click", () => logout().catch((error) => setStatus(error.message, "bad")));
    $("searchOneDrive")?.addEventListener("click", () => searchFiles().catch((error) => setStatus(error.message, "bad")));
    $("useOneDriveFile")?.addEventListener("click", selectSearchResult);
    $("loadEntries")?.addEventListener("click", () => loadMonthEntries().catch((error) => setStatus(error.message, "bad")));
    $("entryMonth")?.addEventListener("change", () => {
      state.monthIndex = Number($("entryMonth").value);
      state.rows = null;
      $("entryRows")?.replaceChildren();
      $("saveEntries")?.classList.add("hidden");
      $("entrySummary").textContent = "Monat geändert – Werte bitte neu laden.";
    });
    $("saveEntries")?.addEventListener("click", () => saveMonthEntries().catch((error) => setStatus(error.message, "bad")));
  }

  function init() {
    fillMonths();
    state.selectedFile = getSavedFile();
    renderSelectedFile();
    const clientId = getClientId();
    if ($("clientIdInput")) $("clientIdInput").value = clientId;
    if ($("redirectUriValue")) $("redirectUriValue").value = currentRedirectUri();
    if (clientId && window.msal) {
      try { ensureMsal(); } catch (_) {}
    } else {
      renderAuthState();
    }
    wire();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init, { once: true });
  } else {
    init();
  }
})();
