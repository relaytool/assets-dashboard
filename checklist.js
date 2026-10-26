(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const SHEET_NAME = CONFIG.SHIFT_CHECKLIST_SHEET_NAME || "Shift Checklist";
  const CHECKLIST_HEADERS = [
    "Shift ID", "Shift", "Shift Start", "Shift End", "Checklist Item", "Checked",
    "Item Note", "Shift Handover Note", "Updated At", "Closed At", "Closed By", "Shift Status"
  ];
  const CHECKLIST_ITEMS = [
    "Verify physical warehouse stock against the dashboard inventory.",
    "Update the Available to Send counts for this shift.",
    "Check that all inbound movements have been recorded correctly.",
    "Check that all outbound movements have been recorded correctly.",
    "Check movement comments and vehicle registrations where applicable.",
    "Review client balances for unexpected negative balances or discrepancies.",
    "Review open alerts and routine outbound expectations.",
    "Confirm required movement photos or attachments are linked.",
    "Flag stock discrepancies, missing assets, or unresolved issues.",
    "Prepare notable movements and handover information for the next shift."
  ];

  const state = {
    accessToken: null,
    idTokenPayload: null,
    headers: CHECKLIST_HEADERS.slice(),
    rawRows: [],
    logs: [],
    selectedWindow: null,
    shiftWindows: [],
    saveQueue: Promise.resolve(),
    saveBusy: false
  };
  let tokenRequestPromise = null;
  const $ = id => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    bindQuickMenu();
    waitForGoogle();
  });

  function bindEvents() {
    $("grant-access")?.addEventListener("click", () => requestAccess(false));
    $("sign-out")?.addEventListener("click", signOut);
    $("refresh")?.addEventListener("click", loadChecklist);
    $("shift-select")?.addEventListener("change", () => {
      const selected = state.shiftWindows.find(window => window.id === $("shift-select").value);
      if (selected) {
        state.selectedWindow = selected;
        renderCurrentShift();
      }
    });
    $("checklist-items")?.addEventListener("change", event => {
      if (event.target.matches("[data-check-index], [data-check-note]")) {
        updateProgressFromForm();
        saveCurrentShift(false);
      }
    });
    $("shift-handover-note")?.addEventListener("change", () => saveCurrentShift(false));
    $("save-checklist")?.addEventListener("click", () => saveCurrentShift(true));
    $("close-shift")?.addEventListener("click", closeSelectedShift);
  }

  function bindQuickMenu() {
    const nav = document.querySelector(".quick-nav");
    const toggle = $("menu-toggle");
    const links = $("quick-links");
    if (!nav || !toggle) return;
    const key = "fmQuickMenuOpen";
    const mq = window.matchMedia("(max-width: 760px)");
    const isMobile = () => mq.matches;
    const setOpen = open => {
      const next = Boolean(open);
      nav.classList.toggle("menu-open", next);
      toggle.setAttribute("aria-expanded", String(next));
      if (!isMobile()) localStorage.setItem(key, next ? "1" : "0");
    };
    setOpen(isMobile() ? false : localStorage.getItem(key) === "1");
    toggle.addEventListener("click", event => {
      event.preventDefault();
      setOpen(!nav.classList.contains("menu-open"));
    });
    links?.addEventListener("click", event => {
      if (event.target.closest("a") && isMobile()) setOpen(false);
    });
    const onResize = event => { if (event.matches) setOpen(false); };
    if (mq.addEventListener) mq.addEventListener("change", onResize); else mq.addListener(onResize);
  }

  function waitForGoogle() {
    if (window.google?.accounts?.id && window.google?.accounts?.oauth2) return initializeGoogle();
    let finished = false;
    const started = Date.now();
    const finish = () => {
      if (finished) return;
      finished = true;
      clearInterval(timer);
      window.removeEventListener("fm-google-loaded", finish);
      initializeGoogle();
    };
    window.addEventListener("fm-google-loaded", finish, { once: true });
    const timer = setInterval(() => {
      if (window.google?.accounts?.id && window.google?.accounts?.oauth2) finish();
      else if (Date.now() - started >= 8000) {
        clearInterval(timer);
        window.removeEventListener("fm-google-loaded", finish);
        setAuthStatus("Google services could not be loaded. Check your internet connection.", true);
      }
    }, 50);
  }

  function initializeGoogle() {
    if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")) {
      setAuthStatus("Add the existing Google Web Client ID to config.js.", true);
      return;
    }
    google.accounts.id.initialize({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      callback: handleCredentialResponse,
      auto_select: true,
      cancel_on_tap_outside: false
    });
    const saved = readSession();
    if (saved) {
      setUserProfile(saved);
      $("google-signin-button")?.classList.add("hidden");
      $("grant-access")?.classList.add("hidden");
      const cached = window.FM_AUTH_CACHE?.read?.(saved.email);
      if (cached?.token) {
        state.accessToken = cached.token;
        window.FM_CONNECTION_UI?.show("Restoring your session…", "Using your saved Google Sheets connection.");
        hideLogin();
        loadChecklist().finally(() => window.FM_CONNECTION_UI?.hide());
      } else {
        setAuthStatus("Restoring your Sheets connection…");
        attemptSilentAccess(saved.email);
      }
    } else {
      window.FM_CONNECTION_UI?.hide();
      google.accounts.id.renderButton($("google-signin-button"), {
        theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280
      });
      google.accounts.id.prompt();
    }
  }

  function handleCredentialResponse(response) {
    try {
      state.idTokenPayload = decodeJwtPayload(response.credential);
      saveSession();
      setUserProfile(state.idTokenPayload);
      $("grant-access")?.classList.remove("hidden");
      setAuthStatus("Signed in. Connecting to Google Sheets…");
      attemptSilentAccess(state.idTokenPayload.email);
    } catch (error) {
      console.error(error);
      setAuthStatus("Google sign-in response could not be read. Please try again.", true);
    }
  }

  function requestAccess(silent = true) {
    const session = readSession();
    if (!state.idTokenPayload && !session) {
      setAuthStatus("Sign in with Google first.", true);
      return;
    }
    window.FM_CONNECTION_UI?.show(
      silent ? "Connecting to Google Sheets…" : "Authorising Google Sheets…",
      silent ? "Restoring your saved connection." : "Approve access in the Google prompt to continue."
    );
    acquireAccessToken(silent ? "none" : "consent")
      .then(async () => {
        hideLogin();
        await loadChecklist();
        window.FM_CONNECTION_UI?.hide();
      })
      .catch(error => {
        window.FM_CONNECTION_UI?.hide();
        $("grant-access")?.classList.remove("hidden");
        setAuthStatus(error?.message || "Google authorisation failed. Tap Connect Google Sheets to try again.", true);
      });
  }

  function attemptSilentAccess(email) {
    window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Checking your saved Google authorisation.");
    acquireAccessToken("none", email)
      .then(async () => {
        hideLogin();
        await loadChecklist();
        window.FM_CONNECTION_UI?.hide();
      })
      .catch(() => {
        window.FM_CONNECTION_UI?.hide();
        $("grant-access")?.classList.remove("hidden");
        setAuthStatus("Your Google session is available. Allow Google Sheets access to continue.");
      });
  }

  function acquireAccessToken(prompt = "none", email, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const expectedEmail = email || state.idTokenPayload?.email || readSession()?.email || "";
    if (!forceRefresh) {
      const cached = window.FM_AUTH_CACHE?.read?.(expectedEmail);
      if (cached?.token) {
        state.accessToken = cached.token;
        return Promise.resolve(cached.token);
      }
    }
    if (tokenRequestPromise) return tokenRequestPromise;
    tokenRequestPromise = new Promise((resolve, reject) => {
      let settled = false;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        tokenRequestPromise = null;
        clearTimeout(timeoutId);
        fn(value);
      };
      const timeoutId = setTimeout(() => finish(reject, new Error("Google authorisation did not complete.")), prompt === "none" ? 5000 : 15000);
      const client = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.GOOGLE_CLIENT_ID,
        scope: CONFIG.OAUTH_SCOPES,
        callback: response => {
          if (response.error) {
            finish(reject, new Error(`Google authorization failed: ${response.error}`));
            return;
          }
          state.accessToken = response.access_token;
          window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, expectedEmail);
          finish(resolve, response.access_token);
        }
      });
      client.requestAccessToken({ prompt, login_hint: expectedEmail || undefined });
    });
    return tokenRequestPromise;
  }

  async function loadChecklist() {
    if (!state.accessToken) return false;
    setSyncStatus("Syncing shift checklist…");
    try {
      $("workbook-name").textContent = CONFIG.INVENTORY_LEDGER_NAME || "Assets Inventory Ledger";
      await ensureChecklistSheet();
      state.shiftWindows = recentShiftWindows(12);
      const currentWindow = state.shiftWindows[0];
      const previousSelection = $("shift-select")?.value;
      const selectedWindow = state.shiftWindows.find(window => window.id === previousSelection) || currentWindow;
      populateShiftSelector(selectedWindow.id);
      state.selectedWindow = selectedWindow;
      renderCurrentShift();
      setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
      return true;
    } catch (error) {
      console.error(error);
      setSyncStatus(error.message || "Unable to load the shift checklist.", true);
      return false;
    }
  }

  async function ensureChecklistSheet() {
    const spreadsheetId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
    let metadata = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}?includeGridData=false`);
    const hasSheet = (metadata.sheets || []).some(sheet => sheet?.properties?.title === SHEET_NAME);
    if (!hasSheet) {
      try {
        await sheetsPost(`/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
          requests: [{ addSheet: { properties: { title: SHEET_NAME } } }]
        });
      } catch (error) {
        // Another open dashboard may have created the checklist sheet first.
        if (!/already exists|duplicate/i.test(String(error.message || ""))) throw error;
      }
      metadata = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}?includeGridData=false`);
      if (!(metadata.sheets || []).some(sheet => sheet?.properties?.title === SHEET_NAME)) {
        throw new Error(`The "${SHEET_NAME}" sheet could not be created.`);
      }
    }

    let rows = await getValues(spreadsheetId, SHEET_NAME);
    if (!rows.length || !(rows[0] || []).some(value => String(value || "").trim())) {
      await updateValues(spreadsheetId, SHEET_NAME, [CHECKLIST_HEADERS]);
      rows = [CHECKLIST_HEADERS.slice()];
    } else {
      const existingHeaders = rows[0].map(value => String(value ?? "").trim());
      const normalized = existingHeaders.map(normalizeHeader);
      const missing = CHECKLIST_HEADERS.filter(header => !normalized.includes(normalizeHeader(header)));
      if (missing.length) {
        const nextHeaders = existingHeaders.concat(missing);
        await updateValues(spreadsheetId, SHEET_NAME, [nextHeaders]);
        rows[0] = nextHeaders;
      }
    }
    state.headers = (rows[0] || CHECKLIST_HEADERS).map(value => String(value ?? "").trim());
    state.rawRows = rows;
    state.logs = parseLogRows(rows);
  }

  function parseLogRows(rows) {
    if (!rows?.length) return [];
    const headerMap = new Map(state.headers.map((header, index) => [normalizeHeader(header), index]));
    const read = (row, field) => row[headerMap.get(normalizeHeader(field)) ?? -1] ?? "";
    return rows.slice(1).map((row, index) => ({
      rowNumber: index + 2,
      shiftId: String(read(row, "Shift ID")).trim(),
      shift: String(read(row, "Shift")).trim(),
      shiftStart: String(read(row, "Shift Start")).trim(),
      shiftEnd: String(read(row, "Shift End")).trim(),
      item: String(read(row, "Checklist Item")).trim(),
      checked: parseChecked(read(row, "Checked")),
      itemNote: String(read(row, "Item Note")).trim(),
      handoverNote: String(read(row, "Shift Handover Note")).trim(),
      updatedAt: String(read(row, "Updated At")).trim(),
      closedAt: String(read(row, "Closed At")).trim(),
      closedBy: String(read(row, "Closed By")).trim(),
      shiftStatus: String(read(row, "Shift Status")).trim().toUpperCase()
    })).filter(record => record.shiftId && record.item);
  }

  function recentShiftWindows(count = 12, now = new Date()) {
    const hour = now.getHours();
    let current;
    if (hour >= 7 && hour < 19) current = makeShiftWindow("DAY", new Date(now.getFullYear(), now.getMonth(), now.getDate()));
    else if (hour >= 19) current = makeShiftWindow("NIGHT", new Date(now.getFullYear(), now.getMonth(), now.getDate()));
    else current = makeShiftWindow("NIGHT", new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1));

    const windows = [current];
    while (windows.length < Math.max(1, count)) windows.push(previousShiftWindow(windows[windows.length - 1]));
    return windows;
  }

  function makeShiftWindow(shift, anchorDate) {
    const year = anchorDate.getFullYear(), month = anchorDate.getMonth(), day = anchorDate.getDate();
    const start = shift === "DAY" ? new Date(year, month, day, 7, 0, 0, 0) : new Date(year, month, day, 19, 0, 0, 0);
    const end = new Date(start);
    if (shift === "DAY") end.setHours(19, 0, 0, 0);
    else end.setDate(end.getDate() + 1), end.setHours(7, 0, 0, 0);
    const anchor = localDateKey(start);
    return { id: `${shift}-${anchor}`, shift, start, end, startText: localStamp(start), endText: localStamp(end) };
  }

  function previousShiftWindow(window) {
    const anchor = new Date(window.start.getFullYear(), window.start.getMonth(), window.start.getDate());
    if (window.shift === "DAY") {
      anchor.setDate(anchor.getDate() - 1);
      return makeShiftWindow("NIGHT", anchor);
    }
    return makeShiftWindow("DAY", anchor);
  }

  function populateShiftSelector(selectedId) {
    const select = $("shift-select");
    if (!select) return;
    select.innerHTML = state.shiftWindows.map((window, index) => {
      const label = `${window.shift === "DAY" ? "Day" : "Night"} shift · ${displayDate(window.start)} · ${formatTime(window.start)}–${formatTime(window.end)}${index === 0 ? " (current shift)" : ""}`;
      return `<option value="${escapeAttr(window.id)}">${escapeHtml(label)}</option>`;
    }).join("");
    select.value = state.shiftWindows.some(window => window.id === selectedId) ? selectedId : state.shiftWindows[0]?.id || "";
  }

  function renderCurrentShift() {
    const window = state.selectedWindow;
    if (!window) return;
    const select = $("shift-select");
    if (select && select.value !== window.id) select.value = window.id;
    const records = state.logs.filter(record => record.shiftId === window.id);
    const byItem = new Map();
    records.slice().sort((a, b) => a.rowNumber - b.rowNumber).forEach(record => byItem.set(record.item, record));
    const latestRecord = records.slice().sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)) || b.rowNumber - a.rowNumber)[0];
    const shiftStatus = String(latestRecord?.shiftStatus || "IN PROGRESS").toUpperCase();
    const closed = ["COMPLETED", "HANDED OVER"].includes(shiftStatus) || Boolean(latestRecord?.closedAt);
    const checkedCount = CHECKLIST_ITEMS.filter(item => Boolean(byItem.get(item)?.checked)).length;
    const completion = Math.round((checkedCount / CHECKLIST_ITEMS.length) * 100);

    $("shift-window-details").textContent = `${window.shift === "DAY" ? "Day shift" : "Night shift"}: ${formatDateTime(window.start)} to ${formatDateTime(window.end)} · ${window.id}`;
    const pill = $("shift-status-pill");
    pill.textContent = closed ? shiftStatus : "IN PROGRESS";
    pill.className = `checklist-status-pill ${shiftStatus === "COMPLETED" ? "status-complete" : shiftStatus === "HANDED OVER" ? "status-handover" : "status-open"}`;

    $("checklist-items").innerHTML = CHECKLIST_ITEMS.map((item, index) => {
      const record = byItem.get(item);
      const checked = Boolean(record?.checked);
      return `<article class="checklist-item ${checked ? "is-checked" : ""}">
        <input id="check-item-${index}" type="checkbox" data-check-index="${index}" ${checked ? "checked" : ""} ${closed ? "disabled" : ""}>
        <div class="checklist-item-body">
          <label class="checklist-item-title" for="check-item-${index}">${escapeHtml(item)}</label>
          <input type="text" class="checklist-item-note" data-check-note="${index}" value="${escapeAttr(record?.itemNote || "")}" placeholder="Optional note or issue for this check" ${closed ? "disabled" : ""}>
        </div>
      </article>`;
    }).join("");

    $("shift-handover-note").value = latestRecord?.handoverNote || "";
    $("shift-handover-note").disabled = closed;
    $("save-checklist").disabled = closed;
    $("close-shift").disabled = closed;
    $("close-shift").textContent = closed ? (shiftStatus === "HANDED OVER" ? "Shift handed over" : "Shift closed") : "Close shift & hand over";
    setProgress(checkedCount, completion);

    const handover = $("handover-result");
    if (closed) {
      const incomplete = CHECKLIST_ITEMS.length - checkedCount;
      handover.classList.remove("hidden");
      if (shiftStatus === "HANDED OVER" || incomplete > 0) {
        handover.className = "checklist-handover-result";
        handover.textContent = `Shift closed and handed over with ${incomplete} check${incomplete === 1 ? "" : "s"} incomplete. Unchecked items and notes are preserved in the Shift Checklist sheet.`;
      } else {
        handover.className = "checklist-handover-result is-complete";
        handover.textContent = "Shift closed with all checks completed. The checklist has been saved to the Shift Checklist sheet.";
      }
      setSaveStatus(`Closed at ${latestRecord?.closedAt || "recorded close time"}${latestRecord?.closedBy ? ` · ${latestRecord.closedBy}` : ""}.`);
    } else {
      handover.classList.add("hidden");
      handover.textContent = "";
      if (records.length) setSaveStatus("This shift is open. Changes save automatically to Google Sheets.");
      else setSaveStatus("No checks saved for this shift yet. Your first save will create its log.");
    }
  }

  function setProgress(checkedCount, percent) {
    $("checklist-progress-label").textContent = `${checkedCount} of ${CHECKLIST_ITEMS.length} checks complete`;
    $("checklist-progress-percent").textContent = `${percent}%`;
    $("checklist-progress-fill").style.width = `${percent}%`;
    $("checklist-progress-track").setAttribute("aria-valuenow", String(percent));
    $("checklist-progress-help").textContent = "Checks, item notes and handover notes are logged to the shift sheet.";
  }

  function updateProgressFromForm() {
    const checkedCount = CHECKLIST_ITEMS.filter((_, index) => $(`check-item-${index}`)?.checked).length;
    const percent = Math.round((checkedCount / CHECKLIST_ITEMS.length) * 100);
    setProgress(checkedCount, percent);
    CHECKLIST_ITEMS.forEach((_, index) => {
      const checked = Boolean($(`check-item-${index}`)?.checked);
      $(`check-item-${index}`)?.closest(".checklist-item")?.classList.toggle("is-checked", checked);
    });
  }

  function snapshotFromForm(shiftStatus = "IN PROGRESS", closedAt = "", closedBy = "") {
    const window = state.selectedWindow;
    const items = CHECKLIST_ITEMS.map((label, index) => ({
      label,
      checked: Boolean($(`check-item-${index}`)?.checked),
      note: String(document.querySelector(`[data-check-note="${index}"]`)?.value || "").trim()
    }));
    return {
      window,
      items,
      handoverNote: String($("shift-handover-note")?.value || "").trim(),
      shiftStatus,
      closedAt,
      closedBy
    };
  }

  function saveCurrentShift(manual = false) {
    if (!state.selectedWindow) return Promise.resolve();
    if (isSelectedShiftClosed()) {
      setSaveStatus("This shift is closed and its handover record is locked.", true);
      return Promise.resolve();
    }
    const snapshot = snapshotFromForm();
    setSaveStatus(manual ? "Saving checklist progress…" : "Saving changes…");
    return queueWrite(snapshot).then(() => {
      const checkedCount = snapshot.items.filter(item => item.checked).length;
      setSaveStatus(`Saved at ${new Date().toLocaleTimeString()} · ${checkedCount}/${CHECKLIST_ITEMS.length} complete.`);
      if (manual) {
        $("handover-result").className = "checklist-handover-result is-complete";
        $("handover-result").textContent = "Progress saved to the Shift Checklist sheet.";
        $("handover-result").classList.remove("hidden");
      }
    }).catch(error => {
      console.error(error);
      setSaveStatus(error.message || "Could not save checklist progress.", true);
    });
  }

  async function closeSelectedShift() {
    const window = state.selectedWindow;
    if (!window || isSelectedShiftClosed()) return;
    const snapshot = snapshotFromForm();
    const incomplete = snapshot.items.filter(item => !item.checked).length;
    if (incomplete > 0) {
      const proceed = windowConfirm(`There ${incomplete === 1 ? "is" : "are"} ${incomplete} unchecked check${incomplete === 1 ? "" : "s"}. Close this shift and hand them over as-is? The unchecked items will remain unchecked and be logged for the next shift.`);
      if (!proceed) return;
    } else if (!windowConfirm("All checks are complete. Close this shift now? The checklist will become read-only.")) {
      return;
    }
    const closedAt = new Date().toISOString();
    const closedBy = state.idTokenPayload?.email || readSession()?.email || state.idTokenPayload?.name || "Google user";
    snapshot.shiftStatus = incomplete > 0 ? "HANDED OVER" : "COMPLETED";
    snapshot.closedAt = closedAt;
    snapshot.closedBy = closedBy;
    setSaveStatus(incomplete > 0 ? "Saving the handover…" : "Closing shift…");
    $("close-shift").disabled = true;
    try {
      await queueWrite(snapshot);
      renderCurrentShift();
      setSyncStatus(`Shift log saved at ${new Date().toLocaleTimeString()}`);
    } catch (error) {
      console.error(error);
      $("close-shift").disabled = false;
      setSaveStatus(error.message || "Shift could not be closed. Your checks remain editable.", true);
    }
  }

  function isSelectedShiftClosed() {
    const id = state.selectedWindow?.id;
    if (!id) return false;
    return state.logs.some(record => record.shiftId === id && ["COMPLETED", "HANDED OVER"].includes(record.shiftStatus));
  }

  function queueWrite(snapshot) {
    state.saveQueue = state.saveQueue.catch(() => {}).then(() => writeSnapshot(snapshot));
    return state.saveQueue;
  }

  async function writeSnapshot(snapshot) {
    if (!state.accessToken) throw new Error("Google Sheets access is unavailable. Reconnect and try again.");
    const headers = state.headers;
    const timestamp = new Date().toISOString();
    const existingMap = new Map();
    state.logs.filter(record => record.shiftId === snapshot.window.id).forEach(record => {
      const key = `${record.shiftId}\u0000${normalizeHeader(record.item)}`;
      const current = existingMap.get(key);
      if (!current || record.rowNumber > current.rowNumber) existingMap.set(key, record);
    });

    const updates = [];
    const appendRows = [];
    const appendItems = [];
    const recordsToSave = [];
    snapshot.items.forEach(item => {
      const row = buildSheetRow(snapshot, item, timestamp, headers);
      const key = `${snapshot.window.id}\u0000${normalizeHeader(item.label)}`;
      const existing = existingMap.get(key);
      if (existing) {
        updates.push({
          range: `${quoteSheetName(SHEET_NAME)}!A${existing.rowNumber}:${columnLetter(headers.length)}${existing.rowNumber}`,
          majorDimension: "ROWS",
          values: [row]
        });
        recordsToSave.push({ ...existing, ...rowToRecord(row, headers, existing.rowNumber) });
      } else {
        appendRows.push(row);
        appendItems.push(item);
      }
    });

    const spreadsheetId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
    if (updates.length) {
      await sheetsPost(`/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`, {
        valueInputOption: "USER_ENTERED",
        data: updates
      });
    }

    let appendStartRow = Math.max(1, ...state.logs.map(record => Number(record.rowNumber) || 1)) + 1;
    if (appendRows.length) {
      const appendRange = `${quoteSheetName(SHEET_NAME)}!A:${columnLetter(headers.length)}`;
      const appended = await sheetsPost(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(appendRange)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`, {
        majorDimension: "ROWS",
        values: appendRows
      });
      const updatedRange = String(appended?.updates?.updatedRange || "");
      const startMatch = updatedRange.match(/![A-Z]+(\d+):/i);
      if (startMatch) appendStartRow = Number(startMatch[1]);
      appendRows.forEach((row, index) => recordsToSave.push(rowToRecord(row, headers, appendStartRow + index)));
    }

    const savedKeys = new Set(recordsToSave.map(record => `${record.shiftId}\u0000${normalizeHeader(record.item)}`));
    state.logs = state.logs.filter(record => !savedKeys.has(`${record.shiftId}\u0000${normalizeHeader(record.item)}`));
    state.logs.push(...recordsToSave);
    state.rawRows = state.rawRows.slice();
    state.rawRows.length = Math.max(state.rawRows.length, ...recordsToSave.map(record => record.rowNumber));
  }

  function buildSheetRow(snapshot, item, timestamp, headers) {
    const row = new Array(headers.length).fill("");
    const put = (name, value) => {
      const index = headers.findIndex(header => normalizeHeader(header) === normalizeHeader(name));
      if (index >= 0) row[index] = value ?? "";
    };
    put("Shift ID", snapshot.window.id);
    put("Shift", snapshot.window.shift);
    put("Shift Start", snapshot.window.startText);
    put("Shift End", snapshot.window.endText);
    put("Checklist Item", item.label);
    put("Checked", item.checked ? "TRUE" : "FALSE");
    put("Item Note", item.note || "");
    put("Shift Handover Note", snapshot.handoverNote || "");
    put("Updated At", timestamp);
    put("Closed At", snapshot.closedAt || "");
    put("Closed By", snapshot.closedBy || "");
    put("Shift Status", snapshot.shiftStatus || "IN PROGRESS");
    return row;
  }

  function rowToRecord(row, headers, rowNumber) {
    const value = name => row[headers.findIndex(header => normalizeHeader(header) === normalizeHeader(name))] ?? "";
    return {
      rowNumber,
      shiftId: String(value("Shift ID") || ""),
      shift: String(value("Shift") || ""),
      shiftStart: String(value("Shift Start") || ""),
      shiftEnd: String(value("Shift End") || ""),
      item: String(value("Checklist Item") || ""),
      checked: parseChecked(value("Checked")),
      itemNote: String(value("Item Note") || ""),
      handoverNote: String(value("Shift Handover Note") || ""),
      updatedAt: String(value("Updated At") || ""),
      closedAt: String(value("Closed At") || ""),
      closedBy: String(value("Closed By") || ""),
      shiftStatus: String(value("Shift Status") || "IN PROGRESS").toUpperCase()
    };
  }

  async function getValues(spreadsheetId, sheetName) {
    const range = `${quoteSheetName(sheetName)}!A:Z`;
    const response = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`);
    return response.values || [];
  }

  async function updateValues(spreadsheetId, sheetName, rows) {
    const range = `${quoteSheetName(sheetName)}!A1`;
    return sheetsPut(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
      range,
      majorDimension: "ROWS",
      values: rows
    });
  }

  async function sheetsGet(path) { return fetchJson(SHEETS_API + path); }
  async function sheetsPost(path, body) { return fetchJson(SHEETS_API + path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); }
  async function sheetsPut(path, body) { return fetchJson(SHEETS_API + path, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); }

  async function fetchJson(url, options = {}) {
    const { __retried, ...fetchOptions } = options;
    const response = await fetch(url, {
      ...fetchOptions,
      headers: { ...authHeaders(), ...(fetchOptions.headers || {}) }
    });
    if (response.status === 401 && !__retried) {
      try {
        await acquireAccessToken("none", state.idTokenPayload?.email || readSession()?.email, { forceRefresh: true });
        return fetchJson(url, { ...fetchOptions, __retried: true });
      } catch (_) {
        state.accessToken = null;
        window.FM_AUTH_CACHE?.clear?.();
        $("grant-access")?.classList.remove("hidden");
        throw new Error("Google Sheets access expired. Allow Google Sheets access to reconnect.");
      }
    }
    const raw = await response.text();
    let data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
    if (!response.ok) throw new Error(data?.error?.message || `Google Sheets request failed (${response.status})`);
    return data;
  }

  function authHeaders() { return { Authorization: `Bearer ${state.accessToken}` }; }
  function quoteSheetName(name) { return `'${String(name).replace(/'/g, "''")}'`; }
  function normalizeHeader(value) { return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " "); }
  function parseChecked(value) { return [true, 1, "1", "true", "yes", "checked", "complete", "completed"].includes(typeof value === "string" ? value.trim().toLowerCase() : value); }
  function localDateKey(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; }
  function localStamp(date) { return `${localDateKey(date)} ${formatTime(date)}`; }
  function formatTime(date) { return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`; }
  function formatDateTime(date) { return `${displayDate(date)} ${formatTime(date)}`; }
  function displayDate(date) { return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }); }
  function columnLetter(number) { let n = Number(number) || 1, result = ""; while (n > 0) { const remainder = (n - 1) % 26; result = String.fromCharCode(65 + remainder) + result; n = Math.floor((n - 1) / 26); } return result || "A"; }
  function decodeJwtPayload(jwt) {
    const parts = String(jwt).split(".");
    if (parts.length !== 3) throw new Error("Invalid Google credential.");
    const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64 + "=".repeat((4 - base64.length % 4) % 4);
    return JSON.parse(decodeURIComponent(Array.from(atob(padded)).map(char => `%${char.charCodeAt(0).toString(16).padStart(2, "0")}`).join("")));
  }
  function saveSession() {
    if (!state.idTokenPayload) return;
    localStorage.setItem(SESSION_KEY, JSON.stringify({ name: state.idTokenPayload.name || "Google user", email: state.idTokenPayload.email || "", picture: state.idTokenPayload.picture || "", sub: state.idTokenPayload.sub || "" }));
  }
  function readSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; } }
  function setUserProfile(profile) {
    const name = String(profile?.given_name || profile?.name || "Google user").trim();
    const firstName = (name.split(/\s+/)[0] || "there").replace(/[^\p{L}\p{M}'-]/gu, "");
    $("user-name").textContent = `Hi, ${firstName || "there"} 👋`;
    $("user-email").textContent = "";
    if (profile?.picture) { $("user-photo").src = profile.picture; $("user-photo").classList.remove("hidden"); }
  }
  function hideLogin() {
    $("login-card")?.classList.add("hidden");
    $("checklist-page")?.classList.remove("hidden");
    $("google-signin-button")?.classList.add("hidden");
    $("grant-access")?.classList.add("hidden");
    $("sign-out")?.classList.remove("hidden");
  }
  function signOut() {
    const saved = readSession();
    try { if (saved?.sub) google.accounts.id.revoke(saved.sub, () => {}); } catch (_) {}
    localStorage.removeItem(SESSION_KEY);
    window.FM_AUTH_CACHE?.clear?.();
    location.reload();
  }
  function setAuthStatus(message, error = false) {
    const el = $("auth-status");
    if (el) { el.textContent = message || ""; el.className = `status ${error ? "error" : ""}`; }
  }
  function setSyncStatus(message, error = false) {
    const el = $("sync-status");
    if (el) { el.textContent = message || ""; el.className = `muted ${error ? "error" : ""}`; }
  }
  function setSaveStatus(message, error = false) {
    const el = $("save-status");
    if (el) { el.textContent = message || ""; el.classList.toggle("is-error", Boolean(error)); }
  }
  function windowConfirm(message) { return window.confirm(message); }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>\'"]/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char])); }
  function escapeAttr(value) { return escapeHtml(value); }

  function refreshSavedSessionIfNeeded() {
    const saved = readSession();
    if (!saved || !window.google?.accounts?.oauth2) return;
    if (window.FM_AUTH_CACHE?.read?.(saved.email)?.token) return;
    attemptSilentAccess(saved.email);
  }
  window.addEventListener("pageshow", refreshSavedSessionIfNeeded);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) refreshSavedSessionIfNeeded(); });
})();
