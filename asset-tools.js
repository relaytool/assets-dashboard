(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const HEADERS = {
    availability: ["Date", "Asset", "Available Count", "Updated By", "Updated At"],
    operations: ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Operation ID"],
    adjustments: ["Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link", "Inventory Balance", "Target Balance", "Previous Balance", "Adjustment ID"],
    rules: ["Rule ID", "Rule Name", "Output Asset", "Output Qty", "Component Asset", "Component Qty", "Active", "Created At"]
  };
  const state = {
    accessToken: null,
    idTokenPayload: null,
    inventory: [],
    clients: [],
    standardTransactions: [],
    operations: [],
    adjustments: [],
    rules: [],
    rulesRows: [],
    operationRows: [],
    adjustmentRows: []
  };
  let tokenRequestPromise = null;
  let pageBusy = false;
  let operationBusy = false;
  const $ = id => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    bindQuickMenu();
    waitForGoogle();
  });

  function bindEvents() {
    $("grant-access")?.addEventListener("click", () => requestAccess(false));
    $("sign-out")?.addEventListener("click", signOut);
    $("refresh")?.addEventListener("click", loadPage);
    $("combine-rule")?.addEventListener("change", renderCombinePreview);
    $("combine-batches")?.addEventListener("input", renderCombinePreview);
    $("combine-form")?.addEventListener("submit", combineAssets);
    $("rule-form")?.addEventListener("submit", saveRule);
    $("add-rule-component")?.addEventListener("click", () => addRuleComponentRow());
    $("rule-components")?.addEventListener("click", event => {
      const remove = event.target.closest("[data-remove-component]");
      if (!remove) return;
      const rows = $("rule-components").querySelectorAll(".rule-component-row");
      if (rows.length <= 2) {
        setStatus("rule-status", "A rule must use at least two components.", true);
        return;
      }
      remove.closest(".rule-component-row")?.remove();
      updateRuleComponentButtons();
    });
    $("rules-list")?.addEventListener("click", event => {
      const button = event.target.closest("[data-disable-rule]");
      if (button) deactivateRule(button.dataset.disableRule);
    });
    $("normalise-form")?.addEventListener("submit", normaliseClientBalance);
    const refreshTargetBalance = () => {
      if ($("normalise-target-balance")) delete $("normalise-target-balance").dataset.userEdited;
      updateCurrentClientBalance();
    };
    $("normalise-client")?.addEventListener("change", refreshTargetBalance);
    $("normalise-asset")?.addEventListener("change", refreshTargetBalance);
  }

  function bindQuickMenu() {
    const nav = document.querySelector(".quick-nav");
    const toggle = $("menu-toggle");
    const links = $("quick-links");
    if (!nav || !toggle) return;
    const key = "fmQuickMenuOpen";
    const mq = window.matchMedia("(max-width: 760px)");
    const setOpen = open => {
      nav.classList.toggle("menu-open", Boolean(open));
      toggle.setAttribute("aria-expanded", String(Boolean(open)));
      if (!mq.matches) { try { localStorage.setItem(key, open ? "1" : "0"); } catch (_) {} }
    };
    let saved = "0";
    try { saved = localStorage.getItem(key) || "0"; } catch (_) {}
    setOpen(!mq.matches && saved === "1");
    toggle.addEventListener("click", () => setOpen(!nav.classList.contains("menu-open")));
    links?.addEventListener("click", event => { if (event.target.closest("a") && mq.matches) setOpen(false); });
    const onResize = event => { if (event.matches) setOpen(false); };
    if (mq.addEventListener) mq.addEventListener("change", onResize); else mq.addListener(onResize);
  }

  function waitForGoogle() {
    if (window.google?.accounts?.id && window.google?.accounts?.oauth2) return initializeGoogle();
    let done = false;
    const started = Date.now();
    const finish = () => { if (done) return; done = true; clearInterval(timer); initializeGoogle(); };
    window.addEventListener("fm-google-loaded", finish, { once: true });
    const timer = setInterval(() => {
      if (window.google?.accounts?.id && window.google?.accounts?.oauth2) finish();
      else if (Date.now() - started >= 8000) { clearInterval(timer); setStatus("auth-status", "Google sign-in could not be loaded. Check your connection and refresh.", true); }
    }, 50);
  }

  function initializeGoogle() {
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
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…");
        hideLogin();
        loadPage().finally(() => window.FM_CONNECTION_UI?.hide());
      } else {
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…");
        attemptSilentAccess(saved.email);
      }
    } else {
      window.FM_CONNECTION_UI?.hide();
      google.accounts.id.renderButton($("google-signin-button"), { theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280 });
      google.accounts.id.prompt();
    }
  }

  function handleCredentialResponse(response) {
    try {
      state.idTokenPayload = decodeJwtPayload(response.credential);
      saveSession();
      setUserProfile(state.idTokenPayload);
      requestAccess(true);
    } catch (error) {
      console.error(error);
      setStatus("auth-status", "Google sign-in could not be completed. Please try again.", true);
    }
  }

  function requestAccess(silent = true) {
    window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…");
    acquireAccessToken(silent ? "none" : "consent").then(async () => {
      hideLogin();
      await loadPage();
      window.FM_CONNECTION_UI?.hide();
    }).catch(error => {
      window.FM_CONNECTION_UI?.hide();
      $("grant-access")?.classList.remove("hidden");
      setStatus("auth-status", error.message || "Google access could not be granted.", true);
    });
  }

  function attemptSilentAccess(email) {
    acquireAccessToken("none", email).then(async () => {
      hideLogin();
      await loadPage();
      window.FM_CONNECTION_UI?.hide();
    }).catch(() => {
      window.FM_CONNECTION_UI?.hide();
      $("grant-access")?.classList.remove("hidden");
      setStatus("auth-status", "Your Google account is signed in. Connect Google Sheets to continue.", true);
    });
  }

  function acquireAccessToken(prompt = "none", email, options = {}) {
    const expectedEmail = email || state.idTokenPayload?.email || readSession()?.email || "";
    if (!options.forceRefresh) {
      const cached = window.FM_AUTH_CACHE?.read?.(expectedEmail);
      if (cached?.token) { state.accessToken = cached.token; return Promise.resolve(cached.token); }
    }
    if (tokenRequestPromise) return tokenRequestPromise;
    tokenRequestPromise = new Promise((resolve, reject) => {
      let settled = false;
      const finish = (fn, value) => { if (settled) return; settled = true; tokenRequestPromise = null; clearTimeout(timeoutId); fn(value); };
      const timeoutId = setTimeout(() => finish(reject, new Error("Google authorisation did not complete.")), prompt === "none" ? 6000 : 15000);
      const client = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.GOOGLE_CLIENT_ID,
        scope: CONFIG.OAUTH_SCOPES,
        callback: response => {
          if (response.error) { finish(reject, new Error(`Google access failed: ${response.error}`)); return; }
          state.accessToken = response.access_token;
          window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, expectedEmail);
          finish(resolve, response.access_token);
        }
      });
      client.requestAccessToken({ prompt, login_hint: expectedEmail || undefined });
    });
    return tokenRequestPromise;
  }

  async function loadPage() {
    if (!state.accessToken || pageBusy) return;
    pageBusy = true;
    setSyncStatus("Refreshing data…");
    try {
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const meta = await sheetsGet(`/${encodeURIComponent(id)}`);
      const titles = (meta.sheets || []).map(sheet => sheet.properties.title);
      const sheetSpecs = [
        [CONFIG.ASSET_AVAILABILITY_SHEET_NAME, HEADERS.availability],
        [CONFIG.ASSET_OPERATIONS_SHEET_NAME, HEADERS.operations],
        [CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME, HEADERS.adjustments],
        [CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME, HEADERS.rules]
      ];
      const missing = sheetSpecs.filter(([name]) => !titles.includes(name));
      if (missing.length) {
        await sheetsPost(`/${encodeURIComponent(id)}:batchUpdate`, { requests: missing.map(([title]) => ({ addSheet: { properties: { title } } })) });
        for (const [name, headers] of missing) await updateValues(id, name, [headers]);
      }
      const [inventoryRows, clientRows, transactionRows, operationRows, adjustmentRows, ruleRows] = await Promise.all([
        getValues(id, CONFIG.INVENTORY_SHEET_NAME),
        getValues(id, CONFIG.CLIENT_LIST_SHEET_NAME),
        getValues(id, CONFIG.TRANSACTIONS_SHEET_NAME),
        getValues(id, CONFIG.ASSET_OPERATIONS_SHEET_NAME),
        getValues(id, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME),
        getValues(id, CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME)
      ]);
      if (!inventoryRows.length) throw new Error("Asset Inventory has no data to load.");
      if (!operationRows.length) await updateValues(id, CONFIG.ASSET_OPERATIONS_SHEET_NAME, [HEADERS.operations]);
      if (!adjustmentRows.length) await updateValues(id, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME, [HEADERS.adjustments]);
      if (!ruleRows.length) await updateValues(id, CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME, [HEADERS.rules]);

      state.inventory = parseInventory(inventoryRows);
      state.clients = parseClients(clientRows);
      state.standardTransactions = parseTransactions(transactionRows);
      state.operations = parseOperations(operationRows.length ? operationRows : [HEADERS.operations]);
      state.adjustments = parseAdjustments(adjustmentRows.length ? adjustmentRows : [HEADERS.adjustments]);
      state.operationRows = operationRows;
      state.adjustmentRows = adjustmentRows;
      state.rulesRows = ruleRows;
      state.rules = parseRules(ruleRows);

      if (!state.rules.length && ruleRows.length <= 1) {
        const defaultRuleRows = makeDefaultMagnumRuleRows();
        if (defaultRuleRows.length) {
          await sheetsPost(`/${encodeURIComponent(id)}/values:batchUpdate`, {
            valueInputOption: "USER_ENTERED",
            data: [{ range: `'${CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME}'!A2:H${defaultRuleRows.length + 1}`, majorDimension: "ROWS", values: defaultRuleRows }]
          });
          state.rulesRows = [HEADERS.rules, ...defaultRuleRows];
          state.rules = parseRules(state.rulesRows);
        }
      }

      $("workbook-name").textContent = CONFIG.INVENTORY_LEDGER_NAME || "Assets Inventory Ledger";
      populateFormOptions();
      renderRules();
      renderCombinePreview();
      renderAdjustmentHistory();
      renderOperationHistory();
      setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
    } catch (error) {
      console.error(error);
      setSyncStatus(error.message || "Could not load asset tools.", true);
    } finally {
      pageBusy = false;
    }
  }

  function makeDefaultMagnumRuleRows() {
    const lids = findAsset("Magnum Lids");
    const sleeves = findAsset("Magnum Sleeves");
    const magnum = findAsset("Magnum");
    if (!lids || !sleeves || !magnum) return [];
    const created = new Date().toISOString();
    return [
      ["RULE-MAGNUM", "Magnum assembly", magnum.asset, 1, lids.asset, 1, true, created],
      ["RULE-MAGNUM", "Magnum assembly", magnum.asset, 1, sleeves.asset, 1, true, created]
    ];
  }

  function parseInventory(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const h = rows[0].map(normalizeHeader);
    const ai = findHeader(h, ["asset", "asset name", "item", "type"]);
    const bi = findHeader(h, ["balance", "current balance", "stock", "quantity"]);
    if (ai < 0) return [];
    return rows.slice(1).map((row, i) => ({
      rowNumber: i + 2,
      asset: String(row[ai] ?? "").trim(),
      balance: bi >= 0 ? numberValue(row[bi]) : 0,
      balanceColumn: bi >= 0 ? bi + 1 : 2
    })).filter(item => item.asset);
  }

  function parseClients(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const h = rows[0].map(normalizeHeader);
    const idx = findHeader(h, ["client", "client name", "name"]);
    return idx < 0 ? [] : [...new Set(rows.slice(1).map(row => String(row[idx] ?? "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }

  function parseTransactions(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const h = rows[0].map(normalizeHeader);
    const idx = {
      timestamp: findHeader(h, ["timestamp", "date", "datetime"]),
      client: findHeader(h, ["client", "client name"]),
      movement: findHeader(h, ["movement", "type", "direction"]),
      asset: findHeader(h, ["asset", "asset name", "item"]),
      quantity: findHeader(h, ["quantity", "qty", "delta"]),
      user: findHeader(h, ["user", "entered by", "email"]),
      comment: findHeader(h, ["comment", "comments", "notes"]),
      balance: findHeader(h, ["inventory balance", "warehouse balance", "balance after movement"])
    };
    return rows.slice(1).map(row => ({
      timestamp: idx.timestamp >= 0 ? row[idx.timestamp] ?? "" : "",
      client: idx.client >= 0 ? row[idx.client] ?? "" : "",
      movement: idx.movement >= 0 ? row[idx.movement] ?? "" : "",
      asset: idx.asset >= 0 ? row[idx.asset] ?? "" : "",
      quantity: idx.quantity >= 0 ? numberValue(row[idx.quantity]) : 0,
      user: idx.user >= 0 ? row[idx.user] ?? "" : "",
      comment: idx.comment >= 0 ? String(row[idx.comment] ?? "") : "",
      inventoryBalance: idx.balance >= 0 ? numberValue(row[idx.balance]) : null
    })).filter(row => row.client || row.asset);
  }

  function parseOperations(rows) {
    return parseTransactions(rows).filter(row => row.asset && ["COMBINE_OUT", "COMBINE_IN"].includes(String(row.movement).toUpperCase()));
  }

  function parseAdjustments(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const headers = rows[0].map(normalizeHeader);
    const parsed = parseTransactions(rows);
    const targetIdx = findHeader(headers, ["target balance"]);
    const previousIdx = findHeader(headers, ["previous balance", "old balance"]);
    const idIdx = findHeader(headers, ["adjustment id", "id"]);
    const timestampIdx = findHeader(headers, ["timestamp", "date", "datetime"]);
    const clientIdx = findHeader(headers, ["client", "client name"]);
    const movementIdx = findHeader(headers, ["movement", "type", "direction"]);
    const assetIdx = findHeader(headers, ["asset", "asset name", "item"]);
    const quantityIdx = findHeader(headers, ["quantity", "qty", "delta"]);
    const userIdx = findHeader(headers, ["user", "entered by", "email"]);
    const commentIdx = findHeader(headers, ["comment", "comments", "notes"]);
    const balanceIdx = findHeader(headers, ["inventory balance", "warehouse balance", "balance after movement"]);
    return rows.slice(1).map((source, index) => {
      if (String(movementIdx >= 0 ? source[movementIdx] : "").trim().toUpperCase() !== "NORMALISATION") return null;
      return {
        timestamp: timestampIdx >= 0 ? source[timestampIdx] ?? "" : "",
        client: clientIdx >= 0 ? source[clientIdx] ?? "" : "",
        movement: movementIdx >= 0 ? source[movementIdx] ?? "" : "NORMALISATION",
        asset: assetIdx >= 0 ? source[assetIdx] ?? "" : "",
        quantity: quantityIdx >= 0 ? numberValue(source[quantityIdx]) : 0,
        user: userIdx >= 0 ? source[userIdx] ?? "" : "",
        comment: commentIdx >= 0 ? String(source[commentIdx] ?? "") : "",
        inventoryBalance: balanceIdx >= 0 && source[balanceIdx] !== "" ? numberValue(source[balanceIdx]) : null,
        targetBalance: targetIdx >= 0 && source[targetIdx] !== "" ? numberValue(source[targetIdx]) : null,
        previousBalance: previousIdx >= 0 && source[previousIdx] !== "" ? numberValue(source[previousIdx]) : null,
        adjustmentId: idIdx >= 0 ? String(source[idIdx] ?? "") : "",
        rowNumber: index + 2
      };
    }).filter(row => row && row.client && row.asset);
  }

  function parseRules(rows) {
    if (!Array.isArray(rows) || rows.length < 2) return [];
    const h = rows[0].map(normalizeHeader);
    const idx = {
      id: findHeader(h, ["rule id", "id"]),
      name: findHeader(h, ["rule name", "name"]),
      output: findHeader(h, ["output asset", "asset created", "output"]),
      outputQty: findHeader(h, ["output qty", "output quantity", "yield"]),
      component: findHeader(h, ["component asset", "input asset", "component"]),
      componentQty: findHeader(h, ["component qty", "component quantity", "input quantity"]),
      active: findHeader(h, ["active", "enabled"])
    };
    if ([idx.id, idx.name, idx.output, idx.component].some(value => value < 0)) return [];
    const grouped = new Map();
    rows.slice(1).forEach((row, i) => {
      const id = String(row[idx.id] ?? "").trim();
      const component = String(row[idx.component] ?? "").trim();
      if (!id || !component) return;
      const activeText = idx.active >= 0 ? String(row[idx.active] ?? "TRUE").trim().toUpperCase() : "TRUE";
      if (["FALSE", "0", "NO", "INACTIVE"].includes(activeText)) return;
      if (!grouped.has(id)) {
        grouped.set(id, {
          id,
          name: String(row[idx.name] ?? "").trim() || "Unnamed rule",
          outputAsset: String(row[idx.output] ?? "").trim(),
          outputQty: idx.outputQty >= 0 ? positiveInteger(row[idx.outputQty], 1) : 1,
          components: [],
          rowNumbers: [],
          rawRows: []
        });
      }
      const rule = grouped.get(id);
      rule.components.push({ asset: component, quantity: idx.componentQty >= 0 ? positiveInteger(row[idx.componentQty], 1) : 1 });
      rule.rowNumbers.push(i + 2);
      rule.rawRows.push(row.slice());
    });
    return [...grouped.values()].filter(rule => rule.outputAsset && rule.components.length >= 2);
  }

  function populateFormOptions() {
    const assetOptions = `<option value="">Select asset</option>${state.inventory.map(item => `<option value="${escapeAttr(item.asset)}">${escapeHtml(item.asset)}</option>`).join("")}`;
    const clientOptions = `<option value="">Select client</option>${state.clients.filter(client => !["HSC London (Self)", "HSC London(Self)"].includes(client)).map(client => `<option value="${escapeAttr(client)}">${escapeHtml(client)}</option>`).join("")}`;
    const ruleSelect = $("combine-rule");
    const prevRule = ruleSelect.value;
    ruleSelect.innerHTML = `<option value="">Select a rule</option>${state.rules.map(rule => `<option value="${escapeAttr(rule.id)}">${escapeHtml(rule.name)} · ${escapeHtml(rule.components.map(part => part.asset).join(" + "))} → ${escapeHtml(rule.outputAsset)}</option>`).join("")}`;
    if (state.rules.some(rule => rule.id === prevRule)) ruleSelect.value = prevRule;
    $("rule-output-asset").innerHTML = assetOptions;
    const selectedClient = $("normalise-client").value;
    const selectedAsset = $("normalise-asset").value;
    $("normalise-client").innerHTML = clientOptions;
    $("normalise-asset").innerHTML = assetOptions;
    if (state.clients.includes(selectedClient)) $("normalise-client").value = selectedClient;
    if (state.inventory.some(item => item.asset === selectedAsset)) $("normalise-asset").value = selectedAsset;
    if (!$("rule-components").children.length) {
      addRuleComponentRow();
      addRuleComponentRow();
    } else {
      $("rule-components").querySelectorAll("select[data-component-asset]").forEach(select => {
        const value = select.value;
        select.innerHTML = assetOptions;
        if (state.inventory.some(item => item.asset === value)) select.value = value;
      });
    }
    updateRuleComponentButtons();
    updateCurrentClientBalance();
  }

  function addRuleComponentRow(assetValue = "", quantity = 1) {
    const container = $("rule-components");
    if (!container) return;
    const row = document.createElement("div");
    row.className = "rule-component-row";
    const options = `<option value="">Select component</option>${state.inventory.map(item => `<option value="${escapeAttr(item.asset)}">${escapeHtml(item.asset)}</option>`).join("")}`;
    row.innerHTML = `<label>Component asset<select data-component-asset required>${options}</select></label><label>Qty<input data-component-quantity type="number" min="1" step="1" value="${quantity}" required></label><button type="button" class="secondary" data-remove-component aria-label="Remove component">×</button>`;
    container.appendChild(row);
    const select = row.querySelector("select");
    if (assetValue && state.inventory.some(item => item.asset === assetValue)) select.value = assetValue;
    updateRuleComponentButtons();
  }

  function updateRuleComponentButtons() {
    const rows = [...($("rule-components")?.querySelectorAll(".rule-component-row") || [])];
    rows.forEach(row => {
      const button = row.querySelector("[data-remove-component]");
      if (button) button.disabled = rows.length <= 2;
    });
  }

  function renderRules() {
    $("rule-count").textContent = String(state.rules.length);
    $("rules-list").innerHTML = state.rules.length ? state.rules.map(rule => `
      <article class="asset-rule-card">
        <div><strong>${escapeHtml(rule.name)}</strong><p>${escapeHtml(rule.components.map(part => `${part.quantity} × ${part.asset}`).join(" + "))} → ${formatNumber(rule.outputQty)} × ${escapeHtml(rule.outputAsset)} per batch</p></div>
        <button type="button" class="secondary danger small-button" data-disable-rule="${escapeAttr(rule.id)}">Remove</button>
      </article>`).join("") : `<div class="client-cards-empty"><strong>No combination rules yet</strong><span>Add a rule to combine asset types.</span></div>`;
  }

  function renderCombinePreview() {
    const rule = state.rules.find(item => item.id === $("combine-rule")?.value);
    const rawBatches = Number($("combine-batches")?.value || 1);
    const batches = Number.isInteger(rawBatches) && rawBatches > 0 ? rawBatches : 0;
    if (!rule) { $("combine-preview").textContent = "Select a rule to preview component usage and output."; return; }
    const lines = rule.components.map(part => {
      const total = part.quantity * batches;
      const item = findAsset(part.asset);
      return `<li>${formatNumber(total)} × ${escapeHtml(part.asset)} required <span class="muted">(stock: ${formatNumber(item?.balance ?? 0)})</span></li>`;
    }).join("");
    const output = rule.outputQty * batches;
    const item = findAsset(rule.outputAsset);
    $("combine-preview").innerHTML = `<strong>${batches ? formatNumber(output) : "—"} × ${escapeHtml(rule.outputAsset)} will be created</strong><ul>${lines}</ul><div>Current ${escapeHtml(rule.outputAsset)} stock: ${formatNumber(item?.balance ?? 0)}</div>`;
  }

  async function saveRule(event) {
    event.preventDefault();
    if (pageBusy || operationBusy) return;
    const name = String($("rule-name").value || "").trim();
    const outputAsset = String($("rule-output-asset").value || "").trim();
    const outputQty = Number($("rule-output-qty").value);
    const componentRows = [...$("rule-components").querySelectorAll(".rule-component-row")];
    const components = componentRows.map(row => ({ asset: String(row.querySelector("[data-component-asset]")?.value || "").trim(), quantity: Number(row.querySelector("[data-component-quantity]")?.value) }));
    if (!name || !outputAsset) { setStatus("rule-status", "Enter a rule name and output asset.", true); return; }
    if (!Number.isInteger(outputQty) || outputQty < 1) { setStatus("rule-status", "Output quantity per batch must be a whole number greater than zero.", true); return; }
    if (components.length < 2 || components.some(item => !item.asset || !Number.isInteger(item.quantity) || item.quantity < 1)) { setStatus("rule-status", "Add at least two components with whole-number quantities.", true); return; }
    const componentNames = components.map(item => item.asset.toLowerCase());
    if (new Set(componentNames).size !== componentNames.length) { setStatus("rule-status", "Each component asset can only appear once in a rule.", true); return; }
    if (componentNames.includes(outputAsset.toLowerCase())) { setStatus("rule-status", "The output asset cannot also be one of its components.", true); return; }
    operationBusy = true;
    try {
      setStatus("rule-status", "Saving rule…");
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const rows = await getValues(id, CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME);
      const ruleId = `RULE-${Date.now().toString(36).toUpperCase()}`;
      const now = new Date().toISOString();
      const dataRows = components.map(part => [ruleId, name, outputAsset, outputQty, part.asset, part.quantity, true, now]);
      const start = Math.max(rows.length, 1) + 1;
      await sheetsPost(`/${encodeURIComponent(id)}/values:batchUpdate`, { valueInputOption: "USER_ENTERED", data: [{ range: `'${CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME}'!A${start}:H${start + dataRows.length - 1}`, majorDimension: "ROWS", values: dataRows }] });
      $("rule-form").reset();
      $("rule-output-qty").value = "1";
      $("rule-components").innerHTML = "";
      addRuleComponentRow(); addRuleComponentRow();
      setStatus("rule-status", "Combination rule saved.");
      await loadPage();
    } catch (error) {
      console.error(error);
      setStatus("rule-status", error.message || "Could not save the rule.", true);
    } finally {
      operationBusy = false;
    }
  }

  async function deactivateRule(ruleId) {
    if (pageBusy || operationBusy) return;
    const rule = state.rules.find(item => item.id === ruleId);
    if (!rule || !window.confirm(`Remove the rule “${rule.name}”? Past combination records and inventory balances will not be changed.`)) return;
    operationBusy = true;
    try {
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const data = rule.rowNumbers.map(rowNumber => ({ range: `'${CONFIG.ASSET_COMBINATION_RULES_SHEET_NAME}'!G${rowNumber}`, majorDimension: "ROWS", values: [[false]] }));
      await sheetsPost(`/${encodeURIComponent(id)}/values:batchUpdate`, { valueInputOption: "USER_ENTERED", data });
      await loadPage();
    } catch (error) {
      console.error(error);
      setSyncStatus(error.message || "Could not remove the rule.", true);
    } finally {
      operationBusy = false;
    }
  }

  async function combineAssets(event) {
    event.preventDefault();
    if (pageBusy || operationBusy) return;
    const rule = state.rules.find(item => item.id === $("combine-rule").value);
    const batches = Number($("combine-batches").value);
    if (!rule) { setStatus("combine-status", "Choose a combination rule.", true); return; }
    if (!Number.isInteger(batches) || batches < 1) { setStatus("combine-status", "Enter a whole number of batches greater than zero.", true); return; }

    const changes = new Map();
    for (const part of rule.components) {
      const item = findAsset(part.asset);
      if (!item) { setStatus("combine-status", `Component ${part.asset} is no longer in inventory. Refresh and update the rule.`, true); return; }
      const quantity = part.quantity * batches;
      const key = item.asset.toLowerCase();
      const change = changes.get(key) || { item, delta: 0, consumed: 0, produced: 0 };
      change.delta -= quantity;
      change.consumed += quantity;
      changes.set(key, change);
    }
    const outputItem = findAsset(rule.outputAsset);
    if (!outputItem) { setStatus("combine-status", `Output asset ${rule.outputAsset} is no longer in inventory. Refresh and update the rule.`, true); return; }
    const outputKey = outputItem.asset.toLowerCase();
    if (changes.has(outputKey)) { setStatus("combine-status", "The output asset cannot also be a component in this rule.", true); return; }
    const outputQty = rule.outputQty * batches;
    changes.set(outputKey, { item: outputItem, delta: outputQty, consumed: 0, produced: outputQty });

    for (const value of changes.values()) {
      const next = value.item.balance + value.delta;
      if (next < 0) { setStatus("combine-status", `Not enough ${value.item.asset}. Available in inventory: ${formatNumber(value.item.balance)}; required: ${formatNumber(value.consumed)}.`, true); return; }
    }

    const usage = rule.components.map(part => `${formatNumber(part.quantity * batches)} ${part.asset}`).join(" + ");
    const confirmText = `Combine ${usage} to create ${formatNumber(outputQty)} ${rule.outputAsset}?\n\nThis will change warehouse stock and add an audit entry.`;
    if (!window.confirm(confirmText)) return;

    operationBusy = true;
    try {
      setStatus("combine-status", "Combining assets…");
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const latestRows = await getValues(id, CONFIG.INVENTORY_SHEET_NAME);
      const latestInventory = parseInventory(latestRows);
      const latestByKey = new Map(latestInventory.map(item => [item.asset.toLowerCase(), item]));
      for (const part of rule.components) {
        const key = part.asset.toLowerCase();
        const latest = latestByKey.get(key);
        const need = part.quantity * batches;
        if (!latest || latest.balance < need) throw new Error(`Stock changed while you were preparing this conversion. Refresh and try again for ${part.asset}.`);
        const planned = changes.get(key);
        planned.item = latest;
      }
      const latestOutput = latestByKey.get(rule.outputAsset.toLowerCase());
      if (!latestOutput) throw new Error(`Output asset ${rule.outputAsset} is missing. Refresh and try again.`);
      changes.get(outputKey).item = latestOutput;
      for (const value of changes.values()) {
        const sourceBalance = value.item.balance;
        value.newBalance = sourceBalance + value.delta;
        if (value.newBalance < 0) throw new Error(`Not enough ${value.item.asset}. Refresh inventory and try again.`);
      }
      const timestamp = new Date().toISOString();
      const user = state.idTokenPayload?.email || readSession()?.email || "";
      const operationId = `COMBINE-${Date.now().toString(36).toUpperCase()}`;
      const comment = `Rule: ${rule.name}. ${usage} → ${formatNumber(outputQty)} ${rule.outputAsset}.`;
      const operationRecords = [];
      for (const part of rule.components) {
        const item = changes.get(part.asset.toLowerCase());
        operationRecords.push([timestamp, "Warehouse", "COMBINE_OUT", item.item.asset, part.quantity * batches, user, comment, "", item.newBalance, operationId]);
      }
      const output = changes.get(outputKey);
      operationRecords.push([timestamp, "Warehouse", "COMBINE_IN", output.item.asset, outputQty, user, comment, "", output.newBalance, operationId]);
      const opRows = await getValues(id, CONFIG.ASSET_OPERATIONS_SHEET_NAME);
      const startOps = Math.max(opRows.length, 1) + 1;
      const data = [];
      for (const value of changes.values()) {
        data.push({ range: `'${CONFIG.INVENTORY_SHEET_NAME}'!${columnLetter(value.item.balanceColumn)}${value.item.rowNumber}`, majorDimension: "ROWS", values: [[value.newBalance]] });
      }
      data.push({ range: `'${CONFIG.ASSET_OPERATIONS_SHEET_NAME}'!A${startOps}:J${startOps + operationRecords.length - 1}`, majorDimension: "ROWS", values: operationRecords });
      await sheetsPost(`/${encodeURIComponent(id)}/values:batchUpdate`, { valueInputOption: "USER_ENTERED", data });
      $("combine-batches").value = "1";
      await loadPage();
      setStatus("combine-status", `Combined successfully: ${formatNumber(outputQty)} ${rule.outputAsset} created.`);
    } catch (error) {
      console.error(error);
      setStatus("combine-status", error.message || "Could not combine the selected assets.", true);
    } finally {
      operationBusy = false;
    }
  }

  function updateCurrentClientBalance() {
    const client = String($("normalise-client")?.value || "").trim();
    const asset = String($("normalise-asset")?.value || "").trim();
    if (!client || !asset) { if ($("normalise-current-balance")) $("normalise-current-balance").textContent = "—"; return; }
    const balance = clientBalance(client, asset, state.standardTransactions, state.adjustments);
    $("normalise-current-balance").textContent = formatSigned(balance);
    if ($("normalise-target-balance") && !$("normalise-target-balance").dataset.userEdited) $("normalise-target-balance").value = String(balance);
  }

  function clientBalance(client, asset, transactions, adjustments) {
    const clientKey = client.trim().toLowerCase();
    const assetKey = asset.trim().toLowerCase();
    let balance = 0;
    (transactions || []).forEach(row => {
      if (String(row.client || "").trim().toLowerCase() !== clientKey || String(row.asset || "").trim().toLowerCase() !== assetKey) return;
      const movement = String(row.movement || "").trim().toUpperCase();
      const quantity = Number(row.quantity) || 0;
      if (movement === "SENT") balance += quantity;
      else if (movement === "RECEIVED") balance -= quantity;
    });
    (adjustments || []).forEach(row => {
      if (String(row.client || "").trim().toLowerCase() === clientKey && String(row.asset || "").trim().toLowerCase() === assetKey) balance += Number(row.quantity) || 0;
    });
    return balance;
  }

  async function normaliseClientBalance(event) {
    event.preventDefault();
    if (pageBusy || operationBusy) return;
    const client = String($("normalise-client").value || "").trim();
    const asset = String($("normalise-asset").value || "").trim();
    const target = Number($("normalise-target-balance").value);
    const reason = String($("normalise-reason").value || "").trim();
    if (!client || !asset || !reason) { setStatus("normalise-status", "Select a client and asset, then enter a reason.", true); return; }
    if (!Number.isInteger(target)) { setStatus("normalise-status", "The corrected balance must be a whole number. Use a minus sign for a negative balance.", true); return; }
    if (reason.length < 5) { setStatus("normalise-status", "Please enter a clear reason for the correction.", true); return; }
    const current = clientBalance(client, asset, state.standardTransactions, state.adjustments);
    const delta = target - current;
    if (delta === 0) { setStatus("normalise-status", "The ledger is already at that balance. No adjustment is needed.", true); return; }
    const confirmText = `Client: ${client}\nAsset: ${asset}\nCurrent balance: ${formatSigned(current)}\nCorrected balance: ${formatSigned(target)}\nAdjustment: ${delta > 0 ? "+" : ""}${formatNumber(delta)}\n\nReason: ${reason}\n\nThis will change the client ledger only. Warehouse inventory will not change.`;
    if (!window.confirm(confirmText)) return;
    operationBusy = true;
    try {
      setStatus("normalise-status", "Saving ledger correction…");
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const adjustmentRows = await getValues(id, CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME);
      const start = Math.max(adjustmentRows.length, 1) + 1;
      const timestamp = new Date().toISOString();
      const user = state.idTokenPayload?.email || readSession()?.email || "";
      const adjustmentId = `ADJUST-${Date.now().toString(36).toUpperCase()}`;
      const comment = `Normalised client balance from ${formatSigned(current)} to ${formatSigned(target)}. Reason: ${reason}`;
      const row = [timestamp, client, "NORMALISATION", asset, delta, user, comment, "", "", target, current, adjustmentId];
      await sheetsPost(`/${encodeURIComponent(id)}/values:batchUpdate`, {
        valueInputOption: "USER_ENTERED",
        data: [{ range: `'${CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME}'!A${start}:L${start}`, majorDimension: "ROWS", values: [row] }]
      });
      $("normalise-reason").value = "";
      $("normalise-target-balance").dataset.userEdited = "";
      await loadPage();
      setStatus("normalise-status", `Client balance updated from ${formatSigned(current)} to ${formatSigned(target)}. Warehouse inventory is unchanged.`);
    } catch (error) {
      console.error(error);
      setStatus("normalise-status", error.message || "Could not save the client-ledger adjustment.", true);
    } finally {
      operationBusy = false;
    }
  }

  function renderAdjustmentHistory() {
    const rows = (state.adjustments || []).slice().sort((a, b) => dateMs(b.timestamp) - dateMs(a.timestamp)).slice(0, 8);
    $("adjustment-history").innerHTML = rows.length ? rows.map(row => `
      <article class="history-row"><div><strong>${escapeHtml(row.client)} · ${escapeHtml(row.asset)}</strong><p>${escapeHtml(row.comment || "Ledger balance corrected")}</p></div><div class="history-quantity">${Number(row.quantity) > 0 ? "+" : ""}${formatNumber(row.quantity)}</div><div class="history-meta">${escapeHtml(formatDateTime(row.timestamp))} · ${escapeHtml(row.user || "")}</div></article>`).join("") : `<div class="client-cards-empty"><span>No normalisations recorded.</span></div>`;
  }

  function renderOperationHistory() {
    const rows = (state.operations || []).slice().sort((a, b) => dateMs(b.timestamp) - dateMs(a.timestamp));
    const groups = new Map();
    rows.forEach(row => {
      const key = `${row.timestamp}|${row.comment}`;
      if (!groups.has(key)) groups.set(key, { timestamp: row.timestamp, comment: row.comment, items: [] });
      groups.get(key).items.push(row);
    });
    const operations = [...groups.values()].slice(0, 8);
    $("operation-history").innerHTML = operations.length ? operations.map(group => {
      const inputs = group.items.filter(item => String(item.movement).toUpperCase() === "COMBINE_OUT");
      const outputs = group.items.filter(item => String(item.movement).toUpperCase() === "COMBINE_IN");
      const detail = `${inputs.map(item => `${formatNumber(item.quantity)} ${item.asset}`).join(" + ")} → ${outputs.map(item => `${formatNumber(item.quantity)} ${item.asset}`).join(" + ")}`;
      return `<article class="history-row"><div><strong>${escapeHtml(detail || "Asset combination")}</strong><p>${escapeHtml(group.comment || "Asset conversion")}</p></div><div class="history-meta">${escapeHtml(formatDateTime(group.timestamp))}</div></article>`;
    }).join("") : `<div class="client-cards-empty"><span>No combinations recorded.</span></div>`;
  }

  function setUserProfile(profile) {
    const rawName = String(profile?.given_name || profile?.name || "Google user").trim();
    const firstName = (rawName.split(/\s+/)[0] || "there").replace(/[^\p{L}\p{M}'-]/gu, "");
    $("user-name").textContent = `Hi, ${firstName || "there"} 👋`;
    $("user-email").textContent = "";
    if (profile?.picture) { $("user-photo").src = profile.picture; $("user-photo").classList.remove("hidden"); }
  }

  function hideLogin() {
    $("google-signin-button")?.classList.add("hidden");
    $("grant-access")?.classList.add("hidden");
    $("sign-out")?.classList.remove("hidden");
    $("login-card")?.classList.add("hidden");
    $("tools-page")?.classList.remove("hidden");
  }

  function signOut() {
    const saved = readSession();
    try { if (state.idTokenPayload?.sub) google.accounts.id.revoke(state.idTokenPayload.sub, () => {}); } catch (_) {}
    try { if (saved?.sub && saved.sub !== state.idTokenPayload?.sub) google.accounts.id.revoke(saved.sub, () => {}); } catch (_) {}
    try { localStorage.removeItem(SESSION_KEY); } catch (_) {}
    window.FM_AUTH_CACHE?.clear?.();
    state.accessToken = null;
    state.idTokenPayload = null;
    $("tools-page")?.classList.add("hidden");
    $("login-card")?.classList.remove("hidden");
    $("google-signin-button")?.classList.remove("hidden");
    $("grant-access")?.classList.add("hidden");
    $("sign-out")?.classList.add("hidden");
    $("user-photo")?.classList.add("hidden");
    $("user-name").textContent = "Not signed in";
  }

  function saveSession() {
    const p = state.idTokenPayload;
    if (!p) return;
    try { localStorage.setItem(SESSION_KEY, JSON.stringify({ name: p.name || "Google user", email: p.email || "", picture: p.picture || "", sub: p.sub || "" })); } catch (_) {}
  }
  function readSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; } }
  function decodeJwtPayload(token) { const parts = String(token || "").split("."); if (parts.length !== 3) throw new Error("Invalid Google credential."); return JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"))); }

  async function getValues(id, sheet) {
    const range = `${quoteSheetName(sheet)}!A:AE`;
    const data = await sheetsGet(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}`);
    return data.values || [];
  }
  async function updateValues(id, sheet, rows) {
    const range = `${quoteSheetName(sheet)}!A1`;
    return sheetsPut(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, { range, majorDimension: "ROWS", values: rows });
  }
  async function sheetsGet(path) { return fetchJson(SHEETS_API + path, { headers: authHeaders() }); }
  async function sheetsPost(path, body) { return fetchJson(SHEETS_API + path, { method: "POST", headers: { ...authHeaders(), "Content-Type": "application/json" }, body: JSON.stringify(body) }); }
  async function sheetsPut(path, body) { return fetchJson(SHEETS_API + path, { method: "PUT", headers: { ...authHeaders(), "Content-Type": "application/json" }, body: JSON.stringify(body) }); }
  function authHeaders() { return { Authorization: `Bearer ${state.accessToken}` }; }

  async function fetchJson(url, options = {}) {
    const response = await fetch(url, options);
    if (!response.ok) {
      let message = `Google Sheets request failed (${response.status}).`;
      try { const body = await response.json(); if (body?.error?.message) message = body.error.message; } catch (_) {}
      if (response.status === 401) { window.FM_AUTH_CACHE?.clear?.(); }
      throw new Error(message);
    }
    if (response.status === 204) return {};
    return response.json();
  }

  function findAsset(name) { return state.inventory.find(item => item.asset.toLowerCase() === String(name || "").trim().toLowerCase()) || null; }
  function normalizeHeader(value) { return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " "); }
  function findHeader(headers, candidates) { const normalized = candidates.map(normalizeHeader); return headers.findIndex(value => normalized.includes(value)); }
  function numberValue(value) { const n = Number(String(value ?? "0").replace(/,/g, "").trim()); return Number.isFinite(n) ? n : 0; }
  function positiveInteger(value, fallback = 1) { const n = Number(value); return Number.isInteger(n) && n > 0 ? n : fallback; }
  function truthy(value) { return ["TRUE", "YES", "Y", "1", "ACTIVE"].includes(String(value ?? "").trim().toUpperCase()); }
  function columnLetter(number) { let n = Number(number); let text = ""; while (n > 0) { const rem = (n - 1) % 26; text = String.fromCharCode(65 + rem) + text; n = Math.floor((n - 1) / 26); } return text || "A"; }
  function quoteSheetName(name) { return `'${String(name).replace(/'/g, "''")}'`; }
  function formatNumber(value) { return (Number(value) || 0).toLocaleString("en-GB", { maximumFractionDigits: 0 }); }
  function formatSigned(value) { const n = Number(value) || 0; return `${n > 0 ? "+" : ""}${formatNumber(n)}`; }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char])); }
  function escapeAttr(value) { return escapeHtml(value); }
  function dateMs(value) { const d = new Date(value); return Number.isNaN(d.getTime()) ? 0 : d.getTime(); }
  function formatDateTime(value) { const d = new Date(value); return Number.isNaN(d.getTime()) ? String(value || "") : d.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }); }
  function setSyncStatus(message, error = false) { const el = $("sync-status"); if (!el) return; el.textContent = message; el.className = error ? "muted error" : "muted"; }
  function setStatus(id, message, error = false) { const el = $(id); if (!el) return; el.textContent = message || ""; el.className = `status${error ? " error" : ""}`; }

  // Keep target balance editable after the user starts entering a correction.
  $("normalise-target-balance")?.addEventListener("input", () => { $("normalise-target-balance").dataset.userEdited = "1"; });
})();
