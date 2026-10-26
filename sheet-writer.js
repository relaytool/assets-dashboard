(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const LEGACY_CHANGE_LOG_SHEET = "Inventory Change Log";
  const TRANSACTION_BALANCE_HEADER = "Inventory Balance";

  function normalizeHeader(value) {
    return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  function quoteSheetName(value) {
    return `'${String(value).replace(/'/g, "''")}'`;
  }

  function columnLetter(value) {
    let n = Number(value) || 0;
    let result = "";
    while (n > 0) {
      const remainder = (n - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result || "A";
  }

  function findColumn(headers, aliases) {
    const wanted = aliases.map(normalizeHeader);
    return headers.findIndex(header => wanted.includes(normalizeHeader(header)));
  }

  function isTransientStatus(status) {
    return [408, 429, 500, 502, 503, 504].includes(Number(status));
  }

  async function request(url, accessToken, options = {}) {
    const method = String(options.method || "GET").toUpperCase();
    const safeRetry = options.safeRetry !== false && ["GET", "HEAD"].includes(method);
    const maxRetries = safeRetry ? 2 : 0;
    let response = null;

    for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
      response = await fetch(url, {
        ...options,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          ...(options.headers || {})
        }
      });

      if (response.ok) {
        const raw = await response.text();
        let data = {};
        try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
        return data;
      }

      if (!safeRetry || !isTransientStatus(response.status) || attempt >= maxRetries) break;
      const retryAfter = Number(response.headers?.get?.("Retry-After"));
      const delay = Number.isFinite(retryAfter) && retryAfter > 0
        ? Math.min(6000, retryAfter * 1000)
        : 500 * (attempt + 1);
      await new Promise(resolve => setTimeout(resolve, delay));
    }

    const raw = await response.text();
    let data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
    const error = new Error(data?.error?.message || `Google Sheets request failed (${response.status})`);
    error.status = response.status;
    error.reason = data?.error?.errors?.[0]?.reason || data?.error?.status || "";
    throw error;
  }

  async function sheetsGet(spreadsheetId, range, accessToken) {
    return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`, accessToken);
  }

  async function sheetsBatchUpdate(spreadsheetId, body, accessToken) {
    return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, accessToken, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      safeRetry: false
    });
  }

  async function sheetsPut(spreadsheetId, range, values, accessToken) {
    return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, accessToken, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ range, majorDimension: "ROWS", values }),
      safeRetry: false
    });
  }

  function sheetByTitle(book, title) {
    return (book?.sheets || []).find(sheet => sheet?.properties?.title === title)?.properties || null;
  }

  function cellValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) return { userEnteredValue: { numberValue: value } };
    return { userEnteredValue: { stringValue: String(value ?? "") } };
  }

  function updateCellsRequest(sheetId, rowNumber, columnNumber, value) {
    return {
      updateCells: {
        start: {
          sheetId,
          rowIndex: Math.max(0, Number(rowNumber) - 1),
          columnIndex: Math.max(0, Number(columnNumber) - 1)
        },
        rows: [{ values: [cellValue(value)] }],
        fields: "userEnteredValue"
      }
    };
  }

  function appendCellsRequest(sheetId, rows) {
    return {
      appendCells: {
        sheetId,
        fields: "userEnteredValue",
        rows: rows.map(row => ({ values: row.map(cellValue) }))
      }
    };
  }

  async function removeLegacyChangeLog({ spreadsheetId, accessToken }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const legacy = sheetByTitle(book, LEGACY_CHANGE_LOG_SHEET);
    if (!legacy?.sheetId) return false;

    await sheetsBatchUpdate(spreadsheetId, {
      requests: [{ deleteSheet: { sheetId: legacy.sheetId } }]
    }, accessToken);
    return true;
  }

  async function readInventory({ spreadsheetId, accessToken }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = sheetByTitle(book, CONFIG.INVENTORY_SHEET_NAME);
    if (!inventorySheet?.sheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    const response = await sheetsGet(spreadsheetId, `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!A:AE`, accessToken);
    const rows = response?.values || [];
    const header = rows[0] || [];
    const normalized = header.map(normalizeHeader);

    if (normalized.includes("asset id")) {
      throw new Error(`"${CONFIG.INVENTORY_SHEET_NAME}" contains an unsupported "Asset ID" column. Remove that column before recording movements.`);
    }

    const assetIdx = findColumn(header, ["asset", "asset name", "item", "type"]);
    const balanceIdx = findColumn(header, ["balance", "current balance", "stock", "quantity"]);
    if (assetIdx < 0 || balanceIdx < 0) {
      throw new Error(`"${CONFIG.INVENTORY_SHEET_NAME}" must contain Asset and Balance columns.`);
    }

    const items = rows.slice(1).map((row, index) => ({
      rowNumber: index + 2,
      asset: String(row?.[assetIdx] ?? "").trim(),
      balance: Number(String(row?.[balanceIdx] ?? "").replace(/,/g, "")) || 0,
      assetColumn: assetIdx + 1,
      balanceColumn: balanceIdx + 1
    })).filter(item => item.asset);

    return { book, sheetId: inventorySheet.sheetId, header, items };
  }

  async function ensureTransactionBalanceColumn({ spreadsheetId, accessToken }) {
    const range = `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A1:AE1`;
    const response = await sheetsGet(spreadsheetId, range, accessToken);
    const header = response?.values?.[0] || [
      "Timestamp", "Client", "Movement", "Asset", "Quantity", "User", "Comment", "Image Link"
    ];
    const normalized = header.map(normalizeHeader);
    if (!normalized.includes("inventory balance")) {
      header.push(TRANSACTION_BALANCE_HEADER);
      await sheetsPut(spreadsheetId, `${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A1`, [header], accessToken);
    }
    return header;
  }

  function buildTransactionRow(header, movement, inventoryBalance) {
    const normalized = header.map(normalizeHeader);
    const row = new Array(header.length).fill("");
    const put = (aliases, value) => {
      const index = findColumn(normalized, aliases);
      if (index >= 0) row[index] = value ?? "";
    };

    put(["timestamp", "date", "datetime"], movement.timestamp);
    put(["client", "client name"], movement.client);
    put(["movement", "type", "direction"], movement.movement);
    put(["asset", "asset name", "item"], movement.asset);
    put(["quantity", "qty"], Number(movement.quantity) || 0);
    put(["user", "entered by", "email"], movement.user);
    put(["comment", "comments", "notes"], movement.comment);
    put(["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"], movement.image);
    put(["inventory balance", "warehouse balance", "balance after movement"], inventoryBalance);
    return row;
  }

  async function appendMovementBatch({
    spreadsheetId,
    accessToken,
    movements = [],
    timestamp,
    user = "Google user"
  }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");
    if (!movements.length) throw new Error("No movements were supplied.");

    // Remove the retired log before doing anything else. This is deliberately
    // a one-time delete, not a recreation/migration path.
    try { await removeLegacyChangeLog({ spreadsheetId, accessToken }); } catch (_) {}

    const inventory = await readInventory({ spreadsheetId, accessToken });
    const transactionHeader = await ensureTransactionBalanceColumn({ spreadsheetId, accessToken });
    const metadata = inventory.book;
    const transactionSheet = sheetByTitle(metadata, CONFIG.TRANSACTIONS_SHEET_NAME);
    if (!transactionSheet?.sheetId) throw new Error(`Sheet "${CONFIG.TRANSACTIONS_SHEET_NAME}" was not found.`);

    const byAsset = new Map(inventory.items.map(item => [normalizeHeader(item.asset), { ...item }]));
    const rows = [];
    const balanceWrites = [];
    const resultingMovements = [];

    for (const movement of movements) {
      const key = normalizeHeader(movement.asset);
      const item = byAsset.get(key);
      if (!item) throw new Error(`Asset "${movement.asset}" is not present in Inventory.`);

      const quantity = Math.abs(Number(movement.quantity) || 0);
      if (!quantity) throw new Error(`Quantity for "${item.asset}" must be greater than zero.`);

      const current = Number(item.balance) || 0;
      const next = String(movement.movement || "").toUpperCase() === "RECEIVED"
        ? current + quantity
        : current - quantity;

      if (String(movement.movement || "").toUpperCase() !== "RECEIVED" && next < 0) {
        throw new Error(`Cannot send ${quantity} ${item.asset}. Current warehouse balance is ${current}.`);
      }

      item.balance = next;
      rows.push(buildTransactionRow(transactionHeader, {
        timestamp: movement.timestamp || timestamp || new Date().toISOString(),
        client: movement.client,
        movement: String(movement.movement || "").toUpperCase(),
        asset: item.asset,
        quantity,
        user: movement.user || user,
        comment: movement.comment || "",
        image: movement.image || ""
      }, next));
      balanceWrites.push(updateCellsRequest(inventory.sheetId, item.rowNumber, item.balanceColumn, next));
      resultingMovements.push({ ...movement, asset: item.asset, quantity, inventoryBalance: next });
    }

    // One write request updates the warehouse and appends the exact rows that
    // describe those updates. There is no change-log sheet and no balance formula.
    await sheetsBatchUpdate(spreadsheetId, {
      requests: [
        appendCellsRequest(transactionSheet.sheetId, rows),
        ...balanceWrites
      ]
    }, accessToken);

    return { movements: resultingMovements };
  }

  async function appendAsset({ spreadsheetId, accessToken, assetName, startingBalance = 0 }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");
    try { await removeLegacyChangeLog({ spreadsheetId, accessToken }); } catch (_) {}

    const inventory = await readInventory({ spreadsheetId, accessToken });
    const cleanName = String(assetName || "").trim();
    if (!cleanName) throw new Error("Enter an asset type.");
    if (inventory.items.some(item => normalizeHeader(item.asset) === normalizeHeader(cleanName))) {
      throw new Error(`Asset "${cleanName}" already exists in Inventory.`);
    }

    const row = new Array(Math.max(inventory.header.length, 2)).fill("");
    row[inventory.items[0]?.assetColumn ? inventory.items[0].assetColumn - 1 : findColumn(inventory.header, ["asset"]) ] = cleanName;
    const balanceColumn = inventory.items[0]?.balanceColumn || findColumn(inventory.header, ["balance"]) + 1;
    row[balanceColumn - 1] = Math.max(0, Number(startingBalance) || 0);

    await sheetsBatchUpdate(spreadsheetId, {
      requests: [appendCellsRequest(inventory.sheetId, [row])]
    }, accessToken);
    return true;
  }

  async function migrateRename({ spreadsheetId, accessToken, newName, rowNumber, assetColumn = 1 }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = sheetByTitle(book, CONFIG.INVENTORY_SHEET_NAME);
    if (!inventorySheet?.sheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);
    await sheetsBatchUpdate(spreadsheetId, {
      requests: [updateCellsRequest(inventorySheet.sheetId, rowNumber, assetColumn, String(newName || "").trim())]
    }, accessToken);
    return true;
  }

  window.FM_SHEET_WRITER = Object.freeze({
    LEGACY_CHANGE_LOG_SHEET,
    TRANSACTION_BALANCE_HEADER,
    removeLegacyChangeLog,
    ensureTransactionBalanceColumn,
    appendMovementBatch,
    appendAsset,
    migrateRename
  });
})();
