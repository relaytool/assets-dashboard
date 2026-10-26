(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const CHANGE_LOG_SHEET_NAME = "Inventory Change Log";
  const CHANGE_LOG_HEADERS = [
    "Timestamp",
    "Event ID",
    "Asset",
    "Inventory Row",
    "Delta",
    "Source",
    "Movement",
    "Quantity",
    "Client",
    "User",
    "Transaction Event ID"
  ];

  const INVENTORY_HEADERS = ["Asset", "Balance"];

  const api = {
    async get(spreadsheetId, range, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`, accessToken);
    },
    async clear(spreadsheetId, range, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}:clear`, accessToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}"
      });
    },
    async updateValues(spreadsheetId, range, values, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, accessToken, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          range,
          majorDimension: "ROWS",
          values
        })
      });
    },
    async batchUpdate(spreadsheetId, body, accessToken) {
      return request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, accessToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
    }
  };

  async function request(url, accessToken, options = {}) {
    const response = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...(options.headers || {})
      }
    });
    const raw = await response.text();
    let data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
    if (!response.ok) {
      throw new Error(data?.error?.message || `Google Sheets request failed (${response.status})`);
    }
    return data;
  }

  function quoteSheetName(value) {
    return `'${String(value).replace(/'/g, "''")}'`;
  }

  function normalizeHeader(value) {
    return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
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

  function randomEventId(prefix = "EVT") {
    const time = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).slice(2, 8).toUpperCase();
    return `${prefix}-${time}-${random}`;
  }

  function cellValue(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return { userEnteredValue: { numberValue: value } };
    }
    return { userEnteredValue: { stringValue: String(value ?? "") } };
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

  function deleteColumnRequest(sheetId, columnIndex) {
    return {
      deleteDimension: {
        range: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: Math.max(0, Number(columnIndex) || 0),
          endIndex: Math.max(0, Number(columnIndex) || 0) + 1
        }
      }
    };
  }

  function balanceFormula(assetColumn = 1) {
    const assetCol = columnLetter(assetColumn);
    return `=SUMIF(${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!$C:$C,INDIRECT("${assetCol}"&ROW()),${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!$E:$E)`;
  }

  function formulaCell(formula) {
    return { userEnteredValue: { formulaValue: formula } };
  }

  function updateBalanceCellRequest(sheetId, item) {
    return {
      updateCells: {
        start: {
          sheetId,
          rowIndex: Math.max(0, Number(item.rowNumber) - 1),
          columnIndex: Math.max(0, Number(item.balanceColumn) - 1)
        },
        rows: [{ values: [formulaCell(balanceFormula(Number(item.assetColumn) || 1))] }],
        fields: "userEnteredValue"
      }
    };
  }

  function buildBaselineRow(item, user, timestamp, eventId) {
    return [
      timestamp,
      eventId,
      String(item.asset || "").trim(),
      Number(item.rowNumber) || 0,
      Number(item.balance) || 0,
      "BASELINE",
      "BASELINE",
      Number(item.balance) || 0,
      "",
      user || "System",
      ""
    ];
  }

  function buildChangeRow(item, movement, quantity, client, user, timestamp, transactionEventId) {
    const delta = movement === "RECEIVED" ? Math.abs(Number(quantity) || 0) : -Math.abs(Number(quantity) || 0);
    return [
      timestamp,
      randomEventId("BAL"),
      String(item.asset || "").trim(),
      Number(item.rowNumber) || 0,
      delta,
      "MOVEMENT",
      String(movement || "").trim(),
      Math.abs(Number(quantity) || 0),
      String(client || "").trim(),
      user || "Google user",
      transactionEventId || ""
    ];
  }

  function metadataSheetInfo(metadata, title) {
    return (metadata?.sheets || []).find(sheet => sheet?.properties?.title === title)?.properties || null;
  }

  function parseInventoryRows(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return [];

    const header = safeRows[0].map(normalizeHeader);
    const assetIdx = ["asset", "asset name", "item", "type"]
      .map(name => header.indexOf(name))
      .find(index => index >= 0);
    const balanceIdx = ["balance", "current balance", "stock", "quantity"]
      .map(name => header.indexOf(name))
      .find(index => index >= 0);

    if (assetIdx === undefined || assetIdx < 0) return [];

    return safeRows.slice(1).map((row, index) => ({
      rowNumber: index + 2,
      asset: String(row?.[assetIdx] ?? "").trim(),
      balance: balanceIdx !== undefined && balanceIdx >= 0 ? Number(String(row?.[balanceIdx] ?? "").replace(/,/g, "")) || 0 : 0,
      assetColumn: assetIdx + 1,
      balanceColumn: balanceIdx !== undefined && balanceIdx >= 0 ? balanceIdx + 1 : 2
    })).filter(item => item.asset);
  }

  async function ensureInventorySchema({ spreadsheetId, accessToken }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");

    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME);
    if (!inventorySheet?.sheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    let response = await api.get(spreadsheetId, `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!A:AE`, accessToken);
    let rows = response?.values || [];
    if (!rows.length) {
      await api.batchUpdate(spreadsheetId, {
        requests: [{
          appendCells: {
            sheetId: inventorySheet.sheetId,
            fields: "userEnteredValue",
            rows: [{ values: INVENTORY_HEADERS.map(cellValue) }]
          }
        }]
      }, accessToken);
      response = await api.get(spreadsheetId, `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!A:AE`, accessToken);
      rows = response?.values || [];
    }

    const originalHeader = (rows[0] || []).slice();
    const header = originalHeader.map(normalizeHeader);
    const assetIdIdx = header.indexOf("asset id");

    if (assetIdIdx >= 0) {
      const withoutAssetId = header.filter((_, index) => index !== assetIdIdx);
      const hasAsset = withoutAssetId.some(value => ["asset", "asset name", "item", "type"].includes(value));
      const hasBalance = withoutAssetId.some(value => ["balance", "current balance", "stock", "quantity"].includes(value));

      if (!hasAsset || !hasBalance) {
        throw new Error(`The "${CONFIG.INVENTORY_SHEET_NAME}" sheet contains an "Asset ID" column but its Asset/Balance columns could not be identified safely. No data was changed.`);
      }

      try {
        await api.batchUpdate(spreadsheetId, {
          requests: [deleteColumnRequest(inventorySheet.sheetId, assetIdIdx)]
        }, accessToken);
      } catch (error) {
        const latest = await api.get(spreadsheetId, `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!A1:AE1`, accessToken);
        const latestHeader = (latest?.values?.[0] || []).map(normalizeHeader);
        if (latestHeader.includes("asset id")) throw error;
      }

      response = await api.get(spreadsheetId, `${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!A:AE`, accessToken);
      rows = response?.values || [];
    }

    const inventoryItems = parseInventoryRows(rows);
    if (!inventoryItems.length && (rows.length > 1 || rows[0]?.length)) {
      throw new Error(`The "${CONFIG.INVENTORY_SHEET_NAME}" sheet does not have a usable Asset and Balance column. No data was changed.`);
    }

    return {
      sheetId: inventorySheet.sheetId,
      rows,
      inventoryItems,
      changed: assetIdIdx >= 0
    };
  }

  async function repairChangeLogHeader(spreadsheetId, accessToken, logSheet, inventoryItems = []) {
    const response = await api.get(
      spreadsheetId,
      `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:ZZ1`,
      accessToken
    );
    let header = response?.values?.[0] || [];
    const normalized = header.map(normalizeHeader);
    const expectedNormalized = CHANGE_LOG_HEADERS.map(normalizeHeader);

    const firstExpected = expectedNormalized.every((name, index) => normalized[index] === name);
    const exactExpectedLength = normalized.length === expectedNormalized.length;
    if (firstExpected && exactExpectedLength) {
      const rawHeaderExact = CHANGE_LOG_HEADERS.every((name, index) => String(header[index] ?? "").trim() === name);
      if (!rawHeaderExact) {
        await api.updateValues(
          spreadsheetId,
          `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:K1`,
          [CHANGE_LOG_HEADERS],
          accessToken
        );
        return { changed: true, headerRewritten: true, header: CHANGE_LOG_HEADERS.slice() };
      }
      return { changed: false, header };
    }

    const assetIdIdx = normalized.indexOf("asset id");

    // Legacy/corrupt schema: Asset ID was inserted into the expected 11-column
    // log. Removing that one column restores Asset/Inventory Row/Delta to C/D/E.
    if (assetIdIdx >= 0) {
      const withoutAssetId = normalized.filter((_, index) => index !== assetIdIdx);
      const matchesAfterRemoval = expectedNormalized.every((name, index) => withoutAssetId[index] === name);
      if (matchesAfterRemoval) {
        try {
          await api.batchUpdate(spreadsheetId, {
            requests: [deleteColumnRequest(logSheet.sheetId, assetIdIdx)]
          }, accessToken);
        } catch (error) {
          const latest = await api.get(
            spreadsheetId,
            `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:K1`,
            accessToken
          );
          const latestHeader = latest?.values?.[0] || [];
          const latestValid = CHANGE_LOG_HEADERS.every((name, index) => String(latestHeader[index] ?? "").trim() === name);
          if (!latestValid) throw error;
        }

        return { changed: true, legacyAssetIdRemoved: true, header: CHANGE_LOG_HEADERS.slice() };
      }
    }

    // If the required 11 headers are correct and there are only trailing/extra
    // headers, clear the header cells beyond K. This does not touch transaction
    // history beneath them.
    const first11Match = expectedNormalized.every((name, index) => normalized[index] === name);
    if (first11Match) {
      let changed = false;
      const rawFirst11Exact = CHANGE_LOG_HEADERS.every((name, index) => String(header[index] ?? "").trim() === name);
      if (!rawFirst11Exact) {
        await api.updateValues(
          spreadsheetId,
          `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:K1`,
          [CHANGE_LOG_HEADERS],
          accessToken
        );
        changed = true;
      }
      if (header.length > CHANGE_LOG_HEADERS.length) {
        await api.clear(
          spreadsheetId,
          `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!${columnLetter(CHANGE_LOG_HEADERS.length + 1)}1:${columnLetter(header.length)}1`,
          accessToken
        );
        changed = true;
      }
      return { changed, trailingHeadersCleared: header.length > CHANGE_LOG_HEADERS.length, headerRewritten: !rawFirst11Exact, header: CHANGE_LOG_HEADERS.slice() };
    }

    // Safe header repair is intentionally narrow. Do not rewrite or shift an
    // unknown schema because doing so could move real audit data.
    throw new Error(`The "${CHANGE_LOG_SHEET_NAME}" sheet exists but has an unexpected header. The first row must be: ${CHANGE_LOG_HEADERS.join(", ")}. No existing log rows were changed.`);
  }

  async function ensureChangeLog({ spreadsheetId, accessToken, inventoryItems = [], user = "System" }) {
    if (!spreadsheetId || !accessToken) throw new Error("Google Sheets connection is not available.");

    const inventorySchema = await ensureInventorySchema({ spreadsheetId, accessToken });
    const resolvedInventoryItems = inventorySchema.inventoryItems;

    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME);
    if (!inventorySheet?.sheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    let logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    if (logSheet?.sheetId) {
      const repair = await repairChangeLogHeader(spreadsheetId, accessToken, logSheet, resolvedInventoryItems);

      // Re-apply the balance formulas after any schema migration or inventory
      // schema cleanup. This is the safety net that prevents old Asset ID columns
      // from leaving formulas pointing at the wrong change-log columns.
      if (repair.changed || inventorySchema.changed) {
        await api.batchUpdate(spreadsheetId, {
          requests: resolvedInventoryItems.map(item => updateBalanceCellRequest(inventorySheet.sheetId, item))
        }, accessToken);
      }

      return {
        initialized: true,
        sheetId: logSheet.sheetId,
        repaired: !!repair.changed,
        inventoryChanged: !!inventorySchema.changed
      };
    }

    const items = resolvedInventoryItems.length ? resolvedInventoryItems : (Array.isArray(inventoryItems) ? inventoryItems : []);
    const existingIds = (book.sheets || []).map(sheet => Number(sheet?.properties?.sheetId)).filter(Number.isFinite);
    let newSheetId = Math.max(1000, ...existingIds) + 1;
    if (newSheetId > 2147483000) newSheetId = 1000000000;

    const timestamp = new Date().toISOString();
    const baselineRows = items.map(item => buildBaselineRow(item, user, timestamp, randomEventId("BASE")));
    const requests = [
      {
        addSheet: {
          properties: {
            title: CHANGE_LOG_SHEET_NAME,
            sheetId: newSheetId
          }
        }
      },
      appendCellsRequest(newSheetId, [CHANGE_LOG_HEADERS, ...baselineRows]),
      ...resolvedInventoryItems.map(item => updateBalanceCellRequest(inventorySheet.sheetId, item))
    ];

    try {
      await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    } catch (error) {
      const message = String(error?.message || "");
      if (/already exists|already used|duplicate/i.test(message)) {
        const latest = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
        const latestLog = metadataSheetInfo(latest, CHANGE_LOG_SHEET_NAME);
        if (latestLog?.sheetId) return { initialized: true, sheetId: latestLog.sheetId, concurrent: true };
      }
      throw error;
    }
    return { initialized: false, sheetId: newSheetId, created: true };
  }

  async function appendMovementBatch({
    spreadsheetId,
    accessToken,
    transactionSheetId,
    transactionRows = [],
    changes = [],
    user = "Google user",
    timestamp = new Date().toISOString()
  }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    const resolvedTransactionSheetId = transactionSheetId || metadataSheetInfo(book, CONFIG.TRANSACTIONS_SHEET_NAME)?.sheetId;
    if (!resolvedTransactionSheetId || !logSheet?.sheetId) {
      throw new Error(`The ${CHANGE_LOG_SHEET_NAME} system sheet is not ready. Refresh the page and try again.`);
    }

    const headerResponse = await api.get(
      spreadsheetId,
      `${quoteSheetName(CHANGE_LOG_SHEET_NAME)}!A1:K1`,
      accessToken
    );
    const header = headerResponse?.values?.[0] || [];
    const validHeader = CHANGE_LOG_HEADERS.every((name, index) => String(header[index] ?? "").trim() === name);
    if (!validHeader || header.length !== CHANGE_LOG_HEADERS.length) {
      throw new Error(`The "${CHANGE_LOG_SHEET_NAME}" sheet is not ready for movement recording. Refresh the page once so its header can be repaired safely.`);
    }

    const transactionEventId = randomEventId("TXN");
    const transactionRowsCopy = transactionRows.map(row => Array.isArray(row) ? row.slice() : []);
    const changeRows = changes.map(change => buildChangeRow(
      change.item,
      change.movement,
      change.quantity,
      change.client,
      user,
      timestamp,
      transactionEventId
    ));

    const requests = [
      appendCellsRequest(resolvedTransactionSheetId, transactionRowsCopy),
      appendCellsRequest(logSheet.sheetId, changeRows)
    ];

    await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    return { transactionEventId, changeCount: changeRows.length };
  }

  async function appendAsset({ spreadsheetId, accessToken, inventorySheetId, assetName, startingBalance, user = "Google user" }) {
    const inventorySchema = await ensureInventorySchema({ spreadsheetId, accessToken });

    // Adding an asset is also a schema-safety checkpoint. If Inventory or the
    // change log had just been repaired, rewrite all existing balance formulas
    // before adding the new row so no legacy column references can survive.
    await ensureChangeLog({
      spreadsheetId,
      accessToken,
      inventoryItems: inventorySchema.inventoryItems,
      user
    });

    const refreshedInventory = await ensureInventorySchema({ spreadsheetId, accessToken });
    const inventoryItems = refreshedInventory.inventoryItems;
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const resolvedInventorySheetId = inventorySheetId || refreshedInventory.sheetId || metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME)?.sheetId;
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    if (!resolvedInventorySheetId || !logSheet?.sheetId) throw new Error(`The ${CHANGE_LOG_SHEET_NAME} system sheet is not ready. Refresh the page and try again.`);

    const existingAsset = inventoryItems.some(item => item.asset.toLowerCase() === String(assetName || "").trim().toLowerCase());
    if (existingAsset) throw new Error(`Asset "${String(assetName || "").trim()}" already exists in Inventory.`);

    const header = refreshedInventory.rows?.[0] || [];
    const normalized = header.map(normalizeHeader);
    const assetIdx = normalized.findIndex(value => ["asset", "asset name", "item", "type"].includes(value));
    const balanceIdx = normalized.findIndex(value => ["balance", "current balance", "stock", "quantity"].includes(value));
    if (assetIdx < 0 || balanceIdx < 0) throw new Error(`The "${CONFIG.INVENTORY_SHEET_NAME}" sheet must contain Asset and Balance columns. No data was changed.`);

    const width = Math.max(assetIdx + 1, balanceIdx + 1);

    // appendAsset used to hard-code A:B. That silently corrupted the inventory
    // whenever a rogue/legacy column (such as Asset ID) existed before Asset.
    // Always write to the actual Asset/Balance columns.
    const rowForSheet = Array.from({ length: width }, () => "");
    rowForSheet[assetIdx] = String(assetName || "").trim();
    rowForSheet[balanceIdx] = null;

    const timestamp = new Date().toISOString();
    const eventId = randomEventId("BASE");
    const changeRow = [
      timestamp,
      eventId,
      String(assetName || "").trim(),
      "",
      Number(startingBalance) || 0,
      "BASELINE",
      "BASELINE",
      Number(startingBalance) || 0,
      "",
      user,
      ""
    ];

    await api.batchUpdate(spreadsheetId, {
      requests: [
        {
          appendCells: {
            sheetId: resolvedInventorySheetId,
            fields: "userEnteredValue",
            rows: [{
              values: rowForSheet.map((value, index) => {
                if (index === balanceIdx) return formulaCell(balanceFormula(assetIdx + 1));
                return cellValue(value);
              })
            }]
          }
        },
        appendCellsRequest(logSheet.sheetId, [changeRow])
      ]
    }, accessToken);

    return { eventId };
  }

  async function migrateRename({ spreadsheetId, accessToken, oldName, newName, inventorySheetId, rowNumber, assetColumn = 1 }) {
    const book = await request(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}?includeGridData=false`, accessToken);
    const inventorySheet = metadataSheetInfo(book, CONFIG.INVENTORY_SHEET_NAME);
    const logSheet = metadataSheetInfo(book, CHANGE_LOG_SHEET_NAME);
    const resolvedInventorySheetId = inventorySheetId || inventorySheet?.sheetId;
    if (!resolvedInventorySheetId) throw new Error(`Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`);

    const requests = [{
      updateCells: {
        start: {
          sheetId: resolvedInventorySheetId,
          rowIndex: Math.max(0, Number(rowNumber) - 1),
          columnIndex: Math.max(0, Number(assetColumn) - 1)
        },
        rows: [{ values: [cellValue(newName)] }],
        fields: "userEnteredValue"
      }
    }];

    if (logSheet?.sheetId) {
      requests.push({
        findReplace: {
          find: String(oldName || "").trim(),
          replacement: String(newName || "").trim(),
          matchCase: false,
          matchEntireCell: true,
          searchByRegex: false,
          includeFormulas: false,
          range: {
            sheetId: logSheet.sheetId,
            startColumnIndex: 2,
            endColumnIndex: 3
          }
        }
      });
    }

    await api.batchUpdate(spreadsheetId, { requests }, accessToken);
    return true;
  }

  window.FM_INVENTORY_LEDGER = Object.freeze({
    CHANGE_LOG_SHEET_NAME,
    CHANGE_LOG_HEADERS,
    INVENTORY_HEADERS,
    ensureInventorySchema,
    ensureChangeLog,
    appendMovementBatch,
    appendAsset,
    migrateRename,
    balanceFormula
  });
})();
