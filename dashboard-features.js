/* Dashboard-only presentation rules. No Google Sheets rows are modified here. */
(() => {
  "use strict";

  const normalise = value => String(value ?? "").trim().replace(/\s+/g, " ").toLocaleLowerCase();

  // Client Types is column-based: each column header is a category, cells
  // underneath are its members. First matching column takes precedence.
  function parseClientTypes(rows = []) {
    const headers = Array.isArray(rows[0]) ? rows[0] : [];
    const categories = [];
    const lookup = Object.create(null);
    const names = new Set();
    headers.forEach((header, column) => {
      const name = String(header ?? "").trim();
      const key = normalise(name);
      if (!key || names.has(key)) return;
      names.add(key);
      categories.push(name);
      rows.slice(1).forEach(row => {
        const client = String(row?.[column] ?? "").trim();
        if (client && lookup[normalise(client)] === undefined) lookup[normalise(client)] = name;
      });
    });
    return { categories, lookup };
  }

  function isAdjustment(row) {
    if (row?.isLedgerAdjustment) return true;
    const movement = normalise(row?.movement).replace(/[ -]+/g, "_");
    return /^(normalis|normaliz|adjust|correct|reconcil|combine_|inventory_adjust)/.test(movement);
  }

  function estimateSheinOutbound(transactions = [], now = Date.now()) {
    const matches = (transactions || []).filter(row =>
      normalise(row?.client) === "shein domestic" &&
      normalise(row?.asset) === "red sleeves" &&
      (normalise(row?.movement) === "sent" || normalise(row?.movement) === "received")
    ).map(row => ({ ...row, time: Date.parse(row?.timestamp) }))
      .filter(row => Number.isFinite(row.time) && row.time <= now);

    const sent = matches.filter(row => normalise(row.movement) === "sent");
    if (!sent.length) return { available: false, quantity: 0, lastSent: null };
    const lastSent = Math.max(...sent.map(row => row.time));
    const quantity = matches
      .filter(row => normalise(row.movement) === "received" && row.time > lastSent)
      .reduce((total, row) => {
        const number = Number(String(row.quantity ?? 0).replace(/,/g, ""));
        return total + (Number.isFinite(number) && number > 0 ? number : 0);
      }, 0);
    return { available: true, quantity, lastSent: new Date(lastSent).toISOString() };
  }

  function applySheinEstimate(outbound = [], estimate, enabled) {
    if (!enabled || !estimate?.available) return outbound.slice();
    let found = false;
    const items = outbound.map(item => {
      if (normalise(item.client) !== "shein domestic" || normalise(item.asset) !== "red sleeves") return item;
      found = true;
      return { ...item, quantity: estimate.quantity, experimentalEstimate: true, lastSent: estimate.lastSent };
    });
    if (!found) items.push({
      direction: "Outbound", client: "SHEIN Domestic", asset: "Red Sleeves",
      quantity: estimate.quantity, plannedArrival: "Estimated", destination: "",
      experimentalEstimate: true, lastSent: estimate.lastSent
    });
    return items;
  }

  window.FM_DASHBOARD_FEATURES = Object.freeze({
    parseClientTypes, isAdjustment, estimateSheinOutbound, applySheinEstimate
  });
})();
