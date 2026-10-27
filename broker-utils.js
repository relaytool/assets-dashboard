(() => {
  "use strict";
  const DEFAULT_BROKERS = Object.freeze([
    "YDH", "NGT", "APEX", "WANB", "CY-UK", "YUNEXPRESS", "GLOBALXAIR",
    "FAST TRACK EXPRESS", "TRANS GLOBAL", "EXPRESS SPEED", "ATA",
    "J&T EXPRESS", "SuperSmartServices"
  ]);
  const clean = value => String(value ?? "").trim().replace(/\s+/g, " ");
  const key = value => clean(value).toLocaleUpperCase("en-GB");
  function parseBrokers(rows) {
    const names = [];
    const seen = new Set();
    const configured = (rows || []).flatMap((row, i) => i === 0 && /^(broker|broker name)$/i.test(clean(row?.[0])) ? [] : [row?.[0]]).map(clean).filter(Boolean);
    for (const name of (configured.length ? configured : DEFAULT_BROKERS)) {
      const value = clean(name);
      if (!value || seen.has(key(value))) continue;
      seen.add(key(value));
      names.push(value);
    }
    return names;
  }
  function summarize(rows) {
    const clients = new Map();
    for (const row of rows || []) {
      const client = clean(row.client);
      const asset = clean(row.asset);
      const direction = key(row.movement);
      if (!client || !asset || !["SENT", "RECEIVED"].includes(direction)) continue;
      const amount = Math.abs(Number(row.quantity) || 0);
      if (!amount) continue;
      const clientKey = key(client), assetKey = key(asset);
      if (!clients.has(clientKey)) clients.set(clientKey, {client, assets: new Map()});
      const record = clients.get(clientKey);
      if (!record.assets.has(assetKey)) record.assets.set(assetKey, {asset, sent:0, received:0, net:0});
      const totals = record.assets.get(assetKey);
      if (direction === "SENT") totals.sent += amount;
      else totals.received += amount;
      totals.net = totals.sent - totals.received;
    }
    return [...clients.values()].map(({client,assets}) => ({
      client, assets: [...assets.values()].sort((a,b) => a.asset.localeCompare(b.asset))
    })).sort((a,b) => a.client.localeCompare(b.client));
  }
  window.FM_BROKERS = Object.freeze({DEFAULT_BROKERS, parseBrokers, clean, key, summarize});
})();
