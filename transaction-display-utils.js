(() => {
  "use strict";

  const ORDER = ["RECEIVED", "SENT", "DISCARD", "COMBINE_OUT", "COMBINE_IN", "NORMALISATION"];

  function canonicalMovement(value) {
    const movement = String(value || "").trim().toUpperCase();
    if (["RECEIVED","INBOUND","IN","RECEIVE"].includes(movement)) return "RECEIVED";
    if (["SENT","OUTBOUND","OUT","SEND"].includes(movement)) return "SENT";
    if (["DISCARDED","DISCARD"].includes(movement)) return "DISCARD";
    if (["COMBINE_OUT", "COMBINATION OUT", "COMPONENT USED"].includes(movement)) return "COMBINE_OUT";
    if (["COMBINE_IN", "COMBINATION IN", "ASSET CREATED"].includes(movement)) return "COMBINE_IN";
    if (["NORMALISATION", "NORMALIZATION", "LEDGER ADJUSTMENT"].includes(movement)) return "NORMALISATION";
    return movement || "OTHER";
  }

  function movementLabel(value) {
    const movement = canonicalMovement(value);
    if (movement === "RECEIVED") return "Received";
    if (movement === "SENT") return "Sent";
    if (movement === "DISCARD") return "Discarded";
    if (movement === "COMBINE_OUT") return "Combination component used";
    if (movement === "COMBINE_IN") return "Combined asset created";
    if (movement === "NORMALISATION") return "Client ledger normalisation";
    return String(value || "Other");
  }

  function movementClass(value) {
    const movement = canonicalMovement(value);
    return movement === "RECEIVED" ? "movement-received" : movement === "SENT" ? "movement-sent" : movement === "DISCARD" ? "movement-discard" : movement === "COMBINE_IN" ? "movement-received" : movement === "COMBINE_OUT" ? "movement-discard" : movement === "NORMALISATION" ? "movement-other" : "movement-other";
  }

  function timestampMs(value) {
    const direct = new Date(value);
    if (!Number.isNaN(direct.getTime())) return direct.getTime();
    const text = String(value || "").trim();
    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const serial = Number(text);
      if (serial > 20000 && serial < 80000) return Math.round((serial - 25569) * 86400000);
    }
    const dmy = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/);
    if (dmy) {
      const ms = Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), Number(dmy[4] || 0), Number(dmy[5] || 0), Number(dmy[6] || 0), Number(String(dmy[7] || "0").slice(0,3).padEnd(3,"0")));
      return Number.isFinite(ms) ? ms : NaN;
    }
    return NaN;
  }

  function timeKey(value) {
    const ms = timestampMs(value);
    return Number.isFinite(ms) ? `date:${ms}` : `raw:${String(value || "").trim()}`;
  }

  // Registration comments are normally entered first, sometimes with a short
  // label such as "Vehicle Reg:". Match complete, familiar UK and European
  // plate formats from the beginning of the comment to avoid guessing from
  // unrelated words later in a note.
  function extractVehicleRegistration(comment) {
    const original = String(comment || "").normalize("NFKC").trim().toUpperCase();
    if (!original) return "";

    // Most comments start with a registration, but also support labelled
    // registrations later in a note (for example: "Dock 4 · Vehicle Reg: AB12 CDE").
    const candidates = [];
    let leading = original.replace(/^(?:(?:VEHICLE\s*)?(?:REG(?:ISTRATION)?|VRM|PLATE|NUMBER\s*PLATE)\s*[:#=-]?\s*)+/i, "");
    leading = leading.replace(/^[\s:;,|()[\]{}]+/, "").trim();
    if (leading) candidates.push(leading);
    const labelledRegistration = /\b(?:VEHICLE\s*(?:REG(?:ISTRATION)?|REGISTRATION)|REG(?:ISTRATION)?|VRM|NUMBER\s*PLATE|PLATE)\s*[:#=-]?\s*/gi;
    for (const match of original.matchAll(labelledRegistration)) {
      const fragment = original.slice(match.index + match[0].length).replace(/^[\s:;,|()[\]{}]+/, "").trim();
      if (fragment && !candidates.includes(fragment)) candidates.push(fragment);
    }
    if (!candidates.length) return "";

    const patterns = [
      // Great Britain current plates, plus common older and dateless styles.
      /^[A-Z]{2}\s?\d{2}\s?[A-Z]{3}(?=\b|$)/,
      /^[A-Z]\s?\d{1,3}\s?[A-Z]{3}(?=\b|$)/,
      /^[A-Z]{3}\s?\d{1,3}\s?[A-Z](?=\b|$)/,
      /^(?:[A-Z]{3}\s?\d{3,4}|\d{1,4}\s?[A-Z]{3})(?=\b|$)/,
      // Ireland.
      /^\d{2,3}[- ]?[A-Z]{1,2}[- ]?\d{1,6}(?=\b|$)/,
      // France, Italy and Spain.
      /^[A-Z]{2}[- ]?\d{3}[- ]?[A-Z]{2}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{3}\s?[A-Z]{2}(?=\b|$)/,
      /^\d{4}[- ]?[BCDFGHJKLMNPRSTVWXYZ]{3}(?=\b|$)/,
      // Portugal and common Dutch series.
      /^\d{2}[- ]?[A-Z]{2}[- ]?\d{2}(?=\b|$)/,
      /^[A-Z]{2}[- ]?\d{2}[- ]?[A-Z]{2,3}(?=\b|$)/,
      /^[A-Z]{2}[- ]?\d{3}[- ]?[A-Z](?=\b|$)/,
      /^\d[- ]?[A-Z]{3}[- ]?\d{2}(?=\b|$)/,
      /^\d{2}[- ]?[A-Z]{3}[- ]?\d(?=\b|$)/,
      /^[A-Z]{3}[- ]?\d{2}[- ]?[A-Z](?=\b|$)/,
      // Germany, Austria and several central/eastern European formats.
      /^[A-ZÄÖÜ]{1,3}[- ]+[A-Z]{1,2}[- ]?\d{1,4}(?=\b|$)/,
      /^[A-Z]{2,3}\s?\d[A-Z0-9]{4,5}(?=\b|$)/,
      /^[A-Z]{3}[- ]?\d{3}(?=\b|$)/,
      /^\d[A-Z]\d\s?\d{4}(?=\b|$)/,
      /^[A-Z]{4}[- ]?\d{3}(?=\b|$)/,
      /^[A-Z]{3}[- ]?\d{4}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{2}\s?[A-Z]{3}(?=\b|$)/,
      /^[A-Z]{1,2}\s?\d{4}\s?[A-Z]{1,2}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{3,4}\s?[A-Z]{1,2}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{3}[A-Z]{2}(?=\b|$)/,
      /^\d[- ]?[A-Z]{3}[- ]?\d{3}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{5}(?=\b|$)/,
      // Additional EU country series: Latvia, Estonia, Romania, Slovenia,
      // Austria and Luxembourg. These are checked before broad legacy UK
      // formats so their trailing letters/numbers are not lost.
      /^[A-Z]{2}[- ]?\d{4}(?=\b|$)/,
      /^\d{3}[- ]?[A-Z]{3}(?=\b|$)/,
      /^[A-Z]{1,2}[- ]?\d{1,5}[A-Z]{1,2}(?=\b|$)/,
      /^[A-Z]{2}[- ]?[A-Z]{2}[- ]?\d{3,4}(?=\b|$)/,
      /^[A-Z]{2}\s?\d{2}\s?[A-Z]{3}(?=\b|$)/,
      /^\d{1,4}[- ]?[A-Z]{1,2}(?=\b|$)/,
      // Northern Ireland / other legacy UK plates. Requiring 3-4 digits
      // avoids treating short notes such as "HSC 10 pallets" as a plate.
      /^[A-Z]{1,3}\s+\d{3,4}(?=\b|$)/
    ];

    for (const candidate of candidates) {
      for (const pattern of patterns) {
        const match = candidate.match(pattern);
        if (match) return match[0].replace(/\s+/g, " ").trim();
      }
    }
    return "";
  }

  function groupTransactions(transactions = []) {
    const groups = new Map();
    (transactions || []).forEach((item, index) => {
      const client = String(item?.client || "Unknown client").trim() || "Unknown client";
      const vehicleReg = extractVehicleRegistration(item?.comment);
      // A shared timestamp is not enough to prove two ledger lines belong to
      // the same vehicle movement. Merge only when both lines carry the same
      // registration; unlabelled/unknown registrations remain separate.
      const registrationKey = vehicleReg
        ? vehicleReg.replace(/[^A-Z0-9]/g, "")
        : `unregistered-row:${index}`;
      const key = `${String(client).toLowerCase()}\u0000${timeKey(item?.timestamp)}\u0000${registrationKey}`;
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          client,
          timestamp: item?.timestamp || "",
          items: [],
          movements: { RECEIVED: [], SENT: [], DISCARD: [], OTHER: [] },
          totalLines: 0,
          totalQuantity: 0,
          image: "",
          comment: "",
          comments: [],
          vehicleReg: vehicleReg || "",
          originalIndexes: [],
          lastIndex: index,
          sortTimestamp: timestampMs(item?.timestamp)
        });
      }
      const group = groups.get(key);
      const movement = canonicalMovement(item?.movement);
      const quantity = Number(item?.quantity) || 0;
      const comment = String(item?.comment || "").trim();
      group.items.push(item);
      group.originalIndexes.push(index);
      group.lastIndex = index;
      if (!Number.isFinite(group.sortTimestamp)) group.sortTimestamp = timestampMs(item?.timestamp);
      group.totalLines += 1;
      group.totalQuantity += quantity;
      if (!group.image && item?.image) group.image = item.image;
      if (comment && !group.comments.includes(comment)) group.comments.push(comment);
      if (!group.comment && comment) group.comment = comment;
      if (!group.vehicleReg && vehicleReg) group.vehicleReg = vehicleReg;
      if (!group.movements[movement]) group.movements[movement] = [];
      group.movements[movement].push({
        asset: String(item?.asset || "Unknown asset").trim() || "Unknown asset",
        quantity,
        raw: item
      });
    });

    return [...groups.values()].sort((a, b) => {
      const da = a.sortTimestamp;
      const db = b.sortTimestamp;
      if (Number.isFinite(db) && Number.isFinite(da) && db !== da) return db - da;
      if (Number.isFinite(db) !== Number.isFinite(da)) return Number.isFinite(db) ? -1 : 1;
      return (b.lastIndex || 0) - (a.lastIndex || 0);
    });
  }


  function movementSections(group, options = {}) {
    const includeEmpty = Boolean(options.includeEmpty);
    const sections = [];
    const addSection = movement => {
      const items = Array.isArray(group?.movements?.[movement]) ? group.movements[movement] : [];
      if (!items.length && !includeEmpty) return;
      const totalQuantity = items.reduce((sum, item) => sum + (Number(item?.quantity) || 0), 0);
      sections.push({
        movement,
        key: movement,
        label: movementLabel(movement),
        items,
        totalQuantity
      });
    };
    ORDER.forEach(addSection);
    addSection("OTHER");
    return sections;
  }

  function aggregateAssets(items = []) {
    const totals = new Map();
    items.forEach(item => {
      const asset = String(item.asset || "Unknown asset").trim() || "Unknown asset";
      totals.set(asset, (totals.get(asset) || 0) + (Number(item.quantity) || 0));
    });
    return [...totals.entries()].map(([asset, quantity]) => ({asset, quantity})).sort((a,b) => a.asset.localeCompare(b.asset));
  }

  function cardClass(group) {
    const movements = movementSections(group).map(section => section.movement);
    if (movements.length === 1) return `movement-card-${String(movements[0] ?? "other").toLowerCase()}`;
    if (movements.length > 1) return "movement-card-mixed";
    return "movement-card-other";
  }

  function movementLabels(group) {
    return movementSections(group).map(section => movementLabel(section.movement));
  }

  function assetSummary(group) {
    const assets = aggregateAssets(group?.items || []);
    return assets.map(item => item.asset).join(", ");
  }

  function formatLineCount(group) {
    const count = Number(group?.totalLines || 0);
    return `${count.toLocaleString()} ledger line${count === 1 ? "" : "s"}`;
  }

  window.FM_TRANSACTION_DISPLAY = {
    canonicalMovement,
    extractVehicleRegistration,
    movementLabel,
    movementClass,
    groupTransactions,
    movementSections,
    aggregateAssets,
    cardClass,
    movementLabels,
    assetSummary,
    formatLineCount
  };
})();
