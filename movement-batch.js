(() => {
  "use strict";

  const MOVEMENTS = [
    ["", "Select movement"],
    ["RECEIVED", "Received"],
    ["SENT", "Sent"],
    ["DISCARD", "Discard"]
  ];

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, char => ({
      "&":"&amp;", "<":"&lt;", ">":"&gt;", "'":"&#39;", '"':"&quot;"
    }[char]));
  }

  function movementClass(movement) {
    const value = String(movement || "").toLowerCase();
    return value === "received" ? "mode-received" : value === "sent" ? "mode-sent" : value === "discard" ? "mode-discard" : "";
  }

  function movementLabel(movement) {
    const value = String(movement || "").toUpperCase();
    return value === "RECEIVED" ? "Received" : value === "SENT" ? "Sent" : value === "DISCARD" ? "Discarded" : value;
  }

  function create(options = {}) {
    const container = document.getElementById(options.containerId || "client-movement-entries");
    const addButton = document.getElementById(options.addButtonId || "add-client-entry");
    if (!container) return null;

    const getClients = typeof options.getClients === "function" ? options.getClients : () => [];
    const getAssets = typeof options.getAssets === "function" ? options.getAssets : () => [];
    const selfClient = options.selfClient || "HSC London (Self)";
    const onChange = typeof options.onChange === "function" ? options.onChange : () => {};
    const onRemoveEntry = typeof options.onRemoveEntry === "function" ? options.onRemoveEntry : () => {};
    let nextId = 1;

    function movementOptions(selected = "") {
      return MOVEMENTS.map(([value, label]) => `<option value="${value}" ${value === selected ? "selected" : ""}>${label}</option>`).join("");
    }

    function clientOptions(movement = "") {
      const clients = [...new Set((getClients() || []).map(v => String(v || "").trim()).filter(Boolean))]
        .filter(client => client !== selfClient && client !== "HSC London(Self)")
        .sort((a,b) => a.localeCompare(b));
      let html = `<option value="">Select client</option>${clients.map(client => `<option value="${escapeHtml(client)}">${escapeHtml(client)}</option>`).join("")}`;
      if (movement === "DISCARD") html = `<option value="${escapeHtml(selfClient)}">${escapeHtml(selfClient)}</option>` + html;
      return html;
    }

    function assetOptions(selected = "") {
      const assets = [...new Set((getAssets() || []).map(v => String(v || "").trim()).filter(Boolean))];
      return `<option value="">Select asset type</option>${assets.map(asset => `<option value="${escapeHtml(asset)}" ${asset === selected ? "selected" : ""}>${escapeHtml(asset)}</option>`).join("")}`;
    }

    function addAssetRow(card, item = {}, focus = false) {
      const wrap = card.querySelector(".batch-asset-rows");
      if (!wrap) return;
      const row = document.createElement("div");
      row.className = "batch-asset-row";
      row.innerHTML = `
        <div class="batch-asset-number" aria-hidden="true"></div>
        <label><span>Asset type</span><select class="batch-asset-select" required>${assetOptions(item.asset || "")}</select></label>
        <label><span>Quantity</span><input class="batch-asset-quantity" type="number" min="1" step="1" inputmode="numeric" placeholder="0" value="${item.quantity ?? ""}" required></label>
        <button type="button" class="secondary batch-remove-asset">Remove</button>`;
      wrap.appendChild(row);
      updateAssetNumbers(card);
      if (focus) row.querySelector(".batch-asset-select")?.focus();
    }

    function updateAssetNumbers(card) {
      const rows = [...card.querySelectorAll(".batch-asset-row")];
      rows.forEach((row, index) => {
        row.querySelector(".batch-asset-number").textContent = String(index + 1).padStart(2, "0");
        const button = row.querySelector(".batch-remove-asset");
        button.textContent = rows.length === 1 ? "Clear" : "Remove";
      });
    }

    function updateEntryNumbers() {
      const cards = [...container.querySelectorAll(".movement-entry-card")];
      cards.forEach((card, index) => {
        card.dataset.entryIndex = String(index);
        const number = card.querySelector(".movement-entry-number");
        if (number) number.textContent = `CLIENT ${String(index + 1).padStart(2, "0")}`;
        const remove = card.querySelector(".batch-remove-client");
        if (remove) remove.classList.toggle("hidden", cards.length === 1);
      });
      container.classList.toggle("is-multi", cards.length > 1);
      onRemoveEntry(cards.length);
    }

    function syncCard(card) {
      const movement = card.querySelector(".batch-movement")?.value || "";
      const clientSelect = card.querySelector(".batch-client");
      if (!clientSelect) return;
      const selectedClient = clientSelect.value || "";
      clientSelect.innerHTML = clientOptions(movement);
      clientSelect.value = selectedClient;

      if (movement === "DISCARD") {
        clientSelect.value = selfClient;
        clientSelect.disabled = true;
        const help = card.querySelector(".batch-client-help");
        if (help) help.textContent = `Discarded assets are recorded against ${selfClient}.`;
      } else {
        clientSelect.disabled = false;
        if (selectedClient === selfClient) clientSelect.value = "";
        const help = card.querySelector(".batch-client-help");
        if (help) help.textContent = "";
      }

      const mode = movementClass(movement);
      card.classList.remove("mode-received", "mode-sent", "mode-discard");
      if (mode) card.classList.add(mode);
      const strip = card.querySelector(".batch-mode-strip");
      if (strip) {
        const copy = {
          RECEIVED:["RECEIVED", "Assets coming into warehouse"],
          SENT:["SENT", "Assets leaving warehouse"],
          DISCARD:["DISCARDED", "Assets removed from inventory"]
        }[movement];
        strip.textContent = copy ? `${copy[0]} · ${copy[1]}` : "Select a movement to set this entry.";
        strip.className = `movement-mode-strip batch-mode-strip ${mode}`.trim();
      }

      const rows = card.querySelectorAll(".batch-asset-select");
      rows.forEach(select => {
        const value = select.value;
        select.innerHTML = assetOptions(value);
        select.value = value;
      });
    }

    function addEntry(data = {}, focus = true) {
      const card = document.createElement("article");
      card.className = "movement-entry-card";
      card.dataset.entryId = String(nextId++);
      card.innerHTML = `
        <div class="movement-entry-head">
          <div>
            <div class="eyebrow movement-entry-number">CLIENT 01</div>
            <h3>Client movement</h3>
          </div>
          <button type="button" class="secondary batch-remove-client hidden">Remove client</button>
        </div>
        <div class="batch-entry-fields">
          <label>Movement
            <select class="batch-movement" required>${movementOptions(data.movement || "")}</select>
          </label>
          <label>Client
            <select class="batch-client" required>${clientOptions(data.movement || "")}</select>
            <span class="batch-client-help field-help"></span>
          </label>
        </div>
        <div class="movement-mode-strip batch-mode-strip">Select a movement to set this entry.</div>
        <div class="batch-assets-section">
          <div class="batch-assets-head">
            <div>
              <div class="eyebrow">ASSETS</div>
              <h4>Asset types and quantities</h4>
            </div>
            <button type="button" class="secondary batch-add-asset">+ Add asset</button>
          </div>
          <div class="batch-asset-rows"></div>
        </div>`;
      container.appendChild(card);
      const initialClient = card.querySelector(".batch-client");
      if (initialClient) initialClient.value = data.client || "";
      (data.items?.length ? data.items : [{}]).forEach((item) => addAssetRow(card, item, false));
      syncCard(card);
      updateEntryNumbers();
      if (focus) card.querySelector(".batch-movement")?.focus();
      onChange();
      return card;
    }

    function removeEntry(card) {
      const cards = [...container.querySelectorAll(".movement-entry-card")];
      if (cards.length === 1) {
        card.querySelector(".batch-movement").value = "";
        card.querySelector(".batch-client").value = "";
        card.querySelector(".batch-asset-rows").innerHTML = "";
        addAssetRow(card, {}, false);
        syncCard(card);
      } else {
        card.remove();
      }
      updateEntryNumbers();
      onChange();
    }

    container.addEventListener("click", event => {
      const addAsset = event.target.closest(".batch-add-asset");
      if (addAsset) {
        const card = addAsset.closest(".movement-entry-card");
        if (card) addAssetRow(card, {}, true);
        onChange();
        return;
      }
      const removeAsset = event.target.closest(".batch-remove-asset");
      if (removeAsset) {
        const card = removeAsset.closest(".movement-entry-card");
        const row = removeAsset.closest(".batch-asset-row");
        if (!card || !row) return;
        const rows = [...card.querySelectorAll(".batch-asset-row")];
        if (rows.length === 1) {
          row.querySelector(".batch-asset-select").value = "";
          row.querySelector(".batch-asset-quantity").value = "";
        } else row.remove();
        updateAssetNumbers(card);
        onChange();
        return;
      }
      const removeClient = event.target.closest(".batch-remove-client");
      if (removeClient) {
        const card = removeClient.closest(".movement-entry-card");
        if (card) removeEntry(card);
      }
    });

    container.addEventListener("change", event => {
      const card = event.target.closest(".movement-entry-card");
      if (!card) return;
      if (event.target.matches(".batch-movement,.batch-client,.batch-asset-select")) syncCard(card);
      onChange();
    });
    container.addEventListener("input", event => {
      if (event.target.closest(".movement-entry-card")) onChange();
    });
    addButton?.addEventListener("click", () => addEntry({}, true));

    function refresh() {
      container.querySelectorAll(".movement-entry-card").forEach(syncCard);
      if (!container.querySelector(".movement-entry-card")) addEntry({}, false);
      updateEntryNumbers();
    }

    function read() {
      return [...container.querySelectorAll(".movement-entry-card")].map(card => ({
        movement: card.querySelector(".batch-movement")?.value || "",
        client: card.querySelector(".batch-client")?.value || "",
        items: [...card.querySelectorAll(".batch-asset-row")].map(row => ({
          asset: row.querySelector(".batch-asset-select")?.value || "",
          quantity: Number(row.querySelector(".batch-asset-quantity")?.value || 0)
        }))
      }));
    }

    function updatePreview(targetId = "preview-text") {
      const target = document.getElementById(targetId);
      if (!target) return;
      const entries = read().filter(entry => entry.movement || entry.client || entry.items.some(item => item.asset || item.quantity));
      if (!entries.length) {
        target.textContent = "Add one or more client movements. Each client can contain multiple asset types.";
        return;
      }
      const lines = entries.map(entry => {
        const client = entry.movement === "DISCARD" ? selfClient : entry.client;
        const items = entry.items.filter(item => item.asset && Number.isInteger(item.quantity) && item.quantity > 0);
        if (!client || !items.length) return `${movementLabel(entry.movement) || "Movement"} · Select client and assets`;
        return `${movementLabel(entry.movement)} · ${client} · ${items.map(item => `${item.quantity.toLocaleString("en-GB")} × ${item.asset}`).join(" · ")}`;
      });
      target.innerHTML = lines.map(line => `<span class="preview-batch-line">${escapeHtml(line)}</span>`).join("");
    }

    function validate() {
      const entries = read();
      if (!entries.length) return "Add at least one client movement.";
      if (entries.some(entry => !["RECEIVED","SENT","DISCARD"].includes(entry.movement))) return "Select a movement for every client entry.";
      for (let index = 0; index < entries.length; index += 1) {
        const entry = entries[index];
        const client = entry.movement === "DISCARD" ? selfClient : entry.client;
        if (!client) return `Select a client for Client ${index + 1}.`;
        if (!entry.items.length || entry.items.some(item => !item.asset || !Number.isInteger(item.quantity) || item.quantity <= 0)) {
          return `Complete the asset type and quantity rows for Client ${index + 1}.`;
        }
        const seen = new Set();
        for (const item of entry.items) {
          const key = String(item?.asset ?? "").trim().toLowerCase();
          if (seen.has(key)) return `Client ${index + 1} has ${item.asset} more than once. Combine the quantities.`;
          seen.add(key);
        }
      }
      return null;
    }

    function setAllDisabled(disabled) {
      container.querySelectorAll("select,input,button").forEach(control => control.disabled = Boolean(disabled));
      addButton && (addButton.disabled = Boolean(disabled));
      container.querySelectorAll(".batch-client").forEach(select => {
        if (!disabled) {
          const movement = select.closest(".movement-entry-card")?.querySelector(".batch-movement")?.value;
          select.disabled = movement === "DISCARD";
        }
      });
    }

    function reset() {
      container.innerHTML = "";
      nextId = 1;
      addEntry({}, false);
      onChange();
    }

    addEntry(options.initialEntry || {}, false);
    refresh();

    return { addEntry, read, refresh, reset, validate, updatePreview, setAllDisabled, container };
  }

  window.FM_BATCH_FORM = { create, movementLabel, movementClass };
})();
