(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const state = { accessToken: null, idTokenPayload: null, routines: [], clients: [], inventory: [], editingId: "", editingActive: true };
  let tokenRequestPromise = null;
  let routineSaveBusy = false;
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
    $("routine-form")?.addEventListener("submit", saveRoutine);
    $("frequency")?.addEventListener("change", toggleWeekday);
    $("cancel-edit")?.addEventListener("click", resetForm);
    $("routine-body")?.addEventListener("click", event => {
      const edit = event.target.closest("[data-edit-routine]");
      if (edit) beginEdit(edit.dataset.editRoutine);
      const toggle = event.target.closest("[data-toggle-routine]");
      if (toggle) toggleActive(toggle.dataset.toggleRoutine);
    });
  }

  function bindQuickMenu() {
    const nav = document.querySelector(".quick-nav");
    const toggle = $("menu-toggle");
    const links = $("quick-links");
    if (!nav || !toggle) return;
    const key = "fmQuickMenuOpen";
    const mobileQuery = window.matchMedia("(max-width: 760px)");
    const isMobile = () => mobileQuery.matches;
    const setOpen = open => {
      nav.classList.toggle("menu-open", Boolean(open));
      toggle.setAttribute("aria-expanded", String(Boolean(open)));
      if (!isMobile()) localStorage.setItem(key, open ? "1" : "0");
    };
    const saved = localStorage.getItem(key);
    setOpen(isMobile() ? false : saved === "1");
    toggle.addEventListener("click", () => setOpen(!nav.classList.contains("menu-open")));
    links?.addEventListener("click", event => { if (event.target.closest("a") && isMobile()) setOpen(false); });
    const onViewportChange = event => { if (event.matches) setOpen(false); };
    if (mobileQuery.addEventListener) mobileQuery.addEventListener("change", onViewportChange); else mobileQuery.addListener(onViewportChange);
  }

  function waitForGoogle() {
    if (window.google?.accounts?.id && window.google?.accounts?.oauth2) return initializeGoogle();
    let done = false;
    const started = Date.now();
    const finish = () => { if (done) return; done = true; clearInterval(timer); initializeGoogle(); };
    window.addEventListener("fm-google-loaded", finish, { once: true });
    const timer = setInterval(() => {
      if (window.google?.accounts?.id && window.google?.accounts?.oauth2) { finish(); }
      else if (Date.now() - started >= 8000) { clearInterval(timer); setStatus("Google services could not be loaded. Check your internet connection.", true); }
    }, 50);
  }

  function initializeGoogle() {
    google.accounts.id.initialize({ client_id: CONFIG.GOOGLE_CLIENT_ID, callback: handleCredentialResponse, auto_select: true, cancel_on_tap_outside: false });
    const saved = readSession();
    if (saved) {
      setUserProfile(saved);
      $("google-signin-button")?.classList.add("hidden");
      $("grant-access").classList.add("hidden");
      const cached = window.FM_AUTH_CACHE?.read?.(saved.email);
      if (cached?.token) {
        state.accessToken = cached.token;
        window.FM_CONNECTION_UI?.show("Restoring your session…", "Using your saved Google Sheets connection.");
        hideLogin();
        loadPage().finally(() => window.FM_CONNECTION_UI?.hide());
      } else {
        setStatus("Restoring your Sheets connection…");
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Restoring your saved Google session.");
        attemptSilentAccess(saved.email);
      }
    } else {
      window.FM_CONNECTION_UI?.hide();
      google.accounts.id.renderButton($("google-signin-button"), { theme: "outline", size: "large", text: "signin_with", shape: "rectangular", width: 280 });
      google.accounts.id.prompt();
    }
  }

  function handleCredentialResponse(response) {
    try { state.idTokenPayload = decodeJwtPayload(response.credential); saveSession(); setUserProfile(state.idTokenPayload); requestAccess(true); }
    catch (e) { console.error(e); setStatus("Google sign-in response could not be read.", true); }
  }

  function requestAccess(silent = true) {
    window.FM_CONNECTION_UI?.show(silent ? "Connecting to Google Sheets…" : "Authorising Google Sheets & Drive…", silent ? "Restoring your saved connection." : "Approve access in the Google prompt to continue.");
    acquireAccessToken(silent ? "none" : "consent").then(async () => { hideLogin(); await loadPage(); window.FM_CONNECTION_UI?.hide(); }).catch(error => { window.FM_CONNECTION_UI?.hide(); $("grant-access").classList.remove("hidden"); setStatus(error.message || "Google authorization failed.", true); });
  }
  function attemptSilentAccess(email) {
    window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Checking your current Google authorisation.");
    acquireAccessToken("none", email).then(async () => { hideLogin(); await loadPage(); window.FM_CONNECTION_UI?.hide(); }).catch(() => { window.FM_CONNECTION_UI?.hide(); $("grant-access").classList.remove("hidden"); setStatus("Your Google session is available. Allow Sheets & Drive access to continue."); });
  }

  function acquireAccessToken(prompt = "none", email, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const expectedEmail = email || state.idTokenPayload?.email || readSession()?.email || "";
    if (!forceRefresh) { const cached = window.FM_AUTH_CACHE?.read?.(expectedEmail); if (cached?.token) { state.accessToken = cached.token; return Promise.resolve(cached.token); } }
    if (tokenRequestPromise) return tokenRequestPromise;
    tokenRequestPromise = new Promise((resolve, reject) => {
      let settled = false;
      const finish = (fn, value) => { if (settled) return; settled = true; tokenRequestPromise = null; clearTimeout(timeoutId); fn(value); };
      const timeoutId = setTimeout(() => finish(reject, new Error("Google authorisation did not complete.")), prompt === "none" ? 5000 : 15000);
      const client = google.accounts.oauth2.initTokenClient({ client_id: CONFIG.GOOGLE_CLIENT_ID, scope: CONFIG.OAUTH_SCOPES, callback: response => {
        if (response.error) { finish(reject, new Error(`Google authorization failed: ${response.error}`)); return; }
        state.accessToken = response.access_token; window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, expectedEmail); finish(resolve, response.access_token);
      }});
      client.requestAccessToken({ prompt, login_hint: expectedEmail || undefined });
    });
    return tokenRequestPromise;
  }

  async function loadPage() {
    if (!state.accessToken) return;
    setSyncStatus("Syncing routine schedules...");
    try {
      const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      let meta = await sheetsGet(`/${encodeURIComponent(id)}`);
      let titles = (meta.sheets || []).map(s => s.properties.title);
      const missing = [CONFIG.ROUTINE_SHEET_NAME, CONFIG.ALERTS_SHEET_NAME].filter(name => !titles.includes(name));
      if (missing.length) { await createSheets(missing); titles = titles.concat(missing); }
      const [routineRows, clientRows, inventoryRows] = await Promise.all([
        getValues(id, CONFIG.ROUTINE_SHEET_NAME), getValues(id, CONFIG.CLIENT_LIST_SHEET_NAME), getValues(id, CONFIG.INVENTORY_SHEET_NAME)
      ]);
      state.routines = FM_ROUTINE_ALERTS.parseRoutineRows(routineRows);
      state.clients = parseClients(clientRows);
      state.inventory = parseInventory(inventoryRows);
      $("workbook-name").textContent = CONFIG.INVENTORY_LEDGER_NAME || "Assets Inventory Ledger";
      populateFormOptions();
      renderRoutines();
      resetForm();
      setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
    } catch (e) { console.error(e); setSyncStatus(e.message || "Unable to load routine schedules.", true); }
  }

  async function createSheets(names) {
    const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
    await sheetsPost(`/${encodeURIComponent(id)}:batchUpdate`, { requests: names.map(title => ({ addSheet: { properties: { title } } })) });
    if (names.includes(CONFIG.ROUTINE_SHEET_NAME)) await updateValues(id, CONFIG.ROUTINE_SHEET_NAME, [FM_ROUTINE_ALERTS.ROUTINE_HEADERS]);
    if (names.includes(CONFIG.ALERTS_SHEET_NAME)) await updateValues(id, CONFIG.ALERTS_SHEET_NAME, [FM_ROUTINE_ALERTS.ALERT_HEADERS]);
  }

  function populateFormOptions() {
    const client = $("client");
    const asset = $("asset");
    const currentClient = client.value, currentAsset = asset.value;
    client.innerHTML = `<option value="">Select client</option>${state.clients.filter(c => c !== "HSC London(Self)" && c !== CONFIG.SELF_CLIENT_NAME).map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`;
    asset.innerHTML = `<option value="">Select asset</option>${state.inventory.map(i => `<option value="${escapeAttr(i.asset)}">${escapeHtml(i.asset)}</option>`).join("")}`;
    if ([...client.options].some(o => o.value === currentClient)) client.value = currentClient;
    if ([...asset.options].some(o => o.value === currentAsset)) asset.value = currentAsset;
  }

  function saveRoutine(event) {
    event.preventDefault();
    if (routineSaveBusy) return;
    const data = {
      id: state.editingId || `RO-${Date.now()}-${Math.random().toString(36).slice(2,7).toUpperCase()}`,
      active: state.editingId ? state.editingActive : true,
      frequency: $("frequency").value,
      weekday: $("frequency").value === "Weekly" ? $("weekday").value : "",
      direction: $("direction").value,
      client: $("client").value.trim(),
      asset: $("asset").value.trim(),
      quantity: Number($("quantity").value),
      destination: $("destination").value.trim(),
      plannedTime: $("planned-time").value,
      notes: $("notes").value.trim(),
      createdBy: state.idTokenPayload?.email || readSession()?.email || "Google user",
      createdAt: state.editingId ? (state.routines.find(r => r.id === state.editingId)?.createdAt || "") : new Date().toISOString()
    };
    if (!data.frequency || (data.frequency === "Weekly" && !data.weekday) || !data.direction || !data.client || !data.asset || !Number.isInteger(data.quantity) || data.quantity <= 0) {
      setFormStatus("Choose a frequency, direction, client, asset and a whole quantity greater than zero.", true); return;
    }
    const existing = state.routines.filter(r => r.id !== data.id);
    existing.push(data);
    routineSaveBusy = true;
    $("save-routine").disabled = true;
    persistRoutines(existing).then(() => { state.routines = existing; setFormStatus(state.editingId ? "Routine updated." : "Routine added."); resetForm(); renderRoutines(); }).catch(e => setFormStatus(e.message || "Unable to save routine.", true)).finally(() => { routineSaveBusy = false; $("save-routine").disabled = false; });
  }

  async function persistRoutines(routines) {
    const id = CONFIG.INVENTORY_LEDGER_SHEET_ID;
    const rows = [FM_ROUTINE_ALERTS.ROUTINE_HEADERS, ...routines.map(r => [r.id, r.active ? "TRUE" : "FALSE", r.frequency, r.weekday, r.direction, r.client, r.asset, r.quantity, r.destination, r.plannedTime, r.notes, r.createdBy, r.createdAt])];
    const range = `${quoteSheetName(CONFIG.ROUTINE_SHEET_NAME)}!A:M`;
    await sheetsPut(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, { range, majorDimension: "ROWS", values: rows });
  }

  function beginEdit(id) {
    const r = state.routines.find(item => item.id === id); if (!r) return;
    state.editingId = id;
    state.editingActive = r.active;
    $("frequency").value = r.frequency; toggleWeekday(); $("weekday").value = r.weekday || "";
    $("direction").value = r.direction; $("client").value = r.client; $("asset").value = r.asset; $("quantity").value = r.quantity; $("destination").value = r.destination; $("planned-time").value = r.plannedTime; $("notes").value = r.notes;
    $("save-routine").textContent = "Update routine"; $("cancel-edit").classList.remove("hidden"); $("routine-form").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function toggleActive(id) {
    const routines = state.routines.map(r => r.id === id ? { ...r, active: !r.active } : r);
    try { await persistRoutines(routines); state.routines = routines; renderRoutines(); setFormStatus("Routine status updated."); }
    catch (e) { setFormStatus(e.message || "Unable to update routine.", true); }
  }

  function resetForm() {
    state.editingId = ""; state.editingActive = true; $("routine-form")?.reset(); $("save-routine").textContent = "Add routine"; $("cancel-edit").classList.add("hidden"); toggleWeekday();
  }

  function toggleWeekday() { $("weekday-wrap")?.classList.toggle("hidden", $("frequency")?.value !== "Weekly"); }

  function renderRoutines() {
    const rows = [...state.routines].sort((a,b) => Number(b.active)-Number(a.active) || a.client.localeCompare(b.client) || a.asset.localeCompare(b.asset));
    $("routine-count").textContent = rows.length.toLocaleString();
    $("routine-body").innerHTML = rows.length ? rows.map(r => `<tr class="${r.active ? "" : "routine-inactive"}">
      <td><strong>${escapeHtml(r.client)}</strong><span class="subtext">${escapeHtml(r.direction)} · ${escapeHtml(r.destination || "No destination")}</span></td>
      <td>${escapeHtml(FM_ROUTINE_ALERTS.routineLabel(r))}${r.plannedTime ? `<span class="subtext">${escapeHtml(r.plannedTime)}</span>` : ""}</td>
      <td><strong>${escapeHtml(r.asset)}</strong></td><td class="num">${FM_ROUTINE_ALERTS.formatQty(r.quantity)}</td>
      <td><span class="status-chip ${r.active ? "status-active" : "status-off"}">${r.active ? "Active" : "Off"}</span></td>
      <td><button class="secondary small-button" type="button" data-edit-routine="${escapeAttr(r.id)}">Edit</button> <button class="secondary small-button" type="button" data-toggle-routine="${escapeAttr(r.id)}">${r.active ? "Deactivate" : "Activate"}</button></td>
    </tr>`).join("") : `<tr><td colspan="6" class="empty">No routine movements added yet.</td></tr>`;
  }

  function parseClients(rows) { if (!rows.length) return []; const h=rows[0].map(v=>String(v??"").trim().toLowerCase()); const idx=h.findIndex(v=>["client","client name","name"].includes(v)); return idx<0?[]:[...new Set(rows.slice(1).map(r=>String(r[idx]??"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b)); }
  function parseInventory(rows) { if (!rows.length) return []; const h=rows[0].map(v=>String(v??"").trim().toLowerCase()); const ai=h.findIndex(v=>["asset","asset name","item","type"].includes(v)); const bi=h.findIndex(v=>["balance","current balance","stock","quantity"].includes(v)); return rows.slice(1).map((r,i)=>({rowNumber:i+2,asset:String(r[ai]??"").trim(),balance:Number(String(r[bi]??"0").replace(/,/g,""))||0})).filter(x=>x.asset); }

  async function getValues(id, sheet) { const data = await sheetsGet(`/${encodeURIComponent(id)}/values/${encodeURIComponent(`${quoteSheetName(sheet)}!A:O`)}`); return data.values || []; }
  async function updateValues(id, sheet, rows) { const range=`${quoteSheetName(sheet)}!A1`; return sheetsPut(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {range,majorDimension:"ROWS",values:rows}); }
  async function sheetsGet(path) { return fetchJson(SHEETS_API + path); }
  async function sheetsPost(path, body) { return fetchJson(SHEETS_API + path, {method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)}); }
  async function sheetsPut(path, body) { return fetchJson(SHEETS_API + path, {method:"PUT",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)}); }
  async function fetchJson(url, options={}) {
    const response=await fetch(url,{...options,headers:{...authHeaders(),...(options.headers||{})}});
    if(response.status===401&&!options.__retried){
      try{
        await acquireAccessToken("none",state.idTokenPayload?.email||readSession()?.email,{forceRefresh:true});
        return fetchJson(url,{...options,__retried:true,headers:{...(options.headers||{}),...authHeaders()}});
      }catch(_){
        state.accessToken=null;window.FM_AUTH_CACHE?.clear?.();$("grant-access")?.classList.remove("hidden");
        throw new Error("Google Sheets access expired. Allow Sheets & Drive access to reconnect.");
      }
    }
    const text=await response.text(); let data={}; try{data=text?JSON.parse(text):{}}catch{}
    if(!response.ok) throw new Error(data.error?.message||`Request failed (${response.status})`);
    return data;
  }
  function authHeaders(){return {Authorization:`Bearer ${state.accessToken}`};}
  function quoteSheetName(name){return `'${String(name).replace(/'/g,"''")}'`;}
  function decodeJwtPayload(token){const part=token.split(".")[1]; const normalized=part.replace(/-/g,"+").replace(/_/g,"/")+"=".repeat((4-part.length%4)%4); return JSON.parse(decodeURIComponent(Array.from(atob(normalized)).map(c=>`%${c.charCodeAt(0).toString(16).padStart(2,"0")}`).join("")));}
  function setUserProfile(p){const rawName=String(p?.given_name||p?.name||"Google user").trim();const firstName=(rawName.split(/\s+/)[0]||"there").replace(/[^\p{L}\p{M}'-]/gu, "");$("user-name").textContent=`Hi, ${firstName||"there"} 👋`;$("user-email").textContent="";if(p?.picture){$("user-photo").src=p.picture;$("user-photo").classList.remove("hidden");}}
  function saveSession(){if(state.idTokenPayload)localStorage.setItem(SESSION_KEY,JSON.stringify({name:state.idTokenPayload.name||"Google user",email:state.idTokenPayload.email||"",picture:state.idTokenPayload.picture||"",sub:state.idTokenPayload.sub||""}));}
  function readSession(){try{return JSON.parse(localStorage.getItem(SESSION_KEY)||"null")}catch{return null}}
  function hideLogin(){$("login-card").classList.add("hidden");$("routine-page").classList.remove("hidden");$("google-signin-button").classList.add("hidden");$("grant-access").classList.add("hidden");$("sign-out").classList.remove("hidden");}
  function signOut(){const s=readSession();try{if(s?.sub)google.accounts.id.revoke(s.sub,()=>{});}catch(_){} localStorage.removeItem(SESSION_KEY);window.FM_AUTH_CACHE?.clear?.();location.reload();}
  function setStatus(text,error=false){const el=$("auth-status");if(el){el.textContent=text;el.className=`status ${error?"error":""}`;}}
  function setFormStatus(text,error=false){const el=$("routine-status");if(el){el.textContent=text;el.className=`status ${error?"error":""}`;}}
  function setSyncStatus(text,error=false){const el=$("sync-status");if(el){el.textContent=text;el.className=`muted ${error?"error":""}`;}}
  function escapeHtml(s){return String(s??"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}
  function escapeAttr(s){return escapeHtml(s);}
  function refreshSavedSessionIfNeeded(){const saved=readSession();if(!saved||!window.google?.accounts?.oauth2)return;if(window.FM_AUTH_CACHE?.read?.(saved.email)?.token)return;attemptSilentAccess(saved.email)}
  window.addEventListener("pageshow",refreshSavedSessionIfNeeded);
  document.addEventListener("visibilitychange",()=>{if(!document.hidden)refreshSavedSessionIfNeeded()});
})();