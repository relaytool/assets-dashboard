(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const state = {
    accessToken: null, idTokenPayload: null, alerts: [], inventory: [], transactions: [], auditTransactions: [], routines: [], clients: [],
    alertHeaders: [], alertFilter: "OPEN", search: "", attending: null, investigationAlert: null, investigationRows: []
  };
  let tokenRequestPromise = null;
  const $ = id => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", () => { bindEvents(); bindQuickMenu(); waitForGoogle(); });

  function bindEvents(){
    $("grant-access")?.addEventListener("click", () => requestAccess(false));
    $("sign-out")?.addEventListener("click", signOut);
    $("refresh")?.addEventListener("click", loadPage);
    $("alert-search")?.addEventListener("input", e => { state.search = e.target.value; renderAlerts(); });
    document.querySelectorAll("[data-alert-filter]").forEach(button => button.addEventListener("click", () => {
      state.alertFilter = button.dataset.alertFilter || "OPEN";
      document.querySelectorAll("[data-alert-filter]").forEach(b => b.classList.toggle("active", b === button));
      renderAlerts();
    }));
    $("alerts-body")?.addEventListener("click", event => {
      const investigate = event.target.closest("[data-investigate-alert]");
      if(investigate){ openInvestigation(investigate.dataset.investigateAlert); return; }
      const ignore = event.target.closest("[data-ignore-alert]");
      if(ignore){ ignoreAlert(ignore.dataset.ignoreAlert); return; }
      const attend = event.target.closest("[data-attend-alert]");
      if(attend){ openAttendModal(attend.dataset.attendAlert); }
    });
    $("cancel-attend")?.addEventListener("click", closeAttendModal);
    $("save-attend")?.addEventListener("click", attendAlert);
    $("close-attend")?.addEventListener("click", closeAttendModal);
    $("attend-comment")?.addEventListener("keydown", e => { if(e.key === "Enter" && (e.ctrlKey || e.metaKey)) attendAlert(); });
    $("investigate-client")?.addEventListener("change", renderInvestigation);
    $("investigate-asset")?.addEventListener("change", renderInvestigation);
    $("investigate-date")?.addEventListener("change", renderInvestigation);
    $("investigate-movement")?.addEventListener("change", renderInvestigation);
    $("clear-investigate-filters")?.addEventListener("click", () => {
      $("investigate-client").value = ""; $("investigate-asset").value = ""; $("investigate-date").value = ""; $("investigate-movement").value = ""; renderInvestigation();
    });
    $("investigate-body")?.addEventListener("click", event => {
      const button = event.target.closest("[data-investigate-info-index]");
      if(button) openAlertInfo(Number(button.dataset.investigateInfoIndex));
    });
    $("close-investigate")?.addEventListener("click", closeInvestigation);
    $("close-alert-info")?.addEventListener("click", closeAlertInfo);
    document.addEventListener("click", event => {
      if(event.target.matches("[data-close-attend]")) closeAttendModal();
      if(event.target.matches("[data-close-investigate]")) closeInvestigation();
      if(event.target.matches("[data-close-alert-info]")) closeAlertInfo();
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

  function waitForGoogle(){
    if(window.google?.accounts?.id&&window.google?.accounts?.oauth2)return initializeGoogle();
    let done=false; const started=Date.now();
    const onLoad=()=>{if(done)return;done=true;clearInterval(timer);initializeGoogle()};
    window.addEventListener("fm-google-loaded",onLoad,{once:true});
    const timer=setInterval(()=>{if(window.google?.accounts?.id&&window.google?.accounts?.oauth2)onLoad();else if(Date.now()-started>=8000){clearInterval(timer);setStatus("Google services could not be loaded. Check your internet connection.",true)}},50);
  }

  function initializeGoogle(){
    google.accounts.id.initialize({client_id:CONFIG.GOOGLE_CLIENT_ID,callback:handleCredentialResponse,auto_select:true,cancel_on_tap_outside:false});
    const saved=readSession();
    if(saved){
      setUserProfile(saved);
      $("google-signin-button")?.classList.add("hidden");
      $("grant-access").classList.add("hidden");
      const cached=window.FM_AUTH_CACHE?.read?.(saved.email);
      if(cached?.token){
        state.accessToken=cached.token;
        window.FM_CONNECTION_UI?.show("Restoring your session…","Using your saved Google Sheets connection.");
        hideLogin();
        loadPage().finally(()=>window.FM_CONNECTION_UI?.hide());
      }else{
        setStatus("Restoring your Sheets connection…");
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…","Restoring your saved Google session.");
        attemptSilentAccess(saved.email);
      }
    }else{
      window.FM_CONNECTION_UI?.hide();
      google.accounts.id.renderButton($("google-signin-button"),{theme:"outline",size:"large",text:"signin_with",shape:"rectangular",width:280});
      google.accounts.id.prompt();
    }
  }
  function handleCredentialResponse(response){try{state.idTokenPayload=decodeJwtPayload(response.credential);saveSession();setUserProfile(state.idTokenPayload);requestAccess(true)}catch(e){console.error(e);setStatus("Google sign-in response could not be read.",true)}}
  function requestAccess(silent=true){window.FM_CONNECTION_UI?.show(silent?"Connecting to Google Sheets…":"Authorising Google Sheets & Drive…",silent?"Restoring your saved connection.":"Approve access in the Google prompt to continue.");acquireAccessToken(silent?"none":"consent").then(async()=>{hideLogin();await loadPage();window.FM_CONNECTION_UI?.hide()}).catch(e=>{window.FM_CONNECTION_UI?.hide();$("grant-access").classList.remove("hidden");setStatus(e.message||"Google authorization failed.",true)})}
  function attemptSilentAccess(email){window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…","Checking your current Google authorisation.");acquireAccessToken("none",email).then(async()=>{hideLogin();await loadPage();window.FM_CONNECTION_UI?.hide()}).catch(()=>{window.FM_CONNECTION_UI?.hide();$("grant-access").classList.remove("hidden");setStatus("Your Google session is available. Allow Sheets & Drive access to continue.")})}
  function acquireAccessToken(prompt="none",email,options={}){const forceRefresh=Boolean(options.forceRefresh);const expectedEmail=email||state.idTokenPayload?.email||readSession()?.email||"";if(!forceRefresh){const cached=window.FM_AUTH_CACHE?.read?.(expectedEmail);if(cached?.token){state.accessToken=cached.token;return Promise.resolve(cached.token)}}if(tokenRequestPromise)return tokenRequestPromise;tokenRequestPromise=new Promise((resolve,reject)=>{let settled=false;const finish=(fn,v)=>{if(settled)return;settled=true;tokenRequestPromise=null;clearTimeout(timeoutId);fn(v)};const timeoutId=setTimeout(()=>finish(reject,new Error("Google authorisation did not complete.")),prompt==="none"?5000:15000);const tc=google.accounts.oauth2.initTokenClient({client_id:CONFIG.GOOGLE_CLIENT_ID,scope:CONFIG.OAUTH_SCOPES,callback:r=>{if(r.error){finish(reject,new Error(`Google authorization failed: ${r.error}`));return}state.accessToken=r.access_token;window.FM_AUTH_CACHE?.write?.(r.access_token,r.expires_in,expectedEmail);finish(resolve,r.access_token)}});tc.requestAccessToken({prompt,login_hint:expectedEmail||undefined})});return tokenRequestPromise}
  async function loadPage(){
    if(!state.accessToken)return;
    setSyncStatus("Refreshing live alerts...");
    try{
      const id=CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const meta=await sheetsGet(`/${encodeURIComponent(id)}`);
      const titles=(meta.sheets||[]).map(s=>s.properties.title);
      const missing=[CONFIG.ROUTINE_SHEET_NAME,CONFIG.ALERTS_SHEET_NAME].filter(n=>!titles.includes(n));
      if(missing.length)await createSheets(missing);
      const [inventoryRows,transactionRows,clientRows,routineRows,alertRows,operationRows,adjustmentRows]=await Promise.all([
        getValues(id,CONFIG.INVENTORY_SHEET_NAME),
        getValues(id,CONFIG.TRANSACTIONS_SHEET_NAME),
        getValues(id,CONFIG.CLIENT_LIST_SHEET_NAME),
        getValues(id,CONFIG.ROUTINE_SHEET_NAME),
        getValues(id,CONFIG.ALERTS_SHEET_NAME),
        titles.includes(CONFIG.ASSET_OPERATIONS_SHEET_NAME)?getValues(id,CONFIG.ASSET_OPERATIONS_SHEET_NAME):Promise.resolve([]),
        titles.includes(CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME)?getValues(id,CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME):Promise.resolve([])
      ]);
      await ensureAlertHeaders(id, alertRows);
      state.inventory=parseInventory(inventoryRows);
      state.transactions=parseTransactions(transactionRows);
      state.auditTransactions=[...state.transactions,...parseTransactions(operationRows),...parseTransactions(adjustmentRows)];
      state.clients=parseClients(clientRows);
      state.routines=FM_ROUTINE_ALERTS.parseRoutineRows(routineRows);
      state.alerts=FM_ROUTINE_ALERTS.parseAlertRows(alertRows);
      await rebuildAlertSnapshot();
      $("workbook-name").textContent=CONFIG.INVENTORY_LEDGER_NAME||"Assets Inventory Ledger";
      renderAlerts();
      updateAlertNavBadge();
      setSyncStatus(`Live snapshot · ${new Date().toLocaleTimeString()}`);
    }catch(e){console.error(e);setSyncStatus(e.message||"Unable to load alerts.",true)}
  }

  async function createSheets(names){
    const id=CONFIG.INVENTORY_LEDGER_SHEET_ID;
    await sheetsPost(`/${encodeURIComponent(id)}:batchUpdate`,{requests:names.map(title=>({addSheet:{properties:{title}}}))});
    if(names.includes(CONFIG.ROUTINE_SHEET_NAME))await updateValues(id,CONFIG.ROUTINE_SHEET_NAME,[FM_ROUTINE_ALERTS.ROUTINE_HEADERS]);
    if(names.includes(CONFIG.ALERTS_SHEET_NAME))await updateValues(id,CONFIG.ALERTS_SHEET_NAME,[FM_ROUTINE_ALERTS.ALERT_HEADERS]);
  }

  async function ensureAlertHeaders(id, rows){
    let headers=rows[0]?.slice()||FM_ROUTINE_ALERTS.ALERT_HEADERS.slice();
    const normalized=headers.map(normalizeHeader);
    const missing=FM_ROUTINE_ALERTS.ALERT_HEADERS.filter(h=>!normalized.includes(normalizeHeader(h)));
    if(missing.length){headers.push(...missing);const range=`${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A1`;await sheetsPut(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:[headers]});}
    state.alertHeaders=headers;
  }

  function alertRowFromEntity(entity,status){
    const headers=state.alertHeaders.length?state.alertHeaders:FM_ROUTINE_ALERTS.ALERT_HEADERS;
    const normalized=headers.map(normalizeHeader); const row=Array.isArray(entity.rawRow)?headers.map((_,index)=>entity.rawRow[index]??""):new Array(headers.length).fill("");
    const put=(aliases,value)=>{const idx=aliases.map(normalizeHeader).map(a=>normalized.indexOf(a)).find(i=>i>=0);if(idx>=0)row[idx]=value??"";};
    put(["Alert Key","Key"],entity.key); put(["Created At","Timestamp"],entity.createdAt||new Date().toISOString()); put(["Alert Type","Type"],entity.type); put(["Status"],status||entity.status||"OPEN");
    put(["Client"],entity.client); put(["Asset"],entity.asset); put(["Required Qty","Required"],entity.requiredQty); put(["Available Qty","Available"],entity.availableQty); put(["Shortfall","Short By"],entity.shortfall); put(["Destination","Location"],entity.destination); put(["Message","Alert"],entity.message);
    put(["Attended At"],entity.attendedAt); put(["Attended By","Attended User"],entity.attendedBy); put(["Resolution Comment","Resolution Note"],entity.resolutionComment); put(["Routine Key"],entity.routineKey);
    put(["Ignored At","Ignore At"],entity.ignoredAt); put(["Ignored By","Ignored User"],entity.ignoredBy); put(["Ignore Comment","Ignore Note"],entity.ignoreComment);
    return row;
  }

  async function rebuildAlertSnapshot(){
    const alertLedgerTransactions=[...state.transactions,...state.auditTransactions.filter(item=>String(item.movement||"").trim().toUpperCase()==="NORMALISATION")];
    const candidates=FM_ROUTINE_ALERTS.buildAlertCandidates({inventory:state.inventory,transactions:alertLedgerTransactions,routines:state.routines,clients:state.clients,date:new Date()});
    const retainedMap=new Map();
    state.alerts.filter(a=>a.status==="ATTENDED"||a.status==="IGNORED").forEach(a=>{if(a.key)retainedMap.set(a.key,{...a,rawRow:Array.isArray(a.rawRow)?a.rawRow.slice():undefined});});
    const retained=[...retainedMap.values()];
    const retainedKeys=new Set(retained.map(a=>a.key));
    const fresh=candidates.filter(c=>c.key&&!retainedKeys.has(c.key));
    const now=new Date().toISOString();
    const generated=fresh.map(c=>({rowNumber:0,key:c.key,createdAt:now,type:c.type,status:"OPEN",client:c.client||"",asset:c.asset||"",requiredQty:c.requiredQty||0,availableQty:c.availableQty||0,shortfall:c.shortfall||0,destination:c.destination||"",message:c.message||"",attendedAt:"",attendedBy:"",resolutionComment:"",routineKey:c.routineKey||"",ignoredAt:"",ignoredBy:"",ignoreComment:""}));
    const finalAlerts=[...retained,...generated].map((a,i)=>({...a,rowNumber:i+2}));
    const headers=state.alertHeaders.length?state.alertHeaders:FM_ROUTINE_ALERTS.ALERT_HEADERS;
    const lastCol=columnLetter(headers.length);
    const clearRange=`${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A2:${lastCol}`;
    await sheetsPost(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(clearRange)}:clear`,{});
    const writeRange=`${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A1:${lastCol}${finalAlerts.length+1}`;
    await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(writeRange)}?valueInputOption=USER_ENTERED`,{range:writeRange,majorDimension:"ROWS",values:[headers,...finalAlerts.map(a=>alertRowFromEntity(a,a.status))]});
    state.alerts=finalAlerts;
  }

  function renderAlerts(){
    const all=[...state.alerts].sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
    const open=all.filter(a=>a.status==="OPEN"), ignored=all.filter(a=>a.status==="IGNORED"), attended=all.filter(a=>a.status==="ATTENDED");
    $("open-count").textContent=open.length.toLocaleString();$("inventory-count").textContent=open.filter(a=>a.type==="INVENTORY").length.toLocaleString();$("client-count").textContent=open.filter(a=>a.type==="CLIENT").length.toLocaleString();$("ignored-count").textContent=ignored.length.toLocaleString();$("attended-count").textContent=attended.length.toLocaleString();
    $("tab-open-count").textContent=open.length.toLocaleString();$("tab-ignored-count").textContent=ignored.length.toLocaleString();$("tab-attended-count").textContent=attended.length.toLocaleString();
    const needle=state.search.trim().toLowerCase();
    let visible=(state.alertFilter==="ALL"?all:all.filter(a=>a.status===state.alertFilter));
    if(needle)visible=visible.filter(a=>[a.client,a.asset,a.message,a.destination,a.type,a.status].some(v=>String(v||"").toLowerCase().includes(needle)));
    $("alerts-body").innerHTML=visible.length?visible.map(renderAlertCard).join(""):"";
    $("alerts-empty").classList.toggle("hidden",visible.length>0);
    $("alerts-empty-title").textContent=state.alertFilter==="OPEN"?"No open alerts":"Nothing in this view";
    $("alerts-empty-text").textContent=state.alertFilter==="OPEN"?"The live snapshot currently has no open alerts.":"Try another filter or search term.";
  }

  function renderAlertCard(a){
    const open=a.status==="OPEN", ignored=a.status==="IGNORED", attended=a.status==="ATTENDED";
    const typeClass=a.type==="INVENTORY"?"inventory":"client";
    const title=a.type==="INVENTORY"?(a.asset||"Inventory"):(a.client||"Client");
    const secondary=a.type==="INVENTORY"?(a.destination||"Warehouse stock"):(a.asset||"Client balance");
    const metric=a.type==="INVENTORY"?`<div><span>Available</span><strong>${FM_ROUTINE_ALERTS.formatQty(a.availableQty)}</strong></div><div><span>Required</span><strong>${FM_ROUTINE_ALERTS.formatQty(a.requiredQty)}</strong></div><div><span>Shortfall</span><strong>${FM_ROUTINE_ALERTS.formatQty(a.shortfall)}</strong></div>`:`<div><span>With client</span><strong>${FM_ROUTINE_ALERTS.formatQty(a.availableQty)}</strong></div><div><span>Asset</span><strong>${escapeHtml(a.asset||"—")}</strong></div>`;
    return `<article class="alert-card alert-card-${typeClass} ${ignored||attended?"is-closed":"is-open"}">\n      <div class="alert-card-top"><span class="alert-type ${a.type==="INVENTORY"?"alert-inventory":"alert-client"}">${a.type==="INVENTORY"?"Inventory":"Client"}</span><span class="alert-status ${open?"open":ignored?"ignored":"attended"}">${open?"Open":ignored?"Ignored":"Attended"}</span></div>\n      <div class="alert-card-title"><div><h3>${escapeHtml(title)}</h3><p>${escapeHtml(secondary)}</p></div><span class="alert-dot" aria-hidden="true"></span></div>\n      <p class="alert-card-message">${escapeHtml(a.message||"")}</p>\n      <div class="alert-metrics ${a.type==="CLIENT"?"alert-metrics-client":""}">${metric}</div>\n      <div class="alert-card-meta"><span>${escapeHtml(a.destination||"No destination")}</span><span>${escapeHtml(formatAlertTime(a.createdAt))}</span></div>\n      ${attended&&a.resolutionComment?`<div class="alert-resolution"><strong>Resolved by ${escapeHtml(a.attendedBy||"user")}</strong><span>${escapeHtml(a.resolutionComment)}</span></div>`:""}\n      ${ignored&&a.ignoreComment?`<div class="alert-resolution"><strong>Ignored by ${escapeHtml(a.ignoredBy||"user")}</strong><span>${escapeHtml(a.ignoreComment)}</span></div>`:""}\n      <div class="alert-card-actions"><button type="button" class="secondary small-button" data-investigate-alert="${escapeAttr(a.key)}">Investigate</button>${open?`<button type="button" class="secondary small-button" data-ignore-alert="${escapeAttr(a.key)}">Ignore</button><button type="button" class="primary small-button" data-attend-alert="${escapeAttr(a.key)}">Attend</button>`:""}</div>\n    </article>`;
  }

  function openAttendModal(key){const alert=state.alerts.find(x=>x.key===key);if(!alert||alert.status!=="OPEN")return;state.attending=alert;$("attend-alert-title").textContent=alert.type==="INVENTORY"?"Attend inventory alert":"Attend client alert";$("attend-alert-message").textContent=alert.message;$("attend-comment").value="";$("attend-status").textContent="";$("attend-modal").classList.remove("hidden");setTimeout(()=>$('attend-comment').focus(),50)}
  async function attendAlert(){const alert=state.attending;if(!alert)return;const comment=$("attend-comment").value.trim();if(comment.length<2){$("attend-status").textContent="Add a short comment explaining how the alert was resolved.";return;}const button=$("save-attend");if(button.disabled)return;button.disabled=true;$("attend-status").textContent="Saving...";const previous={status:alert.status,attendedAt:alert.attendedAt,attendedBy:alert.attendedBy,resolutionComment:alert.resolutionComment};try{const now=new Date().toISOString();const by=state.idTokenPayload?.email||readSession()?.email||"Google user";Object.assign(alert,{status:"ATTENDED",attendedAt:now,attendedBy:by,resolutionComment:comment});await writeAlertRow(alert);closeAttendModal();renderAlerts();updateAlertNavBadge()}catch(e){console.error(e);Object.assign(alert,previous);$("attend-status").textContent=e.message||"Unable to save attendance."}finally{button.disabled=false}}
  async function ignoreAlert(key){const alert=state.alerts.find(x=>x.key===key);if(!alert||alert.status!=="OPEN")return;const button=[...document.querySelectorAll("[data-ignore-alert]")].filter(b=>b.dataset.ignoreAlert===key);button.forEach(b=>b.disabled=true);const previous={status:alert.status,ignoredAt:alert.ignoredAt,ignoredBy:alert.ignoredBy,ignoreComment:alert.ignoreComment};try{const now=new Date().toISOString();const by=state.idTokenPayload?.email||readSession()?.email||"Google user";Object.assign(alert,{status:"IGNORED",ignoredAt:now,ignoredBy:by});await writeAlertRow(alert);renderAlerts();updateAlertNavBadge()}catch(e){console.error(e);Object.assign(alert,previous);setSyncStatus(e.message||"Unable to ignore alert.",true)}finally{button.forEach(b=>b.disabled=false)}}
  async function writeAlertRow(alert){const lastCol=columnLetter(state.alertHeaders.length);const range=`${quoteSheetName(CONFIG.ALERTS_SHEET_NAME)}!A${alert.rowNumber}:${lastCol}${alert.rowNumber}`;await sheetsPut(`/${encodeURIComponent(CONFIG.INVENTORY_LEDGER_SHEET_ID)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:[alertRowFromEntity(alert,alert.status)]})}
  function closeAttendModal(){state.attending=null;$("attend-modal")?.classList.add("hidden")}

  function openInvestigation(key){
    const alert=state.alerts.find(x=>x.key===key);if(!alert)return;state.investigationAlert=alert;
    $("investigate-title").textContent=`Investigate ${alert.type==="INVENTORY"?(alert.asset||"inventory"):(alert.client||"client")}`;
    $("investigate-summary").textContent=alert.message||"Review the related audit history below.";
    const summary=$("investigate-alert-summary");
    if(summary){
      const valueLabel=alert.type==="INVENTORY"?"Available":"With client";
      const value=FM_ROUTINE_ALERTS.formatQty(alert.availableQty);
      summary.innerHTML=`<div class="investigate-summary-item"><span>Alert</span><strong>${escapeHtml(alert.type==="INVENTORY"?"Inventory issue":"Client balance")}</strong></div><div class="investigate-summary-item"><span>${escapeHtml(valueLabel)}</span><strong>${escapeHtml(value)}</strong></div><div class="investigate-summary-item"><span>Required / shortfall</span><strong>${escapeHtml(`${FM_ROUTINE_ALERTS.formatQty(alert.requiredQty)} / ${FM_ROUTINE_ALERTS.formatQty(alert.shortfall)}`)}</strong></div><div class="investigate-summary-item"><span>Destination</span><strong>${escapeHtml(alert.destination||"—")}</strong></div>`;
    }
    const dayFromKey=parseAlertDay(alert.key);
    populateInvestigationFilters(alert);
    $("investigate-client").value=alert.type==="CLIENT"?alert.client:"";
    $("investigate-asset").value=alert.asset||"";
    $("investigate-date").value=dayFromKey;
    $("investigate-movement").value="";
    $("alert-investigate-modal").classList.remove("hidden");
    renderInvestigation();
  }
  function populateInvestigationFilters(alert){
    const auditRows=state.auditTransactions.length?state.auditTransactions:state.transactions;
    const clients=[...new Set(auditRows.map(t=>t.client).filter(Boolean))].sort((a,b)=>a.localeCompare(b));const assets=[...new Set(auditRows.map(t=>t.asset).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    const client=$('investigate-client'),asset=$('investigate-asset'); const currentClient=client.value,currentAsset=asset.value;
    client.innerHTML=`<option value="">All clients</option>${clients.map(v=>`<option value="${escapeAttr(v)}">${escapeHtml(v)}</option>`).join("")}`;
    asset.innerHTML=`<option value="">All assets</option>${assets.map(v=>`<option value="${escapeAttr(v)}">${escapeHtml(v)}</option>`).join("")}`;
    if(clients.includes(currentClient))client.value=currentClient;if(assets.includes(currentAsset))asset.value=currentAsset;
  }
  function investigationMovementClass(value){
    const movement=String(value||"").trim().toUpperCase();
    if(movement==="RECEIVED"||movement==="COMBINE_IN")return "received";
    if(movement==="SENT")return "sent";
    if(movement==="DISCARD"||movement==="COMBINE_OUT")return "discard";
    return "other";
  }
  function renderInvestigation(){
    const client=$("investigate-client")?.value||"", asset=$("investigate-asset")?.value||"", date=$("investigate-date")?.value||"", movement=$("investigate-movement")?.value||"";
    const auditRows=state.auditTransactions.length?state.auditTransactions:state.transactions;
    const rows=auditRows.filter(t=>(!client||sameText(t.client,client))&&(!asset||sameText(t.asset,asset))&&(!date||dateKey(t.timestamp)===date)&&(!movement||String(t.movement||"").toUpperCase()===movement));
    const groups=FM_TRANSACTION_DISPLAY.groupTransactions(rows);
    state.investigationRows=groups;
    $('investigate-count').textContent=`${groups.length.toLocaleString()} grouped movement${groups.length===1?"":"s"}`;
    $('investigate-body').innerHTML=groups.length?groups.map((group,index)=>{
      const sections=FM_TRANSACTION_DISPLAY.movementSections(group);
      return `<article class="transaction-card ${FM_TRANSACTION_DISPLAY.cardClass(group)}">
        <div class="transaction-card-accent" aria-hidden="true"></div>
        <div class="transaction-card-top">
          <div class="transaction-card-icon-wrap"><img class="transaction-card-icon" src="truck-icon.png" alt="" aria-hidden="true"></div>
          <div class="transaction-card-heading"><div class="transaction-card-client">${escapeHtml(group.client||"Unknown client")}</div><div class="transaction-card-meta"><span>${escapeHtml(formatAlertTime(group.timestamp))}</span><span>•</span><span>${escapeHtml(FM_TRANSACTION_DISPLAY.movementLabels(group).join(" + "))}</span>${group.vehicleReg?`<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>`:""}</div></div>
          <button type="button" class="secondary info-button transaction-card-view" data-investigate-info-index="${index}">View details</button>
        </div>
        <div class="transaction-card-body">${sections.map(section=>{const cls=investigationMovementClass(section.movement ?? section.key);return `<section class="transaction-movement-section movement-${cls}"><div class="transaction-movement-heading"><span class="movement-tag ${cls}">${escapeHtml(section.label)}</span><span>${FM_ROUTINE_ALERTS.formatQty(section.totalQuantity)} total</span></div><div class="transaction-asset-list">${section.items.map(item=>`<div class="transaction-asset-pill"><span>${escapeHtml(item.asset)}</span><strong>${FM_ROUTINE_ALERTS.formatQty(item.quantity)}</strong></div>`).join("")}</section>`;}).join("")}</div>
        <div class="transaction-card-footer"><span>${FM_TRANSACTION_DISPLAY.formatLineCount(group)}${group.image?" · attachment":""}</span><strong>${FM_ROUTINE_ALERTS.formatQty(group.totalQuantity)} total units</strong></div>
      </article>`;
    }).join(""):`<div class="transaction-empty">No matching asset movements.</div>`;
    $('investigate-empty').classList.toggle('hidden',groups.length>0);
  }
  function closeInvestigation(){$("alert-investigate-modal")?.classList.add("hidden");state.investigationAlert=null;state.investigationRows=[]}

  async function openAlertInfo(index){
    const group=state.investigationRows[index];if(!group)return;
    const title=$("alert-info-title");
    if(title)title.textContent=`${group.client||"Movement"} · Details`;
    const summary=$("alert-info-summary");
    if(summary){
      const sections=FM_TRANSACTION_DISPLAY.movementSections(group);
      const comments=Array.isArray(group.comments)&&group.comments.length?group.comments:(group.comment?[group.comment]:[]);
      summary.innerHTML=`<div class="transaction-detail-note"><strong>${escapeHtml(group.client||"Unknown client")}</strong><span> · ${escapeHtml(formatAlertTime(group.timestamp)||"Unknown time")}</span><span> · ${escapeHtml(FM_TRANSACTION_DISPLAY.movementLabels(group).join(" + ")||"Movement")}</span><span> · ${FM_ROUTINE_ALERTS.formatQty(group.totalQuantity)} total units</span>${group.vehicleReg?`<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>`:""}</div>${sections.map(section=>`<div class="transaction-detail-section"><div class="transaction-detail-section-head"><span class="movement-tag ${String(section.movement ?? "OTHER").toLowerCase()}">${escapeHtml(section.label)}</span><strong>${FM_ROUTINE_ALERTS.formatQty(section.totalQuantity)}</strong></div><div class="transaction-detail-assets">${section.items.map(item=>`<span>${escapeHtml(item.asset)} <strong>${FM_ROUTINE_ALERTS.formatQty(item.quantity)}</strong></span>`).join("")}</div></div>`).join("")}${comments.length?`<div class="transaction-detail-note"><strong>${comments.length===1?"Comment":"Comments"}</strong><br>${comments.map(comment=>escapeHtml(comment)).join("<br>")}</div>`:""}`;
    }
    const wrap=$("alert-info-photo-wrap"),img=$("alert-info-photo"),status=$("alert-info-photo-status"),comment=$("alert-info-comment"),empty=$("alert-info-empty");
    FM_MEDIA?.revokeObjectUrl?.(img);
    wrap.classList.toggle("hidden",!group.image);status.classList.toggle("hidden",!group.image);status.textContent="Loading photo...";
    comment.classList.toggle("hidden",!group.comment);comment.textContent=group.comment||"";empty.classList.toggle("hidden",Boolean(group.image||group.comment));
    $("alert-info-modal").classList.remove("hidden");
    if(group.image){$("alert-info-photo-link").href=group.image;const loaded=await FM_MEDIA.loadDriveImage(img,group.image,state.accessToken);status.classList.toggle("hidden",loaded);if(!loaded)status.textContent="Preview unavailable here. Use Open full size in Drive.";}
  }
  function closeAlertInfo(){const img=$("alert-info-photo");if(img)FM_MEDIA?.revokeObjectUrl?.(img);$("alert-info-modal")?.classList.add("hidden")}

  function updateAlertNavBadge(){const open=state.alerts.filter(a=>a.status==="OPEN").length;document.querySelectorAll("#alerts-nav-badge").forEach(el=>{el.textContent=open.toLocaleString();el.classList.toggle("hidden",open===0)})}
  function parseAlertDay(key){const parts=String(key||"").split("|");return /^\\d{4}-\\d{2}-\\d{2}$/.test(parts[1]||"")?parts[1]:""}
  function dateKey(value){const d=new Date(value);if(Number.isNaN(d.getTime()))return"";return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}
  function movementLabel(x){const v=String(x||"").toUpperCase();return v==="RECEIVED"?"Received":v==="SENT"?"Sent":v==="DISCARD"?"Discard":String(x||"")}
  function movementClass(x){const v=String(x||"").trim().toLowerCase();return v==="received"?"movement-received":v==="sent"?"movement-sent":v==="discard"?"movement-discard":"movement-other"}
  function sameText(a,b){return String(a||"").trim().toLowerCase()===String(b||"").trim().toLowerCase()}
  function formatAlertTime(value){if(!value)return"";const d=new Date(value);return Number.isNaN(d.getTime())?String(value):d.toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}
  function parseInventory(rows){if(!rows.length)return[];const h=rows[0].map(x=>String(x??"").trim().toLowerCase());const a=h.findIndex(x=>["asset","asset name","item","type"].includes(x));const b=h.findIndex(x=>["balance","current balance","stock","quantity"].includes(x));return rows.slice(1).map((r,i)=>({rowNumber:i+2,asset:String(r[a]??"").trim(),balance:Number(String(r[b]??"0").replace(/,/g,""))||0})).filter(x=>x.asset)}
  function parseTransactions(rows){if(!rows.length)return[];const h=rows[0].map(normalizeHeader);const i={timestamp:h.findIndex(x=>["timestamp","date","datetime"].includes(x)),client:h.findIndex(x=>["client","client name"].includes(x)),movement:h.findIndex(x=>["movement","type","direction"].includes(x)),asset:h.findIndex(x=>["asset","asset name","item"].includes(x)),quantity:h.findIndex(x=>["quantity","qty"].includes(x)),user:h.findIndex(x=>["user","entered by","email"].includes(x)),comment:h.findIndex(x=>["comment","comments","notes"].includes(x)),image:h.findIndex(x=>["image","image link","photo","picture","photo link","attachment","drive link"].includes(x))};return rows.slice(1).map(r=>({timestamp:i.timestamp>=0?r[i.timestamp]??"":"",client:i.client>=0?r[i.client]??"":"",movement:i.movement>=0?r[i.movement]??"":"",asset:i.asset>=0?r[i.asset]??"":"",quantity:i.quantity>=0?Number(String(r[i.quantity]??"0").replace(/,/g,""))||0:0,user:i.user>=0?r[i.user]??"":"",comment:i.comment>=0?String(r[i.comment]??"").trim():"",image:i.image>=0?String(r[i.image]??"").trim():""})).filter(x=>x.asset||x.client)}
  function parseClients(rows){if(!rows.length)return[];const h=rows[0].map(x=>String(x??"").trim().toLowerCase());const i=h.findIndex(x=>["client","client name","name"].includes(x));return i<0?[]:[...new Set(rows.slice(1).map(r=>String(r[i]??"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b))}

  async function getValues(id,sheet){const d=await sheetsGet(`/${encodeURIComponent(id)}/values/${encodeURIComponent(`${quoteSheetName(sheet)}!A:Z`)}`);return d.values||[]}
  async function updateValues(id,sheet,rows){const range=`${quoteSheetName(sheet)}!A1`;return sheetsPut(`/${encodeURIComponent(id)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:rows})}
  async function sheetsGet(path){return fetchJson(SHEETS_API+path)}
  async function sheetsPost(path,body){return fetchJson(SHEETS_API+path,{method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)})}
  async function sheetsPut(path,body){return fetchJson(SHEETS_API+path,{method:"PUT",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)})}
  async function fetchJson(url,options={}){let r=await fetch(url,{...options,headers:{...authHeaders(),...(options.headers||{})}});if(r.status===401&&!options.__retried){try{await acquireAccessToken("none",state.idTokenPayload?.email||readSession()?.email,{forceRefresh:true});return fetchJson(url,{...options,__retried:true,headers:{...(options.headers||{}),...authHeaders()}})}catch(_){state.accessToken=null;window.FM_AUTH_CACHE?.clear?.();$("grant-access")?.classList.remove("hidden");throw new Error("Google Sheets access expired. Allow Sheets & Drive access to reconnect.")}}const t=await r.text();let d={};try{d=t?JSON.parse(t):{}}catch{}if(!r.ok)throw new Error(d.error?.message||`Request failed (${r.status})`);return d}
  function authHeaders(){return{Authorization:`Bearer ${state.accessToken}`}}
  function quoteSheetName(name){return`'${String(name).replace(/'/g,"''")}'`}
  function columnLetter(n){let r="";while(n>0){const rem=(n-1)%26;r=String.fromCharCode(65+rem)+r;n=Math.floor((n-1)/26)}return r}
  function normalizeHeader(x){return String(x??"").trim().toLowerCase().replace(/\s+/g," ")}
  function decodeJwtPayload(token){const part=token.split(".")[1],n=part.replace(/-/g,"+").replace(/_/g,"/")+"=".repeat((4-part.length%4)%4);return JSON.parse(decodeURIComponent(Array.from(atob(n)).map(c=>`%${c.charCodeAt(0).toString(16).padStart(2,"0")}`).join("")))}
  function setUserProfile(p){const rawName=String(p?.given_name||p?.name||"Google user").trim();const firstName=(rawName.split(/\s+/)[0]||"there").replace(/[^\p{L}\p{M}'-]/gu, "");$("user-name").textContent=`Hi, ${firstName||"there"} 👋`;$("user-email").textContent="";if(p?.picture){$("user-photo").src=p.picture;$("user-photo").classList.remove("hidden")}}
  function saveSession(){if(state.idTokenPayload)localStorage.setItem(SESSION_KEY,JSON.stringify({name:state.idTokenPayload.name||"Google user",email:state.idTokenPayload.email||"",picture:state.idTokenPayload.picture||"",sub:state.idTokenPayload.sub||""}))}
  function readSession(){try{return JSON.parse(localStorage.getItem(SESSION_KEY)||"null")}catch{return null}}
  function hideLogin(){$("login-card").classList.add("hidden");$("alerts-page").classList.remove("hidden");$("google-signin-button").classList.add("hidden");$("grant-access").classList.add("hidden");$("sign-out").classList.remove("hidden")}
  function signOut(){const s=readSession();try{if(s?.sub)google.accounts.id.revoke(s.sub,()=>{})}catch(_){}localStorage.removeItem(SESSION_KEY);window.FM_AUTH_CACHE?.clear?.();location.reload()}
  function setStatus(t,e=false){const x=$("auth-status");if(x){x.textContent=t;x.className=`status ${e?"error":""}`}}
  function setSyncStatus(t,e=false){const x=$("sync-status");if(x){x.textContent=t;x.className=`muted ${e?"error":""}`}}
  function emptyRow(colspan,text){return`<tr><td colspan="${colspan}" class="empty">${escapeHtml(text)}</td></tr>`}
  function escapeHtml(s){return String(s??"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]))}
  function escapeAttr(s){return escapeHtml(s)}
  function refreshSavedSessionIfNeeded(){const saved=readSession();if(!saved||!window.google?.accounts?.oauth2)return;if(window.FM_AUTH_CACHE?.read?.(saved.email)?.token)return;attemptSilentAccess(saved.email)}
  window.addEventListener("pageshow",refreshSavedSessionIfNeeded);
  document.addEventListener("visibilitychange",()=>{if(!document.hidden)refreshSavedSessionIfNeeded()});
})();