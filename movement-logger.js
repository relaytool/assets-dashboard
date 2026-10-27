(() => {
  "use strict";

  const SHEETS_API="https://sheets.googleapis.com/v4/spreadsheets";
  const DRIVE_API="https://www.googleapis.com/drive/v3/files";
  const DRIVE_UPLOAD_API="https://www.googleapis.com/upload/drive/v3/files";
  const SELF_CLIENT=CONFIG.SELF_CLIENT_NAME||"HSC London (Self)";
  const SESSION_KEY="fmAssetSession";
  const state={idTokenPayload:null,accessToken:null,workbookId:null,workbookName:null,inventory:[],transactions:[],clients:[],brokers:[...(window.FM_BROKERS?.DEFAULT_BROKERS||[])],transactionIdx:null,transactionHeaderLength:6,driveFolderCache:{},auditVisibleRows:[]};
  let pendingMovement=null;
  let movementSubmitBusy=false;
  let batchForm=null;
  let tokenRequestPromise=null;
  const $=id=>document.getElementById(id);
  const lowerText=value=>String(value??"").trim().toLowerCase();

  document.addEventListener("DOMContentLoaded",()=>{
    bindEvents();
    bindQuickMenu();
    batchForm=FM_BATCH_FORM?.create?.({
      containerId:"client-movement-entries",
      addButtonId:"add-client-entry",
      selfClient:SELF_CLIENT,
      getClients:()=>state.clients,
      getBrokers:()=>state.brokers,
      getAssets:()=>state.inventory.map(item=>String(item?.asset??"").trim()).filter(Boolean),
      onChange:()=>updatePreview()
    });
    setDefaultTimestamp();
    handleMovementChange();
    waitForGoogle();
  });

  function bindEvents(){
    $("grant-access").addEventListener("click",()=>requestSheetAccess(false));
    $("sign-out").addEventListener("click",signOut);
    $("refresh").addEventListener("click",loadLogger);
    $("movement-form").addEventListener("submit",reviewMovement);
    $("cancel-confirm").addEventListener("click",closeConfirm);
    $("approve-confirm").addEventListener("click",approveMovement);
    document.addEventListener("click",e=>{
      const summaryToggle=e.target.closest("[data-summary-toggle]");
      if(summaryToggle){e.preventDefault();toggleTransactionSummary(summaryToggle);return;}
      if(e.target.matches("[data-close-confirm]"))closeConfirm();
      if(e.target.matches("[data-close-info]"))closeInfoModal();
    });
    $("movement-photo").addEventListener("change",handlePhotoChange);
    $("clear-photo").addEventListener("click",clearPhoto);
    $("take-photo")?.addEventListener("click",()=>$("movement-photo-camera")?.click());
    $("movement-photo-camera")?.addEventListener("change",event=>{
      const file=event.target.files?.[0];
      const main=$("movement-photo");
      if(!file||!main)return;
      try{
        const transfer=new DataTransfer();
        transfer.items.add(file);
        main.files=transfer.files;
      }catch(_){
        window._fmCameraFile=file;
      }
      handlePhotoChange();
    });
    $("filter-client").addEventListener("change",()=>renderAudit());
    $("filter-date").addEventListener("change",()=>renderAudit());
     $("filter-asset").addEventListener("change",()=>renderAudit());
    $("clear-filters").addEventListener("click",()=>{$("filter-client").value="";$("filter-date").value="";$("filter-asset").value="";renderAudit();});
    $("transactions-body").addEventListener("click",e=>{const b=e.target.closest("[data-info-index]");if(b)openInfoModal(Number(b.dataset.infoIndex));});
    $("close-info").addEventListener("click",closeInfoModal);
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
    const saved = localStorage.getItem(key);
    setOpen(isMobile() ? false : saved === "1");
    toggle.addEventListener("click", e => { e.preventDefault(); setOpen(!nav.classList.contains("menu-open")); });
    links?.addEventListener("click", e => { if (e.target.closest("a") && isMobile()) setOpen(false); });
    const onResize = e => { if (e.matches) setOpen(false); };
    if (mq.addEventListener) mq.addEventListener("change", onResize); else mq.addListener(onResize);
  }

  function waitForGoogle(){
    if(window.google?.accounts?.id&&window.google?.accounts?.oauth2){initializeGoogle();return;}

    let finished=false;
    const started=Date.now();
    const onLoaded=()=>{
      if(finished)return;
      finished=true;
      window.removeEventListener("fm-google-loaded",onLoaded);
      initializeGoogle();
    };
    window.addEventListener("fm-google-loaded",onLoaded,{once:true});

    const timer=setInterval(()=>{
      if(finished){clearInterval(timer);return;}
      if(window.google?.accounts?.id&&window.google?.accounts?.oauth2){
        clearInterval(timer);
        finished=true;
        window.removeEventListener("fm-google-loaded",onLoaded);
        initializeGoogle();
        return;
      }
      if(Date.now()-started>=8000){
        clearInterval(timer);
        finished=true;
        window.removeEventListener("fm-google-loaded",onLoaded);
        setAuthStatus("Google services could not be loaded. Check your internet connection.",true);
      }
    },50);
  }

  function initializeGoogle(){
    if(!CONFIG.GOOGLE_CLIENT_ID||CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")){
      setAuthStatus("Add your existing Google Web Client ID to config.js.",true);
      return;
    }
    google.accounts.id.initialize({client_id:CONFIG.GOOGLE_CLIENT_ID,callback:handleCredentialResponse,auto_select:true,cancel_on_tap_outside:false});
    const saved=readSavedSession();
    if(saved){
      setUserProfile(saved);
      $("google-signin-button")?.classList.add("hidden");
      $("grant-access").classList.add("hidden");
      const cached=window.FM_AUTH_CACHE?.read?.(saved.email);
      if(cached?.token){
        state.accessToken=cached.token;
        window.FM_CONNECTION_UI?.show("Restoring your session…","Using your saved Google Sheets connection.");
        hideLogin();
        loadLogger().finally(()=>window.FM_CONNECTION_UI?.hide());
      }else{
        setAuthStatus("Restoring your Sheets connection…");
        window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…","Restoring your saved Google session.");
        attemptSilentAccess(saved.email);
      }
    }else{
      window.FM_CONNECTION_UI?.hide();
      google.accounts.id.renderButton($("google-signin-button"),{theme:"outline",size:"large",text:"signin_with",shape:"rectangular",width:280});
      google.accounts.id.prompt();
    }
  }

  function handleCredentialResponse(response){
    try{
      const payload=decodeJwtPayload(response.credential);
      state.idTokenPayload=payload;
      saveSession();
      setUserProfile(payload);
      $("grant-access").classList.remove("hidden");
      setAuthStatus("Signed in. Connecting to Google Sheets…");
      window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…", "Requesting access for the movement logger.");
      attemptSilentAccess(payload.email);
    }catch(e){
      console.error(e);
      setAuthStatus("Google sign-in response could not be read. Please try again.",true);
    }
  }

  function requestSheetAccess(silent=false){
    if(!state.idTokenPayload && !readSavedSession()){
      setAuthStatus("Sign in with Google first.",true);
      return;
    }
    window.FM_CONNECTION_UI?.show(silent?"Connecting to Google Sheets…":"Authorising Google Sheets & Drive…",silent?"Restoring your saved connection.":"Approve access in the Google prompt to continue.");
    acquireAccessToken(silent?"none":"consent")
      .then(async()=>{hideLogin();await loadLogger();window.FM_CONNECTION_UI?.hide();if(pendingMovement)$('confirm-modal').classList.remove('hidden');})
      .catch(e=>{
        window.FM_CONNECTION_UI?.hide();
        $("grant-access").classList.remove("hidden");
        if(silent){setAuthStatus("Your Google session is available. Allow Sheets & Drive access to continue.");}
        else{setAuthStatus(e?.message||"Google authorization failed. Tap Connect Google Sheets to try again.",true);}
      });
  }
  function attemptSilentAccess(email){
    window.FM_CONNECTION_UI?.show("Connecting to Google Sheets…","Checking your current Google authorisation.");
    acquireAccessToken("none",email)
      .then(async()=>{hideLogin();await loadLogger();window.FM_CONNECTION_UI?.hide();})
      .catch(()=>{
        window.FM_CONNECTION_UI?.hide();
        $("grant-access").classList.remove("hidden");
        setAuthStatus("Your Google session is available. Allow Sheets & Drive access to continue.");
      });
  }

  function acquireAccessToken(prompt="none",email,options={}){
    const forceRefresh=Boolean(options.forceRefresh);
    const expectedEmail=email||state.idTokenPayload?.email||readSavedSession()?.email||"";
    if(!forceRefresh){
      const cached=window.FM_AUTH_CACHE?.read?.(expectedEmail);
      if(cached?.token){state.accessToken=cached.token;return Promise.resolve(cached.token);}
    }
    if(tokenRequestPromise)return tokenRequestPromise;
    tokenRequestPromise=new Promise((resolve,reject)=>{
      let settled=false;
      const finish=(fn,value)=>{if(settled)return;settled=true;tokenRequestPromise=null;clearTimeout(timeoutId);fn(value);};
      const timeoutId=setTimeout(()=>finish(reject,new Error("Google authorisation did not complete.")),prompt==="none"?5000:15000);
      const tokenClient=google.accounts.oauth2.initTokenClient({client_id:CONFIG.GOOGLE_CLIENT_ID,scope:CONFIG.OAUTH_SCOPES,callback:response=>{
        if(response.error){finish(reject,new Error(`Google authorization failed: ${response.error}`));return;}
        state.accessToken=response.access_token;
        window.FM_AUTH_CACHE?.write?.(response.access_token,response.expires_in,expectedEmail);
        finish(resolve,response.access_token);
      }});
      tokenClient.requestAccessToken({prompt,login_hint:expectedEmail||undefined});
    });
    return tokenRequestPromise;
  }

  async function ensureWriteAccess(){
    const email=state.idTokenPayload?.email||readSavedSession()?.email||"";
    const cached=window.FM_AUTH_CACHE?.read?.(email);
    if(cached?.token){state.accessToken=cached.token;return true;}
    try{
      window.FM_CONNECTION_UI?.show("Checking Google access…","Refreshing access before saving this movement.");
      await acquireAccessToken("none",email,{forceRefresh:true});
      window.FM_CONNECTION_UI?.hide();
      return true;
    }catch(e){
      try{
        window.FM_CONNECTION_UI?.show("Google permission needed","Approve the Google access prompt to finish saving this movement.");
        await acquireAccessToken("consent",email,{forceRefresh:true});
        window.FM_CONNECTION_UI?.hide();
        return true;
      }catch(err){window.FM_CONNECTION_UI?.hide();throw err;}
    }
  }

  async function loadLogger(){
    if(!state.accessToken)return false;
    setSyncStatus("Syncing with Google Sheets...");
    try{
      const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;
      state.workbookId=ledgerId;
      state.workbookName=CONFIG.INVENTORY_LEDGER_NAME||"Assets Inventory Ledger";
      $("workbook-name").textContent=state.workbookName;
      const metadata=await sheetsGet(`/${encodeURIComponent(ledgerId)}`);
      const titles=(metadata.sheets||[]).map(s=>s.properties.title);
      const missing=[];
      if(!titles.includes(CONFIG.INVENTORY_SHEET_NAME))missing.push(CONFIG.INVENTORY_SHEET_NAME);
      if(!titles.includes(CONFIG.TRANSACTIONS_SHEET_NAME))missing.push(CONFIG.TRANSACTIONS_SHEET_NAME);
      if(!titles.includes(CONFIG.CLIENT_LIST_SHEET_NAME)){
        throw new Error(`The "${CONFIG.CLIENT_LIST_SHEET_NAME}" sheet is missing from "${CONFIG.INVENTORY_LEDGER_NAME}".`);
      }
      if(!titles.includes(CONFIG.ROUTINE_SHEET_NAME))missing.push(CONFIG.ROUTINE_SHEET_NAME);
      if(!titles.includes(CONFIG.BROKERS_SHEET_NAME))missing.push(CONFIG.BROKERS_SHEET_NAME);
      if(!titles.includes(CONFIG.ALERTS_SHEET_NAME))missing.push(CONFIG.ALERTS_SHEET_NAME);
      if(missing.length)await createSheets(missing);

      const [inventoryRows,transactionRows,clientRows,operationRows,adjustmentRows,brokerRows]=await Promise.all([
        getValues(ledgerId,CONFIG.INVENTORY_SHEET_NAME),
        getValues(ledgerId,CONFIG.TRANSACTIONS_SHEET_NAME),
        getValues(ledgerId,CONFIG.CLIENT_LIST_SHEET_NAME),
        titles.includes(CONFIG.ASSET_OPERATIONS_SHEET_NAME)?getValues(ledgerId,CONFIG.ASSET_OPERATIONS_SHEET_NAME):Promise.resolve([]),
        titles.includes(CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME)?getValues(ledgerId,CONFIG.CLIENT_LEDGER_ADJUSTMENTS_SHEET_NAME):Promise.resolve([]),
        getValues(ledgerId,CONFIG.BROKERS_SHEET_NAME).catch(()=>[])
      ]);

      if(brokerRows.length <= 1){
        try{await updateValues(ledgerId,CONFIG.BROKERS_SHEET_NAME,[["Broker"],...FM_BROKERS.DEFAULT_BROKERS.map(name=>[name])]);}
        catch(error){console.warn("Cannot initialise broker list:",error);}
      }
      state.brokers=FM_BROKERS.parseBrokers(brokerRows);
      const columns=await ensureTransactionColumns(ledgerId,transactionRows);
      state.transactionIdx=columns.idx;
      state.transactionHeaderLength=columns.length;
      state.inventory=parseInventory(inventoryRows);
      // The old Inventory Change Log is retired. Remove that legacy sheet once
      // if it still exists; all new audit information lives on Asset Transactions.
      if(window.FM_SHEET_WRITER?.removeLegacyChangeLog){
        try{
          await FM_SHEET_WRITER.removeLegacyChangeLog({spreadsheetId:ledgerId,accessToken:state.accessToken});
        }catch(legacyError){
          console.warn("Legacy inventory log could not be removed yet:",legacyError);
        }
      }
      state.transactions=[
        ...parseTransactions(transactionRows,columns.idx),
        ...parseTransactions(operationRows),
        ...parseTransactions(adjustmentRows)
      ];
      state.clients=parseClients(clientRows);

      batchForm?.refresh?.();
      renderAudit();
      setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
      return true;
    }catch(e){
      console.error(e);
      setSyncStatus(e.message||"Unable to load spreadsheet.",true);
      return false;
    }
  }

  async function createSheets(names){const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;await sheetsPost(`/${encodeURIComponent(ledgerId)}:batchUpdate`,{requests:names.map(title=>({addSheet:{properties:{title}}}))});if(names.includes(CONFIG.INVENTORY_SHEET_NAME))await updateValues(ledgerId,CONFIG.INVENTORY_SHEET_NAME,[["Asset","Balance"]]);if(names.includes(CONFIG.TRANSACTIONS_SHEET_NAME))await updateValues(ledgerId,CONFIG.TRANSACTIONS_SHEET_NAME,[["Timestamp","Client","Movement","Asset","Quantity","User","Comment","Image Link","Inventory Balance","Broker"]]);if(names.includes(CONFIG.BROKERS_SHEET_NAME))await updateValues(ledgerId,CONFIG.BROKERS_SHEET_NAME,[["Broker"],...FM_BROKERS.DEFAULT_BROKERS.map(name=>[name])]);if(names.includes(CONFIG.ROUTINE_SHEET_NAME))await updateValues(ledgerId,CONFIG.ROUTINE_SHEET_NAME,[["Routine ID","Active","Frequency","Weekday","Direction","Client","Asset","Quantity","Destination","Planned Time","Notes","Created By","Created At"]]);if(names.includes(CONFIG.ALERTS_SHEET_NAME))await updateValues(ledgerId,CONFIG.ALERTS_SHEET_NAME,[["Alert Key","Created At","Alert Type","Status","Client","Asset","Required Qty","Available Qty","Shortfall","Destination","Message","Attended At","Attended By","Resolution Comment","Routine Key"]]);}
  async function getValues(spreadsheetId,sheetName){const data=await sheetsGet(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(quoteSheetName(sheetName)+"!A:AE")}`);return data.values||[];}
  async function updateValues(spreadsheetId,sheetName,rows){const range=`${quoteSheetName(sheetName)}!A1`;return sheetsPut(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:rows});}

  // Reads the Transactions header row, adds "Comment" / "Image Link" columns
  // if they're missing (so existing spreadsheets are migrated automatically),
  // and returns a column-name -> index map plus the resulting header width.
  async function ensureTransactionColumns(ledgerId,rows){
    const header=rows.length?rows[0].slice():["Timestamp","Client","Movement","Asset","Quantity","User"];
    const norm=header.map(normalizeHeader);
    let changed=false;
    if(!norm.some(h=>["comment","comments","notes"].includes(h))){header.push("Comment");norm.push("comment");changed=true;}
    if(!norm.some(h=>["image link","image","photo","picture","photo link","attachment","drive link"].includes(h))){header.push("Image Link");norm.push("image link");changed=true;}
    if(!norm.some(h=>["inventory balance","warehouse balance","balance after movement"].includes(h))){header.push("Inventory Balance");norm.push("inventory balance");changed=true;}
    if(!norm.some(value=>["broker","broker name"].includes(value))){header.push("Broker");norm.push("broker");changed=true;}
    if(changed){
      const range=`${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A1`;
      await sheetsPut(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,{range,majorDimension:"ROWS",values:[header]});
      if(rows.length)rows[0]=header;else rows.push(header);
    }
    const idx={
      timestamp:findColumn(norm,["timestamp","date","datetime"]),
      client:findColumn(norm,["client","client name"]),
      movement:findColumn(norm,["movement","type","direction"]),
      asset:findColumn(norm,["asset","asset name","item"]),
      quantity:findColumn(norm,["quantity","qty"]),
      user:findColumn(norm,["user","entered by","email"]),
      comment:findColumn(norm,["comment","comments","notes"]),
      image:findColumn(norm,["image link","image","photo","picture","photo link","attachment","drive link"]),
      inventoryBalance:findColumn(norm,["inventory balance","warehouse balance","balance after movement"]),
      broker:findColumn(norm,["broker","broker name"])
    };
    return {idx,length:header.length};
  }

  function parseInventory(rows){if(!rows.length)return[];const header=rows[0].map(normalizeHeader);const assetIdx=findColumn(header,["asset","asset name","item","type"]);const balanceIdx=findColumn(header,["balance","current balance","stock","quantity"]);if(assetIdx<0)return[];return rows.slice(1).map((row,index)=>({rowNumber:index+2,asset:String(row?.[assetIdx]??"").trim(),balance:balanceIdx>=0?numericValue(row?.[balanceIdx]):0,assetColumn:assetIdx+1,balanceColumn:balanceIdx>=0?balanceIdx+1:2})).filter(x=>String(x?.asset??"").trim());}
  function parseTransactions(rows,idx){if(!rows.length)return[];
    if(!idx){
      const norm=rows[0].map(normalizeHeader);
      idx={timestamp:findColumn(norm,["timestamp","date","datetime"]),client:findColumn(norm,["client","client name"]),movement:findColumn(norm,["movement","type","direction"]),asset:findColumn(norm,["asset","asset name","item"]),quantity:findColumn(norm,["quantity","qty"]),user:findColumn(norm,["user","entered by","email"]),comment:findColumn(norm,["comment","comments","notes"]),image:findColumn(norm,["image link","image","photo","picture","photo link","attachment","drive link"]),inventoryBalance:findColumn(norm,["inventory balance","warehouse balance","balance after movement"]),broker:findColumn(norm,["broker","broker name"])};
    }
    return rows.slice(1).map(row=>({timestamp:idx.timestamp>=0?(row[idx.timestamp]||row[0]||""):(row[0]||""),client:idx.client>=0?row[idx.client]??"":"",movement:idx.movement>=0?row[idx.movement]??"":"",asset:idx.asset>=0?row[idx.asset]??"":"",quantity:idx.quantity>=0?numericValue(row[idx.quantity]):0,user:idx.user>=0?row[idx.user]??"":"",comment:idx.comment>=0?String(row[idx.comment]??"").trim():"",image:idx.image>=0?String(row[idx.image]??"").trim():"",inventoryBalance:idx.inventoryBalance>=0?numericValue(row[idx.inventoryBalance]):null,broker:idx.broker>=0?String(row[idx.broker]??"").trim():""})).filter(x=>x.asset||x.client);}
  function parseClients(rows){if(!rows.length)return[];const header=rows[0].map(normalizeHeader);const idx=findColumn(header,["client","client name","name"]);if(idx<0)return rows.flat().map(x=>String(x).trim()).filter(Boolean).slice(1);return[...new Set(rows.slice(1).map(r=>String(r[idx]??"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));}

  // ---- Photo attachment (optional) ----
  function handlePhotoChange(){const file=$("movement-photo").files[0];const wrap=$("photo-preview-wrap");const clearBtn=$("clear-photo");if(!file){wrap.classList.add("hidden");clearBtn.classList.add("hidden");return;}const url=URL.createObjectURL(file);$("photo-preview").src=url;wrap.classList.remove("hidden");clearBtn.classList.remove("hidden");}
  function clearPhoto(){$("movement-photo").value="";if($("movement-photo-camera"))$("movement-photo-camera").value="";window._fmCameraFile=null;$("photo-preview-wrap").classList.add("hidden");$("photo-preview").src="";$("clear-photo").classList.add("hidden");}

  // ---- Audit trail: filtering + rendering ----
  function dateKey(timestamp){const d=new Date(timestamp);if(Number.isNaN(d.getTime()))return"";return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
  function todayDateStr(){return dateKey(new Date());}
  function filteredTransactions(){
    const client=$("filter-client")?.value||"";
    const date=$("filter-date")?.value||"";
    const asset=$("filter-asset")?.value||"";
    return state.transactions.filter(t=>(!client||t.client===client)&&(!date||dateKey(t.timestamp)===date)&&(!asset||t.asset===asset));
  }
  function transactionSummaryGroups(transactions){
    const groups=new Map();
    transactions.forEach(item=>{
      const client=String(item.client||"Unknown client").trim()||"Unknown client";
      const asset=String(item.asset||"Unknown asset").trim()||"Unknown asset";
      const movement=String(item.movement||"").trim().toUpperCase();
      const direction=movementLabel(movement)||movement||"Other";
      const key=`${lowerText(client)}\u0000${lowerText(asset)}\u0000${movement}`;
      if(!groups.has(key))groups.set(key,{client,asset,movement,direction,quantity:0});
      groups.get(key).quantity+=Number(item.quantity)||0;
    });
    return[...groups.values()].sort((a,b)=>a.client.localeCompare(b.client)||a.asset.localeCompare(b.asset)||a.direction.localeCompare(b.direction));
  }
  function renderTransactionSummary(container,transactions,label="Matching totals"){
    if(!container)return;
    const groups=transactionSummaryGroups(transactions);
    if(!groups.length){container.innerHTML="";container.classList.add("hidden");return;}
    const rows=groups.map(group=>`<div class="transaction-total-row"><strong class="transaction-total-client">${escapeHtml(group.client)}</strong><span class="transaction-total-asset">${escapeHtml(group.asset)}</span><span class="movement-tag ${movementClass(group.movement)}">${escapeHtml(group.direction)}</span><strong class="transaction-total-quantity">${formatNumber(group.quantity)}</strong></div>`).join("");
    container.innerHTML=`<div class="transaction-summary-heading"><div class="transaction-summary-heading-copy"><div class="eyebrow">${escapeHtml(label)}</div><strong>${formatNumber(transactions.length)} matching transaction${transactions.length===1?"":"s"}</strong></div><div class="transaction-summary-heading-actions"><span class="muted transaction-summary-hint">Grouped by client → asset type → direction</span><button type="button" class="secondary transaction-summary-toggle" data-summary-toggle aria-expanded="false">Expand totals</button></div></div><div class="transaction-summary-body hidden" data-summary-body><div class="transaction-summary-table"><div class="transaction-summary-head"><span>Client</span><span>Asset type</span><span>Direction</span><span class="num">Total</span></div>${rows}</div></div>`;
    container.classList.remove("hidden");
  }
  function toggleTransactionSummary(button){
    const container=button.closest(".transaction-summary");
    const body=container?.querySelector("[data-summary-body]");
    if(!container||!body)return;
    const expanded=button.getAttribute("aria-expanded")==="true";
    body.classList.toggle("hidden",expanded);
    button.setAttribute("aria-expanded",String(!expanded));
    button.textContent=expanded?"Expand totals":"Hide totals";
    container.classList.toggle("is-expanded",!expanded);
  }
  function renderTransactionCard(group,index){
    const sections=FM_TRANSACTION_DISPLAY.movementSections(group);
    const sectionMarkup=sections.map(section=>{
      const movement=String(section.movement||"").trim().toUpperCase();
      const cls=movement==="RECEIVED"||movement==="COMBINE_IN"?"received":movement==="SENT"?"sent":movement==="DISCARD"||movement==="COMBINE_OUT"?"discard":"other";
      return `<section class="transaction-movement-section movement-${cls}">
        <div class="transaction-movement-heading"><span class="movement-tag ${cls}">${escapeHtml(section.label)}</span><span>${formatNumber(section.totalQuantity)} total</span></div>
        <div class="transaction-asset-list">${section.items.map(item=>`<div class="transaction-asset-pill"><span>${escapeHtml(item.asset)}</span><strong>${formatNumber(item.quantity)}</strong></div>`).join("")}</div>
      </section>`;
    }).join("");
    const labels=sections.map(section=>section.label).join(" + ");
    return `<article class="transaction-card ${FM_TRANSACTION_DISPLAY.cardClass(group)}">
      <div class="transaction-card-accent" aria-hidden="true"></div>
      <div class="transaction-card-top">
        <div class="transaction-card-icon-wrap"><img class="transaction-card-icon" src="truck-icon.png" alt="" aria-hidden="true"></div>
        <div class="transaction-card-heading"><div class="transaction-card-client">${escapeHtml(group.client||"Unknown client")}</div><div class="transaction-card-meta"><span>${escapeHtml(formatTimestamp(group.timestamp))}</span><span>•</span><span>${escapeHtml(labels||"Movement")}</span>${group.broker?`<span class="broker-transaction-tag">Broker: ${escapeHtml(group.broker)}</span>`:""}${group.vehicleReg?`<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>`:""}</div></div>
        <button type="button" class="secondary info-button transaction-card-view" data-info-index="${index}" aria-label="View details for ${escapeAttr(group.client||"movement")}">View details</button>
      </div>
      <div class="transaction-card-body">${sectionMarkup}</div>
      <div class="transaction-card-footer"><span>${formatNumber(group.totalLines)} ledger line${group.totalLines===1?"":"s"}${group.image?" · attachment":""}</span><strong>${formatNumber(group.totalQuantity)} total units</strong></div>
    </article>`;
  }
  function renderAudit(){
    const client=$("filter-client")?.value||"";const date=$("filter-date")?.value||"";const asset=$("filter-asset")?.value||"";
    const filtering=Boolean(client||date||asset);
    const matches=filtering?filteredTransactions():state.transactions;
    const grouped=FM_TRANSACTION_DISPLAY.groupTransactions(matches);
    const groups=(filtering?grouped.slice(0,200):grouped.slice(0,30));
    state.auditVisibleRows=groups;
    const count=$("audit-count");
    if(count)count.textContent=filtering?`${matches.length.toLocaleString()} lines · ${groups.length.toLocaleString()} grouped movements`:groups.length.toLocaleString();
    const body=$("transactions-body");
    if(!body)return;
    body.innerHTML=groups.length?groups.map((group,i)=>renderTransactionCard(group,i)).join(""):`<div class="transaction-empty">${filtering?"No asset movements match these filters.":"No asset movements recorded yet."}</div>`;
    renderTransactionSummary($("transaction-summary"),matches,filtering?"Filtered totals":"All transaction totals");
  }

  // ---- More info modal ----
  async function openInfoModal(index){
    const group=state.auditVisibleRows[index];if(!group)return;
    const title=$("info-modal-title");
    if(title)title.textContent=`${group.client||"Movement"} · Details`;
    const summary=$("info-summary");
    const comments=Array.isArray(group.comments)&&group.comments.length?group.comments:(group.comment?[group.comment]:[]);
    if(summary){
      const sections=FM_TRANSACTION_DISPLAY.movementSections(group);
      summary.innerHTML=`<div class="transaction-detail-note"><strong>${escapeHtml(group.client||"Unknown client")}</strong>${group.broker?`<span> · Broker: ${escapeHtml(group.broker)}</span>`:""}<span> · ${escapeHtml(formatTimestamp(group.timestamp)||"Unknown time")}</span><span> · ${escapeHtml(FM_TRANSACTION_DISPLAY.movementLabels(group).join(" + ")||"Movement")}</span>${group.vehicleReg?`<span class="transaction-vehicle-reg"><span>Vehicle reg</span><strong>${escapeHtml(group.vehicleReg)}</strong></span>`:""}</div>${sections.map(section=>`<div class="transaction-detail-section"><div class="transaction-detail-section-head"><span class="movement-tag ${lowerText(section.movement)}">${escapeHtml(section.label)}</span><strong>${formatNumber(section.totalQuantity)}</strong></div><div class="transaction-detail-assets">${section.items.map(item=>`<span>${escapeHtml(item.asset)} <strong>${formatNumber(item.quantity)}</strong></span>`).join("")}</div></div>`).join("")}`;
    }
    const photoWrap=$("info-photo-wrap");const commentEl=$("info-comment");const emptyEl=$("info-empty");
    if(group.image){
      const image=$("info-photo");FM_MEDIA?.revokeObjectUrl?.(image);photoWrap.classList.remove("hidden");image.removeAttribute("src");image.classList.add("is-loading");$("info-photo-link").href=group.image;const photoStatus=$("info-photo-status");photoStatus?.classList.remove("hidden");if(photoStatus)photoStatus.textContent="Loading photo...";let loaded=false;
      try{loaded=await FM_MEDIA.loadDriveImage(image,group.image,state.accessToken);if(!loaded&&photoStatus)photoStatus.textContent="Preview unavailable here. Use Open full size in Drive.";}finally{image.classList.remove("is-loading");if(photoStatus&&loaded)photoStatus.classList.add("hidden");}
    }else{
      $("info-photo").removeAttribute("src");FM_MEDIA?.revokeObjectUrl?.($("info-photo"));$("info-photo-status")?.classList.add("hidden");$("info-photo-link").removeAttribute("href");photoWrap.classList.add("hidden");
    }
    if(comments.length){commentEl.textContent=comments.join("\n\n");commentEl.classList.remove("hidden");}else{commentEl.textContent="";commentEl.classList.add("hidden");}
    emptyEl.classList.toggle("hidden",Boolean(group.image||comments.length));
    $("info-modal").classList.remove("hidden");
  }
  function closeInfoModal(){$("info-modal").classList.add("hidden");}
  function driveFileIdFromLink(url){const m=String(url||"").match(/\/d\/([a-zA-Z0-9_-]+)/)||String(url||"").match(/[?&]id=([a-zA-Z0-9_-]+)/);return m?m[1]:"";}

  function handleMovementChange(){
    batchForm?.refresh?.();
    updatePreview();
  }
  function reviewMovement(event){
    event.preventDefault();
    const data=readForm();
    const error=validateMovement(data);
    if(error){setMovementStatus(error,true);return;}
    const balances=new Map(state.inventory.map(i=>[lowerText(i?.asset),i?.balance]).filter(([key])=>key));
    const prepared=[];
    for(const entry of data.entries){
      const movement=entry.movement;
      const client=movement==="DISCARD"?SELF_CLIENT:entry.client;
      for(const itemEntry of entry.items){
        const wantedAssetKey=lowerText(itemEntry?.asset);
        const item=state.inventory.find(x=>lowerText(x?.asset)===wantedAssetKey);
        if(!item){setMovementStatus(`Asset "${itemEntry.asset}" is not present in Inventory.`,true);return;}
        const key=lowerText(item?.asset);
        const current=balances.get(key)??0;
        const next=movement==="RECEIVED"?current+itemEntry.quantity:current-itemEntry.quantity;
        if(next<0){setMovementStatus(`Cannot remove ${formatNumber(itemEntry.quantity)} ${item.asset}. Current balance is ${formatNumber(current)}.`,true);return;}
        balances.set(key,next);
        prepared.push({client,broker:movement==="DISCARD"?"":(entry.broker||""),movement,asset:item.asset,quantity:itemEntry.quantity});
      }
    }
    pendingMovement={...data,prepared,balances};
    const movements=[...new Set(data.entries.map(e=>e.movement).filter(Boolean))];
    const clients=[...new Set(data.entries.map(e=>e.movement==="DISCARD"?SELF_CLIENT:e.client).filter(Boolean))];
    $("confirm-movement").textContent=movements.length===1?movementLabel(movements[0]):"Batch movements";
    $("confirm-client").textContent=clients.length===1?clients[0]:`${clients.length} clients`;
    $("confirm-time").textContent=data.timestamp;
    $("confirm-items").innerHTML=data.entries.map(entry=>{
      const client=entry.movement==="DISCARD"?SELF_CLIENT:entry.client;
      const movement=movementLabel(entry.movement);
      const total=entry.items.reduce((sum,item)=>sum+item.quantity,0);
      return `<div class="confirm-batch-client"><div class="confirm-batch-heading"><strong>${escapeHtml(client)}</strong>${entry.broker?`<span class="broker-transaction-tag">Broker: ${escapeHtml(entry.broker)}</span>`:""}<span class="movement-tag ${movementClass(entry.movement)}">${escapeHtml(movement)}</span><strong>${formatNumber(total)}</strong></div><div class="confirm-batch-assets">${entry.items.map(item=>`<span>${escapeHtml(item.asset)} <strong>${formatNumber(item.quantity)}</strong></span>`).join("")}</div></div>`;
    }).join("");
    $("confirm-total").textContent=formatNumber(prepared.reduce((sum,item)=>sum+item.quantity,0));
    const attachment=$("confirm-attachment");const photoWrap=$("confirm-photo-wrap");const commentEl=$("confirm-comment");let showAttachment=false;
    if(data.photoFile){$("confirm-photo").src=URL.createObjectURL(data.photoFile);photoWrap.classList.remove("hidden");showAttachment=true;}else photoWrap.classList.add("hidden");
    if(data.comment){commentEl.textContent=data.comment;commentEl.classList.remove("hidden");showAttachment=true;}else commentEl.classList.add("hidden");
    attachment.classList.toggle("hidden",!showAttachment);
    const warning=$("confirm-warning");const hasDiscard=data.entries.some(e=>e.movement==="DISCARD");warning.classList.toggle("hidden",!hasDiscard);if(hasDiscard)warning.textContent="This batch includes discarded assets. Discarding permanently removes the selected quantities from inventory.";
    $("confirm-modal").classList.remove("hidden");setTimeout(()=>$("approve-confirm").focus(),50);
  }
  async function approveMovement(){
    if(!pendingMovement||movementSubmitBusy)return;
    const data=pendingMovement;
    movementSubmitBusy=true;
    $("approve-confirm").disabled=true;$("cancel-confirm").disabled=true;
    setMovementSubmitting(true,"Checking Google access...");
    try{await ensureWriteAccess();}
    catch(authError){
      movementSubmitBusy=false;setMovementSubmitting(false);$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;showReconnectUI("Google access expired. Allow Sheets & Drive access, then approve the movement again.");setMovementStatus(authError.message||"Google access could not be refreshed.",true);return;
    }
    setMovementSubmitting(true,data.photoFile?"Uploading photo...":"Checking live inventory...");
    setMovementStatus("Recording movements...");
    try{
      const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const user=state.idTokenPayload?.email||readSavedSession()?.email||"Google user";
      const timestamp=transactionTimestampFromTime(data.timestamp);
      let imageLink="";
      if(data.photoFile){
        setMovementStatus("Uploading photo...");
        const folderId=await getOrCreateDailyFolder(todayDateStr());
        const filename=buildPhotoFilename({...data,timestamp});
        const uploadFile=await FM_MEDIA.optimizeImageForUpload(data.photoFile);
        const uploaded=await uploadImageToDrive(uploadFile,folderId,filename);
        imageLink=uploaded.webViewLink||`https://drive.google.com/file/d/${uploaded.id}/view`;
        setMovementSubmitting(true,"Checking live inventory...");setMovementStatus("Recording movements...");
      }

      // The shared writer re-reads the live inventory immediately before its
      // single Sheets batch update. It records the resulting warehouse balance
      // directly on each Asset Transactions row.
      if(!window.FM_SHEET_WRITER?.appendMovementBatch)throw new Error("The sheet writer is unavailable. Refresh the page and try again.");
      const result=await FM_SHEET_WRITER.appendMovementBatch({
        spreadsheetId:ledgerId,
        accessToken:state.accessToken,
        movements:data.prepared.map(movement=>({
          timestamp,
          client:movement.client,
          broker:movement.broker||"",
          movement:movement.movement,
          asset:movement.asset,
          quantity:movement.quantity,
          user,
          comment:data.comment||"",
          image:imageLink
        })),
        user,
        timestamp
      });
      if(!result?.movements?.length)throw new Error("Google Sheets did not confirm the movement write.");

      const livePrepared=result.movements;

      const clientCount=new Set(livePrepared.map(item=>item.client)).size;
      const movementCount=livePrepared.length;
      closeConfirm();resetMovementForm();setMovementStatus(`${movementCount} movement${movementCount===1?"":"s"} recorded for ${clientCount} client${clientCount===1?"":"s"}.`);await loadLogger();
    }catch(e){console.error(e);setMovementStatus(e.message||"Unable to record movement.",true);}
    finally{movementSubmitBusy=false;setMovementSubmitting(false);$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
  }

  function buildPhotoFilename(data={}){
    const records=Array.isArray(data.prepared)&&data.prepared.length?data.prepared:(Array.isArray(data.entries)?data.entries:[]);
    const clients=[...new Set(records.map(record=>{
      const movement=String(record?.movement||"").toUpperCase();
      return String(movement==="DISCARD"?SELF_CLIENT:(record?.client||"")).trim();
    }).filter(Boolean))];
    const fallbackClient=String(data.client||"").trim();
    if(!clients.length&&fallbackClient&&!(["batch","unknown client","client"].includes(fallbackClient.toLowerCase())))clients.push(fallbackClient);
    const clientLabel=clients.length?(clients.length<=3?clients.join("-AND-"):`${clients[0]}-AND-${clients.length-1}-OTHERS`):"UNKNOWN-CLIENT";
    const safeClient=clientLabel.toUpperCase().replace(/[^A-Z0-9]+/g,"-").replace(/^-+|-+$/g,"").slice(0,90)||"UNKNOWN-CLIENT";
    const movements=[...new Set(records.map(record=>String(record?.movement||"").trim().toUpperCase()).filter(Boolean))];
    const suppliedMovement=String(data.movement||"").trim().toUpperCase();
    if(!movements.length&&suppliedMovement&&suppliedMovement!=="MIXED")movements.push(suppliedMovement);
    const movementPart=movements.length===1?(movements[0]==="DISCARD"?"DISCARDED":movements[0]):"MIXED";
    const selectedDate=data.timestamp?new Date(data.timestamp):null;
    const fileDate=selectedDate&&!Number.isNaN(selectedDate.getTime())?selectedDate:new Date();
    const stamp=fileDate.toISOString().replace(/[:.]/g,"-");
    const uniqueSuffix=Date.now().toString(36).toUpperCase();
    const originalName=String(data.photoFile?.name||"");
    const ext=data.photoFile?.type==="image/jpeg"?".jpg":((originalName.match(/\.[a-zA-Z0-9]+$/)||[".jpg"])[0]);
    return `${safeClient}-${movementPart}-${stamp}-${uniqueSuffix}${ext}`;
  }

  // ---- Drive: daily dated folder + photo upload ----
  async function getOrCreateDailyFolder(dateStr){
    if(state.driveFolderCache[dateStr])return state.driveFolderCache[dateStr];
    const parent=CONFIG.DRIVE_PHOTOS_PARENT_FOLDER_ID;
    const query=[`name = '${escapeDriveQuery(dateStr)}'`,`mimeType = 'application/vnd.google-apps.folder'`,`'${parent}' in parents`,`trashed = false`].join(" and ");
    const found=await fetchJson(`${DRIVE_API}?q=${encodeURIComponent(query)}&pageSize=1&fields=files(id,name)`,{headers:authHeaders()});
    let id=found.files?.[0]?.id;
    if(!id){const created=await fetchJson(DRIVE_API,{method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify({name:dateStr,mimeType:"application/vnd.google-apps.folder",parents:[parent]})});id=created.id;}
    state.driveFolderCache[dateStr]=id;return id;
  }
  async function uploadImageToDrive(file,folderId,filename){
    const metadata={name:filename,parents:[folderId],mimeType:file.type||"image/jpeg"};
    const boundary="fmlogger"+Math.random().toString(36).slice(2);
    const delimiter=`--${boundary}\r\n`;const closeDelim=`\r\n--${boundary}--`;
    const metaPart=delimiter+"Content-Type: application/json; charset=UTF-8\r\n\r\n"+JSON.stringify(metadata)+"\r\n";
    const mediaHeader=delimiter+`Content-Type: ${metadata.mimeType}\r\n\r\n`;
    const body=new Blob([metaPart,mediaHeader,file,closeDelim]);
    return fetchJson(`${DRIVE_UPLOAD_API}?uploadType=multipart&fields=id,webViewLink`,{method:"POST",headers:{...authHeaders(),"Content-Type":`multipart/related; boundary=${boundary}`},body});
  }

  let movementSubmitLoadingTimer=null;
  function setMovementSubmitting(busy,detail=""){
    const overlay=$("movement-submit-loading");
    if(movementSubmitLoadingTimer){clearTimeout(movementSubmitLoadingTimer);movementSubmitLoadingTimer=null;}
    if(overlay&&busy){
      movementSubmitLoadingTimer=setTimeout(()=>{
        if(!movementSubmitBusy)return;
        overlay.classList.remove("hidden");overlay.setAttribute("aria-hidden","false");
      },250);
    }else if(overlay){
      overlay.classList.add("hidden");overlay.setAttribute("aria-hidden","true");
    }
    if($("movement-submit-loading-detail")&&detail)$("movement-submit-loading-detail").textContent=detail;
    const submit=$("movement-form")?.querySelector("button[type=submit]");if(submit){if(!submit.dataset.originalText)submit.dataset.originalText=submit.textContent;submit.disabled=busy;submit.textContent=busy?"Saving...":submit.dataset.originalText;}
    batchForm?.setAllDisabled?.(busy);
    $("movement-photo")?.toggleAttribute("disabled",busy);$("movement-photo-camera")?.toggleAttribute("disabled",busy);$("clear-photo")?.toggleAttribute("disabled",busy);$("movement-comment")?.toggleAttribute("disabled",busy);
  }

  function closeConfirm(){$("confirm-modal").classList.add("hidden");pendingMovement=null;$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
  function readForm(){return{entries:batchForm?.read?.()||[],timestamp:$("timestamp").value,comment:$("movement-comment").value.trim(),photoFile:$("movement-photo").files[0]||window._fmCameraFile||null};}
  function validateMovement(d){
    if(!d.timestamp)return"Please choose a time.";
    if(!d.entries.length)return"Add at least one client movement.";
    const seenClients=new Set();
    for(const entry of d.entries){
      if(!entry.movement)return"Select a movement for every client entry.";
      const client=entry.movement==="DISCARD"?SELF_CLIENT:entry.client;
      if(!client)return"Select a client for every movement.";
      if(!entry.items.length)return`Add at least one asset type for ${client}.`;
      const clientKey=lowerText(client);
      if(seenClients.has(clientKey))return`You have added ${client} more than once. Combine that client's movements into one card.`;
      seenClients.add(clientKey);
      const seenAssets=new Set();
      for(const item of entry.items){
        if(!item.asset||!Number.isInteger(item.quantity)||item.quantity<=0)return`Select an asset type and enter a whole quantity greater than zero for every row under ${client}.`;
        const key=lowerText(item?.asset);
        if(seenAssets.has(key))return`You have selected ${item.asset} more than once for ${client}. Combine the quantities into one row.`;
        seenAssets.add(key);
      }
    }
    return null;
  }
  function updatePreview(){batchForm?.updatePreview?.("preview-text");}
  function resetMovementForm(){
    $("movement-form")?.reset();
    batchForm?.reset?.();
    setDefaultTimestamp();
    clearPhoto();
    window._fmCameraFile=null;
    handleMovementChange();
  }
  function transactionTimestampFromTime(timeValue){
    const value=String(timeValue||"").trim();
    if(!/^\d{2}:\d{2}$/.test(value))return new Date().toISOString();
    const [hours,minutes]=value.split(":").map(Number);
    const now=new Date();
    const selected=new Date(now.getFullYear(),now.getMonth(),now.getDate(),hours,minutes,0,0);
    const offset=-selected.getTimezoneOffset();
    const sign=offset>=0?"+":"-";
    const abs=Math.abs(offset);
    const offsetHours=String(Math.floor(abs/60)).padStart(2,"0");
    const offsetMinutes=String(abs%60).padStart(2,"0");
    const pad=v=>String(v).padStart(2,"0");
    return `${selected.getFullYear()}-${pad(selected.getMonth()+1)}-${pad(selected.getDate())}T${pad(hours)}:${pad(minutes)}:00${sign}${offsetHours}:${offsetMinutes}`;
  }
  function setDefaultTimestamp(){const input=$("timestamp");if(!input)return;const now=new Date();input.value=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;}
  function movementLabel(x){return x==="RECEIVED"?"Received":x==="SENT"?"Sent":x==="DISCARD"?"Discard":x==="COMBINE_OUT"?"Combination component used":x==="COMBINE_IN"?"Combined asset created":x==="NORMALISATION"?"Client ledger normalisation":String(x||"");}function movementClass(x){const m=String(x||"").trim().toUpperCase();return m==="RECEIVED"||m==="COMBINE_IN"?"received":m==="SENT"?"sent":m==="DISCARD"||m==="COMBINE_OUT"?"discard":"other";}function formatTimestamp(x){if(!x)return"";const d=new Date(x);return Number.isNaN(d.getTime())?String(x):d.toLocaleString("en-GB",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"});}function formatNumber(x){return Number(x||0).toLocaleString("en-GB");}function numericValue(x){if(x===null||x===undefined||x==="")return 0;const n=Number(String(x).replace(/,/g,""));return Number.isFinite(n)?n:0;}function normalizeHeader(x){return String(x??"").trim().toLowerCase().replace(/\s+/g," ");}function findColumn(headers,names){for(const name of names){const idx=headers.indexOf(name);if(idx>=0)return idx;}return-1;}function quoteSheetName(x){return `'${String(x).replace(/'/g,"''")}'`;}function columnLetter(n){let r="";while(n>0){const rem=(n-1)%26;r=String.fromCharCode(65+rem)+r;n=Math.floor((n-1)/26);}return r;}function escapeDriveQuery(x){return String(x).replace(/\\/g,"\\\\").replace(/'/g,"\\'");}function escapeHtml(x){return String(x??"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}function escapeAttr(x){return escapeHtml(x);}function emptyRow(colspan,text){return `<tr><td colspan="${colspan}" class="empty">${escapeHtml(text)}</td></tr>`;}
  function authHeaders(){return{Authorization:`Bearer ${state.accessToken}`};}
  async function sheetsGet(path){return fetchJson(SHEETS_API+path,{headers:authHeaders()});}async function sheetsPost(path,body){return fetchJson(SHEETS_API+path,{method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}async function sheetsPut(path,body){return fetchJson(SHEETS_API+path,{method:"PUT",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}
  function isTransientApiStatus(status){return [408,429,500,502,503,504].includes(Number(status));}
  async function fetchJson(url,options={}){
    const method=String(options.method||"GET").toUpperCase();
    const safeRetry=["GET","HEAD"].includes(method);
    const maxRetries=safeRetry?3:0;
    let response;
    for(let attempt=0;attempt<=maxRetries;attempt+=1){
      response=await fetch(url,options);
      if(response.status===401&&!options.__retried){
        try{
          await acquireAccessToken("none",state.idTokenPayload?.email||readSavedSession()?.email,{forceRefresh:true});
          return fetchJson(url,{...options,__retried:true,headers:{...(options.headers||{}),...authHeaders()}});
        }catch(_){
          state.accessToken=null;window.FM_AUTH_CACHE?.clear?.();showReconnectUI("Google access expired. Allow Sheets & Drive access to reconnect.");
          throw new Error("Google access expired. Allow Sheets & Drive access to reconnect.");
        }
      }
      if(response.ok)break;
      if(!safeRetry||!isTransientApiStatus(response.status)||attempt>=maxRetries)break;
      const retryAfter=Number(response.headers?.get?.("Retry-After"));
      const delay=Number.isFinite(retryAfter)&&retryAfter>0?Math.min(32000,retryAfter*1000):Math.min(8000,1000*(2**attempt))+Math.floor(Math.random()*250);
      await new Promise(resolve=>setTimeout(resolve,delay));
    }
    const text=await response.text();let data={};try{data=text?JSON.parse(text):{}}catch(_){}
    if(!response.ok){const error=new Error(data?.error?.message||`Request failed (${response.status})`);error.status=response.status;throw error;}
    return data;
  }
  function decodeJwtPayload(jwt){const parts=String(jwt).split(".");if(parts.length!==3)throw new Error("Invalid Google credential.");const base64=parts[1].replace(/-/g,"+").replace(/_/g,"/");const padded=base64+"=".repeat((4-base64.length%4)%4);return JSON.parse(decodeURIComponent(Array.from(atob(padded)).map(c=>`%${c.charCodeAt(0).toString(16).padStart(2,"0")}`).join("")));}
  function saveSession(){if(!state.idTokenPayload)return;localStorage.setItem(SESSION_KEY,JSON.stringify({name:state.idTokenPayload.name||"Google user",email:state.idTokenPayload.email||"",picture:state.idTokenPayload.picture||"",sub:state.idTokenPayload.sub||""}));}
  function readSavedSession(){try{return JSON.parse(localStorage.getItem(SESSION_KEY)||"null");}catch(_){return null;}}
  function setUserProfile(profile){
    const rawName=String(profile?.given_name||profile?.name||"Google user").trim();
    const firstName=(rawName.split(/\s+/)[0]||"there").replace(/[^\p{L}\p{M}'-]/gu, "");
    $("user-name").textContent=`Hi, ${firstName||"there"} 👋`;
    $("user-email").textContent="";
    if(profile?.picture){
      $("user-photo").src=profile.picture;
      $("user-photo").classList.remove("hidden");
    }
  }

  function setAuthStatus(text,error=false){
    $("auth-status").textContent=text||"";
    $("auth-status").className=`status ${error?"error":""}`;
  }

  function setMovementStatus(text,error=false){$("movement-status").textContent=text;$("movement-status").className=`status ${error?"error":""}`;}
  function setSyncStatus(text,error=false){$("sync-status").textContent=text;$("sync-status").className=`muted ${error?"error":""}`;}

  function hideLogin(){
    $("google-signin-button").classList.add("hidden");
    $("grant-access").classList.add("hidden");
    $("sign-out").classList.remove("hidden");
    $("login-card").classList.add("hidden");
    $("logger").classList.remove("hidden");
  }

  function showReconnectUI(message="Google access expired. Allow Sheets & Drive access to continue."){
    window.FM_CONNECTION_UI?.hide();
    $("confirm-modal")?.classList.add("hidden");
    $("login-card")?.classList.remove("hidden");
    $("google-signin-button")?.classList.add("hidden");
    $("grant-access")?.classList.remove("hidden");
    const title=$("login-card")?.querySelector("h2");
    if(title)title.textContent="Reconnect Google Sheets";
    setAuthStatus(message,true);
  }

  function signOut(){
    const saved=readSavedSession();
    if(state.idTokenPayload?.sub){try{google.accounts.id.revoke(state.idTokenPayload.sub,()=>{});}catch(_){}}
    if(saved?.sub&&saved.sub!==state.idTokenPayload?.sub){try{google.accounts.id.revoke(saved.sub,()=>{});}catch(_){}}
    localStorage.removeItem(SESSION_KEY);
    window.FM_AUTH_CACHE?.clear?.();
    state.idTokenPayload=null;
    state.accessToken=null;
    state.workbookId=null;
    state.workbookName=null;
    state.inventory=[];
    state.transactions=[];
    state.clients=[];
    $("logger").classList.add("hidden");
    $("login-card").classList.remove("hidden");
    $("google-signin-button").classList.remove("hidden");
    $("grant-access").classList.add("hidden");
    $("sign-out").classList.add("hidden");
    $("user-photo").classList.add("hidden");
    $("user-name").textContent="Not signed in";
    $("user-email").textContent="";
    setAuthStatus("");
  }

  function refreshSavedSessionIfNeeded(){
    const saved=readSavedSession();
    if(!saved||!window.google?.accounts?.oauth2)return;
    if(window.FM_AUTH_CACHE?.read?.(saved.email)?.token)return;
    attemptSilentAccess(saved.email);
  }
  window.addEventListener("pageshow",refreshSavedSessionIfNeeded);
  document.addEventListener("visibilitychange",()=>{if(!document.hidden)refreshSavedSessionIfNeeded();});

  // Keep this helper available to any older dashboard/logger code that calls it globally.
  // Backward-compatible bridge for any cached/older page code that still calls renderInputs().
  window.renderInputs = () => batchForm?.refresh?.();
  window.formatTimestamp = formatTimestamp;
})();
