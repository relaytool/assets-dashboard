(() => {
  "use strict";

  const SHEETS_API="https://sheets.googleapis.com/v4/spreadsheets";
  const DRIVE_API="https://www.googleapis.com/drive/v3/files";
  const DRIVE_UPLOAD_API="https://www.googleapis.com/upload/drive/v3/files";
  const SELF_CLIENT=CONFIG.SELF_CLIENT_NAME||"HSC London (Self)";
  const SESSION_KEY="fmAssetSession";
  const state={idTokenPayload:null,accessToken:null,workbookId:null,workbookName:null,inventory:[],transactions:[],clients:[],transactionIdx:null,transactionHeaderLength:6,driveFolderCache:{},auditVisibleRows:[]};
  let pendingMovement=null;
  let tokenRequestPromise=null;
  const $=id=>document.getElementById(id);

  document.addEventListener("DOMContentLoaded",()=>{bindEvents();setDefaultTimestamp();waitForGoogle();});

  function bindEvents(){
    $("grant-access").addEventListener("click",()=>requestSheetAccess(false));
    $("sign-out").addEventListener("click",signOut);
    $("refresh").addEventListener("click",loadLogger);
    $("movement-form").addEventListener("submit",reviewMovement);
    $("movement").addEventListener("change",handleMovementChange);
    $("asset-rows").addEventListener("input",updatePreview);
    $("asset-rows").addEventListener("change",updatePreview);
    $("add-asset-row").addEventListener("click",addAssetRow);
    $("cancel-confirm").addEventListener("click",closeConfirm);
    $("approve-confirm").addEventListener("click",approveMovement);
    $("confirm-items").addEventListener("click",e=>{const b=e.target.closest("[data-remove-row]");if(b)removeAssetRow(Number(b.dataset.removeRow));});
    document.addEventListener("click",e=>{if(e.target.matches("[data-close-confirm]"))closeConfirm();if(e.target.matches("[data-close-info]"))closeInfoModal();});
    $("movement-photo").addEventListener("change",handlePhotoChange);
    $("clear-photo").addEventListener("click",clearPhoto);
    $("filter-client").addEventListener("change",()=>renderAudit());
    $("filter-date").addEventListener("change",()=>renderAudit());
    $("clear-filters").addEventListener("click",()=>{$("filter-client").value="";$("filter-date").value="";renderAudit();});
    $("transactions-body").addEventListener("click",e=>{const b=e.target.closest("[data-info-index]");if(b)openInfoModal(Number(b.dataset.infoIndex));});
    $("close-info").addEventListener("click",closeInfoModal);
  }

  function waitForGoogle(){
    let checks=0;
    const timer=setInterval(()=>{
      checks++;
      if(window.google?.accounts?.id&&window.google?.accounts?.oauth2){
        clearInterval(timer);
        initializeGoogle();
      }
      if(checks>=150){
        clearInterval(timer);
        setAuthStatus("Google services could not be loaded. Check your internet connection.",true);
      }
    },100);
  }

  function initializeGoogle(){
    if(!CONFIG.GOOGLE_CLIENT_ID||CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")){
      setAuthStatus("Add your existing Google Web Client ID to config.js.",true);
      return;
    }
    google.accounts.id.initialize({client_id:CONFIG.GOOGLE_CLIENT_ID,callback:handleCredentialResponse,auto_select:true,cancel_on_tap_outside:false});
    google.accounts.id.renderButton($("google-signin-button"),{theme:"outline",size:"large",text:"signin_with",shape:"rectangular",width:280});
    const saved=readSavedSession();
    if(saved){setUserProfile(saved);attemptSilentAccess(saved.email);}
    google.accounts.id.prompt();
  }

  function handleCredentialResponse(response){
    try{
      const payload=decodeJwtPayload(response.credential);
      state.idTokenPayload=payload;
      saveSession();
      setUserProfile(payload);
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
    acquireAccessToken(silent?"none":"consent")
      .then(async()=>{hideLogin();await loadLogger();})
      .catch(e=>{
        $("grant-access").classList.remove("hidden");
        if(silent){
          setAuthStatus("Google account restored. Connect Google Sheets to continue.");
        }else{
          setAuthStatus(e?.message||"Google authorization failed. Click Connect Google Sheets to try again.",true);
        }
      });
  }

  function attemptSilentAccess(email){
    acquireAccessToken("none",email)
      .then(async()=>{hideLogin();await loadLogger();})
      .catch(()=>{
        $("grant-access").classList.remove("hidden");
        setAuthStatus("Google account restored. Allow Sheets & Drive access to continue.");
      });
  }

  function acquireAccessToken(prompt="none",email){
    if(tokenRequestPromise)return tokenRequestPromise;
    tokenRequestPromise=new Promise((resolve,reject)=>{
      let settled=false;
      const finish=(fn,value)=>{if(settled)return;settled=true;tokenRequestPromise=null;clearTimeout(timeoutId);fn(value);};
      const timeoutId=setTimeout(()=>finish(reject,new Error("Google's sign-in popup didn't open (it may have been blocked). Please allow popups for this site and try again.")),20000);
      const tokenClient=google.accounts.oauth2.initTokenClient({client_id:CONFIG.GOOGLE_CLIENT_ID,scope:CONFIG.OAUTH_SCOPES,callback:response=>{if(response.error){finish(reject,new Error(`Google authorization failed: ${response.error}`));return;}state.accessToken=response.access_token;finish(resolve,response.access_token);}});
      tokenClient.requestAccessToken({prompt,login_hint:email||state.idTokenPayload?.email||readSavedSession()?.email||undefined});
    });
    return tokenRequestPromise;
  }

  async function loadLogger(){
    if(!state.accessToken)return false;
    setSyncStatus("Syncing with Google Sheets...");
    try{
      const workbook=await findMainWorkbook();
      if(!workbook){
        throw new Error(`Workbook "${CONFIG.PRIMARY_WORKBOOK_NAME}" (or "${CONFIG.FALLBACK_WORKBOOK_NAME}") was not found in your Google Drive.`);
      }

      state.workbookId=workbook.id;
      state.workbookName=workbook.name;
      $("workbook-name").textContent=workbook.name;

      const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const metadata=await sheetsGet(`/${encodeURIComponent(ledgerId)}`);
      const titles=(metadata.sheets||[]).map(s=>s.properties.title);
      const missing=[];
      if(!titles.includes(CONFIG.INVENTORY_SHEET_NAME))missing.push(CONFIG.INVENTORY_SHEET_NAME);
      if(!titles.includes(CONFIG.TRANSACTIONS_SHEET_NAME))missing.push(CONFIG.TRANSACTIONS_SHEET_NAME);
      if(!titles.includes(CONFIG.CLIENT_LIST_SHEET_NAME)){
        throw new Error(`The "${CONFIG.CLIENT_LIST_SHEET_NAME}" sheet is missing from "${CONFIG.INVENTORY_LEDGER_NAME}".`);
      }
      if(missing.length)await createSheets(missing);

      const [inventoryRows,transactionRows,clientRows]=await Promise.all([
        getValues(ledgerId,CONFIG.INVENTORY_SHEET_NAME),
        getValues(ledgerId,CONFIG.TRANSACTIONS_SHEET_NAME),
        getValues(ledgerId,CONFIG.CLIENT_LIST_SHEET_NAME)
      ]);

      const columns=await ensureTransactionColumns(ledgerId,transactionRows);
      state.transactionIdx=columns.idx;
      state.transactionHeaderLength=columns.length;
      state.inventory=parseInventory(inventoryRows);
      state.transactions=parseTransactions(transactionRows,columns.idx);
      state.clients=parseClients(clientRows);

      renderInputs();
      renderAudit();
      setSyncStatus(`Synced at ${new Date().toLocaleTimeString()}`);
      return true;
    }catch(e){
      console.error(e);
      setSyncStatus(e.message||"Unable to load spreadsheet.",true);
      return false;
    }
  }

  async function findWorkbook(name){const query=[`name = '${escapeDriveQuery(name)}'`,`mimeType = 'application/vnd.google-apps.spreadsheet'`,`trashed = false`].join(" and ");const data=await fetchJson(`${DRIVE_API}?q=${encodeURIComponent(query)}&pageSize=10&fields=files(id,name,mimeType,modifiedTime,webViewLink)`,{headers:authHeaders()});return data.files?.[0]||null;}
  async function findMainWorkbook(){return(await findWorkbook(CONFIG.PRIMARY_WORKBOOK_NAME))||(await findWorkbook(CONFIG.FALLBACK_WORKBOOK_NAME));}
  async function createSheets(names){const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;await sheetsPost(`/${encodeURIComponent(ledgerId)}:batchUpdate`,{requests:names.map(title=>({addSheet:{properties:{title}}}))});if(names.includes(CONFIG.INVENTORY_SHEET_NAME))await updateValues(ledgerId,CONFIG.INVENTORY_SHEET_NAME,[["Asset","Balance"]]);if(names.includes(CONFIG.TRANSACTIONS_SHEET_NAME))await updateValues(ledgerId,CONFIG.TRANSACTIONS_SHEET_NAME,[["Timestamp","Client","Movement","Asset","Quantity","User","Comment","Image Link"]]);}
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
      image:findColumn(norm,["image link","image","photo","picture","photo link","attachment","drive link"])
    };
    return {idx,length:header.length};
  }

  function parseInventory(rows){if(!rows.length)return[];const header=rows[0].map(normalizeHeader);const assetIdx=findColumn(header,["asset","asset name","item","type"]);const balanceIdx=findColumn(header,["balance","current balance","stock","quantity"]);if(assetIdx<0)return[];return rows.slice(1).map((row,index)=>({rowNumber:index+2,asset:String(row[assetIdx]??"").trim(),balance:balanceIdx>=0?numericValue(row[balanceIdx]):0,assetColumn:assetIdx+1,balanceColumn:balanceIdx>=0?balanceIdx+1:2})).filter(x=>x.asset);}
  function parseTransactions(rows,idx){if(!rows.length)return[];return rows.slice(1).map(row=>({timestamp:idx.timestamp>=0?row[idx.timestamp]??"":"",client:idx.client>=0?row[idx.client]??"":"",movement:idx.movement>=0?row[idx.movement]??"":"",asset:idx.asset>=0?row[idx.asset]??"":"",quantity:idx.quantity>=0?numericValue(row[idx.quantity]):0,user:idx.user>=0?row[idx.user]??"":"",comment:idx.comment>=0?String(row[idx.comment]??"").trim():"",image:idx.image>=0?String(row[idx.image]??"").trim():""})).filter(x=>x.asset||x.client);}
  function parseClients(rows){if(!rows.length)return[];const header=rows[0].map(normalizeHeader);const idx=findColumn(header,["client","client name","name"]);if(idx<0)return rows.flat().map(x=>String(x).trim()).filter(Boolean).slice(1);return[...new Set(rows.slice(1).map(r=>String(r[idx]??"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));}

  function assetOptions(selected=""){return state.inventory.length?`<option value="">Select asset type</option>${state.inventory.map(x=>`<option value="${escapeAttr(x.asset)}" ${x.asset===selected?"selected":""}>${escapeHtml(x.asset)}</option>`).join("")}`:`<option value="">No assets configured</option>`;}
  function renderInputs(){const clients=state.clients.filter(c=>c!==SELF_CLIENT);$("client").innerHTML=clients.length?`<option value="">Select client</option>${clients.map(c=>`<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`:`<option value="">No clients found</option>`;renderAssetRows();populateFilterClients();handleMovementChange();updatePreview();}
  function populateFilterClients(){const select=$("filter-client");if(!select)return;const current=select.value;const names=[...new Set(state.transactions.map(t=>t.client).filter(Boolean))].sort((a,b)=>a.localeCompare(b));select.innerHTML=`<option value="">All clients</option>${names.map(c=>`<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join("")}`;if(names.includes(current))select.value=current;}
  function renderAssetRows(){const rows=[...document.querySelectorAll(".asset-row")];if(!rows.length){addAssetRow(false);return;}rows.forEach(row=>{const select=row.querySelector(".asset-select");select.innerHTML=assetOptions(select.value);});updateRemoveButtons();}
  function addAssetRow(focus=true){const wrap=$("asset-rows");const row=document.createElement("div");row.className="asset-row";row.innerHTML=`<div class="asset-row-number"></div><label class="asset-field"><span>Asset type</span><select class="asset-select" required>${assetOptions()}</select></label><label class="asset-field quantity-field"><span>Quantity</span><input class="asset-quantity" type="number" min="1" step="1" inputmode="numeric" placeholder="0" required></label><button type="button" class="remove-asset secondary" aria-label="Remove asset">Remove</button>`;wrap.appendChild(row);updateRemoveButtons();if(focus)row.querySelector(".asset-select").focus();updatePreview();}
  function removeAssetRow(index){const rows=[...document.querySelectorAll(".asset-row")];if(rows.length<=1)return;rows[index]?.remove();updateRemoveButtons();updatePreview();}
  function updateRemoveButtons(){const rows=[...document.querySelectorAll(".asset-row")];rows.forEach((row,i)=>{row.querySelector(".asset-row-number").textContent=String(i+1).padStart(2,"0");const b=row.querySelector(".remove-asset");b.dataset.removeRow=i;b.disabled=rows.length===1;});}
  function getAssetEntries(){return[...document.querySelectorAll(".asset-row")].map(row=>({asset:row.querySelector(".asset-select").value,quantity:Number(row.querySelector(".asset-quantity").value)}));}

  // ---- Photo attachment (optional) ----
  function handlePhotoChange(){const file=$("movement-photo").files[0];const wrap=$("photo-preview-wrap");const clearBtn=$("clear-photo");if(!file){wrap.classList.add("hidden");clearBtn.classList.add("hidden");return;}const url=URL.createObjectURL(file);$("photo-preview").src=url;wrap.classList.remove("hidden");clearBtn.classList.remove("hidden");}
  function clearPhoto(){$("movement-photo").value="";$("photo-preview-wrap").classList.add("hidden");$("photo-preview").src="";$("clear-photo").classList.add("hidden");}

  // ---- Audit trail: filtering + rendering ----
  function dateKey(timestamp){const d=new Date(timestamp);if(Number.isNaN(d.getTime()))return"";return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
  function todayDateStr(){return dateKey(new Date());}
  function filteredTransactions(){const client=$("filter-client")?.value||"";const date=$("filter-date")?.value||"";return state.transactions.filter(t=>(!client||t.client===client)&&(!date||dateKey(t.timestamp)===date));}
  function renderAudit(){
    const client=$("filter-client")?.value||"";const date=$("filter-date")?.value||"";
    const filtering=Boolean(client||date);
    const matches=filtering?filteredTransactions():state.transactions;
    const rows=(filtering?matches.slice(-200):matches.slice(-30)).reverse();
    state.auditVisibleRows=rows;
    const count=$("audit-count");
    if(count)count.textContent=rows.length.toLocaleString();
    const body=$("transactions-body");
    if(!body)return;
    body.innerHTML=rows.length?rows.map((item,i)=>{const hasInfo=Boolean(item.comment||item.image);return`<tr><td>${escapeHtml(formatTimestamp(item.timestamp))}</td><td><strong>${escapeHtml(item.client||"")}</strong></td><td><span class="movement-tag ${movementClass(item.movement)}">${escapeHtml(item.movement||"")}</span></td><td>${escapeHtml(item.asset||"")}</td><td class="num">${formatNumber(item.quantity)}</td><td>${escapeHtml(item.user||"")}</td><td>${hasInfo?`<button type="button" class="secondary info-button" data-info-index="${i}">View</button>`:`<span class="info-dash">—</span>`}</td></tr>`;}).join(""):emptyRow(7,filtering?"No asset movements match these filters.":"No asset movements recorded yet.");
  }

  // ---- More info modal ----
  function openInfoModal(index){
    const item=state.auditVisibleRows[index];if(!item)return;
    $("info-client").textContent=item.client||"—";
    $("info-movement").textContent=movementLabel(item.movement)||"—";
    $("info-asset").textContent=item.asset||"—";
    $("info-quantity").textContent=formatNumber(item.quantity);
    const photoWrap=$("info-photo-wrap");const commentEl=$("info-comment");const emptyEl=$("info-empty");
    if(item.image){const fileId=driveFileIdFromLink(item.image);$("info-photo").src=fileId?`https://drive.google.com/thumbnail?id=${fileId}&sz=w800`:item.image;$("info-photo-link").href=item.image;photoWrap.classList.remove("hidden");}else photoWrap.classList.add("hidden");
    if(item.comment){commentEl.textContent=item.comment;commentEl.classList.remove("hidden");}else commentEl.classList.add("hidden");
    emptyEl.classList.toggle("hidden",Boolean(item.image||item.comment));
    $("info-modal").classList.remove("hidden");
  }
  function closeInfoModal(){$("info-modal").classList.add("hidden");}
  function driveFileIdFromLink(url){const m=String(url||"").match(/\/d\/([a-zA-Z0-9_-]+)/)||String(url||"").match(/[?&]id=([a-zA-Z0-9_-]+)/);return m?m[1]:"";}

  function handleMovementChange(){
    const discard=$("movement").value==="DISCARD";
    const client=$("client");
    let selfOption=client.querySelector("option[data-self-client]");
    if(discard){
      if(!selfOption){
        selfOption=document.createElement("option");
        selfOption.dataset.selfClient="1";
        selfOption.value=SELF_CLIENT;
        selfOption.textContent=SELF_CLIENT;
        client.insertBefore(selfOption,client.firstChild);
      }
      client.value=SELF_CLIENT;
      client.disabled=true;
      $("client-help").textContent=`Discarded assets are automatically recorded against ${SELF_CLIENT}.`;
    }else{
      if(selfOption)selfOption.remove();
      client.disabled=false;
      if(client.value===SELF_CLIENT)client.value="";
      $("client-help").textContent="";
    }
    updatePreview();
  }
  function reviewMovement(event){
    event.preventDefault();const data=readForm();const error=validateMovement(data);if(error){setMovementStatus(error,true);return;}
    const balances=new Map(state.inventory.map(i=>[i.asset.toLowerCase(),i.balance]));
    for(const entry of data.items){const item=state.inventory.find(x=>x.asset.toLowerCase()===entry.asset.toLowerCase());if(!item){setMovementStatus(`Asset "${entry.asset}" is not present in Inventory.`,true);return;}if(data.movement!=="RECEIVED"){const next=(balances.get(item.asset.toLowerCase())??0)-entry.quantity;if(next<0){setMovementStatus(`Cannot remove ${entry.quantity} ${item.asset}. Current balance is ${formatNumber(balances.get(item.asset.toLowerCase()))}.`,true);return;}balances.set(item.asset.toLowerCase(),next);}else balances.set(item.asset.toLowerCase(),(balances.get(item.asset.toLowerCase())??0)+entry.quantity);}
    pendingMovement={...data,balances};
    $("confirm-movement").textContent=movementLabel(data.movement);$("confirm-client").textContent=data.client;$("confirm-time").textContent=data.timestamp;
    $("confirm-items").innerHTML=data.items.map((x,i)=>`<div class="confirm-item"><span>${escapeHtml(x.asset)}</span><strong>${formatNumber(x.quantity)}</strong><button type="button" class="icon-button" data-remove-row="${i}" aria-label="Remove">×</button></div>`).join("");
    $("confirm-total").textContent=formatNumber(data.items.reduce((s,x)=>s+x.quantity,0));
    const attachment=$("confirm-attachment");const photoWrap=$("confirm-photo-wrap");const commentEl=$("confirm-comment");
    let showAttachment=false;
    if(data.photoFile){$("confirm-photo").src=URL.createObjectURL(data.photoFile);photoWrap.classList.remove("hidden");showAttachment=true;}else photoWrap.classList.add("hidden");
    if(data.comment){commentEl.textContent=data.comment;commentEl.classList.remove("hidden");showAttachment=true;}else commentEl.classList.add("hidden");
    attachment.classList.toggle("hidden",!showAttachment);
    const warning=$("confirm-warning");warning.classList.toggle("hidden",data.movement!=="DISCARD");if(data.movement==="DISCARD")warning.textContent="Discard is permanent in the warehouse balance. Please make sure all quantities are correct.";
    $("confirm-modal").classList.remove("hidden");setTimeout(()=>$("approve-confirm").focus(),50);
  }
  async function approveMovement(){
    if(!pendingMovement)return;const data=pendingMovement;$("approve-confirm").disabled=true;$("cancel-confirm").disabled=true;setMovementStatus("Recording movements...");
    try{
      const ledgerId=CONFIG.INVENTORY_LEDGER_SHEET_ID;const user=state.idTokenPayload?.email||readSavedSession()?.email||"Google user";const timestamp=new Date().toISOString();
      let imageLink="";
      if(data.photoFile){
        setMovementStatus("Uploading photo...");
        const folderId=await getOrCreateDailyFolder(todayDateStr());
        const filename=buildPhotoFilename(data);
        const uploaded=await uploadImageToDrive(data.photoFile,folderId,filename);
        imageLink=uploaded.webViewLink||`https://drive.google.com/file/d/${uploaded.id}/view`;
        setMovementStatus("Recording movements...");
      }
      const idx=state.transactionIdx;const len=state.transactionHeaderLength||6;
      const rows=data.items.map(x=>buildTransactionRow(idx,len,{timestamp,client:data.client,movement:data.movement,asset:x.asset,quantity:x.quantity,user,comment:data.comment||"",image:imageLink}));
      const transactionRange=`${quoteSheetName(CONFIG.TRANSACTIONS_SHEET_NAME)}!A:${columnLetter(len)}`;
      await sheetsPost(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(transactionRange)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,{values:rows});
      for(const x of data.items){const item=state.inventory.find(i=>i.asset.toLowerCase()===x.asset.toLowerCase());const balanceCell=columnLetter(item.balanceColumn)+item.rowNumber;const inventoryRange=`${quoteSheetName(CONFIG.INVENTORY_SHEET_NAME)}!${balanceCell}`;await sheetsPut(`/${encodeURIComponent(ledgerId)}/values/${encodeURIComponent(inventoryRange)}?valueInputOption=USER_ENTERED`,{range:inventoryRange,majorDimension:"ROWS",values:[[data.balances.get(item.asset.toLowerCase())]]});}
      closeConfirm();resetMovementForm();setMovementStatus(`${data.items.length} movement${data.items.length===1?"":"s"} recorded successfully.`);await loadLogger();
    }catch(e){console.error(e);setMovementStatus(e.message||"Unable to record movement.",true);}
    finally{$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
  }
  function buildTransactionRow(idx,len,values){const row=new Array(len).fill("");const set=(i,v)=>{if(i>=0&&i<len)row[i]=v;};set(idx.timestamp,values.timestamp);set(idx.client,values.client);set(idx.movement,values.movement);set(idx.asset,values.asset);set(idx.quantity,values.quantity);set(idx.user,values.user);set(idx.comment,values.comment);set(idx.image,values.image);return row;}
  function buildPhotoFilename(data){const safeClient=String(data.client||"client").replace(/[^a-z0-9]+/gi,"-").replace(/^-+|-+$/g,"")||"client";const stamp=new Date().toISOString().replace(/[:.]/g,"-");const ext=(data.photoFile.name.match(/\.[a-zA-Z0-9]+$/)||[".jpg"])[0];return`${safeClient}-${data.movement}-${stamp}${ext}`;}

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

  function resetMovementForm(){$("movement-form").reset();$("client").disabled=false;$("asset-rows").innerHTML="";addAssetRow(false);clearPhoto();setDefaultTimestamp();handleMovementChange();}
  function closeConfirm(){$("confirm-modal").classList.add("hidden");pendingMovement=null;$("approve-confirm").disabled=false;$("cancel-confirm").disabled=false;}
  function readForm(){const movement=$("movement").value;const client=movement==="DISCARD"?SELF_CLIENT:$("client").value.trim();return{movement,client,timestamp:$("timestamp").value,items:getAssetEntries(),comment:$("movement-comment").value.trim(),photoFile:$("movement-photo").files[0]||null};}
  function validateMovement(d){if(!d.movement||!d.client||!d.timestamp)return"Please complete the movement, client and time fields.";if(!d.items.length)return"Add at least one asset type.";if(d.items.some(x=>!x.asset||!Number.isInteger(x.quantity)||x.quantity<=0))return"Select an asset type and enter a whole quantity greater than zero for every row.";const seen=new Set();for(const x of d.items){const k=x.asset.toLowerCase();if(seen.has(k))return`You have selected ${x.asset} more than once. Combine the quantities into one row.`;seen.add(k);}return null;}
  function updatePreview(){const d=readForm();if(!d.client||!d.items.length||d.items.some(x=>!x.asset||!Number.isInteger(x.quantity)||x.quantity<=0)){$("preview-text").textContent="Select a client and add one or more asset types with quantities.";return;}const list=d.items.map(x=>`${formatNumber(x.quantity)} × ${x.asset}`).join(" • ");$("preview-text").textContent=`${movementLabel(d.movement)} ${list} for ${d.client}`;}
  function setDefaultTimestamp(){const input=$("timestamp");if(!input)return;const now=new Date();input.value=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;}
  function movementLabel(x){return x==="RECEIVED"?"Received":x==="SENT"?"Sent":x==="DISCARD"?"Discard":String(x||"");}function movementClass(x){return String(x||"").toLowerCase();}function formatTimestamp(x){if(!x)return"";const d=new Date(x);return Number.isNaN(d.getTime())?String(x):d.toLocaleString("en-GB",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"});}function formatNumber(x){return Number(x||0).toLocaleString("en-GB");}function numericValue(x){if(x===null||x===undefined||x==="")return 0;const n=Number(String(x).replace(/,/g,""));return Number.isFinite(n)?n:0;}function normalizeHeader(x){return String(x??"").trim().toLowerCase().replace(/\s+/g," ");}function findColumn(headers,names){for(const name of names){const idx=headers.indexOf(name);if(idx>=0)return idx;}return-1;}function quoteSheetName(x){return `'${String(x).replace(/'/g,"''")}'`;}function columnLetter(n){let r="";while(n>0){const rem=(n-1)%26;r=String.fromCharCode(65+rem)+r;n=Math.floor((n-1)/26);}return r;}function escapeDriveQuery(x){return String(x).replace(/\\/g,"\\\\").replace(/'/g,"\\'");}function escapeHtml(x){return String(x??"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}function escapeAttr(x){return escapeHtml(x);}function emptyRow(colspan,text){return `<tr><td colspan="${colspan}" class="empty">${escapeHtml(text)}</td></tr>`;}
  function authHeaders(){return{Authorization:`Bearer ${state.accessToken}`};}
  async function sheetsGet(path){return fetchJson(SHEETS_API+path,{headers:authHeaders()});}async function sheetsPost(path,body){return fetchJson(SHEETS_API+path,{method:"POST",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}async function sheetsPut(path,body){return fetchJson(SHEETS_API+path,{method:"PUT",headers:{...authHeaders(),"Content-Type":"application/json"},body:JSON.stringify(body)});}
  async function fetchJson(url,options={}){
    let response=await fetch(url,options);
    if(response.status===401&&!options.__retried){
      try{
        await acquireAccessToken("none",state.idTokenPayload?.email||readSavedSession()?.email);
        const retry={...options,__retried:true,headers:{...(options.headers||{}),...authHeaders()}};
        return fetchJson(url,retry);
      }catch(e){
        state.accessToken=null;
        throw new Error("Google access expired. Click Connect Google Sheets to reconnect.");
      }
    }
    const text=await response.text();
    let data={};
    try{data=text?JSON.parse(text):{};}catch(_){}
    if(!response.ok)throw new Error(data?.error?.message||`Request failed (${response.status})`);
    return data;
  }
  function decodeJwtPayload(jwt){const parts=String(jwt).split(".");if(parts.length!==3)throw new Error("Invalid Google credential.");const base64=parts[1].replace(/-/g,"+").replace(/_/g,"/");const padded=base64+"=".repeat((4-base64.length%4)%4);return JSON.parse(decodeURIComponent(Array.from(atob(padded)).map(c=>`%${c.charCodeAt(0).toString(16).padStart(2,"0")}`).join("")));}
  function saveSession(){if(!state.idTokenPayload)return;localStorage.setItem(SESSION_KEY,JSON.stringify({name:state.idTokenPayload.name||"Google user",email:state.idTokenPayload.email||"",picture:state.idTokenPayload.picture||"",sub:state.idTokenPayload.sub||""}));}
  function readSavedSession(){try{return JSON.parse(localStorage.getItem(SESSION_KEY)||"null");}catch(_){return null;}}
  function setUserProfile(profile){
    $("user-name").textContent=profile?.name||"Google user";
    $("user-email").textContent=profile?.email||"";
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

  function signOut(){
    const saved=readSavedSession();
    if(state.idTokenPayload?.sub){try{google.accounts.id.revoke(state.idTokenPayload.sub,()=>{});}catch(_){}}
    if(saved?.sub&&saved.sub!==state.idTokenPayload?.sub){try{google.accounts.id.revoke(saved.sub,()=>{});}catch(_){}}
    localStorage.removeItem(SESSION_KEY);
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

  // Keep this helper available to any older dashboard/logger code that calls it globally.
  window.formatTimestamp = formatTimestamp;
})();
