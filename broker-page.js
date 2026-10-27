(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const spreadsheet = CONFIG.INVENTORY_LEDGER_SHEET_ID;
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheet)}`;
  let token = "";
  let profile = null;
  let busy = false;
  const sessionKey = "fmAssetSession";
  const cachedSession = () => {
    try { return JSON.parse(localStorage.getItem(sessionKey) || "null"); }
    catch (_) { return null; }
  };
  const setStatus = (message, failure = false) => {
    const node = $("auth-status"); if (!node) return;
    node.textContent = message;
    node.classList.toggle("error", failure);
  };
  const report = (message, failure = false) => {
    const node = $("sync-status"); if (!node) return;
    node.textContent = message; node.classList.toggle("error", failure);
  };
  function setProfile(user) {
    profile = user;
    const first = String(user?.given_name || user?.name || "there").trim().split(/\s+/)[0];
    $("user-name").textContent = `Hi, ${first} 👋`;
    $("user-email").textContent = "";
    if (user?.picture) { $("user-photo").src = user.picture; $("user-photo").classList.remove("hidden"); }
    $("sign-out").classList.remove("hidden");
    $("google-signin-button").classList.add("hidden");
  }
  function decodeJwt(jwt) {
    const encoded = String(jwt || "").split(".")[1];
    if (!encoded) throw new Error("No Google identity returned");
    const bytes = atob(encoded.replace(/-/g,"+").replace(/_/g,"/"));
    const chars = Uint8Array.from(bytes, c => c.charCodeAt(0));
    return JSON.parse(new TextDecoder().decode(chars));
  }
  function initGoogle() {
    if (!window.google?.accounts?.id || !window.google?.accounts?.oauth2) {
      setStatus("Google sign-in could not load. Refresh the page.", true);
      return;
    }
    google.accounts.id.initialize({client_id:CONFIG.GOOGLE_CLIENT_ID, callback: result => {
      try {
        const user = decodeJwt(result.credential);
        localStorage.setItem(sessionKey, JSON.stringify({name:user.name||"Google user", email:user.email||"", picture:user.picture||"", sub:user.sub||""}));
        setProfile(user);
        $("grant-access").classList.remove("hidden");
        setStatus("Select Connect Google Sheets to authorise broker data.");
      } catch (error) { setStatus(error.message, true); }
    }, auto_select:true});
    const saved = cachedSession();
    if (saved?.email) {
      setProfile(saved);
      const cached = window.FM_AUTH_CACHE?.read?.(saved.email);
      if (cached?.token) { token=cached.token; showWorkspace(); refreshData(); }
      else { $("grant-access").classList.remove("hidden"); setStatus("Reconnect Google Sheets to view your broker ledger."); }
    } else {
      google.accounts.id.renderButton($("google-signin-button"),{theme:"outline",size:"large",text:"signin_with",shape:"rectangular",width:280});
      google.accounts.id.prompt();
    }
  }
  function connect() {
    if (!window.google?.accounts?.oauth2) { setStatus("Google sign-in is not ready. Refresh and try again.",true); return; }
    if (!profile?.email) { setStatus("Sign in with Google first.",true); return; }
    $("grant-access").disabled = true;
    setStatus("Confirm access in the Google permission window.");
    let resolved = false;
    const client = google.accounts.oauth2.initTokenClient({
      client_id:CONFIG.GOOGLE_CLIENT_ID,scope:CONFIG.OAUTH_SCOPES,
      callback: result => {
        resolved = true;
        $("grant-access").disabled = false;
        if (!result?.access_token) { setStatus("Google access was not granted. Select Connect to retry.",true); return; }
        token=result.access_token;
        window.FM_AUTH_CACHE?.write?.(token,result.expires_in,profile.email);
        showWorkspace(); refreshData();
      },
      error_callback: error => {
        resolved = true; $("grant-access").disabled=false;
        setStatus(error?.type === "popup_failed_to_open" ? "Your browser blocked the Google prompt. Allow popups and try again." : "Could not connect to Google Sheets.",true);
      }
    });
    try { client.requestAccessToken({prompt:"consent",login_hint:profile.email}); }
    catch(error) { $("grant-access").disabled=false; setStatus(error.message,true); }
  }
  function showWorkspace() {
    $("login-card").classList.add("hidden");
    $("broker-workspace").classList.remove("hidden");
    $("grant-access").classList.add("hidden");
  }
  async function getRows(sheetName, optional = false) {
    const range = `'${sheetName.replace(/'/g,"''")}'!A:AE`;
    const url = `${base}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
    const response = await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
    if (!response.ok) {
      if (optional && response.status === 400) return [];
      if (response.status === 401) { token=""; window.FM_AUTH_CACHE?.clear?.(); throw new Error("Your Sheets session expired. Please reconnect."); }
      const body = await response.json().catch(()=>({}));
      throw new Error(body.error?.message || `Unable to read ${sheetName} (${response.status})`);
    }
    return (await response.json()).values || [];
  }
  function parseTransactions(rows) {
    if (!rows?.length) return [];
    const headings=rows[0].map(x=>String(x||"").toLowerCase().trim().replace(/\s+/g," "));
    const index=(...names)=>headings.findIndex(h=>names.includes(h));
    const fields={timestamp:index("timestamp","date","datetime"),client:index("client","client name"),movement:index("movement","type","direction"),asset:index("asset","asset name","item"),quantity:index("quantity","qty"),broker:index("broker","broker name")};
    const pick=(row,id)=>fields[id]<0?"":String(row[fields[id]]??"").trim();
    return rows.slice(1).map(row=>({timestamp:pick(row,"timestamp"),client:pick(row,"client"),movement:pick(row,"movement"),asset:pick(row,"asset"),quantity:Number(pick(row,"quantity").replace(/,/g,""))||0,broker:pick(row,"broker")}))
      .filter(row=>row.client&&row.asset);
  }
  async function refreshData() {
    if (busy || !token) return;
    busy=true; $("refresh").disabled=true;
    report("Refreshing broker movements…");
    try {
      const [primary,operations,brokerRows] = await Promise.all([
        getRows(CONFIG.TRANSACTIONS_SHEET_NAME),
        getRows(CONFIG.ASSET_OPERATIONS_SHEET_NAME,true),
        getRows(CONFIG.BROKERS_SHEET_NAME,true)
      ]);
      const rows=[...parseTransactions(primary),...parseTransactions(operations)]
        .filter(row=>!["NORMALISATION","NORMALIZATION","LEDGER ADJUSTMENT"].includes(row.movement.toUpperCase()));
      const brokers=window.FM_BROKERS.parseBrokers(brokerRows);
      window.FM_BROKER_VIEW.render({rows,brokers});
      report(`Updated ${new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})} · ${rows.filter(r=>r.broker).length.toLocaleString("en-GB")} tagged ledger lines`);
    } catch(error) {
      report(error.message||"Unable to load broker data.",true);
      if (!token) { $("broker-workspace").classList.add("hidden"); $("login-card").classList.remove("hidden"); $("grant-access").classList.remove("hidden"); }
    } finally { busy=false; $("refresh").disabled=false; }
  }
  function initNavigation() {
    const nav=document.querySelector(".quick-nav"), toggle=$("menu-toggle");
    toggle?.addEventListener("click",()=>{const open=nav.classList.toggle("menu-open");toggle.setAttribute("aria-expanded",String(open));});
    document.addEventListener("click",e=>{if(nav?.classList.contains("menu-open")&&!nav.contains(e.target)){nav.classList.remove("menu-open");toggle?.setAttribute("aria-expanded","false");}});
  }
  document.addEventListener("DOMContentLoaded",()=>{
    initNavigation();
    $("grant-access").addEventListener("click",connect);
    $("refresh").addEventListener("click",refreshData);
    $("sign-out").addEventListener("click",()=>{
      window.FM_AUTH_CACHE?.clear?.(); localStorage.removeItem(sessionKey); location.reload();
    });
    if (window.google?.accounts?.oauth2 && window.google?.accounts?.id) initGoogle();
    else window.addEventListener("fm-google-loaded",initGoogle,{once:true});
  });
})();
