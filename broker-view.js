(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const key = value => window.FM_BROKERS?.key(value) || String(value || "").trim().toUpperCase();
  const clean = value => String(value ?? "").trim();
  const esc = value => String(value ?? "").replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  const number = value => Number(value || 0).toLocaleString("en-GB");
  const dateFormatter = new Intl.DateTimeFormat("en-GB", {timeZone:"Europe/London",year:"numeric",month:"2-digit",day:"2-digit"});
  const monthFormatter = new Intl.DateTimeFormat("en-GB", {month:"short", day:"numeric", timeZone:"UTC"});
  const model = {rows:[],brokers:[],selected:"",client:"",asset:"",period:"all",from:"",to:"",graphRange:"7d",more:20};
  let bound = false;
  function day(value) {
    const d = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(d.getTime())) return "";
    const parts = Object.fromEntries(dateFormatter.formatToParts(d).filter(p=>p.type!=="literal").map(p=>[p.type,p.value]));
    return `${parts.year}-${parts.month}-${parts.day}`;
  }
  function dayOffset(date, count) {
    const d = new Date(`${date}T12:00:00Z`);
    d.setUTCDate(d.getUTCDate() + count);
    return d.toISOString().slice(0,10);
  }
  function periodWindow() {
    const today=day(new Date());
    if (model.period==='all') return {start:'',end:''};
    if (model.period==='custom') return {start:model.from,end:model.to};
    return {start:dayOffset(today,1-Number(model.period)),end:today};
  }
  function inWindow(row,range=periodWindow()) {
    const date=day(row.timestamp);
    return Boolean(date) && (!range.start || date>=range.start) && (!range.end || date<=range.end);
  }
  function tagged() { return model.rows.filter(row=>clean(row.broker) && ['SENT','RECEIVED'].includes(key(row.movement))); }
  function brokerMatches(row) { return key(row.broker)===key(model.selected); }
  function selectedRows() {
    return tagged().filter(row=>brokerMatches(row) && (!model.client || key(row.client)===key(model.client)) &&
      (!model.asset || key(row.asset)===key(model.asset)) && inWindow(row));
  }
  function activeAssets(rows) { return [...new Set(rows.map(r=>clean(r.asset)).filter(Boolean))].sort((a,b)=>a.localeCompare(b)); }
  function optionHTML(label,items,value) {return `<option value="">${esc(label)}</option>`+items.map(v=>`<option value="${esc(v)}" ${key(v)===key(value)?'selected':''}>${esc(v)}</option>`).join('');}
  function brokersFromRows() {
    return window.FM_BROKERS.parseBrokers([...model.brokers.map(n=>[n]),...tagged().map(row=>[row.broker])]);
  }
  function renderCards() {
    const root=$('broker-cards'); if(!root)return;
    const dateRange=periodWindow();
    const brokers=brokersFromRows();
    const bucket=new Map(brokers.map(name=>[key(name),{name,rows:[]}]));
    tagged().filter(row=>inWindow(row,dateRange)).forEach(row=>{
      const bucketKey=key(row.broker);
      if(!bucket.has(bucketKey))bucket.set(bucketKey,{name:clean(row.broker),rows:[]});
      bucket.get(bucketKey).rows.push(row);
    });
    root.innerHTML=[...bucket.values()].map(({name,rows})=>{
      const filtered=rows.filter(r=>(!model.client||key(r.client)===key(model.client))&&(!model.asset||key(r.asset)===key(model.asset)));
      const clients=new Set(filtered.map(r=>key(r.client)));
      const assets=activeAssets(filtered);
      return `<button type="button" class="broker-card${key(model.selected)===key(name)?' is-selected':''}" data-broker-card="${esc(name)}" aria-pressed="${key(model.selected)===key(name)}">
        <span class="broker-card-top"><span class="broker-card-monogram">${esc(name.split(/\s/).map(w=>w[0]).join('').slice(0,2))}</span><span class="broker-card-indicator"></span></span>
        <strong>${esc(name)}</strong>
        <span class="broker-card-metrics"><span><b>${clients.size}</b> client${clients.size===1?'':'s'}</span><span><b>${assets.length}</b> asset type${assets.length===1?'':'s'}</span></span>
        <span class="broker-card-foot">${filtered.length?`${number(filtered.length)} ledger lines`:'No tagged movements yet'}</span>
      </button>`;
    }).join('');
  }
  function renderBreakdown() {
    const el=$('broker-details');if(!el)return;
    if(!model.selected){el.innerHTML=`<div class="broker-empty"><div class="broker-empty-icon">↗</div><strong>Select a broker to view their clients</strong><span>Tap any broker card above to see assets sent, received, and the remaining net by asset type.</span></div>`;
      if($('broker-history'))$('broker-history').innerHTML='<div class="broker-history-empty">Choose a broker to see movements.</div>';
      renderTrend();return;}
    const rows=selectedRows();
    const summaries=window.FM_BROKERS.summarize(rows);
    const allRows=summaries.flatMap(entry=>entry.assets.map(asset=>({client:entry.client,...asset})));
    const label=model.period==='all'?'All-time broker ledger':'Selected period';
    el.innerHTML=`<div class="broker-details-top"><div><div class="eyebrow">${esc(label)}</div><h3>${esc(model.selected)} · Client assets</h3><p>Net with broker = Sent − Received. Figures are per asset type, not combined across different assets.</p></div><span class="broker-details-lines">${number(rows.length)} lines</span></div>
      <div class="broker-table-scroll"><table class="broker-table"><thead><tr><th>Client</th><th>Asset type</th><th class="num">Sent</th><th class="num">Received</th><th class="num">Net with broker</th></tr></thead><tbody>
      ${allRows.length?allRows.map(entry=>`<tr><td data-label="Client"><strong>${esc(entry.client)}</strong></td><td data-label="Asset">${esc(entry.asset)}</td><td data-label="Sent" class="num">${number(entry.sent)}</td><td data-label="Received" class="num">${number(entry.received)}</td><td data-label="Net" class="num"><span class="broker-net ${entry.net>0?'broker-net-positive':entry.net<0?'broker-net-negative':''}">${entry.net>0?'+':''}${number(entry.net)}</span></td></tr>`).join(''):`<tr><td colspan="5" class="broker-empty-table">No tagged transactions match these filters.</td></tr>`}
      </tbody></table></div>`;
    const history=$('broker-history');
    const sorted=rows.slice().sort((a,b)=>new Date(b.timestamp)-new Date(a.timestamp));
    if(history){
      const visible=sorted.slice(0,model.more);
      history.innerHTML=visible.length?visible.map(row=>`<div class="broker-history-row"><span class="broker-history-date">${esc(new Date(row.timestamp).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}))}</span><span><strong>${esc(row.client)}</strong><small>${esc(row.asset)}</small></span><span class="broker-direction ${key(row.movement)==='SENT'?'out':'in'}">${key(row.movement)==='SENT'?'Sent':'Received'}</span><strong class="broker-history-qty">${number(row.quantity)}</strong></div>`).join('')+(sorted.length>visible.length?`<button type="button" class="secondary broker-history-more" data-broker-more>Load 20 more (${number(sorted.length-visible.length)} remaining)</button>`:''):`<div class="broker-history-empty">No transactions in this view.</div>`;
    }
    renderTrend();
  }
  function graphWindow() {
    const today=day(new Date());
    const n={"7d":7,"30d":30,"90d":90}[model.graphRange];
    const dates=model.rows.filter(brokerMatches).map(r=>day(r.timestamp)).filter(Boolean).sort();
    return {start:n?dayOffset(today,1-n):(dates[0]||today),end:today};
  }
  function renderTrend() {
    const el=$('broker-trend'); const title=$('broker-trend-title');if(!el||!title)return;
    const all=tagged().filter(brokerMatches).filter(row=>!model.client||key(row.client)===key(model.client));
    const asset=model.asset||activeAssets(all)[0]||'';
    title.textContent=asset?`${asset} · Sent vs received`:'Sent vs received';
    document.querySelectorAll('[data-broker-graph-range]').forEach(b=>{const yes=b.dataset.brokerGraphRange===model.graphRange;b.classList.toggle('is-active',yes);b.setAttribute('aria-pressed',String(yes));});
    if(!asset){el.innerHTML='<div class="broker-chart-empty">Select a broker with transactions to display its trend.</div>';return;}
    const range=graphWindow();
    const rows=all.filter(row=>key(row.asset)===key(asset)&&inWindow(row,range));
    const daily=new Map();
    for(const r of rows){const d=day(r.timestamp);if(!d)continue;if(!daily.has(d))daily.set(d,{sent:0,received:0});daily.get(d)[key(r.movement)==='SENT'?'sent':'received']+=Math.abs(Number(r.quantity)||0);}
    if(!rows.length){el.innerHTML='<div class="broker-chart-empty">No movement for this asset during this timeframe. Try 30D, 90D or All.</div>';return;}
    const days=Math.max(1,Math.round((Date.parse(range.end+'T12:00Z')-Date.parse(range.start+'T12:00Z'))/86400000)+1);
    const granularity=days>180?'month':days>45?'week':'day';
    const buckets=new Map();
    function bucket(date){if(granularity==='month')return date.slice(0,7)+'-01';if(granularity==='week'){const d=new Date(date+'T12:00Z');const weekday=(d.getUTCDay()+6)%7;d.setUTCDate(d.getUTCDate()-weekday);return d.toISOString().slice(0,10);}return date;}
    for(const [date,value] of daily){const b=bucket(date);const prior=buckets.get(b)||{sent:0,received:0};prior.sent+=value.sent;prior.received+=value.received;buckets.set(b,prior);}
    let cursor=bucket(range.start);const end=bucket(range.end),points=[];
    for(let i=0;cursor<=end&&i<360;i++){
      points.push([cursor,buckets.get(cursor)||{sent:0,received:0}]);
      const d=new Date(cursor+'T12:00Z');
      if(granularity==='month')d.setUTCMonth(d.getUTCMonth()+1);
      else d.setUTCDate(d.getUTCDate()+(granularity==='week'?7:1));
      cursor=d.toISOString().slice(0,10);
    }
    const width=Math.max(540,points.length*44+95),height=256,left=54,right=22,top=24,bottom=45;
    const innerW=width-left-right,innerH=height-top-bottom;
    const rawMax=Math.max(1,...points.map(([,v])=>Math.max(v.sent,v.received)));
    const max=Math.ceil(rawMax/4)*4;
    const X=i=>left+(points.length===1?innerW/2:(i/(points.length-1))*innerW);
    const Y=v=>top+innerH*(1-v/max);
    const path=k=>points.map(([,v],i)=>`${i?'L':'M'}${X(i).toFixed(1)},${Y(v[k]).toFixed(1)}`).join(' ');
    const label=date=>monthFormatter.format(new Date(date+'T12:00Z'));
    const lines=Array.from({length:5},(_,i)=>{const y=top+innerH*i/4;return `<line x1="${left}" y1="${y}" x2="${width-right}" y2="${y}" class="broker-chart-grid"/><text x="${left-11}" y="${y+4}" text-anchor="end" class="broker-chart-axis">${number(Math.round(max*(1-i/4)))}</text>`;}).join('');
    const ticks=points.filter((_,i)=>i%Math.max(1,Math.ceil(points.length/7))===0||i===points.length-1).map(([d],idx)=>{const i=points.findIndex(p=>p[0]===d);return `<text x="${X(i)}" y="${height-13}" text-anchor="middle" class="broker-chart-axis">${esc(label(d))}</text>`;}).join('');
    const area=k=>`${path(k)} L${X(points.length-1)},${Y(0)} L${X(0)},${Y(0)} Z`;
    const dots=points.map(([d,v],i)=>['sent','received'].map(k=>`<circle cx="${X(i)}" cy="${Y(v[k])}" r="4" class="broker-chart-dot ${k}"><title>${esc(label(d))} · ${k==='sent'?'Sent':'Received'}: ${number(v[k])}</title></circle>`).join('')).join('');
    el.innerHTML=`<div class="broker-chart-key"><span><i class="sent"></i>Sent</span><span><i class="received"></i>Received</span></div><div class="broker-chart-scroll"><svg viewBox="0 0 ${width} ${height}" style="min-width:${width}px" aria-label="${esc(asset)} sent and received graph" role="img"><defs><linearGradient id="broker-blue-area" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#4377ef" stop-opacity=".17"/><stop offset="1" stop-color="#4377ef" stop-opacity="0"/></linearGradient></defs>${lines}${ticks}<path d="${area('sent')}" fill="url(#broker-blue-area)"/><path d="${path('sent')}" class="broker-chart-sent"/><path d="${path('received')}" class="broker-chart-received"/>${dots}</svg></div>`;
  }
  function render({rows,brokers}) {
    if(!$('brokers-section'))return;
    model.rows=Array.isArray(rows)?rows:[];
    model.brokers=Array.isArray(brokers)?brokers:[];
    const names=brokersFromRows();
    if(model.selected&&!names.some(n=>key(n)===key(model.selected)))model.selected='';
    const broker=$('broker-select');if(broker)broker.innerHTML=optionHTML('Choose a broker',names,model.selected);
    const eligible=tagged().filter(r=>!model.selected||brokerMatches(r));
    const clients=[...new Set(eligible.map(r=>clean(r.client)).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    const assets=activeAssets(eligible);
    if(model.client&&!clients.some(c=>key(c)===key(model.client)))model.client='';
    if(model.asset&&!assets.some(a=>key(a)===key(model.asset)))model.asset='';
    $('broker-client').innerHTML=optionHTML('All clients',clients,model.client);
    $('broker-asset').innerHTML=optionHTML('All asset types',assets,model.asset);
    $('broker-period').value=model.period;
    $('broker-date-fields').classList.toggle('hidden',model.period!=='custom');
    $('broker-from').value=model.from;$('broker-to').value=model.to;
    renderCards();renderBreakdown();
    if(!bound){bound=true;bind();}
  }
  function bind() {
    $('broker-cards').addEventListener('click',e=>{
      const el=e.target.closest('[data-broker-card]');if(!el)return;
      model.selected=el.dataset.brokerCard;
      model.client='';model.asset='';model.more=20;
      render({rows:model.rows,brokers:model.brokers});
      if(window.matchMedia('(max-width:760px)').matches) $('broker-details')?.scrollIntoView({block:'nearest',behavior:'smooth'});
    });
    ['broker-select','broker-client','broker-asset','broker-period','broker-from','broker-to'].forEach(id=>{
      $(id)?.addEventListener('change',()=>{
        model.selected=$('broker-select').value;
        model.client=$('broker-client').value;
        model.asset=$('broker-asset').value;
        model.period=$('broker-period').value;
        model.from=$('broker-from').value;
        model.to=$('broker-to').value;
        model.more=20;
        render({rows:model.rows,brokers:model.brokers});
      });
    });
    $('broker-reset')?.addEventListener('click',()=>{Object.assign(model,{selected:'',client:'',asset:'',period:'all',from:'',to:'',graphRange:'7d',more:20});render({rows:model.rows,brokers:model.brokers});});
    $('brokers-section').addEventListener('click',e=>{
      const range=e.target.closest('[data-broker-graph-range]');
      if(range){model.graphRange=range.dataset.brokerGraphRange;renderTrend();return;}
      if(e.target.closest('[data-broker-more]')){model.more+=20;renderBreakdown();}
    });
  }
  window.FM_BROKER_VIEW={render};
})();
