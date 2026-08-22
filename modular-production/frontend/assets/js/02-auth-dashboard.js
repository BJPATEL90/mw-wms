function doLogin(){
  const u   = document.getElementById('loginUser').value.trim();
  const p   = document.getElementById('loginPass').value.trim();
  const err = document.getElementById('loginErr');
  const btn = document.querySelector('#loginOverlay .btn-primary');

  if(!u || !p){ err.innerText = 'Enter User ID and Password'; return; }

  err.innerText = '';
  if(btn){ btn.disabled = true; btn.innerText = 'Signing in…'; }

  // Try sheet-based login first
  callSheet('login', { userId: u, password: p },
    function(data){
      if(btn){ btn.disabled = false; btn.innerText = 'Sign In'; }
      if(data && data.success){
        _completeLogin(data.name, data.role, data.userId, false, data.allowedSections);
      } else {
        err.innerText = (data && data.message) || 'Invalid credentials';
      }
    },
    function(sheetErr){
      if(btn){ btn.disabled = false; btn.innerText = 'Sign In'; }
      const gasConfigured = !!(GAS_URL || ls('gasUrl'));
      if (gasConfigured) {
        // GAS URL is set but call failed — server error, NOT offline
        // Don't use offline fallback — show the real error
        err.innerText = 'Server error: ' + sheetErr + '. Check GAS deployment is up to date.';
        return;
      }
      // GAS URL not configured — use offline fallback
      const offline = OFFLINE_USERS.find(x => x.id === u && x.pw === p);
      if(offline){
        _completeLogin(offline.name, offline.role, offline.id, true, offline.allowedSections);
      } else {
        err.innerText = 'Server not configured and credentials not in offline list.';
      }
    }
  );
}

/** Shared post-login setup — called whether online or offline */
function _completeLogin(name, role, userId, isOffline, allowedSections){
  userName     = name;
  userRole     = (role || 'OPERATOR').toUpperCase();
  userIdGlobal = userId;
  userSections = parseAllowedSections_(allowedSections, userRole);

  // Cache last successful login so offline works on re-open
  lsSet('lastUser', { name, role: userRole, userId, allowedSections: userSections.join(','), ts: Date.now() });

  // Apply admin class to body — shows/hides admin-only elements
  if(userRole === 'ADMIN'){
    document.body.classList.add('is-admin');
  } else {
    document.body.classList.remove('is-admin');
  }

  document.getElementById('loginName').innerText = name;
  document.getElementById('avatarEl').innerText  = name.charAt(0).toUpperCase();
  applySectionAccess();

  // Show offline warning if needed
  if(isOffline){
    showOfflineBanner(true);
    toast('⚠ Working offline — sync unavailable until connection restored', 'warn', 6000);
  } else {
    showOfflineBanner(false);
  }

  document.getElementById('loginOverlay').style.display = 'none';
  document.getElementById('appShell').style.display     = 'block';

  if(inventory.length > 0) dumpRows = inventory;
  initGRNSection();
  loadDashboard();
  if(_syncPollTimer) clearInterval(_syncPollTimer);
  _syncPollTimer = setInterval(pollSyncStatus, 10*60*1000);
  setTimeout(pollSyncStatus, 2000);
  // Gatepass dump is large; load it only when diff check needs it.

  // Load master data from sheet
  callSheet('loadAllWMSData', {},
    function(data){
      if(!data || !data.success) return;
      if(data.skuMaster  && data.skuMaster.length)  { skuMaster = data.skuMaster;   saveAll(); renderSKUTable(); document.getElementById('db-sku').innerText = skuMaster.length; }
      if(data.binMaster  && data.binMaster.length)  { binMaster = normalizeBinMasterRows(data.binMaster);   saveAll(); loadDashboard(); }
      if(data.binSummary) applyDashboardBinSummary(data.binSummary);
      if(data.gatepasses && data.gatepasses.length) { gatepasses = data.gatepasses; saveAll(); renderGatepassSummary(); }
    },
    function(){}
  );
}

function applyDashboardBinSummary(summary){
  if(!summary) return;
  const localTotal = normalizeBinMasterRows(binMaster).length;
  if(localTotal && Number(summary.total || 0) < localTotal) return;
  const set = (id, val) => {
    const el = document.getElementById(id);
    if(el) el.innerText = Number(val || 0).toLocaleString();
  };
  set('db-total', summary.total);
  set('db-empty', summary.empty);
  set('db-occ', summary.occupied);
  set('db-res', summary.reserved);
}

function parseDashboardBinId_(binId){
  const m = String(binId||'').match(/^R(\d+)-C(\d+)-(\d+)$/i);
  if(!m) return null;
  return { row:Number(m[1]), col:Number(m[2]), level:Number(m[3]) };
}

function renderBinUtilizationMap(bins){
  const grid=document.getElementById('binUtilGrid');
  if(!grid) return;
  const rows={};
  (bins||[]).forEach(b=>{
    const p=parseDashboardBinId_(b.id);
    if(!p) return;
    if(!rows[p.row]) rows[p.row]=[];
    rows[p.row].push({...b,_p:p});
  });
  const rowNos=Object.keys(rows).map(Number).sort((a,b)=>a-b);
  if(!rowNos.length){ grid.innerHTML='<div style="font-size:12px;color:var(--slate4);padding:8px">No valid rack bins found</div>'; return; }
  grid.className='bin-map';
  grid.innerHTML=rowNos.map(r=>{
    const line=rows[r].sort((a,b)=>a._p.col-b._p.col||a._p.level-b._p.level);
    const occ=line.filter(b=>b.status==='Occupied').length;
    const res=line.filter(b=>b.status==='Reserved').length;
    const empty=line.filter(b=>b.status==='Empty').length;
    return '<div class="bin-line">'
      +'<div class="bin-line-head"><span class="bin-line-title">R'+r+'</span><span>'+line.length+' bins</span><span>'+empty+' empty · '+occ+' occ · '+res+' res</span></div>'
      +'<div class="bin-line-grid">'
      +line.map(b=>'<div class="bin-cell bc-'+String(b.status||'Empty').toLowerCase()+'" title="'+b.id+' — '+b.status+' ('+(b.fsn||'')+')"><span class="bin-cell-id">C'+b._p.col+'-'+String(b._p.level).padStart(3,'0')+'</span></div>').join('')
      +'</div></div>';
  }).join('');
}

// v6: Track last active section to avoid re-rendering on same tab click
let _lastSection = '';
function showSection(name,el){
  const group = SECTION_GROUP_BY_VIEW[name] || 'OVERVIEW';
  if(!canAccessSection_(group)){
    toast('You do not have access to this section. Ask Admin to update USERS → Allowed Sections.', 'warn', 4000);
    return false;
  }
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById('s-'+name)?.classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  if(el) el.classList.add('active');
  _lastSection = name;
  return true;
}

/* ════════════════════ DASHBOARD ════════════════════ */
function loadDashboard(){
  const cleanBins=normalizeBinMasterRows(binMaster);
  if(cleanBins.length!==binMaster.length){ binMaster=cleanBins; lsSet('bin', binMaster); }
  // v8: gracefully handle empty arrays (post-flush state)
  const total=cleanBins.length;
  const empty=cleanBins.filter(b=>b.status==='Empty').length;
  const occ=cleanBins.filter(b=>b.status==='Occupied').length;
  const res=cleanBins.filter(b=>b.status==='Reserved').length;
  document.getElementById('db-total').innerText=total||'—';
  document.getElementById('db-empty').innerText=empty||'—';
  document.getElementById('db-occ').innerText=occ||'—';
  document.getElementById('db-res').innerText=res||'—';
  document.getElementById('db-sku').innerText=skuMaster.length||'—';
  document.getElementById('db-grns').innerText=grnList.length||'—';
  const gpTotal=gatepasses.reduce((s,r)=>s+(r.qty||0),0);
  document.getElementById('db-gp').innerText=gatepasses.length>0?gatepasses.length+' lines ('+gpTotal.toLocaleString()+' units)':'—';

  const pct=n=>total>0?Math.round(n/total*100)+'%':'0%';
  document.getElementById('occ-pct').innerText=pct(occ);
  document.getElementById('res-pct').innerText=pct(res);
  document.getElementById('avail-pct').innerText=pct(empty);
  document.getElementById('occ-bar').style.width=pct(occ);
  document.getElementById('res-bar').style.width=pct(res);
  document.getElementById('avail-bar').style.width=pct(empty);

  // Class breakdown
  const cg={A:0,B:0,C:0};
  cleanBins.forEach(b=>{if(cg[b.fsn]!==undefined)cg[b.fsn]++;});
  document.getElementById('classBreakdown').innerHTML=
    Object.entries(cg).map(([k,v])=>`<span class="status-badge sb-${k.toLowerCase()}">${k}: ${v} bins</span>`).join('');

  // Bin utilization map grouped by rack line
  renderBinUtilizationMap(cleanBins);

  // Inventory level cards
  const totalInv=inventory.reduce((s,r)=>s+r.qty,0);
  const goodInv=inventory.filter(r=>['GOOD','UNRESTRICTED'].includes((r.stockType||'GOOD').toUpperCase())).reduce((s,r)=>s+r.qty,0);
  const dmgInv=inventory.filter(r=>['DAMAGE','DAMAGED'].includes((r.stockType||'').toUpperCase())).reduce((s,r)=>s+r.qty,0);
  const activeSKUs=new Set(inventory.map(r=>r.skuCode)).size;
  document.getElementById('db-invTotal').innerText=totalInv.toLocaleString();
  document.getElementById('db-invGood').innerText=goodInv.toLocaleString();
  document.getElementById('db-invDamage').innerText=dmgInv.toLocaleString();
  document.getElementById('db-invSKUs').innerText=activeSKUs;

  renderDashboardInventory();
  renderRecentGRN();
}

function renderDashboardInventory(){
  const q=String(document.getElementById('dbInvSearch')?.value||'').toLowerCase();
  const skuMap={};
  const _dbToday=new Date(); _dbToday.setHours(0,0,0,0);
  function _dbParseExp(s){
    if(!s)return null;
    const str=String(s).trim();
    const m=str.match(/^(\d\d)-(\d\d)-(\d{4})$/);
    if(m)return new Date(+m[3],+m[2]-1,+m[1]);
    const d=new Date(str); return isNaN(d.getTime())?null:d;
  }
  inventory.forEach(r=>{
    const key=r.skuCode;
    if(!skuMap[key]) skuMap[key]={skuCode:r.skuCode,skuName:r.skuName||'',good:0,damage:0,bins:new Set(),minDays:null};
    const t=(r.stockType||'GOOD').toUpperCase();
    if(['GOOD','UNRESTRICTED'].includes(t)) skuMap[key].good+=r.qty;
    else if(['DAMAGE','DAMAGED'].includes(t)) skuMap[key].damage+=r.qty;
    if(r.bin) skuMap[key].bins.add(r.bin);
    const exp=_dbParseExp(r.expDate);
    if(exp){
      const days=Math.round((exp-_dbToday)/(1000*60*60*24));
      if(skuMap[key].minDays===null||days<skuMap[key].minDays) skuMap[key].minDays=days;
    }
  });
  const rows=Object.values(skuMap).filter(r=>`${r.skuCode} ${r.skuName}`.toLowerCase().includes(q));
  const body=document.getElementById('dbInvBody');
  if(!rows.length){body.innerHTML='<tr class="empty-row"><td colspan="7">Upload inventory dump to see stock levels</td></tr>';return;}
  body.innerHTML=rows.map(r=>{
    var exBadge='<span style="font-size:10px;color:var(--slate5)">—</span>';
    if(r.minDays!==null){
      if(r.minDays<0)        exBadge='<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:var(--red-l);color:var(--red-d)">🔴 Expired</span>';
      else if(r.minDays<=90) exBadge='<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:var(--amber-l);color:var(--amber-d)">🟡 Near Expiry ('+r.minDays+'d)</span>';
      else                   exBadge='<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:var(--green-l);color:var(--green-d)">🟢 Active</span>';
    }
    return '<tr>'
      +'<td><span class="sku-badge">'+r.skuCode+'</span></td>'
      +'<td style="font-size:12px">'+r.skuName+'</td>'
      +'<td style="text-align:right;font-weight:700;color:var(--green)">'+r.good.toLocaleString()+'</td>'
      +'<td style="text-align:right;font-weight:700;color:var(--red)">'+r.damage.toLocaleString()+'</td>'
      +'<td style="text-align:right;font-weight:700">'+(r.good+r.damage).toLocaleString()+'</td>'
      +'<td style="font-size:11px">'+[...r.bins].slice(0,3).map(b=>'<span class="bin-badge" style="margin:1px 2px;font-size:9px">'+b+'</span>').join('')+(r.bins.size>3?' <span style="color:var(--slate4);font-size:10px">+'+( r.bins.size-3)+' more</span>':'')+'</td>'
      +'<td>'+exBadge+'</td>'
      +'</tr>';
  }).join('');
}

function renderRecentGRN(){
  const el=document.getElementById('recentGRN');
  if(!grnList.length){el.innerHTML='<p style="padding:20px;color:var(--slate5);font-size:12px">No GRN entries yet.</p>';return;}
  const recent=grnList.slice(-5).reverse();
  el.innerHTML=`<table style="width:100%"><thead><tr><th>GRN No</th><th>Date</th><th>Supplier</th><th style="text-align:center">Lines</th><th style="text-align:right">Total Qty</th></tr></thead><tbody>${
    recent.map(g=>`<tr><td><code style="font-family:'DM Mono',monospace;font-size:11px">${g.grnNo}</code></td><td>${g.date}</td><td>${g.supplier||'—'}</td><td style="text-align:center">${g.lines?.length||0}</td><td style="text-align:right;font-weight:700">${(g.lines||[]).reduce((s,l)=>s+Number(l.qty||0),0).toLocaleString()}</td></tr>`).join('')
  }</tbody></table>`;
}

/* ════════════════════ GRN MODE ════════════════════ */
