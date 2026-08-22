function toast(msg, type, duration){
  type = type || 'info';
  duration = duration || 3500;
  const container = document.getElementById('toastContainer');
  if(!container) return;
  const el = document.createElement('div');
  el.className = 'toast t-' + type;
  const icons = { success:'✅', warn:'⚠️', error:'❌', info:'ℹ️' };
  el.innerHTML = `<span>${icons[type]||'ℹ️'}</span><span>${msg}</span>`;
  container.appendChild(el);
  requestAnimationFrame(()=>{ el.classList.add('show'); });
  setTimeout(()=>{
    el.classList.remove('show');
    setTimeout(()=>el.remove(), 300);
  }, duration);
}

/* ════════════════════ v7: GLOBAL REFRESH ════════════════════ */
/**
 * Reload ALL data from Google Sheet in one server call.
 * Updates: skuMaster, binMaster, gatepasses, inventory (dumpRows).
 * Re-renders all active views.
 */
function globalRefresh(){
  const btn  = document.getElementById('globalRefreshBtn');
  const icon = document.getElementById('refreshIcon');
  if(btn)  { btn.disabled = true; }
  if(icon) { icon.classList.add('refresh-spin'); }
  toast('Refreshing data from Google Sheet…', 'info', 8000);

  callSheet('refreshAllData', {},
    function(data){
      if(btn)  { btn.disabled = false; }
      if(icon) { icon.classList.remove('refresh-spin'); }
      if(!data || !data.success){
        toast('Refresh failed: ' + (data&&data.message||'unknown'), 'error');
        return;
      }
      // v8: ALWAYS write to localStorage — even empty arrays
      // This prevents stale local data from overriding a freshly emptied sheet
      const updated = [];
      skuMaster  = data.skuMaster  || [];  lsSet('sku', skuMaster);
      binMaster  = normalizeBinMasterRows(data.binMaster || []);  lsSet('bin', binMaster);
      gatepasses = data.gatepasses || [];  lsSet('gp',  gatepasses);
      // v9: normalize stockType for each inventory row on refresh
      inventory  = (data.inventory || []).map(r=>({...r, bin: normalizeBinIdForMaster(r.bin), stockType: normalizeStockType_(r.stockType)}));
      lsSet('inv', inventory);
      dumpRows   = inventory;
      _ciFlushMode = false; // fresh data — CI can rebuild

      // Mark this as loaded so pollSyncStatus doesn't double-refresh
      callSheet('getLastSyncStatus', {},
        function(s){ if(s && (s.syncId||s.timestamp)) _lastLoadedSyncId = s.syncId||s.timestamp; },
        function(){}
      );

      if(data.skuMaster  && data.skuMaster.length)  updated.push('SKU ('+data.skuMaster.length+')');
      if(binMaster.length)  updated.push('Bins ('+binMaster.length+')');
      if(data.gatepasses && data.gatepasses.length) updated.push('Gatepasses ('+data.gatepasses.length+')');
      if(data.inventory  && data.inventory.length)  updated.push('Inventory ('+data.inventory.length+')');
      // Re-render the currently visible section
      const active = document.querySelector('.section.active');
      if(active){
        const id = active.id.replace('s-','');
        _reRenderTab_(id, true /* silent */);
      }
      loadDashboard();
      renderCurrentInventory();
      toast('✅ Refreshed: ' + (updated.join(', ')||'no changes') +
            (data.refreshedAt?' at '+data.refreshedAt:''), 'success', 4000);
    },
    function(err){
      if(btn)  { btn.disabled = false; }
      if(icon) { icon.classList.remove('refresh-spin'); }
      toast('Refresh error: ' + err, 'error');
    }
  );
}

/* ════════════════════ v7: PER-TAB REFRESH ════════════════════ */
/**
 * Refresh the data and re-render a specific tab.
 * Does NOT make a server call — works from in-memory state.
 * For a full server reload, use globalRefresh().
 */
/**
 * refreshTab — re-renders the selected tab from already-loaded data.
 * Use the top-bar Refresh button for a full Google Sheet reload.
 */
function refreshTab(tabName, silent){
  _reRenderTab_(tabName, silent);
}

function _reRenderTab_(tabName, silent){
  switch(tabName){
    case 'skumaster':        renderSKUTable();                              break;
    case 'bins':             renderBinTable(); loadDashboard();             break;
    case 'dump':             renderDumpSummary(); renderDumpTable();        break;
    case 'gatepass':         renderGatepassSummary(); renderGatepassTable();break;
    case 'cyclecount':       initCycleCount();                              break;
    case 'adjustment':       renderAdjLog(); updateAdjStats();              break;
    case 'currentinventory': renderCurrentInventory();                      break;
    case 'dashboard':        loadDashboard();                               break;
    case 'grnList':          renderGRNList();                               break;
    default: break;
  }
  if(!silent) toast('✅ ' + tabName + ' refreshed', 'success', 2000);
}

/* ════════════════════ v7: FLUSH DATA ════════════════════ */
/**
 * Flush (delete all data rows from) a Google Sheet.
 * sheetName: exact sheet name in Google Sheets
 * tabName:   the local tab to refresh after flush
 */
/**
 * v8 FIX — flushData now clears BOTH:
 *   1. Google Sheet (server call)
 *   2. localStorage (mwwms2_* keys)
 *   3. In-memory JS variables
 * So UI reflects the flush immediately and doesn't reload stale data.
 *
 * sheetName  — exact Google Sheet tab name
 * tabName    — UI section id to re-render after flush
 */
function flushData(sheetName, tabName){
  // ── Admin check ──
  if(userRole !== 'ADMIN'){
    toast('⛔ Flush is restricted to Admin users only', 'error', 4000);
    return;
  }

  const friendlyNames = {
    SKU_MASTER:       'SKU Master',
    BIN_MASTER_WMS:   'Bin Master',
    INVENTORY_DUMP:   'Inventory Dump',
    GATEPASS:         'Gatepass',
    CYCLE_COUNT_LOG:  'Cycle Count Log',
    ADJUSTMENT_LOG:   'Adjustment Log',
    CURRENT_INVENTORY:'Current Inventory'
  };
  const name = friendlyNames[sheetName] || sheetName;

  // ── Type-to-confirm (admin must type FLUSH) ──
  const typed = prompt(
    'ADMIN ACTION: Delete ALL rows from [' + name + '] sheet.\n\nType FLUSH to confirm (cannot be undone):'
  );
  if(typed === null) return;                          // cancelled
  if(typed.trim().toUpperCase() !== 'FLUSH'){
    toast('Flush cancelled — you must type FLUSH exactly', 'warn', 3000);
    return;
  }

  const localMap = {
    SKU_MASTER:        { reset: ()=>{ skuMaster  = [];           lsSet('sku', []); } },
    BIN_MASTER_WMS:    { reset: ()=>{ binMaster  = [];           lsSet('bin', []); } },
    INVENTORY_DUMP:    { reset: ()=>{ inventory  = []; dumpRows=[]; lsSet('inv', []); } },
    GATEPASS:          { reset: ()=>{ gatepasses = [];           lsSet('gp',  []); } },
    CYCLE_COUNT_LOG:   { reset: ()=>{ ccLog      = [];           lsSet('ccLog', []); } },
    ADJUSTMENT_LOG:    { reset: ()=>{ adjLog     = [];           saveAdjLog(); } },
    CURRENT_INVENTORY: { reset: ()=>{ currentInventory = []; _ciFlushMode = true; } }
  };

  // Clear local state immediately
  if(localMap[sheetName]) {
    try { localMap[sheetName].reset(); } catch(e) { console.warn(e); }
  }
  _reRenderAfterFlush(sheetName, tabName);
  toast('Flushing ' + name + '…', 'info', 6000);

  // Clear Google Sheet
  callSheet('clearSheetData', { sheetName },
    function(r){
      if(r && r.success){
        toast('✅ ' + name + ' flushed — ' + (r.count||0) + ' rows deleted', 'success', 4000);
      } else {
        toast('⚠ Local cleared but sheet error: ' + (r&&r.message||'unknown'), 'warn', 5000);
      }
    },
    function(err){
      toast('⚠ Local cleared. Sheet sync failed: ' + err, 'warn', 6000);
    }
  );
}

/** Re-render whichever section was flushed */
function _reRenderAfterFlush(sheetName, tabName){
  switch(sheetName){
    case 'SKU_MASTER':
      renderSKUTable();
      document.getElementById('db-sku').innerText = '0';
      break;
    case 'BIN_MASTER_WMS':
      renderBinTable();
      loadDashboard();
      break;
    case 'INVENTORY_DUMP':
      dumpRows = [];
      if(document.getElementById('dumpSummary')) document.getElementById('dumpSummary').innerHTML = '';
      if(document.getElementById('dumpMsg'))     document.getElementById('dumpMsg').innerText = 'Data flushed.';
      renderDumpTable();
      renderCurrentInventory();
      loadDashboard();
      break;
    case 'GATEPASS':
      renderGatepassSummary();
      renderGatepassTable();
      if(document.getElementById('gpSummary')) document.getElementById('gpSummary').innerHTML = '';
      loadDashboard();
      break;
    case 'CYCLE_COUNT_LOG':
      updateCCStats();
      renderCCLog();
      break;
    case 'ADJUSTMENT_LOG':
      renderAdjLog();
      updateAdjStats();
      break;
    case 'CURRENT_INVENTORY':
      // _ciFlushMode is already set to true by localMap reset
      // renderCurrentInventory will show empty state without rebuilding
      renderCurrentInventory();
      break;
  }
}

/* ════════════════════ CONNECTIVITY ════════════════════ */

function showOfflineBanner(show){
  _isOnline = !show;
  const banner = document.getElementById('offlineBanner');
  if(banner) banner.classList.toggle('show', show);
  const st = document.getElementById('loginStatus');
  if(st && document.getElementById('loginOverlay') &&
     document.getElementById('loginOverlay').style.display !== 'none'){
    st.className = 'login-status ' + (show ? 'offline' : 'online');
    st.innerText = show ? '⚠ Server unavailable — check connection'
                        : '✅ Connected to Mosaic Wellness WMS';
  }
}

/** Ping GAS — only shows offline banner after 3 consecutive failures */
function pingServer(){
  callSheet('ping', {},
    function(){
      _pingFailCount = 0;          // reset on success
      showOfflineBanner(false);
    },
    function(){
      _pingFailCount++;
      if(_pingFailCount >= PING_FAIL_THRESHOLD){
        showOfflineBanner(true);   // only show after 3 fails
      }
      // else: transient blip — don't alarm user
    }
  );
}

// ── Restore session if user already logged in (page reload) ──
(function restoreSession(){
  const cached = ls('lastUser');
  if(!cached) return;
  // Only auto-restore if logged in within last 8 hours
  if(Date.now() - cached.ts > 8 * 60 * 60 * 1000) return;
  // Don't auto-restore — always require explicit login for security
  // But pre-fill the username field as convenience
  const uField = document.getElementById('loginUser');
  if(uField && cached.userId) uField.value = cached.userId;
})();

/* ════════════════════ GAS URL SETTINGS ════════════════════ */
function saveGASUrl(){
  const url=document.getElementById('gasUrlInput').value.trim();
  if(!url){document.getElementById('gasStatus').innerText='⚠ Paste the URL first';return;}
  if(!url.includes('script.google.com')){document.getElementById('gasStatus').innerText='⚠ Must be a script.google.com URL';return;}
  if(!url.endsWith('/exec')){document.getElementById('gasStatus').innerText='⚠ URL must end with /exec (not /dev)';return;}
  lsSet('gasUrl',url);
  document.getElementById('gasStatus').innerText='✅ Saved! Testing connection…';
  setTimeout(testGASConnection,500);
}
async function testGASConnection(){
  const url=(document.getElementById('gasUrlInput').value.trim()||ls('gasUrl')||'');
  if(!url){document.getElementById('gasStatus').innerText='⚠ No URL saved yet';return;}
  document.getElementById('gasStatus').innerText='Testing…';
  try{
    const res=await fetch(url+'?action=ping',{method:'GET'});
    const text=await res.text();
    if(text.includes('pong')||text.includes('success')||res.ok){
      document.getElementById('gasStatus').innerText='✅ Connected! Data will now sync to Google Sheet.';
    } else {
      document.getElementById('gasStatus').innerText='⚠ Connected but unexpected response. Sync may still work.';
    }
  }catch(e){
    document.getElementById('gasStatus').innerText='❌ Cannot reach URL: '+e.message+'. Check URL and make sure deployment is set to "Anyone".';
  }
}
window.addEventListener('DOMContentLoaded',function(){
  const saved=ls('gasUrl')||GAS_URL||'';
  const input=document.getElementById('gasUrlInput');
  if(input)input.value=saved;
  if(saved){
    const btn=document.querySelector('button[title^="Connect to Google Sheet"]');
    if(btn){btn.style.borderColor='var(--green)';btn.style.color='var(--green)';btn.innerText='✅ Sheet';}
  }
  const st = document.getElementById('loginStatus');
  if(st) st.innerText = saved ? 'Ready — sign in to continue' : '⚠ Sheet URL not set';
  if(saved) setTimeout(function(){
    pingServer();
    setInterval(pingServer, 300000);
  }, 4000);
});

/* ════════════════════ INIT ════════════════════ */
window.onload=function(){
  if(inventory.length>0){dumpRows=inventory;}
  if(gatepasses.length>0)renderGatepassSummary();
  renderGRNTable();
};

/* ════════════════════════════════════════════════════════
   UNIWARE SYNC — Frontend
   ════════════════════════════════════════════════════════ */
let _syncPollTimer       = null;
let _lastLoadedSyncId   = '';   // msgid of the last sync we loaded — reliable change detection

/* ════════════════════════════════════════════════════════
   GATEPASS SYNC — Frontend
   ════════════════════════════════════════════════════════ */

/** Sync gatepasses from latest email */
function runSyncGatepass(){
  const res = document.getElementById('syncActionResult');
  if(res) res.innerHTML = '<span style="color:var(--amber)">⏳ Fetching Gatepass report…</span>';
  callSheet('syncGatepasses', {},
    function(data){
      if(data && data.success){
        if(res) res.innerHTML = '<span style="color:var(--green)">✅ Gatepasses synced — '+data.rowCount+' rows (CREATED/RETURN_AWAITED from SL Mother Hub)</span>';
        toast('Gatepasses synced — '+data.rowCount+' rows','success',4000);
        // Load into memory for diff check
        loadGatepasses();
      } else {
        if(res) res.innerHTML = '<span style="color:var(--amber)">⚠ '+(data&&data.message||'No new gatepass email')+'</span>';
      }
    },
    function(err){ if(res) res.innerHTML='<span style="color:var(--red)">❌ '+err+'</span>'; }
  );
}

/** Force re-import gatepass (ignores processed check) */
function runForceReimportGatepass(){
  const res = document.getElementById('syncActionResult');
  if(res) res.innerHTML = '<span style="color:var(--amber)">⏳ Force-fetching Gatepass…</span>';
  callSheet('forceReimportGatepass', {},
    function(data){
      if(data && data.success){
        if(res) res.innerHTML = '<span style="color:var(--green)">✅ '+data.message+'</span>';
        toast(data.message,'success',4000);
        loadGatepasses();
      } else {
        if(res) res.innerHTML = '<span style="color:var(--red)">❌ '+(data&&data.message||'Failed')+'</span>';
      }
    },
    function(err){ if(res) res.innerHTML='<span style="color:var(--red)">❌ '+err+'</span>'; }
  );
}

/** Load gatepasses from sheet into memory for diff check */
function loadGatepasses(done){
  callSheet('getGatepasses', {},
    function(data){
      if(data && data.success){
        gatewayPasses = data.rows || [];
        _gatewayPassesLoaded = true;
        // Refresh diff check if open
        if(diffRows.length) renderDiffTable();
      }
      if(done) done();
    },
    function(){ _gatewayPassesLoaded = true; if(done) done(); }
  );
}

/**
 * Get total pending gatepass qty for a specific bin+sku+batch+stockType.
 * Used by diff check to show net variance.
 */
function getGpPendingQty(bin, skuCode, batch, stockType){
  return gatewayPasses
    .filter(g =>
      g.bin       === bin &&
      g.skuCode   === skuCode &&
      (g.batch    === batch || !batch) &&
      normalizeStockType_(g.stockType) === normalizeStockType_(stockType)
    )
    .reduce((sum, g) => sum + (g.qty||0), 0);
}

function pollSyncStatus(){
  callSheet('getLastSyncStatus', {},
    function(data){
      if(!data || !data.success) return;
      _updateSyncBadge(data);

      // Auto-refresh inventory if a newer successful sync has happened
      const st = (data.status||'').toUpperCase();
      const sid = data.syncId || data.timestamp || '';
      if(st === 'SUCCESS' && sid && sid !== _lastLoadedSyncId){
        _lastLoadedSyncId = sid;   // mark as loaded immediately to prevent double-refresh
        setTimeout(function(){
          _silentRefreshInventory(data.rowCount);
        }, 800);
      }
    },
    function(){ /* GAS not connected — badge stays grey */ }
  );
}

/**
 * Pull fresh inventory from sheet WITHOUT spinning the topbar button.
 * Called automatically after background syncs are detected.
 */
function _silentRefreshInventory(rowCount){
  callSheet('refreshAllData', {},
    function(data){
      if(!data || !data.success) return;

      // Update all in-memory state + localStorage
      skuMaster  = data.skuMaster  || [];  lsSet('sku', skuMaster);
      binMaster  = normalizeBinMasterRows(data.binMaster || []);  lsSet('bin', binMaster);
      gatepasses = data.gatepasses || [];  lsSet('gp',  gatepasses);
      inventory  = (data.inventory || []).map(r=>({...r, bin: normalizeBinIdForMaster(r.bin), stockType: normalizeStockType_(r.stockType)}));
      lsSet('inv', inventory);
      dumpRows      = inventory;
      _ciFlushMode  = false;

      // Re-render every visible section
      loadDashboard();
      renderDumpSummary();
      renderDumpTable();
      renderCurrentInventory();
      renderBinTable();
      updateBinsFromDump();
      renderSKUTable();

      // Re-init cycle count bin/sku lists so fresh bins are searchable
      const ccBinList = document.getElementById('cc_bin_list');
      const ccSkuList = document.getElementById('cc_sku_list');
      if(ccBinList){
        const bins=[...new Set(inventory.map(r=>r.bin).filter(Boolean))].sort();
        ccBinList.innerHTML = bins.map(b=>'<option value="'+b+'">').join('');
      }
      if(ccSkuList){
        const skus=[...new Set(inventory.map(r=>r.skuCode).filter(Boolean))].sort();
        ccSkuList.innerHTML = skus.map(s=>'<option value="'+s+'">').join('');
      }

      toast('Inventory auto-updated from Uniware — '+(rowCount||inventory.length).toLocaleString()+' rows', 'success', 4000);
    },
    function(){ /* silent fail — badge will retry next poll */ }
  );
}

function _updateSyncBadge(data){
  const dot   = document.getElementById('syncDot');
  const label = document.getElementById('syncLabel');
  if(!dot || !label) return;

  const st = (data.status||'NEVER').toUpperCase();
  const dotClass = {
    SUCCESS:'s-success', FAILED:'s-failed', ERROR:'s-failed',
    SKIPPED:'s-skipped', NEVER:'s-never', RUNNING:'s-running'
  }[st] || 's-never';
  dot.className = 'sync-dot ' + dotClass;

  if(st === 'SUCCESS' && data.timestamp){
    const t = data.timestamp.split(' ')[1] || data.timestamp;
    label.innerHTML = 'Synced <b style="color:#fff">' + t + '</b> <span style="color:#f1f5f9">· ' + (data.rowCount||0).toLocaleString() + ' rows</span>';
  } else if(st === 'FAILED' || st === 'ERROR'){
    label.innerHTML = '<b style="color:#fca5a5;font-weight:700">⚠ Sync Failed</b>';
  } else if(st === 'SKIPPED'){
    const skipTime = (data.timestamp||'').split(' ')[1] || '';
    label.innerHTML = '<span style="color:#f1f5f9">No new report</span>' +
                      (skipTime ? ' <b style="color:#fff">' + skipTime + '</b>' : '');
  } else {
    label.innerHTML = 'Uniware · not synced';
  }

  // Update modal cards
  const smS = document.getElementById('sm-status');
  const smR = document.getElementById('sm-rows');
  const smT = document.getElementById('sm-time');
  if(smS){ smS.innerText = st; smS.style.color = {SUCCESS:'var(--green)',FAILED:'var(--red)',ERROR:'var(--red)',SKIPPED:'var(--amber)'}[st]||'var(--slate4)'; }
  if(smR) smR.innerText = (data.rowCount||0).toLocaleString() + ' rows';
  if(smT) smT.innerText = data.timestamp || 'Never';
}

function openSyncModal(){
  document.getElementById('syncModal').classList.add('open');
  pollSyncStatus();
}

function runManualSync(){
  const btn = document.getElementById('manualSyncBtn');
  const res = document.getElementById('syncActionResult');
  if(btn){ btn.disabled=true; btn.innerText='Syncing…'; }
  if(res) res.innerText = '';

  const dot = document.getElementById('syncDot');
  if(dot) dot.className = 'sync-dot s-running';
  const lbl = document.getElementById('syncLabel');
  if(lbl) lbl.innerHTML = 'Syncing…';

  callSheet('manualUniwareSync', {},
    function(data){
      if(btn){ btn.disabled=false; btn.innerHTML='<svg viewBox="0 0 24 24"><path d="M1 4v6h6"/><path d="M23 20v-6h-6"/><path d="M20.49 9A9 9 0 005.64 5.64L1 10M23 14l-4.64 4.36A9 9 0 013.51 15"/></svg>Sync Now'; }
      if(!data){ if(res) res.innerHTML='<span style="color:var(--red)">No response from server</span>'; return; }
      _updateSyncBadge(data);
      const st = (data.status||'').toUpperCase();
      if(st === 'SUCCESS'){
        // Show result in modal — keep modal open, show timestamp
        var ts = data.timestamp || data.syncedAt || new Date().toLocaleString('en-IN');
        var rowCount = (data.rowCount||0).toLocaleString();
        if(res) res.innerHTML = '<div style="background:var(--green-l);border:1px solid var(--green);border-radius:8px;padding:12px 14px;margin-top:12px">'
          + '<div style="font-size:14px;font-weight:700;color:var(--green-d);margin-bottom:4px">✅ Re-imported ' + rowCount + ' rows from ' + ts + '. Refreshing…</div>'
          + '<div style="font-size:12px;color:var(--green-d);opacity:.8">Inventory data is being updated. This may take a few seconds.</div>'
          + '</div>';
        _lastLoadedSyncId = data.syncId || data.timestamp || '';
        // Refresh data in background — modal stays open
        setTimeout(function(){
          globalRefresh();
          toast('✅ Inventory refreshed — ' + rowCount + ' rows from Uniware', 'success', 5000);
        }, 1500);
      } else if(st === 'SKIPPED'){
        if(res) res.innerHTML = '<div style="background:var(--amber-l);border:1px solid var(--amber);border-radius:8px;padding:12px 14px;margin-top:12px">'
          + '<div style="font-size:13px;font-weight:700;color:var(--amber-d)">⚠ No new Shelfwise email found</div>'
          + '<div style="font-size:12px;color:var(--amber-d);margin-top:4px">Inventory unchanged. Last import was already the latest available.</div>'
          + '</div>';
      } else {
        if(res) res.innerHTML = '<div style="background:var(--red-l);border:1px solid var(--red);border-radius:8px;padding:12px 14px;margin-top:12px">'
          + '<div style="font-size:13px;font-weight:700;color:var(--red-d)">❌ Sync failed</div>'
          + '<div style="font-size:12px;color:var(--red-d);margin-top:4px">' + (data.detail||data.message||'Unknown error') + '</div>'
          + '</div>';
      }
    },
    function(err){
      if(btn){ btn.disabled=false; btn.innerText='Sync Now'; }
      if(res) res.innerHTML='<span style="color:var(--red)">❌ '+err+'</span>';
    }
  );
}

function runForceReimport(){
  const res = document.getElementById('syncActionResult');
  if(!confirm('Force re-import from the most recent Shelfwise email?\n\nThis will REPLACE the current INVENTORY_DUMP regardless of whether it was already imported.\n\nUse this after flushing inventory to restore data.')) return;
  if(res) res.innerHTML = '<span style="color:var(--amber)">⏳ Re-importing from latest email…</span>';
  const dot = document.getElementById('syncDot');
  if(dot) dot.className = 'sync-dot s-running';
  callSheet('forceReimportLatest', {},
    function(data){
      if(data && data.success){
        if(res) res.innerHTML = '<span style="color:var(--green)">✅ '+data.message+'. Refreshing…</span>';
        _lastLoadedSyncId = ''; // force next poll to reload
        setTimeout(function(){ globalRefresh(); toast('Inventory re-imported — '+data.rowCount+' rows','success',5000); }, 1500);
      } else {
        if(res) res.innerHTML = '<span style="color:var(--red)">❌ '+(data&&data.message||'Failed')+'</span>';
      }
    },
    function(err){ if(res) res.innerHTML='<span style="color:var(--red)">❌ '+err+'</span>'; }
  );
}

function runRestoreBackup(){
  const res = document.getElementById('syncActionResult');
  if(!confirm('Restore INVENTORY_DUMP from previous backup?\n\nThis replaces current data with the snapshot taken before the last sync.\nUse only if the last import was incorrect.')) return;
  callSheet('restorePreviousDump', {},
    function(data){
      if(data && data.success){
        if(res) res.innerHTML='<span style="color:var(--green)">✅ '+data.message+'. Refreshing…</span>';
        setTimeout(globalRefresh, 1500);
        toast('Previous dump restored','warn',4000);
      } else {
        if(res) res.innerHTML='<span style="color:var(--red)">❌ '+(data&&data.message||'Restore failed')+'</span>';
      }
    },
    function(err){ if(res) res.innerHTML='<span style="color:var(--red)">❌ '+err+'</span>'; }
  );
}



/* ══ Paginator ══ */
function createPager(getRows, pageSize, pagerId, renderFn) {
  var page = 0;
  var p = {
    render: function() {
      var all = getRows(); var total = all.length;
      var pages = Math.max(1, Math.ceil(total / pageSize));
      if (page >= pages) page = pages - 1;
      renderFn(all.slice(page * pageSize, (page + 1) * pageSize));
      var el = document.getElementById(pagerId);
      if (!el) return;
      var from = total === 0 ? 0 : page * pageSize + 1;
      var to = Math.min((page + 1) * pageSize, total);
      el.innerHTML = '<span class="pager-info">Showing ' + from + '–' + to + ' of ' + total.toLocaleString() + ' rows</span>'
        + '<button class="pager-btn" id="' + pagerId + '-p"' + (page===0?' disabled':'') + '>← Prev</button>'
        + '<button class="pager-btn" id="' + pagerId + '-n"' + (page>=pages-1?' disabled':'') + '>Next →</button>';
      var pb=document.getElementById(pagerId+'-p'), nb=document.getElementById(pagerId+'-n');
      if(pb) pb.onclick=function(){if(page>0){page--;p.render();}};
      if(nb) nb.onclick=function(){if(page<pages-1){page++;p.render();}};
    },
    refresh: function() { page=0; p.render(); }
  };
  return p;
}

/* ══ Rate limiter ══ */
var _rl = {};
function rateLimit(key, ms) {
  var now = Date.now();
  if (_rl[key] && now-_rl[key] < ms) {
    toast('Please wait '+Math.ceil((ms-(now-_rl[key]))/1000)+'s before refreshing','warn',2500);
    return false;
  }
  _rl[key] = now;
  return true;
}

/* ════════════════════════════════════════════════════════
   PLANNING MODULE — FSN REPLENISHMENT  v1.0
   Reads: inventory[], skuMaster[], binMaster[] (already in memory — zero extra API calls)
   ════════════════════════════════════════════════════════ */
