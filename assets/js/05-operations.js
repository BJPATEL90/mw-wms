let adjLog = ls('adjLog') || [];
let _lastBinScanResult = null; // v5: stores last shelf scan response from server
let _serverGpBlockByBin = {};  // v6: server-provided shelf-level GP block qtys

function saveAdjLog(){ lsSet('adjLog', adjLog); }

function initCycleCount(){
  const bins=[...new Set(inventory.map(r=>r.bin).filter(Boolean))].sort();
  const skus=[...new Set(inventory.map(r=>r.skuCode).filter(Boolean))].sort();
  document.getElementById('cc_bin_list').innerHTML  = bins.map(b=>`<option value="${b}">`).join('');
  document.getElementById('cc_sku_list').innerHTML  = skus.map(s=>`<option value="${s}">`).join('');
  const msg = document.getElementById('ccSearchMsg');
  msg.innerText = inventory.length ? '' : '⚠ Upload inventory dump first to enable cycle counting.';
  renderCCLog(); updateCCStats();
}

/* ── Search by Bin ── */
function ccSearchByBin(){
  const bin = document.getElementById('cc_bin').value.trim();
  if(!bin) return;

  // v5: Try server first to get live MFG/EXP + cross-shelf data
  const loadEl = document.getElementById('ccLoadingMsg');
  if(loadEl) loadEl.style.display = 'block';

  callSheet('getInventoryByBin', { bin },
    function(res){
      if(loadEl) loadEl.style.display = 'none';
      if(res && res.success){
        _lastBinScanResult = res;
        if(res.rows && res.rows.length){
          const skus=[...new Set(res.rows.map(r=>r.skuCode))];
          if(skus.length===1) document.getElementById('cc_sku').value = skus[0];
          document.getElementById('ccSearchMsg').innerText = '';
          document.getElementById('ccAddStockCard').style.display = 'none';
          // v6: use gpBlockByBin from server for accurate shelf-level GP
          if(res.gpBlockByBin !== undefined){
            // Override local GP calc for this shelf with server's shelf-level value
            _serverGpBlockByBin = { [bin]: res.gpBlockByBin };
          }
          loadCCLines(res.rows, bin, res.sameBatchElsewhere || []);
        } else {
          document.getElementById('ccSearchMsg').innerText = '📭 Shelf is empty — fill in details to add stock.';
          hideCCLines();
          showAddStockForm(bin);
          renderCCOtherShelves([], bin);
        }
      } else {
        _ccSearchByBinLocal(bin);
      }
    },
    function(){
      if(loadEl) loadEl.style.display = 'none';
      _ccSearchByBinLocal(bin);
    }
  );
}

function _ccSearchByBinLocal(bin){
  const rows = inventory.filter(r => r.bin === bin);
  if(!rows.length){
    document.getElementById('ccSearchMsg').innerText = '📭 Shelf is empty — fill in details to add stock.';
    hideCCLines();
    showAddStockForm(bin);
    renderCCOtherShelves([], bin);
    return;
  }
  const skus=[...new Set(rows.map(r=>r.skuCode))];
  if(skus.length===1) document.getElementById('cc_sku').value = skus[0];
  document.getElementById('ccSearchMsg').innerText = '';
  document.getElementById('ccAddStockCard').style.display = 'none';
  // Build cross-shelf from local inventory
  const sbe=[];
  rows.forEach(r=>{
    if(!r.batch)return;
    inventory.filter(x=>x.bin!==bin&&x.skuCode===r.skuCode&&x.batch===r.batch).forEach(o=>{
      if(!sbe.find(s=>s.fromBin===o.bin&&s.skuCode===o.skuCode&&s.batch===o.batch))
        sbe.push({fromBin:o.bin,skuCode:o.skuCode,skuName:o.skuName,batch:o.batch,
                  mfgDate:o.mfgDate||'',expDate:o.expDate||'',stockType:o.stockType,qty:o.qty});
    });
  });
  loadCCLines(rows, bin, sbe);
}

/* ── Search by SKU ── */
function ccSearchBySKU(){
  const skuInput = document.getElementById('cc_sku').value.trim();
  if(!skuInput) return;
  const rows = inventory.filter(r =>
    r.skuCode === skuInput ||
    r.skuCode.toLowerCase().includes(skuInput.toLowerCase()) ||
    (r.skuName||'').toLowerCase().includes(skuInput.toLowerCase())
  );
  if(!rows.length){
    document.getElementById('ccSearchMsg').innerText = 'No inventory found for: ' + skuInput;
    hideCCLines(); return;
  }
  const bins=[...new Set(rows.map(r=>r.bin))];
  if(bins.length===1) document.getElementById('cc_bin').value = bins[0];
  document.getElementById('ccSearchMsg').innerText = '';
  document.getElementById('ccAddStockCard').style.display = 'none';
  loadCCLines(rows, document.getElementById('cc_bin').value.trim()||null);
}

/* ── Load existing inventory into count table ── */
function loadCCLines(rows, binFilter, sameBatchElsewhere){
  const filtered = binFilter ? rows.filter(r=>r.bin===binFilter) : rows;
  ccLines = filtered.map(r=>{
    // Look up MFG/EXP from inventory[] if not on this row
    let mfg = r.mfgDate || '';
    let exp = r.expDate  || '';
    if((!mfg || !exp) && r.batch){
      const inv = inventory.find(i =>
        i.skuCode === r.skuCode && i.batch === r.batch && i.bin === r.bin
      ) || inventory.find(i =>
        i.skuCode === r.skuCode && i.batch === r.batch
      );
      if(inv){ mfg = mfg || normDate(inv.mfgDate||''); exp = exp || normDate(inv.expDate||''); }
    }
    return {
      bin:        r.bin,
      skuCode:    r.skuCode,
      skuName:    r.skuName    || '',
      ean:        r.ean        || '',
      batch:      r.batch      || '',
      mfgDate:    normDate(mfg),
      expDate:    normDate(exp),
      stockType:  r.stockType  || 'GOOD',
      systemQty:  r.qty,
      gpBlock:    getGpBlock(r.skuCode, r.batch, binFilter||r.bin),
      countedQty: r.qty,
      remarks:    ''
    };
  });
  document.getElementById('ccLinesCard').style.display = 'block';
  document.getElementById('ccBinLabel').innerText = binFilter || (ccLines.length + ' rows');

  // Show "Add more stock" card below existing lines
  if(binFilter){
    showAddStockForm(binFilter, true);
  }

  renderCCLines();
  // v5: pass server-authoritative cross-shelf data
  renderCCOtherShelves(sameBatchElsewhere || [], binFilter || '');
}

function hideCCLines(){
  document.getElementById('ccLinesCard').style.display       = 'none';
  document.getElementById('ccOtherShelvesCard').style.display = 'none';
  ccLines = [];
}

/* ── Show Add-Stock form ── */
function showAddStockForm(bin, isPartial){
  const card  = document.getElementById('ccAddStockCard');
  const label = document.getElementById('ccAddBinLabel');
  card.style.display = 'block';
  label.innerText    = bin;
  if(isPartial){
    card.querySelector('.notice').className = 'notice info';
    card.querySelector('.notice').innerHTML = `<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>Add <b>additional</b> stock to this shelf (new batch or product).`;
  } else {
    card.querySelector('.notice').className = 'notice success';
    card.querySelector('.notice').innerHTML = `<svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>No inventory found on this shelf — enter details to add new stock.`;
  }
  // Clear fields
  ['ns_ean','ns_skuCode','ns_skuName','ns_batch','ns_qty'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });
  document.getElementById('ns_mfg').value = '';
  document.getElementById('ns_exp').value = '';
  document.getElementById('ns_stockType').value = 'GOOD';
  document.getElementById('ns_nameHint').innerText = '';
  document.getElementById('ns_mfgHint').innerText  = '';
  document.getElementById('ns_expHint').innerText  = '';
  document.getElementById('nsAddMsg').innerText     = '';
  // v5: clear batch hint
  const bh = document.getElementById('ns_batchHint');
  if(bh) bh.innerText = '';
}

/* ── EAN lookup ── */
function ccLookupByEAN(){
  const ean = document.getElementById('ns_ean').value.trim();
  if(!ean) return;
  // Search inventory for this EAN
  const match = inventory.find(r => r.ean && r.ean === ean);
  if(match){
    document.getElementById('ns_skuCode').value  = match.skuCode;
    document.getElementById('ns_skuName').value  = match.skuName || '';
    document.getElementById('ns_nameHint').innerText = '(from inventory)';
    if(match.mfgDate){ document.getElementById('ns_mfg').value = match.mfgDate; document.getElementById('ns_mfgHint').innerText='(from inventory)'; }
    if(match.expDate){ document.getElementById('ns_exp').value = match.expDate; document.getElementById('ns_expHint').innerText='(from inventory)'; }
    return;
  }
  // Fallback: search SKU master by EAN
  const skuM = skuMaster.find(s => s.ean === ean);
  if(skuM){
    document.getElementById('ns_skuCode').value  = skuM.code;
    document.getElementById('ns_skuName').value  = skuM.name;
    document.getElementById('ns_nameHint').innerText = '(from SKU master)';
  }
}

/* ── SKU Code lookup ── */
function ccLookupBySKU(){
  const code = document.getElementById('ns_skuCode').value.trim();
  if(!code) return;
  // Check existing inventory for MFG/EXP auto-fill
  const invMatch = inventory.find(r => r.skuCode === code && (r.mfgDate || r.expDate));
  if(invMatch){
    if(invMatch.skuName && !document.getElementById('ns_skuName').value)
      document.getElementById('ns_skuName').value = invMatch.skuName;
    if(invMatch.mfgDate){ document.getElementById('ns_mfg').value = dateToHTMLFormat(invMatch.mfgDate); document.getElementById('ns_mfgHint').innerText='(auto from inventory)'; }
    if(invMatch.expDate){ document.getElementById('ns_exp').value = dateToHTMLFormat(invMatch.expDate); document.getElementById('ns_expHint').innerText='(auto from inventory)'; }
    document.getElementById('ns_nameHint').innerText='(from inventory)';
    return;
  }
  const skuM = skuMaster.find(s => s.code === code);
  if(skuM){
    document.getElementById('ns_skuName').value = skuM.name || '';
    document.getElementById('ns_nameHint').innerText = '(from SKU master)';
  }
}

/* ── v5: Batch date auto-fill — calls getBatchDates on GAS ── */
function ccLookupBatchDates(){
  const skuCode = document.getElementById('ns_skuCode').value.trim();
  const batch   = document.getElementById('ns_batch').value.trim();
  const hint    = document.getElementById('ns_batchHint');
  if(hint) hint.innerText = '';
  if(!skuCode || !batch || batch.length < 3) return;

  // First try local inventory
  const local = inventory.find(r => r.skuCode===skuCode && r.batch===batch && (r.mfgDate||r.expDate));
  if(local){ _applyBatchDates(local.mfgDate, local.expDate, '(auto from inventory)'); return; }

  // Then try server
  callSheet('getBatchDates', { skuCode, batch },
    function(res){
      if(res && res.found){ _applyBatchDates(res.mfgDate, res.expDate, '(auto from Google Sheet)'); }
      else { if(hint) hint.innerText = 'New batch — enter MFG & EXP manually'; }
    },
    function(){ if(hint) hint.innerText = 'Enter MFG & EXP manually'; }
  );
}

function _applyBatchDates(mfg, exp, label){
  if(mfg){ document.getElementById('ns_mfg').value=dateToHTMLFormat(mfg); document.getElementById('ns_mfgHint').innerText=label; }
  if(exp){ document.getElementById('ns_exp').value=dateToHTMLFormat(exp); document.getElementById('ns_expHint').innerText=label; }
  const hint=document.getElementById('ns_batchHint');
  if(hint) hint.innerText = (mfg||exp) ? '✓ Dates auto-filled' : 'New batch — enter dates manually';
}

/* ── Save new stock on shelf (from Add Stock form) ── */
function saveNewStockOnShelf(){
  const bin      = document.getElementById('cc_bin').value.trim();
  const skuCode  = document.getElementById('ns_skuCode').value.trim();
  const skuName  = document.getElementById('ns_skuName').value.trim();
  const batch    = document.getElementById('ns_batch').value.trim();
  const mfgDate  = dateFromHTMLFormat(document.getElementById('ns_mfg').value.trim());
  const expDate  = dateFromHTMLFormat(document.getElementById('ns_exp').value.trim());
  const qty      = Number(document.getElementById('ns_qty').value)||0;
  const stockType= document.getElementById('ns_stockType').value;
  const ean      = document.getElementById('ns_ean').value.trim();
  const msg      = document.getElementById('nsAddMsg');

  if(!bin)     { msg.innerText='Search a shelf first'; return; }
  if(!skuCode) { msg.innerText='SKU Code required'; return; }
  if(!batch)   { msg.innerText='Batch No required'; return; }
  if(qty <= 0) { msg.innerText='Quantity must be > 0'; return; }
  msg.innerText = '';

  // Add to inventory (in-memory + localStorage)
  const existing = inventory.find(r =>
    r.bin===bin && r.skuCode===skuCode && r.batch===batch && r.stockType===stockType
  );
  if(existing){
    existing.qty += qty;
  } else {
    inventory.push({ bin, skuCode, skuName, batch, mfgDate, expDate, stockType, qty, ean });
  }
  dumpRows = inventory;
  updateBinsFromDump();
  saveAll();

  // v5: Record in Adjustment Log as ADD — send only this ONE new row to sheet
  const now = new Date().toISOString();
  const newAdjRow = {
    time: now, bin, skuCode, skuName, batch, mfgDate, expDate, stockType,
    qty: +qty, action: 'ADD', ref: 'New stock via Cycle Count', by: userName
  };
  adjLog.push(newAdjRow);
  saveAdjLog();
  renderAdjLog(); updateAdjStats();

  // Re-load the shelf view
  showResult('ccResult', `✅ Added ${qty} units of ${skuCode} (${batch}) to ${bin}`, 'success');
  document.getElementById('ns_qty').value = '';
  ccSearchByBin();

  // Sync to sheet — send only the 1 new ADD row (APPEND in Code.gs)
  callSheet('saveInventoryDump', {rows: inventory, mode:'replace'}, ()=>{}, ()=>{});
  callSheet('saveAdjustmentLog', [newAdjRow], ()=>{}, ()=>{});
  // v6: Sync current inventory view after adding stock
  callSheet('syncCurrentInventory', {}, r=>{ if(r&&r.success) renderCurrentInventory(); }, ()=>{});
}

/* ── Render existing count lines ── */
function renderCCLines(){
  const body = document.getElementById('ccLinesBody');
  if(!ccLines.length){ body.innerHTML='<tr class="empty-row"><td colspan="12">No inventory rows</td></tr>'; return; }
  body.innerHTML = ccLines.map((l,i)=>{
    const diff = l.countedQty - l.systemQty;
    const diffStyle = diff<0 ? 'color:var(--red);font-weight:700' : diff>0 ? 'color:var(--green);font-weight:700' : 'color:var(--slate4)';
    // v5: use CSS classes for date inputs; amber border = missing date
    const mfgDisplay = `<input type="date" class="date-input-sm ${l.mfgDate?'':'missing'}" value="${dateToHTMLFormat(l.mfgDate)||''}" onchange="ccLines[${i}].mfgDate=dateFromHTMLFormat(this.value);this.classList.remove('missing')">`;
    const expDisplay = `<input type="date" class="date-input-sm ${l.expDate?'':'missing'}" value="${dateToHTMLFormat(l.expDate)||''}" onchange="ccLines[${i}].expDate=dateFromHTMLFormat(this.value);this.classList.remove('missing')">`;
    return `<tr style="${diff!==0?'background:#fef9f9':''}">
      <td><span class="sku-badge" style="font-size:10px">${l.skuCode}</span></td>
      <td style="font-size:12px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${l.skuName}">${l.skuName||'—'}</td>
      <td style="font-size:11px;color:var(--slate4)">${l.ean||'—'}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px;font-weight:600">${l.batch||'—'}</td>
      <td>${mfgDisplay}</td>
      <td>${expDisplay}</td>
      <td><span class="status-badge ${l.stockType==='DAMAGE'||l.stockType==='DAMAGED'?'sb-occupied':l.stockType==='BLOCKED'?'sb-reserved':'sb-ok'}" style="font-size:9px">${l.stockType}</span></td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--indigo)">${l.systemQty.toLocaleString()}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--amber);font-weight:700" title="${l.gpBlock>0?'GP blocked qty for shelf '+l.bin:'No GP block for this shelf'}">${l.gpBlock>0?l.gpBlock.toLocaleString():'—'}</td>
      <td style="text-align:right"><input type="number" value="${l.countedQty}" min="0"
        onchange="ccLines[${i}].countedQty=Number(this.value);renderCCDiff(${i})"
        style="width:90px;padding:4px 7px;border:1.5px solid var(--teal);border-radius:5px;font-size:13px;text-align:right;font-family:'DM Mono',monospace;font-weight:700;background:#f0fdfa"></td>
      <td style="text-align:right;font-family:'DM Mono',monospace;min-width:50px" id="cc-diff-${i}">
        <span style="${diffStyle}">${diff>0?'+':''}${diff}</span>
      </td>
      <td><input class="tbl-input" placeholder="Remarks…" onchange="ccLines[${i}].remarks=this.value" style="width:120px;border:1px solid var(--border);border-radius:5px;font-size:11px;padding:3px 6px"></td>
    </tr>`;
  }).join('');
}

function renderCCDiff(i){
  const diff = ccLines[i].countedQty - ccLines[i].systemQty;
  const ds   = diff<0?'color:var(--red);font-weight:700':diff>0?'color:var(--green);font-weight:700':'color:var(--slate4)';
  const cell = document.getElementById('cc-diff-'+i);
  if(cell) cell.innerHTML = `<span style="${ds}">${diff>0?'+':''}${diff}</span>`;
}

/* ── Other shelves panel — v5: uses server sameBatchElsewhere data ── */
function renderCCOtherShelves(sameBatchElsewhere, currentBin){
  currentBin = currentBin || document.getElementById('cc_bin').value.trim();
  const card  = document.getElementById('ccOtherShelvesCard');
  const body  = document.getElementById('ccOtherShelvesBody');

  // If no server data passed, build from local inventory as fallback
  if(!sameBatchElsewhere || !sameBatchElsewhere.length){
    if(!ccLines.length && !currentBin){ card.style.display='none'; return; }
    const sbe=[];
    ccLines.forEach(l=>{
      if(!l.batch) return;
      inventory.filter(r=>r.skuCode===l.skuCode&&r.batch===l.batch&&r.bin!==currentBin).forEach(r=>{
        if(!sbe.find(s=>s.fromBin===r.bin&&s.skuCode===r.skuCode&&s.batch===r.batch))
          sbe.push({fromBin:r.bin,skuCode:r.skuCode,skuName:r.skuName,batch:r.batch,
                    mfgDate:r.mfgDate||'',expDate:r.expDate||'',stockType:r.stockType,qty:r.qty});
      });
    });
    if(!sbe.length){ card.style.display='none'; return; }
    sameBatchElsewhere=sbe;
  }

  // Group by SKU+Batch
  const groups={};
  sameBatchElsewhere.forEach(r=>{
    const k=r.skuCode+'|'+r.batch;
    if(!groups[k])groups[k]=[];
    groups[k].push(r);
  });

  let html='';
  Object.entries(groups).forEach(([key,others])=>{
    const [sku,bat]=key.split('|');
    html+=`<div style="padding:14px 18px;border-bottom:1px solid var(--border)">
      <div style="font-size:12px;font-weight:700;color:var(--slate3);margin-bottom:10px">
        <span class="sku-badge">${sku}</span>&nbsp; Batch: <b>${bat||'—'}</b>
        &nbsp;→&nbsp; Move to <span style="color:var(--teal);font-family:'DM Mono',monospace">${currentBin||'?'}</span>
        <span style="font-size:11px;font-weight:400;color:var(--amber);margin-left:10px">⚠ Same batch on other shelf(s)</span>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:var(--surface)">
          <th style="padding:6px 10px;text-align:left;border-bottom:1px solid var(--border)">From Shelf</th>
          <th style="padding:6px 10px;text-align:left;border-bottom:1px solid var(--border)">MFG Date</th>
          <th style="padding:6px 10px;text-align:left;border-bottom:1px solid var(--border)">EXP Date</th>
          <th style="padding:6px 10px;text-align:right;border-bottom:1px solid var(--border)">Available Qty</th>
          <th style="padding:6px 10px;text-align:left;border-bottom:1px solid var(--border)">Type</th>
          <th style="padding:6px 10px;border-bottom:1px solid var(--border)">Move Qty → ${currentBin||'current shelf'}</th>
        </tr></thead>
        <tbody>${others.map(r=>{
          const safeid=(r.fromBin+'_'+sku+'_'+bat).replace(/[^a-z0-9]/gi,'_');
          return `<tr>
            <td style="padding:6px 10px"><span class="bin-badge" style="font-size:10px">${r.fromBin}</span></td>
            <td style="padding:6px 10px;font-size:11px;color:var(--slate4)">${r.mfgDate||'—'}</td>
            <td style="padding:6px 10px;font-size:11px;color:var(--slate4)">${r.expDate||'—'}</td>
            <td style="padding:6px 10px;text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--indigo)">${r.qty.toLocaleString()}</td>
            <td style="padding:6px 10px"><span class="status-badge ${r.stockType==='DAMAGE'?'sb-occupied':'sb-ok'}" style="font-size:9px">${r.stockType}</span></td>
            <td style="padding:6px 10px">
              <div style="display:flex;align-items:center;gap:8px">
                <input type="number" id="tqty_${safeid}" value="${r.qty}" min="1" max="${r.qty}"
                  style="width:80px;padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:12px;font-family:'DM Mono',monospace">
                <button class="btn btn-primary btn-sm" onclick="doTransferToCurrent('${r.fromBin}','${currentBin}','${sku}','${bat}','${r.stockType}',${r.qty},'${safeid}','${r.mfgDate||''}','${r.expDate||''}')">
                  Move → Current Shelf
                </button>
              </div>
            </td>
          </tr>`;
        }).join('')}</tbody>
      </table>
    </div>`;
  });

  if(html){card.style.display='block';body.innerHTML=html;}
  else{card.style.display='none';}
}

/* ── Transfer stock to current shelf — records REMOVE + ADD in adj log ── */
// ⑧ v6: doTransferToCurrent with idempotency key to prevent duplicate writes
function doTransferToCurrent(fromBin, toBin, sku, batch, stockType, maxQty, safeid, mfgDate, expDate){
  if(!toBin){ alert('Search a destination shelf first (use the Shelf / Bin search above)'); return; }

  // Disable button immediately to prevent double-click
  const btn = document.querySelector(`button[onclick*="${safeid}"]`);
  if(btn){ if(btn.dataset.moving==='1'){ alert('Move already in progress…'); return; } btn.dataset.moving='1'; btn.disabled=true; btn.innerText='Moving…'; }

  const qtyEl   = document.getElementById('tqty_'+safeid);
  const moveQty = qtyEl ? Number(qtyEl.value)||0 : maxQty;
  if(moveQty <= 0)      { if(btn){btn.disabled=false;btn.dataset.moving='';btn.innerText='Move → Current Shelf';} alert('Enter qty to move'); return; }
  if(moveQty > maxQty)  { if(btn){btn.disabled=false;btn.dataset.moving='';btn.innerText='Move → Current Shelf';} alert('Cannot move more than available: '+maxQty); return; }
  if(fromBin === toBin) { if(btn){btn.disabled=false;btn.dataset.moving='';btn.innerText='Move → Current Shelf';} alert('Source and destination are the same shelf'); return; }

  if(!confirm('Move '+moveQty+' units of '+sku+' (Batch: '+batch+') from '+fromBin+' to '+toBin+'. Confirm?'))
    { if(btn){btn.disabled=false;btn.dataset.moving='';btn.innerText='Move → Current Shelf';} return; }

  const idemKey = fromBin+'_'+toBin+'_'+sku+'_'+batch+'_'+moveQty+'_'+Date.now();

  const now = new Date().toISOString();

  // ── Update inventory in memory ──
  let remaining = moveQty;
  inventory.forEach(r=>{
    if(r.bin===fromBin && r.skuCode===sku && r.batch===batch && r.stockType===stockType && remaining>0){
      const cut = Math.min(remaining, r.qty);
      r.qty -= cut; remaining -= cut;
    }
  });
  inventory = inventory.filter(r=>r.qty>0);

  const existing = inventory.find(r=>r.bin===toBin && r.skuCode===sku && r.batch===batch && r.stockType===stockType);
  if(existing){
    existing.qty += moveQty;
  } else {
    const src = inventory.find(r=>r.skuCode===sku&&r.batch===batch)||{};
    inventory.push({
      bin:toBin, skuCode:sku, skuName:src.skuName||sku, batch,
      stockType, qty:moveQty,
      mfgDate: mfgDate||src.mfgDate||'',
      expDate:  expDate||src.expDate||'',
      ean: src.ean||''
    });
  }

  dumpRows = inventory;
  updateBinsFromDump();
  saveAll();

  // ── v5: Adjustment log — named vars, send only 2 new rows (APPEND in Code.gs) ──
  const skuName = inventory.find(r=>r.skuCode===sku)?.skuName || sku;
  const removeRow = { time:now, bin:fromBin, skuCode:sku, skuName, batch, mfgDate, expDate, stockType, qty:-moveQty, action:'REMOVE', ref:`Moved to ${toBin}`,   by:userName };
  const addRow    = { time:now, bin:toBin,   skuCode:sku, skuName, batch, mfgDate, expDate, stockType, qty:+moveQty, action:'ADD',    ref:`Moved from ${fromBin}`, by:userName };
  adjLog.push(removeRow);
  adjLog.push(addRow);
  saveAdjLog();
  renderAdjLog(); updateAdjStats();

  // Sync to sheet — use idempotent call to prevent duplicate writes
  callSheet('saveInventoryDump',  {rows:inventory, mode:'replace'}, ()=>{}, ()=>{});
  callSheet('saveAdjLogIdempotent', { rows:[removeRow, addRow], key: idemKey },
    r=>{
      if(r && r.duplicate) console.log('Duplicate move prevented by server');
      // Re-enable button after server confirms
      const btn2 = document.querySelector(`button[onclick*="${safeid}"]`);
      if(btn2){ btn2.disabled=false; btn2.dataset.moving=''; btn2.innerText='Move → Current Shelf'; }
    },
    e=>{ console.warn('Adj log sync failed:', e);
      const btn2 = document.querySelector(`button[onclick*="${safeid}"]`);
      if(btn2){ btn2.disabled=false; btn2.dataset.moving=''; btn2.innerText='Move → Current Shelf'; }
    }
  );
  // Sync current inventory view
  callSheet('syncCurrentInventory', {}, r=>{ if(r&&r.success) renderCurrentInventory(); }, ()=>{});

  // ── Update ccLines for BOTH source and destination bins ──

  // Destination bin (current shelf being counted): increase both system + counted
  ccLines.forEach(l => {
    if(l.bin === toBin && l.skuCode === sku && l.batch === batch && l.stockType === stockType){
      l.systemQty  += moveQty;
      l.countedQty += moveQty;
    }
  });

  // Source bin: decrease systemQty to reflect the move — counted stays as physically seen
  // This is KEY: if user counted 1200 and we moved 96 from elsewhere,
  // the source bin's system qty dropped by 96 — so its diff recalculates correctly
  ccLines.forEach(l => {
    if(l.bin === fromBin && l.skuCode === sku && l.batch === batch && l.stockType === stockType){
      l.systemQty -= moveQty;   // system qty reduced (stock moved out)
      // countedQty stays — that's what was physically counted at that shelf
    }
  });

  renderCCLines();

  // ── FIX7: Remove source row from sameBatchElsewhere panel ──
  // If full qty moved → remove that row entirely
  // If partial → reduce qty in that row
  if(_lastBinScanResult && _lastBinScanResult.sameBatchElsewhere){
    const sbe = _lastBinScanResult.sameBatchElsewhere;
    const idx  = sbe.findIndex(r => r.fromBin===fromBin && r.skuCode===sku && r.batch===batch);
    if(idx >= 0){
      if(moveQty >= maxQty){
        sbe.splice(idx, 1);  // full move — remove source row
      } else {
        sbe[idx].qty -= moveQty;  // partial move — reduce qty
      }
      renderCCOtherShelves(sbe, toBin);
    }
  }

  showResult('ccResult', '✅ Moved '+moveQty+' units of '+sku+' from '+fromBin+' → '+toBin, 'success');

  // Re-enable button after move completes
  if(btn){ btn.disabled=false; btn.dataset.moving=''; btn.innerText='Move → Current Shelf'; }
}

/* ── Save Count Lines ── */
function saveCountLines(){
  if(!ccLines.length){ alert('No count lines to save'); return; }
  const now = new Date().toISOString();
  const newEntries = [];

  ccLines.forEach(l=>{
    // Recalculate systemQty from current inventory[] at save time
    // This accounts for any stock moves done during the count session
    const liveQty = inventory
      .filter(r => r.bin===l.bin && r.skuCode===l.skuCode &&
                   r.batch===l.batch && normalizeStockType_(r.stockType)===normalizeStockType_(l.stockType))
      .reduce((s,r) => s + (r.qty||0), 0);

    // Use live qty if available, otherwise use ccLines systemQty (offline fallback)
    const sysQty = liveQty > 0 ? liveQty : l.systemQty;

    const diff = l.countedQty - sysQty;
    const entry = {
      time:now, bin:l.bin, skuCode:l.skuCode, skuName:l.skuName||'',
      batch:l.batch, systemQty:sysQty, countedQty:l.countedQty, diff,
      mfgDate:l.mfgDate, expDate:l.expDate,
      stockType:l.stockType||'GOOD', gpBlock:l.gpBlock||0,
      remarks:l.remarks||'', by:userName
    };
    // Update ccLines systemQty to match live (so display is accurate too)
    l.systemQty = sysQty;
    ccLog.push(entry);
    newEntries.push(entry);
  });

  updateCCStats(); renderCCLog();
  const diffs = ccLines.filter(l=>l.countedQty!==l.systemQty).length;
  showResult('ccResult', 'Count saved. '+diffs+' difference(s) recorded.', diffs?'warn':'success');

  callSheet('saveCycleCountLog', newEntries,
    r=>{
      showResult('ccResult','Count saved to sheet. '+diffs+' difference(s).', diffs?'warn':'success');
      // ⑮ v7: Sync Current Inventory after every cycle count save
      callSheet('syncCurrentInventory', {},
        r2=>{ if(r2&&r2.success){ renderCurrentInventory(); toast('Current Inventory updated','success'); } },
        ()=>{}
      );
    },
    e=>showResult('ccResult','Count saved locally. ⚠ Sheet sync failed: '+e,'warn')
  );
}

function updateCCStats(){
  document.getElementById('cc-counted').innerText = ccLog.length;
  document.getElementById('cc-match').innerText   = ccLog.filter(l=>l.diff===0).length;
  document.getElementById('cc-diff').innerText    = ccLog.filter(l=>l.diff!==0).length;
}

function renderCCLog(){
  const body = document.getElementById('ccLogBody');
  if(!ccLog.length){ body.innerHTML='<tr class="empty-row"><td colspan="15">No count entries yet</td></tr>'; return; }
  body.innerHTML = [...ccLog].reverse().map(l=>{
    const dt  = new Date(l.time);
    const diff= Number(l.diff||0);
    const ds  = diff<0?'color:var(--red);font-weight:700':diff>0?'color:var(--green);font-weight:700':'';
    const st  = diff===0?'MATCH':diff>0?'EXCESS':'SHORT';
    const stClass = diff===0?'sb-ok':diff>0?'sb-reserved':'sb-diff';
    return `<tr>
      <td style="font-size:11px;white-space:nowrap">${dt.toLocaleDateString('en-IN')}</td>
      <td style="font-size:11px;white-space:nowrap">${dt.toLocaleTimeString('en-IN')}</td>
      <td><span class="bin-badge" style="font-size:10px">${l.bin}</span></td>
      <td><span class="sku-badge" style="font-size:10px">${l.skuCode}</span></td>
      <td style="font-size:11px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${l.skuName||'—'}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${l.batch||'—'}</td>
      <td style="font-size:11px">${l.mfgDate||'—'}</td>
      <td style="font-size:11px">${l.expDate||'—'}</td>
      <td><span class="status-badge sb-ok" style="font-size:9px">${l.stockType||'GOOD'}</span></td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${l.systemQty}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700">${l.countedQty}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace"><span style="${ds}">${diff>0?'+':''}${diff}</span></td>
      <td><span class="status-badge ${stClass}" style="font-size:9px">${st}</span></td>
      <td style="font-size:11px;color:var(--slate4)">${l.remarks||'—'}</td>
      <td style="font-size:11px">${l.by||'—'}</td>
    </tr>`;
  }).join('');
}

function exportCCLogXLSX(){
  if(!ccLog.length){ alert('No count log to export'); return; }
  const ws = XLSX.utils.json_to_sheet(ccLog.map(l=>({
    'Date':new Date(l.time).toLocaleDateString('en-IN'),
    'Time':new Date(l.time).toLocaleTimeString('en-IN'),
    'Bin / Shelf':l.bin, 'SKU Code':l.skuCode, 'SKU Name':l.skuName||'',
    'Batch No':l.batch, 'MFG Date':l.mfgDate||'', 'EXP Date':l.expDate||'',
    'Stock Type':l.stockType||'GOOD', 'System Qty':l.systemQty,
    'GP Block Qty':l.gpBlock||0, 'Counted Qty':l.countedQty,
    'Difference':l.diff,
    'Status':l.diff===0?'MATCH':l.diff>0?'EXCESS':'SHORT',
    'Remarks':l.remarks||'', 'Counted By':l.by
  })));
  ws['!cols']=[{wch:12},{wch:10},{wch:14},{wch:28},{wch:34},{wch:14},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:10},{wch:20},{wch:16}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Cycle Count Log');
  XLSX.writeFile(wb,'Cycle_Count_Log.xlsx');
}

function downloadCCTemplate(){
  const ws=XLSX.utils.aoa_to_sheet([
    ['Date','Bin / Shelf','SKU Code','SKU Name','Batch No','MFG Date','EXP Date','MRP','Stock Type','SKU Status','System Qty','GP Block Qty','Counted Qty','Remarks'],
    ['','R1-C1-001','MWLJNTP.0003.B0_N','LJ NutriMix 2+ 350gm Chocolate Jar','BATCH001','01-Jan-2024','31-Dec-2025','399','GOOD','ACTIVE',100,0,'',''],
    ['','R1-C2-001','MWLJNTP.00059.B0_N','LJ NutriMix 7+ 350gm Chocolate Jar','BATCH002','15-Feb-2024','28-Feb-2026','449','GOOD','ACTIVE',200,0,'',''],
  ]);
  ws['!cols']=[{wch:12},{wch:14},{wch:28},{wch:34},{wch:14},{wch:12},{wch:12},{wch:8},{wch:10},{wch:10},{wch:12},{wch:12},{wch:12},{wch:20}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Cycle Count Template');
  XLSX.writeFile(wb,'Cycle_Count_Template.xlsx');
}

/* ════════════════════ ADJUSTMENT LOG ════════════════════ */
function initAdjustment(){
  renderAdjLog(); updateAdjStats();
}

function renderAdjLog(){
  const af  = document.getElementById('adjActionFilter')?.value || '';
  const q   = (document.getElementById('adjSearch')?.value||'').toLowerCase();
  let rows  = adjLog;
  if(af) rows = rows.filter(l=>l.action===af);
  if(q)  rows = rows.filter(l=>`${l.skuCode} ${l.bin} ${l.batch} ${l.skuName}`.toLowerCase().includes(q));

  const body = document.getElementById('adjLogBody');
  if(!rows.length){ body.innerHTML='<tr class="empty-row"><td colspan="13">No adjustments yet</td></tr>'; return; }
  body.innerHTML = [...rows].reverse().map(l=>{
    const dt = new Date(l.time);
    const isRemove = l.action==='REMOVE';
    // v5: qty stored negative for REMOVE, positive for ADD — display with correct sign
    const absQty = Math.abs(Number(l.qty||0));
    const qtyDisplay = isRemove
      ? `<span style="color:var(--red);font-weight:700">−${absQty.toLocaleString()}</span>`
      : `<span style="color:var(--green);font-weight:700">+${absQty.toLocaleString()}</span>`;
    return `<tr style="${isRemove?'background:#fff8f8':'background:#f0fdf4'}">
      <td style="font-size:11px;white-space:nowrap">${dt.toLocaleDateString('en-IN')}</td>
      <td style="font-size:11px;white-space:nowrap">${dt.toLocaleTimeString('en-IN')}</td>
      <td><span class="bin-badge" style="font-size:10px">${l.bin}</span></td>
      <td><span class="sku-badge" style="font-size:10px">${l.skuCode}</span></td>
      <td style="font-size:11px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${l.skuName||'—'}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${l.batch||'—'}</td>
      <td style="font-size:11px">${l.mfgDate||'—'}</td>
      <td style="font-size:11px">${l.expDate||'—'}</td>
      <td><span class="status-badge ${l.stockType==='DAMAGE'?'sb-occupied':'sb-ok'}" style="font-size:9px">${l.stockType}</span></td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700">${qtyDisplay}</td>
      <td><span class="status-badge ${isRemove?'sb-remove':'sb-add'}" style="font-size:10px">${l.action}</span></td>
      <td style="font-size:11px;color:var(--slate4)">${l.ref||'—'}</td>
      <td style="font-size:11px">${l.by||'—'}</td>
    </tr>`;
  }).join('');
}

function updateAdjStats(){
  const total   = adjLog.length;
  const removes = adjLog.filter(l=>l.action==='REMOVE').length;
  const adds    = adjLog.filter(l=>l.action==='ADD').length;
  const skus    = new Set(adjLog.map(l=>l.skuCode)).size;
  document.getElementById('adj-total').innerText  = total;
  document.getElementById('adj-remove').innerText = removes;
  document.getElementById('adj-add').innerText    = adds;
  document.getElementById('adj-skus').innerText   = skus;
}

function exportAdjLogXLSX(){
  if(!adjLog.length){ alert('No adjustment log to export'); return; }
  const ws = XLSX.utils.json_to_sheet(adjLog.map(l=>({
    'Date':new Date(l.time).toLocaleDateString('en-IN'),
    'Time':new Date(l.time).toLocaleTimeString('en-IN'),
    'Bin / Shelf':l.bin, 'SKU Code':l.skuCode, 'SKU Name':l.skuName||'',
    'Batch No':l.batch, 'MFG Date':l.mfgDate||'', 'EXP Date':l.expDate||'',
    'Stock Type':l.stockType, 'Qty Moved':l.qty,
    'Action':l.action, 'Reference':l.ref||'', 'Counted By':l.by
  })));
  ws['!cols']=[{wch:12},{wch:10},{wch:14},{wch:28},{wch:34},{wch:14},{wch:12},{wch:12},{wch:12},{wch:10},{wch:8},{wch:22},{wch:16}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Adjustment Log');
  XLSX.writeFile(wb,'Adjustment_Log.xlsx');
}

function clearAdjLog(){
  // Local-only clear (non-admin view)
  adjLog=[]; saveAdjLog(); renderAdjLog(); updateAdjStats();
}

/** Admin: clears local + deletes all rows from ADJUSTMENT_LOG sheet */
function clearAdjLogFull(){
  if(userRole !== 'ADMIN'){ toast('Admin only','error'); return; }
  const typed = prompt('Type CLEAR to delete all Adjustment Log entries from Google Sheet:');
  if(!typed || typed.trim().toUpperCase() !== 'CLEAR'){ toast('Cancelled','warn'); return; }
  adjLog=[]; saveAdjLog(); renderAdjLog(); updateAdjStats();
  callSheet('clearSheetData', { sheetName: 'ADJUSTMENT_LOG' },
    function(r){ toast(r&&r.success ? '✅ Adjustment Log cleared from sheet ('+r.count+' rows)' : '❌ '+(r&&r.message||'Error'), r&&r.success?'success':'error', 4000); },
    function(e){ toast('❌ Sheet error: '+e, 'error'); }
  );
}

/** Admin: clears local CC log + deletes from CYCLE_COUNT_LOG sheet */
function clearCCSessionFull(){
  if(userRole !== 'ADMIN'){ toast('Admin only','error'); return; }
  const typed = prompt('Type CLEAR to delete all Cycle Count Log entries from Google Sheet:');
  if(!typed || typed.trim().toUpperCase() !== 'CLEAR'){ toast('Cancelled','warn'); return; }
  ccLog=[]; lsSet('ccLog',[]); updateCCStats(); renderCCLog();
  callSheet('clearSheetData', { sheetName: 'CYCLE_COUNT_LOG' },
    function(r){ toast(r&&r.success ? '✅ Cycle Count Log cleared from sheet ('+r.count+' rows)' : '❌ '+(r&&r.message||'Error'), r&&r.success?'success':'error', 4000); },
    function(e){ toast('❌ Sheet error: '+e, 'error'); }
  );
}

// v5: Manual "Sync to Sheet" — pushes full in-memory log (useful after reconnecting GAS)
function syncAdjLogToSheet(){
  if(!adjLog.length){alert('No entries to sync');return;}
  callSheet('saveAdjustmentLog', adjLog,
    r=>showResult('adjResult','☁ Full adjustment log synced to Google Sheet ('+adjLog.length+' rows)','success'),
    e=>showResult('adjResult','⚠ Sync failed: '+e,'warn')
  );
}

/* ════════════════════ CURRENT INVENTORY ════════════════════ */

// In-memory current inventory (aggregated view, rebuilt from inventory[])
let currentInventory = [];
let _ciFlushMode    = false; // v9: when true, shows empty state (post-flush)

function initCurrentInventory(){
  buildCurrentInventoryLocal();
  renderCurrentInventory();
}

/**
 * Build aggregated current inventory from local inventory[] array.
 * Key: bin|skuCode|batch|stockType
 * GP block is shelf-level (sum of all GP for that bin).
 */
function buildCurrentInventoryLocal(){
  const map = {};
  inventory.forEach(r=>{
    if(!r.bin || !r.skuCode || r.qty <= 0) return;
    const key = r.bin+'|'+r.skuCode+'|'+(r.batch||'')+'|'+(r.stockType||'GOOD');
    if(!map[key]) map[key] = {
      bin: r.bin, skuCode: r.skuCode, skuName: r.skuName||'',
      batch: r.batch||'', mfgDate: r.mfgDate||'', expDate: r.expDate||'',
      mrp: r.mrp||'', stockType: normalizeStockType_(r.stockType), // v9: BAD_INVENTORY→DAMAGE etc
      skuStatus: r.skuStatus||'ACTIVE', qty: 0
    };
    map[key].qty += r.qty;
  });

  // Build GP map per bin
  const gpMap = {};
  gatepasses.forEach(gp=>{ if(gp.bin) gpMap[gp.bin] = (gpMap[gp.bin]||0) + (gp.qty||0); });

  currentInventory = Object.values(map).map(r=>({
    ...r,
    gpBlock: gpMap[r.bin] || 0
  })).sort((a,b)=> a.bin.localeCompare(b.bin, undefined, {numeric:true}) || a.skuCode.localeCompare(b.skuCode));

  // Rebuild bin filter dropdown
  const bf = document.getElementById('ciBinFilter');
  if(bf){
    const bins = [...new Set(currentInventory.map(r=>r.bin))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    bf.innerHTML = '<option value="">All Bins</option>' + bins.map(b=>`<option value="${b}">${b}</option>`).join('');
  }
}

function renderCurrentInventory(){
  // v9: if user explicitly flushed current inventory, show empty until
  // they click Sync or upload new data — do NOT rebuild from inventory[]
  if(!_ciFlushMode){
    buildCurrentInventoryLocal();
  }
  const q   = (document.getElementById('ciSearch')?.value||'').toLowerCase();
  const bf  = document.getElementById('ciBinFilter')?.value||'';
  const tf  = document.getElementById('ciTypeFilter')?.value||'';

  const rows = currentInventory.filter(r=>{
    if(bf && r.bin !== bf) return false;
    if(tf && r.stockType !== tf) return false;
    if(q && !`${r.bin} ${r.skuCode} ${r.skuName} ${r.batch}`.toLowerCase().includes(q)) return false;
    return true;
  });

  // Summary chips
  const totalQty = rows.reduce((s,r)=>s+r.qty,0);
  const totalGP  = rows.reduce((s,r)=>s+r.gpBlock,0);
  const skus     = new Set(rows.map(r=>r.skuCode)).size;
  const bins     = new Set(rows.map(r=>r.bin)).size;
  document.getElementById('ciSummary').innerHTML =
    `<span class="chip chip-total">${rows.length} Lines</span>` +
    `<span class="chip chip-ok">${skus} SKUs</span>` +
    `<span class="chip chip-ok">${bins} Bins</span>` +
    `<span class="chip chip-total">Qty: ${totalQty.toLocaleString()}</span>` +
    (totalGP > 0 ? `<span class="chip chip-warn">GP Block: ${totalGP.toLocaleString()}</span>` : '');

  const body = document.getElementById('ciBody');
  if(_ciFlushMode){
    body.innerHTML = '<tr class="empty-row"><td colspan="11">🗑 Current Inventory flushed. Click <b>☁ Sync Inventory</b> to reload from sheet, or upload new Inventory Dump.</td></tr>';
    document.getElementById('ciSummary').innerHTML = '<span class="chip chip-total">0 Lines (flushed)</span>';
    return;
  }
  if(!rows.length){
    body.innerHTML = '<tr class="empty-row"><td colspan="11">No inventory found — upload a dump and click Sync Inventory</td></tr>';
    return;
  }

  const visible=rows.slice(0,500);
  const more=rows.length-visible.length;
  body.innerHTML = (more?`<tr><td colspan="11" style="font-size:11px;color:var(--slate4);background:#f8fafc">Showing first 500 of ${rows.length.toLocaleString()} inventory lines. Use search/filter to narrow the list.</td></tr>`:'')+
  visible.map(r=>{
    const net = Math.max(0, r.qty - r.gpBlock);
    // v9: After normalization, only GOOD/DAMAGE/BLOCKED possible
    const typeClass = r.stockType==='GOOD' ? 'sb-ok' : r.stockType==='DAMAGE' ? 'sb-occupied' : 'sb-reserved';
    return `<tr>
      <td><span class="bin-badge">${r.bin}</span></td>
      <td><span class="sku-badge">${r.skuCode}</span></td>
      <td style="font-size:12px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.skuName||'—'}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${r.batch||'—'}</td>
      <td style="font-size:11px">${r.mfgDate||'—'}</td>
      <td style="font-size:11px">${r.expDate||'—'}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${r.mrp||'—'}</td>
      <td><span class="status-badge ${typeClass}" style="font-size:9px">${r.stockType}</span></td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--indigo)">${r.qty.toLocaleString()}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--amber)">${r.gpBlock > 0 ? r.gpBlock.toLocaleString() : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:${net<r.qty?'var(--red)':'var(--green)'}">${net.toLocaleString()}</td>
    </tr>`;
  }).join('');
}

function doSyncCurrentInventory(){
  const btn = document.querySelector('button[onclick="doSyncCurrentInventory()"]');
  if(btn){ btn.disabled=true; btn.innerText='Syncing…'; }
  buildCurrentInventoryLocal();
  callSheet('syncCurrentInventory', {},
    r=>{
      if(btn){ btn.disabled=false; btn.innerText='☁ Sync Inventory'; }
      if(r && r.success){
        _ciFlushMode = false; // v9: sync clears flush mode — data is fresh from sheet
        const ts = new Date().toLocaleTimeString('en-IN');
        const el = document.getElementById('ciLastSync');
        if(el) el.innerText = 'Last synced: ' + ts + ' (' + (r.count||0) + ' rows)';
        showResult('ciNotice','✅ Current Inventory synced to Google Sheet — '+r.count+' rows written.','success');
        renderCurrentInventory();
      } else {
        showResult('ciNotice','⚠ Sync failed: '+(r&&r.message||'unknown'),'warn');
      }
    },
    e=>{
      if(btn){ btn.disabled=false; btn.innerText='☁ Sync Inventory'; }
      showResult('ciNotice','⚠ Sync error: '+e,'warn');
    }
  );
}

function exportCurrentInventoryXLSX(){
  buildCurrentInventoryLocal();
  if(!currentInventory.length){ alert('No inventory data to export'); return; }
  const ws = XLSX.utils.json_to_sheet(currentInventory.map(r=>({
    'Bin / Shelf': r.bin, 'SKU Code': r.skuCode, 'SKU Name': r.skuName||'',
    'Batch No': r.batch||'', 'MFG Date': r.mfgDate||'', 'EXP Date': r.expDate||'',
    'MRP': r.mrp||'', 'Stock Type': r.stockType, 'SKU Status': r.skuStatus||'ACTIVE',
    'System Qty': r.qty, 'GP Block Qty': r.gpBlock, 'Net Available': Math.max(0, r.qty-r.gpBlock)
  })));
  ws['!cols']=[{wch:14},{wch:28},{wch:34},{wch:14},{wch:12},{wch:12},{wch:8},{wch:10},{wch:10},{wch:12},{wch:12},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Current Inventory');
  XLSX.writeFile(wb, 'Current_Inventory_' + new Date().toISOString().slice(0,10) + '.xlsx');
}

/* ════════════════════ v7: TOAST NOTIFICATIONS ════════════════════ */
/**
 * Show a non-blocking toast notification.
 * type: 'success' | 'warn' | 'error' | 'info'
 * duration: ms (default 3500)
 */
