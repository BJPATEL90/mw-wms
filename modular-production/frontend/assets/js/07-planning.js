let fsnPlanRows = [];

function inferBrandFromPrefix_(skuCode) {
  const p = String(skuCode || '').slice(0, 4).toUpperCase();
  if (p === 'MWBW') return 'Bodywise';
  if (p === 'MWLJ') return 'LittleJoys';
  if (p === 'MWMM') return 'ManMatters';
  return 'Unknown';
}

function initFSNReplen() {
  if (!inventory || !inventory.length) {
    showResult('fsn-notice', '⚠ No inventory loaded — upload a dump first.', 'warn');
    return;
  }
  if (!binMaster || !binMaster.length) {
    showResult('fsn-notice', '⚠ Bin Master is empty — load Bin Master data first (Bin Status tab → Bulk Upload).', 'warn');
    return;
  }

  /* Build SKU lookup: code → sku object */
  const skuMap = {};
  (skuMaster || []).forEach(s => { skuMap[String(s.code || '').trim()] = s; });

  /* Which bins currently have stock */
  const occupiedBins = new Set();
  (inventory || []).forEach(r => { if (r.bin && r.qty > 0) occupiedBins.add(r.bin); });

  /* Build empty-bin pools: "BrandZone|FSNClass" → [binId, ...] sorted by level asc */
  const emptyPool = {};
  (binMaster || [])
    .filter(b => b.id && !occupiedBins.has(b.id) && String(b.status || '').toLowerCase() !== 'allocated')
    .sort((a, b) => (a.level || 0) - (b.level || 0) || (a.row || 0) - (b.row || 0) || (a.col || 0) - (b.col || 0))
    .forEach(b => {
      const key = String(b.brand || 'Unknown') + '|' + String(b.fsn || 'B').toUpperCase();
      if (!emptyPool[key]) emptyPool[key] = [];
      emptyPool[key].push(b.id);
    });

  /* Pointer per pool key so we never suggest the same empty bin twice */
  const poolPtr = {};
  function nextEmpty(brand, targetFSN) {
    const zones = [brand, 'Overflow', 'Unknown'];
    for (const zone of zones) {
      const key = zone + '|' + targetFSN;
      if (!poolPtr[key]) poolPtr[key] = 0;
      const arr = emptyPool[key] || [];
      if (poolPtr[key] < arr.length) return arr[poolPtr[key]++];
    }
    return null;
  }

  /* Aggregate GOOD inventory: bin+sku+batch (batch-level rows) */
  function parseDate_(s) {
    if (!s) return null;
    const str = String(s).trim();
    const m = str.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) return new Date(+m[3], +m[2]-1, +m[1]);
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  }
  const _today = new Date(); _today.setHours(0,0,0,0);
  const binSkuBatchAgg = {};
  (inventory || []).forEach(r => {
    if (!r.bin || !r.skuCode || r.qty <= 0) return;
    const st = String(r.stockType || '').toUpperCase();
    if (st === 'DAMAGE' || st === 'BAD' || st === 'BAD_INVENTORY' || st === 'BLOCKED') return;
    const key = r.bin + '|||' + r.skuCode + '|||' + (r.batch || '');
    if (!binSkuBatchAgg[key]) binSkuBatchAgg[key] = { bin:r.bin, skuCode:r.skuCode, skuName:r.skuName||'', batch:r.batch||'', expDate:r.expDate||'', mfgDate:r.mfgDate||'', qty:0 };
    binSkuBatchAgg[key].qty += (r.qty || 0);
  });

  const binById = {};
  (binMaster || []).forEach(b => { if (b.id) binById[b.id] = b; });

  const rows = [];
  let eligibleLines = 0;
  let compliantLines = 0;

  Object.values(binSkuBatchAgg).forEach(entry => {
    const binInfo = binById[entry.bin];
    if (!binInfo) return;
    const sku = skuMap[entry.skuCode];
    const skuCls = String((sku && sku.cls) || 'B').toUpperCase();
    const binCls = String(binInfo.fsn || 'B').toUpperCase();
    const brand  = inferBrandFromPrefix_(entry.skuCode);
    const skuName = (sku && sku.name) || entry.skuName || '—';
    if (!['A','B','C'].includes(skuCls) || !['A','B','C'].includes(binCls)) return;
    eligibleLines++;

    let sl = null;
    const exp = parseDate_(entry.expDate), mfg = parseDate_(entry.mfgDate);
    if (exp) {
      const msL = exp - _today;
      sl = (mfg && exp > mfg) ? Math.max(0, Math.round(msL/(exp-mfg)*100))
                                : Math.max(0, Math.round(msL/(1000*60*60*24*30.44)*10)/10);
    }

    let moveType = null, targetClass = null, reason = '';
    if (skuCls === 'A' && (binCls === 'B' || binCls === 'C')) {
      moveType = 'BRING_DOWN'; targetClass = 'A';
      reason = 'A-class SKU on Class ' + binCls + ' bin — bring to Class A (Level 1–6)';
    } else if ((skuCls === 'B' || skuCls === 'C') && binCls === 'A') {
      moveType = 'MOVE_UP'; targetClass = skuCls;
      reason = skuCls + '-class SKU on Class A bin — free for fast movers';
    }
    if (!moveType) { compliantLines++; return; }

    rows.push({
      moveType, brand, skuCode:entry.skuCode, skuName,
      batch:entry.batch, expDate:entry.expDate, shelfLifePct:sl,
      cls:skuCls, qty:entry.qty, currentBin:entry.bin, currentClass:binCls,
      suggestedBin: nextEmpty(brand, targetClass) || '— No empty ' + targetClass + ' bin',
      targetClass, reason,
    });
  });

  /* Sort: MOVE_UP (P1 — clear Class A) first, then BRING_DOWN, both by qty desc */
  rows.sort((a, b) => {
    if (a.moveType !== b.moveType) return a.moveType === 'MOVE_UP' ? -1 : 1;
    return b.qty - a.qty;
  });

  fsnPlanRows = rows;

  /* KPIs */
  const up     = rows.filter(r => r.moveType === 'MOVE_UP').length;
  const down   = rows.filter(r => r.moveType === 'BRING_DOWN').length;
  const emptyA = Object.entries(emptyPool)
    .filter(([k]) => k.endsWith('|A'))
    .reduce((s, [, v]) => s + v.length, 0);

  document.getElementById('fsn-k-misplaced').innerText = rows.length.toLocaleString();
  document.getElementById('fsn-k-down').innerText      = down.toLocaleString();
  document.getElementById('fsn-k-up').innerText        = up.toLocaleString();
  document.getElementById('fsn-k-empty').innerText     = emptyA.toLocaleString();
  document.getElementById('fsn-k-compliant').innerText = compliantLines.toLocaleString();
  document.getElementById('fsn-k-compliance-pct').innerText =
    (eligibleLines ? (compliantLines / eligibleLines * 100).toFixed(2) : '0.00') + '% compliance';

  if (!rows.length) {
    showResult('fsn-notice', '✅ All SKUs are already in the correct bin class — no moves required!', 'success');
  } else {
    const el = document.getElementById('fsn-notice');
    if (el) el.innerHTML = '';
    toast('FSN plan ready — ' + rows.length + ' moves found', 'info', 3000);
  }

  renderFSNTable();
}

function renderFSNTable() {
  const brandF = (document.getElementById('fsn-filter-brand') || {}).value || '';
  const moveF  = (document.getElementById('fsn-filter-move')  || {}).value || '';
  const clsF   = (document.getElementById('fsn-filter-cls')   || {}).value || '';
  const q      = ((document.getElementById('fsn-search')      || {}).value || '').toLowerCase();

  const slMin = parseInt(document.getElementById('fsn-sl-min')?.value||'0');
  const slMax = parseInt(document.getElementById('fsn-sl-max')?.value||'100');

  const rows = fsnPlanRows.filter(r => {
    if (brandF && r.brand    !== brandF) return false;
    if (moveF  && r.moveType !== moveF)  return false;
    if (clsF   && r.cls      !== clsF)   return false;
    if (slMin > 0 || slMax < 100) {
      if (r.shelfLifePct == null) return false;
      if (r.shelfLifePct < slMin || r.shelfLifePct > slMax) return false;
    }

    if (q && !(r.skuCode+' '+r.skuName+' '+(r.batch||'')+' '+r.currentBin).toLowerCase().includes(q)) return false;
    return true;
  });

  const body = document.getElementById('fsn-body');
  if (!rows.length) {
    body.innerHTML = '<tr class="empty-row"><td colspan="14">No results</td></tr>';
    const fp = document.getElementById('fsnPager'); if(fp) fp.innerHTML = '';
    return;
  }

  const bClr = { Bodywise:'#dbeafe', LittleJoys:'#ede9fe', ManMatters:'#d1fae5' };
  const cClr = { A:'var(--green-l)', B:'var(--amber-l)', C:'var(--red-l)' };
  const cTxt = { A:'var(--green-d)', B:'var(--amber-d)', C:'var(--red-d)' };

  function formatDateSafe(dateVal) {
    if (!dateVal) return '&mdash;';
    const str = String(dateVal).trim();
    
    // Already in DD-MM-YYYY → return as is
    if (/^\d{2}-\d{2}-\d{4}$/.test(str)) return str;
    
    // Try to fix if it's in MM-DD-YYYY or other format
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return d.getDate().toString().padStart(2,'0') + '-' + 
             (d.getMonth()+1).toString().padStart(2,'0') + '-' + 
             d.getFullYear();
    }
    return str;
  }

    function _fsnPg(pr) {
    body.innerHTML = pr.map((r, i) => {
      const isUp = r.moveType === 'MOVE_UP';
      const sl = r.shelfLifePct || 0;
      const slColor = sl >= 100 ? '#166534' : sl >= 75 ? 'var(--green)' : sl >= 50 ? 'var(--amber)' : 'var(--red)';

      return '<tr style="background:' + (isUp ? '#fff8f8' : '#f0fdf4') + '">'
        + '<td style="font-size:11px;color:var(--slate4)">' + (i+1) + '</td>'
        + '<td><span style="font-size:10px;font-weight:700;padding:3px 8px;border-radius:4px;background:' 
          + (isUp ? 'var(--red-l)' : 'var(--green-l)') + ';color:' 
          + (isUp ? 'var(--red-d)' : 'var(--green-d)') + '">'
          + (isUp ? '🔴 Move Up' : '🟢 Bring Down') + '</span></td>'
        + '<td><span style="font-size:10px;font-weight:600;padding:2px 7px;border-radius:4px;background:' 
          + (bClr[r.brand] || 'var(--surface)') + '">' + r.brand + '</span></td>'
        + '<td><span class="sku-badge" style="font-size:10px">' + r.skuCode + '</span></td>'
        + '<td style="font-size:11px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + r.skuName + '</td>'
        + '<td style="font-size:10px;color:var(--slate4)">' + (r.batch || '&mdash;') + '</td>'
        + '<td style="font-size:10px;color:var(--slate4)">' + (r.expDate || '&mdash;') + '</td>'
        + '<td style="text-align:right;font-weight:700;color:' + slColor + '">' + (sl ? sl + '%' : '&mdash;') + '</td>'
        + '<td style="text-align:center"><span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:4px;background:' 
          + (cClr[r.cls] || 'var(--surface)') + ';color:' + (cTxt[r.cls] || 'var(--slate)') + '">' + r.cls + '</span></td>'
        + '<td style="text-align:right;font-weight:700">' + r.qty.toLocaleString() + '</td>'
        + '<td><span class="bin-badge" style="font-size:10px;background:var(--amber-l);color:var(--amber-d)">' + r.currentBin + '</span></td>'
        + '<td><span style="font-size:10px;padding:2px 6px;border-radius:4px;background:' 
          + (cClr[r.currentClass] || 'var(--surface)') + ';color:' + (cTxt[r.currentClass] || 'var(--slate)') + '">Class ' + r.currentClass + '</span></td>'
        + '<td><span class="bin-badge" style="font-size:10px;background:var(--green-l);color:var(--green-d)">' + (r.suggestedBin || '&mdash;') + '</span></td>'
        + '<td style="font-size:10px;color:var(--slate4);max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + r.reason + '</td>'
        + '</tr>';
    }).join('');
  }

  createPager(function(){return rows;}, 1000, 'fsnPager', _fsnPg).refresh();
}

function fsnSliderChange(which) {
  var minInput = document.getElementById('fsn-sl-min');
  var maxInput = document.getElementById('fsn-sl-max');
  var lo = Math.max(0, Math.min(100, parseInt(minInput && minInput.value !== '' ? minInput.value : '0')));
  var hi = Math.max(0, Math.min(100, parseInt(maxInput && maxInput.value !== '' ? maxInput.value : '100')));
  if (Number.isNaN(lo)) lo = 0;
  if (Number.isNaN(hi)) hi = 100;
  // Prevent crossover
  if (lo > hi) {
    if (which === 'min') lo = hi;
    else hi = lo;
  }
  if (minInput) minInput.value = lo;
  if (maxInput) maxInput.value = hi;
  var loEl  = document.getElementById('fsn-sl-lo-val');
  var hiEl  = document.getElementById('fsn-sl-hi-val');
  var range = document.getElementById('fsn-sl-range-label');
  var hint  = document.getElementById('fsn-sl-hint');
  if (loEl)  loEl.innerText  = lo + '%';
  if (hiEl)  hiEl.innerText  = hi + '%';
  var isDefault = (lo === 0 && hi === 100);
  if (range) range.innerText = isDefault ? 'Any' : lo + '% – ' + hi + '%';
  if (hint) {
    if (isDefault) hint.innerText = '';
    else hint.innerText = 'Showing batches with shelf life between ' + lo + '% and ' + hi + '%';
  }
  renderFSNTable();
}

function exportFSNXLSX() {
  if (!fsnPlanRows.length) { alert('No FSN plan — click ↻ Refresh first'); return; }
  const brandF = (document.getElementById('fsn-filter-brand')||{}).value||'';
  const moveF  = (document.getElementById('fsn-filter-move') ||{}).value||'';
  const clsF   = (document.getElementById('fsn-filter-cls')  ||{}).value||'';
  const rows   = fsnPlanRows.filter(r =>
    (!brandF || r.brand === brandF) && (!moveF || r.moveType === moveF) && (!clsF || r.cls === clsF));

  const data = rows.map((r, i) => ({
    'S.No':            i + 1,
    'Move Type':       r.moveType === 'MOVE_UP' ? '🔴 Move Up (B/C off Class A)' : '🟢 Bring Down (A → Class A)',
    'Brand':           r.brand,
    'SKU Code':        r.skuCode,
    'SKU Name':        r.skuName,
    'ABC Class':       r.cls,
    'Qty (units)':     r.qty,
    'Current Bin':     r.currentBin,
    'Current Class':   'Class ' + r.currentClass,
    'Suggested Bin':   r.suggestedBin,
    'Target Class':    'Class ' + r.targetClass,
    'Reason':          r.reason,
    'Generated At':    new Date().toLocaleString('en-IN'),
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:6},{wch:30},{wch:13},{wch:28},{wch:40},{wch:8},{wch:12},{wch:16},{wch:12},{wch:16},{wch:12},{wch:52},{wch:18}];

  const sumRows = [
    ['Metric','Value'],
    ['Generated At', new Date().toLocaleString('en-IN')],
    ['Total Misplaced Lines', fsnPlanRows.length],
    ['🔴 B/C-class to Move Up (free Class A)', fsnPlanRows.filter(r=>r.moveType==='MOVE_UP').length],
    ['🟢 A-class to Bring Down (fill Class A)', fsnPlanRows.filter(r=>r.moveType==='BRING_DOWN').length],
    ['Rows in this export', rows.length],
    ['',''],
    ['Rule 1 (Priority)', 'B/C-class SKU on Class A bin → Move Up — frees picking zone FIRST'],
    ['Rule 2', 'A-class SKU on Class B/C bin → Bring Down — replenish picking zone'],
    ['Brand zones', 'Bodywise=R1-R13 | LittleJoys=R14-R23 | ManMatters=R24-R26 | Overflow=R27+'],
    ['Suggested bins', 'Empty bins only from Bin Master, sorted by lowest level (most accessible) first'],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(sumRows);
  ws2['!cols'] = [{wch:38},{wch:60}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'FSN Replenishment Plan');
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
  XLSX.writeFile(wb, 'FSN_Replenishment_' + new Date().toISOString().slice(0,10) + '.xlsx');
  toast('FSN plan exported — ' + rows.length + ' rows', 'success', 3000);
}


/* ════════════════════════════════════════════════════════
   PLANNING MODULE — BIN CONSOLIDATION  v1.0
   Reads: inventory[], skuMaster[] (already in memory — zero extra API calls)
   ════════════════════════════════════════════════════════ */

let conPlanRows = [];

function initBinConsolidate() {
  if (!inventory || !inventory.length) {
    showResult('con-notice', '⚠ No inventory loaded — upload a dump first.', 'warn');
    return;
  }

  const skuMap = {};
  (skuMaster || []).forEach(s => { skuMap[String(s.code || '').trim()] = s; });

  /* Aggregate GOOD inventory: skuCode+batch → binId → qty (batch-aware) */
  const skuBatchBinMap = {};
  (inventory || []).forEach(r => {
    if (!r.bin || !r.skuCode || r.qty <= 0) return;
    const st = String(r.stockType || '').toUpperCase();
    if (st === 'DAMAGE' || st === 'BAD' || st === 'BAD_INVENTORY' || st === 'BLOCKED') return;
    const key = r.skuCode + '|||' + (r.batch || '');
    if (!skuBatchBinMap[key]) skuBatchBinMap[key] = {};
    skuBatchBinMap[key][r.bin] = (skuBatchBinMap[key][r.bin] || 0) + (r.qty || 0);
  });

  const rows = [];

  Object.entries(skuBatchBinMap).forEach(([compositeKey, binQtyObj]) => {
    const [skuCode, batch] = compositeKey.split('|||');
    const bins = Object.entries(binQtyObj);
    if (bins.length < 2) return;
    const sku      = skuMap[skuCode];
    const casePack = sku ? (Number(sku.casePack) || 0) : 0;
    const skuName  = sku ? (sku.name || '—') : '—';
    const cls      = String((sku && sku.cls) || 'B').toUpperCase();
    const brand    = inferBrandFromPrefix_(skuCode);
    bins.sort((a, b) => b[1] - a[1]);
    const [maxBin, maxQtyRaw] = bins[0];
    const maxQty = Number(maxQtyRaw) || 0;
    const sourceBins = bins.slice(1)
      .map(([fromBin, looseQty]) => ({ bin: fromBin, qty: Number(looseQty) || 0 }))
      .filter(src => casePack > 0 ? src.qty < casePack : src.qty < maxQty);
    if (!sourceBins.length) return;

    const totalLooseQty = sourceBins.reduce((sum, src) => sum + src.qty, 0);
    const combinedQty = maxQty + totalLooseQty;
    rows.push({
      brand, skuCode, batch: batch||'', skuName, cls, casePack,
      fromBin: sourceBins.map(src => src.bin + ' (' + src.qty + ')').join(', '),
      fromBins: sourceBins,
      looseQty: totalLooseQty,
      toBin: maxBin,
      maxBinQty: maxQty,
      combinedQty,
      totalBins: bins.length,
      reason: 'Batch '+(batch||'—')+' · pick '+sourceBins.length+' bin(s), total '+totalLooseQty+' → '+maxBin+' final '+combinedQty,
    });
  });

  /* Stable grouping by SKU + batch */
  rows.sort((a, b) =>
    a.brand.localeCompare(b.brand) ||
    a.skuCode.localeCompare(b.skuCode) ||
    String(a.batch).localeCompare(String(b.batch)) ||
    a.toBin.localeCompare(b.toBin, undefined, {numeric:true})
  );
  conPlanRows = rows;

  /* KPIs */
  const uniqSKUs = new Set(rows.map(r => r.skuCode)).size;
  const uniqBins = new Set(rows.flatMap(r => (r.fromBins||[]).map(src => src.bin))).size;
  const totalQty = rows.reduce((s, r) => s + r.looseQty, 0);

  document.getElementById('con-k-skus').innerText  = uniqSKUs.toLocaleString();
  document.getElementById('con-k-bins').innerText  = uniqBins.toLocaleString();
  document.getElementById('con-k-qty').innerText   = totalQty.toLocaleString();
  document.getElementById('con-k-moves').innerText = rows.length.toLocaleString();

  if (!rows.length) {
    showResult('con-notice', '✅ No loose stock found — all bins are at or above case-pack quantity!', 'success');
  } else {
    const el = document.getElementById('con-notice');
    if (el) el.innerHTML = '';
    toast('Consolidation plan ready — ' + rows.length + ' moves found', 'info', 3000);
  }

  renderConTable();
}

function renderConTable() {
  const brandF = (document.getElementById('con-filter-brand') || {}).value || '';
  const clsF   = (document.getElementById('con-filter-cls')   || {}).value || '';
  const q      = ((document.getElementById('con-search')       || {}).value || '').toLowerCase();

  const rows = conPlanRows.filter(r => {
    if (brandF && r.brand !== brandF) return false;
    if (clsF   && r.cls   !== clsF)   return false;
    if (q && !(r.skuCode + ' ' + r.skuName + ' ' + r.fromBin + ' ' + r.toBin).toLowerCase().includes(q)) return false;
    return true;
  });

  const body = document.getElementById('con-body');
  if (!rows.length) {
    body.innerHTML = '<tr class="empty-row"><td colspan="14">No loose stock found matching these filters</td></tr>';
    return;
  }

  const bClr = { Bodywise:'#dbeafe', LittleJoys:'#ede9fe', ManMatters:'#d1fae5' };
  const cClr = { A:'var(--green-l)', B:'var(--amber-l)', C:'var(--red-l)' };
  const cTxt = { A:'var(--green-d)', B:'var(--amber-d)', C:'var(--red-d)' };

  function _conPg(pr) {
    body.innerHTML = pr.map((r, i) => {
      const fromHtml = (r.fromBins && r.fromBins.length ? r.fromBins : [{bin:r.fromBin, qty:r.looseQty}])
        .map(src => '<span class="bin-badge" style="font-size:10px;background:var(--red-l);color:var(--red-d);margin:1px 3px 1px 0">'+src.bin+' ('+Number(src.qty||0).toLocaleString()+')</span>')
        .join('');
      return '<tr>'
      +'<td style="font-size:11px;color:var(--slate4)">'+(i+1)+'</td>'
      +'<td><span style="font-size:10px;font-weight:600;padding:2px 7px;border-radius:4px;background:'+(bClr[r.brand]||'var(--surface)')+'">'+r.brand+'</span></td>'
      +'<td><span class="sku-badge" style="font-size:10px">'+r.skuCode+'</span></td>'
      +'<td style="font-size:11px;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.skuName+'</td>'
      +'<td style="text-align:center"><span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:4px;background:'+(cClr[r.cls]||'var(--surface)')+';color:'+(cTxt[r.cls]||'var(--slate)')+'">'+r.cls+'</span></td>'
      +'<td style="text-align:center;font-size:11px;color:var(--indigo);font-weight:700">'+(r.casePack||'—')+'</td>'
      +'<td style="font-size:10px;color:var(--slate4)">'+(r.batch||'—')+'</td>'
      +'<td style="max-width:260px">'+fromHtml+'</td>'
      +'<td style="text-align:right;font-weight:700;color:var(--red)">'+r.looseQty.toLocaleString()+'</td>'
      +'<td><span class="bin-badge" style="font-size:10px;background:var(--green-l);color:var(--green-d)">'+r.toBin+'</span></td>'
      +'<td style="text-align:right;font-weight:700;color:var(--green)">'+r.maxBinQty.toLocaleString()+'</td>'
      +'<td style="text-align:right;font-weight:700;color:var(--indigo)">'+r.combinedQty.toLocaleString()+'</td>'
      +'<td style="text-align:center;font-size:11px">'+r.totalBins+' bins</td>'
      +'<td style="font-size:10px;color:var(--slate4);max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.reason+'</td>'
      +'</tr>';
    }).join('');
  }
  const cp = document.getElementById('conPager'); if(cp) cp.innerHTML='';
  createPager(function(){return rows;}, 1000, 'conPager', _conPg).refresh();
}

function exportConXLSX() {
  if (!conPlanRows.length) { alert('No consolidation plan — click ↻ Refresh first'); return; }
  const brandF = (document.getElementById('con-filter-brand')||{}).value||'';
  const clsF   = (document.getElementById('con-filter-cls')  ||{}).value||'';
  const rows   = conPlanRows.filter(r => (!brandF||r.brand===brandF) && (!clsF||r.cls===clsF));

  const data = rows.map((r, i) => ({
    'S.No':                i + 1,
    'Brand':               r.brand,
    'SKU Code':            r.skuCode,
    'SKU Name':            r.skuName,
    'ABC Class':           r.cls,
    'Case Pack (units)':   r.casePack || '—',
    'Pick FROM Bin':       r.fromBin,
    'Pick FROM Details':   (r.fromBins||[]).map(src => src.bin + ' (' + src.qty + ')').join(', '),
    'Loose Qty':           r.looseQty,
    'Consolidate TO Bin':  r.toBin,
    'Max Bin Qty':         r.maxBinQty,
    'Combined Qty':        r.combinedQty,
    'Total Bins (SKU)':    r.totalBins,
    'Reason':              r.reason,
    'Generated At':        new Date().toLocaleString('en-IN'),
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:6},{wch:13},{wch:28},{wch:40},{wch:8},{wch:14},{wch:50},{wch:50},{wch:10},{wch:16},{wch:12},{wch:14},{wch:14},{wch:55},{wch:18}];

  const uniqSKUs = new Set(rows.map(r=>r.skuCode)).size;
  const totalQty = rows.reduce((s,r)=>s+r.looseQty,0);
  const sumRows = [
    ['Metric','Value'],
    ['Generated At', new Date().toLocaleString('en-IN')],
    ['Unique SKUs with Loose Stock', uniqSKUs],
    ['Total Loose Units', totalQty],
    ['Move Suggestions (rows)', rows.length],
    ['',''],
    ['Definition','Loose = bin qty < Case Pack for that SKU'],
    ['Target bin','Bin with highest qty for the same SKU (confirm space before moving)'],
    ['Excluded','BAD_INVENTORY / DAMAGE / BLOCKED stock'],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(sumRows);
  ws2['!cols'] = [{wch:36},{wch:55}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bin Consolidation Plan');
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
  XLSX.writeFile(wb, 'Bin_Consolidation_' + new Date().toISOString().slice(0,10) + '.xlsx');
  toast('Consolidation plan exported — ' + rows.length + ' rows', 'success', 3000);
}


/* ════════════════════════════════════════════════════════
   EXPIRY WATCH — Expired & Near-Expiry Inventory
   Reads inventory[] already in memory — zero GAS calls
   ════════════════════════════════════════════════════════ */

var expiryWatchRows = [];

function initExpiryWatch() {
  if (!inventory || !inventory.length) {
    document.getElementById('ew-body').innerHTML = '<tr class="empty-row"><td colspan="11">No inventory loaded — upload a dump first</td></tr>';
    return;
  }
  var today = new Date(); today.setHours(0,0,0,0);
  function parseExp(s) {
    if (!s) return null;
    var str = String(s).trim();
    var m = str.match(/^(\d\d)-(\d\d)-(\d{4})$/);
    if (m) return new Date(+m[3], +m[2]-1, +m[1]);
    var d = new Date(str); return isNaN(d.getTime()) ? null : d;
  }
  function brandOf(code) {
    var p = String(code||'').slice(0,4).toUpperCase();
    if (p==='MWBW') return 'Bodywise';
    if (p==='MWLJ') return 'LittleJoys';
    if (p==='MWMM') return 'ManMatters';
    return 'Other';
  }

  // Aggregate by bin+sku+batch (GOOD stock only)
  var agg = {};
  (inventory || []).forEach(function(r) {
    if (!r.bin || !r.skuCode || r.qty <= 0) return;
    var st = String(r.stockType||'').toUpperCase();
    if (st === 'DAMAGE' || st === 'BAD' || st === 'BAD_INVENTORY' || st === 'BLOCKED') return;
    if (!r.expDate) return;
    var exp = parseExp(r.expDate); if (!exp) return;
    var days = Math.round((exp - today) / (1000*60*60*24));
    if (days > 90) return; // Only expired or near-expiry
    var key = r.bin + '|||' + r.skuCode + '|||' + (r.batch||'');
    if (!agg[key]) agg[key] = {
      skuCode:r.skuCode, skuName:r.skuName||'', batch:r.batch||'',
      expDate:r.expDate, bin:r.bin, days:days,
      brand:brandOf(r.skuCode), qty:0
    };
    agg[key].qty += (r.qty||0);
  });

  expiryWatchRows = Object.values(agg).sort(function(a,b){ return a.days - b.days; });

  // KPIs
  var expired = expiryWatchRows.filter(function(r){return r.days < 0;}).length;
  var near    = expiryWatchRows.filter(function(r){return r.days >= 0;}).length;
  var qty     = expiryWatchRows.reduce(function(s,r){return s+r.qty;}, 0);
  var bins    = new Set(expiryWatchRows.map(function(r){return r.bin;})).size;
  document.getElementById('ew-k-expired').innerText = expired.toLocaleString();
  document.getElementById('ew-k-near').innerText    = near.toLocaleString();
  document.getElementById('ew-k-qty').innerText     = qty.toLocaleString();
  document.getElementById('ew-k-bins').innerText    = bins.toLocaleString();

  if (!expiryWatchRows.length) {
    document.getElementById('ew-body').innerHTML = '<tr class="empty-row"><td colspan="11">No expired or near-expiry stock found — all good!</td></tr>';
  } else {
    renderExpiryWatch();
    toast('Expiry Watch: ' + expired + ' expired, ' + near + ' near-expiry batches', expired ? 'error' : 'warn', 4000);
  }
}

function renderExpiryWatch() {
  var typeF  = (document.getElementById('ew-filter-type')  || {}).value || '';
  var brandF = (document.getElementById('ew-filter-brand') || {}).value || '';
  var q      = ((document.getElementById('ew-search')      || {}).value || '').toLowerCase();

  var rows = expiryWatchRows.filter(function(r) {
    if (typeF === 'expired' && r.days >= 0) return false;
    if (typeF === 'near'    && r.days <  0) return false;
    if (brandF && r.brand !== brandF) return false;
    if (q && (r.skuCode + ' ' + r.skuName + ' ' + r.bin + ' ' + (r.batch||'')).toLowerCase().indexOf(q) === -1) return false;
    return true;
  });

  var body = document.getElementById('ew-body');
  if (!rows.length) {
    body.innerHTML = '<tr class="empty-row"><td colspan="11">No results match filters</td></tr>';
    return;
  }

  body.innerHTML = rows.map(function(r, i) {
    var isExp = r.days < 0;
    var statusBadge = isExp
      ? '<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:var(--red-l);color:var(--red-d)">🔴 EXPIRED</span>'
      : '<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:var(--amber-l);color:var(--amber-d)">🟡 NEAR</span>';
    var daysText = isExp
      ? '<span style="color:var(--red);font-weight:700">' + Math.abs(r.days) + 'd ago</span>'
      : '<span style="color:var(--amber);font-weight:700">' + r.days + 'd</span>';
    var bClr = {Bodywise:'#dbeafe',LittleJoys:'#ede9fe',ManMatters:'#d1fae5'};
    return '<tr style="background:' + (isExp ? '#fff8f8' : '#fffbeb') + '">'
      + '<td style="font-size:11px;color:var(--slate4)">' + (i+1) + '</td>'
      + '<td>' + statusBadge + '</td>'
      + '<td><span style="font-size:10px;font-weight:600;padding:2px 7px;border-radius:4px;background:' + (bClr[r.brand]||'var(--surface)') + '">' + r.brand + '</span></td>'
      + '<td><span class="sku-badge" style="font-size:10px">' + r.skuCode + '</span></td>'
      + '<td style="font-size:11px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + r.skuName + '">' + r.skuName + '</td>'
      + '<td style="font-size:11px">' + (r.batch||'—') + '</td>'
      + '<td style="font-size:11px">' + r.expDate + '</td>'
      + '<td style="text-align:center">' + daysText + '</td>'
      + '<td><span class="bin-badge" style="font-size:10px">' + r.bin + '</span></td>'
      + '<td style="text-align:right;font-weight:700">' + r.qty.toLocaleString() + '</td>'
      + '<td style="font-size:10px;color:var(--slate4)">Move to quarantine</td>'
      + '</tr>';
  }).join('');
}

// Poll every 5 minutes while app is open
// pollSyncStatus starts in _completeLogin() after user logs in
