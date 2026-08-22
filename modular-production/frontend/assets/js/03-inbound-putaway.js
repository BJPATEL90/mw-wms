function setGRNMode(mode,el){
  grnMode=mode;
  document.querySelectorAll('.mode-tab').forEach(t=>t.classList.remove('active'));
  if(el) el.classList.add('active');
  document.getElementById('grnUploadCard').style.display=mode==='upload'?'block':'none';
  document.getElementById('grnManualCard').style.display=mode==='manual'?'block':'none';
}

function initGRNSection(){
  const no='GRN-'+String(grnCounter).padStart(5,'0');
  document.getElementById('grnNo').value=no;
  if(!document.getElementById('grnDate').value) document.getElementById('grnDate').valueAsDate=new Date();
  renderGRNTable();
}

/* ════════════════════ CALCULATION ENGINE ════════════════════ */
/*
 * Logic:
 *  totalUnits = e.g. 1000
 *  casePack   = units per box, e.g. 24  → totalBoxes = ceil(1000/24) = 42 boxes (41 full + 1 with 16 loose)
 *  boxPallet  = boxes per pallet/shelf, e.g. 25
 *  Pallets: floor(42/25) = 1 full pallet (25 boxes), remainder = 17 boxes → 2nd pallet
 *  Each bin gets one pallet's worth of boxes.
 *  Bin 1: 25 boxes × 24 units = 600 units
 *  Bin 2: 17 boxes: 16 full (16×24=384) + 1 box with 16 loose = 384+16 = 400 units
 */
function computeBreakdown(qty,casePack,boxPallet){
  qty=Number(qty)||0;casePack=Number(casePack)||0;boxPallet=Number(boxPallet)||0;
  if(qty<=0) return null;
  const fullBoxes=casePack>0?Math.floor(qty/casePack):0;
  const looseUnits=casePack>0?qty%casePack:qty;
  const totalBoxes=fullBoxes+(looseUnits>0?1:0);  // total cartons (last one may be partial)
  let pallets=[];
  if(boxPallet>0){
    let rem=totalBoxes;
    while(rem>0){
      const bInThisPallet=Math.min(rem,boxPallet);
      // count units in this pallet
      let units=0;
      for(let b=0;b<bInThisPallet;b++){
        const boxIdx=totalBoxes-rem+b;  // which box number (0-indexed)
        if(boxIdx<fullBoxes) units+=casePack;  // full box
        else units+=looseUnits;                 // last partial box
      }
      pallets.push({boxes:bInThisPallet,units});
      rem-=bInThisPallet;
    }
  }
  return {qty,casePack,boxPallet,fullBoxes,looseUnits,totalBoxes,pallets};
}

function calcBreakdown(){
  const qty=Number(document.getElementById('grnQty').value)||0;
  const cp=Number(document.getElementById('grnCasePack').value)||0;
  const bp=Number(document.getElementById('grnBoxPallet').value)||0;
  const preview=document.getElementById('calcPreview');
  const txt=document.getElementById('calcText');
  if(!qty){preview.style.display='none';return;}
  const b=computeBreakdown(qty,cp,bp);
  if(!b){preview.style.display='none';return;}
  preview.style.display='block';
  let html='';
  if(cp>0) html+=`<div>Total units: <b>${qty}</b> ÷ Case Pack <b>${cp}</b> = <b>${b.fullBoxes}</b> full boxes + <b>${b.looseUnits}</b> loose units = <b>${b.totalBoxes}</b> total boxes/cartons</div>`;
  else html+=`<div>Total units: <b>${qty}</b> (No Case Pack — each unit treated as 1 box)</div>`;
  if(bp>0&&b.pallets.length>0){
    html+=`<div style="margin-top:4px;color:var(--slate4)">Box/Pallet = ${bp} → <b>${b.pallets.length} bin(s)/shelf(s)</b> will be allocated:</div>`;
    b.pallets.forEach((p,i)=>{ html+=`<div style="margin-left:12px">🗂 Bin ${i+1}: <b>${p.boxes} boxes</b> = <b>${p.units} units</b></div>`; });
  } else if(bp<=0){
    html+=`<div style="margin-top:4px;color:var(--amber)">⚠ Enter Box/Pallet to see bin allocation preview</div>`;
  }
  txt.innerHTML=html;
}

/* ════════════════════ SKU LOOKUP ════════════════════ */
function lookupSKU(){
  const code=document.getElementById('grnSKU').value.trim();
  const cpHint=document.getElementById('cpHint');
  const bpHint=document.getElementById('bpHint');
  const msg=document.getElementById('grnLookupMsg');
  if(!code){clearGRNSKUFields();return;}
  const sku=skuMaster.find(s=>s.code.toLowerCase()===code.toLowerCase());
  if(sku){
    document.getElementById('grnSKUName').value=sku.name;
    document.getElementById('grnClass').value=sku.cls||'';
    if(sku.casePack){document.getElementById('grnCasePack').value=sku.casePack;cpHint.innerText='(from master)';}
    else{document.getElementById('grnCasePack').value='';cpHint.innerText='(enter manually)';}
    if(sku.boxPallet){document.getElementById('grnBoxPallet').value=sku.boxPallet;bpHint.innerText='(from master)';}
    else{document.getElementById('grnBoxPallet').value='';bpHint.innerText='(enter manually)';}
    msg.innerText='';calcBreakdown();
  }else{
    clearGRNSKUFields();
    msg.innerText='⚠ SKU not in master — add it first via SKU Master';
    cpHint.innerText='';bpHint.innerText='';
  }
}
/** Cancel current GRN and release all Reserved bins back to Empty */
function openAmendStockModal(){
  if(userRole !== 'ADMIN'){ toast('Admin only','error'); return; }
  document.getElementById('amendStockModal').classList.add('open');
}

function doAmendStock(){
  const bin    = document.getElementById('amend_bin').value.trim().toUpperCase();
  const sku    = document.getElementById('amend_sku').value.trim();
  const batch  = document.getElementById('amend_batch').value.trim();
  const qty    = Number(document.getElementById('amend_qty').value);
  const type   = document.getElementById('amend_type').value;
  const reason = document.getElementById('amend_reason').value.trim() || 'Manual amendment';
  const msg    = document.getElementById('amendMsg');

  if(!bin || !sku || !batch){ msg.innerHTML='<span style="color:var(--red)">Bin, SKU and Batch are required</span>'; return; }
  if(isNaN(qty) || qty < 0){ msg.innerHTML='<span style="color:var(--red)">Enter valid qty (0 or more)</span>'; return; }

  // Find row in local inventory
  const row = inventory.find(r => r.bin===bin && r.skuCode===sku && r.batch===batch && r.stockType===type);
  if(!row){ msg.innerHTML='<span style="color:var(--amber)">Row not found in local inventory. It will be added.</span>'; }

  const oldQty = row ? row.qty : 0;
  const diff   = qty - oldQty;

  // Update local inventory
  if(row){ row.qty = qty; }
  else {
    const skuM = skuMaster.find(s => s.code === sku);
    inventory.push({ bin, skuCode:sku, skuName:skuM?skuM.name:'', batch,
      mfgDate:'', expDate:'', mrp:'', stockType:type, skuStatus:'ACTIVE', qty, ean:'' });
  }
  lsSet('inv', inventory); dumpRows = inventory;

  // Log to adjustment log
  const ts = new Date();
  adjLog.push({
    time: ts.toISOString(), bin, skuCode: sku, batch, stockType: type,
    fromBin: bin, toBin: bin, moveQty: diff, remarks: reason + ' (Admin amend)',
    by: userName, type: 'AMEND'
  });
  saveAdjLog();

  // Sync to sheet
  msg.innerHTML = '<span style="color:var(--amber)">Saving to sheet…</span>';
  callSheet('saveInventoryDump', { rows: inventory.map(r=>[
    r.skuCode,r.skuName,r.batch,r.mfgDate,r.expDate,r.mrp,r.bin,r.stockType,r.skuStatus,r.qty,r.ean,'',''
  ]), mode:'replace' },
    function(r){
      renderCurrentInventory(); renderDumpTable();
      toast('Stock amended: ' + sku + ' in ' + bin + ' → ' + qty + ' units','success',4000);
      msg.innerHTML = '<span style="color:var(--green)">✅ Amendment applied. Old: '+oldQty+' → New: '+qty+' (Δ'+diff+')</span>';
    },
    function(e){ msg.innerHTML = '<span style="color:var(--red)">Local updated but sheet error: '+e+'</span>'; }
  );
}

function cancelGRNAndReleaseBins(){
  if(!putawayPlan.length && !grnLines.length){
    toast('No active GRN or putaway to cancel','warn'); return;
  }
  if(!confirm('Cancel current GRN and release all Reserved bins back to Empty?')) return;

  // Release bins that were Reserved by this putaway
  putawayPlan.forEach(p => {
    if(!p.binId || p.binId === 'NO_BIN') return;
    const bm = binMaster.find(b => b.id === p.binId);
    if(bm && bm.status === 'Reserved') bm.status = 'Empty';
  });
  const _cancelPlanId = currentGRN ? currentGRN.grnNo : '';

  // Clear state
  putawayPlan = [];
  grnLines    = [];
  currentGRN  = null;
  _overflowRows = [];
  saveAll();
  renderGRNTable();
  loadDashboard();
  toast('GRN cancelled — Reserved bins released back to Empty','success',3000);

  // Cancel plan on sheet
  if(_cancelPlanId){
    callSheet('cancelPutawayPlan',{planId:_cancelPlanId,releasedBy:userName},function(){},function(){});
  }
}

/** Change bin for a specific pallet in the putaway plan */
function changePalletBin(palletIdx){
  const row = putawayPlan[palletIdx];
  if(!row){ toast('Pallet not found','error'); return; }

  const brand    = row.binBrand || getBrandFromSKUCode_(row.sku) || '';
  const fsn      = row.cls || 'B';
  const usedBins = new Set(putawayPlan.map(p=>p.binId).filter(b=>b&&b!=='NO_BIN'&&b!==row.binId));

  const sameBrandFsn = binMaster.filter(b=>b.status==='Empty'&&!usedBins.has(b.id)&&b.brand===brand&&b.fsn===fsn);
  const sameBrandOth = binMaster.filter(b=>b.status==='Empty'&&!usedBins.has(b.id)&&b.brand===brand&&b.fsn!==fsn);
  const anyEmpty     = binMaster.filter(b=>b.status==='Empty'&&!usedBins.has(b.id)&&b.brand!==brand);

  if(!sameBrandFsn.length&&!sameBrandOth.length&&!anyEmpty.length){
    toast('No empty bins available','error',4000); return;
  }

  const el = document.getElementById('changeBinModal');
  if(el) el.remove();

  // Build select options
  let opts = '';
  const addGroup = (bins, label) => {
    if(!bins.length) return;
    opts += '<optgroup label="' + label + '">';
    bins.forEach(b=>{ opts += '<option value="' + b.id + '">' + b.id + ' — ' + b.brand + ' FSN-' + b.fsn + '</option>'; });
    opts += '</optgroup>';
  };
  addGroup(sameBrandFsn, 'Same Brand + FSN (recommended)');
  addGroup(sameBrandOth, 'Same Brand, Different FSN');
  addGroup(anyEmpty,     'Other Brand / Overflow');

  // Build modal using DOM API — avoids all quote escaping issues
  const overlay = document.createElement('div');
  overlay.id = 'changeBinModal';
  overlay.className = 'modal-overlay open';

  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.style.maxWidth = '440px';
  modal.innerHTML =
    '<div class="modal-header">' +
      '<div class="modal-title">Change Bin — Pallet #' + (row.palletNo||1) + '</div>' +
      '<button class="modal-close" id="cbmClose">&#10005;</button>' +
    '</div>' +
    '<div class="modal-body">' +
      '<div class="notice info" style="margin-bottom:12px">' +
        '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>' +
        '<div>SKU: <b>' + row.sku + '</b> | Batch: <b>' + (row.batch||'—') + '</b><br>' +
             'Current bin: <b>' + row.binId + '</b><br>' +
             'Brand: <b>' + (brand||'—') + '</b> | FSN: <b>' + fsn + '</b></div>' +
      '</div>' +
      '<label class="form-label">Select New Bin</label>' +
      '<select id="changeBinSelect" class="form-control" style="font-family:DM Mono,monospace;margin-bottom:8px">' + opts + '</select>' +
      '<div style="font-size:11px;color:var(--slate4);margin-bottom:8px">Only empty bins shown. Bins already used in this plan are excluded.</div>' +
      '<div id="changeBinMsg" style="font-size:12px;min-height:14px"></div>' +
    '</div>' +
    '<div class="modal-footer">' +
      '<button class="btn btn-ghost" id="cbmCancel">Cancel</button>' +
      '<button class="btn btn-primary" id="cbmConfirm">Confirm Change</button>' +
    '</div>';

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // Attach events via addEventListener — no inline onclick needed
  document.getElementById('cbmClose').addEventListener('click',  function(){ overlay.remove(); });
  document.getElementById('cbmCancel').addEventListener('click', function(){ overlay.remove(); });
  document.getElementById('cbmConfirm').addEventListener('click', function(){ applyBinChange(palletIdx); });
}


function applyBinChange(palletIdx){
  const row  = putawayPlan[palletIdx];
  const sel  = document.getElementById('changeBinSelect');
  const msgEl = document.getElementById('changeBinMsg');
  if(!sel||!sel.value){ if(msgEl) msgEl.innerHTML='<span style="color:var(--red)">Select a bin</span>'; return; }
  const newBinId = sel.value;
  const bm = binMaster.find(b=>b.id===newBinId);
  if(!bm||bm.status!=='Empty'){
    if(msgEl) msgEl.innerHTML='<span style="color:var(--red)">Bin no longer available</span>'; return;
  }
  const oldBinId = row.binId;
  const oldBm = binMaster.find(b=>b.id===oldBinId);
  if(oldBm&&oldBm.status==='Reserved') oldBm.status='Empty';
  bm.status='Reserved';
  putawayPlan[palletIdx].binId    = newBinId;
  putawayPlan[palletIdx].binBrand = bm.brand||'';
  putawayPlan[palletIdx].binFsn   = bm.fsn||'';
  putawayPlan[palletIdx].warning  = false;
  putawayPlan[palletIdx].warnReason='';
  // Sync bin swap to sheet
  const _planIdForSwap = currentGRN ? currentGRN.grnNo : '';
  if(_planIdForSwap){
    showProcessing('Updating bin reservation…');
    callSheet('swapPutawayBin',{planId:_planIdForSwap,lineNo:putawayPlan[palletIdx].lineNo,oldBinId:oldBinId,newBinId:newBinId,updatedBy:userName},
      function(r){ hideProcessing(); if(!r||!r.success) toast('Bin swap sync: '+(r&&r.message||'error'),'warn',3000); },
      function(e){ hideProcessing(); toast('Bin swap error: '+e,'warn',3000); }
    );
  }
  const m=document.getElementById('changeBinModal'); if(m) m.remove();
  saveAll(); showPutaway();
  toast('Pallet #'+(row.palletNo||1)+' → '+newBinId,'success',3000);
}

function clearGRNSKUFields(){
  ['grnSKUName','grnBoxPallet','grnCasePack'].forEach(id=>{const e=document.getElementById(id);if(e)e.value='';});
  document.getElementById('grnClass').value='';
  document.getElementById('cpHint').innerText='';
  document.getElementById('bpHint').innerText='';
  document.getElementById('calcPreview').style.display='none';
}

function lookupEditSKU(){
  const code=document.getElementById('eg_sku').value.trim();
  const sku=skuMaster.find(s=>s.code.toLowerCase()===code.toLowerCase());
  document.getElementById('eg_skuName').value=sku?sku.name:'';
  if(sku){
    if(sku.casePack) document.getElementById('eg_cp').value=sku.casePack;
    if(sku.boxPallet) document.getElementById('eg_bp').value=sku.boxPallet;
    document.getElementById('eg_class').value=sku.cls||'';
  }
}

/* ════════════════════ ADD GRN LINE ════════════════════ */
function addGRNLine(){
  const sku=document.getElementById('grnSKU').value.trim();
  const skuName=document.getElementById('grnSKUName').value.trim();
  const batch=document.getElementById('grnBatch').value.trim();
  const qty=Number(document.getElementById('grnQty').value);
  const cp=Number(document.getElementById('grnCasePack').value)||0;
  const bp=Number(document.getElementById('grnBoxPallet').value)||0;
  const cls=document.getElementById('grnClass').value;
  const msg=document.getElementById('grnLookupMsg');
  if(!sku){msg.innerText='⚠ SKU Code required';return;}
  if(!skuName){msg.innerText='⚠ SKU not found in master';return;}
  if(!batch){msg.innerText='⚠ Batch No required';return;}
  if(qty<=0){msg.innerText='⚠ Quantity must be > 0';return;}
  msg.innerText='';
  const bd=computeBreakdown(qty,cp,bp);
  grnLines.push({sku,skuName,batch,qty,casePack:cp||null,boxPallet:bp||null,cls:cls||'B',breakdown:bd});
  saveAll();renderGRNTable();
  ['grnSKU','grnSKUName','grnBatch','grnQty','grnBoxPallet','grnCasePack'].forEach(id=>{const e=document.getElementById(id);if(e)e.value='';});
  document.getElementById('grnClass').value='';
  document.getElementById('cpHint').innerText='';document.getElementById('bpHint').innerText='';
  document.getElementById('calcPreview').style.display='none';
}

function renderGRNTable(){
  const body=document.getElementById('grnBody');
  document.getElementById('grnLineCount').innerText=grnLines.length+' line'+(grnLines.length!==1?'s':'');
  document.getElementById('grnClearBtn').style.display=grnLines.length?'inline-flex':'none';
  if(!grnLines.length){body.innerHTML='<tr class="empty-row"><td colspan="11">No lines added yet</td></tr>';return;}
  body.innerHTML=grnLines.map((l,i)=>{
    const bd=l.breakdown||computeBreakdown(l.qty,l.casePack,l.boxPallet);
    return `<tr>
      <td style="color:var(--slate5);font-size:11px">${i+1}</td>
      <td><span class="sku-badge">${l.sku}</span></td>
      <td style="font-size:12px">${l.skuName}</td>
      <td>${l.batch}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700">${l.qty}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${bd?bd.fullBoxes:'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${bd?bd.looseUnits:'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${bd?bd.totalBoxes:'—'}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace;font-weight:700;color:var(--indigo)">${bd&&bd.pallets?bd.pallets.length:'—'}</td>
      <td><span class="status-badge sb-${(l.cls||'b').toLowerCase()}">${l.cls||'—'}</span></td>
      <td><button class="btn btn-ghost btn-xs" onclick="removeGRNLine(${i})" style="color:var(--red);border-color:#fecaca">✕</button></td>
    </tr>`;
  }).join('');
}

function removeGRNLine(i){grnLines.splice(i,1);saveAll();renderGRNTable();}
function clearGRNLines(){if(confirm('Clear all GRN lines?')){grnLines=[];saveAll();renderGRNTable();}}

/* ════════════════════ GRN EXCEL UPLOAD ════════════════════ */
function handleGRNExcelDrop(e){e.preventDefault();e.currentTarget.classList.remove('drag');const f=e.dataTransfer.files[0];if(f)parseGRNExcel({target:{files:[f]}});}
function parseGRNExcel(e){
  const file=e.target.files[0];if(!file)return;
  const msg=document.getElementById('grnExcelMsg');
  msg.innerText='Reading…';
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){msg.innerText='No rows found';return;}
      const heads=Object.keys(rows[0]);
      const fa=aliases=>heads.find(h=>aliases.includes(String(h).toLowerCase().trim()));
      const skuCol=fa(['sku code','sku_code','sku','material code','material_code','item code']);
      const batchCol=fa(['batch','batch no','batch_no','lot','lot no']);
      const qtyCol=fa(['qty','quantity','units','total qty']);
      const cpCol=fa(['case pack','casepack','case_pack','units/box','units per box']);
      const bpCol=fa(['box/pallet','box per pallet','boxes per pallet','pallet','shelf','boxes/shelf']);
      const clsCol=fa(['class','classification','fsn','category']);
      if(!skuCol||!qtyCol){msg.innerText='⚠ SKU Code and Qty columns required';return;}
      let added=0,skipped=0;
      rows.forEach(r=>{
        const code=String(r[skuCol]||'').trim();
        const batch=batchCol?String(r[batchCol]||'').trim():'';
        const qty=Number(String(r[qtyCol]||'0').replace(/,/g,''));
        if(!code||qty<=0){skipped++;return;}
        const sku=skuMaster.find(s=>s.code.toLowerCase()===code.toLowerCase());
        const cp=cpCol?Number(r[cpCol]||0)||(sku?.casePack||0):sku?.casePack||0;
        const bp=bpCol?Number(r[bpCol]||0)||(sku?.boxPallet||0):sku?.boxPallet||0;
        const cls=clsCol?String(r[clsCol]||'').trim()||(sku?.cls||'B'):(sku?.cls||'B');
        const bd=computeBreakdown(qty,cp,bp);
        grnLines.push({sku:code,skuName:sku?.name||code,batch,qty,casePack:cp||null,boxPallet:bp||null,cls,breakdown:bd});
        added++;
      });
      saveAll();renderGRNTable();
      msg.innerText=`✓ ${added} line(s) loaded${skipped?`, ${skipped} skipped`:''}`;
    }catch(err){msg.innerText='Error: '+err.message;}
  };
  reader.readAsArrayBuffer(file);
  e.target.value='';
}

function downloadGRNTemplate(){
  const ws=XLSX.utils.aoa_to_sheet([
    ['SKU Code','Batch No','Qty','Case Pack (Units/Box)','Box/Pallet (Boxes/Shelf)','Classification'],
    ['MWLJNTP.0003.B0_N','BATCH001',1000,24,32,'A'],
    ['MWMMHRP.0004.AAAA.B0_N','BATCH002',500,42,24,'A'],
  ]);
  ws['!cols']=[{wch:26},{wch:14},{wch:8},{wch:22},{wch:28},{wch:14}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'GRN Template');XLSX.writeFile(wb,'GRN_Upload_Template.xlsx');
}

/* ════════════════════ SAVE GRN ════════════════════ */
function saveGRN(){
  if(!grnLines.length){alert('Add at least one line');return;}
  const no=document.getElementById('grnNo').value;
  const date=document.getElementById('grnDate').value;
  const supplier=document.getElementById('grnSupplier').value||'';
  const vehicle=document.getElementById('grnVehicle').value||'';
  if(!date){alert('GRN Date required');return;}
  const entry={grnNo:no,date,supplier,vehicle,lines:[...grnLines],savedBy:userName,savedAt:new Date().toISOString()};
  // update existing or add
  const existIdx=grnList.findIndex(g=>g.grnNo===no);
  if(existIdx>=0) grnList[existIdx]=entry; else grnList.push(entry);
  currentGRN=entry;
  grnCounter++;
  saveAll();
  showResult('grnResult','GRN #'+no+' saved. Now click "Generate Putaway" to allocate bins.','success');
  loadDashboard();
  // Sync to Google Sheet
  showProcessing('Saving GRN to Google Sheet…');
  callSheet('saveGRNToSheet', entry,
    r=>{ hideProcessing(); if(r&&r.success) showResult('grnResult','GRN #'+no+' saved and synced to Google Sheet.','success'); },
    ()=>{}
  );
}

/* ════════════════════ GRN LIST ════════════════════ */
function renderGRNList(){
  const q=String(document.getElementById('grnListSearch')?.value||'').toLowerCase();
  const body=document.getElementById('grnListBody');
  const filtered=grnList.filter(g=>`${g.grnNo} ${g.supplier||''} ${g.vehicle||''}`.toLowerCase().includes(q)).slice().reverse();
  if(!filtered.length){body.innerHTML='<tr class="empty-row"><td colspan="8">No GRN entries found</td></tr>';return;}
  body.innerHTML=filtered.map((g,i)=>{
    const totalQty=(g.lines||[]).reduce((s,l)=>s+Number(l.qty||0),0);
    return `<tr>
      <td><code style="font-family:'DM Mono',monospace;font-size:11px">${g.grnNo}</code></td>
      <td>${g.date}</td>
      <td>${g.supplier||'—'}</td>
      <td>${g.vehicle||'—'}</td>
      <td style="text-align:center">${(g.lines||[]).length}</td>
      <td style="text-align:right;font-weight:700;font-family:'DM Mono',monospace">${totalQty.toLocaleString()}</td>
      <td>${g.savedBy||'—'}</td>
      <td style="display:flex;gap:6px">
        <button class="btn btn-warn btn-xs" onclick="openEditGRN('${g.grnNo}')">✏ Edit</button>
        <button class="btn btn-indigo btn-xs" onclick="reloadGRNToPutaway('${g.grnNo}')">📦 Putaway</button>
      </td>
    </tr>`;
  }).join('');
}

function openEditGRN(grnNo){
  const grn=grnList.find(g=>g.grnNo===grnNo);
  if(!grn){alert('GRN not found');return;}
  editGRNIdx=grnList.indexOf(grn);
  editGRNLines=JSON.parse(JSON.stringify(grn.lines||[]));
  document.getElementById('eg_no').value=grn.grnNo;
  document.getElementById('eg_date').value=grn.date;
  document.getElementById('eg_supplier').value=grn.supplier||'';
  document.getElementById('eg_vehicle').value=grn.vehicle||'';
  ['eg_sku','eg_skuName','eg_batch','eg_qty','eg_cp','eg_bp'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('eg_class').value='';
  document.getElementById('egMsg').innerText='';
  renderEditGRNLines();
  document.getElementById('editGRNModal').classList.add('open');
}

function renderEditGRNLines(){
  const body=document.getElementById('editGRNLines');
  if(!editGRNLines.length){body.innerHTML='<tr class="empty-row"><td colspan="9">No lines</td></tr>';return;}
  body.innerHTML=editGRNLines.map((l,i)=>`<tr>
    <td style="color:var(--slate5);font-size:11px">${i+1}</td>
    <td><input class="tbl-input" value="${l.sku}" onchange="editGRNLines[${i}].sku=this.value" style="width:160px"></td>
    <td><input class="tbl-input" value="${l.skuName}" onchange="editGRNLines[${i}].skuName=this.value" style="width:140px"></td>
    <td><input class="tbl-input" value="${l.batch}" onchange="editGRNLines[${i}].batch=this.value" style="width:90px"></td>
    <td style="text-align:right"><input class="tbl-input" type="number" value="${l.qty}" onchange="editGRNLines[${i}].qty=Number(this.value);editGRNLines[${i}].breakdown=computeBreakdown(Number(this.value),l.casePack,l.boxPallet)" style="width:70px;text-align:right"></td>
    <td style="text-align:right"><input class="tbl-input" type="number" value="${l.casePack||''}" onchange="editGRNLines[${i}].casePack=Number(this.value)||null" style="width:60px;text-align:right"></td>
    <td style="text-align:right"><input class="tbl-input" type="number" value="${l.boxPallet||''}" onchange="editGRNLines[${i}].boxPallet=Number(this.value)||null" style="width:60px;text-align:right"></td>
    <td><span class="status-badge sb-${(l.cls||'b').toLowerCase()}">${l.cls||'—'}</span></td>
    <td><button class="btn btn-ghost btn-xs" onclick="editGRNLines.splice(${i},1);renderEditGRNLines()" style="color:var(--red);border-color:#fecaca">✕</button></td>
  </tr>`).join('');
}

function addEditGRNLine(){
  const sku=document.getElementById('eg_sku').value.trim();
  const skuName=document.getElementById('eg_skuName').value.trim()||sku;
  const batch=document.getElementById('eg_batch').value.trim();
  const qty=Number(document.getElementById('eg_qty').value);
  const cp=Number(document.getElementById('eg_cp').value)||null;
  const bp=Number(document.getElementById('eg_bp').value)||null;
  const cls=document.getElementById('eg_class').value||'B';
  if(!sku||qty<=0){document.getElementById('egMsg').innerText='SKU Code and Qty required';return;}
  editGRNLines.push({sku,skuName,batch,qty,casePack:cp,boxPallet:bp,cls,breakdown:computeBreakdown(qty,cp,bp)});
  renderEditGRNLines();
  document.getElementById('egMsg').innerText='';
}

function saveEditGRN(){
  if(editGRNIdx===null){alert('No GRN loaded');return;}
  const no=document.getElementById('eg_no').value;
  grnList[editGRNIdx]={...grnList[editGRNIdx],date:document.getElementById('eg_date').value,supplier:document.getElementById('eg_supplier').value,vehicle:document.getElementById('eg_vehicle').value,lines:[...editGRNLines],editedBy:userName,editedAt:new Date().toISOString()};
  saveAll();closeModal('editGRNModal');renderGRNList();loadDashboard();
  alert('GRN #'+no+' updated successfully');
}

function deleteGRN(){
  if(editGRNIdx===null)return;
  const no=grnList[editGRNIdx]?.grnNo;
  if(!confirm('Delete GRN #'+no+'? This cannot be undone.'))return;
  grnList.splice(editGRNIdx,1);
  saveAll();closeModal('editGRNModal');renderGRNList();loadDashboard();
}

function reloadGRNToPutaway(grnNo){
  const grn=grnList.find(g=>g.grnNo===grnNo);
  if(!grn){alert('GRN not found');return;}
  currentGRN=grn;
  grnLines=[...grn.lines];
  saveAll();
  generatePutaway();
}

/* ════════════════════ PUTAWAY ════════════════════ */
/* ─────────────────────────────────────────────────────
   Brand-aware putaway plan generator
   ───────────────────────────────────────────────────── */

// Overflow rows selected by user (persists within session)
let _overflowRows = [];

function generatePutaway(){
  if(!grnLines.length){ alert('No GRN lines to process'); return; }
  if(!currentGRN){
    const no = document.getElementById('grnNo').value || 'GRN-DRAFT';
    currentGRN = { grnNo: no, date: document.getElementById('grnDate').value || new Date().toISOString().slice(0,10), lines:[...grnLines] };
  }

  // Check if any lines might need overflow — fetch options first
  callSheet('getOverflowOptions', {},
    function(res){
      _runPutawayAllocation(res && res.rows ? res.rows : []);
    },
    function(){
      // Offline — run without overflow info
      _runPutawayAllocation([]);
    }
  );
}

function _runPutawayAllocation(overflowRowOptions){
  const plan     = [];
  const usedBins = new Set();
  let   needsOverflow = false;
  const warnings = [];

  grnLines.forEach((line, li) => {
    const bd      = line.breakdown || computeBreakdown(line.qty, line.casePack, line.boxPallet);
    const brand   = getBrandFromSKUCode_(line.sku);
    const fsnPref = line.cls || 'B';
    const pallets = (bd && bd.pallets && bd.pallets.length) ? bd.pallets : [{ units: line.qty, boxes: bd ? bd.totalBoxes : 0 }];

    pallets.forEach((pallet, pi) => {
      const result = findBestBinLocal_(brand, fsnPref, usedBins, _overflowRows);

      if(!result){
        needsOverflow = true;
        warnings.push('No bin for ' + line.sku + ' (Brand: ' + (brand||'Unknown') + ')');
        plan.push({
          ...line, lineNo: li+1, palletNo: pi+1,
          binId: 'NO_BIN', binBrand: brand||'?',
          allocQty: pallet.units, allocBoxes: pallet.boxes,
          isFullPallet: pallet.boxes === Number(line.boxPallet||0),
          warning: true, warnReason: 'No empty bin — zone full'
        });
      } else {
        usedBins.add(result.binId);
        // Mark bin reserved in local binMaster
        const bm = binMaster.find(b => b.id === result.binId);
        if(bm) bm.status = 'Reserved';
        plan.push({
          ...line, lineNo: li+1, palletNo: pi+1,
          binId: result.binId, binBrand: result.zone,
          binFsn: result.fsn, isOverflow: result.isOverflow,
          suggestion: result.suggestion,
          allocQty: pallet.units, allocBoxes: pallet.boxes,
          isFullPallet: pallet.boxes === Number(line.boxPallet||0),
          warning: false
        });
      }
    });
  });

  // If overflow needed and no overflow rows selected → prompt user
  if(needsOverflow && _overflowRows.length === 0){
    showOverflowSelector_(overflowRowOptions, warnings, function(selectedRows){
      _overflowRows = selectedRows;
      _runPutawayAllocation(overflowRowOptions); // re-run with overflow rows
    });
    return;
  }

  putawayPlan = plan; saveAll();
  showPutaway();
  // Save plan to Google Sheet (source of truth)
  const _planGrnNo = currentGRN ? currentGRN.grnNo : ('GRN-'+Date.now());
  showProcessing('Saving putaway plan to Google Sheet…');
  callSheet('savePutawayPlan', { rows: plan, grnNo: _planGrnNo, createdBy: userName },
    function(r){
      hideProcessing();
      if(r && r.success){ toast('Putaway plan saved — visible to all users','success',3000); }
      else toast('Plan sheet sync: '+(r&&r.message||'error'),'warn',3000);
    },
    function(e){ hideProcessing(); toast('Plan sheet sync error: '+e,'warn',3000); }
  );
  showSection('putaway', document.querySelectorAll('.nav-item')[3]);
  grnLines = []; saveAll(); renderGRNTable();
  document.getElementById('grnNo').value = 'GRN-' + String(grnCounter).padStart(5,'0');
  loadDashboard();
}

/** Identify brand from SKU code prefix */
/** Infer brand from Bin ID on client side — mirrors server inferBrandFromBin_ */
function inferBrandClient_(binId){
  const m = String(binId).match(/^R(\d+)-C(\d+)/i);
  if(!m) return 'Unknown';
  const row=parseInt(m[1]), col=parseInt(m[2]);
  for(const [brand,zones] of Object.entries(BRAND_ZONES_CLIENT)){
    if(zones.some(z=>row>=z.rMin&&row<=z.rMax&&col>=z.cMin&&col<=z.cMax)) return brand;
  }
  if(row>=27&&row<=35) return 'Overflow';
  return 'Unknown';
}

function getBrandFromSKUCode_(skuCode){
  skuCode = String(skuCode||'').trim().toUpperCase();
  if(skuCode.startsWith('MWBW')) return 'Bodywise';
  if(skuCode.startsWith('OWN'))  return 'OWN';
  if(skuCode.startsWith('MWLJ')) return 'LittleJoys';
  if(skuCode.startsWith('MWMM')) return 'ManMatters';
  return null; // unknown — will fall back to FSN-only search
}

/** Brand zone definitions (mirrors server-side BRAND_ZONES) */
const BRAND_ZONES_CLIENT = {
  'Bodywise':   [ {rMin:1,rMax:6,cMin:1,cMax:10}, {rMin:7,rMax:13,cMin:1,cMax:20} ],
  'OWN':        [ {rMin:1,rMax:6,cMin:11,cMax:20} ],
  'LittleJoys': [ {rMin:14,rMax:23,cMin:1,cMax:20} ],
  'ManMatters': [ {rMin:24,rMax:26,cMin:1,cMax:20} ],
};

function parseBinIdClient_(binId){
  const m = String(binId).match(/^R(\d+)-C(\d+)-(\d+)$/i);
  return m ? {row:parseInt(m[1]),col:parseInt(m[2]),level:parseInt(m[3])} : null;
}

function binInBrandZone_(binId, brand){
  const p = parseBinIdClient_(binId);
  if(!p) return false;
  const zones = BRAND_ZONES_CLIENT[brand];
  if(!zones) return false;
  return zones.some(z => p.row>=z.rMin && p.row<=z.rMax && p.col>=z.cMin && p.col<=z.cMax);
}

/** Find best empty bin from local binMaster — brand-aware with FSN fallback */
function findBestBinLocal_(brand, fsnPref, usedBins, overflowRows){
  const fsns = fsnPref==='A' ? ['A','B','C'] : fsnPref==='B' ? ['B','C','A'] : ['C','B','A'];

  // Search brand zone first
  if(brand && BRAND_ZONES_CLIENT[brand]){
    for(const fsn of fsns){
      const bin = binMaster.find(b =>
        b.status === 'Empty' &&
        b.fsn    === fsn &&
        !usedBins.has(b.id) &&
        binInBrandZone_(b.id, brand)
      );
      if(bin) return { binId: bin.id, fsn: bin.fsn, zone: brand, isOverflow: false,
                       suggestion: brand + ' zone — FSN ' + fsn };
    }
  }

  // Brand zone full — try overflow rows
  if(overflowRows && overflowRows.length){
    for(const fsn of fsns){
      const bin = binMaster.find(b => {
        if(b.status !== 'Empty' || b.fsn !== fsn || usedBins.has(b.id)) return false;
        const p = parseBinIdClient_(b.id);
        return p && overflowRows.includes(p.row);
      });
      if(bin){
        const p = parseBinIdClient_(bin.id);
        return { binId: bin.id, fsn: bin.fsn, zone: 'Overflow R'+(p&&p.row),
                 isOverflow: true, suggestion: 'Overflow zone (brand zone full)' };
      }
    }
  }

  return null;
}

/** Show overflow row selector modal */
function showOverflowSelector_(rowOptions, warnings, onConfirm){
  const existing = document.getElementById('overflowModal');
  if(existing) existing.remove();

  const availHtml = rowOptions.length
    ? rowOptions.map(r => '<label style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)"><input type="checkbox" value="'+r.row+'" '+(r.emptyBins===0?'disabled':'')+'> <span>Row R'+r.row+' &nbsp;<span style="color:var(--slate4);font-size:11px">('+r.emptyBins+' empty bins)</span></span></label>').join('')
    : '<p style="color:var(--amber);font-size:12px">Could not load overflow availability — select rows manually.</p>' +
      [27,28,29,30,31,32,33,34,35].map(r=>'<label style="display:flex;align-items:center;gap:8px;padding:6px 0"><input type="checkbox" value="'+r+'"> Row R'+r+'</label>').join('');

  const modal = document.createElement('div');
  modal.id        = 'overflowModal';
  modal.className = 'modal-overlay open';
  modal.innerHTML = `
    <div class="modal" style="max-width:440px">
      <div class="modal-header">
        <div class="modal-title">⚠ Brand Zone Full — Select Overflow Rows</div>
      </div>
      <div class="modal-body">
        <div class="notice warn" style="margin-bottom:12px">
          <svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
          <div><b>No empty bins in brand zone for:</b><br>${warnings.join('<br>')}<br><br>Select overflow rows (R27–R35) to use:</div>
        </div>
        <div id="overflowCheckboxes" style="max-height:280px;overflow-y:auto">${availHtml}</div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-ghost" onclick="document.getElementById('overflowModal').remove();_overflowRows=[];">Cancel</button>
        <button class="btn btn-primary" onclick="
          const rows=[...document.querySelectorAll('#overflowCheckboxes input:checked')].map(x=>parseInt(x.value));
          if(!rows.length){alert('Select at least one overflow row');return;}
          document.getElementById('overflowModal').remove();
          overflowConfirmCallback(rows);
        ">Use Selected Rows</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  window.overflowConfirmCallback = onConfirm;
}

function getFreeBin(cls, usedBins){
  // Legacy fallback — used only if brand not identified
  return binMaster.find(b=>b.status==='Empty'&&b.fsn===cls&&!usedBins.has(b.id))||
         binMaster.find(b=>b.status==='Empty'&&!usedBins.has(b.id))||null;
}

/* ════════════════════ PUTAWAY SHEET SYNC ════════════════════ */

function loadPutawayFromSheet(){
  showProcessing('Loading putaway plan…');
  callSheet('loadActivePutawayPlan', {},
    function(data){
      hideProcessing();
      if(!data || !data.success){ toast('Could not load plan: '+(data&&data.message||'error'),'warn'); return; }
      if(!data.rows || !data.rows.length){ toast('No active putaway plan found','info',3000); return; }
      // Map sheet rows back to local putawayPlan format
      putawayPlan = data.rows.map(function(r){
        return {
          sku: r.sku, skuName: r.skuName, batch: r.batch,
          binId: r.binId, allocQty: r.allocQty, allocBoxes: r.allocBoxes,
          lineNo: r.lineNo, palletNo: r.palletNo,
          cls: r.cls, binBrand: r.binBrand,
          isFullPallet: r.isFullPallet, warning: r.warning,
          _confirmed: r.status === 'CONFIRMED',
          _planId: r.planId
        };
      });
      lsSet('putaway', putawayPlan);
      if(currentGRN && data.planId) currentGRN.grnNo = data.planId;
      showPutaway();
      toast('Putaway plan loaded — '+putawayPlan.length+' pallets','success',2000);
    },
    function(e){ hideProcessing(); toast('Load error: '+e,'warn'); }
  );
}

function confirmPutawayRow_(palletIdx){
  const row = putawayPlan[palletIdx];
  if(!row){ toast('Row not found','error'); return; }
  if(row._confirmed){ toast('Already confirmed','info',2000); return; }
  const planId = (currentGRN && currentGRN.grnNo) || row._planId || '';
  showProcessing('Confirming bin '+row.binId+'…');
  callSheet('confirmPutawayRow', { planId, lineNo: row.lineNo, binId: row.binId, confirmedBy: userName },
    function(r){
      hideProcessing();
      if(!r || !r.success){ toast('Confirm failed: '+(r&&r.message||'error'),'error'); return; }
      putawayPlan[palletIdx]._confirmed = true;
      // Mark bin Occupied locally
      const bm = binMaster.find(b=>b.id===row.binId);
      if(bm) bm.status = 'Occupied';
      lsSet('putaway', putawayPlan);
      saveAll();
      showPutaway();
      if(r.allConfirmed){
        toast('✅ All pallets confirmed! Putaway complete.','success',5000);
        putawayPlan = []; currentGRN = null; lsSet('putaway',[]); lsSet('curGRN',null); loadDashboard();
      } else {
        toast('Bin '+row.binId+' confirmed — Occupied','success',2000);
      }
    },
    function(e){ hideProcessing(); toast('Confirm error: '+e,'error'); }
  );
}

function confirmAllPutaway_(){
  if(!putawayPlan.length){ toast('No active putaway plan','warn'); return; }
  const unconfirmed = putawayPlan.filter(r=>!r._confirmed&&!r.warning);
  if(!unconfirmed.length){ toast('All pallets already confirmed','info',2000); return; }
  if(!confirm('Confirm all '+unconfirmed.length+' remaining pallets as placed? All bins will be marked Occupied.')) return;
  const planId = (currentGRN && currentGRN.grnNo) || (putawayPlan[0]&&putawayPlan[0]._planId) || '';
  showProcessing('Confirming all pallets…');
  callSheet('confirmAllPutaway', { planId, confirmedBy: userName },
    function(r){
      hideProcessing();
      if(!r||!r.success){ toast('Confirm all failed: '+(r&&r.message||'error'),'error'); return; }
      // Mark all bins Occupied locally
      putawayPlan.forEach(p=>{
        if(!p.warning){ p._confirmed=true; const bm=binMaster.find(b=>b.id===p.binId); if(bm) bm.status='Occupied'; }
      });
      saveAll();
      toast('✅ Putaway complete! All bins marked Occupied.','success',5000);
      putawayPlan=[]; currentGRN=null; lsSet('putaway',[]); lsSet('curGRN',null); loadDashboard();
      showPutaway();
    },
    function(e){ hideProcessing(); toast('Error: '+e,'error'); }
  );
}

function showPutaway(){
  const el=document.getElementById('putawayRows');
  const grnEl=document.getElementById('putawayGRN');
  const notice=document.getElementById('putawayNotice');
  if(currentGRN) grnEl.innerText='GRN: '+currentGRN.grnNo;
  if(!putawayPlan.length){el.innerHTML='<p style="color:var(--slate4);font-size:13px">No putaway plan yet. Generate from GRN Entry.</p>';return;}
  const warns=putawayPlan.filter(r=>r.warning).length;
  notice.innerHTML=warns?`<div class="notice warn"><svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>${warns} pallet(s) have no bin available.</div>`:'<div class="notice success"><svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>All pallets allocated!</div>';
  // Group by lineNo+SKU for cleaner view
  let lastKey='';
  el.innerHTML=putawayPlan.map((r,i)=>{
    const key=r.sku+'|'+r.batch;
    const isNewGroup=key!==lastKey;lastKey=key;
    const groupHeader=isNewGroup?`<div style="font-size:11px;font-weight:700;color:var(--slate3);margin:${i>0?'16px':0} 0 6px;padding-bottom:4px;border-bottom:1px solid var(--border)">
      <span class="sku-badge">${r.sku}</span> &nbsp;${r.skuName}&nbsp;|&nbsp;Batch: ${r.batch}&nbsp;|&nbsp;Total: ${putawayPlan.filter(x=>x.sku===r.sku&&x.batch===r.batch).reduce((s,x)=>s+x.allocQty,0)} units
    </div>`:'';
    const cp=Number(r.casePack||0);
    return groupHeader+`<div class="putaway-row" style="${r.warning?'border-color:#fecaca;background:#fff8f8':''}">
      <div class="putaway-bin" style="display:flex;flex-direction:column;align-items:center;gap:6px">
        <span style="${r.warning?'color:var(--red);font-size:13px;font-weight:700':''}">${r.warning?'⚠ NO BIN':r.binId}</span>
        <div style="display:flex;gap:6px;width:100%;justify-content:space-between;margin-top:4px">
          <button onclick="changePalletBin(${i})" style="flex:1;font-size:11px;padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;background:var(--surface);cursor:pointer;color:var(--slate3);font-weight:600" title="Change bin for this pallet">✎ Change</button>
          ${!r.warning?`<button onclick="confirmPutawayRow_(${i})" style="flex:1;font-size:11px;padding:5px 8px;border:1.5px solid #059669;border-radius:6px;background:${r._confirmed?'#d1fae5':'var(--surface)'};cursor:pointer;color:${r._confirmed?'#065f46':'#059669'};font-weight:600" title="Mark this pallet as placed">${r._confirmed?'✅ Done':'✓ Confirm'}</button>`:'<div style="flex:1"></div>'}
        </div>
      </div>
      <div class="putaway-info">
        <div class="pi-sku">
          Pallet #${r.palletNo||'1'} &nbsp;|&nbsp;
          <span class="status-badge" style="background:#dbeafe;color:#1e40af;font-size:9px">${r.binBrand||r.cls||'—'}</span>
          &nbsp;FSN: <span class="status-badge sb-${(r.cls||'b').toLowerCase()}">${r.cls||'—'}</span>
          &nbsp;|&nbsp; ${r.isOverflow?'<span style="color:var(--amber);font-weight:700;font-size:11px">⚠ OVERFLOW</span>':r.isFullPallet?'<span style="color:var(--green);font-weight:700;font-size:11px">✓ FULL PALLET</span>':'<span style="color:var(--amber);font-weight:700;font-size:11px">PARTIAL PALLET</span>'}
          ${r.suggestion?'<span style="color:var(--slate4);font-size:10px;margin-left:6px">'+r.suggestion+'</span>':''}
        </div>
        <div class="pi-detail">${r.allocBoxes} boxes${cp>0?` (${r.allocBoxes-1} full + 1 partial)`:''} on this pallet</div>
      </div>
      <div class="putaway-breakdown">
        <div style="font-weight:700;font-size:10px;color:var(--slate4);margin-bottom:4px">THIS BIN</div>
        <div style="color:var(--slate)"><b>${r.allocQty}</b> units</div>
        <div style="color:var(--slate4)">${r.allocBoxes} boxes × ${cp>0?cp+' u/box':'-'}</div>
      </div>
      <div class="putaway-qty"><div class="pq-num">${r.allocQty}</div><div class="pq-label">UNITS</div></div>
    </div>`;
  }).join('');
}

/* ════════════════════ INVENTORY DUMP ════════════════════ */
