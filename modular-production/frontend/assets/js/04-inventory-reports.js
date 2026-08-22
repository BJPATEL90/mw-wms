const DUMP_ALIASES={
  skuCode:  ['sku code','sku_code','material code','material_code','sku','item code','product code'],
  skuName:  ['sku name','sku_name','material name','material_name','description','name','product name'],
  batch:    ['batch','batch no','batch_no','batchno','lot','lot no'],
  bin:      ['bin','bin no','bin_no','bin id','location','shelf','bin / shelf'],
  qty:      ['qty','quantity','balance','balance qty','stock qty','units'],
  stockType:['stock type','stock_type','type','category'],
  ean:      ['ean','barcode','ean code','ean no'],
  mfgDate:  ['mfg date','mfg_date','manufacture date','mfd date','mfg','mfg. date'],
  expDate:  ['exp date','exp_date','expiry date','expiry','best before','exp. date'],
  mrp:      ['mrp','mrp (rs)','mrp rs','price','unit price','selling price'],
  skuStatus:['sku status','sku_status','product status','item status'],
};
function findA(heads,als){return heads.find(h=>als.includes(String(h).toLowerCase().trim()))||null;}

function handleDumpDrop(e){e.preventDefault();e.currentTarget.classList.remove('drag');parseDumpFile({target:{files:[e.dataTransfer.files[0]]}});}
function parseDumpFile(e){
  const file=e.target.files[0];if(!file)return;
  const msg=document.getElementById('dumpMsg');msg.innerText='Reading…';
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){msg.innerText='No rows found';return;}
      const heads=Object.keys(rows[0]);
      const cm={};Object.keys(DUMP_ALIASES).forEach(k=>cm[k]=findA(heads,DUMP_ALIASES[k]));
      if(!cm.skuCode||!cm.qty){msg.innerText='⚠ SKU Code and Qty columns required';return;}
      dumpRows=rows.map(r=>({
        skuCode:String(r[cm.skuCode]||'').trim(),
        brand:getBrandFromSKUCode_(String(r[cm.skuCode]||'').trim()),
        skuName:cm.skuName?String(r[cm.skuName]||'').trim():'',
        batch:cm.batch?String(r[cm.batch]||'').trim():'',
        bin:cm.bin?normalizeBinIdForMaster(r[cm.bin]):'',
        qty:Number(String(r[cm.qty]||'0').replace(/,/g,'')),
        stockType:normalizeStockType_(cm.stockType?String(r[cm.stockType]||'').trim():''),
        ean:cm.ean?String(r[cm.ean]||'').trim():'',
        mfgDate:normDate(cm.mfgDate?r[cm.mfgDate]:''),    // v7: resolve Excel serials
        expDate:normDate(cm.expDate?r[cm.expDate]:''),
        mrp:cm.mrp?String(r[cm.mrp]||'').trim():'',
        skuStatus:cm.skuStatus?String(r[cm.skuStatus]||'ACTIVE').trim().toUpperCase():'ACTIVE',
      })).filter(r=>r.skuCode&&r.qty>0);

      inventory=dumpRows;saveAll();

      // ── UPDATE BIN STATUS FROM DUMP ──
      updateBinsFromDump();

      msg.innerText=`✓ ${dumpRows.length} rows loaded from ${file.name}. Bin status updated.`;
      renderDumpSummary();renderDumpTable();loadDashboard();
      // Sync to Google Sheet
      msg.innerText+=' Syncing to Google Sheet…';
      callSheet('saveInventoryDump', {rows:dumpRows, mode:'replace'},
        r=>{
          msg.innerText=`✓ ${dumpRows.length} rows synced to Google Sheet.`;
          _ciFlushMode = false; // v9: new dump uploaded — allow CI to rebuild
          // v6: Auto-sync Current Inventory after dump upload
          callSheet('syncCurrentInventory', {},
            r2=>{ if(r2&&r2.success){ renderCurrentInventory(); msg.innerText+=` Current Inventory updated (${r2.count} rows).`; } },
            ()=>{}
          );
        },
        e=>{ msg.innerText+=` ⚠ Sheet sync failed: ${e}`; }
      );
    }catch(err){msg.innerText='Error: '+err.message;}
  };
  reader.readAsArrayBuffer(file);
  e.target.value='';
}

function updateBinsFromDump(){
  // Build map: which normalized master bins have stock.
  const binsWithStock={};
  dumpRows.forEach(r=>{
    const binId=normalizeBinIdForMaster(r.bin);
    if(binId){
      if(!binsWithStock[binId]) binsWithStock[binId]={qty:0,skus:new Set()};
      binsWithStock[binId].qty+=r.qty;
      binsWithStock[binId].skus.add(r.skuCode);
    }
  });

  let updated=0,ignoredBins=0;
  const masterIdMap={};
  binMaster.forEach(b=>{ const id=normalizeBinIdForMaster(b.id); if(id) masterIdMap[id]=b; });

  binMaster.forEach(b=>{
    const masterId=normalizeBinIdForMaster(b.id);
    if(masterId&&binsWithStock[masterId]){
      // Bin has stock -> Occupied
      if(b.status!=='Occupied'){b.status='Occupied';updated++;}
      b.skuFromDump=[...binsWithStock[masterId].skus].slice(0,2).join(', ');
    } else {
      // No stock in dump -> mark Empty (unless reserved by putaway)
      if(b.status==='Blocked') return;
      const reservedByPutaway=putawayPlan.some(p=>normalizeBinIdForMaster(p.binId)===masterId&&!p.warning);
      if(!reservedByPutaway&&b.status!=='Reserved'){
        if(b.status==='Occupied'){b.status='Empty';updated++;}
      }
      b.skuFromDump='';
    }
  });

  // Dump bins that are not in master are ignored. Bin Master is source of truth.
  Object.keys(binsWithStock).forEach(binId=>{
    if(!masterIdMap[binId]) ignoredBins++;
  });

  saveAll();
  const notice=document.getElementById('dumpBinUpdateNotice');
  notice.innerHTML=`<div class="notice success" style="margin-top:8px"><svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>Bin status auto-updated: <b>${updated}</b> existing bin(s) changed. <b>${ignoredBins}</b> dump bin(s) ignored because they are not in Bin Master.</div>`;
}

function renderDumpSummary(){
  const total=dumpRows.length,totalQty=dumpRows.reduce((s,r)=>s+r.qty,0),skus=new Set(dumpRows.map(r=>r.skuCode)).size;
  const dmg=dumpRows.filter(r=>['DAMAGE','DAMAGED'].includes(r.stockType)).length;
  document.getElementById('dumpSummary').innerHTML=`<span class="chip chip-total">${total} Lines</span><span class="chip chip-ok">${skus} SKUs</span><span class="chip chip-ok">Total: ${totalQty.toLocaleString()}</span>${dmg?`<span class="chip chip-err">${dmg} Damage Rows</span>`:''}`;
}
function renderDumpTable(){
  const q=String(document.getElementById('dumpSearch')?.value||'').toLowerCase();
  const tf=String(document.getElementById('dumpTypeFilter')?.value||'');
  const rows=dumpRows.filter(r=>`${r.skuCode} ${r.skuName} ${r.batch} ${r.bin}`.toLowerCase().includes(q)&&(!tf||r.stockType===tf));
  const body=document.getElementById('dumpBody');
  if(!rows.length){body.innerHTML='<tr class="empty-row"><td colspan="6">No inventory rows found</td></tr>';return;}
  const visible=rows.slice(0,500);
  const more=rows.length-visible.length;
  body.innerHTML=(more?`<tr><td colspan="6" style="font-size:11px;color:var(--slate4);background:#f8fafc">Showing first 500 of ${rows.length.toLocaleString()} inventory rows. Use search/filter to narrow the list.</td></tr>`:'')+
  visible.map(r=>`<tr>
    <td><span class="sku-badge">${r.skuCode}</span></td>
    <td style="font-size:12px">${r.skuName}</td>
    <td>${r.batch||'—'}</td>
    <td><span class="bin-badge">${r.bin||'—'}</span></td>
    <td><span class="status-badge ${['GOOD','UNRESTRICTED'].includes(r.stockType)?'sb-ok':r.stockType==='DAMAGE'||r.stockType==='DAMAGED'?'sb-diff':'sb-reserved'}">${r.stockType}</span></td>
    <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700">${r.qty.toLocaleString()}</td>
  </tr>`).join('');
}
function downloadDumpTemplate(){
  const ws=XLSX.utils.aoa_to_sheet([
    ['SKU Code','SKU Name','Batch No','MFG Date','EXP Date','MRP','Bin / Shelf','Stock Type','SKU Status','Qty','EAN'],
    ['MWLJNTP.0003.B0_N','LJ NutriMix 2+ 350gm Chocolate Jar','BATCH001','01-Jan-2024','31-Dec-2025','399','R1-C1-001','GOOD','ACTIVE',100,''],
    ['MWLJNTP.00059.B0_N','LJ NutriMix 7+ 350gm Chocolate Jar','BATCH002','15-Feb-2024','28-Feb-2026','449','R1-C2-001','GOOD','ACTIVE',200,''],
    ['MWMMHRP.0004.AAAA.B0_N','MM Hair Gummies 30mcg - 30s','BATCH003','01-Mar-2024','31-Mar-2026','599','R1-C3-001','DAMAGE','ACTIVE',5,''],
  ]);
  ws['!cols']=[{wch:28},{wch:34},{wch:12},{wch:12},{wch:12},{wch:8},{wch:14},{wch:10},{wch:10},{wch:8},{wch:14}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Inventory Dump');XLSX.writeFile(wb,'Inventory_Dump_Template.xlsx');
}
function exportDumpXLSX(){
  if(!dumpRows.length){alert('Upload inventory dump first');return;}
  const ws=XLSX.utils.json_to_sheet(dumpRows.map(r=>({'SKU Code':r.skuCode,'SKU Name':r.skuName,'Batch':r.batch,'Bin':r.bin,'Stock Type':r.stockType,'Qty':r.qty})));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Inventory');XLSX.writeFile(wb,'WMS_Inventory.xlsx');
}

/* ════════════════════ BIN TABLE ════════════════════ */
function renderBinTable(){
  const cleanBins=normalizeBinMasterRows(binMaster);
  if(cleanBins.length!==binMaster.length){ binMaster=cleanBins; lsSet('bin', binMaster); }
  const q=String(document.getElementById('binSearch')?.value||'').toLowerCase();
  const sf=String(document.getElementById('binStatusFilter')?.value||'');
  const cf=String(document.getElementById('binClassFilter')?.value||'');
  const filtered=cleanBins.filter(b=>`${b.id} ${b.zone||''} ${b.fsn}`.toLowerCase().includes(q)&&(!sf||b.status===sf)&&(!cf||b.fsn===cf));
  const total=cleanBins.length,empty=cleanBins.filter(b=>b.status==='Empty').length,occ=cleanBins.filter(b=>b.status==='Occupied').length,res=cleanBins.filter(b=>b.status==='Reserved').length;
  const _blk=cleanBins.filter(b=>b.status==='Blocked').length;
  document.getElementById('binSummaryChips').innerHTML='<span class="chip chip-total">'+total+' Total</span><span class="chip chip-ok">'+empty+' Empty</span><span class="chip chip-err">'+occ+' Occupied</span><span class="chip chip-warn">'+res+' Reserved</span>'+(_blk?'<span class="chip" style="background:#fce4d6;color:#7f1d1d">🚫 '+_blk+' Blocked</span>':'');
  const body=document.getElementById('binBody');
  // v8: binMaster could be empty after flush — show correct message
  if(!cleanBins.length){
    document.getElementById('binSummaryChips').innerHTML='<span class="chip chip-total">0 bins (data flushed)</span>';
    body.innerHTML='<tr class="empty-row"><td colspan="7">No bin data — upload Bin Master or click ↻ Refresh</td></tr>';
    return;
  }
  if(!filtered.length){body.innerHTML='<tr class="empty-row"><td colspan="7">No bins match filter</td></tr>';return;}
  const visible=filtered.slice(0,500);
  const more=filtered.length-visible.length;
  body.innerHTML=(more?`<tr><td colspan="8" style="font-size:11px;color:var(--slate4);background:#f8fafc">Showing first 500 of ${filtered.length.toLocaleString()} bins. Use search/filter to narrow the list.</td></tr>`:'')+
  visible.map(b=>`<tr>
    <td><span class="bin-badge">${b.id}</span></td>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${b.row||'—'}</td>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${b.col||'—'}</td>
    <td><span class="status-badge" style="background:#dbeafe;color:#1e40af;font-size:9px">${b.brand||'—'}</span></td>
    <td><span class="status-badge sb-${(b.fsn||'b').toLowerCase()}">${b.fsn||'—'}</span></td>
    <td style="font-size:10px;color:var(--slate4);max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${b.skuFromDump||'—'}</td>
    <td><span class="status-badge ${b.status==='Empty'?'sb-empty':b.status==='Occupied'?'sb-occupied':'sb-reserved'}">${b.status}</span></td>
    <td>
      <select class="tbl-input" onchange="updateBinStatus('${b.id}',this.value)" style="font-size:11px;width:100px">
        ${['Empty','Occupied','Reserved','Blocked'].map(s=>`<option ${b.status===s?'selected':''}>${s==='Blocked'?'🚫 Blocked':s}</option>`).join('')}
      </select>
    </td>
  </tr>`).join('');
}
function updateBinStatus(id,status){
  var b=binMaster.find(function(x){return x.id===id;});
  if(!b)return;
  var wasBlocked=b.status==='Blocked', goingBlocked=status==='Blocked';
  b.status=status; saveAll(); loadDashboard();
  if(goingBlocked||wasBlocked) toast(goingBlocked?'Bin '+id+' blocked locally':'Bin '+id+' unblocked locally','success',3000);
}
function openAddBinModal(){['nb_id','nb_zone','nb_qty'].forEach(id=>document.getElementById(id).value='');document.getElementById('nbMsg').innerText='';document.getElementById('addBinModal').classList.add('open');}
function saveNewBin(){
  const id=normalizeBinIdForMaster(document.getElementById('nb_id').value);
  if(!id){document.getElementById('nbMsg').innerText='Bin ID required';return;}
  if(binMaster.find(b=>normalizeBinIdForMaster(b.id)===id)){document.getElementById('nbMsg').innerText='Bin ID exists';return;}
  const newBinEntry={id,zone:document.getElementById('nb_zone').value.trim(),brand:document.getElementById('nb_zone').value.trim(),fsn:document.getElementById('nb_fsn').value,capacity:Number(document.getElementById('nb_qty').value)||null,status:document.getElementById('nb_status').value,skuFromDump:''};
  binMaster.push(newBinEntry);
  saveAll();closeModal('addBinModal');renderBinTable();loadDashboard();
  showProcessing('Saving bin to Google Sheet…');
  callSheet('addOrUpdateBinRow',{'Bin ID':id,'Brand Zone':newBinEntry.brand,'FSN Class':newBinEntry.fsn,'Status':newBinEntry.status},
    function(r){ hideProcessing(); if(r&&r.success) toast('Bin '+id+' saved to sheet','success',2000); else toast('Sheet sync: '+(r&&r.message||'error'),'warn',3000); },
    function(e){ hideProcessing(); toast('Sheet sync error: '+e,'warn',3000); }
  );
}
function exportBinXLSX(){
  const ws=XLSX.utils.json_to_sheet(binMaster.map(b=>({'Bin ID':b.id,'Zone':b.zone||'','FSN Class':b.fsn,'Capacity (Boxes)':b.capacity||'','Status':b.status})));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Bin_Master');XLSX.writeFile(wb,'Bin_Master.xlsx');
}

function downloadBinTemplate(){
  // Generate sample rows covering all brand zones
  // Columns: Bin ID | Row | Column No | Level | Brand Zone | FSN Class | Status
  // NOTE: Capacity removed — pallet size comes from SKU Master (Box/Pallet field)
  const samples = [
    ['Bin ID','Row','Column No','Level','Brand Zone','FSN Class (A/B/C)','Status'],
    // Bodywise — R1-R6, C1-C10
    ['R1-C1-001',1,1,1,'Bodywise','A','Empty'],
    ['R1-C1-002',1,1,2,'Bodywise','A','Empty'],
    ['R1-C5-001',1,5,1,'Bodywise','B','Empty'],
    ['R4-C1-001',4,1,1,'Bodywise','A','Empty'],
    // OWN — R1-R6, C11-C12
    ['R1-C11-001',1,11,1,'OWN','B','Empty'],
    ['R2-C12-001',2,12,1,'OWN','B','Empty'],
    // LittleJoys — R14-R23
    ['R14-C1-001',14,1,1,'LittleJoys','A','Empty'],
    ['R18-C10-001',18,10,1,'LittleJoys','B','Empty'],
    // ManMatters — R24-R26
    ['R24-C1-001',24,1,1,'ManMatters','A','Empty'],
    ['R25-C5-001',25,5,1,'ManMatters','B','Empty'],
    // Overflow — R27-R35
    ['R27-C1-001',27,1,1,'Overflow','B','Empty'],
    ['R30-C1-001',30,1,1,'Overflow','C','Empty'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(samples);
  ws['!cols']=[{wch:14},{wch:6},{wch:10},{wch:7},{wch:14},{wch:18},{wch:10}];
  // Bold header row
  const hdr = ws['A1'];
  const wb  = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bin_Master');
  XLSX.writeFile(wb, 'Bin_Master_Template.xlsx');
  toast('Bin Master template downloaded — fill in all bins and upload via ⬆ Bulk Upload', 'info', 5000);
}

const BIN_UPLOAD_ALIASES={
  id:     ['bin id','bin_id','bin no','binid','bin'],
  row:    ['row'],
  col:    ['column no','column','col','col no'],
  level:  ['level'],
  brand:  ['brand zone','brand','zone'],
  fsn:    ['fsn class (a/b/c)','fsn class','fsn','class','classification'],
  status: ['status','bin status']
  // Capacity removed — pallet size comes from SKU Master (Box/Pallet field)
};

function parseBinUpload(e){
  const file=e.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){showBinUploadResult('No data rows found in file.','danger');return;}
      const heads=Object.keys(rows[0]).map(h=>String(h).toLowerCase().trim());
      const fa=als=>Object.keys(rows[0]).find(h=>als.includes(String(h).toLowerCase().trim()))||null;
      const cm={
        id:     fa(BIN_UPLOAD_ALIASES.id),
        row:    fa(BIN_UPLOAD_ALIASES.row),
        col:    fa(BIN_UPLOAD_ALIASES.col),
        level:  fa(BIN_UPLOAD_ALIASES.level),
        brand:  fa(BIN_UPLOAD_ALIASES.brand),
        fsn:    fa(BIN_UPLOAD_ALIASES.fsn),
        status: fa(BIN_UPLOAD_ALIASES.status),
      };
      if(!cm.id){showBinUploadResult('⚠ "Bin ID" column is required. Check your file headers.','danger');return;}
      let added=0,updated=0,skipped=0;
      rows.forEach(r=>{
        const id=normalizeBinIdForMaster(r[cm.id]);if(!id){skipped++;return;}
        // Parse row/col/level from ID if not in separate columns
        const mParsed=id.match(/^R(\d+)-C(\d+)-(\d+)$/i);
        const rowNo  =cm.row?Number(r[cm.row])||'':(mParsed?parseInt(mParsed[1]):'');
        const colNo  =cm.col?Number(r[cm.col])||'':(mParsed?parseInt(mParsed[2]):'');
        const level  =cm.level?Number(r[cm.level])||'':(mParsed?parseInt(mParsed[3]):'');
        const brand  =cm.brand?String(r[cm.brand]||'').trim():inferBrandClient_(id);
        const fsn    =(cm.fsn?String(r[cm.fsn]||'').trim().toUpperCase():'B');
        const validFsn=['A','B','C'].includes(fsn)?fsn:'B';
        const stRaw  =cm.status?String(r[cm.status]||'').trim():'Empty';
        const status =['Empty','Occupied','Reserved'].find(s=>s.toLowerCase()===stRaw.toLowerCase())||'Empty';
        const existing=binMaster.find(b=>normalizeBinIdForMaster(b.id)===id);
        if(existing){
          existing.row=rowNo||existing.row; existing.col=colNo||existing.col;
          existing.level=level||existing.level; existing.brand=brand||existing.brand;
          existing.fsn=validFsn; existing.status=status; updated++;
        } else {
          // Capacity not stored — pallet size comes from SKU Master
          binMaster.push({id,row:rowNo,col:colNo,level,brand,fsn:validFsn,status,skuFromDump:''});
          added++;
        }
      });
      saveAll();renderBinTable();loadDashboard();
      const msg=`✅ Bin Master updated from <b>${file.name}</b> — ${added} added, ${updated} updated, ${skipped} skipped.`;
      showBinUploadResult(msg,'success');
      // Sync to Google Sheet if connected
      if(typeof google!=='undefined'&&google.script){
        syncBinMasterToSheet();
      }
    }catch(err){showBinUploadResult('Error reading file: '+err.message,'danger');}
  };
  reader.readAsArrayBuffer(file);e.target.value='';
}

function showBinUploadResult(msg,type){
  const el=document.getElementById('binUploadResult');
  el.style.display='block';
  el.className='notice '+(type||'info');
  el.innerHTML=`<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>${msg}`;
}

function syncBinMasterToSheet(){
  const byBin={};
  binMaster.forEach(b=>{
    const id=normalizeBinIdForMaster(b.id);
    if(!id) return;
    byBin[id]={'Bin ID':id,'Zone':b.zone||'','FSN Class':b.fsn,'Capacity (Boxes)':b.capacity||'','Status':b.status};
  });
  const rows=Object.values(byBin);
  callSheet('bulkSaveBinMaster', rows,
    r=>showBinUploadResult(r&&r.success?'✅ Bin Master synced to Google Sheet ('+rows.length+' bins)':'⚠ Sync issue: '+(r&&r.message||'unknown'), r&&r.success?'success':'warn'),
    err=>showBinUploadResult('⚠ Sheet sync error: '+err,'warn')
  );
}

/* ════════════════════ SKU MASTER ════════════════════ */
function renderSKUTable(){
  const q=String(document.getElementById('skuSearch')?.value||'').toLowerCase();
  const filtered=skuMaster.filter(s=>`${s.code} ${s.name}`.toLowerCase().includes(q));
  const body=document.getElementById('skuBody');
  // v8: skuMaster could be empty after flush
  if(!skuMaster.length){
    body.innerHTML='<tr class="empty-row"><td colspan="6">No SKU data — upload SKU Master or click ↻ Refresh</td></tr>';
    return;
  }
  if(!filtered.length){body.innerHTML='<tr class="empty-row"><td colspan="6">No SKUs match search</td></tr>';return;}
  body.innerHTML=filtered.map(s=>`<tr>
    <td><span class="sku-badge">${s.code}</span></td>
    <td style="font-size:12px">${s.name}</td>
    <td><span class="status-badge" style="background:#dbeafe;color:#1e40af;font-size:9px">${s.brand||getBrandFromSKUCode_(s.code)||'—'}</span></td>
    <td style="text-align:center;font-family:'DM Mono',monospace">${s.casePack!=null?s.casePack:'<span style="color:var(--amber)">—</span>'}</td>
    <td style="text-align:center;font-family:'DM Mono',monospace">${s.boxPallet!=null?s.boxPallet:'<span style="color:var(--amber)">—</span>'}</td>
    <td><span class="status-badge sb-${(s.cls||'b').toLowerCase()}">${s.cls||'—'}</span></td>
    <td><button class="btn btn-ghost btn-xs" onclick="editSKU(${skuMaster.indexOf(s)})">Edit</button></td>
  </tr>`).join('');
}
let editSKUIdx=null;
function openAddSKUModal(){editSKUIdx=null;['ns_code','ns_name','ns_cp','ns_bp'].forEach(id=>document.getElementById(id).value='');document.getElementById('ns_class').value='';document.getElementById('nsMsg').innerText='';document.getElementById('addSKUModal').classList.add('open');}
function editSKU(idx){
  editSKUIdx=idx;const s=skuMaster[idx];
  document.getElementById('ns_code').value=s.code;document.getElementById('ns_name').value=s.name;
  document.getElementById('ns_cp').value=s.casePack!=null?s.casePack:'';document.getElementById('ns_bp').value=s.boxPallet!=null?s.boxPallet:'';
  document.getElementById('ns_class').value=s.cls||'';document.getElementById('nsMsg').innerText='';
  document.getElementById('addSKUModal').classList.add('open');
}
function saveNewSKU(){
  const code=document.getElementById('ns_code').value.trim(),name=document.getElementById('ns_name').value.trim();
  if(!code||!name){document.getElementById('nsMsg').innerText='Code and Name required';return;}
  const entry={code,name,brand:getBrandFromSKUCode_(code)||'',casePack:document.getElementById('ns_cp').value?Number(document.getElementById('ns_cp').value):null,boxPallet:document.getElementById('ns_bp').value?Number(document.getElementById('ns_bp').value):null,cls:document.getElementById('ns_class').value||'B'};
  if(editSKUIdx!==null){skuMaster[editSKUIdx]=entry;}
  else{if(skuMaster.find(s=>s.code===code)){document.getElementById('nsMsg').innerText='Code already exists';return;}skuMaster.push(entry);}
  saveAll();closeModal('addSKUModal');renderSKUTable();document.getElementById('db-sku').innerText=skuMaster.length;
  showProcessing('Saving SKU to Google Sheet…');
  const skuPayload={'SKU Code':entry.code,'SKU Name':entry.name,'Brand':entry.brand||'','Case Pack (Units/Box)':entry.casePack||'','Box/Pallet (Boxes/Shelf)':entry.boxPallet||'','Classification (A/B/C)':entry.cls||'B'};
  callSheet('addOrUpdateSKURow',skuPayload,
    function(r){ hideProcessing(); if(r&&r.success) toast('SKU '+entry.code+' saved to sheet','success',2000); else toast('Sheet sync: '+(r&&r.message||'error'),'warn',3000); },
    function(e){ hideProcessing(); toast('Sheet sync error: '+e,'warn',3000); }
  );
}
function exportSKUXLSX(){
  const ws=XLSX.utils.json_to_sheet(skuMaster.map(s=>({'SKU Code':s.code,'SKU Name':s.name,'Case Pack (Units/Box)':s.casePack||'','Box/Pallet (Boxes/Shelf)':s.boxPallet||'','Classification (A/B/C)':s.cls||''})));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'SKU_Master');XLSX.writeFile(wb,'SKU_Master.xlsx');
}

function downloadSKUTemplate(){
  const ws=XLSX.utils.aoa_to_sheet([
    ['SKU Code','SKU Name','Case Pack (Units/Box)','Box/Pallet (Boxes/Shelf)','Classification (A/B/C)'],
    ['MWLJNTP.0003.B0_N','LJ NutriMix 2+ 350gm Chocolate Jar',24,32,'A'],
    ['MWLJNTP.00059.B0_N','LJ NutriMix 7+ 350gm Chocolate Jar',24,32,'A'],
    ['MWMMHRP.0004.AAAA.B0_N','MM Hair Gummies 30mcg - 30s',42,24,'A'],
    ['MWMMHRP.6286.AAAA.B0_N','MM Activator Regular Black 1mm',100,25,'A'],
    ['MWBWHFP.00531.B0_N','BW Hair Health Biotin Gummies - 30s',30,24,'B'],
  ]);
  ws['!cols']=[{wch:28},{wch:38},{wch:20},{wch:22},{wch:20}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'SKU_Master');XLSX.writeFile(wb,'SKU_Master_Template.xlsx');
}

const SKU_UPLOAD_ALIASES={
  code:     ['sku code','sku_code','material code','material_code','sku','item code','product code','code'],
  name:     ['sku name','sku_name','material name','material_name','description','name','product name'],
  brand:    ['brand','brand zone','brand name'],
  casePack: ['case pack','case pack (units/box)','case_pack','casepack','units per box','units/box','cp'],
  boxPallet:['box/pallet','box/pallet (boxes/shelf)','box_pallet','boxpallet','boxes per shelf','boxes/shelf','pallet size','bp'],
  cls:      ['classification','classification (a/b/c)','class','fsn','fsn class','category']
};

function parseSKUUpload(e){
  const file=e.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){showSKUUploadResult('No data rows found in file.','danger');return;}
      const fa=als=>Object.keys(rows[0]).find(h=>als.includes(String(h).toLowerCase().trim()))||null;
      const cm={code:fa(SKU_UPLOAD_ALIASES.code),name:fa(SKU_UPLOAD_ALIASES.name),casePack:fa(SKU_UPLOAD_ALIASES.casePack),boxPallet:fa(SKU_UPLOAD_ALIASES.boxPallet),cls:fa(SKU_UPLOAD_ALIASES.cls)};
      if(!cm.code){showSKUUploadResult('⚠ "SKU Code" column is required. Check your file headers.','danger');return;}
      if(!cm.name){showSKUUploadResult('⚠ "SKU Name" column is required. Check your file headers.','danger');return;}
      let added=0,updated=0,skipped=0;
      rows.forEach(r=>{
        const code=String(r[cm.code]||'').trim();if(!code){skipped++;return;}
        const name=String(r[cm.name]||'').trim();if(!name){skipped++;return;}
        const casePack=cm.casePack&&r[cm.casePack]!==''?Number(r[cm.casePack])||null:null;
        const boxPallet=cm.boxPallet&&r[cm.boxPallet]!==''?Number(r[cm.boxPallet])||null:null;
        const clsRaw=cm.cls?String(r[cm.cls]||'').trim().toUpperCase():'B';
        const cls=['A','B','C'].includes(clsRaw)?clsRaw:'B';
        const brand = cm.brand && r[cm.brand] ? String(r[cm.brand]).trim() : getBrandFromSKUCode_(code);
        const existing=skuMaster.find(s=>s.code===code);
        if(existing){
          existing.name=name; existing.brand=brand||existing.brand;
          existing.casePack=casePack!=null?casePack:existing.casePack;
          existing.boxPallet=boxPallet!=null?boxPallet:existing.boxPallet;
          existing.cls=cls; updated++;
        } else {
          skuMaster.push({code,name,brand,casePack,boxPallet,cls}); added++;
        }
      });
      saveAll();renderSKUTable();
      document.getElementById('db-sku').innerText=skuMaster.length;
      // Refresh SKU datalist for GRN entry
      const dl=document.getElementById('skuList');
      if(dl)dl.innerHTML=skuMaster.map(s=>`<option value="${s.code}">${s.name}</option>`).join('');
      const msg=`✅ SKU Master updated from <b>${file.name}</b> — ${added} added, ${updated} updated, ${skipped} skipped (${skuMaster.length} total SKUs).`;
      showSKUUploadResult(msg,'success');
      if(typeof google!=='undefined'&&google.script){
        syncSKUMasterToSheet();
      }
    }catch(err){showSKUUploadResult('Error reading file: '+err.message,'danger');}
  };
  reader.readAsArrayBuffer(file);e.target.value='';
}

function showSKUUploadResult(msg,type){
  const el=document.getElementById('skuUploadResult');
  el.style.display='block';
  el.className='notice '+(type||'info');
  el.innerHTML=`<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>${msg}`;
}

function syncSKUMasterToSheet(){
  const rows=skuMaster.map(s=>({'SKU Code':s.code,'SKU Name':s.name,'Brand':s.brand||getBrandFromSKUCode_(s.code)||'','Case Pack (Units/Box)':s.casePack||'','Box/Pallet (Boxes/Shelf)':s.boxPallet||'','Classification (A/B/C)':s.cls||''}));
  callSheet('bulkSaveSKUMaster', rows,
    r=>showSKUUploadResult(r&&r.success?'✅ SKU Master synced to Google Sheet ('+skuMaster.length+' SKUs)':'⚠ Sync issue: '+(r&&r.message||'unknown'), r&&r.success?'success':'warn'),
    err=>showSKUUploadResult('⚠ Sheet sync error: '+err,'warn')
  );
}

/* ════════════════════ UNIWIRE DIFF ════════════════════ */
const DIFF_ALIASES={
  skuCode:     ['item type sku code','sku code','sku_code','material code','material','sku'],
  skuName:     ['item type name','sku name','sku_name','material name','description','name'],
  batch:       ['vendor batch code','vendor batch','batch code','batch','batch no','batch_no','lot'],
  qty:         ['quantity','qty','unrestricted','unrestricted qty','good qty','available qty'],
  stockType:   ['inventory type','stock type','type'],
  bin:         ['shelf','bin','bin id','bin / shelf','location'],
  mfgDate:     ['manufacturing','mfg date','manufacturing date'],
  expDate:     ['expiry','exp date','expiry date','best before'],
  unrestricted:['unrestricted','unrestricted qty','unrestricted quantity'],
  blocked:     ['blocked','blocked qty','bad inventory','bad stock'],
};
function runDiffFromCCLog(){
  if(!ccLog.length){toast('No Cycle Count log entries. Run a cycle count first.','warn',4000);return;}
  if(!inventory.length){toast('No inventory data — sync from Uniware first.','warn',4000);return;}
  const msg=document.getElementById('diffMsg');
  if(!_gatewayPassesLoaded){
    msg.innerHTML='<span style="color:var(--amber)">Loading gatepass pending data...</span>';
    loadGatepasses(function(){ runDiffFromCCLog(); });
    return;
  }
  // Build latest-count-per-bin+sku+batch+type map
  const ccMap={};
  ccLog.forEach(r=>{
    const key=(r.bin||r.Bin||r['Bin / Shelf']||'')+'|'+(r.skuCode||r['SKU Code']||'')+'|'+(r.batch||r.Batch||r['Batch No']||'')+'|'+normalizeStockType_(r.stockType||r['Stock Type']||'GOOD');
    const existing=ccMap[key];
    const ts=new Date(r.time||r.Date||0).getTime();
    if(!existing||ts>new Date(existing.time||existing.Date||0).getTime()) ccMap[key]=r;
  });
  const ccRows=Object.values(ccMap).map(r=>({
    bin:      r.bin||r.Bin||r['Bin / Shelf']||'',
    skuCode:  r.skuCode||r['SKU Code']||'',
    skuName:  r.skuName||r['SKU Name']||'',
    batch:    r.batch||r.Batch||r['Batch No']||'',
    stockType:normalizeStockType_(r.stockType||r['Stock Type']||'GOOD'),
    qty:      Number(r.countedQty||r['Counted Qty']||r.counted||0),
    mfgDate:  normDate(r.mfgDate||r['MFG Date']||''),
    expDate:  normDate(r.expDate||r['EXP Date']||''),
    ccDate:   r.time||r.Date||'',
  })).filter(r=>r.skuCode&&r.batch);
  if(!ccRows.length){toast('CC log found but no valid rows with SKU+Batch','warn');return;}
  msg.innerHTML='<span style="color:var(--amber)">Comparing '+ccRows.length+' CC entries vs '+inventory.length+' inventory rows…</span>';
  _buildDiffRows(ccRows,'CC Log');
}

function _buildDiffRows(inputRows,source){
  // WMS inventory map: bin|sku|batch|type → qty
  const wmsMap={};
  inventory.forEach(r=>{
    const k=(r.bin||'')+'|'+r.skuCode+'|'+(r.batch||'')+'|'+normalizeStockType_(r.stockType||'');
    wmsMap[k]=(wmsMap[k]||0)+r.qty;
  });
  // Input map
  const inMap={};
  inputRows.forEach(r=>{
    const k=(r.bin||'')+'|'+r.skuCode+'|'+(r.batch||'')+'|'+normalizeStockType_(r.stockType||'GOOD');
    if(!inMap[k])inMap[k]={...r,qty:0};
    inMap[k].qty+=r.qty;
  });
  const allKeys=new Set([...Object.keys(inMap),...Object.keys(wmsMap)]);
  diffRows=[];
  allKeys.forEach(key=>{
    const [bin,skuCode,batch,stockType]=key.split('|');
    const inQ  = inMap[key] ? inMap[key].qty : 0;
    const wQ   = wmsMap[key] || 0;
    const rawDiff = wQ - inQ;

    // ── Gatepass adjustment ──
    // If there is a pending gatepass (CREATED/RETURN_AWAITED) for this bin+sku+batch,
    // that qty is already picked/pending dispatch — accounts for apparent excess
    const gpPending = getGpPendingQty(bin, skuCode, batch, stockType);
    // Net variance = raw diff minus pending gatepass
    // e.g. WMS=1300, Physical=1296, rawDiff=+4, GP pending=4 → netDiff=0 → MATCH
    const netDiff  = rawDiff - gpPending;

    const inR = inMap[key] || {};
    const ivR = inventory.find(r=>r.bin===bin&&r.skuCode===skuCode&&r.batch===batch)||{};

    diffRows.push({
      bin, skuCode, batch, stockType,
      skuName:   inR.skuName  || ivR.skuName  || '',
      mfgDate:   inR.mfgDate  || ivR.mfgDate  || '',
      expDate:   inR.expDate  || ivR.expDate  || '',
      sourceQty: inQ,
      wmsQty:    wQ,
      gpPending,                              // gatepass pending qty for this line
      rawDiff,                                // WMS - Physical (before GP adjustment)
      diff:      netDiff,                     // net variance after GP adjustment
      status:    netDiff===0 ? 'MATCH' : netDiff>0 ? 'EXCESS' : 'SHORT',
      rawStatus: rawDiff===0 ? 'MATCH' : rawDiff>0 ? 'EXCESS' : 'SHORT',
      ccDate:    inR.ccDate || '',
      source
    });
  });
  diffRows.sort((a,b)=>{
    const o={SHORT:0,EXCESS:1,MATCH:2};
    return o[a.status]!==o[b.status]?o[a.status]-o[b.status]:(a.bin||'').localeCompare(b.bin||'');
  });
  renderDiffTable();
  const sh=diffRows.filter(r=>r.status==='SHORT').length;
  const ex=diffRows.filter(r=>r.status==='EXCESS').length;
  const msg=document.getElementById('diffMsg');
  msg.innerHTML='<span style="color:var(--green)">✅ Done — Source: '+source+' | Total: '+diffRows.length+' | Short: '+sh+' | Excess: '+ex+'</span>';
  if(sh+ex===0) toast('All entries match!','success',4000);
  else toast((sh+ex)+' discrepancies found','warn',3000);
}

function handleDiffDrop(e){e.preventDefault();e.currentTarget.classList.remove('drag');parseDiffFile({target:{files:[e.dataTransfer.files[0]]}});}
function parseDiffFile(e){
  const file=e.target.files[0]; if(!file)return;
  const msg=document.getElementById('diffMsg');
  msg.innerHTML='<span style="color:var(--amber)">Reading…</span>';
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb   = XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){msg.innerText='⚠ File is empty';return;}
      const fa = al=>{const lo=al.map(a=>a.toLowerCase());return Object.keys(rows[0]).find(h=>lo.includes(h.toLowerCase().trim()))||null;};
      const cm = {};
      Object.keys(DIFF_ALIASES).forEach(k=>cm[k]=fa(DIFF_ALIASES[k]));
      if(!cm.skuCode){msg.innerText='⚠ SKU Code column not found. Expected: "Item Type SKU Code"';return;}
      if(!cm.batch)  {msg.innerText='⚠ Batch column not found. Expected: "Vendor batch code"';return;}
      const uploaded = rows.map(r=>({
        bin:       cm.bin       ? String(r[cm.bin]||'').trim()       : '',
        skuCode:   String(r[cm.skuCode]||'').trim(),
        skuName:   cm.skuName  ? String(r[cm.skuName]||'').trim()   : '',
        batch:     String(r[cm.batch]||'').trim(),
        stockType: cm.stockType? normalizeStockType_(r[cm.stockType]||'') : 'GOOD',
        qty:       Number(String((cm.qty?r[cm.qty]:0)||'0').replace(/,/g,''))||0,
        mfgDate:   cm.mfgDate  ? normDate(r[cm.mfgDate]||'')        : '',
        expDate:   cm.expDate  ? normDate(r[cm.expDate]||'')         : '',
      })).filter(r=>r.skuCode&&r.batch);
      msg.innerHTML='<span style="color:var(--green)">✅ '+uploaded.length+' rows read — comparing with WMS…</span>';
      _buildDiffRows(uploaded,'Shelfwise Upload');
    }catch(err){ msg.innerHTML='<span style="color:var(--red)">❌ '+err.message+'</span>'; }
  };
  reader.readAsArrayBuffer(file);
}


function renderDiffTable(){
  const filterVal=(document.getElementById('diffFilterSelect')||{}).value||'DIFF';
  const body=document.getElementById('diffBody');
  if(!diffRows.length){body.innerHTML='<tr class="empty-row"><td colspan="10">Run ⚡ From CC Log or upload Shelfwise file to compare</td></tr>';return;}
  let filtered;
  if(filterVal==='ALL')         filtered=diffRows;
  else if(filterVal==='SHORT')  filtered=diffRows.filter(r=>r.status==='SHORT');
  else if(filterVal==='EXCESS') filtered=diffRows.filter(r=>r.status==='EXCESS');
  else                          filtered=diffRows.filter(r=>r.status!=='MATCH');
  const tot=diffRows.length, mat=diffRows.filter(r=>r.status==='MATCH').length,
        sh=diffRows.filter(r=>r.status==='SHORT').length, ex=diffRows.filter(r=>r.status==='EXCESS').length;
  document.getElementById('diffSummary').innerHTML=
    '<span class="chip chip-total">'+tot+' Total</span>'+
    '<span class="chip chip-ok">'+mat+' Match</span>'+
    '<span class="chip" style="background:#fef9c3;color:#854d0e">'+sh+' Short</span>'+
    '<span class="chip chip-err">'+ex+' Excess</span>'+
    '<span class="chip" style="background:#f1f5f9;color:#475569">'+filtered.length+' Showing</span>';
  if(!filtered.length){body.innerHTML='<tr class="empty-row"><td colspan="10">No rows match filter</td></tr>';return;}
  body.innerHTML=filtered.map(r=>{
    const dc=r.status==='SHORT'?'color:var(--red);font-weight:700':r.status==='EXCESS'?'color:var(--green);font-weight:700':'color:var(--slate4)';
    const sb=r.status==='SHORT'?'<span class="status-badge sb-reserved">SHORT</span>':r.status==='EXCESS'?'<span class="status-badge sb-ok">EXCESS</span>':'<span class="status-badge sb-empty">MATCH</span>';
    const ds=(r.diff>0?'+':'')+r.diff;
    const rds = (r.rawDiff>0?'+':'')+r.rawDiff;
    const nds = (r.diff>0?'+':'')+r.diff;
    const ndc = r.diff===0?'color:var(--slate4)':r.diff>0?'color:var(--green);font-weight:700':'color:var(--red);font-weight:700';
    const gpCell = r.gpPending>0
      ? '<span style="color:var(--amber);font-weight:700">'+r.gpPending+'</span>'
      : '<span style="color:var(--slate4)">—</span>';
    return '<tr>'+
      '<td><span class="bin-badge">'+(r.bin||'—')+'</span></td>'+
      '<td><span class="sku-badge">'+r.skuCode+'</span></td>'+
      '<td style="font-size:11px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.skuName+'</td>'+
      '<td><b>'+r.batch+'</b></td>'+
      '<td><span class="status-badge '+(r.stockType==='GOOD'?'sb-ok':'sb-occupied')+'">'+r.stockType+'</span></td>'+
      '<td style="text-align:right;font-family:DM Mono,monospace">'+r.sourceQty+'</td>'+
      '<td style="text-align:right;font-family:DM Mono,monospace;font-weight:700">'+r.wmsQty+'</td>'+
      '<td style="text-align:right;font-family:DM Mono,monospace;'+dc+'">'+rds+'</td>'+
      '<td style="text-align:right">'+gpCell+'</td>'+
      '<td style="text-align:right;font-family:DM Mono,monospace;'+ndc+'">'+nds+'</td>'+
      '<td>'+sb+'</td>'+
      '<td style="font-size:10px;color:var(--slate4)">'+(r.ccDate?(r.ccDate+'').substring(0,10):'—')+'</td>'+
    '</tr>';
  }).join('');
}


function downloadDiffTemplate(){
  const rows=[
    ['Shelf','Item Type SKU Code','Item Type Name','Vendor batch code','Manufacturing','Expiry','Inventory Type','Quantity','MRP'],
    ['R10-C12-002','MWBWSKP.00684.B0_N','Sample Product','NJ5179','01-03-2026','30-04-2027','GOOD_INVENTORY',100,399],
    ['R10-C12-002','MWBWSKP.00684.B0_N','Sample Product','NJ5179','01-03-2026','30-04-2027','BAD_INVENTORY',5,399],
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:14},{wch:22},{wch:30},{wch:16},{wch:12},{wch:12},{wch:18},{wch:10},{wch:8}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Template');
  XLSX.writeFile(wb,'Shelfwise_Diff_Template.xlsx');
  toast('Template downloaded — same format as Shelfwise CSV from email','success',3000);
}


function exportDiffXLSX(){
  if(!diffRows.length){alert('Run diff check first');return;}
  const fv=(document.getElementById('diffFilterSelect')||{}).value||'DIFF';
  let rows;
  if(fv==='ALL')         rows=diffRows;
  else if(fv==='SHORT')  rows=diffRows.filter(r=>r.status==='SHORT');
  else if(fv==='EXCESS') rows=diffRows.filter(r=>r.status==='EXCESS');
  else                   rows=diffRows.filter(r=>r.status!=='MATCH');
  if(!rows.length){toast('No rows to export','warn');return;}
  const data=rows.map(r=>({'Bin/Shelf':r.bin||'','SKU Code':r.skuCode,'SKU Name':r.skuName,'Batch':r.batch,
    'Stock Type':r.stockType,'MFG Date':r.mfgDate||'','EXP Date':r.expDate||'',
    'Physical Qty':r.sourceQty,'WMS Qty':r.wmsQty,'Raw Diff':r.rawDiff,'GP Pending':r.gpPending||0,'Net Variance':r.diff,'Status':r.status,
    'Count Date':r.ccDate||'','Source':r.source||''}));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Diff Mismatches');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
    ['WMS Diff Report — '+new Date().toLocaleDateString('en-IN')],
    ['Total',diffRows.length],['Match',diffRows.filter(r=>r.status==='MATCH').length],
    ['Short',diffRows.filter(r=>r.status==='SHORT').length],
    ['Excess',diffRows.filter(r=>r.status==='EXCESS').length],['Exported',rows.length]
  ]),'Summary');
  const fname='WMS_Diff_'+new Date().toISOString().slice(0,10)+'.xlsx';
  XLSX.writeFile(wb,fname);
  toast('Exported '+rows.length+' rows to '+fname,'success',3000);
}


function printPalletLabels(){
  const plan=putawayPlan.filter(r=>!r.warning);
  if(!plan.length){alert('Generate Putaway Plan first');return;}
  const ps=`@page{size:4in 2in;margin:0}body{font-family:Arial,sans-serif}
    .label{width:4in;height:2in;box-sizing:border-box;padding:6px 8px 4px;page-break-after:always;border:1px solid #000;display:flex;flex-direction:column;overflow:hidden}
    .label:last-child{page-break-after:avoid}
    .top{display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #000;padding-bottom:3px;margin-bottom:3px}
    .co{font-size:9px;font-weight:900;letter-spacing:.5px}.gn{font-size:9px;font-weight:700;border:1.5px solid #000;padding:1px 6px;border-radius:3px}
    .bin{text-align:center;font-size:48px;font-weight:900;line-height:1;letter-spacing:-2px;flex:1;display:flex;align-items:center;justify-content:center}
    .sku-line{border-top:1.5px solid #000;padding-top:3px;margin-top:3px;font-size:7.5px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .grid{display:grid;grid-template-columns:repeat(6,1fr);border-top:1px solid #000;margin-top:2px}
    .gc{padding:2px 3px;border-right:1px solid #ddd;font-size:6px;color:#555}
    .gc:last-child{border-right:none}.gv{font-size:9px;font-weight:900;line-height:1.2}`;
  let html='';
  plan.forEach(r=>{
    const cp=Number(r.casePack||0);const bp=Number(r.boxPallet||0);
    html+=`<div class="label">
      <div class="top"><span class="co">MOSAIC WELLNESS — PUTAWAY LABEL</span></div>
      <div class="bin">${r.binId}</div>
      <div class="sku-line">${r.sku} | ${r.skuName}</div>
      <div class="grid">
        <div class="gc"><div>Batch</div><div class="gv">${r.batch}</div></div>
        <div class="gc"><div>Units</div><div class="gv">${r.allocQty}</div></div>
        <div class="gc"><div>Boxes</div><div class="gv">${r.allocBoxes||'—'}</div></div>
        <div class="gc"><div>Case Pack</div><div class="gv">${cp||'—'}</div></div>
        <div class="gc"><div>Pallet #</div><div class="gv">${r.palletNo||'1'}</div></div>
        <div class="gc"><div>Date</div><div class="gv">${new Date().toLocaleDateString('en-IN')}</div></div>
      </div>
    </div>`;
  });
  doPrint(html,ps);
}

function printPutawayPlan(){
  if(!putawayPlan.length){alert('Generate Putaway Plan first');return;}
  const grnNo=currentGRN?.grnNo||'—';
  const now=new Date().toLocaleString('en-IN');
  const ps=`@page{size:A4 portrait;margin:8mm}body{font-family:Arial,sans-serif;font-size:9pt}
    .title{text-align:center;font-size:15pt;font-weight:900;margin-bottom:8px}.meta{display:flex;gap:20px;margin-bottom:10px;font-size:8pt;font-weight:700}
    table{width:100%;border-collapse:collapse}th{background:#1e293b;color:#fff;font-size:7.5pt;padding:5px;text-align:left;border:1px solid #000}
    td{font-size:8.5pt;padding:4px 5px;border:1px solid #ccc}tr:nth-child(even) td{background:#f8fafc}
    .total-row td{font-weight:900;background:#e2e8f0}.warn-row td{background:#fff5f5}
    .footer{margin-top:16px;border-top:1px solid #000;padding-top:8px;display:flex;justify-content:space-between;font-size:8pt}
    .sign{border-top:1px solid #000;width:160px;text-align:center;padding-top:4px;font-size:8pt}`;
  let total=0;
  const rows=putawayPlan.map((r,i)=>{total+=r.allocQty;return `<tr class="${r.warning?'warn-row':''}"><td>${i+1}</td><td>${r.warning?'NO BIN':r.binId}</td><td>${r.sku}</td><td>${r.skuName}</td><td>${r.batch}</td><td>${r.cls||'—'}</td><td style="text-align:center">${r.palletNo||'1'}</td><td style="text-align:center">${r.allocBoxes||'—'}</td><td style="text-align:right">${r.allocQty}</td><td>${r.isFullPallet?'Full':'Partial'}</td></tr>`;}).join('');
  const html=`<div class="title">PUTAWAY PLAN — MOSAIC WELLNESS</div>
    <div class="meta"><span>GRN: ${grnNo}</span><span>Date: ${now}</span><span>By: ${userName}</span></div>
    <table><thead><tr><th>#</th><th>Bin</th><th>SKU Code</th><th>SKU Name</th><th>Batch</th><th>Class</th><th>Pallet#</th><th>Boxes</th><th>Units</th><th>Type</th></tr></thead>
    <tbody>${rows}<tr class="total-row"><td colspan="8" style="text-align:right">GRAND TOTAL</td><td>${total}</td><td></td></tr></tbody></table>
    <div class="footer"><div class="sign">Warehouse Supervisor</div><div class="sign">Receiver</div><div style="font-size:7pt;color:#888">Printed: ${now}</div></div>`;
  doPrint(html,ps);
}

function doPrint(body,style){
  const win=window.open('','_blank','width=900,height=700');
  if(!win){alert('Popup blocked — allow popups for this page');return;}
  // IMPORTANT: Only inject print-specific styles, NOT the main app CSS
  // This prevents sync modal, offline banner etc from appearing in print
  win.document.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{background:#fff;color:#000;font-family:Arial,sans-serif}' +
    style +
    '</style></head><body>' + body + '</body></html>');
  win.document.close();
  win.focus();
  setTimeout(function(){ win.print(); win.close(); }, 600);
}


/* ════════════════════ HELPERS ════════════════════ */
function showResult(id,msg,type){
  const el=document.getElementById(id);if(!el)return;
  el.style.display='flex';el.className='notice '+(type||'info');
  el.innerHTML=`<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>${msg}`;
}
function closeModal(id){document.getElementById(id).classList.remove('open');}


/* ════════════════════ GATEPASS UPLOAD ════════════════════ */
const GP_ALIASES={
  gpNo:['gatepass no','gatepass_no','gp no','gp_no','gatepass','gp number','gate pass no'],
  date:['date','gatepass date','gp date'],
  skuCode:['sku code','sku_code','material code','sku','item code','product code'],
  skuName:['sku name','sku_name','material name','description','product name'],
  bin:['bin','shelf','bin no','bin id','location','shelf no'],
  qty:['qty','quantity','units'],
  reason:['reason','remarks','note','notes','comment']
};
function handleGpDrop(e){e.preventDefault();e.currentTarget.classList.remove('drag');parseGatepassFile({target:{files:[e.dataTransfer.files[0]]}});}
function parseGatepassFile(e){
  const file=e.target.files[0];if(!file)return;
  const msg=document.getElementById('gpMsg');msg.innerText='Reading…';
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      if(!rows.length){msg.innerText='No rows found';return;}
      const heads=Object.keys(rows[0]);
      const fa=als=>heads.find(h=>als.includes(String(h).toLowerCase().trim()))||null;
      const cm={};Object.keys(GP_ALIASES).forEach(k=>cm[k]=fa(GP_ALIASES[k]));
      if(!cm.skuCode||!cm.qty){msg.innerText='⚠ SKU Code and Qty columns required';return;}
      let added=0;
      rows.forEach(r=>{
        const skuCode=String(r[cm.skuCode]||'').trim();
        const qty=Number(String(r[cm.qty]||'0').replace(/,/g,''));
        if(!skuCode||qty<=0)return;
        gatepasses.push({
          gpNo:cm.gpNo?String(r[cm.gpNo]||'').trim():'',
          date:cm.date?String(r[cm.date]||'').trim():'',
          skuCode,
          skuName:cm.skuName?String(r[cm.skuName]||'').trim():'',
          bin:cm.bin?String(r[cm.bin]||'').trim():'',
          qty,
          reason:cm.reason?String(r[cm.reason]||'').trim():'',
          uploadedAt:new Date().toISOString(),uploadedBy:userName
        });
        added++;
      });
      saveAll();
      msg.innerText=`✓ ${added} gatepass line(s) added from ${file.name} (total: ${gatepasses.length})`;
      renderGatepassSummary();renderGatepassTable();loadDashboard();
      // Sync to Google Sheet
      callSheet('saveGatepasses', gatepasses,
        r=>{ msg.innerText=`✓ ${added} gatepass line(s) synced to Google Sheet (total: ${gatepasses.length}).`; },
        e=>{ msg.innerText+=` ⚠ Sheet sync failed: ${e}`; }
      );
    }catch(err){msg.innerText='Error: '+err.message;}
  };
  reader.readAsArrayBuffer(file);e.target.value='';
}
function renderGatepassSummary(){
  const total=gatepasses.reduce((s,r)=>s+r.qty,0);
  const skus=new Set(gatepasses.map(r=>r.skuCode)).size;
  document.getElementById('gpSummary').innerHTML=`<span class="chip chip-total">${gatepasses.length} Lines</span><span class="chip chip-ok">${skus} SKUs</span><span class="chip chip-err">Blocked: ${total.toLocaleString()} units</span>`;
}
function renderGatepassTable(){
  const q=String(document.getElementById('gpSearch')?.value||'').toLowerCase();
  const filtered=gatepasses.filter(r=>`${r.gpNo} ${r.skuCode} ${r.skuName} ${r.bin}`.toLowerCase().includes(q));
  const body=document.getElementById('gpBody');
  if(!filtered.length){body.innerHTML='<tr class="empty-row"><td colspan="7">No gatepass records found</td></tr>';return;}
  body.innerHTML=filtered.map(r=>`<tr>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${r.gpNo||'—'}</td>
    <td style="font-size:11px">${r.date||'—'}</td>
    <td><span class="sku-badge">${r.skuCode}</span></td>
    <td style="font-size:12px">${r.skuName||'—'}</td>
    <td><span class="bin-badge" style="font-size:10px">${r.bin||'—'}</span></td>
    <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--red)">${r.qty.toLocaleString()}</td>
    <td style="font-size:11px;color:var(--slate4)">${r.reason||'—'}</td>
  </tr>`).join('');
  if(!gatepasses.length){document.getElementById('gpSummary').innerHTML='';}else renderGatepassSummary();
}
function clearGatepasses(){
  if(!confirm('Clear ALL gatepass records? This cannot be undone.'))return;
  gatepasses=[];saveAll();renderGatepassTable();loadDashboard();
  document.getElementById('gpSummary').innerHTML='';document.getElementById('gpMsg').innerText='Cleared.';
  callSheet('clearGatepasses', {}, ()=>{}, ()=>{});
}
function downloadGatepassTemplate(){
  const ws=XLSX.utils.aoa_to_sheet([['Gatepass No','Date','SKU Code','SKU Name','Bin/Shelf','Qty','Reason'],['GP-2024-001','2024-01-15','MWLJNTP.0003.B0_N','LJ NutriMix 2+ 350gm','R1-C1-001',50,'Quality Hold']]);
  ws['!cols']=[{wch:16},{wch:12},{wch:28},{wch:30},{wch:14},{wch:8},{wch:20}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Gatepass');XLSX.writeFile(wb,'Gatepass_Template.xlsx');
}
function exportGatepassXLSX(){
  if(!gatepasses.length){alert('No gatepass data to export');return;}
  const ws=XLSX.utils.json_to_sheet(gatepasses.map(r=>({'Gatepass No':r.gpNo,'Date':r.date,'SKU Code':r.skuCode,'SKU Name':r.skuName,'Bin/Shelf':r.bin,'Qty':r.qty,'Reason':r.reason})));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Gatepasses');XLSX.writeFile(wb,'Gatepass_Data.xlsx');
}
// v6: GP block is now BIN-aware — pass bin for shelf-level accuracy
// Uses server-provided value if available (more accurate), otherwise local gatepass data
function getGpBlock(skuCode, batch, bin){
  if(bin){
    // Prefer server-provided shelf-level GP (set during ccSearchByBin from server)
    if(_serverGpBlockByBin && _serverGpBlockByBin[bin] !== undefined) {
      return _serverGpBlockByBin[bin];
    }
    // Shelf-level: sum all GP for this shelf regardless of SKU (warehouse reality)
    const binTotal = gatepasses.filter(r=>r.bin===bin).reduce((s,r)=>s+r.qty,0);
    return binTotal;
  }
  // Fallback: SKU+batch level (original behaviour)
  return gatepasses.filter(r=>r.skuCode===skuCode&&(!batch||!r.batch||r.batch===batch)).reduce((s,r)=>s+r.qty,0);
}

/* ════════════════════ CYCLE COUNT ════════════════════ */

// Adjustment log: persisted across session
