/* ════════════════════ STATE ════════════════════ */
// v6: Users now loaded from Google Sheet USERS tab via login()
// Hardcoded users kept ONLY as emergency offline fallback
const OFFLINE_USERS=[
  {id:'admin', pw:'admin123', name:'Admin User',  role:'ADMIN', allowedSections:'ALL'},
  {id:'wms1',  pw:'wms123',  name:'WMS Operator', role:'OPERATOR', allowedSections:'INBOUND'},
];

function ls(k){try{return JSON.parse(localStorage.getItem('mwwms2_'+k)||'null');}catch{return null;}}
function lsSet(k,v){try{localStorage.setItem('mwwms2_'+k,JSON.stringify(v));}catch{}}

// ══════════════════════════════════════════════════════════════
//  GOOGLE APPS SCRIPT WEB APP URL
//  Paste your deployed web app URL here (ending with /exec).
//  Get it from: Apps Script → Deploy → Manage Deployments
//  It looks like: https://script.google.com/macros/s/AKfycb.../exec
// ══════════════════════════════════════════════════════════════
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbwOftOsHKyol0NPhQFy8waSCT5GFdaPkrqt45U8PEip8kLVn9tXJdd3HR1_ni5xbcM9/exec';
const GAS_URL = ls('gasUrl') || DEFAULT_GAS_URL;

function setGasUrl(url){
  lsSet('gasUrl', url.trim());
  location.reload();
}

// ── Universal Google Sheet caller ──
// Works whether HTML is opened as a file, hosted anywhere, or served by Apps Script.
// action: string name of the server function
// payload: object to send as JSON
// onSuccess: function(result)
// onError: optional function(errMsg)
async function callSheet(action, payload, onSuccess, onError) {
  // Method 1: Running inside Apps Script (served via doGet)
  if (typeof google !== 'undefined' && google.script) {
    const runner = google.script.run
      .withSuccessHandler(r => onSuccess && onSuccess(r))
      .withFailureHandler(e => onError ? onError(e.message) : console.warn('Sheet error:', e.message));
    if      (action === 'bulkSaveSKUMaster')  runner.bulkSaveSKUMaster(payload);
    else if (action === 'bulkSaveBinMaster')  runner.bulkSaveBinMaster(payload);
    else if (action === 'saveInventoryDump')  runner.saveInventoryDump(payload.rows, payload.mode);
    else if (action === 'saveGatepasses')     runner.saveGatepasses(payload);
    else if (action === 'clearGatepasses')    runner.clearGatepasses();
    else if (action === 'saveCycleCountLog')  runner.saveCycleCountLog(payload);
    else if (action === 'saveGRNToSheet')     runner.saveGRNToSheet(payload);
    else if (action === 'loadAllWMSData')     runner.loadAllWMSData();
    // ── NEW v5 ──
    else if (action === 'saveAdjustmentLog')  runner.saveAdjustmentLog(payload);
    else if (action === 'getInventoryByBin')  runner.getInventoryByBin(payload.bin);
    else if (action === 'getBatchDates')      runner.getBatchDates(payload.skuCode, payload.batch);
    else if (action === 'syncCurrentInventory') runner.syncCurrentInventory(payload);
    else if (action === 'getGPBlockByBin')    runner.getGPBlockByBin(payload.bin);
    else if (action === 'saveAdjLogIdempotent') runner.saveAdjustmentLogIdempotent(payload.rows, payload.key);
    else if (action === 'clearSheetData')       runner.clearSheetData(payload.sheetName);
    else if (action === 'refreshAllData')       runner.refreshAllData();
    else if (action === 'login')                runner.login(payload.userId, payload.password);
    else if (action === 'getLastSyncStatus')    runner.getLastSyncStatus();
    else if (action === 'manualUniwareSync')    runner.manualUniwareSync();
    else if (action === 'restorePreviousDump')  runner.restorePreviousDump();
    else if (action === 'forceReimportLatest')  runner.forceReimportLatest();

    else if (action === 'getOverflowOptions')   runner.getOverflowOptions();
    else if (action === 'ping')                 runner.ping ? runner.ping() : onSuccess({success:true});
    else if (action === 'syncGatepasses')       runner.syncGatepasses();
    else if (action === 'forceReimportGatepass') runner.forceReimportGatepass();
    else if (action === 'getGatepasses')        runner.getGatepasses();
    else {
      // Unknown action — call onSuccess with null so callers handle gracefully
      if (onSuccess) onSuccess({ success: false, message: 'Unknown action: ' + action });
    }
    return;
  }
  // Method 2: Fetch to deployed web app URL
  const url = GAS_URL || ls('gasUrl');
  if (!url) {
    console.warn('GAS_URL not set');
    if (onError) onError('GAS URL not configured — open ⚙ Sheet settings to connect');
    return;
  }
  try {
    const body = JSON.stringify({ action, payload });
    const res  = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body
    });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch(e) { data = { success: false, message: text }; }
    if (onSuccess) onSuccess(data);
  } catch(e) {
    console.warn('callSheet fetch error [' + action + ']:', e.message);
    if (onError) onError(e.message);
  }
}

let skuMaster   = ls('sku')     || initSKUs();
let binMaster   = ls('bin')     || initBins();
let grnList     = ls('grn')     || [];
let grnLines    = ls('grnLines')|| [];
let inventory   = ls('inv')     || [];
let gatepasses  = ls('gp')      || [];
let putawayPlan = ls('putaway') || [];
let currentGRN  = ls('curGRN')  || null;
let grnCounter  = ls('grnCtr')  || 1;
let editGRNIdx  = null;
let editGRNLines= [];
let grnMode     = 'manual';
let userName     = '';
let userRole     = 'OPERATOR'; // ADMIN or OPERATOR — controls UI visibility
let userIdGlobal = '';
let userSections = ['OVERVIEW','REPORTS'];
let _isOnline        = true;  // tracks GAS connectivity
let gatewayPasses    = [];   // gatepass rows loaded from GATEPASS_DUMP
let _gatewayPassesLoaded = false;
let _pingFailCount   = 0;    // consecutive ping failures before showing offline banner
const PING_FAIL_THRESHOLD = 3; // show offline after 3 consecutive failures (~3 min)
let diffRows    = [];
let dumpRows    = [];
let ccLog       = [];
let ccLines     = [];

const ALL_SECTION_GROUPS = ['OVERVIEW','INBOUND','INVENTORY','OPERATIONS','PLANNING','REPORTS'];
const SECTION_GROUP_BY_VIEW = {
  dashboard:'OVERVIEW',
  grn:'INBOUND',
  grnList:'INBOUND',
  putaway:'INBOUND',
  dump:'INVENTORY',
  gatepass:'INVENTORY',
  bins:'INVENTORY',
  skumaster:'INVENTORY',
  currentinventory:'OPERATIONS',
  cyclecount:'OPERATIONS',
  adjustment:'OPERATIONS',
  fsnReplen:'PLANNING',
  binConsolidate:'PLANNING',
  expiryWatch:'PLANNING',
  diffcheck:'REPORTS'
};

function parseAllowedSections_(raw, role){
  const normalizedRole = String(role || '').trim().toUpperCase();
  const allowed = new Set(['OVERVIEW','REPORTS']); // Dashboard + Reports for every user
  String(raw || '').split(/[,\|;]/).forEach(part => {
    const section = part.trim().toUpperCase();
    if(section) allowed.add(section);
  });
  if(normalizedRole === 'ADMIN') allowed.add('ALL');
  if(allowed.has('ALL')) ALL_SECTION_GROUPS.forEach(section => allowed.add(section));
  if(normalizedRole !== 'ADMIN') allowed.delete('INVENTORY'); // Inventory remains admin-only
  return Array.from(allowed);
}

function canAccessSection_(section){
  const s = String(section || '').trim().toUpperCase();
  return userSections.includes('ALL') || userSections.includes(s);
}

function applySectionAccess(){
  document.querySelectorAll('[data-section-group]').forEach(el => {
    const section = el.getAttribute('data-section-group');
    el.style.display = canAccessSection_(section) ? '' : 'none';
  });
  document.querySelectorAll('.section').forEach(el => {
    const name = String(el.id || '').replace(/^s-/,'');
    const group = SECTION_GROUP_BY_VIEW[name] || 'OVERVIEW';
    el.style.display = canAccessSection_(group) ? '' : 'none';
  });
  if(canAccessSection_('INVENTORY')) document.body.classList.add('can-inventory');
  else document.body.classList.remove('can-inventory');

  const active = document.querySelector('.section.active');
  if(active){
    const activeName = String(active.id || '').replace(/^s-/,'');
    const activeGroup = SECTION_GROUP_BY_VIEW[activeName] || 'OVERVIEW';
    if(!canAccessSection_(activeGroup)) showSection('dashboard', document.querySelector('.nav-item[onclick*="dashboard"]'));
  }
}

function saveAll(){
  lsSet('sku',skuMaster);    // SKU master
  lsSet('bin',binMaster);    // Bin master
  lsSet('grn',grnList);      // GRN list
  lsSet('grnLines',grnLines);// GRN lines
  lsSet('inv',inventory);    // Inventory dump rows
  lsSet('putaway',putawayPlan);
  lsSet('curGRN',currentGRN);
  lsSet('grnCtr',grnCounter);
  lsSet('gp',gatepasses);    // Gatepasses
  // Note: adjLog is saved separately via saveAdjLog()
  // Note: ccLog is session-only (intentionally not persisted across sessions)
}

/**
 * v8: clearLocalData — clears ALL mwwms2_* localStorage keys.
 * Called when user wants a completely fresh state.
 */
function clearLocalData(){
  const keys = ['sku','bin','grn','grnLines','inv','putaway','curGRN','grnCtr','gp','adjLog','ccLog','gasUrl'];
  keys.forEach(k => { try{ localStorage.removeItem('mwwms2_'+k); }catch(e){} });
  skuMaster=initSKUs(); binMaster=initBins(); grnList=[]; grnLines=[];
  inventory=[]; dumpRows=[]; gatepasses=[]; putawayPlan=[];
  currentGRN=null; grnCounter=1; ccLog=[]; adjLog=[];
  saveAll(); saveAdjLog();
  toast('All local data cleared. Page will reload…','warn',2000);
  setTimeout(()=>location.reload(), 2000);
}

/* ════════════════════ v9: STOCK TYPE NORMALIZER ════════════════════ */
/**
 * Maps any stockType string to a known WMS value.
 * Handles: BAD_INVENTORY, bad inventory, BLOCKED, damage, GOOD, etc.
 * Unknown values default to GOOD so data is never lost.
 */
function normalizeStockType_(val){
  if(!val || val === '') return 'GOOD';
  const v = String(val).trim().toUpperCase().replace(/[_\s-]+/g, '_');
  // DAMAGE variants
  if(['DAMAGE','DAMAGED','DAMAGE_STOCK','BAD_INVENTORY','BAD_STOCK',
      'BAD','DEFECTIVE','REJECTED','SCRAP','QUARANTINE','HOLD',
      'BAD_INVENTORY','EXPIRED','EXPIRY_HOLD'].includes(v)) return 'DAMAGE';
  // BLOCKED variants  
  if(['BLOCKED','BLOCK','BLOCKED_STOCK','QC_HOLD','CQA','QUALITY',
      'QI','QUALITY_HOLD','UNDER_INSPECTION','INSPECTION'].includes(v)) return 'BLOCKED';
  // GOOD variants
  if(['GOOD','GOOD_STOCK','UNRESTRICTED','AVAILABLE','SALEABLE',
      'FREE','OK','ACTIVE','SELLABLE'].includes(v)) return 'GOOD';
  // If still unknown, log it and default to GOOD
  console.warn('WMS: Unknown stockType "' + val + '" → defaulting to GOOD');
  return 'GOOD';
}

function normalizeBinIdForMaster(value){
  var bin = String(value || '').trim().toUpperCase();
  if(!bin) return '';
  bin = bin.replace(/\s+/g,'').replace(/[.。]+$/g,'');
  var m = bin.match(/^R(\d{1,3})-C(\d{1,3})-(\d{1,3})$/);
  if(!m) return '';
  var row = Number(m[1]), col = Number(m[2]), level = Number(m[3]);
  if(row < 1 || col < 1 || level < 1) return '';
  return 'R' + row + '-C' + col + '-' + String(level).padStart(3,'0');
}

function normalizeBinMasterRows(rows){
  var byBin={}, rank={EMPTY:1,OCCUPIED:2,RESERVED:3,BLOCKED:4};
  (rows||[]).forEach(function(b){
    var id=normalizeBinIdForMaster(b && b.id);
    if(!id) return;
    var m=id.match(/^R(\d+)-C(\d+)-(\d+)$/);
    var status=String((b && b.status) || 'Empty').trim() || 'Empty';
    var item={
      id:id,
      row:(b && Number(b.row)) || (m ? Number(m[1]) : 0),
      col:(b && Number(b.col)) || (m ? Number(m[2]) : 0),
      level:(b && Number(b.level)) || (m ? Number(m[3]) : 0),
      brand:(b && (b.brand || b.zone)) || '',
      zone:(b && (b.zone || b.brand)) || '',
      fsn:String((b && b.fsn) || 'B').trim().toUpperCase(),
      capacity:b && b.capacity,
      status:status,
      skuFromDump:(b && b.skuFromDump) || ''
    };
    var prev=byBin[id];
    if(!prev || (rank[status.toUpperCase()]||1) > (rank[String(prev.status).toUpperCase()]||1)){
      byBin[id]=item;
    }
  });
  return Object.values(byBin);
}

binMaster = normalizeBinMasterRows(binMaster);
lsSet('bin', binMaster);

/* ════════════════════ v7: CLIENT-SIDE DATE NORMALISER ════════════════════ */
/**
 * Convert Excel serial numbers and various date strings to DD-MM-YYYY.
 * Mirrors the server-side normaliseDate_() in Code.gs.
 */
function normDate(val){
  if(val===null||val===undefined||val==='') return '';
  const s = String(val).trim();
  if(!s) return '';
  // Pure 5-digit number = Excel serial
  if(/^\d{4,6}(\.\d+)?$/.test(s) && Number(s) > 40000 && Number(s) < 60000){
    const serial = Math.floor(Number(s));
    const d = new Date((serial - 25569) * 86400 * 1000);
    if(!isNaN(d.getTime())){
      const dd=String(d.getUTCDate()).padStart(2,'0');
      const mm=String(d.getUTCMonth()+1).padStart(2,'0');
      const yy=d.getUTCFullYear();
      return dd+'-'+mm+'-'+yy;
    }
  }
  // ISO or JS date string
  const parsed = new Date(s);
  if(!isNaN(parsed.getTime()) && s.length > 6){
    const dd=String(parsed.getDate()).padStart(2,'0');
    const mm=String(parsed.getMonth()+1).padStart(2,'0');
    const yy=parsed.getFullYear();
    return dd+'-'+mm+'-'+yy;
  }
  return s; // already formatted (e.g. 01-Jan-2024)
}

/**
 * Convert DD-MM-YYYY format to YYYY-MM-DD (for HTML date input)
 * Input: "26-05-2023" → Output: "2023-05-26"
 */
function dateToHTMLFormat(dateStr){
  if(!dateStr) return '';
  const s = String(dateStr).trim();
  const m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if(m) return m[3] + '-' + m[2] + '-' + m[1];
  // If already in YYYY-MM-DD format, return as is
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return '';
}

/**
 * Convert YYYY-MM-DD format to DD-MM-YYYY (from HTML date input)
 * Input: "2023-05-26" → Output: "26-05-2023"
 */
function dateFromHTMLFormat(dateStr){
  if(!dateStr) return '';
  const s = String(dateStr).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(m) return m[3] + '-' + m[2] + '-' + m[1];
  // If already in DD-MM-YYYY format, return as is
  if(/^\d{2}-\d{2}-\d{4}$/.test(s)) return s;
  return '';
}

/* ════════════════════ INITIAL DATA ════════════════════ */
function initSKUs(){
  return [
    {code:'MWLJNTP.0003.B0_N',name:'LJ NutriMix 2+ 350gm Chocolate Jar',casePack:24,boxPallet:32,cls:'A'},
    {code:'MWLJNTP.00059.B0_N',name:'LJ NutriMix 7+ 350gm Chocolate Jar',casePack:24,boxPallet:32,cls:'A'},
    {code:'MWMMHRP.0004.AAAA.B0_N',name:'MM Hair Gummies 30mcg - 30s',casePack:42,boxPallet:24,cls:'A'},
    {code:'MWMMHRP.6286.AAAA.B0_N',name:'MM Activator Regular Black 1mm',casePack:100,boxPallet:25,cls:'A'},
    {code:'MWBWHFP.00531.B0_N',name:'BW Hair Health Biotin Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWLJNTP.00122.B0_N',name:'LJ NutriMix 2+ 350gm Vanilla Jar',casePack:24,boxPallet:32,cls:'A'},
    {code:'MWLJNTP.0002.B0_N',name:'LJ Multivitamin 2+ Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWBWSKP.00436.B0_N',name:'BB 5% Pigmentation Cream',casePack:96,boxPallet:40,cls:'A'},
    {code:'MWMMHRP.1001.AAAA.B0_N',name:'MM Anti Dandruff Shampoo 100ml',casePack:64,boxPallet:36,cls:'A'},
    {code:'MWLJNTP.00027.B0_N',name:'LJ NutriMix 2+ 350gm Chocolate Pouch',casePack:15,boxPallet:30,cls:'A'},
    {code:'MWLJNTP.0007.B0_N',name:'LJ Brain Health 2+ Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWLJNTS.0001.B0_N',name:'LJ NutriMix 2+ 20g Chocolate Mini',casePack:192,boxPallet:30,cls:'A'},
    {code:'MWBWSKP.00681.B0_N',name:'Be Bodywise 4% AHA BHA Underarm Roll-On (Aqua)',casePack:48,boxPallet:40,cls:'A'},
    {code:'MWLJNTP.00380.B0_N',name:'LJ Calcium Gummies (Strawberry) - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWMMHRP.0001.AAAA.B0_N',name:'MM DHT Blocking Shampoo 200ml',casePack:48,boxPallet:30,cls:'A'},
    {code:'MWLJNTP.00248.B0_N',name:'LJ Brain Health 7+ Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWMMHRP.2050.AAAA.B0_N',name:'MM Activator Regular Black 0.5mm',casePack:100,boxPallet:25,cls:'A'},
    {code:'MWLJNTP.00500.B0_N',name:'LJ NutriMix 7+ (Chocolate) - 1kg',casePack:11,boxPallet:30,cls:'B'},
    {code:'MWMMHRP.0003.AAAA.B0_N',name:'MM Hair Oil 100ml',casePack:70,boxPallet:48,cls:'B'},
    {code:'MWLJNTP.00507.B0_N',name:'LJ NutriMix 2+ (Chocolate) - 1kg',casePack:11,boxPallet:30,cls:'A'},
    {code:'MWBWHFP.00145.B0_N',name:'BW Hair Health Biotin Gummies - 60s',casePack:40,boxPallet:15,cls:'A'},
    {code:'MWMMPRK.2097.B0_N',name:'MM Shilajit Gummies - 60s',casePack:40,boxPallet:15,cls:'A'},
    {code:'MWLJNTP.00051.B0_N',name:'LJ Multivitamin 7+ Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWMMPRP.2092.B0_N',name:'MM Shilajit Gummies - 30s',casePack:30,boxPallet:24,cls:'A'},
    {code:'MWBWSKP.00648.B0_N',name:'BB 10% Urea Lotion 200ml',casePack:39,boxPallet:40,cls:'A'},
    {code:'MWMMHRP.0010.AAAA.B0_N',name:'MM Nourish Hair 30mcg Gummies - 60s',casePack:40,boxPallet:15,cls:'B'},
    {code:'MWBWSKP.00240.B0_N',name:'BeBodywise 10% Niacinamide Body Lotion 200ml',casePack:40,boxPallet:40,cls:'B'},
    {code:'MWLJNTP.00034.B0_N',name:'LJ NutriMix 2+ 350gm Strawberry Jar',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00177.B0_N',name:'LJ NutriMix 7+ 350gm Strawberry Jar',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWBWSKP.00241.B0_N',name:'BeBodywise 5% Lactic Acid Body Lotion 200ml',casePack:48,boxPallet:40,cls:'A'},
    {code:'MWLJNTP.00518.B0_N',name:'LJ Mother\'s Calcium Gummies - 30s',casePack:40,boxPallet:24,cls:'C'},
    {code:'MWBWHFP.00172.B0_N',name:'BeBodywise 1% Ketoconazole Shampoo 150ml',casePack:36,boxPallet:36,cls:'B'},
    {code:'MWLJNTP.00023.B0_N',name:'LJ Multivitamin 2+ Gummies - 60s',casePack:40,boxPallet:24,cls:'B'},
    {code:'MWLJNTP.00033.B0_N',name:'LJ NutriMix 2+ 700gm Chocolate Pouch',casePack:15,boxPallet:30,cls:'A'},
    {code:'MWMMHRP.0017.AAAA.B0_N',name:'MM Anti Dandruff Shampoo 200ml',casePack:48,boxPallet:36,cls:'B'},
    {code:'MWLJNTP.00056.B0_N',name:'LJ Multivitamin 2+ Gummies - 90s',casePack:24,boxPallet:12,cls:'B'},
    {code:'MWLJNTP.00339.B0_N',name:'LJ NutriMix for Mothers 400g Pouch',casePack:22,boxPallet:32,cls:'B'},
    {code:'MWMMNTL.0171.B0_N',name:'MM Shilajit BP Gummies - 60s',casePack:40,boxPallet:15,cls:'A'},
    {code:'MWLJNTS.0002.B0_N',name:'LJ Multivitamin 2+ Gummies - 4s Mini',casePack:250,boxPallet:24,cls:'A'},
    {code:'MWLJNTP.00241.B0_N',name:'LJ Calcium Gummies (Mango) - 30s',casePack:30,boxPallet:24,cls:'B'},
    {code:'MWLJNTP.00060.B0_N',name:'LJ Eye Gummies (Strawberry) - 30s',casePack:30,boxPallet:24,cls:'B'},
    {code:'MWLJNTP.00482.B0_N',name:'LJ NutriMix 2+ Hazelnut Chocolate 350g',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00529.B0_N',name:'LJ NutriMix 7+ Chocolate Hazelnut 350g',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00636.B0_N',name:'LJ NutriMix 7+ Vanilla 350gm Jar',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00534.B0_N',name:'LJ Multivitamin 7+ Gummies - 60s',casePack:40,boxPallet:15,cls:'B'},
    {code:'MWLJNTP.00698.B0_N',name:'LJ NutriMix Lite 7+ 350gm Chocolate Jar',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00699.B0_N',name:'LJ NutriMix 2+ Unsweetened 350gm Pouch',casePack:24,boxPallet:30,cls:'C'},
    {code:'MWLJNTP.000836.B0_N',name:'LJ Nutrimix 2+ Vanilla Lite 350gm',casePack:24,boxPallet:32,cls:'C'},
    {code:'MWBWNTL.00356.B0_N',name:'BB Magnesium Beads Gummies 60s',casePack:18,boxPallet:20,cls:'B'},
    {code:'MWMMNTL.00205.B0_N',name:'MM Creatine Monohydrate Powder Advanced 150g',casePack:36,boxPallet:30,cls:'C'},
    {code:'MWMMNTL.00207.B0_N',name:'MM Creatine Mixed Fruit Flavored 125g',casePack:24,boxPallet:32,cls:'B'},
    {code:'MWLJNTP.00536.B0_N',name:'LJ Millet Choco Crunch 250g',casePack:32,boxPallet:20,cls:'A'},
    {code:'MWLJNTP.00668.B0_N',name:'LJ Nutrimix 2+ Vanilla 1kg',casePack:11,boxPallet:30,cls:'B'},
  ];
}

function initBins(){
  const bins=[];
  for(let c=1;c<=17;c++){
    for(let i=1;i<=12;i++){
      const id=`R1-C${c}-${String(i).padStart(3,'0')}`;
      const classes=['A','B','C'];
      const fsn=classes[(c+i)%3];
      let status;
      if(c<=9) status=Math.random()<0.05?'Reserved':'Occupied';
      else if(c<=11) status=Math.random()<0.4?'Empty':Math.random()<0.5?'Reserved':'Occupied';
      else status='Empty';
      bins.push({id,zone:'R1',fsn,capacity:null,status,skuFromDump:''});
    }
  }
  return bins;
}

/* ════════════════════ LOGIN ════════════════════ */
/* ─────────────────────────────────────────────
   Sheet-based login with offline fallback
   ───────────────────────────────────────────── */
