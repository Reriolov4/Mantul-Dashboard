/**
 * HALAMAN 2 — OUTLETS | SALES FORCE
 * Sumber: RAW_DATA_SALES pada spreadsheet utama.
 * Kolom mengikuti struktur A:AH:
 * A tahun, B bulan, C tanggal, F subbrand, G brand, K distributor/subdist,
 * M customer name, AA tim, AB nama akun, AC category, AE SBU.
 */
function page2Source_() {
  const ss = getDatabase_();
  const sh = ss.getSheetByName('RAW_DATA_SALES');
  if (!sh || sh.getLastRow() < 2) throw new Error('Sheet RAW_DATA_SALES belum tersedia atau belum berisi data.');
  return sh;
}
function page2Month_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  const map = {JANUARI:'JAN',FEBRUARI:'FEB',MARET:'MAR',APRIL:'APR',MEI:'MAY',JUNI:'JUN',JULI:'JUL',AGUSTUS:'AUG',SEPTEMBER:'SEP',OKTOBER:'OCT',NOVEMBER:'NOV',DESEMBER:'DEC'};
  const en = {JANUARY:'JAN',FEBRUARY:'FEB',MARCH:'MAR',APRIL:'APR',MAY:'MAY',JUNE:'JUN',JULY:'JUL',AUGUST:'AUG',SEPTEMBER:'SEP',OCTOBER:'OCT',NOVEMBER:'NOV',DECEMBER:'DEC'};
  return map[s] || en[s] || s.slice(0,3);
}
function page2Text_(v) { return String(v == null ? '' : v).trim(); }
function page2Key_(v) { return page2Text_(v).toUpperCase(); }
function page2Read_() {
  const sh = page2Source_();
  // A:AH = 34 kolom; jangan membatasi pembacaan pada struktur file Excel lama.
  return sh.getRange(2, 1, sh.getLastRow() - 1, 31).getValues();
}
function page2Date_(year, month, day) {
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const mi = months.indexOf(page2Month_(month));
  let d = day;
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) d = d.getDate();
  d = Number(String(d).replace(/[^0-9]/g,''));
  if (!Number(year) || mi < 0 || !d || d > 31) return '';
  return String(Number(year)) + '-' + String(mi + 1).padStart(2,'0') + '-' + String(d).padStart(2,'0');
}
function page2Unique_(rows, index) {
  const o = {};
  rows.forEach(r => { const v = page2Text_(r[index]); if (v) o[v] = true; });
  return Object.keys(o).sort((a,b) => a.localeCompare(b,'id'));
}
function page2BuildMeta_() {
  const rows = page2Read_().filter(r => page2Key_(r[21]) === 'SOO');
  const years = page2Unique_(rows,0).sort((a,b) => Number(a)-Number(b));
  const monthOrder = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const monthsPresent = page2Unique_(rows,1).map(page2Month_);
  const dateKeys = page2Unique_(rows.map(r => [page2Date_(r[0],r[1],r[2])]).filter(r => r[0]),0);
  const months = monthOrder.filter(m => monthsPresent.includes(m));
  return {
    years: years,
    months: months,
    dateKeys: dateKeys,
    teams: page2Unique_(rows,26),
    subdists: page2Unique_(rows,10),
    brands: page2Unique_(rows,6),
    subbrands: page2Unique_(rows,5),
    accounts: page2Unique_(rows,27),
    categories: page2Unique_(rows,28),
    salesTypes: ['SOO'],
    distributors: ['ALL'],
    updatedAt: PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}
function page2FilterList_(v) {
  if (v == null || v === '' || v === 'ALL') return [];
  return (Array.isArray(v) ? v : [v]).map(page2Key_).filter(Boolean);
}
function page2Growth_(cy,ly) { return ly ? (cy / ly - 1) * 100 : (cy ? null : 0); }
function page2BuildDashboardData_(filters) {
  filters = filters || {};
  const all = page2Read_();
  const year = Number(filters.year) || new Date().getFullYear();
  const ly = year - 1;
  const selectedMonths = page2FilterList_(filters.months);
  const selectedDates = page2FilterList_(filters.dates);
  const selectedMonthDays = selectedDates;
  const teams = page2FilterList_(filters.teams);
  const subdists = page2FilterList_(filters.subdists);
  const brands = page2FilterList_(filters.brands);
  const subbrands = page2FilterList_(filters.subbrands);
  const accounts = page2FilterList_(filters.accounts);
  const categories = page2FilterList_(filters.categories);
  const matches = r => {
    if (page2Key_(r[21]) !== 'SOO') return false; // Jenis Sales dikunci SOO.
    if (teams.length && !teams.includes(page2Key_(r[26]))) return false;
    if (subdists.length && !subdists.includes(page2Key_(r[10]))) return false;
    if (brands.length && !brands.includes(page2Key_(r[6]))) return false;
    if (subbrands.length && !subbrands.includes(page2Key_(r[5]))) return false;
    if (accounts.length && !accounts.includes(page2Key_(r[27]))) return false;
    if (categories.length && !categories.includes(page2Key_(r[28]))) return false;
    return true;
  };
  const rows = all.filter(matches);
  const monthNames = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const periodMonths = selectedMonths.length ? selectedMonths : monthNames;
  const monthCount = periodMonths.length;
  const makeBucket = () => ({ ao:new Set(), ec:new Set(), groups:{subdist:{},account:{},category:{},sbu:{}} });
  const CY = makeBucket(), LY = makeBucket();
  const monthly = monthNames.map(m => ({month:m,ao:0,aoLy:0,ec:0,ecLy:0}));
  const dailyMap = {};
  const mtdMonth = periodMonths.slice().sort((a,b)=>monthNames.indexOf(a)-monthNames.indexOf(b)).slice(-1)[0] || monthNames[new Date().getMonth()];
  const customer = r => page2Key_(r[12]);
  const addGroup = (bucket, dim, label, cust, date) => {
    if (!label) return;
    if (!bucket.groups[dim][label]) bucket.groups[dim][label] = {name:label,ao:new Set(),ec:new Set()};
    bucket.groups[dim][label].ao.add(cust);
    bucket.groups[dim][label].ec.add(cust+'|'+date);
  };
  rows.forEach(r => {
    const y = Number(r[0]), m = page2Month_(r[1]);
    if ((y !== year && y !== ly) || !periodMonths.includes(m)) return;
    const cust = customer(r);
    const date = page2Date_(y,m,r[2]);
    if (!cust || !date || (selectedMonthDays.length && !selectedMonthDays.includes(date.slice(8,10)))) return;
    const b = y === year ? CY : LY;
    const ecKey = cust + '|' + date;
    b.ao.add(cust); b.ec.add(ecKey);
    addGroup(b,'subdist',page2Text_(r[10]),cust,date);
    addGroup(b,'account',page2Text_(r[27]),cust,date);
    addGroup(b,'category',page2Text_(r[28]),cust,date);
    addGroup(b,'sbu',page2Text_(r[30]),cust,date);
    const mi = monthNames.indexOf(m);
    const mBucket = monthly[mi];
    if (y === year) { /* monthly counts built below from month-specific sets */ }
    const dk = date;
    if (y === year) {
      if (!dailyMap[dk]) dailyMap[dk] = {date:dk,ao:new Set(),ec:new Set()};
      dailyMap[dk].ao.add(cust); dailyMap[dk].ec.add(ecKey);
    }
  });
  // Hitung AO/EC bulanan secara terpisah; AO = customer unik per bulan,
  // EC = pasangan customer + tanggal (transaksi harian outlet), bukan baris SKU.
  monthNames.forEach((m,idx) => {
    const sets = {cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()};
    rows.forEach(r => {
      const y=Number(r[0]); if (page2Month_(r[1])!==m || (y!==year && y!==ly) || !periodMonths.includes(m)) return;
      const c=customer(r), d=page2Date_(y,m,r[2]); if(!c||!d||(selectedMonthDays.length&&!selectedMonthDays.includes(d.slice(8,10))))return;
      (y===year?sets.cyAO:sets.lyAO).add(c);
      (y===year?sets.cyEC:sets.lyEC).add(c+'|'+d);
    });
    monthly[idx].ao=sets.cyAO.size; monthly[idx].aoLy=sets.lyAO.size;
    monthly[idx].ec=sets.cyEC.size; monthly[idx].ecLy=sets.lyEC.size;
  });
  const mtdRows = rows.filter(r => Number(r[0])===year && page2Month_(r[1])===mtdMonth && (!selectedMonthDays.length || selectedMonthDays.includes(page2Date_(year,r[1],r[2]).slice(8,10))));
  const mtdRowsLY = rows.filter(r => Number(r[0])===ly && page2Month_(r[1])===mtdMonth && (!selectedMonthDays.length || selectedMonthDays.includes(page2Date_(ly,r[1],r[2]).slice(8,10))));
  const uniqueCount = (arr, fn) => new Set(arr.map(fn).filter(Boolean)).size;
  const aoMTD = uniqueCount(mtdRows,customer), aoMTDLY = uniqueCount(mtdRowsLY,customer);
  const ecMTD = uniqueCount(mtdRows,r=>{const c=customer(r),d=page2Date_(year,r[1],r[2]);return c&&d?c+'|'+d:'';});
  const ecMTDLY = uniqueCount(mtdRowsLY,r=>{const c=customer(r),d=page2Date_(ly,r[1],r[2]);return c&&d?c+'|'+d:'';});
  const groupResult = (dim) => {
    const a=CY.groups[dim],b=LY.groups[dim],names=Array.from(new Set(Object.keys(a).concat(Object.keys(b))));
    return names.map(name=>{
      const cy=a[name],prev=b[name],ao=cy?cy.ao.size:0,aoLy=prev?prev.ao.size:0,ec=cy?cy.ec.size:0;
      return {name,ao,aoLy,ec,growth:page2Growth_(ao,aoLy)};
    }).sort((x,y)=>y.ao-x.ao||x.name.localeCompare(y.name,'id'));
  };
  const groupSbu = (sbu) => {
    const cy=CY.groups.sbu[sbu],prev=LY.groups.sbu[sbu];
    const ao=cy?cy.ao.size:0,aoLy=prev?prev.ao.size:0;
    return {name:sbu,aoCY:ao,aoLY:aoLy,growth:page2Growth_(ao,aoLy),ctr:CY.ao.size?ao/CY.ao.size*100:0};
  };
  const daily = Object.keys(dailyMap).sort().map(k=>({date:k,ao:dailyMap[k].ao.size,ec:dailyMap[k].ec.size}));
  const ao=CY.ao.size, aoLy=LY.ao.size, ec=CY.ec.size, ecLy=LY.ec.size;
  const compositionGroups = dim => groupResult(dim);
  return {
    kpi:{year,ly,months:periodMonths,monthCount,ao,aoLy,aoGrowth:page2Growth_(ao,aoLy),ec,ecLy,ecGrowth:page2Growth_(ec,ecLy),aoMTD,aoMTDLY,ecMTD,ecMTDLY,mtdMonth,rows:rows.length},
    monthly,daily,
    donut:{
      subdist:compositionGroups('subdist'),
      account:(() => {
        const groups={};
        rows.forEach(r=>{
          if(page2Key_(r[26]).indexOf('SALESMAN')<0)return;
          const y=Number(r[0]),m=page2Month_(r[1]);
          if((y!==year&&y!==ly)||!periodMonths.includes(m))return;
          const name=page2Text_(r[27]),c=customer(r),d=page2Date_(y,m,r[2]);
          if(!name||!c||!d||(selectedMonthDays.length&&!selectedMonthDays.includes(d.slice(8,10))))return;
          if(!groups[name])groups[name]={name,cyAO:new Set(),lyAO:new Set()};
          (y===year?groups[name].cyAO:groups[name].lyAO).add(c);
        });
        return Object.keys(groups).map(n=>{const ao=groups[n].cyAO.size,aoLy=groups[n].lyAO.size;return {name:n,ao,aoLy,growth:page2Growth_(ao,aoLy)};}).sort((a,b)=>b.ao-a.ao);
      })(),
      category:compositionGroups('category')
    },
    sbu:{b7:groupSbu('B7'),saka:groupSbu('SAKA')},
    byTeam:groupResult('sbu'), // data ringkas untuk ekspor; filter TIM tetap di atas.
    updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED')||'Live dari Google Sheets'
  };
}


/**
 * Cache layer for Page 2.
 * Cache version follows PAGE2_RAW_UPDATED, so a completed import naturally
 * starts a fresh cache namespace without attempting to enumerate cache keys.
 */
function page2CacheKey_(prefix, payload) {
  const version = PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'initial';
  const raw = String(version) + '|' + String(payload || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return prefix + '_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
}

function page2Cached_(key, builder, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); }
    catch (err) { cache.remove(key); }
  }
  const result = builder();
  const json = JSON.stringify(result);
  // Apps Script CacheService limits an individual cache value to about 100 KB.
  if (json.length <= 90000) {
    try { cache.put(key, json, ttlSeconds || 180); }
    catch (err) { /* Cache is an optimization, never a dependency. */ }
  }
  return result;
}

function page2GetMeta() {
  const key = page2CacheKey_('p2meta', 'v2');
  return page2Cached_(key, page2BuildMeta_, 300);
}

function getPage2DashboardData(filters) {
  const safeFilters = filters || {};
  const key = page2CacheKey_('p2data', JSON.stringify(safeFilters));
  return page2Cached_(key, function() {
    return page2BuildDashboardData_(safeFilters);
  }, 180);
}
