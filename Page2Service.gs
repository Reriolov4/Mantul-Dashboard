/**
 * HALAMAN 2 — OUTLETS | SALES FORCE
 * Sumber: RAW_DATA_SALES pada spreadsheet utama.
 * Kolom mengikuti struktur A:AH:
 * A tahun, B bulan, C tanggal, F subbrand, G brand, M customer name,
 * K distributors (label filter SUBDIST), V jenis sales, AA tim, AB nama akun (label filter PIC),
 * AC category, AE SBU. K adalah DISTRIBUTORS, bukan SUBDIST.
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
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  // Ambil hanya 13 field yang benar-benar dipakai Page 2.
  // Lima rentang dipilih agar tidak menarik kolom transaksi/teks yang tidak relevan.
  const count = lastRow - 1;
  const abc = sh.getRange(2, 1, count, 3).getValues();    // A:C tahun, bulan, tanggal
  const fg = sh.getRange(2, 6, count, 2).getValues();     // F:G subbrand, brand
  const k = sh.getRange(2, 11, count, 1).getValues();      // K distributors
  const mno = sh.getRange(2, 13, count, 1).getValues();   // M customer name
  const vy = sh.getRange(2, 22, count, 1).getValues();    // V jenis sales
  const aaAe = sh.getRange(2, 27, count, 5).getValues();  // AA:AE tim, nama akun, category, DOT, SBU

  return abc.map((row, i) => [
    row[0], row[1], row[2],
    fg[i][0], fg[i][1],
    mno[i][0], aaAe[i][1],
    vy[i][0], k[i][0],
    aaAe[i][0], aaAe[i][1], aaAe[i][2], aaAe[i][4]
  ]);
}
function page2Date_(year, month, day) {
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  let mi = months.indexOf(page2Month_(month));
  if (mi < 0 && /^\d{1,2}$/.test(String(month).trim())) mi = Number(month) - 1;
  let d = null;
  if (Object.prototype.toString.call(day) === '[object Date]' && !isNaN(day)) {
    d = day.getDate();
  } else if (typeof day === 'number' && Number.isFinite(day)) {
    // Numeric values above 31 may be Google Sheets date serials.
    d = day > 31 ? new Date(Date.UTC(1899, 11, 30) + day * 86400000).getUTCDate() : day;
  } else {
    const raw = String(day == null ? '' : day).trim();
    let m = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) d = Number(m[3]);
    else {
      m = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
      if (m) d = Number(m[1]);
      else if (/^\d{1,2}$/.test(raw)) d = Number(raw);
      else {
        const parsed = Date.parse(raw);
        if (!isNaN(parsed)) d = new Date(parsed).getDate();
      }
    }
  }
  if (!Number(year) || mi < 0 || !d || d > 31) return '';
  return String(Number(year)) + '-' + String(mi + 1).padStart(2,'0') + '-' + String(d).padStart(2,'0');
}
function page2Unique_(rows, index) {
  const o = {};
  rows.forEach(r => { const v = page2Text_(r[index]); if (v) o[v] = true; });
  return Object.keys(o).sort((a,b) => a.localeCompare(b,'id'));
}
function page2BuildMeta_(sourceRows) {
  // Bangun seluruh opsi filter dalam satu lintasan; hindari membuat beberapa
  // array perantara besar dari RAW_DATA_SALES.
  const values = sourceRows || page2Read_();
  const fields = {years:0, months:1, pics:6, subdists:8, teams:9, brands:4, subbrands:3, accounts:10, categories:11};
  const sets = {};
  Object.keys(fields).forEach(k => sets[k] = new Set());
  const dateSet = new Set(), monthsByYear = {};
  const salesTypeCounts = {};
  let sooRowCount = 0;
  values.forEach(r => {
    const salesType = page2Key_(r[7]);
    if (salesType) salesTypeCounts[salesType] = (salesTypeCounts[salesType] || 0) + 1;
    if (salesType !== 'SOO') return;
    sooRowCount++;
    Object.keys(fields).forEach(k => {
      const value = page2Text_(r[fields[k]]);
      if (value) sets[k].add(k === 'months' ? page2Month_(value) : value);
    });
    const normalizedMonth = page2Month_(r[1]);
    const yearKey = page2Text_(r[0]);
    if (yearKey && normalizedMonth) {
      if (!monthsByYear[yearKey]) monthsByYear[yearKey] = new Set();
      monthsByYear[yearKey].add(normalizedMonth);
    }
    const date = page2Date_(r[0], r[1], r[2]);
    if (date) dateSet.add(date);
  });
  const monthOrder = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const sorted = key => Array.from(sets[key]).sort((a,b) => a.localeCompare(b,'id'));
  return {
    years: sorted('years').sort((a,b) => Number(a)-Number(b)),
    months: monthOrder.filter(m => sets.months.has(m)),
    monthsByYear: Object.fromEntries(Object.keys(monthsByYear).map(y => [y, monthOrder.filter(m => monthsByYear[y].has(m))])),
    dateKeys: Array.from(dateSet).sort(),
    pics: sorted('pics'),
    subdists: sorted('subdists'),
    brands: sorted('brands'),
    subbrands: sorted('subbrands'),
    accounts: sorted('accounts'),
    categories: sorted('categories'),
    salesTypes: ['SOO'],
    distributors: ['ALL'],
    sourceRowCount: values.length,
    sooRowCount: sooRowCount,
    salesTypeCounts: salesTypeCounts,
    updatedAt: PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}
function page2FilterList_(v) {
  if (v == null || v === '' || v === 'ALL') return [];
  return (Array.isArray(v) ? v : [v]).map(page2Key_).filter(Boolean);
}
function page2Growth_(cy,ly) { return ly ? (cy / ly - 1) * 100 : (cy ? null : 0); }
function page2BuildDashboardData_(filters, sourceRows) {
  filters = filters || {};
  const all = sourceRows || page2Read_();
  const year = Number(filters.year) || new Date().getFullYear();
  const ly = year - 1;
  const selectedMonths = page2FilterList_(filters.months);
  // Date checklist intentionally does not constrain KPI or composition.
  const pics = page2FilterList_(filters.pics);
  const subdists = page2FilterList_(filters.subdists);
  const brands = page2FilterList_(filters.brands);
  const subbrands = page2FilterList_(filters.subbrands);
  const accounts = page2FilterList_(filters.accounts);
  const categories = page2FilterList_(filters.categories);
  const teams = page2FilterList_(filters.teams);
  const matches = r => {
    if (page2Key_(r[7]) !== 'SOO') return false; // Jenis Sales dikunci SOO.
    if (pics.length && !pics.includes(page2Key_(r[6]))) return false;
    if (subdists.length && !subdists.includes(page2Key_(r[8]))) return false;
    if (brands.length && !brands.includes(page2Key_(r[4]))) return false;
    if (subbrands.length && !subbrands.includes(page2Key_(r[3]))) return false;
    if (accounts.length && !accounts.includes(page2Key_(r[10]))) return false;
    if (categories.length && !categories.includes(page2Key_(r[11]))) return false;
    if (teams.length && !teams.includes(page2Key_(r[9]))) return false;
    return true;
  };
  const rows = all.filter(matches);
  const monthNames = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const periodMonths = Array.isArray(filters.months) ? selectedMonths : monthNames;
  const monthCount = periodMonths.length;
  const makeBucket = () => ({ ao:new Set(), ec:new Set(), groups:{subdist:{},account:{},category:{},sbu:{}} });
  const CY = makeBucket(), LY = makeBucket();
  const monthly = monthNames.map(m => ({month:m,ao:0,aoLy:0,ec:0,ecLy:0}));
  const monthlySets = monthNames.map(() => ({cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()}));
  const dailyMap = {};
  // MTD mengikuti bulan terpilih paling akhir. Jika filter bulan kosong
  // (artinya semua bulan), gunakan bulan terbaru yang benar-benar tersedia
  // pada tahun berjalan, bukan otomatis Desember.
  const selectedMtdMonth = periodMonths.slice().sort((a,b)=>monthNames.indexOf(a)-monthNames.indexOf(b)).slice(-1)[0];
  const latestAvailableMonth = rows.reduce((latest, r) => {
    if (Number(r[0]) !== year) return latest;
    const idx = monthNames.indexOf(page2Month_(r[1]));
    return idx > latest ? idx : latest;
  }, -1);
  const mtdMonth = periodMonths.length ? (selectedMtdMonth || (latestAvailableMonth >= 0 ? monthNames[latestAvailableMonth] : monthNames[new Date().getMonth()])) : null;
  const mtdAO = {cy:new Set(),ly:new Set(),cyEC:new Set(),lyEC:new Set()};
  const customer = r => page2Key_(r[5]);
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
    if (!cust || !date) return;
    const b = y === year ? CY : LY;
    const ecKey = cust + '|' + date;
    b.ao.add(cust); b.ec.add(ecKey);
    const monthIndex = monthNames.indexOf(m);
    if (monthIndex >= 0) {
      const ms = monthlySets[monthIndex];
      (y === year ? ms.cyAO : ms.lyAO).add(cust);
      (y === year ? ms.cyEC : ms.lyEC).add(ecKey);
    }
    addGroup(b,'subdist',page2Text_(r[8]),cust,date);
    addGroup(b,'account',page2Text_(r[10]),cust,date);
    addGroup(b,'category',page2Text_(r[11]),cust,date);
    addGroup(b,'sbu',page2Text_(r[12]),cust,date);
    if (m === mtdMonth) {
      (y === year ? mtdAO.cy : mtdAO.ly).add(cust);
      (y === year ? mtdAO.cyEC : mtdAO.lyEC).add(ecKey);
    }
    const dk = date;
    if (y === year) {
      if (!dailyMap[dk]) dailyMap[dk] = {date:dk,ao:new Set(),ec:new Set()};
      dailyMap[dk].ao.add(cust); dailyMap[dk].ec.add(ecKey);
    }
  });
  // Monthly AO/EC sudah dikumpulkan pada lintasan utama di atas.
  monthlySets.forEach((sets,idx) => {
    monthly[idx].ao=sets.cyAO.size; monthly[idx].aoLy=sets.lyAO.size;
    monthly[idx].ec=sets.cyEC.size; monthly[idx].ecLy=sets.lyEC.size;
  });
  const aoMTD = mtdAO.cy.size, aoMTDLY = mtdAO.ly.size;
  const ecMTD = mtdAO.cyEC.size, ecMTDLY = mtdAO.lyEC.size;
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
  // Roll-up per nomor tanggal: jumlahkan nilai harian dari seluruh bulan
  // yang dipilih. AO tetap unik per tanggal; EC tetap unik per customer-tanggal.
  const dailyTotals = Array.from({length:31},(_,i)=>({day:String(i+1).padStart(2,'0'),ao:0,ec:0}));
  Object.keys(dailyMap).forEach(k => {
    const dayIndex = Number(k.slice(8,10)) - 1;
    if (dayIndex < 0 || dayIndex >= 31) return;
    dailyTotals[dayIndex].ao += dailyMap[k].ao.size;
    dailyTotals[dayIndex].ec += dailyMap[k].ec.size;
  });
  const daily = dailyTotals;
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
          if(page2Key_(r[9]).indexOf('SALESMAN')<0)return;
          const y=Number(r[0]),m=page2Month_(r[1]);
          if((y!==year&&y!==ly)||!periodMonths.includes(m))return;
          const name=page2Text_(r[10]),c=customer(r),d=page2Date_(y,m,r[2]);
          if(!name||!c||!d)return;
          if(!groups[name])groups[name]={name,cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()};
          const bucket=groups[name];
          (y===year?bucket.cyAO:bucket.lyAO).add(c);
          (y===year?bucket.cyEC:bucket.lyEC).add(c+'|'+d);
        });
        return Object.keys(groups).map(n=>{const ao=groups[n].cyAO.size,aoLy=groups[n].lyAO.size,ec=groups[n].cyEC.size;return {name:n,ao,aoLy,ec,growth:page2Growth_(ao,aoLy)};}).sort((a,b)=>b.ao-a.ao);
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
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID || '').trim();
  const version = props.getProperty('PAGE2_RAW_UPDATED') || 'initial';
  // Cache dipisahkan per Spreadsheet agar hasil database lama tidak terbaca
  // setelah konfigurasi sumber data berubah.
  const raw = spreadsheetId + '|' + String(version) + '|' + String(payload || '');
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
  const key = page2CacheKey_('p2meta', 'v6-teams');
  return page2Cached_(key, page2BuildMeta_, 300);
}

function getPage2DashboardData(filters) {
  const safeFilters = filters || {};
  const key = page2CacheKey_('p2data_v3', JSON.stringify(safeFilters));
  return page2Cached_(key, function() {
    return page2BuildDashboardData_(safeFilters);
  }, 180);
}


/**
 * Initial Page 2 payload: read RAW_DATA_SALES once, then build both filter
 * metadata and dashboard aggregates from the same compact row set.
 */
function getPage2InitialPayload() {
  const key = page2CacheKey_('p2initial', 'v5-teams');
  return page2Cached_(key, function() {
    const sourceRows = page2Read_();
    const meta = page2BuildMeta_(sourceRows);
    const years = meta.years || [];
    const year = years.includes('2026')
      ? 2026
      : (Number(years[years.length - 1]) || new Date().getFullYear());
    // Default filter bulan sama dengan bulan yang tersedia untuk tahun awal.
    // Dengan begitu MTD initial-load konsisten dengan pilihan filter di UI.
    const monthOrder = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    const initialMonths = Array.from(new Set(sourceRows
      .filter(r => Number(r[0]) === year && page2Key_(r[7]) === 'SOO')
      .map(r => page2Month_(r[1]))
      .filter(m => monthOrder.includes(m))))
      .sort((a,b) => monthOrder.indexOf(a)-monthOrder.indexOf(b));
    const data = page2BuildDashboardData_({ year: year, months: initialMonths }, sourceRows);
    return { meta: meta, data: data };
  }, 180);
}
