/**
 * HALAMAN 2 — OUTLETS | SALES FORCE
 * Membaca RAW_DATA_SALES berdasarkan NAMA HEADER, bukan posisi kolom lama.
 * Skema aktif workbook: A:AA transaksi; AH:AL tabel mapping/notes.
 *
 * Definisi:
 * - AO = distinct CUSTOMER NAME pada periode (tidak digandakan antar-subdist).
 * - EC = distinct tanggal + CUSTOMER NAME + SUBDIST.
 * - Salesman filter = daftar CUSTOMER NAME yang TIM-nya SALESMAN.
 * - KPI total tetap mencakup seluruh SOO; filter Salesman membatasi trend/
 *   komposisi, mengikuti perilaku dashboard sebelumnya.
 */
const PAGE2_MONTHS_ = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function page2Source_() {
  const ss = getDatabase_();
  const sh = ss.getSheetByName(RAW_DATA_SHEET_NAME || 'RAW_DATA_SALES');
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('Sheet RAW_DATA_SALES belum tersedia atau belum berisi data transaksi.');
  }
  return sh;
}

function page2Month_(value) {
  const s = String(value == null ? '' : value).trim().toUpperCase();
  const aliases = {
    JANUARI:'JAN',FEBRUARI:'FEB',MARET:'MAR',APRIL:'APR',MEI:'MAY',
    JUNI:'JUN',JULI:'JUL',AGUSTUS:'AUG',SEPTEMBER:'SEP',
    OKTOBER:'OCT',NOVEMBER:'NOV',DESEMBER:'DEC',
    JANUARY:'JAN',FEBRUARY:'FEB',MARCH:'MAR',MAY:'MAY',
    JUNE:'JUN',JULY:'JUL',AUGUST:'AUG',OCTOBER:'OCT',
    DECEMBER:'DEC'
  };
  if (aliases[s]) return aliases[s];
  if (/^\d{1,2}$/.test(s)) {
    const n = Number(s);
    return n >= 1 && n <= 12 ? PAGE2_MONTHS_[n - 1] : '';
  }
  return s.slice(0,3);
}

function page2Text_(value) {
  return String(value == null ? '' : value).trim();
}

function page2Key_(value) {
  return page2Text_(value).toUpperCase().replace(/\s+/g,' ');
}

function page2HeaderKey_(value) {
  return page2Key_(value).replace(/[_-]+/g,' ');
}

function page2FindHeader_(headers, aliases, fieldName) {
  const normalized = headers.map(page2HeaderKey_);
  for (let i = 0; i < aliases.length; i++) {
    const at = normalized.indexOf(page2HeaderKey_(aliases[i]));
    if (at >= 0) return at;
  }
  throw new Error(
    'Kolom wajib "' + fieldName + '" tidak ditemukan di RAW_DATA_SALES. Header yang terbaca: ' +
    headers.map(page2Text_).filter(Boolean).join(' | ')
  );
}

/**
 * Membaca hanya kolom yang dibutuhkan. Kolom dikelompokkan menjadi rentang
 * berurutan sehingga jumlah panggilan Spreadsheet tetap rendah tanpa membaca
 * seluruh 38 kolom (termasuk tabel mapping dan catatan di AH:AL).
 */
function page2Read_(sheet) {
  const sh = sheet || page2Source_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const fields = {
    year: page2FindHeader_(headers, ['TAHUN','YEAR'], 'TAHUN'),
    month: page2FindHeader_(headers, ['BULAN','MONTH'], 'BULAN'),
    day: page2FindHeader_(headers, ['TANGGAL','DATE','DAY'], 'TANGGAL'),
    subbrand: page2FindHeader_(headers, ['SUB BRAND','SUBBRAND'], 'SUB BRAND'),
    brand: page2FindHeader_(headers, ['BRAND'], 'BRAND'),
    subdist: page2FindHeader_(headers, ['SUBDIST','DISTRIBUTORS','DISTRIBUTOR'], 'SUBDIST'),
    customer: page2FindHeader_(headers, ['CUSTOMER NAME','NAMA AKUN','CUSTOMER'], 'CUSTOMER NAME'),
    category: page2FindHeader_(headers, ['CATEGORY','KATEGORI'], 'CATEGORY'),
    salesType: page2FindHeader_(headers, ['JENIS SALES','SALES TYPE'], 'JENIS SALES'),
    pic: page2FindHeader_(headers, ['PIC'], 'PIC'),
    tim: page2FindHeader_(headers, ['TIM','TEAM'], 'TIM'),
    sbu: page2FindHeader_(headers, ['SBU'], 'SBU')
  };

  const columns = Object.keys(fields).map(k => fields[k]).sort((a,b) => a-b);
  const runs = [];
  columns.forEach(col => {
    const current = runs[runs.length - 1];
    if (!current || col > current.end + 1) runs.push({start:col,end:col,values:null});
    else current.end = col;
  });

  const rowCount = lastRow - 1;
  runs.forEach(run => {
    run.values = sh.getRange(2, run.start + 1, rowCount, run.end - run.start + 1).getValues();
  });

  const colRun = {};
  runs.forEach(run => {
    for (let c = run.start; c <= run.end; c++) colRun[c] = run;
  });
  const read = (rowIndex, colIndex) => {
    const run = colRun[colIndex];
    return run.values[rowIndex][colIndex - run.start];
  };

  const rows = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    const rawTim = page2Key_(read(i, fields.tim));
    // Semua akun yang tidak masuk mapping Motoris/Canvaser diperlakukan
    // sebagai Salesman, sesuai aturan klasifikasi yang disepakati.
    const tim = rawTim || 'SALESMAN';
    rows[i] = [
      read(i, fields.year),                         // 0 year
      page2Month_(read(i, fields.month)),            // 1 month
      read(i, fields.day),                          // 2 day (DD supported)
      read(i, fields.subbrand),                     // 3 subbrand
      read(i, fields.brand),                        // 4 brand
      read(i, fields.customer),                     // 5 customer / Nama Akun
      read(i, fields.pic),                          // 6 PIC
      read(i, fields.salesType),                    // 7 SO / SOO
      read(i, fields.subdist),                      // 8 distributor / subdist
      tim,                                          // 9 team classification
      read(i, fields.customer),                     // 10 Nama Akun
      read(i, fields.category),                     // 11 category
      read(i, fields.sbu)                           // 12 SBU
    ];
  }
  return rows;
}

function page2Date_(year, month, day) {
  const mi = PAGE2_MONTHS_.indexOf(page2Month_(month));
  let d = null;
  if (Object.prototype.toString.call(day) === '[object Date]' && !isNaN(day)) {
    d = day.getDate();
  } else if (typeof day === 'number' && Number.isFinite(day)) {
    d = day > 31
      ? new Date(Date.UTC(1899, 11, 30) + day * 86400000).getUTCDate()
      : day;
  } else {
    const raw = page2Text_(day);
    let match = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) d = Number(match[3]);
    else {
      match = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
      if (match) d = Number(match[1]);
      else if (/^\d{1,2}$/.test(raw)) d = Number(raw);
      else {
        const parsed = Date.parse(raw);
        if (!isNaN(parsed)) d = new Date(parsed).getDate();
      }
    }
  }
  if (!Number(year) || mi < 0 || !d || d < 1 || d > 31) return '';
  return String(Number(year)) + '-' + String(mi + 1).padStart(2,'0') + '-' + String(d).padStart(2,'0');
}

function page2BuildMeta_(sourceRows) {
  const values = sourceRows || page2Read_();
  const years = new Set(), months = new Set(), pics = new Set();
  const salesmen = new Set(), subdists = new Set(), brands = new Set();
  const subbrands = new Set(), accounts = new Set(), categories = new Set();
  const dateSet = new Set(), monthsByYear = {};
  const salesTypeCounts = {};
  let sooRowCount = 0;

  values.forEach(r => {
    const type = page2Key_(r[7]);
    if (type) salesTypeCounts[type] = (salesTypeCounts[type] || 0) + 1;
    if (type !== 'SOO') return;
    sooRowCount++;

    const year = page2Text_(r[0]);
    const month = page2Month_(r[1]);
    if (year) years.add(year);
    if (month) months.add(month);
    if (page2Text_(r[6])) pics.add(page2Text_(r[6]));
    if (page2Text_(r[8])) subdists.add(page2Text_(r[8]));
    if (page2Text_(r[4])) brands.add(page2Text_(r[4]));
    if (page2Text_(r[3])) subbrands.add(page2Text_(r[3]));
    if (page2Text_(r[10])) accounts.add(page2Text_(r[10]));
    if (page2Text_(r[11])) categories.add(page2Text_(r[11]));
    if (page2Key_(r[9]) === 'SALESMAN' && page2Text_(r[10])) salesmen.add(page2Text_(r[10]));

    if (year && month) {
      if (!monthsByYear[year]) monthsByYear[year] = new Set();
      monthsByYear[year].add(month);
    }
    const date = page2Date_(year, month, r[2]);
    if (date) dateSet.add(date);
  });

  const sorted = set => Array.from(set).sort((a,b) => a.localeCompare(b,'id'));
  return {
    years: sorted(years).sort((a,b) => Number(a)-Number(b)),
    months: PAGE2_MONTHS_.filter(m => months.has(m)),
    monthsByYear: Object.fromEntries(Object.keys(monthsByYear).map(y => [
      y, PAGE2_MONTHS_.filter(m => monthsByYear[y].has(m))
    ])),
    dateKeys: Array.from(dateSet).sort(),
    pics: sorted(pics),
    salesmen: sorted(salesmen),
    subdists: sorted(subdists),
    brands: sorted(brands),
    subbrands: sorted(subbrands),
    accounts: sorted(accounts),
    categories: sorted(categories),
    salesTypes: ['SOO'],
    distributors: ['ALL'],
    sourceRowCount: values.length,
    sooRowCount: sooRowCount,
    salesTypeCounts: salesTypeCounts,
    updatedAt: PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}

function page2FilterList_(value) {
  if (value == null || value === '' || value === 'ALL') return [];
  return (Array.isArray(value) ? value : [value]).map(page2Key_).filter(Boolean);
}

function page2Growth_(current, previous) {
  return previous ? (current / previous - 1) * 100 : (current ? null : 0);
}

function page2CompositeKey_(parts) {
  return parts.map(value => {
    const s = page2Key_(value);
    return s.length + ':' + s;
  }).join('|');
}

function page2BuildDashboardData_(filters, sourceRows) {
  filters = filters || {};
  const all = sourceRows || page2Read_();
  const year = Number(filters.year) || new Date().getFullYear();
  const ly = year - 1;
  const selectedMonths = page2FilterList_(filters.months);
  const pics = page2FilterList_(filters.pics);
  const subdists = page2FilterList_(filters.subdists);
  const brands = page2FilterList_(filters.brands);
  const subbrands = page2FilterList_(filters.subbrands);
  const categories = page2FilterList_(filters.categories);
  const salesmen = page2FilterList_(filters.salesmen);
  const selectedDates = page2FilterList_(filters.dates);
  const hasDateFilter = Array.isArray(filters.dates);
  const periodMonths = Array.isArray(filters.months) ? selectedMonths : PAGE2_MONTHS_;
  const monthCount = periodMonths.length;
  const matches = r => {
    if (page2Key_(r[7]) !== 'SOO') return false;
    if (pics.length && !pics.includes(page2Key_(r[6]))) return false;
    if (subdists.length && !subdists.includes(page2Key_(r[8]))) return false;
    if (brands.length && !brands.includes(page2Key_(r[4]))) return false;
    if (subbrands.length && !subbrands.includes(page2Key_(r[3]))) return false;
    if (categories.length && !categories.includes(page2Key_(r[11]))) return false;
    return true;
  };

  const makeBucket = () => ({ao:new Set(),ec:new Set(),groups:{subdist:{},account:{},category:{},sbu:{}}});
  const CY = makeBucket(), LY = makeBucket(), CHART_CY = makeBucket(), CHART_LY = makeBucket();
  const monthly = PAGE2_MONTHS_.map(month => ({month:month,ao:0,aoLy:0,ec:0,ecLy:0}));
  const monthlySets = PAGE2_MONTHS_.map(() => ({cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()}));
  const dailyMap = {};
  const orderedSelected = periodMonths.slice().sort((a,b) => PAGE2_MONTHS_.indexOf(a)-PAGE2_MONTHS_.indexOf(b));
  const mtdMonth = orderedSelected.length ? orderedSelected[orderedSelected.length - 1] : null;
  const mtd = {cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()};
  const customer = r => page2Key_(r[5]);
  const outlet = r => customer(r); // AO counts unique account names
  const ecKeyFor = (r, date) => page2CompositeKey_([date, customer(r), r[8]]);

  const addGroup = (bucket, dimension, label, outletKey, ecKey) => {
    if (!label) return;
    if (!bucket.groups[dimension][label]) {
      bucket.groups[dimension][label] = {name:label,ao:new Set(),ec:new Set()};
    }
    bucket.groups[dimension][label].ao.add(outletKey);
    bucket.groups[dimension][label].ec.add(ecKey);
  };

  let matchedRowCount = 0;
  all.forEach(r => {
    if (!matches(r)) return;
    matchedRowCount++;

    const y = Number(r[0]), month = page2Month_(r[1]);
    if ((y !== year && y !== ly) || periodMonths.indexOf(month) < 0) return;
    const customerName = customer(r);
    if (!customerName) return;
    const date = page2Date_(y, month, r[2]);
    if (!date) return;

    const outletKey = outlet(r);
    const ecKey = ecKeyFor(r, date);
    const isCY = y === year;
    const totalBucket = isCY ? CY : LY;

    totalBucket.ao.add(outletKey);
    totalBucket.ec.add(ecKey);
    addGroup(totalBucket, 'sbu', page2Key_(r[12]), outletKey, ecKey);

    if (month === mtdMonth) {
      (isCY ? mtd.cyAO : mtd.lyAO).add(outletKey);
      (isCY ? mtd.cyEC : mtd.lyEC).add(ecKey);
    }

    // Dropdown SALESMAN hanya berisi Nama Akun yang diklasifikasikan SALESMAN.
    // Saat dipilih, filter membatasi trend dan komposisi; KPI total tetap total SOO.
    if (salesmen.length && !salesmen.includes(customerName)) return;

    const chartBucket = isCY ? CHART_CY : CHART_LY;
    addGroup(chartBucket, 'subdist', page2Text_(r[8]), outletKey, ecKey);
    addGroup(chartBucket, 'account', page2Text_(r[10]), outletKey, ecKey);
    addGroup(chartBucket, 'category', page2Text_(r[11]), outletKey, ecKey);

    const monthIndex = PAGE2_MONTHS_.indexOf(month);
    if (monthIndex >= 0) {
      const ms = monthlySets[monthIndex];
      (isCY ? ms.cyAO : ms.lyAO).add(outletKey);
      (isCY ? ms.cyEC : ms.lyEC).add(ecKey);
    }

    const dayKey = date.slice(8,10);
    if (isCY && (!hasDateFilter || selectedDates.includes(dayKey))) {
      if (!dailyMap[dayKey]) dailyMap[dayKey] = {ao:new Set(),ec:new Set()};
      dailyMap[dayKey].ao.add(outletKey);
      dailyMap[dayKey].ec.add(ecKey);
    }
  });

  monthlySets.forEach((sets,index) => {
    monthly[index].ao = sets.cyAO.size;
    monthly[index].aoLy = sets.lyAO.size;
    monthly[index].ec = sets.cyEC.size;
    monthly[index].ecLy = sets.lyEC.size;
  });

  const groupResult = (dimension, currentBucket, previousBucket) => {
    const current = currentBucket.groups[dimension];
    const previous = previousBucket.groups[dimension];
    const names = Array.from(new Set(Object.keys(current).concat(Object.keys(previous))));
    return names.map(name => {
      const cy = current[name], lyGroup = previous[name];
      const ao = cy ? cy.ao.size : 0;
      const aoLy = lyGroup ? lyGroup.ao.size : 0;
      const ec = cy ? cy.ec.size : 0;
      return {name:name,ao:ao,aoLy:aoLy,ec:ec,growth:page2Growth_(ao,aoLy)};
    }).sort((a,b) => b.ao-a.ao || a.name.localeCompare(b.name,'id'));
  };

  const groupSbu = sbu => {
    const current = CY.groups.sbu[sbu], previous = LY.groups.sbu[sbu];
    const ao = current ? current.ao.size : 0;
    const aoLy = previous ? previous.ao.size : 0;
    return {
      name:sbu,aoCY:ao,aoLY:aoLy,growth:page2Growth_(ao,aoLy),
      ctr:CY.ao.size ? ao / CY.ao.size * 100 : 0
    };
  };

  const daily = Array.from({length:31},(_,index) => {
    const day = String(index + 1).padStart(2,'0');
    const item = dailyMap[day];
    return {day:day,ao:item ? item.ao.size : 0,ec:item ? item.ec.size : 0};
  });

  const ao = CY.ao.size, aoLy = LY.ao.size, ec = CY.ec.size, ecLy = LY.ec.size;
  const aoMTD = mtd.cyAO.size, aoMTDLY = mtd.lyAO.size;
  const ecMTD = mtd.cyEC.size, ecMTDLY = mtd.lyEC.size;
  const composition = dimension => groupResult(dimension, CHART_CY, CHART_LY);

  return {
    kpi:{
      year:year,ly:ly,months:periodMonths,monthCount:monthCount,
      ao:ao,aoLy:aoLy,aoGrowth:page2Growth_(ao,aoLy),
      ec:ec,ecLy:ecLy,ecGrowth:page2Growth_(ec,ecLy),
      aoMTD:aoMTD,aoMTDLY:aoMTDLY,ecMTD:ecMTD,ecMTDLY:ecMTDLY,
      mtdMonth:mtdMonth,rows:matchedRowCount
    },
    monthly:monthly,
    daily:daily,
    donut:{
      subdist:composition('subdist'),
      account:composition('account'),
      category:composition('category')
    },
    sbu:{b7:groupSbu('B7'),saka:groupSbu('SAKA')},
    byTeam:groupResult('sbu',CY,LY),
    updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}

function page2SourceVersion_(sheet) {
  const sh = sheet || page2Source_();
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID || '').trim();
  return [
    spreadsheetId, sh.getSheetId(), sh.getLastRow(), sh.getLastColumn(),
    props.getProperty('PAGE2_RAW_UPDATED') || 'manual'
  ].join('|');
}

function page2CacheKey_(prefix, payload, sourceVersion) {
  const raw = String(sourceVersion || '') + '|' + String(payload || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return prefix + '_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g,'');
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
  if (json.length <= 90000) {
    try { cache.put(key, json, ttlSeconds || 300); }
    catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }
  }
  return result;
}

function page2GetMeta() {
  const sh = page2Source_();
  const version = page2SourceVersion_(sh);
  const key = page2CacheKey_('p2meta_schema2', 'soo-header-mapped', version);
  return page2Cached_(key, function() {
    return page2BuildMeta_(page2Read_(sh));
  }, 300);
}

function getPage2DashboardData(filters) {
  const sh = page2Source_();
  const safeFilters = filters || {};
  const version = page2SourceVersion_(sh);
  const key = page2CacheKey_('p2data_schema2', JSON.stringify(safeFilters), version);
  return page2Cached_(key, function() {
    return page2BuildDashboardData_(safeFilters, page2Read_(sh));
  }, 300);
}

/**
 * Initial payload membaca sumber satu kali dan membangun metadata + KPI/grafik.
 * Versi cache memuat sheetId, jumlah baris/kolom, serta timestamp import,
 * sehingga append data baru tidak memakai hasil cache dataset sebelumnya.
 */
function getPage2InitialPayload() {
  const sh = page2Source_();
  const version = page2SourceVersion_(sh);
  const metaKey = page2CacheKey_('p2meta_schema2', 'initial', version);
  const dataKey = page2CacheKey_('p2initial_schema2', 'initial', version);
  const cache = CacheService.getScriptCache();
  const cachedMeta = cache.get(metaKey);
  const cachedData = cache.get(dataKey);
  if (cachedMeta && cachedData) {
    try {
      return {meta:JSON.parse(cachedMeta),data:JSON.parse(cachedData)};
    } catch (err) {
      cache.remove(metaKey);
      cache.remove(dataKey);
    }
  }

  const sourceRows = page2Read_(sh);
  const meta = page2BuildMeta_(sourceRows);
  const years = meta.years || [];
  const year = Number(years.indexOf('2026') >= 0 ? '2026' : years[years.length - 1]) || new Date().getFullYear();
  const monthSet = new Set();
  sourceRows.forEach(r => {
    if (Number(r[0]) === year && page2Key_(r[7]) === 'SOO') {
      const month = page2Month_(r[1]);
      if (PAGE2_MONTHS_.indexOf(month) >= 0) monthSet.add(month);
    }
  });
  const initialMonths = PAGE2_MONTHS_.filter(month => monthSet.has(month));
  const data = page2BuildDashboardData_({year:year,months:initialMonths}, sourceRows);

  const metaJson = JSON.stringify(meta);
  const dataJson = JSON.stringify(data);
  try {
    if (metaJson.length <= 90000) cache.put(metaKey, metaJson, 300);
    if (dataJson.length <= 90000) cache.put(dataKey, dataJson, 300);
  } catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }
  return {meta:meta,data:data};
}
