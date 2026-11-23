/**
 * HALAMAN 2 — OUTLETS | SALES FORCE
 * Membaca RAW_DATA_SALES berdasarkan NAMA HEADER, bukan posisi kolom lama.
 * Skema aktif workbook: A:AA transaksi; AH:AI catatan dan tabel mapping terpisah.
 *
 * Definisi:
 * - AO = distinct KODE CUSTOMER valid pada periode; baris tanpa kode (termasuk OTHERS) tidak dihitung sebagai outlet.
 * - EC = distinct tanggal + identitas outlet + SUBDIST.
 * - Customer picker memfilter KPI, trend, dan komposisi secara konsisten.
 */
const PAGE2_MONTHS_ = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function page2Source_() {
  // Prefer the verified normalized index. If it is missing or stale, keep the
  // existing RAW_DATA_SALES route as a safe fallback.
  const indexed = analyticsGetReadySheet_('P2_EVENT_INDEX', RAW_DATA_SHEET_NAME, 'P2');
  if (indexed && indexed.getLastRow() > 1) return indexed;

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
 * Membaca hanya kolom transaksi yang dibutuhkan. Kolom dikelompokkan menjadi
 * rentang berurutan A:C, F:G, I:K, R:X; AH:AI helper formula tidak dibaca.
 */
function page2Read_(sheet) {
  const sh = sheet || page2Source_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  // Analytics index is already normalized to 15 columns (including VAL) and contains SOO only.
  // Convert it into the exact internal row shape consumed by existing KPI logic.
  if (sh.getName() === 'P2_EVENT_INDEX') {
    const values = sh.getRange(2, 1, lastRow - 1, 15).getValues();
    return values.map(r => [
      r[0], page2Month_(r[1]), r[2], r[8], r[7], r[4], r[6],
      r[11], r[5], page2Key_(r[13]) || 'SALESMAN', r[4], r[9],
      r[10], r[3], Number(r[14]) || 0
    ]);
  }

  // Header hanya A:AA. AH:AI merupakan helper formula workbook, bukan data dashboard.
  const headers = sh.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT).getValues()[0];
  const fields = {
    year: page2FindHeader_(headers, ['TAHUN','YEAR'], 'TAHUN'),
    month: page2FindHeader_(headers, ['BULAN','MONTH'], 'BULAN'),
    day: page2FindHeader_(headers, ['TANGGAL','DATE','DAY'], 'TANGGAL'),
    subbrand: page2FindHeader_(headers, ['SUB BRAND','SUBBRAND'], 'SUB BRAND'),
    brand: page2FindHeader_(headers, ['BRAND'], 'BRAND'),
    customerCode: page2FindHeader_(headers, ['KODE CUSTOMER','CUSTOMER CODE','KODE AKUN'], 'KODE CUSTOMER'),
    customer: page2FindHeader_(headers, ['CUSTOMER NAME','NAMA AKUN','CUSTOMER'], 'CUSTOMER NAME'),
    subdist: page2FindHeader_(headers, ['SUBDIST','DISTRIBUTORS','DISTRIBUTOR'], 'SUBDIST'),
    category: page2FindHeader_(headers, ['CATEGORY','KATEGORI'], 'CATEGORY'),
    salesType: page2FindHeader_(headers, ['JENIS SALES','SALES TYPE'], 'JENIS SALES'),
    pic: page2FindHeader_(headers, ['PIC'], 'PIC'),
    tim: page2FindHeader_(headers, ['TIM','TEAM'], 'TIM'),
    sbu: page2FindHeader_(headers, ['SBU'], 'SBU'),
    value: page2FindHeader_(headers, ['VAL','VALUE'], 'VAL')
  };

  // Empat batch berurutan menggantikan banyak getRange kecil:
  // A:C, F:G, I:K, R:X. Total 15 kolom yang benar-benar dibutuhkan.
  const rowCount = lastRow - 1;
  const blocks = [
    {start:0,end:2,values:sh.getRange(2,1,rowCount,3).getValues()},
    {start:5,end:6,values:sh.getRange(2,6,rowCount,2).getValues()},
    {start:8,end:10,values:sh.getRange(2,9,rowCount,3).getValues()},
    {start:12,end:12,values:sh.getRange(2,13,rowCount,1).getValues()},
    {start:17,end:23,values:sh.getRange(2,18,rowCount,7).getValues()}
  ];
  // Bangun lookup kolom satu kali. Hindari Array.find() untuk setiap sel
  // pada 132 ribu baris (lebih dari 1,8 juta pencarian blok).
  const columnReaders = {};
  blocks.forEach(block => {
    for (let col = block.start; col <= block.end; col++) {
      columnReaders[col] = {values:block.values, offset:col - block.start};
    }
  });
  const read = (rowIndex,colIndex) => {
    const column = columnReaders[colIndex];
    return column ? column.values[rowIndex][column.offset] : '';
  };

  const rows = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    const rawTim = page2Key_(read(i, fields.tim));
    rows[i] = [
      read(i, fields.year),                         // 0 year
      page2Month_(read(i, fields.month)),            // 1 month
      read(i, fields.day),                          // 2 day
      read(i, fields.subbrand),                     // 3 subbrand
      read(i, fields.brand),                        // 4 brand
      read(i, fields.customer),                     // 5 display customer name
      read(i, fields.pic),                          // 6 PIC
      read(i, fields.salesType),                    // 7 SO / SOO
      read(i, fields.subdist),                      // 8 distributor / subdist
      rawTim || 'SALESMAN',                         // 9 team classification
      read(i, fields.customer),                     // 10 Nama Akun
      read(i, fields.category),                     // 11 category
      read(i, fields.sbu),                          // 12 SBU
      read(i, fields.customerCode),                 // 13 stable outlet identity
      Number(read(i, fields.value)) || 0            // 14 VAL
    ];
  }
  return rows;
}

function page2Day_(day) {
  let d = null;
  if (Object.prototype.toString.call(day) === '[object Date]' && !isNaN(day)) d = day.getDate();
  else if (typeof day === 'number' && Number.isFinite(day)) {
    d = day > 31 ? new Date(Date.UTC(1899, 11, 30) + day * 86400000).getUTCDate() : day;
  } else {
    const raw = page2Text_(day);
    let match = raw.match(/^\d{4}[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) d = Number(match[2]);
    else {
      match = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
      if (match) d = Number(match[1]);
      else if (/^\d{1,2}$/.test(raw)) d = Number(raw);
      else { const parsed = Date.parse(raw); if (!isNaN(parsed)) d = new Date(parsed).getDate(); }
    }
  }
  return d && d >= 1 && d <= 31 ? String(d).padStart(2,'0') : '';
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
  // Set.has() jauh lebih murah daripada Array.includes() di setiap baris.
  const pics = new Set(page2FilterList_(filters.pics));
  const subdists = new Set(page2FilterList_(filters.subdists));
  const brands = new Set(page2FilterList_(filters.brands));
  const subbrands = new Set(page2FilterList_(filters.subbrands));
  const categories = new Set(page2FilterList_(filters.categories));
  const customers = new Set(page2FilterList_(filters.customers));
  const selectedDates = new Set(page2FilterList_(filters.dates));
  // Bulan kosong bermakna semua bulan, konsisten dengan filter dimensi lain.
  const periodMonths = selectedMonths.length ? selectedMonths : PAGE2_MONTHS_.slice();
  const periodMonthSet = new Set(periodMonths);
  const monthCount = periodMonths.length;
  const hasDateFilter = Array.isArray(filters.dates) && selectedDates.size > 0;
  const matches = r => {
    if (page2Key_(r[7]) !== 'SOO') return false;
    if (pics.size && !pics.has(page2Key_(r[6]))) return false;
    if (subdists.size && !subdists.has(page2Key_(r[8]))) return false;
    if (brands.size && !brands.has(page2Key_(r[4]))) return false;
    if (subbrands.size && !subbrands.has(page2Key_(r[3]))) return false;
    if (categories.size && !categories.has(page2Key_(r[11]))) return false;
    if (customers.size && !customers.has(page2Key_(r[10]))) return false;
    return true;
  };

  const makeBucket = () => ({ao:new Set(),ec:new Set(),groups:{subdist:{},account:{},category:{},subbrand:{},sbu:{},team:{}}});
  const CY = makeBucket(), LY = makeBucket(), CHART_CY = makeBucket(), CHART_LY = makeBucket();
  const monthly = PAGE2_MONTHS_.map(month => ({month:month,ao:0,aoLy:0,ec:0,ecLy:0,value:0,valueLy:0}));
  const monthlySets = PAGE2_MONTHS_.map(() => ({cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()}));
  const dailyMap = {cy:{},ly:{}};
  const orderedSelected = periodMonths.slice().sort((a,b) => PAGE2_MONTHS_.indexOf(a)-PAGE2_MONTHS_.indexOf(b));
  let mtdMonth = selectedMonths.length && orderedSelected.length ? orderedSelected[orderedSelected.length - 1] : null;
  // Jika semua bulan dipilih (checkbox kosong = semua), gunakan bulan terakhir
  // yang benar-benar memiliki data pada tahun dan filter dimensi terpilih.
  if (!mtdMonth) {
    let latestIndex = -1;
    all.forEach(r => {
      if (!matches(r) || Number(r[0]) !== year) return;
      const index = PAGE2_MONTHS_.indexOf(page2Month_(r[1]));
      if (index > latestIndex) latestIndex = index;
    });
    mtdMonth = latestIndex >= 0 ? PAGE2_MONTHS_[latestIndex] : null;
  }
  const mtd = {cyAO:new Set(),lyAO:new Set(),cyEC:new Set(),lyEC:new Set()};
  const paretoByOutlet = new Map();
  const mtdIndex = mtdMonth ? PAGE2_MONTHS_.indexOf(mtdMonth) : -1;
  // Trend Sales selalu kumulatif Januari sampai bulan terakhir yang dipilih.
  // Contoh NOV => JAN-NOV; YTD OCT => JAN-OCT (NOV-DES tidak ditampilkan).
  const trendMonths = mtdIndex >= 0 ? PAGE2_MONTHS_.slice(0,mtdIndex+1) : []; 
  const paretoMonths = new Set(trendMonths);
  let paretoGrandTotalCY = 0;
  const customer = r => page2Key_(r[10] || r[5]);
  // Hanya KODE CUSTOMER valid yang dihitung. Kode placeholder OTHERSB,
  // kode kosong, dan baris tanpa identitas outlet tidak masuk KPI AO maupun EC.
  const outletCode = r => page2Key_(r[13]);
  const validOutletCode = r => {
    const code = outletCode(r);
    // Abaikan placeholder OTHERSB dan hasil error formula (#VALUE!)
    // karena keduanya bukan identitas outlet/transaksi SOO yang valid.
    return !!code && code !== 'OTHERSB' && code !== '#VALUE!';
  };
  // AO dihitung per kombinasi KODE CUSTOMER + SUBDIST agar total KPI
  // konsisten dengan penjumlahan AO masing-masing subdist.
  const outlet = r => validOutletCode(r)
    ? page2CompositeKey_([outletCode(r), r[8]])
    : '';
  // EC = kombinasi TANGGAL (DD) + KODE CUSTOMER valid + SUBDIST.
  // Hari yang sama pada outlet/subdist yang sama dihitung satu EC.
  const ecKeyFor = (r, dayKey) => page2CompositeKey_([
    dayKey, outletCode(r), r[8]
  ]);

  const addGroup = (bucket, dimension, label, outletKey, ecKey, value) => {
    if (!label) return;
    if (!bucket.groups[dimension][label]) {
      bucket.groups[dimension][label] = {name:label,ao:new Set(),ec:new Set(),value:0};
    }
    if (outletKey) bucket.groups[dimension][label].ao.add(outletKey);
    bucket.groups[dimension][label].ec.add(ecKey);
    bucket.groups[dimension][label].value += Number(value) || 0;
  };

  let matchedRowCount = 0;
  all.forEach(r => {
    if (!matches(r)) return;
    matchedRowCount++;

    const y = Number(r[0]), month = page2Month_(r[1]);
    if (y !== year && y !== ly) return;
    // Agregasi bulanan untuk insight YTD/L3M membaca seluruh bulan Jan–Des,
    // tidak terpotong oleh pilihan bulan pada grafik.
    const isCY = y === year;
    const monthIndex = PAGE2_MONTHS_.indexOf(month);
    if (monthIndex >= 0) {
      const monthlyOutletKey = outlet(r);
      const monthlyDay = page2Day_(r[2]);
      const monthlyEcKey = monthlyDay && monthlyOutletKey ? ecKeyFor(r, monthlyDay) : '';
      // Gunakan guard yang sama dengan KPI: kode outlet valid, nama akun,
      // dan tanggal harus tersedia sebelum nilai/EC/AO masuk agregasi.
      if (monthlyOutletKey && customer(r) && monthlyDay) {
        monthly[monthIndex][isCY ? 'value' : 'valueLy'] += Number(r[14]) || 0;
        const ms = monthlySets[monthIndex];
        (isCY ? ms.cyAO : ms.lyAO).add(monthlyOutletKey);
        (isCY ? ms.cyEC : ms.lyEC).add(monthlyEcKey);
      }
    }
    if (paretoMonths.has(month) && validOutletCode(r)) {
      const pKey = outlet(r);
      if (pKey) {
        let item = paretoByOutlet.get(pKey);
        if (!item) { item = {key:pKey,customerName:page2Text_(r[10] || r[5]),subdist:page2Text_(r[8]),cy:{},ly:{}}; paretoByOutlet.set(pKey,item); }
        const val = Number(r[14]) || 0;
        const target = y === year ? item.cy : item.ly;
        target[month] = (target[month] || 0) + val;
        if (y === year && paretoMonths.has(month)) paretoGrandTotalCY += val;
      }
    }
    if (!periodMonthSet.has(month)) return;
    const customerName = customer(r);
    if (!customerName) return;
    const dayKey = page2Day_(r[2]);
    if (!dayKey) return;

    const outletKey = outlet(r);
    // Baris tanpa kode customer valid (termasuk OTHERSB) dikeluarkan
    // dari seluruh agregasi KPI, trend, dan komposisi.
    if (!outletKey) return;
    const ecKey = ecKeyFor(r, dayKey);
    const totalBucket = isCY ? CY : LY;

    if (outletKey) totalBucket.ao.add(outletKey);
    totalBucket.ec.add(ecKey);
    addGroup(totalBucket, 'sbu', page2Key_(r[12]), outletKey, ecKey);
    addGroup(totalBucket, 'team', page2Key_(r[9]) || 'SALESMAN', outletKey, ecKey, Number(r[14]) || 0);

    if (month === mtdMonth) {
      if (outletKey) (isCY ? mtd.cyAO : mtd.lyAO).add(outletKey);
      (isCY ? mtd.cyEC : mtd.lyEC).add(ecKey);
    }

    const chartBucket = isCY ? CHART_CY : CHART_LY;
    addGroup(chartBucket, 'subdist', page2Text_(r[8]), outletKey, ecKey, Number(r[14]) || 0);
    addGroup(chartBucket, 'account', page2Text_(r[10]), outletKey, ecKey, Number(r[14]) || 0);
    addGroup(chartBucket, 'team', page2Key_(r[9]) || 'SALESMAN', outletKey, ecKey, Number(r[14]) || 0);
    addGroup(chartBucket, 'category', page2Text_(r[11]), outletKey, ecKey, Number(r[14]) || 0);
    addGroup(chartBucket, 'subbrand', page2Text_(r[3]), outletKey, ecKey, Number(r[14]) || 0);


    if (!hasDateFilter || selectedDates.has(dayKey)) {
      const dailyKey = month + '|' + dayKey;
      const dailyYear = isCY ? dailyMap.cy : dailyMap.ly;
      if (!dailyYear[dailyKey]) dailyYear[dailyKey] = {ao:new Set(),ec:new Set(),value:0};
      if (outletKey) dailyYear[dailyKey].ao.add(outletKey);
      dailyYear[dailyKey].ec.add(ecKey);
      dailyYear[dailyKey].value += Number(r[14]) || 0;
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
      const value = cy ? cy.value : 0;
      const valueLy = lyGroup ? lyGroup.value : 0;
      return {name:name,ao:ao,aoLy:aoLy,ec:ec,value:value,valueLy:valueLy,valueGrowth:page2Growth_(value,valueLy),growth:page2Growth_(ao,aoLy)};
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
    const cyItems = periodMonths.map(month => dailyMap.cy[month + '|' + day]).filter(Boolean);
    const lyItems = periodMonths.map(month => dailyMap.ly[month + '|' + day]).filter(Boolean);
    const divisor = monthCount || 1;
    return {
      day:day,
      ao:cyItems.reduce((total,item)=>total+item.ao.size,0)/divisor,
      aoLy:lyItems.reduce((total,item)=>total+item.ao.size,0)/divisor,
      ec:cyItems.reduce((total,item)=>total+item.ec.size,0)/divisor,
      ecLy:lyItems.reduce((total,item)=>total+item.ec.size,0)/divisor,
      value:cyItems.reduce((total,item)=>total+(Number(item.value)||0),0)/divisor,
      valueLy:lyItems.reduce((total,item)=>total+(Number(item.value)||0),0)/divisor
    };
  });
  const dailyMonthStats = (source, month) => {
    const items = Object.keys(source).filter(key=>key.startsWith(month+'|')).map(key=>({day:key.split('|')[1],...source[key]}));
    const total = key => items.reduce((sum,item)=>sum+(key==='value'?(Number(item.value)||0):item[key].size),0);
    const best = items.reduce((winner,item)=>{
      const value=Number(item.value)||0;
      return !winner || value>winner.value ? {day:item.day,value:value} : winner;
    },null);
    return {
      ao:total('ao'),ec:total('ec'),value:total('value'),
      days:items.length,bestDay:best?best.day:null,bestValue:best?best.value:0
    };
  };
  const dailyCYMTD = mtdMonth ? dailyMonthStats(dailyMap.cy,mtdMonth) : {ao:0,ec:0,value:0,days:0,bestDay:null,bestValue:0};
  const dailyLYMTD = mtdMonth ? dailyMonthStats(dailyMap.ly,mtdMonth) : {ao:0,ec:0,value:0,days:0,bestDay:null,bestValue:0};

  const ao = CY.ao.size, aoLy = LY.ao.size, ec = CY.ec.size, ecLy = LY.ec.size;
  const aoMTD = mtd.cyAO.size, aoMTDLY = mtd.lyAO.size;
  const ecMTD = mtd.cyEC.size, ecMTDLY = mtd.lyEC.size;

  const composition = dimension => groupResult(dimension, CHART_CY, CHART_LY);
  const paretoExclude = new Set([
    'TIM PROMOSI BINTANG TOEDJOE','BINTANG TOEDJOE','TEAM PROMOSI B7','PT. BINTANG TOEDJOE',
    'END USER TEAM PROMOSI B7 LSE','PT.BINTANG TOEDJOE','TEAM PROMOSI BINTANG TOEDJOE','TEAM PROMOSI BINTANG TOEDJOE / RABITAH',
    'RAJA ALANSYAH MMTW-BAC-MEULABOH','DAVID FIRDAUS MMTW-BAC MEULABOH BARAT',
    'MUHAMMAD RISKY RAMADHANA MMTW BAC MEULABOH TIMUR','ANDRIAN MMTW BAC MEULABOH TIMUR',
    'FIKRI MMTW BAC MEULABOH','VICKY MMTW BAC MEULABOH','DONI KURNIAWAN','ZULFADHLI',
    'MUNTAZAR MMTW - BAC- MEULABOH TIMUR','MUHMMAD ALI MMTW-BAC MEULABOH KOTA',
    'GUNAWAN MMTW BAC MEULABOH','AFRIZAL MMTW BAC MEULABOH','SALMAN MMTW BAC MEULABOH',
    'MUHAMMAD YATIM MMTW BAC MEULABOH','RAHMAD ZULFAHMY MMTW BAC MEULABOH',
    'MMTW BANDA ACEH ACEH KOTA','MMTW LSE LHOKSEUMAWE','MMTW LSE MUARA DUA','MPR LSE LHOKSEUMAWE',
    'MPR LSE BIREUEN','MMTW LSE LANGSA','MPR LSE IDI','MOTORIS KCH SAKA FARMA','MOTORIS KCH LSE'
  ].map(v => page2Key_(v).replace(/[^A-Z0-9]/g,'')));
  const normParetoName = value => page2Key_(value).replace(/[^A-Z0-9]/g,'');
  const previousMonth = mtdIndex > 0 ? PAGE2_MONTHS_[mtdIndex - 1] : null;
  const last3Months = mtdIndex > 0 ? PAGE2_MONTHS_.slice(Math.max(0,mtdIndex-3),mtdIndex) : [];
  const completedMonths = trendMonths.filter(m => PAGE2_MONTHS_.indexOf(m) < mtdIndex);
  const pareto = Array.from(paretoByOutlet.values()).filter(x =>
    x.customerName && !paretoExclude.has(normParetoName(x.customerName))
  ).map(x => {
    const sum = (obj,months) => months.reduce((v,m)=>v+(Number(obj[m])||0),0);
    const actualCY = Number(x.cy[mtdMonth]) || 0;
    const actualLY = Number(x.ly[mtdMonth]) || 0;
    const actualYTD = sum(x.cy,trendMonths);
    const actualLYYTD = sum(x.ly,trendMonths);
    const previous = previousMonth ? Number(x.cy[previousMonth]) || 0 : 0;
    const averageLast3 = last3Months.length ? sum(x.cy,last3Months)/last3Months.length : 0;
    const averageMonth = completedMonths.length ? sum(x.cy,completedMonths)/completedMonths.length : null;
    return {
      customerName:x.customerName,subdist:x.subdist,actualCY,actualLY,
      monthValues:Object.fromEntries(trendMonths.map(month=>[month,Number(x.cy[month])||0])),
      grwMTD:page2Growth_(actualCY,actualLY),actualYTD,actualLYYTD,
      grwYTD:page2Growth_(actualYTD,actualLYYTD),
      grwVsLM:page2Growth_(actualCY,previous),
      grwVsL3M:page2Growth_(actualCY,averageLast3),
      previousValue:previous,averageLast3Value:averageLast3,
      ctr:paretoGrandTotalCY ? actualYTD/paretoGrandTotalCY*100 : 0,
      avgMonth:averageMonth
    };
  }).sort((a,b)=>b.actualYTD-a.actualYTD || a.customerName.localeCompare(b.customerName,'id')).slice(0,25);

  return {
    kpi:{
      year:year,ly:ly,months:periodMonths,trendMonths:trendMonths,monthCount:monthCount,
      ao:ao,aoLy:aoLy,aoGrowth:page2Growth_(ao,aoLy),
      ec:ec,ecLy:ecLy,ecGrowth:page2Growth_(ec,ecLy),
      aoMTD:aoMTD,aoMTDLY:aoMTDLY,ecMTD:ecMTD,ecMTDLY:ecMTDLY,
      mtdMonth:mtdMonth,rows:matchedRowCount,
      dailyCYMTD:dailyCYMTD,dailyLYMTD:dailyLYMTD
    },
    monthly:monthly,
    daily:daily,
    donut:{
      subdist:composition('subdist'),
      team:composition('team'),
      subbrand:composition('subbrand')
    },
    sbu:{b7:groupSbu('B7'),saka:groupSbu('SAKA')},
    byTeam:groupResult('team',CY,LY),
    pareto:pareto,
    updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}

function page2SourceVersion_(sheet) {
  const sh = sheet || page2Source_();
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID || '').trim();
  return [
    'page2-value-matrix-subbrand-v4',
    spreadsheetId, sh.getParent().getId(), sh.getSheetId(), sh.getLastRow(), sh.getLastColumn(),
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
    const meta = page2BuildMeta_(page2Read_(sh));
    if (sh.getName() === 'P2_EVENT_INDEX') {
      const raw = getDatabase_().getSheetByName(RAW_DATA_SHEET_NAME);
      meta.sourceRowCount = raw ? Math.max(0, raw.getLastRow() - 1) : meta.sourceRowCount;
    }
    return meta;
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
function page2CachePutChunked_(cache, key, value, ttlSeconds) {
  const chunkSize = 30000;
  const count = Math.ceil(value.length / chunkSize);
  // Batasi jumlah item agar cache tetap ringan; bila terlalu besar, cache dilewati.
  if (!count || count > 20) return false;
  for (let i = 0; i < count; i++) {
    cache.put(key + '_' + i, value.slice(i * chunkSize, (i + 1) * chunkSize), ttlSeconds);
  }
  cache.put(key + '_manifest', String(count), ttlSeconds);
  return true;
}

function page2CacheGetChunked_(cache, key) {
  const manifest = cache.get(key + '_manifest');
  if (!manifest) return null;
  const count = Number(manifest);
  if (!count || count > 20) return null;
  let value = '';
  for (let i = 0; i < count; i++) {
    const part = cache.get(key + '_' + i);
    if (part == null) return null;
    value += part;
  }
  return value;
}

function getPage2InitialPayload() {
  const totalStarted = Date.now();
  const sh = page2Source_();
  const version = page2SourceVersion_(sh);
  const metaKey = page2CacheKey_('p2meta_schema3', 'initial', version);
  const dataKey = page2CacheKey_('p2initial_schema3', 'initial', version);
  const cache = CacheService.getScriptCache();
  const cachedMeta = page2CacheGetChunked_(cache, metaKey);
  const cachedData = cache.get(dataKey);
  if (cachedMeta && cachedData) {
    try {
      return {
        meta:JSON.parse(cachedMeta),
        data:JSON.parse(cachedData),
        diagnostics:{cacheHit:true,totalMs:Date.now()-totalStarted}
      };
    } catch (err) {
      cache.remove(metaKey + '_manifest');
      cache.remove(dataKey);
    }
  }

  const readStarted = Date.now();
  const sourceRows = page2Read_(sh);
  const readMs = Date.now() - readStarted;

  const metaStarted = Date.now();
  const meta = page2BuildMeta_(sourceRows);
  if (sh.getName() === 'P2_EVENT_INDEX') {
    const raw = getDatabase_().getSheetByName(RAW_DATA_SHEET_NAME);
    meta.sourceRowCount = raw ? Math.max(0, raw.getLastRow() - 1) : meta.sourceRowCount;
  }
  const metaMs = Date.now() - metaStarted;
  const years = meta.years || [];
  const year = Number(years.indexOf('2026') >= 0 ? '2026' : years[years.length - 1]) || new Date().getFullYear();
  // monthsByYear sudah dibangun dalam page2BuildMeta_; tidak perlu scan 132 ribu baris lagi.
  const initialMonths = (meta.monthsByYear || {})[String(year)] || [];

  const aggregateStarted = Date.now();
  const data = page2BuildDashboardData_({year:year,months:initialMonths}, sourceRows);
  const aggregateMs = Date.now() - aggregateStarted;

  const metaJson = JSON.stringify(meta);
  const dataJson = JSON.stringify(data);
  try {
    page2CachePutChunked_(cache, metaKey, metaJson, 300);
    if (dataJson.length <= 90000) cache.put(dataKey, dataJson, 300);
  } catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }

  return {
    meta:meta,
    data:data,
    diagnostics:{
      cacheHit:false,
      sourceRows:sourceRows.length,
      readMs:readMs,
      metaMs:metaMs,
      aggregateMs:aggregateMs,
      totalMs:Date.now()-totalStarted,
      metaBytes:metaJson.length,
      dataBytes:dataJson.length
    }
  };
}

/**
 * Diagnostik manual sebelum deployment.
 * Jalankan dari Apps Script editor untuk membandingkan sumber live dengan
 * angka referensi dashboard tanpa mengubah sheet atau cache.
 */
function testPage2SourceSchema() {
  const sh = page2Source_();
  const rows = page2Read_(sh);
  const meta = page2BuildMeta_(rows);
  const year = Number(meta.years.indexOf('2026') >= 0 ? '2026' : meta.years[meta.years.length - 1]);
  const months = (meta.monthsByYear || {})[String(year)] || [];
  const data = page2BuildDashboardData_({year:year,months:months}, rows);
  return {
    ok:true,
    spreadsheetName:sh.getParent().getName(),
    sheetName:sh.getName(),
    lastRow:sh.getLastRow(),
    lastColumn:sh.getLastColumn(),
    sourceRows:meta.sourceRowCount,
    sooRows:meta.sooRowCount,
    salesTypeCounts:meta.salesTypeCounts,
    years:meta.years,
    selectedYear:year,
    availableMonths:months,
    ao:data.kpi.ao,
    ec:data.kpi.ec,
    aoTrend:data.monthly.filter(x => months.indexOf(x.month) >= 0).map(x => ({month:x.month,ao:x.ao})),
    ecTrend:data.monthly.filter(x => months.indexOf(x.month) >= 0).map(x => ({month:x.month,ec:x.ec})),
    salesmanNamaAkunOptions:meta.salesmen.length,
    updatedAt:data.updatedAt
  };
}


/** Diagnostik Page 2 read-only; tanpa nama/kode customer. */
function testPage2FullDiagnostics() {
  const sh=page2Source_(),rows=page2Read_(sh),meta=page2BuildMeta_(rows);
  const year=Number(meta.years.indexOf('2026')>=0?'2026':meta.years[meta.years.length-1])||new Date().getFullYear();
  const months=(meta.monthsByYear||{})[String(year)]||[];
  const counts={sourceRows:rows.length,sooRows:0,rowsWithYearMonth:0,rowsWithValidCustomerCode:0,rowsWithCustomerName:0,rowsWithSubdist:0,rowsWithValidDate:0,rowsWithPositiveValue:0,years:{},months:{},codeTypes:{},valueTypes:{}};
  rows.forEach(r=>{if(page2Key_(r[7])!=='SOO')return;counts.sooRows++;const y=String(r[0]||'').trim(),m=page2Month_(r[1]);if(y&&m)counts.rowsWithYearMonth++;counts.years[y]=(counts.years[y]||0)+1;if(m)counts.months[m]=(counts.months[m]||0)+1;const code=page2Key_(r[13]),valid=!!code&&code!=='OTHERSB'&&code!=='#VALUE!';if(valid)counts.rowsWithValidCustomerCode++;const type=!code?'BLANK':code==='OTHERSB'?'OTHERSB':code==='#VALUE!'?'#VALUE!':'VALID';counts.codeTypes[type]=(counts.codeTypes[type]||0)+1;if(page2Text_(r[10]||r[5]))counts.rowsWithCustomerName++;if(page2Text_(r[8]))counts.rowsWithSubdist++;if(page2Day_(r[2]))counts.rowsWithValidDate++;if(Number(r[14])>0)counts.rowsWithPositiveValue++;const vt=typeof r[14];counts.valueTypes[vt]=(counts.valueTypes[vt]||0)+1;});
  const data=page2BuildDashboardData_({year:year,months:months},rows);
  return {ok:true,sourceSheet:sh.getName(),sourceSpreadsheet:sh.getParent().getName(),lastRow:sh.getLastRow(),lastColumn:sh.getLastColumn(),selectedYear:year,availableMonths:months,counts:counts,payload:{kpi:data.kpi,donutCounts:{subdist:data.donut.subdist.length,team:data.donut.team.length,subbrand:data.donut.subbrand.length},paretoRows:data.pareto.length,monthlyPositive:data.monthly.filter(x=>x.ao>0||x.ec>0||x.value>0).length},updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED')||'not set'};
}
