/**
 * HALAMAN 2 — OUTLETS | SALES FORCE
 * Mengutamakan P2_EVENT_INDEX yang sudah tersedia; fallback ke RAW_DATA_SALES.
 * SKU dan MAPPING TEAM tetap menjadi tabel referensi untuk pengembangan berikutnya.
 *
 * Definisi:
 * - AO = distinct KODE CUSTOMER valid pada periode; baris tanpa kode (termasuk OTHERS) tidak dihitung sebagai outlet.
 * - EC = distinct tanggal + identitas outlet + SUBDIST.
 * - Customer picker memfilter KPI, trend, dan komposisi secara konsisten.
 */
const PAGE2_MONTHS_ = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function page2Source_() {
  // Page 2 is bound to the maintained analytics index; never silently fall
  // back to RAW_DATA_SALES because its schema and filter universe differ.
  const ss=getDatabase_();
  const derived=analyticsSummaryReady_()?ss.getSheetByName('P2_DASHBOARD_FACT'):null;
  const sh=derived||ss.getSheetByName('P2_EVENT_INDEX');
  if(!sh||sh.getLastRow()<2)throw new Error('P2_EVENT_INDEX tidak ditemukan atau belum berisi data.');
  const expected=['YEAR','MONTH','DD','CUSTOMER CODE','CUSTOMER NAME','SUBDIST','PIC','BRAND','SUBBRAND','CATEGORY','SBU','SALES TYPE','SALESMAN','TEAM','VAL','CARTON'];
  const actual=sh.getRange(1,1,1,16).getValues()[0].map(page2HeaderKey_);
  if(!expected.every((v,i)=>actual[i]===v))throw new Error('Header P2_EVENT_INDEX harus 16 kolom termasuk CARTON di kolom P. Terbaca: '+actual.join(' | '));
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
  if (sh.getName() === 'P2_EVENT_INDEX' || sh.getName() === 'P2_DASHBOARD_FACT') {
    const expectedHeaders = [
      'YEAR','MONTH','DD','CUSTOMER CODE','CUSTOMER NAME','SUBDIST','PIC',
      'BRAND','SUBBRAND','CATEGORY','SBU','SALES TYPE','SALESMAN','TEAM','VAL','CARTON'
    ];
    const actualHeaders = sh.getRange(1, 1, 1, 16).getValues()[0].map(page2HeaderKey_);
    if (actualHeaders.length !== expectedHeaders.length ||
        !actualHeaders.every((header, index) => header === expectedHeaders[index])) {
      throw new Error(
        'Header P2_EVENT_INDEX tidak sesuai. Diperlukan: ' +
        expectedHeaders.join(', ') + '. Terbaca: ' + actualHeaders.join(', ')
      );
    }
    // Page 2 only consumes A:O. Column P (CARTON) is reserved for Page 3.
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
  const subbrands = new Set(), categories = new Set();
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

// Kunci identitas persis: tidak mengubah kapitalisasi atau spasi internal nama.
function page2ExactCompositeKey_(parts) {
  return parts.map(value => {
    const s = value == null ? '' : String(value);
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
  const mtdIndex = mtdMonth ? PAGE2_MONTHS_.indexOf(mtdMonth) : -1;
  // Trend Sales selalu kumulatif Januari sampai bulan terakhir yang dipilih.
  // Contoh NOV => JAN-NOV; YTD OCT => JAN-OCT (NOV-DES tidak ditampilkan).
  const trendMonths = mtdIndex >= 0 ? PAGE2_MONTHS_.slice(0,mtdIndex+1) : []; 
  const trendMonthSet = new Set(trendMonths);
  const trendCY = {ao:new Set(),ec:new Set(),value:0};
  const trendLY = {ao:new Set(),ec:new Set(),value:0};
  // Identitas outlet mengikuti definisi bisnis: DISTRIBUTOR + CUSTOMER NAME.
  // Nama dipertahankan apa adanya; kode customer tidak menjadi kunci AO/EC.
  const outletKeyFor = (customerName, subdist) => customerName !== ''
    ? page2ExactCompositeKey_([subdist, customerName])
    : '';
  // EC = DISTRIBUTOR + CUSTOMER NAME + tanggal kalender lengkap (YYYY-MM-DD).
  // Tahun dan bulan ikut tercakup dalam dateKey, sehingga tanggal yang sama
  // pada bulan/tahun berbeda tetap menjadi event yang berbeda.

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

    const y = Number(r[0]);
    if (y !== year && y !== ly) return;
    const month = page2Month_(r[1]);
    const isCY = y === year;
    const monthIndex = PAGE2_MONTHS_.indexOf(month);
    const value = Number(r[14]) || 0;

    // Normalisasi identitas dan tanggal cukup sekali per transaksi.
    // Sebelumnya page2Day_, page2Date_, outlet(), dan customer() dipanggil
    // berulang untuk agregasi bulanan, YTD, KPI, dan dailyMap.
    const customerName = String(r[10] == null || r[10] === '' ? (r[5] || '') : r[10]);
    const outletKey = customerName !== '' ? page2ExactCompositeKey_([r[8], customerName]) : '';
    const dayKey = page2Day_(r[2]);
    const dateKey = dayKey && monthIndex >= 0
      ? String(y) + '-' + String(monthIndex + 1).padStart(2,'0') + '-' + dayKey
      : '';
    const ecKey = outletKey && dayKey
      ? page2ExactCompositeKey_([dateKey || dayKey, r[8], customerName])
      : '';

    // Agregasi bulanan untuk insight YTD/L3M membaca seluruh bulan Jan–Des.
    if (monthIndex >= 0 && outletKey && customerName && dayKey) {
      monthly[monthIndex][isCY ? 'value' : 'valueLy'] += value;
      const ms = monthlySets[monthIndex];
      (isCY ? ms.cyAO : ms.lyAO).add(outletKey);
      (isCY ? ms.cyEC : ms.lyEC).add(ecKey);
    }

    // YTD AO/EC memakai identitas unik sepanjang Jan–bulan aktif.
    if (trendMonthSet.has(month) && outletKey && customerName && dayKey) {
      const bucket = isCY ? trendCY : trendLY;
      bucket.ao.add(outletKey);
      bucket.ec.add(ecKey);
      bucket.value += value;
    }

    if (!customerName || !dayKey || !outletKey) return;

    // KPI utama, SBU, dan byTeam harus mengikuti bulan yang dipilih.
    // Jangan memasukkan Desember LY saat periode CY hanya Jan-Nov,
    // dan jangan menghitung full-year ketika pengguna memilih bulan tertentu.
    if (periodMonthSet.has(month)) {
      const totalBucket = isCY ? CY : LY;
      totalBucket.ao.add(outletKey);
      totalBucket.ec.add(ecKey);
      addGroup(totalBucket, 'sbu', page2Key_(r[12]), outletKey, ecKey);
      addGroup(totalBucket, 'team', page2Key_(r[9]) || 'SALESMAN', outletKey, ecKey, Number(r[14]) || 0);

      if (month === mtdMonth) {
        (isCY ? mtd.cyAO : mtd.lyAO).add(outletKey);
        (isCY ? mtd.cyEC : mtd.lyEC).add(ecKey);
      }

    }

    // Donut composition hanya untuk bulan terakhir yang dipilih (bulan berjalan pada filter).
    // KPI utama dan tabel tetap mengikuti seluruh bulan yang dipilih.
    if (month === mtdMonth) {
      const chartBucket = isCY ? CHART_CY : CHART_LY;
      addGroup(chartBucket, 'subdist', page2Text_(r[8]), outletKey, ecKey, value);
      addGroup(chartBucket, 'account', page2Text_(r[10]), outletKey, ecKey, value);
      addGroup(chartBucket, 'team', page2Key_(r[9]) || 'SALESMAN', outletKey, ecKey, value);
      addGroup(chartBucket, 'category', page2Text_(r[11]), outletKey, ecKey, value);
      addGroup(chartBucket, 'subbrand', page2Text_(r[3]), outletKey, ecKey, value);
    }

    // Simpan hanya bulan YTD sampai bulan aktif agar insight harian bisa
    // menghitung MTD, LM, L3M, dan YTD tanpa memuat seluruh histori.
    if (trendMonthSet.has(month) && (!hasDateFilter || selectedDates.has(dayKey))) {
      const dailyKey = month + '|' + dayKey;
      const dailyYear = isCY ? dailyMap.cy : dailyMap.ly;
      if (!dailyYear[dailyKey]) dailyYear[dailyKey] = {ao:new Set(),ec:new Set(),value:0};
      dailyYear[dailyKey].ao.add(outletKey);
      dailyYear[dailyKey].ec.add(ecKey);
      dailyYear[dailyKey].value += value;
    }
    if (!periodMonthSet.has(month)) return;
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
  // Ringkas dailyMap sekali per sumber (CY/LY). Sebelumnya setiap metrik
  // insight memindai Object.keys(dailyMap) berulang kali hingga puluhan kali.
  // Satu pass per tahun menjaga hasil identik sambil mengurangi kerja agregasi.
  const buildDailyMonthTotals = source => {
    const totals = Object.create(null);
    Object.keys(source).forEach(key => {
      const month = key.slice(0,3);
      const item = source[key];
      let bucket = totals[month];
      if (!bucket) {
        bucket = totals[month] = {ao:0,ec:0,value:0,days:0,bestDay:null,bestValue:0};
      }
      const value = Number(item.value) || 0;
      bucket.ao += item.ao.size;
      bucket.ec += item.ec.size;
      bucket.value += value;
      bucket.days++;
      if (bucket.bestDay === null || value > bucket.bestValue) {
        bucket.bestDay = key.slice(4);
        bucket.bestValue = value;
      }
    });
    return totals;
  };
  const dailyMonthlyCYTotals = buildDailyMonthTotals(dailyMap.cy);
  const dailyMonthlyLYTotals = buildDailyMonthTotals(dailyMap.ly);
  const emptyDailyMonth = () => ({ao:0,ec:0,value:0,days:0,bestDay:null,bestValue:0});
  const dailyMonthStats = (source, month) => {
    const totals = source[month];
    return totals ? {
      ao:totals.ao,ec:totals.ec,value:totals.value,
      days:totals.days,bestDay:totals.bestDay,bestValue:totals.bestValue
    } : emptyDailyMonth();
  };
  const dailyCYMTD = mtdMonth ? dailyMonthStats(dailyMonthlyCYTotals,mtdMonth) : emptyDailyMonth();
  const dailyPeriodStats = (source, months) => (months || []).reduce((total,month) => {
    const item = source[month];
    if (!item) return total;
    total.ao += item.ao;
    total.ec += item.ec;
    total.value += item.value;
    total.days += item.days;
    return total;
  }, {ao:0,ec:0,value:0,days:0});
  const previousDailyMonth = mtdIndex > 0 ? PAGE2_MONTHS_[mtdIndex-1] : null;
  const dailyL3MMonths = mtdIndex > 0 ? PAGE2_MONTHS_.slice(Math.max(0,mtdIndex-3),mtdIndex) : [];
  const dailyInsight = {
    cyMtd:dailyPeriodStats(dailyMonthlyCYTotals,mtdMonth?[mtdMonth]:[]),
    cyLm:dailyPeriodStats(dailyMonthlyCYTotals,previousDailyMonth?[previousDailyMonth]:[]),
    cyL3M:dailyPeriodStats(dailyMonthlyCYTotals,dailyL3MMonths),
    cyYtd:dailyPeriodStats(dailyMonthlyCYTotals,trendMonths),
    lyYtd:dailyPeriodStats(dailyMonthlyLYTotals,trendMonths),
    // Bobot bulanan setara; format ini tetap kompatibel dengan renderer.
    monthlyCY:Object.fromEntries(trendMonths.map(month=>[month,dailyPeriodStats(dailyMonthlyCYTotals,[month])])),
    monthlyLY:Object.fromEntries(trendMonths.map(month=>[month,dailyPeriodStats(dailyMonthlyLYTotals,[month])]))
  };
  const dailyLYMTD = mtdMonth ? dailyMonthStats(dailyMonthlyLYTotals,mtdMonth) : emptyDailyMonth();

  const ao = CY.ao.size, aoLy = LY.ao.size, ec = CY.ec.size, ecLy = LY.ec.size;
  const aoMTD = mtd.cyAO.size, aoMTDLY = mtd.lyAO.size;
  const ecMTD = mtd.cyEC.size, ecMTDLY = mtd.lyEC.size;

  const composition = dimension => groupResult(dimension, CHART_CY, CHART_LY);


  return {
    kpi:{
      year:year,ly:ly,months:periodMonths,trendMonths:trendMonths,monthCount:monthCount,
      ao:ao,aoLy:aoLy,aoGrowth:page2Growth_(ao,aoLy),
      ec:ec,ecLy:ecLy,ecGrowth:page2Growth_(ec,ecLy),
      aoMTD:aoMTD,aoMTDLY:aoMTDLY,ecMTD:ecMTD,ecMTDLY:ecMTDLY,
      mtdMonth:mtdMonth,rows:matchedRowCount,
      trendYtdCY:{ao:trendCY.ao.size,ec:trendCY.ec.size,value:trendCY.value},
      trendYtdLY:{ao:trendLY.ao.size,ec:trendLY.ec.size,value:trendLY.value},
      dailyCYMTD:dailyCYMTD,dailyLYMTD:dailyLYMTD,dailyInsight:dailyInsight
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
    updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED') || 'Live dari Google Sheets'
  };
}

function page2ApplySourceStats_(meta){if(!meta)return meta;const s=analyticsGetSourceStats_();if(!s)return meta;meta.sourceRowCount=Number(s.sourceRowCount)||meta.sourceRowCount;meta.sooRowCount=Number(s.salesTypeCounts&&s.salesTypeCounts.SOO)||meta.sooRowCount;meta.salesTypeCounts=s.salesTypeCounts||meta.salesTypeCounts;meta.summaryRowCount=Number(s.p2SummaryRows)||0;return meta;}

function page2SourceVersion_(sheet) {
  const sh = sheet || page2Source_();
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = String(props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID || '').trim();
  return [
    'page2-salesforce-only-v9',
    spreadsheetId, getDatabaseRevision_(), sh.getParent().getId(), sh.getSheetId(), sh.getLastRow(), sh.getLastColumn(),
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
  const cached = cache.get(key) || page2CacheGetChunked_(cache, key);
  if (cached) {
    try { return JSON.parse(cached); }
    catch (err) {
      cache.remove(key);
      cache.remove(key + '_manifest');
    }
  }
  const result = builder();
  const json = JSON.stringify(result);
  try {
    if (json.length <= 90000) cache.put(key, json, ttlSeconds || 300);
    else page2CachePutChunked_(cache, key, json, ttlSeconds || 300);
  } catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }
  return result;
}

function page2GetMeta() {
  const sh = page2Source_();
  const version = page2SourceVersion_(sh);
  const key = page2CacheKey_('p2meta_schema2', 'soo-header-mapped', version);
  return page2Cached_(key, function() {
    const meta = page2BuildMeta_(page2Read_(sh));
    page2ApplySourceStats_(meta);
    if (sh.getName() === 'P2_EVENT_INDEX' || sh.getName() === 'P2_DASHBOARD_FACT') {
      const raw = getDatabase_().getSheetByName(RAW_DATA_SHEET_NAME);
      meta.sourceRowCount = raw ? Math.max(0, raw.getLastRow() - 1) : meta.sourceRowCount;
    }
    return meta;
  }, 300);
}

function page2ReadFiltered_(sheet, filters) {
  const sh=sheet||page2Source_(),safe=filters||{};
  // P2_EVENT_INDEX is already normalized and contains SOO only. Apply the
  // dimension filters before materializing rows for the dashboard aggregation.
  // selected dimensions while rows are being decoded, so the expensive KPI
  // aggregation only receives rows that can appear in the current view.
  if(sh.getName()!=='P2_EVENT_INDEX'&&sh.getName()!=='P2_DASHBOARD_FACT')return page2Read_(sh);
  const lastRow=sh.getLastRow();if(lastRow<=1)return [];
  const expected=['YEAR','MONTH','DD','CUSTOMER CODE','CUSTOMER NAME','SUBDIST','PIC','BRAND','SUBBRAND','CATEGORY','SBU','SALES TYPE','SALESMAN','TEAM','VAL','CARTON'];
  const actual=sh.getRange(1,1,1,16).getValues()[0].map(page2HeaderKey_);
  if(!actual.every((v,i)=>v===expected[i]))throw new Error('Header P2_EVENT_INDEX tidak sesuai. Terbaca: '+actual.join(', '));
  const pics=new Set(page2FilterList_(safe.pics)),subdists=new Set(page2FilterList_(safe.subdists)),
    brands=new Set(page2FilterList_(safe.brands)),subbrands=new Set(page2FilterList_(safe.subbrands)),
    categories=new Set(page2FilterList_(safe.categories));
  const values=sh.getRange(2,1,lastRow-1,15).getValues(),rows=[];
  const analysisYear=Number(safe.year)||0, comparisonYear=analysisYear?analysisYear-1:0;
  for(let i=0;i<values.length;i++){
    const r=values[i];
    const rowYear=Number(r[0]);
    if(analysisYear && rowYear!==analysisYear && rowYear!==comparisonYear)continue;
    if(page2Key_(r[11])!=='SOO')continue;
    if(pics.size&&!pics.has(page2Key_(r[6])))continue;
    if(subdists.size&&!subdists.has(page2Key_(r[5])))continue;
    if(brands.size&&!brands.has(page2Key_(r[7])))continue;
    if(subbrands.size&&!subbrands.has(page2Key_(r[8])))continue;
    if(categories.size&&!categories.has(page2Key_(r[9])))continue;
    rows.push([r[0],page2Month_(r[1]),r[2],r[8],r[7],r[4],r[6],r[11],r[5],
      page2Key_(r[13])||'SALESMAN',r[4],r[9],r[10],r[3],Number(r[14])||0]);
  }
  return rows;
}

function getPage2DashboardData(filters) {
  const started = Date.now();
  const sh = page2Source_();
  const safeFilters = filters || {};
  const version = page2SourceVersion_(sh);
  const key = page2CacheKey_('p2data_schema3', JSON.stringify(safeFilters), version);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key) || page2CacheGetChunked_(cache, key);
  if (cached) {
    try {
      const result = JSON.parse(cached);
      result.diagnostics = {cacheHit:true,totalMs:Date.now()-started};
      return result;
    } catch (err) {
      cache.remove(key);
      cache.remove(key + '_manifest');
    }
  }

  const readStarted = Date.now();
  const sourceRows = page2ReadFiltered_(sh, safeFilters);
  const readMs = Date.now() - readStarted;
  const aggregateStarted = Date.now();
  const result = page2BuildDashboardData_(safeFilters, sourceRows);
  const aggregateMs = Date.now() - aggregateStarted;
  const json = JSON.stringify(result);
  result.diagnostics = {
    cacheHit:false,sourceRows:sourceRows.length,readMs:readMs,
    aggregateMs:aggregateMs,serializeBytes:json.length,totalMs:Date.now()-started
  };
  try {
    if (json.length <= 90000) cache.put(key, json, 300);
    else page2CachePutChunked_(cache, key, json, 300);
  } catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }
  return result;
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
  const parts = [];
  for (let i = 0; i < count; i++) {
    const part = cache.get(key + '_' + i);
    if (part == null) return null;
    parts.push(part);
  }
  return parts.join('');
}

function getPage2InitialPayload() {
  const totalStarted = Date.now();
  const sh = page2Source_();
  const version = page2SourceVersion_(sh);
  const metaKey = page2CacheKey_('p2meta_schema3', 'initial', version);
  const dataKey = page2CacheKey_('p2initial_schema3', 'initial', version);
  const cache = CacheService.getScriptCache();
  const cachedMeta = page2CacheGetChunked_(cache, metaKey);
  const cachedData = cache.get(dataKey) || page2CacheGetChunked_(cache, dataKey);
  if (cachedMeta && cachedData) {
    try {
      return {
        meta:JSON.parse(cachedMeta),
        data:JSON.parse(cachedData),
        diagnostics:{cacheHit:true,source:sh.getName(),totalMs:Date.now()-totalStarted}
      };
    } catch (err) {
      cache.remove(metaKey + '_manifest');
      cache.remove(dataKey + '_manifest');
      cache.remove(dataKey);
    }
  }

  const readStarted = Date.now();
  const sourceRows = page2Read_(sh);
  const readMs = Date.now() - readStarted;

  const metaStarted = Date.now();
  const meta = page2BuildMeta_(sourceRows);
  page2ApplySourceStats_(meta);
  if (sh.getName() === 'P2_EVENT_INDEX' || sh.getName() === 'P2_DASHBOARD_FACT') {
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
    else page2CachePutChunked_(cache, dataKey, dataJson, 300);
  } catch (err) { /* Cache hanya optimasi, bukan dependensi. */ }

  return {
    meta:meta,
    data:data,
    diagnostics:{
      cacheHit:false,
      source:sh.getName(),
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
  const result={ok:true,sourceSheet:sh.getName(),sourceSpreadsheet:sh.getParent().getName(),lastRow:sh.getLastRow(),lastColumn:sh.getLastColumn(),selectedYear:year,availableMonths:months,counts:counts,payload:{kpi:data.kpi,donutCounts:{subdist:data.donut.subdist.length,team:data.donut.team.length,subbrand:data.donut.subbrand.length},paretoRows:data.pareto.length,monthlyPositive:data.monthly.filter(x=>x.ao>0||x.ec>0||x.value>0).length},updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED')||'not set'};
  Logger.log('PAGE2_FULL_DIAGNOSTICS\n' + JSON.stringify(result, null, 2));
  return result;
}

/** Verifikasi AO/EC berdasarkan DISTRIBUTOR + CUSTOMER NAME untuk 2026 Jan-Nov. */
function testPage2CustomerNameCounts() {
  const sh = page2Source_();
  const rows = page2Read_(sh);
  const year = 2026;
  const months = new Set(PAGE2_MONTHS_.slice(0, 11));
  const byDist = Object.create(null);

  rows.forEach(r => {
    if (Number(r[0]) !== year || page2Key_(r[7]) !== 'SOO') return;
    const month = page2Month_(r[1]);
    if (!months.has(month)) return;
    const name = String(r[10] == null || r[10] === '' ? (r[5] || '') : r[10]);
    const dist = String(r[8] == null ? '' : r[8]);
    const day = page2Day_(r[2]);
    if (name === '' || dist === '' || !day) return;
    if (!byDist[dist]) byDist[dist] = {ao:new Set(),ec:new Set()};
    byDist[dist].ao.add(page2ExactCompositeKey_([year, dist, name]));
    const monthIndex = PAGE2_MONTHS_.indexOf(month);
    const dateKey = String(year) + '-' + String(monthIndex + 1).padStart(2,'0') + '-' + day;
    byDist[dist].ec.add(page2ExactCompositeKey_([dist, year, name, month, dateKey]));
  });

  const result = Object.keys(byDist).sort().map(dist => ({
    distributor:dist, ao:byDist[dist].ao.size, ec:byDist[dist].ec.size
  }));
  const totals = result.reduce((sum, item) => ({
    ao:sum.ao + item.ao, ec:sum.ec + item.ec
  }), {ao:0,ec:0});
  const output = {
    ok:true, sourceSheet:sh.getName(), year:year, months:Array.from(months),
    definition:'AO = unique DISTRIBUTOR + YEAR + CUSTOMER NAME; EC = unique DISTRIBUTOR + YEAR + CUSTOMER NAME + MONTH + DATE',
    byDistributor:result, totals:totals
  };
  Logger.log('PAGE2_CUSTOMER_NAME_COUNTS\\n' + JSON.stringify(output, null, 2));
  return output;
}


/* =========================================================
   PAGE 3 — OUTLETS PARETO
   Sumber khusus RAW_DATA_SALES; terpisah dari kalkulasi Page 2.
   ========================================================= */
function page3IndexSheet_(){const ss=getDatabase_(),d=analyticsSummaryReady_()?ss.getSheetByName('P3_OUTLET_MONTHLY'):null,sh=d||ss.getSheetByName('P2_EVENT_INDEX');if(!sh||sh.getLastRow()<2)throw new Error('P2_EVENT_INDEX/P3_OUTLET_MONTHLY tidak ditemukan atau belum berisi data.');if(sh.getName()==='P3_OUTLET_MONTHLY'){const e=P3_OUTLET_HEADERS_.map(page2HeaderKey_),a=sh.getRange(1,1,1,6).getDisplayValues()[0].map(page2HeaderKey_);if(!e.every((v,i)=>a[i]===v))throw new Error('Header P3_OUTLET_MONTHLY tidak sesuai: '+a.join(' | '));return sh;}const e=['YEAR','MONTH','DD','CUSTOMER CODE','CUSTOMER NAME','SUBDIST','PIC','BRAND','SUBBRAND','CATEGORY','SBU','SALES TYPE','SALESMAN','TEAM','VAL','CARTON'],a=sh.getRange(1,1,1,16).getDisplayValues()[0].map(page2HeaderKey_);if(!e.every((v,i)=>a[i]===v))throw new Error('Header P2_EVENT_INDEX tidak sesuai: '+a.join(' | '));return sh;}
function page3SourceVersion_(sh) {
  const props=PropertiesService.getScriptProperties();
  return ['page3-outlet-monthly-v4',page2SourceVersion_(sh),props.getProperty('PAGE2_RAW_UPDATED')||'manual'].join('|');
}
function page3Month_(value) {
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value))return PAGE2_MONTHS_[value.getMonth()];
  const normalized=page2Month_(value);
  return PAGE2_MONTHS_.indexOf(normalized)>=0?normalized:'';
}
function page3Year_(value) {
  const direct=Number(value);if(direct>=1900&&direct<=2200)return direct;
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value))return value.getFullYear();
  const match=String(value==null?'':value).match(/(?:19|20|21)\d{2}/);return match?Number(match[0]):0;
}
/*
 * Pareto hanya memerlukan YEAR, MONTH, CUSTOMER, SUBDIST, SALES TYPE,
 * TEAM, VAL dan CARTON. Baca empat rentang sempit, bukan A:P sekaligus.
 * Simpan bentuk array ringkas [year,month,customer,subdist,value,karton].
 */
function page3ReadIndex_(sh) {
  const sheet=sh||page3IndexSheet_(),lastRow=sheet.getLastRow();
  if(lastRow<2)return {rows:[],diagnostics:{indexRowCount:0,salesTypeCounts:{},teamCounts:{},readMs:0}};
  const n=lastRow-1,started=Date.now();
  if(sheet.getName()==='P3_OUTLET_MONTHLY'){const v=sheet.getRange(2,1,n,6).getValues(),rows=v.map(r=>[page3Year_(r[0]),page3Month_(r[1]),page2Text_(r[2]),page2Text_(r[3]),Number(r[4])||0,Number(r[5])||0]),s=analyticsGetSourceStats_()||{};return {rows:rows,diagnostics:{indexRowCount:Number(s.sourceRowCount)||n,salesTypeCounts:s.salesTypeCounts||{SOO:n},teamCounts:s.teamCounts||{SALESMAN:n},readMs:Date.now()-started,preAggregated:true}};}
  const ab=sheet.getRange(2,1,n,2).getValues();       // A:B YEAR, MONTH
  const df=sheet.getRange(2,4,n,3).getValues();      // D:F CODE, CUSTOMER, SUBDIST
  const sales=sheet.getRange(2,12,n,1).getValues();  // L SALES TYPE
  const np=sheet.getRange(2,14,n,3).getValues();     // N:P TEAM, VAL, CARTON
  const rows=[],salesTypeCounts={},teamCounts={};
  for(let i=0;i<n;i++){
    const type=page2Key_(sales[i][0]),team=page2Key_(np[i][0]);
    if(type)salesTypeCounts[type]=(salesTypeCounts[type]||0)+1;
    if(team)teamCounts[team]=(teamCounts[team]||0)+1;
    if(type!=='SOO'||team!=='SALESMAN')continue;
    const customer=page2Text_(df[i][1]);if(!customer)continue;
    const year=page3Year_(ab[i][0]),month=page3Month_(ab[i][1]);
    if(!year||!month)continue;
    rows.push([year,month,customer,page2Text_(df[i][2]),Number(np[i][1])||0,Number(np[i][2])||0]);
  }
  return {rows:rows,diagnostics:{indexRowCount:n,salesTypeCounts:salesTypeCounts,teamCounts:teamCounts,readMs:Date.now()-started}};
}
function page3RowsCacheKey_(sh) {
  return page2CacheKey_('p3salesman_rows_gzip_v3','compact-soo-salesman',page3SourceVersion_(sh));
}
function page3CacheRows_(cache,key,dataset) {
  try {
    const json=JSON.stringify(dataset);
    const zipped=Utilities.gzip(Utilities.newBlob(json,'application/json'));
    const encoded=Utilities.base64Encode(zipped.getBytes());
    // CacheService item limit: split compressed payload into safe 30 KB chunks.
    return page2CachePutChunked_(cache,key,encoded,300);
  } catch(err) { return false; }
}
function page3GetCachedRows_(cache,key) {
  const encoded=page2CacheGetChunked_(cache,key);if(!encoded)return null;
  try {
    const bytes=Utilities.base64Decode(encoded);
    const json=Utilities.ungzip(Utilities.newBlob(bytes)).getDataAsString();
    return JSON.parse(json);
  } catch(err) {
    cache.remove(key+'_manifest');return null;
  }
}
function page3GetDataset_(sh,cache) {
  const key=page3RowsCacheKey_(sh),cached=page3GetCachedRows_(cache,key);
  if(cached&&Array.isArray(cached.rows))return {dataset:cached,cacheHit:true};
  const dataset=page3ReadIndex_(sh);
  const cacheStored=page3CacheRows_(cache,key,dataset);
  return {dataset:dataset,cacheHit:false,cacheStored:cacheStored};
}
function page3BuildMeta_(dataset) {
  const rows=dataset.rows||[],years=new Set(),subdists=new Set();let monthCount=0;
  for(let i=0;i<rows.length;i++){const r=rows[i];if(r[0])years.add(String(r[0]));if(r[3])subdists.add(r[3]);if(r[1])monthCount++;}
  const diag=dataset.diagnostics||{};
  return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),
    subdists:['EPM BAC','EPM LSE','GMP BAC','GMP LSE','GMP LGS','DB MBO','DB BAC'].filter(x=>subdists.has(x)),
    rowCount:rows.length,rawRowCount:diag.indexRowCount||0,salesTypeCounts:diag.salesTypeCounts||{},
    teamCounts:diag.teamCounts||{},monthCount:monthCount,yearCount:years.size,
    updatedAt:PropertiesService.getScriptProperties().getProperty('PAGE2_RAW_UPDATED')||'P2_EVENT_INDEX'};
}
function page3BuildPareto_(dataset,filters) {
  filters=filters||{};
  const sourceRows=dataset.rows||[];
  let year=Number(filters.year)||0,latestMonthIndex=0;
  if(!year){for(let i=0;i<sourceRows.length;i++)if(sourceRows[i][0]>year)year=sourceRows[i][0];}
  if(!year)year=new Date().getFullYear();
  for(let i=0;i<sourceRows.length;i++){const r=sourceRows[i];if(r[0]===year){const mi=PAGE2_MONTHS_.indexOf(r[1])+1;if(mi>latestMonthIndex)latestMonthIndex=mi;}}
  const months=latestMonthIndex>0?PAGE2_MONTHS_.slice(0,latestMonthIndex):[],monthSet=new Set(months);
  const metric=String(filters.metric||'VALUE').toUpperCase()==='KARTON'?'KARTON':'VALUE';
  const customerQuery=page2Key_(filters.customerQuery||''),subdist=page2Key_(filters.subdist||'SEMUA');
  const map=new Map();let filteredTotal=0;
  for(let i=0;i<sourceRows.length;i++){
    const r=sourceRows[i];if(r[0]!==year||!monthSet.has(r[1]))continue;
    if(customerQuery&&!page2Key_(r[2]).includes(customerQuery))continue;
    if(subdist&&subdist!=='SEMUA'&&page2Key_(r[3])!==subdist)continue;
    const amount=Number(metric==='KARTON'?r[5]:r[4])||0;filteredTotal+=amount;
    const key=page2ExactCompositeKey_([r[3],r[2]]);let item=map.get(key);
    if(!item){item={customerName:r[2],subdist:r[3],monthValues:{},total:0};map.set(key,item);}
    item.monthValues[r[1]]=(item.monthValues[r[1]]||0)+amount;item.total+=amount;
  }
  const rows=Array.from(map.values()).sort((a,b)=>b.total-a.total||a.customerName.localeCompare(b.customerName,'id')).slice(0,25);
  let top25Total=0;for(let i=0;i<rows.length;i++){top25Total+=rows[i].total;rows[i].contribution=filteredTotal?rows[i].total/filteredTotal*100:0;}
  return {year:year,metric:metric,months:months,rows:rows,filteredTotal:filteredTotal,top25Total:top25Total,top25Contribution:filteredTotal?top25Total/filteredTotal*100:0};
}
function getPage3OutletInitialPayload() {
  const started=Date.now(),sh=page3IndexSheet_(),cache=CacheService.getScriptCache();
  const version=page3SourceVersion_(sh),initialKey=page2CacheKey_('p3initial_salesman_v3','initial',version);
  const cached=page2CacheGetChunked_(cache,initialKey);
  if(cached){try{const result=JSON.parse(cached);result.diagnostics={cacheHit:true,source:sh.getName(),totalMs:Date.now()-started};return result;}catch(err){cache.remove(initialKey+'_manifest');}}
  const loaded=page3GetDataset_(sh,cache),dataset=loaded.dataset;
  const meta=page3BuildMeta_(dataset),aggregateStarted=Date.now(),data=page3BuildPareto_(dataset,{});
  const aggregateMs=Date.now()-aggregateStarted,diag=dataset.diagnostics||{};
  const result={meta:meta,data:data,diagnostics:{cacheHit:loaded.cacheHit,source:sh.getName(),sourceRows:meta.rowCount,
    indexRows:meta.rawRowCount,readMs:diag.readMs||0,aggregateMs:aggregateMs,totalMs:Date.now()-started,
    cacheStored:!!loaded.cacheStored}};
  try{page2CachePutChunked_(cache,initialKey,JSON.stringify(result),300);}catch(err){}
  return result;
}
function getPage3OutletData(filters) {
  const started=Date.now(),sh=page3IndexSheet_(),cache=CacheService.getScriptCache(),safe=filters||{};
  const version=page3SourceVersion_(sh),key=page2CacheKey_('p3data_salesman_v3',JSON.stringify(safe),version);
  const cached=cache.get(key);
  if(cached){try{const result=JSON.parse(cached);result.diagnostics={cacheHit:true,totalMs:Date.now()-started};return result;}catch(err){cache.remove(key);}}
  const loaded=page3GetDataset_(sh,cache),aggregateStarted=Date.now();
  const result=page3BuildPareto_(loaded.dataset,safe),aggregateMs=Date.now()-aggregateStarted;
  result.diagnostics={cacheHit:loaded.cacheHit,source:sh.getName(),sourceRows:loaded.dataset.rows.length,
    readMs:loaded.dataset.diagnostics.readMs||0,aggregateMs:aggregateMs,totalMs:Date.now()-started};
  try{const json=JSON.stringify(result);if(json.length<=90000)cache.put(key,json,300);else page2CachePutChunked_(cache,key,json,300);}catch(err){}
  return result;
}
