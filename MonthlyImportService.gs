/**
 * MANTUL — MONTHLY PARTITION IMPORT
 *
 * Import workbook yang berisi P1_FACT, P2_EVENT_INDEX, dan RAW_DATA_SALES.
 * Browser hanya mengirim baris untuk satu YEAR+MONTH; sheet live diganti
 * per periode (bukan append), sehingga upload ulang bulan yang sama idempotent.
 *
 * SKU dan MAPPING TEAM tidak disentuh oleh proses bulanan.
 */
const MONTHLY_IMPORT_PREFIX_ = 'MANTUL_MONTHLY_IMPORT_';
const MONTHLY_RAW_HEADERS_ = [
  'TAHUN','BULAN','DATE','KODE ITEM','NAMA ITEM','SUB BRAND','BRAND','KARTON',
  'DISTRIBUTORS','KODE CUSTOMER','CUSTOMER NAME','KODE SALUR','VAL','QTY',
  'JENIS SALES','SUBDIST','PIC','TIM','NAMA AKUN','CATEGORY','DOT','SBU',
  'PF Bulanan','PF Triwulan','SKU KCH'
];
const MONTHLY_STAGE_NAMES_ = {
  p1:'P1_MONTH_STAGING',
  p2:'P2_MONTH_STAGING',
  raw:'RAW_MONTH_STAGING'
};
const MONTHLY_BACKUP_NAMES_ = {
  p1:'P1_MONTH_BACKUP',
  p2:'P2_MONTH_BACKUP',
  raw:'RAW_MONTH_BACKUP'
};

function monthlyImportMonth_(value) {
  const s = String(value == null ? '' : value).trim().toUpperCase();
  const aliases = {
    JANUARI:'JAN',FEBRUARI:'FEB',MARET:'MAR',APRIL:'APR',MEI:'MAY',
    JUNI:'JUN',JULI:'JUL',AGUSTUS:'AUG',SEPTEMBER:'SEP',
    OKTOBER:'OCT',NOVEMBER:'NOV',DESEMBER:'DEC',
    JANUARY:'JAN',FEBRUARY:'FEB',MARCH:'MAR',JUNE:'JUN',
    JULY:'JUL',AUGUST:'AUG',OCTOBER:'OCT',DECEMBER:'DEC'
  };
  if (aliases[s]) return aliases[s];
  if (/^\d{1,2}$/.test(s)) {
    const n = Number(s);
    return n >= 1 && n <= 12 ? PAGE2_MONTHS_[n - 1] : '';
  }
  return s.slice(0,3);
}

function monthlyImportAssertHeaders_(sheet, expected) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < expected.length) {
    throw new Error('Sheet ' + (sheet ? sheet.getName() : '(tidak ada)') + ' tidak memiliki kolom yang diperlukan.');
  }
  const actual = sheet.getRange(1,1,1,expected.length).getDisplayValues()[0]
    .map(v => String(v == null ? '' : v).trim().toUpperCase().replace(/[_-]+/g,' ').replace(/\s+/g,' '));
  const wanted = expected.map(v => String(v).trim().toUpperCase().replace(/[_-]+/g,' ').replace(/\s+/g,' '));
  if (!actual.every((v,i) => v === wanted[i])) {
    throw new Error('Header ' + sheet.getName() + ' tidak sesuai. Terbaca: ' + actual.join(' | '));
  }
}

function beginMonthlyDataImport(meta) {
  assertUploadPermission_();
  meta = meta || {};
  const year = Number(meta.year);
  const month = monthlyImportMonth_(meta.month);
  const p1Total = Number(meta.p1Total);
  const p2Total = Number(meta.p2Total);
  const rawTotal = Number(meta.rawTotal);
  const allowedMonths = PAGE2_MONTHS_;

  if (!Number.isInteger(year) || year < 2000 || year > 2100 ||
      allowedMonths.indexOf(month) < 0 ||
      !Number.isInteger(p1Total) || p1Total < 1 ||
      !Number.isInteger(p2Total) || p2Total < 1 ||
      !Number.isInteger(rawTotal) || rawTotal < 1 ||
      p2Total !== rawTotal) {
    throw new Error('Periode atau jumlah baris bulanan tidak valid. P2_EVENT_INDEX dan RAW_DATA_SALES harus memiliki jumlah baris yang sama.');
  }
  if (p1Total > 5000 || p2Total > 30000 || rawTotal > 30000) {
    throw new Error('Jumlah baris periode melebihi batas aman import bulanan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const prefix = MONTHLY_IMPORT_PREFIX_;
    const existingToken = props.getProperty(prefix+'TOKEN');
    const startedAt = Number(props.getProperty(prefix+'STARTED') || 0);
    if (existingToken && startedAt && Date.now()-startedAt < 2*60*60*1000) {
      throw new Error('Import bulanan lain masih berlangsung. Selesaikan atau tunggu sesi kedaluwarsa.');
    }

    const ss = getDatabase_();
    const p1 = ss.getSheetByName('P1_FACT');
    const p2 = ss.getSheetByName('P2_EVENT_INDEX');
    const raw = ss.getSheetByName('RAW_DATA_SALES');
    if (!p1 || !p2 || !raw) throw new Error('P1_FACT, P2_EVENT_INDEX, dan RAW_DATA_SALES wajib tersedia.');
    analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
    analyticsAssertHeaders_(p2, P2_EVENT_HEADERS_);
    monthlyImportAssertHeaders_(raw, MONTHLY_RAW_HEADERS_);

    const stageSpecs = [
      ['p1',P1_FACT_HEADERS_],['p2',P2_EVENT_HEADERS_],['raw',MONTHLY_RAW_HEADERS_]
    ];
    stageSpecs.forEach(([key,headers]) => {
      const stage = getOrCreateSheet_(MONTHLY_STAGE_NAMES_[key]);
      stage.clearContents();
      ensureSheetGrid_(stage,1,headers.length);
      stage.getRange(1,1,1,headers.length).setValues([headers]);
    });

    const token = Utilities.getUuid();
    props.setProperties({
      [prefix+'TOKEN']:token,
      [prefix+'YEAR']:String(year),
      [prefix+'MONTH']:month,
      [prefix+'P1_EXPECTED']:String(p1Total),
      [prefix+'P2_EXPECTED']:String(p2Total),
      [prefix+'RAW_EXPECTED']:String(rawTotal),
      [prefix+'P1_WRITTEN']:'0',
      [prefix+'P2_WRITTEN']:'0',
      [prefix+'RAW_WRITTEN']:'0',
      [prefix+'STARTED']:String(Date.now()),
      [prefix+'FILENAME']:safeText_(meta.filename || 'Workbook sales',200)
    });
    return {ok:true,token:token,year:year,month:month,p1Expected:p1Total,p2Expected:p2Total,rawExpected:rawTotal};
  } finally {
    lock.releaseLock();
  }
}

function monthlyImportAssertSession_(props, token) {
  const prefix = MONTHLY_IMPORT_PREFIX_;
  const saved = props.getProperty(prefix+'TOKEN');
  const started = Number(props.getProperty(prefix+'STARTED') || 0);
  if (!saved || !token || saved !== token) throw new Error('Sesi import bulanan tidak valid atau sudah diganti.');
  if (!started || Date.now()-started >= 2*60*60*1000) {
    monthlyImportClearSession_(props);
    throw new Error('Sesi import bulanan kedaluwarsa. Silakan upload ulang.');
  }
}

function monthlyImportClearSession_(props) {
  const prefix = MONTHLY_IMPORT_PREFIX_;
  ['TOKEN','YEAR','MONTH','P1_EXPECTED','P2_EXPECTED','RAW_EXPECTED',
   'P1_WRITTEN','P2_WRITTEN','RAW_WRITTEN','STARTED','FILENAME']
    .forEach(key => props.deleteProperty(prefix+key));
}

function monthlyImportValidateChunk_(rows, width, year, month, sheetLabel, yearIndex, monthIndex) {
  (rows || []).forEach((row,i) => {
    if (!Array.isArray(row) || row.length !== width) {
      throw new Error(sheetLabel + ' chunk baris ' + (i+1) + ' harus memiliki ' + width + ' kolom.');
    }
    if (Number(row[yearIndex]) !== year || monthlyImportMonth_(row[monthIndex]) !== month) {
      throw new Error(sheetLabel + ' berisi baris di luar periode ' + year + '-' + month + '.');
    }
  });
}

function appendMonthlyDataImportChunk(payload) {
  assertUploadPermission_();
  payload = payload || {};
  const p1Chunk = Array.isArray(payload.p1Chunk) ? payload.p1Chunk : [];
  const p2Chunk = Array.isArray(payload.p2Chunk) ? payload.p2Chunk : [];
  const rawChunk = Array.isArray(payload.rawChunk) ? payload.rawChunk : [];
  if (!p1Chunk.length && !p2Chunk.length && !rawChunk.length) return {written:0};
  if (p1Chunk.length>1000 || p2Chunk.length>1000 || rawChunk.length>1000) {
    throw new Error('Maksimal 1.000 baris per sheet per batch.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    monthlyImportAssertSession_(props,payload.token);
    const prefix = MONTHLY_IMPORT_PREFIX_;
    const year = Number(props.getProperty(prefix+'YEAR'));
    const month = props.getProperty(prefix+'MONTH');
    const specs = [
      ['p1',p1Chunk,14,0,1,'P1_FACT'],
      ['p2',p2Chunk,15,0,1,'P2_EVENT_INDEX'],
      ['raw',rawChunk,25,0,1,'RAW_DATA_SALES']
    ];
    specs.forEach(([key,rows,width,yearIndex,monthIndex,label]) => {
      if (!rows.length) return;
      const expected = Number(props.getProperty(prefix+key.toUpperCase()+'_EXPECTED') || 0);
      const written = Number(props.getProperty(prefix+key.toUpperCase()+'_WRITTEN') || 0);
      if (written + rows.length > expected) throw new Error(label + ' melebihi jumlah baris yang didaftarkan.');
      monthlyImportValidateChunk_(rows,width,year,month,label,yearIndex,monthIndex);
      const stage = getOrCreateSheet_(MONTHLY_STAGE_NAMES_[key]);
      ensureSheetGrid_(stage,written+rows.length+1,width);
      stage.getRange(written+2,1,rows.length,width).setValues(rows);
      props.setProperty(prefix+key.toUpperCase()+'_WRITTEN',String(written+rows.length));
    });
    return {
      p1Written:Number(props.getProperty(prefix+'P1_WRITTEN')||0),
      p1Expected:Number(props.getProperty(prefix+'P1_EXPECTED')||0),
      p2Written:Number(props.getProperty(prefix+'P2_WRITTEN')||0),
      p2Expected:Number(props.getProperty(prefix+'P2_EXPECTED')||0),
      rawWritten:Number(props.getProperty(prefix+'RAW_WRITTEN')||0),
      rawExpected:Number(props.getProperty(prefix+'RAW_EXPECTED')||0)
    };
  } finally {
    lock.releaseLock();
  }
}

function monthlyImportReadStage_(key, expected, width) {
  const sheet = getDatabase_().getSheetByName(MONTHLY_STAGE_NAMES_[key]);
  if (!sheet || sheet.getLastRow() !== expected+1) {
    throw new Error('Staging ' + key + ' tidak lengkap. Diharapkan ' + expected + ' baris.');
  }
  return expected ? sheet.getRange(2,1,expected,width).getValues() : [];
}

function monthlyImportSummary_(rows, distributorIndex, salesTypeIndex, valueIndex) {
  const summary = Object.create(null);
  rows.forEach(row => {
    const distributor = String(row[distributorIndex] == null ? '' : row[distributorIndex]).trim().toUpperCase();
    const salesType = String(row[salesTypeIndex] == null ? '' : row[salesTypeIndex]).trim().toUpperCase();
    const key = distributor + '|' + salesType;
    if (!summary[key]) summary[key] = {count:0,value:0};
    summary[key].count++;
    summary[key].value += Number(row[valueIndex]) || 0;
  });
  return summary;
}

function monthlyImportCompareP2Raw_(p2Rows, rawRows) {
  const p2 = monthlyImportSummary_(p2Rows,5,11,14);
  const raw = monthlyImportSummary_(rawRows,8,14,12);
  const keys = Object.keys(p2).concat(Object.keys(raw).filter(k => !p2[k]));
  for (let i=0;i<keys.length;i++) {
    const key = keys[i], a = p2[key] || {count:0,value:0}, b = raw[key] || {count:0,value:0};
    if (a.count !== b.count || Math.abs(a.value-b.value) > 0.01) {
      throw new Error('P2_EVENT_INDEX dan RAW_DATA_SALES tidak cocok pada ' + key +
        '. P2=' + a.count + ' baris/' + a.value + '; RAW=' + b.count + ' baris/' + b.value + '.');
    }
  }
}

function monthlyImportFindBlocks_(sheet, year, month) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const keys = sheet.getRange(2,1,lastRow-1,2).getValues();
  const blocks = [];
  let start = -1, count = 0;
  for (let i=0;i<keys.length;i++) {
    const match = Number(keys[i][0]) === year && monthlyImportMonth_(keys[i][1]) === month;
    if (match) {
      if (start < 0) start = i+2;
      count++;
    } else if (start >= 0) {
      blocks.push([start,count]);
      start = -1;
      count = 0;
    }
  }
  if (start >= 0) blocks.push([start,count]);
  return blocks;
}

function monthlyImportReadPeriodRows_(sheet, year, month) {
  const width = sheet.getLastColumn();
  const blocks = monthlyImportFindBlocks_(sheet,year,month);
  const rows = [];
  blocks.forEach(([start,count]) => {
    const part = sheet.getRange(start,1,count,width).getValues();
    Array.prototype.push.apply(rows,part);
  });
  return {blocks:blocks,rows:rows,width:width};
}

function monthlyImportWriteBackup_(key, headers, rows) {
  const backup = getOrCreateSheet_(MONTHLY_BACKUP_NAMES_[key]);
  backup.clearContents();
  ensureSheetGrid_(backup,rows.length+1,headers.length);
  backup.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) backup.getRange(2,1,rows.length,headers.length).setValues(rows);
  return rows;
}

function monthlyImportDeletePeriod_(sheet, year, month) {
  const blocks = monthlyImportFindBlocks_(sheet,year,month).sort((a,b)=>b[0]-a[0]);
  let removed = 0;
  blocks.forEach(([start,count]) => {
    sheet.deleteRows(start,count);
    removed += count;
  });
  return removed;
}

function monthlyImportAppendRows_(sheet, rows, width) {
  if (!rows.length) return;
  const start = sheet.getLastRow()+1;
  ensureSheetGrid_(sheet,start+rows.length-1,width);
  sheet.getRange(start,1,rows.length,width).setValues(rows);
}

function finishMonthlyDataImport(payload) {
  assertUploadPermission_();
  payload = payload || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    monthlyImportAssertSession_(props,payload.token);
    const prefix = MONTHLY_IMPORT_PREFIX_;
    const year = Number(props.getProperty(prefix+'YEAR'));
    const month = props.getProperty(prefix+'MONTH');
    const p1Expected = Number(props.getProperty(prefix+'P1_EXPECTED')||0);
    const p2Expected = Number(props.getProperty(prefix+'P2_EXPECTED')||0);
    const rawExpected = Number(props.getProperty(prefix+'RAW_EXPECTED')||0);
    if (Number(props.getProperty(prefix+'P1_WRITTEN')||0)!==p1Expected ||
        Number(props.getProperty(prefix+'P2_WRITTEN')||0)!==p2Expected ||
        Number(props.getProperty(prefix+'RAW_WRITTEN')||0)!==rawExpected) {
      throw new Error('Import belum lengkap. P1/P2/RAW belum mencapai jumlah baris yang didaftarkan.');
    }

    const p1Rows = monthlyImportReadStage_('p1',p1Expected,14);
    const p2Rows = monthlyImportReadStage_('p2',p2Expected,15);
    const rawRows = monthlyImportReadStage_('raw',rawExpected,25);
    monthlyImportValidateChunk_(p1Rows,14,year,month,'P1_FACT',0,1);
    monthlyImportValidateChunk_(p2Rows,15,year,month,'P2_EVENT_INDEX',0,1);
    monthlyImportValidateChunk_(rawRows,25,year,month,'RAW_DATA_SALES',0,1);
    monthlyImportCompareP2Raw_(p2Rows,rawRows);

    const p1Soo = p1Rows.filter(r=>String(r[11]||'').trim().toUpperCase()==='SOO')
      .reduce((sum,r)=>sum+(Number(r[10])||0),0);
    const p2Soo = p2Rows.filter(r=>String(r[11]||'').trim().toUpperCase()==='SOO')
      .reduce((sum,r)=>sum+(Number(r[14])||0),0);
    const p1SooDiff = p1Soo-p2Soo;

    const ss = getDatabase_();
    const live = {
      p1:ss.getSheetByName('P1_FACT'),
      p2:ss.getSheetByName('P2_EVENT_INDEX'),
      raw:ss.getSheetByName('RAW_DATA_SALES')
    };
    const headers = {p1:P1_FACT_HEADERS_,p2:P2_EVENT_HEADERS_,raw:MONTHLY_RAW_HEADERS_};
    const widths = {p1:14,p2:15,raw:25};
    const newRows = {p1:p1Rows,p2:p2Rows,raw:rawRows};
    const oldRows = {};
    Object.keys(live).forEach(key => {
      if (!live[key]) throw new Error('Sheet live ' + key + ' tidak ditemukan.');
      const captured = monthlyImportReadPeriodRows_(live[key],year,month);
      oldRows[key] = captured.rows;
      monthlyImportWriteBackup_(key,headers[key],captured.rows);
    });

    const oldSalesUpdated = props.getProperty('SALES_LAST_UPDATED') || '';
    const oldPage2Updated = props.getProperty('PAGE2_RAW_UPDATED') || '';
    try {
      Object.keys(live).forEach(key => monthlyImportDeletePeriod_(live[key],year,month));
      Object.keys(live).forEach(key => monthlyImportAppendRows_(live[key],newRows[key],widths[key]));

      const now = new Date();
      const formatted = Utilities.formatDate(now,Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
      props.setProperty('SALES_LAST_UPDATED',formatted);
      props.setProperty('PAGE2_RAW_UPDATED',formatted);
      analyticsStartInitialBuild();
      cacheDashboardData_();

      const log = getOrCreateSheet_('MONTHLY_IMPORT_LOG');
      if (log.getLastRow()===0) {
        log.appendRow(['WAKTU','FILE','PERIODE','P1_FACT ROWS','P2_EVENT_INDEX ROWS','RAW_DATA_SALES ROWS','P1 SOO','P2 SOO','SELISIH SOO','UPLOADER','STATUS']);
      }
      log.appendRow([now,props.getProperty(prefix+'FILENAME')||'',year+'-'+month,
        p1Rows.length,p2Rows.length,rawRows.length,p1Soo,p2Soo,p1SooDiff,
        Session.getActiveUser().getEmail()||'email tidak tersedia','SUCCESS']);

      Object.keys(MONTHLY_STAGE_NAMES_).forEach(key => {
        const stage = ss.getSheetByName(MONTHLY_STAGE_NAMES_[key]);
        if (stage) stage.clearContents();
      });
      monthlyImportClearSession_(props);
      return {
        ok:true,year:year,month:month,
        p1Rows:p1Rows.length,p2Rows:p2Rows.length,rawRows:rawRows.length,
        p1Soo:p1Soo,p2Soo:p2Soo,p1SooDiff:p1SooDiff,updatedAt:formatted,
        message:'Periode berhasil diganti. Baris lama untuk periode yang sama tidak digandakan.'
      };
    } catch (publishError) {
      let rollbackError = null;
      try {
        Object.keys(live).forEach(key => monthlyImportDeletePeriod_(live[key],year,month));
        Object.keys(live).forEach(key => monthlyImportAppendRows_(live[key],oldRows[key],widths[key]));
        props.setProperty('SALES_LAST_UPDATED',oldSalesUpdated);
        props.setProperty('PAGE2_RAW_UPDATED',oldPage2Updated);
        analyticsStartInitialBuild();
        cacheDashboardData_();
      } catch (err) {
        rollbackError = err;
      }
      if (rollbackError) {
        throw new Error('Import gagal: '+publishError.message+'. Rollback juga gagal: '+rollbackError.message+'. Periksa backup P1_MONTH_BACKUP, P2_MONTH_BACKUP, dan RAW_MONTH_BACKUP.');
      }
      throw new Error('Import gagal dan data periode sebelumnya sudah dipulihkan: '+publishError.message);
    }
  } finally {
    lock.releaseLock();
  }
}
