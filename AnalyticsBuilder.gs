/**
 * MANTUL ANALYTICS — SOURCE INDEX VALIDATOR
 *
 * Spreadsheet ini SUDAH berisi P1_FACT dan P2_EVENT_INDEX sebagai tabel
 * analitik jadi. Jangan membangun ulang/kosongkan keduanya dari RAW_DATA_SALES.
 * RAW_DATA_SALES tetap sumber transaksi; SKU dan MAPPING TEAM adalah referensi.
 *
 * Jalankan analyticsRebuildAll() untuk memvalidasi struktur dan mengaktifkan
 * routing dashboard ke P1_FACT/P2_EVENT_INDEX. Tidak ada data sheet yang diubah.
 */
const ANALYTICS_BUILD_PREFIX_ = 'MANTUL_ANALYTICS_BUILD_';
const P1_FACT_HEADERS_ = ['YEAR','MONTH','SUBDIST','SUBBRAND','BRAND','CATEGORY','SBU','PIC','PF_MONTHLY','PF_QUARTERLY','VAL','SALES_TYPE','TARGET','BE'];
const P2_EVENT_HEADERS_ = ['YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','PIC','BRAND','SUBBRAND','CATEGORY','SBU','SALES_TYPE','SALESMAN','TEAM','VAL','CARTON'];

function analyticsRebuildAll() {
  return analyticsStartInitialBuild();
}

function analyticsStartInitialBuild() {
  const ss = getDatabase_();
  const p1 = ss.getSheetByName('P1_FACT');
  const p2 = ss.getSheetByName('P2_EVENT_INDEX');
  const raw = ss.getSheetByName('RAW_DATA_SALES');
  const sku = ss.getSheetByName('SKU');
  const mapping = ss.getSheetByName('MAPPING TEAM');

  if (!p1 || p1.getLastRow() < 2) throw new Error('P1_FACT tidak ditemukan atau belum berisi data.');
  if (!p2 || p2.getLastRow() < 2) throw new Error('P2_EVENT_INDEX tidak ditemukan atau belum berisi data.');
  if (!raw || raw.getLastRow() < 2) throw new Error('RAW_DATA_SALES tidak ditemukan atau belum berisi data.');

  analyticsAssertHeaders_(p1, P1_FACT_HEADERS_);
  analyticsAssertHeaders_(p2, P2_EVENT_HEADERS_);
  // RAW_DATA_SALES memiliki skema transaksi aktual yang berbeda dari template
  // impor legacy RAW_SALES_HEADERS. Sheet ini hanya diverifikasi keberadaan,
  // isi, dan keberadaan baris header; jangan paksa urutan/header template lama.
  analyticsAssertRawSheet_(raw);
  if (!sku) console.warn('Sheet SKU belum tersedia; fitur Halaman 3 belum dapat memakai referensi SKU.');
  if (!mapping) console.warn('Sheet MAPPING TEAM belum tersedia; mapping team harus diperiksa.');

  const props = PropertiesService.getScriptProperties();
  const prefix = ANALYTICS_BUILD_PREFIX_;
  const now = new Date().toISOString();
  props.setProperties({
    [prefix+'STATUS']:'COMPLETE',
    [prefix+'PHASE']:'COMPLETE',
    [prefix+'SOURCE_ID']:ss.getId(),
    [prefix+'P1_SOURCE_ROWS']:String(p1.getLastRow()-1),
    [prefix+'P2_SOURCE_ROWS']:String(p2.getLastRow()-1),
    [prefix+'P1_WRITTEN']:String(p1.getLastRow()-1),
    [prefix+'P2_WRITTEN']:String(p2.getLastRow()-1),
    [prefix+'P2_SOURCE_UPDATED_AT']:String(props.getProperty('PAGE2_RAW_UPDATED')||''),
    [prefix+'STARTED_AT']:now,
    [prefix+'COMPLETED_AT']:now
  });
  const result = {
    status:'COMPLETE',
    message:'Index yang sudah tersedia berhasil divalidasi. Tidak ada sheet/data yang dibersihkan atau ditulis ulang.',
    sourceSpreadsheetId:ss.getId(),
    spreadsheetUrl:ss.getUrl(),
    sheets:{
      P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
      P2_EVENT_INDEX:{rows:p2.getLastRow()-1,columns:p2.getLastColumn()},
      RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:raw.getLastColumn()},
      SKU:{exists:!!sku,rows:sku?Math.max(0,sku.getLastRow()-1):0},
      'MAPPING TEAM':{exists:!!mapping,rows:mapping?Math.max(0,mapping.getLastRow()-1):0}
    }
  };
  Logger.log(JSON.stringify(result,null,2));
  return result;
}

function analyticsGetBuildStatus() {
  const props=PropertiesService.getScriptProperties(),p=ANALYTICS_BUILD_PREFIX_;
  return {
    status:props.getProperty(p+'STATUS')||'NOT_STARTED',
    phase:props.getProperty(p+'PHASE')||'',
    sourceSpreadsheetId:props.getProperty(p+'SOURCE_ID')||'',
    p1:{sourceRows:Number(props.getProperty(p+'P1_SOURCE_ROWS')||0),writtenRows:Number(props.getProperty(p+'P1_WRITTEN')||0)},
    p2:{sourceRows:Number(props.getProperty(p+'P2_SOURCE_ROWS')||0),writtenRows:Number(props.getProperty(p+'P2_WRITTEN')||0)}
  };
}

function analyticsAssertRawSheet_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error('RAW_DATA_SALES tidak memiliki header pada baris pertama.');
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(v => String(v == null ? '' : v).trim());
  const nonEmpty = headers.filter(Boolean);
  if (!nonEmpty.length) {
    throw new Error('Baris header RAW_DATA_SALES kosong.');
  }
  // Struktur detail RAW_DATA_SALES dipertahankan apa adanya. P1_FACT dan
  // P2_EVENT_INDEX adalah index analitik resmi yang divalidasi secara ketat.
  return { columns:lastColumn, headers:headers };
}

function analyticsAssertHeaders_(sheet, expected) {
  const actual=sheet.getRange(1,1,1,expected.length).getValues()[0]
    .map(v=>String(v==null?'':v).trim().toUpperCase().replace(/[ _-]+/g,' '));
  const wanted=expected.map(v=>String(v).trim().toUpperCase().replace(/[ _-]+/g,' '));
  if(actual.length!==wanted.length||!actual.every((v,i)=>v===wanted[i])) {
    throw new Error('Header sheet '+sheet.getName()+' tidak sesuai. Harus: '+expected.join(' | ')+'. Terbaca: '+actual.join(' | '));
  }
}

function analyticsGetReadySheet_(analyticsSheetName, sourceSheetName, phase) {
  const props=PropertiesService.getScriptProperties(),p=ANALYTICS_BUILD_PREFIX_;
  if(props.getProperty(p+'STATUS')!=='COMPLETE') return null;
  const ss=getDatabase_();
  if(ss.getId()!==String(props.getProperty(p+'SOURCE_ID')||'')) return null;
  // Index sudah merupakan sheet sumber resmi, bukan output rebuild.
  if(analyticsSheetName!==sourceSheetName) return null;
  const sheet=ss.getSheetByName(analyticsSheetName);
  if(!sheet||sheet.getLastRow()<2) return null;
  const expected=Number(props.getProperty(p+phase+'_WRITTEN')||0);
  if(sheet.getLastRow()-1!==expected) return null;
  return sheet;
}

function testAnalyticsDashboardRouting() {
  const ss=getDatabase_();
  const p1=analyticsGetReadySheet_('P1_FACT','P1_FACT','P1');
  const p2=analyticsGetReadySheet_('P2_EVENT_INDEX','P2_EVENT_INDEX','P2');
  const result={
    buildStatus:PropertiesService.getScriptProperties().getProperty(ANALYTICS_BUILD_PREFIX_+'STATUS')||'NOT_STARTED',
    sourceSpreadsheetId:ss.getId(),
    p1:{route:p1?'P1_FACT':'not-ready',rows:p1? p1.getLastRow()-1:0},
    p2:{route:p2?'P2_EVENT_INDEX':'not-ready',rows:p2? p2.getLastRow()-1:0}
  };
  Logger.log(JSON.stringify(result,null,2));
  return result;
}
