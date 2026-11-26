/**
 * MANTUL — SAFE DASHBOARD REFRESH
 * Supports canonical Database V2 while preserving the legacy validation path.
 * Never rebuilds or overwrites any database sheet.
 */
function mantulV2ValidateLiveDatabase_(){
 if(PropertiesService.getScriptProperties().getProperty('MANTUL_V2_DATA_STALE')==='TRUE')throw new Error('Database V2 masih stale setelah import. Jalankan build, finalize, dan publish sebelum refresh.');
 const ss=getDatabase_(),raw=ss.getSheetByName('RAW_DATA_SALES'),
  event=ss.getSheetByName('P2_EVENT_INDEX'),p2=ss.getSheetByName('P2_DASHBOARD_FACT'),
  p3=ss.getSheetByName('P3_OUTLET_MONTHLY'),p1=ss.getSheetByName('P1_FACT');
 if(!p1||p1.getLastRow()<2)throw new Error('P1_FACT belum tersedia atau belum berisi data.');
 if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
 if(!event||event.getLastRow()<2||!p2||p2.getLastRow()<2||!p3||p3.getLastRow()<2)
  throw new Error('Database V2 belum lengkap. Dibutuhkan P2_EVENT_INDEX, P2_DASHBOARD_FACT, dan P3_OUTLET_MONTHLY.');
 mantulV2AssertRawHeaders_(raw);
 const check=(sheet,headers,label)=>{
  if(sheet.getLastColumn()<headers.length)throw new Error(label+' tidak memiliki '+headers.length+' kolom.');
  const actual=sheet.getRange(1,1,1,headers.length).getDisplayValues()[0].map(page2HeaderKey_);
  if(!headers.every((v,i)=>page2HeaderKey_(v)===actual[i]))
   throw new Error('Header '+label+' tidak sesuai schema V2: '+actual.join(' | '));
 };
 check(event,MANTUL_V2_SCHEMA_.P2_EVENT_INDEX,'P2_EVENT_INDEX');
 check(p2,MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT,'P2_DASHBOARD_FACT');
 check(p3,MANTUL_V2_SCHEMA_.P3_OUTLET_MONTHLY,'P3_OUTLET_MONTHLY');
 const rawTotals=mantulV2SheetTotals_(raw,27,13,8),
  eventTotals=mantulV2SheetTotals_(event,19,15,16),
  p2Totals=mantulV2SheetTotals_(p2,17,15,16),
  p3Totals=mantulV2SheetTotals_(p3,18,16,17),
  close=(a,b)=>Math.abs(a-b)<=Math.max(.01,Math.abs(b)*1e-9);
 const checks={eventValue:close(eventTotals.value,rawTotals.value),eventCarton:close(eventTotals.carton,rawTotals.carton),
  p2Value:close(p2Totals.value,rawTotals.value),p2Carton:close(p2Totals.carton,rawTotals.carton),
  p3Value:close(p3Totals.value,rawTotals.value),p3Carton:close(p3Totals.carton,rawTotals.carton)};
 const failed=Object.keys(checks).filter(k=>!checks[k]);
 if(failed.length)throw new Error('Rekonsiliasi Database V2 gagal ('+failed.join(', ')+'). RAW='+JSON.stringify(rawTotals)+
  '; EVENT='+JSON.stringify(eventTotals)+'; P2='+JSON.stringify(p2Totals)+'; P3='+JSON.stringify(p3Totals));
 return {ok:true,version:'V2',spreadsheetId:ss.getId(),sheets:{
  P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
  RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:27},
  P2_EVENT_INDEX:{rows:eventTotals.rows,columns:19},P2_DASHBOARD_FACT:{rows:p2Totals.rows,columns:17},
  P3_OUTLET_MONTHLY:{rows:p3Totals.rows,columns:18}},totals:{raw:rawTotals,event:eventTotals,p2:p2Totals,p3:p3Totals},checks:checks};
}

function refreshMantulDashboard(){
 if(PropertiesService.getScriptProperties().getProperty('MANTUL_V2_DATA_STALE')==='TRUE')throw new Error('Import data terbaru sudah mengubah RAW_DATA_SALES. Selesaikan Database V2 build, finalize, dan publish sebelum refresh.');
 const props=PropertiesService.getScriptProperties(),oldSales=props.getProperty('SALES_LAST_UPDATED')||'',
  oldP2=props.getProperty('PAGE2_RAW_UPDATED')||'',started=Date.now();
 try{
  const ss=getDatabase_(),p2=ss.getSheetByName('P2_EVENT_INDEX');
  let validation;
  if(p2&&p2.getLastColumn()>=19&&page2HeaderKey_(p2.getRange(1,18).getDisplayValue())==='SUB DIST')
   validation=mantulV2ValidateLiveDatabase_();
  else{
   const p1=ss.getSheetByName('P1_FACT'),raw=ss.getSheetByName('RAW_DATA_SALES');
   if(!p1||p1.getLastRow()<2)throw new Error('P1_FACT belum tersedia atau belum berisi data.');
   if(!p2||p2.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
   if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
   analyticsAssertHeaders_(p1,P1_FACT_HEADERS_);analyticsAssertHeaders_(p2,P2_EVENT_HEADERS_);
   const rawCheck=analyticsAssertRawSheet_(raw);
   validation={ok:true,version:'LEGACY',spreadsheetId:ss.getId(),sheets:{
    P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
    P2_EVENT_INDEX:{rows:p2.getLastRow()-1,columns:p2.getLastColumn()},
    RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:rawCheck.columns}}};
   // Legacy registration remains limited to the legacy schemas.
   const audit=analyticsRegisterExistingIndexes();
   validation.derived=audit.derived;
  }
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
  props.setProperty('SALES_LAST_UPDATED',stamp);props.setProperty('PAGE2_RAW_UPDATED',stamp);
  if(typeof cacheDashboardData_==='function')cacheDashboardData_();
  const result={ok:true,status:'COMPLETE',buildExecuted:false,message:'Refresh selesai. Database divalidasi; tidak ada sheet data yang dibangun ulang atau ditimpa.',
   validation:validation,updatedAt:stamp,elapsedMs:Date.now()-started};
  Logger.log('MANTUL SAFE REFRESH\\n'+JSON.stringify(result,null,2));return result;
 }catch(err){
  if(oldSales)props.setProperty('SALES_LAST_UPDATED',oldSales);else props.deleteProperty('SALES_LAST_UPDATED');
  if(oldP2)props.setProperty('PAGE2_RAW_UPDATED',oldP2);else props.deleteProperty('PAGE2_RAW_UPDATED');
  Logger.log('MANTUL SAFE REFRESH FAILED: '+(err&&err.stack||err));throw err;
 }
}

function validateMantulDashboardSources(){
 if(PropertiesService.getScriptProperties().getProperty('MANTUL_V2_DATA_STALE')==='TRUE')throw new Error('Database V2 stale setelah import. Selesaikan build, finalize, dan publish terlebih dahulu.');
 const ss=getDatabase_(),p2=ss.getSheetByName('P2_EVENT_INDEX');
 if(p2&&p2.getLastColumn()>=19&&page2HeaderKey_(p2.getRange(1,18).getDisplayValue())==='SUB DIST')
  return Object.assign({readOnly:true},mantulV2ValidateLiveDatabase_());
 const p1=ss.getSheetByName('P1_FACT'),raw=ss.getSheetByName('RAW_DATA_SALES');
 if(!p1||p1.getLastRow()<2)throw new Error('P1_FACT belum tersedia atau belum berisi data.');
 if(!p2||p2.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
 if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
 analyticsAssertHeaders_(p1,P1_FACT_HEADERS_);analyticsAssertHeaders_(p2,P2_EVENT_HEADERS_);
 const rawCheck=analyticsAssertRawSheet_(raw);
 return {ok:true,readOnly:true,version:'LEGACY',spreadsheetId:ss.getId(),sources:{
  P1_FACT:p1.getLastRow()-1,P2_EVENT_INDEX:p2.getLastRow()-1,RAW_DATA_SALES:raw.getLastRow()-1},rawColumns:rawCheck.columns};
}
