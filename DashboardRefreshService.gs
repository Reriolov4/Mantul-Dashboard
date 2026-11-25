/**
 * MANTUL — SAFE DASHBOARD REFRESH
 * Validates existing source and derived sheets. It never runs analyticsBuildDashboardIndexes().
 */
function refreshMantulDashboard(){
 const props=PropertiesService.getScriptProperties(),oldSales=props.getProperty('SALES_LAST_UPDATED')||'',oldP2=props.getProperty('PAGE2_RAW_UPDATED')||'',started=Date.now();
 try{
  const ss=getDatabase_(),p1=ss.getSheetByName('P1_FACT'),p2=ss.getSheetByName('P2_EVENT_INDEX'),raw=ss.getSheetByName('RAW_DATA_SALES');
  if(!p1||p1.getLastRow()<2)throw new Error('P1_FACT belum tersedia atau belum berisi data.');
  if(!p2||p2.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
  if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
  analyticsAssertHeaders_(p1,P1_FACT_HEADERS_);analyticsAssertHeaders_(p2,P2_EVENT_HEADERS_);
  const rawCheck=analyticsAssertRawSheet_(raw),audit=analyticsRegisterExistingIndexes();
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
  props.setProperty('SALES_LAST_UPDATED',stamp);props.setProperty('PAGE2_RAW_UPDATED',stamp);
  if(typeof cacheDashboardData_==='function')cacheDashboardData_();
  const result={ok:true,status:'COMPLETE',buildExecuted:false,message:'Refresh selesai. Source dan derived sheets dipertahankan; hanya status/cache yang diperbarui.',
   spreadsheetId:ss.getId(),sources:{P1_FACT:{rows:p1.getLastRow()-1,columns:p1.getLastColumn()},
   P2_EVENT_INDEX:{rows:p2.getLastRow()-1,columns:p2.getLastColumn()},RAW_DATA_SALES:{rows:raw.getLastRow()-1,columns:rawCheck.columns}},
   derived:audit.derived,indexAudit:{ok:audit.ok,auditOnly:audit.auditOnly,sheetsMutated:audit.sheetsMutated},
   updatedAt:stamp,elapsedMs:Date.now()-started};
  Logger.log('MANTUL SAFE REFRESH\n'+JSON.stringify(result,null,2));return result;
 }catch(err){
  if(oldSales)props.setProperty('SALES_LAST_UPDATED',oldSales);else props.deleteProperty('SALES_LAST_UPDATED');
  if(oldP2)props.setProperty('PAGE2_RAW_UPDATED',oldP2);else props.deleteProperty('PAGE2_RAW_UPDATED');
  Logger.log('MANTUL SAFE REFRESH FAILED: '+(err&&err.stack||err));throw err;
 }
}
function validateMantulDashboardSources(){
 const ss=getDatabase_(),p1=ss.getSheetByName('P1_FACT'),p2=ss.getSheetByName('P2_EVENT_INDEX'),raw=ss.getSheetByName('RAW_DATA_SALES');
 if(!p1||p1.getLastRow()<2)throw new Error('P1_FACT belum tersedia atau belum berisi data.');
 if(!p2||p2.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum tersedia atau belum berisi data.');
 if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum tersedia atau belum berisi data.');
 analyticsAssertHeaders_(p1,P1_FACT_HEADERS_);analyticsAssertHeaders_(p2,P2_EVENT_HEADERS_);
 const rawCheck=analyticsAssertRawSheet_(raw);
 return {ok:true,readOnly:true,spreadsheetId:ss.getId(),sources:{P1_FACT:p1.getLastRow()-1,P2_EVENT_INDEX:p2.getLastRow()-1,RAW_DATA_SALES:raw.getLastRow()-1},rawColumns:rawCheck.columns};
}