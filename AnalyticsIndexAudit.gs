/**
 * MANTUL — SAFE DERIVED INDEX AUDIT
 * Never rebuilds, clears, renames, sorts, or deletes spreadsheet sheets.
 * Registration writes only Script Properties after existing aggregates reconcile.
 */
function analyticsRegisterExistingIndexes(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=getDatabase_(),src=ss.getSheetByName('P2_EVENT_INDEX'),p2=ss.getSheetByName('P2_DASHBOARD_FACT'),
   sub=ss.getSheetByName('P2_SUBDIST_FACT'),p3=ss.getSheetByName('P3_OUTLET_MONTHLY');
  if(!src||src.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum berisi data.');
  if(!p2||p2.getLastRow()<2||!sub||sub.getLastRow()<2||!p3||p3.getLastRow()<2)
   throw new Error('Derived sheet belum lengkap. Audit berhenti tanpa mengubah sheet.');
  analyticsAssertHeaders_(src,P2_EVENT_HEADERS_);analyticsAssertHeaders_(p2,P2_DASHBOARD_HEADERS_);
  analyticsAssertHeaders_(sub,P2_DASHBOARD_HEADERS_);analyticsAssertHeaders_(p3,P3_OUTLET_HEADERS_);
  const sourceCount=src.getLastRow()-1,meta=src.getRange(2,12,sourceCount,3).getValues(),vals=src.getRange(2,15,sourceCount,2).getValues();
  let sooRows=0,sooValue=0,sooCarton=0;const salesTypeCounts={},teamCounts={};
  for(let i=0;i<sourceCount;i++){const type=page2Key_(meta[i][0]),team=page2Key_(meta[i][2]);if(type)salesTypeCounts[type]=(salesTypeCounts[type]||0)+1;if(team)teamCounts[team]=(teamCounts[team]||0)+1;
   if(type!=='SOO')continue;sooRows++;sooValue+=Number(vals[i][0])||0;sooCarton+=Number(vals[i][1])||0;
  }
  const readTotals=(sh,sourceRows,start)=>{const n=sh.getLastRow()-1,a=sh.getRange(2,start,n,sourceRows?3:2).getValues();
   return a.reduce((t,r)=>{t.value+=Number(r[0])||0;t.carton+=Number(r[1])||0;if(sourceRows)t.sourceRows+=Number(r[2])||0;return t;},{value:0,carton:0,sourceRows:0});};
  const a=readTotals(p2,true,15),b=readTotals(sub,true,15),c=readTotals(p3,false,5),close=(x,y,t)=>Math.abs(x-y)<=Math.max(t,Math.abs(y)*1e-9);
  const checks={p2SourceRows:a.sourceRows===sooRows,p2ValueMatches:close(a.value,sooValue,.01),p2CartonMatches:close(a.carton,sooCarton,.0001),
   subdistRowCountMatches:sub.getLastRow()===p2.getLastRow(),subdistValueMatches:close(b.value,a.value,.01),
   subdistCartonMatches:close(b.carton,a.carton,.0001),subdistSourceRowsMatches:b.sourceRows===a.sourceRows,
   outletMonthlyValueMatches:close(c.value,salesmanValue,.01),outletMonthlyCartonMatches:close(c.carton,salesmanCarton,.0001)};
  const failed=Object.keys(checks).filter(k=>!checks[k]);if(failed.length)throw new Error('Audit derived index gagal: '+failed.join(', ')+'. Tidak ada sheet yang diubah.');
  const n=sub.getLastRow()-1,values=sub.getRange(2,6,n,1).getDisplayValues(),ranges={};let current='',start=2,count=0,previous='';
  values.forEach((r,i)=>{const key=page2Key_(r[0]),row=i+2;if(previous&&key<previous)throw new Error('P2_SUBDIST_FACT tidak terurut berdasarkan SUBDIST.');previous=key;
   if(key!==current){if(current)ranges[current]={startRow:start,numRows:count};current=key;start=row;count=1;}else count++;});
  if(current)ranges[current]={startRow:start,numRows:count};
  const now=new Date().toISOString(),props=PropertiesService.getScriptProperties();
  const stats={sourceRowCount:sourceCount,sourceLastRow:src.getLastRow(),salesTypeCounts:salesTypeCounts,teamCounts:teamCounts,
   sooSourceRows:sooRows,p2SummaryRows:p2.getLastRow()-1,p2SubdistRows:sub.getLastRow()-1,p3SummaryRows:p3.getLastRow()-1,builtAt:now,auditOnly:true};
  props.setProperty(ANALYTICS_SUMMARY_STATS_KEY_,JSON.stringify(stats));props.setProperty(ANALYTICS_SUMMARY_STATUS_KEY_,'COMPLETE');
  props.setProperty('MANTUL_P2_SUBDIST_RANGES',JSON.stringify({rowCount:n,ranges:ranges,builtAt:now}));
  return {ok:true,status:'COMPLETE',auditOnly:true,sheetsMutated:false,checks:checks,sourceRows:sourceCount,sooRows:sooRows,
   derived:{p2Rows:p2.getLastRow()-1,subdistRows:n,p3OutletRows:p3.getLastRow()-1},
   totals:{sooValue:sooValue,sooCarton:sooCarton,salesmanValue:salesmanValue,salesmanCarton:salesmanCarton}};
 }finally{lock.releaseLock();}
}
function analyticsAuditDerivedIndexesReadOnly(){
 const ss=getDatabase_(),names=['RAW_DATA_SALES','P2_EVENT_INDEX','P2_DASHBOARD_FACT','P2_SUBDIST_FACT','P3_OUTLET_MONTHLY','SKU','MAPPING TEAM'],sheets={};
 names.forEach(name=>{const sh=ss.getSheetByName(name);sheets[name]=sh?{exists:true,rows:Math.max(0,sh.getLastRow()-1),columns:sh.getLastColumn(),
  headers:sh.getLastColumn()?sh.getRange(1,1,1,Math.min(sh.getLastColumn(),30)).getDisplayValues()[0]:[]}:{exists:false,rows:0,columns:0,headers:[]};});
 const props=PropertiesService.getScriptProperties();let stats=null,ranges=null;
 try{stats=JSON.parse(props.getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){}
 try{ranges=JSON.parse(props.getProperty('MANTUL_P2_SUBDIST_RANGES')||'null');}catch(e){}
 const checks={sourceIndexExists:sheets.P2_EVENT_INDEX.exists,dashboardFactExists:sheets.P2_DASHBOARD_FACT.exists,
  subdistFactExists:sheets.P2_SUBDIST_FACT.exists,outletMonthlyExists:sheets.P3_OUTLET_MONTHLY.exists,
  dashboardAndSubdistRowCountsEqual:sheets.P2_DASHBOARD_FACT.rows===sheets.P2_SUBDIST_FACT.rows,
  sourceStatsMatch:!!stats&&Number(stats.sourceLastRow)===sheets.P2_EVENT_INDEX.rows+1,
  subdistRangeIndexMatches:!!ranges&&Number(ranges.rowCount)===sheets.P2_SUBDIST_FACT.rows};
 const result={ok:true,readOnly:true,buildExecuted:false,spreadsheetId:ss.getId(),
  buildStatus:props.getProperty(ANALYTICS_SUMMARY_STATUS_KEY_)||'NOT_BUILT',checks:checks,sheets:sheets,
  storedStats:stats?{sourceLastRow:stats.sourceLastRow,p2SummaryRows:stats.p2SummaryRows,p2SubdistRows:stats.p2SubdistRows,p3SummaryRows:stats.p3SummaryRows,builtAt:stats.builtAt}:null,
  subdistRangeIndex:ranges?{rowCount:ranges.rowCount,subdistCount:Object.keys(ranges.ranges||{}).length,builtAt:ranges.builtAt}:null};
 Logger.log('MANTUL DERIVED INDEX AUDIT\n'+JSON.stringify(result,null,2));return result;
}

/** Register existing derived indexes once, on first Page 2/3 visit. Safe: no sheet writes. */
function analyticsEnsureExistingIndexesRegistered_(){
 try{
  if(typeof analyticsSummaryReady_==='function'&&analyticsSummaryReady_())return {ok:true,skipped:true,status:'ALREADY_READY'};
  if(typeof analyticsRegisterExistingIndexes!=='function')return {ok:false,status:'AUDIT_FUNCTION_UNAVAILABLE'};
  return analyticsRegisterExistingIndexes();
 }catch(err){Logger.log('MANTUL INDEX REGISTRATION SKIPPED: '+(err&&err.stack||err));return {ok:false,status:'AUDIT_FAILED',message:String(err)};}
}
