/**
 * MANTUL — SAFE DERIVED INDEX AUDIT (SINGLE P2 FACT)
 * Audit never mutates spreadsheet sheets. Registration writes Script Properties only.
 */
function analyticsRegisterExistingIndexes(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=getDatabase_(),src=ss.getSheetByName('P2_EVENT_INDEX'),p2=ss.getSheetByName('P2_DASHBOARD_FACT'),
    p3=ss.getSheetByName('P3_OUTLET_MONTHLY');
  if(!src||src.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum berisi data.');
  if(!p2||p2.getLastRow()<2||!p3||p3.getLastRow()<2)
   throw new Error('Derived sheet belum lengkap. Dibutuhkan P2_DASHBOARD_FACT dan P3_OUTLET_MONTHLY.');
  analyticsAssertHeaders_(src,P2_EVENT_HEADERS_);
  if(typeof P2_EVENT_INDEX_HEADERS_!=='undefined')analyticsAssertHeaders_(src,P2_EVENT_INDEX_HEADERS_);
  analyticsAssertHeaders_(p2,P2_DASHBOARD_HEADERS_);analyticsAssertHeaders_(p3,P3_OUTLET_HEADERS_);
  const sourceCount=src.getLastRow()-1,meta=src.getRange(2,12,sourceCount,3).getValues(),
    vals=src.getRange(2,15,sourceCount,2).getValues();
  let sooRows=0,sooValue=0,sooCarton=0;const salesTypeCounts={},teamCounts={};
  for(let i=0;i<sourceCount;i++){
   const type=page2Key_(meta[i][0]),team=page2Key_(meta[i][2]);
   if(type)salesTypeCounts[type]=(salesTypeCounts[type]||0)+1;
   if(team)teamCounts[team]=(teamCounts[team]||0)+1;
   if(type!=='SOO')continue;
   sooRows++;sooValue+=Number(vals[i][0])||0;sooCarton+=Number(vals[i][1])||0;
  }
  const readTotals=(sh,start,width)=>sh.getRange(2,start,sh.getLastRow()-1,width).getValues()
    .reduce((t,r)=>{t.value+=Number(r[0])||0;t.carton+=Number(r[1])||0;if(width===3)t.sourceRows+=Number(r[2])||0;return t;},
      {value:0,carton:0,sourceRows:0});
  const a=readTotals(p2,15,3),c=readTotals(p3,5,2),
    close=(x,y,t)=>Math.abs(x-y)<=Math.max(t,Math.abs(y)*1e-9);
  const n=p2.getLastRow()-1,subdistValues=p2.getRange(2,6,n,1).getDisplayValues(),
    ranges={};let current='',start=2,count=0,previous='';
  subdistValues.forEach((r,i)=>{
   const key=page2Key_(r[0]),row=i+2;
   if(previous&&key<previous)throw new Error('P2_DASHBOARD_FACT belum terurut berdasarkan SUBDIST.');
   previous=key;
   if(key!==current){if(current)ranges[current]={startRow:start,numRows:count};current=key;start=row;count=1;}else count++;
  });
  if(current)ranges[current]={startRow:start,numRows:count};
  const checks={p2SourceRows:a.sourceRows===sooRows,p2ValueMatches:close(a.value,sooValue,.01),
    p2CartonMatches:close(a.carton,sooCarton,.0001),outletMonthlyValueMatches:close(c.value,sooValue,.01),
    outletMonthlyCartonMatches:close(c.carton,sooCarton,.0001)};
  const failed=Object.keys(checks).filter(k=>!checks[k]);
  if(failed.length)throw new Error('Audit derived index gagal: '+failed.join(', ')+'. Tidak ada sheet yang diubah.');
  const now=new Date().toISOString(),props=PropertiesService.getScriptProperties();
  const stats={sourceRowCount:sourceCount,sourceLastRow:src.getLastRow(),salesTypeCounts:salesTypeCounts,
    teamCounts:teamCounts,sooSourceRows:sooRows,p2SummaryRows:n,p3SummaryRows:p3.getLastRow()-1,
    builtAt:now,auditOnly:true};
  props.setProperty(ANALYTICS_SUMMARY_STATS_KEY_,JSON.stringify(stats));
  props.setProperty(ANALYTICS_SUMMARY_STATUS_KEY_,'COMPLETE');
  props.setProperty('MANTUL_P2_SUBDIST_RANGES',JSON.stringify({rowCount:n,ranges:ranges,builtAt:now}));
  return {ok:true,status:'COMPLETE',auditOnly:true,sheetsMutated:false,checks:checks,
    sourceRows:sourceCount,sooRows:sooRows,derived:{p2Rows:n,p3OutletRows:p3.getLastRow()-1},
    legacyDuplicateSheetPresent:!!ss.getSheetByName('P2_SUBDIST_FACT'),
    totals:{sooValue:sooValue,sooCarton:sooCarton,outletMonthlyValue:c.value,outletMonthlyCarton:c.carton}};
 }finally{lock.releaseLock();}
}
function analyticsAuditDerivedIndexesReadOnly(){
 const ss=getDatabase_(),names=['RAW_DATA_SALES','P2_EVENT_INDEX','P2_DASHBOARD_FACT','P2_SUBDIST_FACT','P3_OUTLET_MONTHLY','SKU','MAPPING TEAM'],sheets={};
 names.forEach(name=>{const sh=ss.getSheetByName(name);sheets[name]=sh?{exists:true,rows:Math.max(0,sh.getLastRow()-1),columns:sh.getLastColumn(),
  headers:sh.getLastColumn()?sh.getRange(1,1,1,Math.min(sh.getLastColumn(),30)).getDisplayValues()[0]:[]}:{exists:false,rows:0,columns:0,headers:[]};});
 const props=PropertiesService.getScriptProperties();let stats=null,ranges=null;
 try{stats=JSON.parse(props.getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){}
 try{ranges=JSON.parse(props.getProperty('MANTUL_P2_SUBDIST_RANGES')||'null');}catch(e){}
 const p2Rows=sheets.P2_DASHBOARD_FACT.rows;
 const checks={sourceIndexExists:sheets.P2_EVENT_INDEX.exists,dashboardFactExists:sheets.P2_DASHBOARD_FACT.exists,
  outletMonthlyExists:sheets.P3_OUTLET_MONTHLY.exists,
  sourceStatsMatch:!!stats&&Number(stats.sourceLastRow)===sheets.P2_EVENT_INDEX.rows+1,
  subdistRangeIndexMatches:!!ranges&&Number(ranges.rowCount)===p2Rows};
 const result={ok:true,readOnly:true,buildExecuted:false,spreadsheetId:ss.getId(),buildStatus:props.getProperty(ANALYTICS_SUMMARY_STATUS_KEY_)||'NOT_BUILT',
  checks:checks,sheets:sheets,legacyDuplicateSheetPresent:sheets.P2_SUBDIST_FACT.exists,
  storedStats:stats?{sourceLastRow:stats.sourceLastRow,p2SummaryRows:stats.p2SummaryRows,p3SummaryRows:stats.p3SummaryRows,builtAt:stats.builtAt}:null,
  subdistRangeIndex:ranges?{rowCount:ranges.rowCount,subdistCount:Object.keys(ranges.ranges||{}).length,builtAt:ranges.builtAt}:null};
 Logger.log('MANTUL DERIVED INDEX AUDIT\\n'+JSON.stringify(result,null,2));return result;
}

/** Register existing derived indexes once, on first Page 2/3 visit. Safe: no sheet writes. */
function analyticsEnsureExistingIndexesRegistered_(){
 try{
  if(typeof analyticsSummaryReady_==='function'&&analyticsSummaryReady_())return {ok:true,skipped:true,status:'ALREADY_READY'};
  if(typeof analyticsRegisterExistingIndexes!=='function')return {ok:false,status:'AUDIT_FUNCTION_UNAVAILABLE'};
  return analyticsRegisterExistingIndexes();
 }catch(err){Logger.log('MANTUL INDEX REGISTRATION SKIPPED: '+(err&&err.stack||err));return {ok:false,status:'AUDIT_FAILED',message:String(err)};}
}


/**
 * Remove the legacy duplicate only after the canonical fact reconciles to source
 * and every value matches. Sorting both sheets first makes row order irrelevant.
 */
function analyticsRemoveLegacySubdistFact(confirmation){
 if(confirmation!=='REMOVE_DUPLICATE_P2_SUBDIST_FACT')
  throw new Error('Penghapusan ditahan. Gunakan konfirmasi REMOVE_DUPLICATE_P2_SUBDIST_FACT.');
 const ss=getDatabase_(),canonical=ss.getSheetByName('P2_DASHBOARD_FACT'),
  legacy=ss.getSheetByName('P2_SUBDIST_FACT');
 if(!legacy)return {ok:true,removed:false,message:'P2_SUBDIST_FACT sudah tidak ada.'};
 if(!canonical||canonical.getLastRow()<2)throw new Error('P2_DASHBOARD_FACT tidak tersedia; sheet legacy tidak dihapus.');
 const rows=canonical.getLastRow()-1,width=P2_DASHBOARD_HEADERS_.length;
 if(legacy.getLastRow()-1!==rows||legacy.getLastColumn()<width)
  throw new Error('Jumlah baris/kolom legacy berbeda; sheet tidak dihapus.');
 // Canonical sort order starts with SUBDIST, then stable dimensions and measures.
 const sortSpecs=[6,1,2,3,4,5,7,8,9,10,11,12,13,14,15,16,17]
  .map(column=>({column:column,ascending:true}));
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const currentCanonical=ss.getSheetByName('P2_DASHBOARD_FACT'),
   currentLegacy=ss.getSheetByName('P2_SUBDIST_FACT');
  if(!currentCanonical||!currentLegacy)throw new Error('Sheet berubah saat proses validasi; batal.');
  currentCanonical.getRange(2,1,rows,width).sort(sortSpecs);
  currentLegacy.getRange(2,1,rows,width).sort(sortSpecs);
  const equalValue=(a,b)=>{
   if(a instanceof Date&&b instanceof Date)return a.getTime()===b.getTime();
   return a===b || (a==null&&b==='') || (a===''&&b==null);
  };
  const batch=2000;
  for(let start=2;start<=rows+1;start+=batch){
   const count=Math.min(batch,rows+2-start);
   const a=currentCanonical.getRange(start,1,count,width).getValues();
   const b=currentLegacy.getRange(start,1,count,width).getValues();
   for(let i=0;i<count;i++)for(let j=0;j<width;j++){
    if(!equalValue(a[i][j],b[i][j]))
     throw new Error('Isi P2_SUBDIST_FACT berbeda dari P2_DASHBOARD_FACT pada baris '+(start+i)+', kolom '+(j+1)+'. Sheet tidak dihapus.');
   }
  }
 }finally{lock.releaseLock();}
 // Audit canonical against P2_EVENT_INDEX and P3 before deleting the duplicate.
 const audit=analyticsRegisterExistingIndexes();
 if(!audit||!audit.ok)throw new Error('Audit canonical P2 gagal; sheet duplikat tidak dihapus.');
 const deleteLock=LockService.getScriptLock();deleteLock.waitLock(30000);
 try{
  const latest=ss.getSheetByName('P2_SUBDIST_FACT');
  if(latest)ss.deleteSheet(latest);
  return {ok:true,removed:!!latest,rowsCompared:rows,columnsCompared:width,
   sourceAuditPassed:true,remainingCanonicalSheet:'P2_DASHBOARD_FACT'};
 }finally{deleteLock.releaseLock();}
}

/** One-click guarded cleanup; compares all rows and audits source before deletion. */
function runMantulRemoveLegacyP2Duplicate(){
 return analyticsRemoveLegacySubdistFact('REMOVE_DUPLICATE_P2_SUBDIST_FACT');
}
