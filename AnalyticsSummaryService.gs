/** MANTUL DASHBOARD — DERIVED ANALYTICS READ MODELS */
const ANALYTICS_SUMMARY_STATUS_KEY_='MANTUL_DERIVED_INDEX_STATUS';
const ANALYTICS_SUMMARY_STATS_KEY_='MANTUL_DERIVED_INDEX_STATS';
const P2_DASHBOARD_HEADERS_=P2_EVENT_HEADERS_.concat(['SOURCE_ROWS']);
const P3_OUTLET_HEADERS_=['YEAR','MONTH','CUSTOMER_NAME','SUBDIST','VAL','CARTON'];
function analyticsMonth_(value){
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value))return PAGE2_MONTHS_[value.getMonth()];
  const normalized=page2Month_(value);
  return PAGE2_MONTHS_.indexOf(normalized)>=0?normalized:'';
}
function analyticsYear_(value){
  const direct=Number(value);if(direct>=1900&&direct<=2200)return direct;
  if(Object.prototype.toString.call(value)==='[object Date]'&&!isNaN(value))return value.getFullYear();
  const match=String(value==null?'':value).match(/(?:19|20|21)\d{2}/);return match?Number(match[0]):0;
}

function analyticsSummaryReadiness_(){
  const props=PropertiesService.getScriptProperties(),status=props.getProperty(ANALYTICS_SUMMARY_STATUS_KEY_)||'NOT_BUILT';
  const ss=getDatabase_(),source=ss.getSheetByName('P2_EVENT_INDEX'),p2=ss.getSheetByName('P2_DASHBOARD_FACT'),
    sub=ss.getSheetByName('P2_SUBDIST_FACT'),p3=ss.getSheetByName('P3_OUTLET_MONTHLY');
  let stats=null,ranges=null;
  try{stats=JSON.parse(props.getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){}
  try{ranges=JSON.parse(props.getProperty('MANTUL_P2_SUBDIST_RANGES')||'null');}catch(e){}
  const checks={
    statusComplete:status==='COMPLETE',
    sourcePresent:!!(source&&source.getLastRow()>1),
    dashboardFactPresent:!!(p2&&p2.getLastRow()>1),
    subdistFactPresent:!!(sub&&sub.getLastRow()>1),
    outletMonthlyPresent:!!(p3&&p3.getLastRow()>1),
    sourceRowCountMatches:!!(source&&stats&&Number(stats.sourceLastRow)===source.getLastRow()),
    subdistRangeIndexMatches:!!(sub&&ranges&&Number(ranges.rowCount)===sub.getLastRow()-1)
  };
  return {ready:Object.keys(checks).every(k=>checks[k]),status:status,checks:checks,
    sheets:{sourceRows:source?Math.max(0,source.getLastRow()-1):0,dashboardFactRows:p2?Math.max(0,p2.getLastRow()-1):0,
      subdistFactRows:sub?Math.max(0,sub.getLastRow()-1):0,outletMonthlyRows:p3?Math.max(0,p3.getLastRow()-1):0},
    stats:stats?{sourceLastRow:stats.sourceLastRow,p2SummaryRows:stats.p2SummaryRows,p2SubdistRows:stats.p2SubdistRows,p3SummaryRows:stats.p3SummaryRows,builtAt:stats.builtAt}:null,
    rangeIndex:ranges?{rowCount:ranges.rowCount,subdistCount:Object.keys(ranges.ranges||{}).length,builtAt:ranges.builtAt}:null};
}
function analyticsSummaryReady_(){return analyticsSummaryReadiness_().ready;}
function analyticsGetSourceStats_(){if(!analyticsSummaryReady_())return null;try{return JSON.parse(PropertiesService.getScriptProperties().getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){return null;}}

// Resumable derived-index builder. Each invocation stays bounded and schedules
// the next phase; active published facts remain available while a rebuild runs.
const ANALYTICS_BUILD_STATE_KEY_='MANTUL_DERIVED_INDEX_BUILD_STATE';
const ANALYTICS_BUILD_BATCH_ROWS_=4000;
const ANALYTICS_BUILD_TRIGGER_HANDLER_='analyticsBuildDashboardIndexes';

function analyticsBuildDashboardIndexes(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return {ok:false,status:'BUSY'};
  const props=PropertiesService.getScriptProperties();
  let state=null;
  try{
    state=analyticsGetBuildState_();
    if(!state)state=analyticsStartBuild_();
    const result=analyticsBuildStep_(state);
    if(result.done){
      props.deleteProperty(ANALYTICS_BUILD_STATE_KEY_);
      analyticsClearBuildTriggers_();
      Logger.log('ANALYTICS BUILD COMPLETE\\n'+JSON.stringify(result,null,2));
      return result;
    }
    props.setProperty(ANALYTICS_BUILD_STATE_KEY_,JSON.stringify(state));
    analyticsScheduleBuildContinuation_();
    Logger.log('ANALYTICS BUILD CHECKPOINT\\n'+JSON.stringify(result,null,2));
    return result;
  }catch(e){
    if(state&&state.previousStatus!=='COMPLETE')
      props.setProperty(ANALYTICS_SUMMARY_STATUS_KEY_,'FAILED');
    if(state)state.lastError=String(e&&e.stack||e);
    if(state)props.setProperty(ANALYTICS_BUILD_STATE_KEY_,JSON.stringify(state));
    Logger.log('ANALYTICS BUILD STEP FAILED: '+(e&&e.stack||e));
    throw e;
  }finally{lock.releaseLock();}
}

function analyticsGetBuildState_(){
  try{return JSON.parse(PropertiesService.getScriptProperties().getProperty(ANALYTICS_BUILD_STATE_KEY_)||'null');}
  catch(e){throw Error('Checkpoint BUILD rusak; jangan hapus sheet. Detail: '+e);}
}

function analyticsStartBuild_(){
  const ss=getDatabase_(),src=ss.getSheetByName('P2_EVENT_INDEX');
  if(!src||src.getLastRow()<2)throw Error('P2_EVENT_INDEX belum berisi transaksi.');
  const actual=src.getRange(1,1,1,16).getDisplayValues()[0].map(page2HeaderKey_);
  if(!P2_EVENT_HEADERS_.every((x,i)=>page2HeaderKey_(x)===actual[i]))
    throw Error('Header P2_EVENT_INDEX tidak sesuai 16 kolom: '+actual.join(' | '));
  const ready=analyticsSummaryReadiness_();
  const id=Utilities.getUuid().replace(/-/g,'').slice(0,8);
  const state={id:id,phase:'SOURCE',sourceLastRow:src.getLastRow(),nextRow:2,
    periods:[],periodRows:{},periodDone:0,p2OutRows:0,p3OutRows:0,
    p2Stage:'P2_DASHBOARD_FACT_BUILD_'+id,p3Stage:'P3_OUTLET_MONTHLY_BUILD_'+id,
    subStage:'P2_SUBDIST_FACT_BUILD_'+id,previousStatus:ready.ready?'COMPLETE':'NOT_READY',
    stats:{salesTypeCounts:{},teamCounts:{}}};
  analyticsCreateBuildSheet_(ss,state.p2Stage,P2_DASHBOARD_HEADERS_);
  analyticsCreateBuildSheet_(ss,state.p3Stage,P3_OUTLET_HEADERS_);
  if(!ready.ready)PropertiesService.getScriptProperties().setProperty(ANALYTICS_SUMMARY_STATUS_KEY_,'BUILDING');
  PropertiesService.getScriptProperties().setProperty(ANALYTICS_BUILD_STATE_KEY_,JSON.stringify(state));
  return state;
}

function analyticsBuildStep_(state){
  const ss=getDatabase_(),props=PropertiesService.getScriptProperties();
  if(state.phase==='SOURCE'){
    const src=ss.getSheetByName('P2_EVENT_INDEX');
    if(!src||src.getLastRow()!==state.sourceLastRow)throw Error('P2_EVENT_INDEX berubah saat BUILD berjalan. Hentikan dan mulai BUILD baru.');
    if(state.nextRow<=state.sourceLastRow){
      const count=Math.min(ANALYTICS_BUILD_BATCH_ROWS_,state.sourceLastRow-state.nextRow+1);
      const rows=src.getRange(state.nextRow,1,count,16).getValues(),groups={};
      rows.forEach(r=>{
        const y=analyticsYear_(r[0]),m=analyticsMonth_(r[1]);
        const key=y&&m?('Y'+y+'M'+String(PAGE2_MONTHS_.indexOf(m)+1).padStart(2,'0')):'INVALID';
        (groups[key]||(groups[key]=[])).push(r);
      });
      Object.keys(groups).forEach(key=>{
        const name=analyticsSourceStageName_(state.id,key),sh=analyticsCreateBuildSheet_(ss,name,P2_EVENT_HEADERS_);
        const offset=Number(state.periodRows[key]||0);
        analyticsWriteAt_(sh,2+offset,groups[key],16);
        state.periodRows[key]=offset+groups[key].length;
        if(state.periods.indexOf(key)<0)state.periods.push(key);
      });
      analyticsMergeBuildCounts_(state.stats,analyticsCountSourceRows_(rows));
      state.nextRow+=count;
      return {done:false,phase:state.phase,readRows:count,nextSourceRow:state.nextRow,sourceRows:state.sourceLastRow-1};
    }
    state.periods.sort();
    state.phase='AGGREGATE';
    return {done:false,phase:state.phase,periods:state.periods.length};
  }
  if(state.phase==='AGGREGATE'){
    if(state.periodDone<state.periods.length){
      const key=state.periods[state.periodDone],name=analyticsSourceStageName_(state.id,key),sh=ss.getSheetByName(name);
      const raw=sh&&sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,16).getValues():[];
      const parsed=analyticsParsePeriodKey_(key);
      const p2=analyticsAggregateP2Rows_(raw),p3=analyticsAggregateP3Rows_(raw,parsed.year,parsed.month);
      const p2sh=ss.getSheetByName(state.p2Stage),p3sh=ss.getSheetByName(state.p3Stage);
      analyticsWriteAt_(p2sh,2+state.p2OutRows,p2,P2_DASHBOARD_HEADERS_.length);
      analyticsWriteAt_(p3sh,2+state.p3OutRows,p3,P3_OUTLET_HEADERS_.length);
      state.p2OutRows+=p2.length;state.p3OutRows+=p3.length;state.periodDone++;
      if(sh)ss.deleteSheet(sh);
      return {done:false,phase:state.phase,period:key,p2Rows:p2.length,p3Rows:p3.length,periodsDone:state.periodDone,periodsTotal:state.periods.length};
    }
    state.phase='PUBLISH';
    return {done:false,phase:state.phase,p2Rows:state.p2OutRows,p3Rows:state.p3OutRows};
  }
  if(state.phase==='PUBLISH'){
    const p2=ss.getSheetByName(state.p2Stage),p3=ss.getSheetByName(state.p3Stage);
    if(!p2||!p3)throw Error('Sheet staging output hilang; data aktif tidak dihapus.');
    const sub=ss.getSheetByName(state.subStage)||p2.copyTo(ss).setName(state.subStage);
    const subRows=sub.getLastRow()-1;
    if(subRows>1)sub.getRange(2,1,subRows,P2_DASHBOARD_HEADERS_.length)
      .sort([{column:6,ascending:true},{column:1,ascending:true},{column:2,ascending:true}]);
    analyticsPublishStagedSheet_(ss,state.p2Stage,'P2_DASHBOARD_FACT',state.id);
    analyticsPublishStagedSheet_(ss,state.subStage,'P2_SUBDIST_FACT',state.id);
    analyticsStoreSubdistRangesFromSheet_(ss.getSheetByName('P2_SUBDIST_FACT'));
    analyticsPublishStagedSheet_(ss,state.p3Stage,'P3_OUTLET_MONTHLY',state.id);
    const src=ss.getSheetByName('P2_EVENT_INDEX');
    if(!src||src.getLastRow()!==state.sourceLastRow)throw Error('P2_EVENT_INDEX berubah sebelum publikasi selesai.');
    const stats=Object.assign(state.stats,{sourceRowCount:state.sourceLastRow-1,sourceLastRow:state.sourceLastRow,
      p2SummaryRows:state.p2OutRows,p2SubdistRows:state.p2OutRows,p3SummaryRows:state.p3OutRows,builtAt:new Date().toISOString()});
    props.setProperty(ANALYTICS_SUMMARY_STATS_KEY_,JSON.stringify(stats));
    props.setProperty(ANALYTICS_SUMMARY_STATUS_KEY_,'COMPLETE');
    analyticsCleanupBuildSheets_(ss,state.id);
    return {done:true,ok:true,status:'COMPLETE',sourceRows:state.sourceLastRow-1,
      p2SummaryRows:state.p2OutRows,p3SummaryRows:state.p3OutRows,stats:stats};
  }
  throw Error('Tahap BUILD tidak dikenal: '+state.phase);
}

function analyticsCreateBuildSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);ensureSheetGrid_(sh,1,headers.length);sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);}
  return sh;
}
function analyticsWriteAt_(sh,startRow,rows,width){
  if(!rows||!rows.length)return;
  ensureSheetGrid_(sh,startRow+rows.length-1,width);
  sh.getRange(startRow,1,rows.length,width).setValues(rows);
}
function analyticsSourceStageName_(id,key){return ('MANTUL_SRC_'+id+'_'+key).slice(0,99);}
function analyticsParsePeriodKey_(key){
  if(key==='INVALID')return {year:0,month:''};
  const m=String(key).match(/^Y(\\d+)M(\\d{2})$/);
  return m?{year:Number(m[1]),month:PAGE2_MONTHS_[Number(m[2])-1]||''}:{year:0,month:''};
}
function analyticsMergeBuildCounts_(target,source){
  ['salesTypeCounts','teamCounts'].forEach(k=>Object.keys(source[k]||{}).forEach(x=>target[k][x]=(target[k][x]||0)+source[k][x]));
}
function analyticsPublishStagedSheet_(ss,stageName,targetName,id){
  const stage=ss.getSheetByName(stageName),target=ss.getSheetByName(targetName),backupName=(targetName+'_OLD_'+id).slice(0,99);
  if(!stage){if(target)return;throw Error('Staging '+stageName+' tidak ditemukan.');}
  if(target){const prior=ss.getSheetByName(backupName);if(prior)ss.deleteSheet(prior);target.setName(backupName);}
  try{stage.setName(targetName);const backup=ss.getSheetByName(backupName);if(backup)ss.deleteSheet(backup);}
  catch(e){const backup=ss.getSheetByName(backupName);if(backup&&!ss.getSheetByName(targetName))backup.setName(targetName);throw e;}
}
function analyticsCleanupBuildSheets_(ss,id){
  ss.getSheets().forEach(sh=>{const n=sh.getName();if(n.indexOf('MANTUL_SRC_'+id+'_')===0||n.indexOf('P2_DASHBOARD_FACT_BUILD_'+id)===0||n.indexOf('P2_SUBDIST_FACT_BUILD_'+id)===0||n.indexOf('P3_OUTLET_MONTHLY_BUILD_'+id)===0||n.indexOf('_OLD_'+id)>=0)ss.deleteSheet(sh);});
}
function analyticsScheduleBuildContinuation_(){
  analyticsClearBuildTriggers_();
  ScriptApp.newTrigger(ANALYTICS_BUILD_TRIGGER_HANDLER_).timeBased().after(60*1000).create();
}
function analyticsClearBuildTriggers_(){
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()===ANALYTICS_BUILD_TRIGGER_HANDLER_)ScriptApp.deleteTrigger(t);});
}
function analyticsSortP2BySubdist_(rows){
  return (rows||[]).slice().sort((a,b)=>{const x=page2Key_(a[5]),y=page2Key_(b[5]);return x<y?-1:x>y?1:0;});
}
function analyticsStoreSubdistRanges_(rows){
  const ranges={};let current='',start=2,count=0;
  (rows||[]).forEach((r,i)=>{const key=page2Key_(r[5]),row=i+2;if(key!==current){if(current)ranges[current]={startRow:start,numRows:count};current=key;start=row;count=1;}else count++;});
  if(current)ranges[current]={startRow:start,numRows:count};
  PropertiesService.getScriptProperties().setProperty('MANTUL_P2_SUBDIST_RANGES',JSON.stringify({rowCount:(rows||[]).length,ranges:ranges,builtAt:new Date().toISOString()}));return ranges;
}
function analyticsStoreSubdistRangesFromSheet_(sheet){
  const n=sheet.getLastRow()-1,ranges={};if(n<=0){PropertiesService.getScriptProperties().setProperty('MANTUL_P2_SUBDIST_RANGES',JSON.stringify({rowCount:0,ranges:ranges,builtAt:new Date().toISOString()}));return ranges;}
  const values=sheet.getRange(2,6,n,1).getValues();let current='',start=2,count=0;
  values.forEach((r,i)=>{const key=page2Key_(r[0]),row=i+2;if(key!==current){if(current)ranges[current]={startRow:start,numRows:count};current=key;start=row;count=1;}else count++;});
  if(current)ranges[current]={startRow:start,numRows:count};
  PropertiesService.getScriptProperties().setProperty('MANTUL_P2_SUBDIST_RANGES',JSON.stringify({rowCount:n,ranges:ranges,builtAt:new Date().toISOString()}));return ranges;
}

function analyticsCountSourceRows_(rows){const salesTypeCounts={},teamCounts={};rows.forEach(r=>{const t=page2Key_(r[11]),u=page2Key_(r[13]);if(t)salesTypeCounts[t]=(salesTypeCounts[t]||0)+1;if(u)teamCounts[u]=(teamCounts[u]||0)+1;});return {salesTypeCounts:salesTypeCounts,teamCounts:teamCounts};}
function analyticsAggregateP2Rows_(rows){const map=new Map();rows.forEach(r=>{if(page2Key_(r[11])!=='SOO')return;const k=JSON.stringify(r.slice(0,14));let x=map.get(k);if(!x){x=r.slice(0,16);x[14]=0;x[15]=0;x[16]=0;map.set(k,x);}x[14]+=Number(r[14])||0;x[15]+=Number(r[15])||0;x[16]++;});return Array.from(map.values());}
function analyticsAggregateP3Rows_(rows,yf,mf){const map=new Map();rows.forEach(r=>{if(page2Key_(r[11])!=='SOO'||page2Key_(r[13])!=='SALESMAN')return;const y=analyticsYear_(r[0]),m=analyticsMonth_(r[1]),c=page2Text_(r[4]),s=page2Text_(r[5]);if(!y||!m||!c||(yf&&y!==Number(yf))||(mf&&m!==mf))return;const k=JSON.stringify([y,m,c,s]);let x=map.get(k);if(!x){x=[y,m,c,s,0,0];map.set(k,x);}x[4]+=Number(r[14])||0;x[5]+=Number(r[15])||0;});return Array.from(map.values());}
function analyticsWriteSheetAtomic_(ss,name,headers,rows){const id=Utilities.getUuid().replace(/-/g,'').slice(0,8),tmp=(name+'_BUILD_'+id).slice(0,99),bak=(name+'_OLD_'+id).slice(0,99),sh=ss.insertSheet(tmp);let old=null,renamed=false;try{ensureSheetGrid_(sh,Math.max(1,rows.length+1),headers.length);sh.getRange(1,1,1,headers.length).setValues([headers]);if(rows.length)sh.getRange(2,1,rows.length,headers.length).setValues(rows);sh.setFrozenRows(1);old=ss.getSheetByName(name);if(old){old.setName(bak);renamed=true;}sh.setName(name);if(renamed)ss.deleteSheet(old);}catch(e){try{if(renamed&&ss.getSheetByName(bak)&&!ss.getSheetByName(name))ss.getSheetByName(bak).setName(name);}catch(x){}try{const q=ss.getSheetByName(tmp);if(q)ss.deleteSheet(q);}catch(x){}throw e;}}
function analyticsDeleteSummaryPeriod_(sh,y,m){const n=sh.getLastRow()-1;if(n<=0)return 0;const a=sh.getRange(2,1,n,2).getValues(),blocks=[];let st=-1,c=0;for(let i=0;i<a.length;i++){const ok=Number(a[i][0])===Number(y)&&page2Month_(a[i][1])===page2Month_(m);if(ok){if(st<0)st=i+2;c++;}else if(st>=0){blocks.push([st,c]);st=-1;c=0;}}if(st>=0)blocks.push([st,c]);blocks.sort((a,b)=>b[0]-a[0]).forEach(x=>sh.deleteRows(x[0],x[1]));return blocks.reduce((s,x)=>s+x[1],0);}
function analyticsUpdateDashboardPeriod_(y,m,rows){
  if(!analyticsSummaryReady_())return {updated:false,reason:'Initial derived index belum dibangun.'};
  const ss=getDatabase_(),a=ss.getSheetByName('P2_DASHBOARD_FACT'),b=ss.getSheetByName('P3_OUTLET_MONTHLY'),sub=ss.getSheetByName('P2_SUBDIST_FACT'),
    p2=analyticsAggregateP2Rows_(rows||[]),p3=analyticsAggregateP3Rows_(rows||[],y,m);
  analyticsDeleteSummaryPeriod_(a,y,m);analyticsDeleteSummaryPeriod_(b,y,m);if(sub)analyticsDeleteSummaryPeriod_(sub,y,m);
  if(p2.length){
    let r=a.getLastRow()+1;ensureSheetGrid_(a,r+p2.length-1,P2_DASHBOARD_HEADERS_.length);a.getRange(r,1,p2.length,P2_DASHBOARD_HEADERS_.length).setValues(p2);
    if(sub){r=sub.getLastRow()+1;ensureSheetGrid_(sub,r+p2.length-1,P2_DASHBOARD_HEADERS_.length);sub.getRange(r,1,p2.length,P2_DASHBOARD_HEADERS_.length).setValues(p2);}
  }
  if(sub){const n=sub.getLastRow()-1;if(n>1)sub.getRange(2,1,n,P2_DASHBOARD_HEADERS_.length).sort([{column:6,ascending:true},{column:1,ascending:true},{column:2,ascending:true}]);analyticsStoreSubdistRangesFromSheet_(sub);}
  if(p3.length){const r=b.getLastRow()+1;ensureSheetGrid_(b,r+p3.length-1,P3_OUTLET_HEADERS_.length);b.getRange(r,1,p3.length,P3_OUTLET_HEADERS_.length).setValues(p3);}
  analyticsRefreshSourceStats_();return {updated:true,p2Rows:p2.length,p3Rows:p3.length,subdistIndexReady:!!sub};
}
function analyticsRefreshSourceStats_(){const ss=getDatabase_(),sh=ss.getSheetByName('P2_EVENT_INDEX');if(!sh||sh.getLastRow()<2)return null;const a=sh.getRange(2,12,sh.getLastRow()-1,3).getValues(),salesTypeCounts={},teamCounts={};a.forEach(r=>{const t=page2Key_(r[0]),u=page2Key_(r[2]);if(t)salesTypeCounts[t]=(salesTypeCounts[t]||0)+1;if(u)teamCounts[u]=(teamCounts[u]||0)+1;});const stats={sourceRowCount:sh.getLastRow()-1,sourceLastRow:sh.getLastRow(),salesTypeCounts:salesTypeCounts,teamCounts:teamCounts,p2SummaryRows:ss.getSheetByName('P2_DASHBOARD_FACT').getLastRow()-1,p2SubdistRows:(ss.getSheetByName('P2_SUBDIST_FACT')||{getLastRow:()=>1}).getLastRow()-1,p3SummaryRows:ss.getSheetByName('P3_OUTLET_MONTHLY').getLastRow()-1,builtAt:new Date().toISOString()};PropertiesService.getScriptProperties().setProperty(ANALYTICS_SUMMARY_STATS_KEY_,JSON.stringify(stats));return stats;}
function analyticsGetDashboardIndexStatus(){const readiness=analyticsSummaryReadiness_(),result={ready:readiness.ready,status:readiness.status,checks:readiness.checks,sheets:readiness.sheets,stats:readiness.stats,rangeIndex:readiness.rangeIndex};Logger.log('ANALYTICS INDEX STATUS\\n'+JSON.stringify(result,null,2));console.log('ANALYTICS INDEX STATUS',result);return result;}
