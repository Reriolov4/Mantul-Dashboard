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
    p3=ss.getSheetByName('P3_OUTLET_MONTHLY');
  let stats=null,ranges=null;
  try{stats=JSON.parse(props.getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){}
  try{ranges=JSON.parse(props.getProperty('MANTUL_P2_SUBDIST_RANGES')||'null');}catch(e){}
  const p2Rows=p2?Math.max(0,p2.getLastRow()-1):0;
  const checks={
    statusComplete:status==='COMPLETE',
    sourcePresent:!!(source&&source.getLastRow()>1),
    dashboardFactPresent:!!(p2&&p2Rows>0),
    outletMonthlyPresent:!!(p3&&p3.getLastRow()>1),
    sourceRowCountMatches:!!(source&&stats&&Number(stats.sourceLastRow)===source.getLastRow()),
    subdistRangeIndexMatches:!!(ranges&&Number(ranges.rowCount)===p2Rows)
  };
  return {ready:Object.keys(checks).every(k=>checks[k]),status:status,checks:checks,
    sheets:{sourceRows:source?Math.max(0,source.getLastRow()-1):0,dashboardFactRows:p2Rows,
      outletMonthlyRows:p3?Math.max(0,p3.getLastRow()-1):0},
    stats:stats?{sourceLastRow:stats.sourceLastRow,p2SummaryRows:stats.p2SummaryRows,
      p3SummaryRows:stats.p3SummaryRows,builtAt:stats.builtAt}:null,
    rangeIndex:ranges?{rowCount:ranges.rowCount,subdistCount:Object.keys(ranges.ranges||{}).length,builtAt:ranges.builtAt}:null};
}
function analyticsSummaryReady_(){
 const p2=getDatabase_().getSheetByName('P2_DASHBOARD_FACT');
 if(p2&&p2.getLastRow()>1&&p2.getLastColumn()>=17){
  const h=p2.getRange(1,1,1,17).getDisplayValues()[0].map(page2HeaderKey_);
  if(MANTUL_V2_SCHEMA_.P2_DASHBOARD_FACT.every((v,i)=>page2HeaderKey_(v)===h[i]))return false;
 }
 return analyticsSummaryReadiness_().ready;
}
function analyticsGetSourceStats_(){if(!analyticsSummaryReady_())return null;try{return JSON.parse(PropertiesService.getScriptProperties().getProperty(ANALYTICS_SUMMARY_STATS_KEY_)||'null');}catch(e){return null;}}
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
function analyticsAggregateP3Rows_(rows,yf,mf){const map=new Map();rows.forEach(r=>{if(page2Key_(r[11])!=='SOO')return;const y=analyticsYear_(r[0]),m=analyticsMonth_(r[1]),c=page2Text_(r[4]),s=page2Text_(r[5]);if(!y||!m||!c||(yf&&y!==Number(yf))||(mf&&m!==mf))return;const k=JSON.stringify([y,m,c,s]);let x=map.get(k);if(!x){x=[y,m,c,s,0,0];map.set(k,x);}x[4]+=Number(r[14])||0;x[5]+=Number(r[15])||0;});return Array.from(map.values());}
function analyticsWriteSheetAtomic_(ss,name,headers,rows){const id=Utilities.getUuid().replace(/-/g,'').slice(0,8),tmp=(name+'_BUILD_'+id).slice(0,99),bak=(name+'_OLD_'+id).slice(0,99),sh=ss.insertSheet(tmp);let old=null,renamed=false;try{ensureSheetGrid_(sh,Math.max(1,rows.length+1),headers.length);sh.getRange(1,1,1,headers.length).setValues([headers]);if(rows.length)sh.getRange(2,1,rows.length,headers.length).setValues(rows);sh.setFrozenRows(1);old=ss.getSheetByName(name);if(old){old.setName(bak);renamed=true;}sh.setName(name);if(renamed)ss.deleteSheet(old);}catch(e){try{if(renamed&&ss.getSheetByName(bak)&&!ss.getSheetByName(name))ss.getSheetByName(bak).setName(name);}catch(x){}try{const q=ss.getSheetByName(tmp);if(q)ss.deleteSheet(q);}catch(x){}throw e;}}
function analyticsDeleteSummaryPeriod_(sh,y,m){const n=sh.getLastRow()-1;if(n<=0)return 0;const a=sh.getRange(2,1,n,2).getValues(),blocks=[];let st=-1,c=0;for(let i=0;i<a.length;i++){const ok=Number(a[i][0])===Number(y)&&page2Month_(a[i][1])===page2Month_(m);if(ok){if(st<0)st=i+2;c++;}else if(st>=0){blocks.push([st,c]);st=-1;c=0;}}if(st>=0)blocks.push([st,c]);blocks.sort((a,b)=>b[0]-a[0]).forEach(x=>sh.deleteRows(x[0],x[1]));return blocks.reduce((s,x)=>s+x[1],0);}
function analyticsUpdateDashboardPeriod_(y,m,rows){
  if(!analyticsSummaryReady_())return {updated:false,reason:'Initial derived index belum dibangun.'};
  const ss=getDatabase_(),a=ss.getSheetByName('P2_DASHBOARD_FACT'),b=ss.getSheetByName('P3_OUTLET_MONTHLY'),
    p2=analyticsAggregateP2Rows_(rows||[]),p3=analyticsAggregateP3Rows_(rows||[],y,m);
  analyticsDeleteSummaryPeriod_(a,y,m);analyticsDeleteSummaryPeriod_(b,y,m);
  if(p2.length){
    const r=a.getLastRow()+1;ensureSheetGrid_(a,r+p2.length-1,P2_DASHBOARD_HEADERS_.length);
    a.getRange(r,1,p2.length,P2_DASHBOARD_HEADERS_.length).setValues(p2);
  }
  // Re-sort the canonical P2 fact and rebuild its SUBDIST range index.
  const p2Count=a.getLastRow()-1;
  if(p2Count>1)a.getRange(2,1,p2Count,P2_DASHBOARD_HEADERS_.length)
    .sort([{column:6,ascending:true},{column:1,ascending:true},{column:2,ascending:true}]);
  analyticsStoreSubdistRangesFromSheet_(a);
  if(p3.length){
    const r=b.getLastRow()+1;ensureSheetGrid_(b,r+p3.length-1,P3_OUTLET_HEADERS_.length);
    b.getRange(r,1,p3.length,P3_OUTLET_HEADERS_.length).setValues(p3);
  }
  analyticsRefreshSourceStats_();
  return {updated:true,p2Rows:p2.length,p3Rows:p3.length,subdistIndexReady:true};
}
function analyticsRefreshSourceStats_(){const ss=getDatabase_(),sh=ss.getSheetByName('P2_EVENT_INDEX');if(!sh||sh.getLastRow()<2)return null;const a=sh.getRange(2,12,sh.getLastRow()-1,3).getValues(),salesTypeCounts={},teamCounts={};a.forEach(r=>{const t=page2Key_(r[0]),u=page2Key_(r[2]);if(t)salesTypeCounts[t]=(salesTypeCounts[t]||0)+1;if(u)teamCounts[u]=(teamCounts[u]||0)+1;});const stats={sourceRowCount:sh.getLastRow()-1,sourceLastRow:sh.getLastRow(),salesTypeCounts:salesTypeCounts,teamCounts:teamCounts,p2SummaryRows:ss.getSheetByName('P2_DASHBOARD_FACT').getLastRow()-1,p3SummaryRows:ss.getSheetByName('P3_OUTLET_MONTHLY').getLastRow()-1,builtAt:new Date().toISOString()};PropertiesService.getScriptProperties().setProperty(ANALYTICS_SUMMARY_STATS_KEY_,JSON.stringify(stats));return stats;}
function analyticsGetDashboardIndexStatus(){const readiness=analyticsSummaryReadiness_(),result={ready:readiness.ready,status:readiness.status,checks:readiness.checks,sheets:readiness.sheets,stats:readiness.stats,rangeIndex:readiness.rangeIndex};Logger.log('ANALYTICS INDEX STATUS\\n'+JSON.stringify(result,null,2));console.log('ANALYTICS INDEX STATUS',result);return result;}


