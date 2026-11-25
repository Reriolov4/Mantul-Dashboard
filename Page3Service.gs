/**
 * PAGE 3 — OUTLETS
 * Read from the maintained SOO aggregate. All TIM values are included.
 * Monthly rows group CUSTOMER_NAME + SUBDIST; no target/achievement fields.
 */
const PAGE3_MONTHS_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function page3Source_(){
 const ss=getDatabase_();let sh=ss.getSheetByName('P2_DASHBOARD_FACT');
 if(!sh||sh.getLastRow()<2)sh=ss.getSheetByName('P2_EVENT_INDEX');
 if(!sh||sh.getLastRow()<2)throw new Error('P2_DASHBOARD_FACT/P2_EVENT_INDEX belum berisi data.');
 const expected=P2_EVENT_HEADERS_.map(page2HeaderKey_);
 const actual=sh.getRange(1,1,1,16).getDisplayValues()[0].map(page2HeaderKey_);
 if(!expected.every((v,i)=>v===actual[i]))throw new Error('Header sumber Page 3 tidak sesuai: '+actual.join(' | '));
 return sh;
}
function page3ReadSource_(filters){
 const ss=getDatabase_(),selected=page2FilterList_(filters&&filters.subdists);
 if(selected.length===1){
  const sub=ss.getSheetByName('P2_SUBDIST_FACT');let index=null;
  try{index=JSON.parse(PropertiesService.getScriptProperties().getProperty('MANTUL_P2_SUBDIST_RANGES')||'null');}catch(e){}
  const entry=index&&index.ranges&&index.ranges[selected[0]];
  if(sub&&sub.getLastRow()>1&&index&&Number(index.rowCount)===sub.getLastRow()-1&&entry&&entry.numRows>0){
   const expected=P2_EVENT_HEADERS_.map(page2HeaderKey_);
   const actual=sub.getRange(1,1,1,16).getDisplayValues()[0].map(page2HeaderKey_);
   if(!expected.every((v,i)=>v===actual[i]))throw new Error('Header P2_SUBDIST_FACT tidak sesuai: '+actual.join(' | '));
   return {sheet:sub,rows:sub.getRange(entry.startRow,1,entry.numRows,16).getValues()};
  }
 }
 const sh=page3Source_(),n=sh.getLastRow()-1;
 return {sheet:sh,rows:sh.getRange(2,1,n,16).getValues()};
}
function page3BuildMeta_(rows){
 const years=new Set(),subdists=new Set(),subbrands=new Set(),monthsByYear={};let soo=0;
 rows.forEach(r=>{if(page2Key_(r[11])!=='SOO')return;soo++;
  const y=String(Number(r[0])||r[0]||'').trim(),m=page2Month_(r[1]),sd=page2Text_(r[5]),sb=page2Text_(r[8]);
  if(y)years.add(y);if(sd)subdists.add(sd);if(sb)subbrands.add(sb);
  if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
 });
 const sort=a=>Array.from(a).sort((x,y)=>x.localeCompare(y,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),
  monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE3_MONTHS_.filter(m=>monthsByYear[y].has(m))])),
  subdists:sort(subdists),subbrands:sort(subbrands),sourceRows:rows.length,sooRows:soo};
}
function page3BuildData_(rows,filters,meta,sourceName){
 filters=filters||{};const year=Number(filters.year)||Number((meta.years||[]).slice(-1)[0])||new Date().getFullYear();
 const available=(meta.monthsByYear||{})[String(year)]||[];
 const requested=(Array.isArray(filters.months)?filters.months:[]).map(page2Month_).filter(m=>PAGE3_MONTHS_.includes(m));
 const months=requested.length?PAGE3_MONTHS_.filter(m=>requested.includes(m)):available.slice();
 const monthSet=new Set(months),subdistSet=new Set(page2FilterList_(filters.subdists)),subbrandSet=new Set(page2FilterList_(filters.subbrands));
 const query=page2Key_(filters.customerName||''),metric=String(filters.metric||'value').toLowerCase()==='carton'?'carton':'value';
 const map=new Map();let matchedSourceRows=0;
 rows.forEach(r=>{
  if(page2Key_(r[11])!=='SOO'||Number(r[0])!==year)return;
  const m=page2Month_(r[1]),sd=page2Text_(r[5]),sb=page2Text_(r[8]),customer=page2Text_(r[4]);
  if(!monthSet.has(m)||!customer)return;
  if(subdistSet.size&&!subdistSet.has(page2Key_(sd)))return;
  if(subbrandSet.size&&!subbrandSet.has(page2Key_(sb)))return;
  if(query&&!page2Key_(customer).includes(query))return;
  matchedSourceRows++;
  const key=page2ExactCompositeKey_([sd,customer]);let x=map.get(key);
  if(!x){x={customerName:customer,subdist:sd,monthly:{},total:{value:0,carton:0}};map.set(key,x);}
  if(!x.monthly[m])x.monthly[m]={value:0,carton:0};
  x.monthly[m].value+=Number(r[14])||0;x.monthly[m].carton+=Number(r[15])||0;
  x.total.value+=Number(r[14])||0;x.total.carton+=Number(r[15])||0;
 });
 const all=Array.from(map.values()).sort((a,b)=>(b.total[metric]-a.total[metric])||a.customerName.localeCompare(b.customerName,'id'));
 const monthTotals={};months.forEach(m=>monthTotals[m]={value:0,carton:0});
 const totals={value:0,carton:0};all.forEach(r=>{totals.value+=r.total.value;totals.carton+=r.total.carton;months.forEach(m=>{monthTotals[m].value+=r.monthly[m]?.value||0;monthTotals[m].carton+=r.monthly[m]?.carton||0;});});
 const page=Math.max(1,Number(filters.page)||1),pageSize=Math.max(10,Math.min(100,Number(filters.pageSize)||50)),start=(page-1)*pageSize;
 return {year:year,months:months,metric:metric,rows:all.slice(start,start+pageSize),totalRows:all.length,page:page,pageSize:pageSize,
  totals:totals,monthTotals:monthTotals,matchedSourceRows:matchedSourceRows,source:sourceName||'P2_DASHBOARD_FACT'};
}
function getPage3InitialPayload(){
 const started=Date.now(),src=page3ReadSource_(),meta=page3BuildMeta_(src.rows);
 const year=meta.years.includes('2026')?'2026':meta.years[meta.years.length-1],months=meta.monthsByYear[year]||[];
 const data=page3BuildData_(src.rows,{year:year,months:months,page:1,pageSize:50,metric:'value'},meta,src.sheet.getName());
 return {meta:meta,data:data,diagnostics:{source:src.sheet.getName(),readAndAggregateMs:Date.now()-started}};
}
function getPage3OutletData(filters){
 const started=Date.now(),src=page3ReadSource_(filters),meta=page3BuildMeta_(src.rows);
 const data=page3BuildData_(src.rows,filters,meta,src.sheet.getName());
 data.diagnostics={elapsedMs:Date.now()-started};return data;
}