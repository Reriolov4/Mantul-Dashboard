/**
 * PAGE 3 — OUTLETS. Reads validated P34_EVENT_FACT (fallback P2_EVENT_INDEX); SOO uses SUBDIST, SO uses SUB_DIST.
 */
const PAGE3_MONTHS_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const PAGE34_ENTITIES_=['EPM BAC','EPM LSE','GMP BAC','GMP LSE','GMP LGS','DB MBO','DB BAC'];
function page34ReadRawIndex_(startRow, requestedRows) {
  const ss = getDatabase_();
  const sh = mantulGetReadyReadIndex_();
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('P2_EVENT_INDEX belum siap atau kedaluwarsa. Jalankan refreshMantulDashboard() terlebih dahulu.');
  }
  const actual = sh.getRange(1, 1, 1, RAW_SALES_COLUMN_COUNT)
    .getDisplayValues()[0].map(page2HeaderKey_);
  const expected = RAW_SALES_HEADERS.map(page2HeaderKey_);
  if (!expected.every((value, index) => actual[index] === value)) {
    throw new Error('Header RAW_DATA_SALES tidak sesuai. Terbaca: ' + actual.join(' | '));
  }

  const firstRow = Math.max(2, Number(startRow) || 2);
  if (sh.getLastRow() < firstRow) return {sheet:sh,rows:[],source:'P2_EVENT_INDEX'};
  const rowCount = Math.min(sh.getLastRow() - firstRow + 1, Number(requestedRows) > 0 ? Number(requestedRows) : sh.getLastRow() - firstRow + 1);
  // Baca hanya 20 kolom yang dipakai Page 3, bukan seluruh 25 kolom mentah.
  const blocks = [
    {start:0,end:2,values:sh.getRange(firstRow,1,rowCount,3).getValues()},
    {start:5,end:10,values:sh.getRange(firstRow,6,rowCount,6).getValues()},
    {start:12,end:21,values:sh.getRange(firstRow,13,rowCount,10).getValues()},
    {start:24,end:24,values:sh.getRange(firstRow,25,rowCount,1).getValues()}
  ];
  const readers = {};
  blocks.forEach(block => {
    for (let col=block.start; col<=block.end; col++) {
      readers[col] = {values:block.values,offset:col-block.start};
    }
  });
  const read = (i,col) => {
    const source=readers[col];
    return source ? source.values[i][source.offset] : '';
  };
  const rows = new Array(rowCount);
  for (let i=0; i<rowCount; i++) {
    rows[i] = [
      read(i,0),read(i,1),read(i,2),read(i,9),read(i,10),read(i,8),
      read(i,16),read(i,6),read(i,5),read(i,19),read(i,21),read(i,14),
      read(i,18),read(i,17),read(i,12),read(i,7),read(i,24),read(i,15),1
    ];
  }
  return {sheet:sh, rows:rows, source:'P2_EVENT_INDEX'};
}

/**
 * PAGE 3/4 SHARED FACT — grain transaksi dengan dimensi filter lengkap.
 * Kolom berurutan sesuai row adapter Page 3/4; SOURCE ROWS menjaga hitungan asal.
 */
const PAGE34_FACT_SHEET_='P34_EVENT_FACT';
const PAGE34_FACT_STAGE_='P34_EVENT_FACT_BUILD';
const PAGE34_FACT_HEADERS_=[
 'YEAR','MONTH','DD','CUSTOMER CODE','CUSTOMER NAME','SUBDIST','PIC','BRAND',
 'SUBBRAND','CATEGORY','SBU','SALES TYPE','SALESMAN','TEAM','VAL','CARTON',
 'SKU','SUB_DIST','SOURCE ROWS'
];
function page34FactReady_(ss,index){
 const props=PropertiesService.getScriptProperties(),fact=ss.getSheetByName(PAGE34_FACT_SHEET_);
 if(!fact||fact.getLastRow()<2||props.getProperty('PAGE34_FACT_STATUS')!=='COMPLETE'||
  props.getProperty('PAGE34_FACT_SOURCE_ID')!==ss.getId()||
  Number(props.getProperty('PAGE34_FACT_SOURCE_ROWS'))!==index.getLastRow()-1||
  Number(props.getProperty('PAGE34_FACT_WRITTEN'))!==fact.getLastRow()-1)return null;
 const actual=fact.getRange(1,1,1,PAGE34_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_);
 if(!PAGE34_FACT_HEADERS_.every((v,i)=>actual[i]===page2HeaderKey_(v)))return null;
 return fact;
}
function page34ReadEventIndex_(){
 const ss=getDatabase_(),index=mantulGetReadyReadIndex_();
 if(!index||index.getLastRow()<2)throw new Error('P2_EVENT_INDEX belum siap atau kedaluwarsa. Jalankan refreshMantulDashboard() terlebih dahulu.');
 const fact=page34FactReady_(ss,index);
 if(fact){
  const values=fact.getRange(2,1,fact.getLastRow()-1,PAGE34_FACT_HEADERS_.length).getValues();
  return {sheet:fact,rows:values.map(r=>r.slice(0,18).concat([Number(r[18])||1])),source:PAGE34_FACT_SHEET_};
 }
 return page34ReadRawIndex_();
}
/** Full rebuild terkontrol; bila index bertambah, fact lama otomatis dianggap stale. */
function buildPage34EventFact(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);let stage=null;
 try{
  const ss=getDatabase_(),src=page34ReadRawIndex_(),grouped=new Map();let sourceRows=0,sourceValue=0,sourceCarton=0;
  src.rows.forEach(r=>{
   const key=JSON.stringify(r.slice(0,14).concat(r.slice(16,18)));let item=grouped.get(key);
   if(!item){item=r.slice();item[14]=0;item[15]=0;item[18]=0;grouped.set(key,item);}
   item[18]+=Number(r[18])||1;item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||0;
   sourceRows+=Number(r[18])||1;sourceValue+=Number(r[14])||0;sourceCarton+=Number(r[15])||0;
  });
  const rows=Array.from(grouped.values()),oldStage=ss.getSheetByName(PAGE34_FACT_STAGE_);
  if(oldStage)ss.deleteSheet(oldStage);
  stage=ss.insertSheet(PAGE34_FACT_STAGE_);
  stage.getRange(1,1,1,PAGE34_FACT_HEADERS_.length).setValues([PAGE34_FACT_HEADERS_]);
  const batch=3000;
  for(let i=0;i<rows.length;i+=batch){const part=rows.slice(i,i+batch);stage.getRange(i+2,1,part.length,PAGE34_FACT_HEADERS_.length).setValues(part);}
  SpreadsheetApp.flush();
  if(stage.getLastRow()!==rows.length+1||stage.getLastColumn()!==PAGE34_FACT_HEADERS_.length)throw new Error('Ukuran P34_EVENT_FACT tidak valid.');
  let factRows=0,factValue=0,factCarton=0;
  for(let i=0;i<rows.length;i+=batch){const part=stage.getRange(i+2,1,Math.min(batch,rows.length-i),PAGE34_FACT_HEADERS_.length).getValues();part.forEach(r=>{factRows+=Number(r[18])||0;factValue+=Number(r[14])||0;factCarton+=Number(r[15])||0;});}
  const close=(a,b)=>Math.abs(a-b)<=Math.max(0.01,Math.abs(a)*1e-10);
  if(factRows!==sourceRows||!close(factValue,sourceValue)||!close(factCarton,sourceCarton))throw new Error('Validasi P34_EVENT_FACT gagal: SOURCE ROWS/VAL/CARTON tidak cocok.');
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'),current=ss.getSheetByName(PAGE34_FACT_SHEET_),backup='P34_EVENT_FACT_OLD_'+stamp;
  if(current)current.setName(backup);
  try{
   stage.setName(PAGE34_FACT_SHEET_);stage.setFrozenRows(1);
   PropertiesService.getScriptProperties().setProperties({
    PAGE34_FACT_STATUS:'COMPLETE',PAGE34_FACT_SOURCE_ID:ss.getId(),
    PAGE34_FACT_SOURCE_ROWS:String(src.rows.length),PAGE34_FACT_WRITTEN:String(rows.length),
    PAGE34_FACT_SOURCE_COUNT:String(sourceRows),PAGE34_FACT_VAL:String(factValue),
    PAGE34_FACT_CARTON:String(factCarton),PAGE34_FACT_BUILT_AT:new Date().toISOString()
   });
   if(current)ss.deleteSheet(ss.getSheetByName(backup));
  }catch(e){
   const published=ss.getSheetByName(PAGE34_FACT_SHEET_);if(published&&published.getSheetId()===stage.getSheetId())ss.deleteSheet(published);
   const prev=ss.getSheetByName(backup);if(prev&&!ss.getSheetByName(PAGE34_FACT_SHEET_))prev.setName(PAGE34_FACT_SHEET_);throw e;
  }
  const result={ok:true,status:'COMPLETE',sheet:PAGE34_FACT_SHEET_,sourceRows:sourceRows,factRows:rows.length,
   compressedRows:sourceRows-rows.length,sourceValue:sourceValue,factValue:factValue,
   sourceCarton:sourceCarton,factCarton:factCarton,validated:true};
  Logger.log('PAGE34 EVENT FACT BUILD\\n'+JSON.stringify(result,null,2));return result;
 }catch(err){
  if(stage){const sh=stage.getParent().getSheetByName(PAGE34_FACT_STAGE_);if(sh&&sh.getSheetId()===stage.getSheetId())stage.getParent().deleteSheet(sh);}
  PropertiesService.getScriptProperties().setProperty('PAGE34_FACT_STATUS','FAILED');throw err;
 }finally{lock.releaseLock();}
}

function refreshPage34EventFactIncremental_(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 let fact=null,appendStart=0,appendCount=0;const props=PropertiesService.getScriptProperties(),oldStatus=props.getProperty('PAGE34_FACT_STATUS')||'';
 try{
  const ss=getDatabase_(),index=mantulGetReadyReadIndex_();if(!index)throw new Error('P2_EVENT_INDEX belum siap.');
  const sourceRows=index.getLastRow()-1;fact=ss.getSheetByName(PAGE34_FACT_SHEET_);
  const previousRows=Number(props.getProperty('PAGE34_FACT_SOURCE_ROWS')||0),previousWritten=Number(props.getProperty('PAGE34_FACT_WRITTEN')||0);
  const ready=fact&&fact.getLastRow()>1&&oldStatus==='COMPLETE'&&props.getProperty('PAGE34_FACT_SOURCE_ID')===ss.getId()&&previousRows>0&&previousRows<=sourceRows&&fact.getLastRow()-1===previousWritten;
  if(!ready)return {ok:true,status:'NEEDS_FULL_BUILD',sheet:PAGE34_FACT_SHEET_,sourceRows:sourceRows};
  const actual=fact.getRange(1,1,1,PAGE34_FACT_HEADERS_.length).getDisplayValues()[0].map(page2HeaderKey_);
  if(!PAGE34_FACT_HEADERS_.every((v,i)=>actual[i]===page2HeaderKey_(v)))return {ok:true,status:'NEEDS_FULL_BUILD',reason:'HEADER_MISMATCH'};
  if(sourceRows===previousRows)return {ok:true,status:'UNCHANGED',sheet:PAGE34_FACT_SHEET_,sourceRows:sourceRows};
  const incoming=page34ReadRawIndex_(previousRows+2,sourceRows-previousRows),grouped=new Map();let sourceCount=0,sourceValue=0,sourceCarton=0;
  incoming.rows.forEach(r=>{const key=JSON.stringify(r.slice(0,14).concat(r.slice(16,18)));let item=grouped.get(key);if(!item){item=r.slice();item[14]=0;item[15]=0;item[18]=0;grouped.set(key,item);}item[18]+=Number(r[18])||1;item[14]+=Number(r[14])||0;item[15]+=Number(r[15])||0;sourceCount+=Number(r[18])||1;sourceValue+=Number(r[14])||0;sourceCarton+=Number(r[15])||0;});
  const rows=Array.from(grouped.values());if(sourceCount!==sourceRows-previousRows)throw new Error('Jumlah baris baru P34 tidak cocok dengan index.');
  const oldCount=Number(props.getProperty('PAGE34_FACT_SOURCE_COUNT')||previousRows),oldValue=Number(props.getProperty('PAGE34_FACT_VAL')||0),oldCarton=Number(props.getProperty('PAGE34_FACT_CARTON')||0);
  appendStart=fact.getLastRow()+1;appendCount=rows.length;props.setProperty('PAGE34_FACT_STATUS','BUILDING');
  const batch=3000;for(let i=0;i<rows.length;i+=batch){const part=rows.slice(i,i+batch);fact.getRange(appendStart+i,1,part.length,PAGE34_FACT_HEADERS_.length).setValues(part);}
  SpreadsheetApp.flush();let checkCount=0,checkValue=0,checkCarton=0;
  for(let i=0;i<rows.length;i+=batch){const part=fact.getRange(appendStart+i,1,Math.min(batch,rows.length-i),PAGE34_FACT_HEADERS_.length).getValues();part.forEach(x=>{checkCount+=Number(x[18])||0;checkValue+=Number(x[14])||0;checkCarton+=Number(x[15])||0;});}
  const close=(a,b)=>Math.abs(a-b)<=Math.max(0.01,Math.abs(a)*1e-10);
  if(checkCount!==sourceCount||!close(checkValue,sourceValue)||!close(checkCarton,sourceCarton)||fact.getLastRow()!==appendStart+appendCount-1)throw new Error('Validasi append P34_EVENT_FACT gagal.');
  const totalCount=oldCount+checkCount,totalValue=oldValue+checkValue,totalCarton=oldCarton+checkCarton;
  props.setProperties({PAGE34_FACT_STATUS:'COMPLETE',PAGE34_FACT_SOURCE_ID:ss.getId(),PAGE34_FACT_SOURCE_ROWS:String(sourceRows),PAGE34_FACT_WRITTEN:String(fact.getLastRow()-1),PAGE34_FACT_SOURCE_COUNT:String(totalCount),PAGE34_FACT_VAL:String(totalValue),PAGE34_FACT_CARTON:String(totalCarton),PAGE34_FACT_BUILT_AT:new Date().toISOString()});
  return {ok:true,status:'APPENDED',sheet:PAGE34_FACT_SHEET_,appendedSourceRows:sourceCount,appendedFactRows:appendCount,sourceRows:sourceRows,factRows:fact.getLastRow()-1,sourceValue:totalValue,sourceCarton:totalCarton,validated:true};
 }catch(err){
  if(fact&&appendStart>0&&appendCount>0&&fact.getLastRow()>=appendStart)fact.getRange(appendStart,1,Math.min(appendCount,fact.getLastRow()-appendStart+1),PAGE34_FACT_HEADERS_.length).clearContent();
  if(oldStatus)props.setProperty('PAGE34_FACT_STATUS',oldStatus);else props.deleteProperty('PAGE34_FACT_STATUS');throw err;
 }finally{lock.releaseLock();}
}
function page34Mode_(v){return String(v||'SOO').trim().toUpperCase()==='SO'?'SO':'SOO';}
function page34Entity_(row,mode){
 const raw=page2Text_(row[mode==='SOO'?5:17]),key=page2Key_(raw);
 if(mode==='SOO')return PAGE34_ENTITIES_.includes(key)?key:'';
 if(!key||key==='0')return '';
 const normalized=key.replace(/[.,]/g,'').replace(/\s+/g,' ').trim();
 if(normalized==='AP GEUBRINA RIZKI')return 'EPM LSE';
 if(normalized==='PT RESTU INDO MEDICA')return 'EPM BAC';
 return PAGE34_ENTITIES_.includes(key)?key:'';
}
function page3BuildMeta_(rows){
 const years=new Set(),subbrands=new Set(),monthsByYear={},pics=new Set(),brands=new Set(),categories=new Set(),customers=new Set(),dateKeys=new Set();let soo=0,soModeRows=0,sourceRows=0;
 (rows||[]).forEach(r=>{
  const type=page2Key_(r[11]),sourceCount=Number(r[18])||1;sourceRows+=sourceCount;if(type==='SOO')soo+=sourceCount;soModeRows+=sourceCount;
  const y=String(Number(r[0])||r[0]||'').trim(),m=page2Month_(r[1]),sb=page2Text_(r[8]);
  if(y)years.add(y);if(sb)subbrands.add(sb);if(page2Text_(r[6]))pics.add(page2Text_(r[6]));if(page2Text_(r[7]))brands.add(page2Text_(r[7]));if(page2Text_(r[9]))categories.add(page2Text_(r[9]));if(page2Text_(r[4]))customers.add(page2Text_(r[4]));const dateKey=page2Date_(y,m,r[2]);if(dateKey)dateKeys.add(dateKey);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
 });
 const sort=a=>Array.from(a).sort((x,y)=>x.localeCompare(y,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE3_MONTHS_.filter(m=>monthsByYear[y].has(m))])),
  subdists:PAGE34_ENTITIES_.slice(),subbrands:sort(subbrands),pics:sort(pics),brands:sort(brands),categories:sort(categories),customers:sort(customers),dateKeys:Array.from(dateKeys).sort(),sourceRows:sourceRows,sooRows:soo,soModeRows:soModeRows};
}
function page3BuildData_(rows,filters,meta){
 filters=filters||{};const mode=page34Mode_(filters.mode),year=Number(filters.year)||Number((meta.years||[]).slice(-1)[0])||new Date().getFullYear();
 const available=(meta.monthsByYear||{})[String(year)]||[],requested=(Array.isArray(filters.months)?filters.months:[]).map(page2Month_).filter(m=>PAGE3_MONTHS_.includes(m));
 const months=requested.length?PAGE3_MONTHS_.filter(m=>requested.includes(m)):available.slice(),monthSet=new Set(months);
 const subdistSet=new Set(page2FilterList_(filters.subdists).map(page2Key_)),subbrandSet=new Set(page2FilterList_(filters.subbrands).map(page2Key_)),picSet=new Set(page2FilterList_(filters.pics).map(page2Key_)),brandSet=new Set(page2FilterList_(filters.brands).map(page2Key_)),categorySet=new Set(page2FilterList_(filters.categories).map(page2Key_)),customerSet=new Set(page2FilterList_(filters.customers).map(page2Key_)),dateSet=new Set(page2FilterList_(filters.dates).map(v=>String(v).padStart(2,'0')));
 const metric='value',map=new Map();let matchedSourceRows=0;
 (rows||[]).forEach(r=>{
  const type=page2Key_(r[11]);if(mode==='SOO'&&type!=='SOO')return;
  if(Number(r[0])!==year)return;
  const m=page2Month_(r[1]),sd=page34Entity_(r,mode),sb=page2Text_(r[8]),customer=page2Text_(r[4]);
  if(!monthSet.has(m)||!sd||!customer)return;const dateKey=page2Date_(r[0],r[1],r[2]),day=dateKey?dateKey.slice(8,10):'';if(dateSet.size&&!dateSet.has(day))return;if(subdistSet.size&&!subdistSet.has(page2Key_(sd)))return;
  if(subbrandSet.size&&!subbrandSet.has(page2Key_(sb)))return;if(picSet.size&&!picSet.has(page2Key_(r[6])))return;if(brandSet.size&&!brandSet.has(page2Key_(r[7])))return;if(categorySet.size&&!categorySet.has(page2Key_(r[9])))return;if(customerSet.size&&!customerSet.has(page2Key_(customer)))return;
  matchedSourceRows+=Number(r[18])||1;const key=page2ExactCompositeKey_([sd,customer]);let x=map.get(key);
  if(!x){x={customerName:customer,subdist:sd,monthly:{},total:{value:0,carton:0}};map.set(key,x);}
  if(!x.monthly[m])x.monthly[m]={value:0,carton:0};x.monthly[m].value+=Number(r[14])||0;x.monthly[m].carton+=Number(r[15])||0;x.total.value+=Number(r[14])||0;x.total.carton+=Number(r[15])||0;
 });
 const all=Array.from(map.values()).sort((a,b)=>(b.total[metric]-a.total[metric])||a.customerName.localeCompare(b.customerName,'id'));
 const monthTotals={};months.forEach(m=>monthTotals[m]={value:0,carton:0});const totals={value:0,carton:0};
 all.forEach(r=>{totals.value+=r.total.value;totals.carton+=r.total.carton;months.forEach(m=>{monthTotals[m].value+=r.monthly[m]?.value||0;monthTotals[m].carton+=r.monthly[m]?.carton||0;});});
 const page=Math.max(1,Number(filters.page)||1),pageSize=Math.max(10,Math.min(100,Number(filters.pageSize)||50)),start=(page-1)*pageSize;
 return {mode:mode,modeSalesType:mode==='SOO'?'SOO':'Semua',label:'SUBDIST',year:year,months:months,metric:metric,rows:all.slice(start,start+pageSize),totalRows:all.length,page:page,pageSize:pageSize,totals:totals,monthTotals:monthTotals,matchedSourceRows:matchedSourceRows,source:'P34_EVENT_FACT'};
}
function getPage3InitialPayload(mode){
 const started=Date.now(),readStarted=Date.now(),src=page34ReadEventIndex_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page3BuildMeta_(src.rows),metaMs=Date.now()-metaStarted;
 const selectedMode=page34Mode_(mode),year=meta.years.includes('2026')?'2026':meta.years[meta.years.length-1],months=meta.monthsByYear[year]||[];
 const aggregateStarted=Date.now(),data=page3BuildData_(src.rows,{mode:selectedMode,year:year,months:months,page:1,pageSize:50,metric:'value'},meta),aggregateMs=Date.now()-aggregateStarted;
 return {meta:meta,data:data,diagnostics:{source:src.sheet.getName(),sourceRows:src.rows.length,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started}};
}
function getPage3OutletData(filters){
 const started=Date.now(),readStarted=Date.now(),src=page34ReadEventIndex_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page3BuildMeta_(src.rows),metaMs=Date.now()-metaStarted;
 const aggregateStarted=Date.now(),data=page3BuildData_(src.rows,filters,meta),aggregateMs=Date.now()-aggregateStarted;
 data.diagnostics={source:src.sheet.getName(),sourceRows:src.rows.length,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started};return data;
}
