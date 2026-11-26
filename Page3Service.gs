/**
 * PAGE 3 — OUTLETS. SOO uses SUBDIST; SO uses SUB_DIST and all sales types.
 */
const PAGE3_MONTHS_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const PAGE34_ENTITIES_=['EPM BAC','EPM LSE','GMP BAC','GMP LSE','GMP LGS','DB MBO','DB BAC'];
function page34ReadEventIndex_(){
 if(PropertiesService.getScriptProperties().getProperty('MANTUL_V2_DATA_STALE')==='TRUE')throw new Error('Database V2 belum diperbarui setelah import. Selesaikan build, finalize, dan publish terlebih dahulu.');
 const ss=getDatabase_(),p3=ss.getSheetByName('P3_OUTLET_MONTHLY');
 const p3Headers=['YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','SUB_DIST','SALES_TYPE','PIC','BRAND','SUBBRAND','CATEGORY','SBU','ITEM_CODE','SKU','VAL','CARTON','SOURCE_ROWS'];
 if(p3&&p3.getLastRow()>1&&p3.getLastColumn()>=18){
  const actual=p3.getRange(1,1,1,18).getDisplayValues()[0].map(page2HeaderKey_);
  if(p3Headers.every((v,i)=>page2HeaderKey_(v)===actual[i])){
   const values=p3.getRange(2,1,p3.getLastRow()-1,18).getValues();
   const rows=values.map(r=>[r[0],r[1],r[2],r[3],r[4],r[5],r[8],r[9],r[10],r[11],r[12],r[7],'', '',r[15],r[16],r[14],r[6],r[17]]);
   return {sheet:p3,rows:rows,source:'P3_OUTLET_MONTHLY'};
  }
 }
 const sh=ss.getSheetByName('P2_EVENT_INDEX');
 if(!sh||sh.getLastRow()<2)throw new Error('P3_OUTLET_MONTHLY V2 atau P2_EVENT_INDEX belum berisi data.');
 const expected=(typeof P2_EVENT_INDEX_HEADERS_!=='undefined'?P2_EVENT_INDEX_HEADERS_:P2_EVENT_HEADERS_.concat(['SKU','SUB_DIST']));
 const actual=sh.getRange(1,1,1,18).getDisplayValues()[0].map(page2HeaderKey_);
 if(!expected.every((v,i)=>page2HeaderKey_(v)===actual[i]))throw new Error('Header P2_EVENT_INDEX harus 18 kolom termasuk SKU dan SUB_DIST. Terbaca: '+actual.join(' | '));
 return {sheet:sh,rows:sh.getRange(2,1,sh.getLastRow()-1,18).getValues(),source:'P3_OUTLET_MONTHLY'};
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
 const years=new Set(),subbrands=new Set(),monthsByYear={},pics=new Set(),brands=new Set(),categories=new Set(),customers=new Set(),dateKeys=new Set();let soo=0,soModeRows=0;
 (rows||[]).forEach(r=>{
  const type=page2Key_(r[11]);if(type==='SOO')soo++;soModeRows++;
  const y=String(Number(r[0])||r[0]||'').trim(),m=page2Month_(r[1]),sb=page2Text_(r[8]);
  if(y)years.add(y);if(sb)subbrands.add(sb);if(page2Text_(r[6]))pics.add(page2Text_(r[6]));if(page2Text_(r[7]))brands.add(page2Text_(r[7]));if(page2Text_(r[9]))categories.add(page2Text_(r[9]));if(page2Text_(r[4]))customers.add(page2Text_(r[4]));const dateKey=page2Date_(y,m,r[2]);if(dateKey)dateKeys.add(dateKey);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
 });
 const sort=a=>Array.from(a).sort((x,y)=>x.localeCompare(y,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE3_MONTHS_.filter(m=>monthsByYear[y].has(m))])),
  subdists:PAGE34_ENTITIES_.slice(),subbrands:sort(subbrands),pics:sort(pics),brands:sort(brands),categories:sort(categories),customers:sort(customers),dateKeys:Array.from(dateKeys).sort(),sourceRows:(rows||[]).length,sooRows:soo,soModeRows:soModeRows};
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
 return {mode:mode,modeSalesType:mode==='SOO'?'SOO':'Semua',label:'SUBDIST',year:year,months:months,metric:metric,rows:all.slice(start,start+pageSize),totalRows:all.length,page:page,pageSize:pageSize,totals:totals,monthTotals:monthTotals,matchedSourceRows:matchedSourceRows,source:'P2_EVENT_INDEX'};
}
function getPage3InitialPayload(mode){
 if(typeof analyticsEnsureExistingIndexesRegistered_==='function')analyticsEnsureExistingIndexesRegistered_();
 const started=Date.now(),src=page34ReadEventIndex_(),meta=page3BuildMeta_(src.rows),selectedMode=page34Mode_(mode),year=meta.years.includes('2026')?'2026':meta.years[meta.years.length-1],months=meta.monthsByYear[year]||[];
 return {meta:meta,data:page3BuildData_(src.rows,{mode:selectedMode,year:year,months:months,page:1,pageSize:50,metric:'value'},meta),diagnostics:{source:src.sheet.getName(),readAndAggregateMs:Date.now()-started}};
}
function getPage3OutletData(filters){const started=Date.now(),src=page34ReadEventIndex_(),meta=page3BuildMeta_(src.rows),data=page3BuildData_(src.rows,filters,meta);data.diagnostics={elapsedMs:Date.now()-started};return data;}
