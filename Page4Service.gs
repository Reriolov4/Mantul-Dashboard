/**
 * PAGE 4 — Sales by Brand/SKU. Reads P34_EVENT_FACT (fallback P2_EVENT_INDEX); SO and SOO are separated. Historical quantity/stock metrics stay blank until valid monthly unit data exists.
 */
const PAGE4_MONTHS_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function page4HeaderIndex_(headers,aliases,label){
 const normalized=headers.map(page2HeaderKey_);for(let i=0;i<aliases.length;i++){const at=normalized.indexOf(page2HeaderKey_(aliases[i]));if(at>=0)return at;}
 throw new Error('Kolom '+label+' tidak ditemukan. Header: '+headers.join(' | '));
}
function page4ReadSelectedColumns_(sheet,headerMap){
 const n=sheet.getLastRow()-1,keys=Object.keys(headerMap),indices=keys.map(k=>headerMap[k]),groups=[];
 indices.slice().sort((a,b)=>a-b).forEach(col=>{const last=groups[groups.length-1];if(last&&col<=last.end+1)last.end=col;else groups.push({start:col,end:col,values:null});});
 groups.forEach(g=>{g.values=sheet.getRange(2,g.start+1,n,g.end-g.start+1).getValues();});
 const result={};keys.forEach(key=>{const col=headerMap[key],g=groups.find(x=>col>=x.start&&col<=x.end);result[key]=g.values.map(row=>row[col-g.start]);});return {rows:n,columns:result};
}
function page4ReadSource_() {
  const ss=getDatabase_(),sku=ss.getSheetByName(SKU_SHEET_NAME);
  const event=page34ReadEventIndex_();
  if(!sku||sku.getLastRow()<2)throw new Error('Sheet SKU belum tersedia atau belum berisi master produk.');
  const rows=event.rows,n=rows.length;
  const columns={
    year:rows.map(r=>r[0]),month:rows.map(r=>r[1]),dd:rows.map(r=>r[2]),customer:rows.map(r=>r[4]),pic:rows.map(r=>r[6]),
    subdist:rows.map(r=>r[5]),subDist:rows.map(r=>r[17]),brand:rows.map(r=>r[7]),subbrand:rows.map(r=>r[8]),
    category:rows.map(r=>r[9]),sbu:rows.map(r=>r[10]),salesType:rows.map(r=>r[11]),
    value:rows.map(r=>r[14]),sku:rows.map(r=>r[16]),
    sourceRows:rows.map(r=>Number(r[18])||1)
  };
  const mh=sku.getRange(1,1,1,sku.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const mm={
    sku:page4HeaderIndex_(mh,['SKU'],'SKU master SKU'),
    subbrand:page4HeaderIndex_(mh,['SUB_BRAND','SUB BRAND','SUBBRAND'],'SKU master SUB_BRAND'),
    brand:page4HeaderIndex_(mh,['BRAND'],'SKU master BRAND'),
    category:page4HeaderIndex_(mh,['CATEGORY'],'SKU master CATEGORY'),
    entity:null
  };
  delete mm.entity;
  const masterData=page4ReadSelectedColumns_(sku,mm);
  return {sourceRows:n,columns:columns,master:masterData.columns,masterRows:masterData.rows,sourceName:event.source};
}
function page4BuildMeta_(src){
 const years=new Set(),monthsByYear={},skuAttrs=new Map(),customers=new Set(),pics=new Set(),brands=new Set(),subbrands=new Set(),categories=new Set(),dateKeys=new Set();let sooRows=0,soModeRows=0,sourceRows=0;const c=src.columns;
 const add=(map,key,field,value)=>{if(!key||value==null||String(value).trim()==='')return;let obj=map.get(key);if(!obj){obj={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};map.set(key,obj);}obj[field].add(String(value).trim());};
 for(let i=0;i<src.sourceRows;i++){const type=page2Key_(c.salesType[i]),sourceCount=Number(c.sourceRows[i])||1;sourceRows+=sourceCount;soModeRows+=sourceCount;if(type!=='SOO')continue;sooRows+=sourceCount;
  const y=String(Number(c.year[i])||c.year[i]||'').trim(),m=page2Month_(c.month[i]),name=page2Text_(c.sku[i]);if(page2Text_(c.customer[i]))customers.add(page2Text_(c.customer[i]));if(page2Text_(c.pic[i]))pics.add(page2Text_(c.pic[i]));if(page2Text_(c.brand[i]))brands.add(page2Text_(c.brand[i]));if(page2Text_(c.subbrand[i]))subbrands.add(page2Text_(c.subbrand[i]));if(page2Text_(c.category[i]))categories.add(page2Text_(c.category[i]));const dateKey=page2Date_(y,m,c.dd[i]);if(dateKey)dateKeys.add(dateKey);if(y)years.add(y);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
  add(skuAttrs,name,'subbrand',c.subbrand[i]);add(skuAttrs,name,'brand',c.brand[i]);add(skuAttrs,name,'category',c.category[i]);add(skuAttrs,name,'entity',c.sbu[i]);
 }
 const masterMap=new Map();for(let i=0;i<src.masterRows;i++){const name=page2Text_(src.master.sku[i]);if(!name)continue;let x=masterMap.get(name);if(!x){x={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};masterMap.set(name,x);}['subbrand','brand','category','entity'].forEach(k=>{const v=src.master[k][i];if(v!=null&&String(v).trim())x[k].add(String(v).trim());});}
 const choose=(set,fallback)=>set&&set.size===1?Array.from(set)[0]:set&&set.size>1?'MULTI':fallback||'',names=new Set([...masterMap.keys(),...skuAttrs.keys()]);
 const skuMaster=Array.from(names).map(name=>{const raw=skuAttrs.get(name)||{},master=masterMap.get(name)||{};return {sku:name,subbrand:choose(raw.subbrand,choose(master.subbrand,'')),brand:choose(raw.brand,choose(master.brand,'')),category:choose(raw.category,choose(master.category,'')),entity:choose(raw.entity,choose(master.entity,''))};}).sort((a,b)=>a.sku.localeCompare(b.sku,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE4_MONTHS_.filter(m=>monthsByYear[y].has(m))])),subdists:PAGE34_ENTITIES_.slice(),customers:Array.from(customers).sort((a,b)=>a.localeCompare(b,'id')),pics:Array.from(pics).sort((a,b)=>a.localeCompare(b,'id')),brands:Array.from(brands).sort((a,b)=>a.localeCompare(b,'id')),subbrands:Array.from(subbrands).sort((a,b)=>a.localeCompare(b,'id')),categories:Array.from(categories).sort((a,b)=>a.localeCompare(b,'id')),dateKeys:Array.from(dateKeys).sort(),skuMaster:skuMaster,skuMasterCount:skuMaster.length,sourceRows:sourceRows,sooRows:sooRows,soModeRows:soModeRows};
}
function page4BuildData_(src,meta,filters){
 filters=filters||{};
 const years=(meta.years||[]).map(Number).filter(Number.isFinite);
 const year=Number(filters.year)||(years.includes(2026)?2026:(years[years.length-1]||new Date().getFullYear()));
 const available=(meta.monthsByYear||{})[String(year)]||[];
 const requested=(Array.isArray(filters.months)?filters.months:[]).map(page2Month_).filter(m=>PAGE4_MONTHS_.includes(m));
 const monthsUsed=requested.length?PAGE4_MONTHS_.filter(m=>requested.includes(m)&&available.includes(m)):available.slice();
 const monthSet=new Set(monthsUsed),selected=new Set(page2FilterList_(filters.subdists).map(page2Key_)),customers=new Set(page2FilterList_(filters.customers).map(page2Key_)),pics=new Set(page2FilterList_(filters.pics).map(page2Key_)),brands=new Set(page2FilterList_(filters.brands).map(page2Key_)),subbrands=new Set(page2FilterList_(filters.subbrands).map(page2Key_)),categories=new Set(page2FilterList_(filters.categories).map(page2Key_)),dates=new Set(page2FilterList_(filters.dates).map(v=>String(v).padStart(2,'0')));
 const map=new Map(meta.skuMaster.map(s=>[s.sku,{...s,total:{value:0,carton:0}}]));
 let matchedSourceRows=0;const c=src.columns;
 for(let i=0;i<src.sourceRows;i++){
  if(page2Key_(c.salesType[i])!=='SOO'||Number(c.year[i])!==year||!monthSet.has(page2Month_(c.month[i])))continue;
  const row=[null,null,null,null,null,c.subdist[i],null,null,null,null,null,c.salesType[i],null,null,null,null,c.sku[i]];
  const entity=page34Entity_(row,'SOO'),customer=page2Text_(c.customer[i]),day=String(Number(c.dd[i])||c.dd[i]||'').padStart(2,'0');
  if(!entity||(selected.size&&!selected.has(page2Key_(entity)))||(customers.size&&!customers.has(page2Key_(customer)))||(pics.size&&!pics.has(page2Key_(c.pic[i])))||(brands.size&&!brands.has(page2Key_(c.brand[i])))||(subbrands.size&&!subbrands.has(page2Key_(c.subbrand[i])))||(categories.size&&!categories.has(page2Key_(c.category[i])))||(dates.size&&!dates.has(day)))continue;
  const name=page2Text_(c.sku[i]);if(!name)continue;
  matchedSourceRows+=Number(c.sourceRows[i])||1;
  let x=map.get(name);
  if(!x){x={sku:name,brand:page2Text_(c.brand[i]),subbrand:page2Text_(c.subbrand[i]),category:page2Text_(c.category[i]),entity:page2Text_(c.sbu[i]),total:{value:0,carton:0}};map.set(name,x);}
  x.total.value+=Number(c.value[i])||0;x.total.carton+=Number(c.carton[i])||0;
 }
 const all=Array.from(map.values()).sort((a,b)=>(b.total.value-a.total.value)||a.sku.localeCompare(b.sku,'id'));
 const totals=all.reduce((a,r)=>({value:a.value+r.total.value,carton:a.carton+r.total.carton}),{value:0,carton:0});
 return {mode:'SOO',modeSalesType:'SOO',label:'SUBDIST',year:year,monthsUsed:monthsUsed,subdist:selected.size?Array.from(selected).join(', '):'Semua SUBDIST',rows:all,skuCount:all.length,totals:totals,matchedSourceRows:matchedSourceRows,sourceRows:meta.sourceRows};
}
function getPage4InitialPayload(mode){
 const started=Date.now(),readStarted=Date.now(),src=page4ReadSource_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page4BuildMeta_(src),metaMs=Date.now()-metaStarted;
 const selectedMode=page34Mode_(mode),year=meta.years.includes('2026')?2026:Number(meta.years[meta.years.length-1]);
 const aggregateStarted=Date.now(),data=page4BuildData_(src,meta,{mode:selectedMode,year:year}),aggregateMs=Date.now()-aggregateStarted;
 return {meta:meta,data:data,diagnostics:{source:src.sourceName,sourceRows:src.sourceRows,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started}};
}
function getPage4SkuData(filters){
 const started=Date.now(),readStarted=Date.now(),src=page4ReadSource_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page4BuildMeta_(src),metaMs=Date.now()-metaStarted;
 const aggregateStarted=Date.now(),data=page4BuildData_(src,meta,filters),aggregateMs=Date.now()-aggregateStarted;
 data.diagnostics={source:src.sourceName,sourceRows:src.sourceRows,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started};return data;
}
