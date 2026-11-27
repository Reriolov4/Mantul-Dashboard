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
  masterData.columns.entity=Array(masterData.rows).fill(''); // kolom ENTITAS/O sengaja tidak dipakai Page 4
  return {sourceRows:n,columns:columns,master:masterData.columns,masterRows:masterData.rows,sourceName:event.source};
}
function page4BuildMeta_(src){
 const years=new Set(),monthsByYear={},skuAttrs=new Map(),customers=new Set(),pics=new Set(),brands=new Set(),subbrands=new Set(),categories=new Set(),dateKeys=new Set();let sooRows=0,soModeRows=0,sourceRows=0;const c=src.columns;
 const add=(map,key,field,value)=>{if(!key||value==null||String(value).trim()==='')return;let obj=map.get(key);if(!obj){obj={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};map.set(key,obj);}obj[field].add(String(value).trim());};
 for(let i=0;i<src.sourceRows;i++){const type=page2Key_(c.salesType[i]),sourceCount=Number(c.sourceRows[i])||1;sourceRows+=sourceCount;soModeRows+=sourceCount;
  const y=String(Number(c.year[i])||c.year[i]||'').trim(),m=page2Month_(c.month[i]),name=page2Text_(c.sku[i]);if(type==='SOO')sooRows+=sourceCount;if(page2Text_(c.customer[i]))customers.add(page2Text_(c.customer[i]));if(page2Text_(c.pic[i]))pics.add(page2Text_(c.pic[i]));if(page2Text_(c.brand[i]))brands.add(page2Text_(c.brand[i]));if(page2Text_(c.subbrand[i]))subbrands.add(page2Text_(c.subbrand[i]));if(page2Text_(c.category[i]))categories.add(page2Text_(c.category[i]));const dateKey=page2Date_(y,m,c.dd[i]);if(dateKey)dateKeys.add(dateKey);if(y)years.add(y);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
  add(skuAttrs,name,'subbrand',c.subbrand[i]);add(skuAttrs,name,'brand',c.brand[i]);add(skuAttrs,name,'category',c.category[i]);add(skuAttrs,name,'entity',c.sbu[i]);
 }
 const masterMap=new Map();for(let i=0;i<src.masterRows;i++){const name=page2Text_(src.master.sku[i]);if(!name)continue;let x=masterMap.get(name);if(!x){x={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};masterMap.set(name,x);}['subbrand','brand','category','entity'].forEach(k=>{const v=src.master[k][i];if(v!=null&&String(v).trim())x[k].add(String(v).trim());});}
 const choose=(set,fallback)=>set&&set.size===1?Array.from(set)[0]:set&&set.size>1?'MULTI':fallback||'',names=new Set([...masterMap.keys(),...skuAttrs.keys()]);
 const skuMaster=Array.from(names).map(name=>{const raw=skuAttrs.get(name)||{},master=masterMap.get(name)||{};return {sku:name,subbrand:choose(raw.subbrand,choose(master.subbrand,'')),brand:choose(raw.brand,choose(master.brand,'')),category:choose(raw.category,choose(master.category,'')),entity:choose(raw.entity,choose(master.entity,''))};}).sort((a,b)=>a.sku.localeCompare(b.sku,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE4_MONTHS_.filter(m=>monthsByYear[y].has(m))])),subdists:PAGE34_ENTITIES_.slice(),customers:Array.from(customers).sort((a,b)=>a.localeCompare(b,'id')),pics:Array.from(pics).sort((a,b)=>a.localeCompare(b,'id')),brands:Array.from(brands).sort((a,b)=>a.localeCompare(b,'id')),subbrands:Array.from(subbrands).sort((a,b)=>a.localeCompare(b,'id')),categories:Array.from(categories).sort((a,b)=>a.localeCompare(b,'id')),dateKeys:Array.from(dateKeys).sort(),skuMaster:skuMaster,skuMasterCount:skuMaster.length,sourceRows:sourceRows,sooRows:sooRows,soModeRows:soModeRows};
}
function page4BuildData_(src,meta,filters){
 filters=filters||{};
 const mode=String(filters.mode||'SOO').toUpperCase()==='SO'?'SO':'SOO';
 const years=(meta.years||[]).map(Number).filter(Number.isFinite);
 const year=Number(filters.year)||(years.includes(2026)?2026:(years[years.length-1]||new Date().getFullYear()));
 const now=new Date(),currentYear=now.getFullYear(),currentMonth=now.getMonth()+1;
 const visibleMonthCount=year<currentYear?12:(year===currentYear?currentMonth:0);
 const monthsUsed=PAGE4_MONTHS_.slice(0,visibleMonthCount);
 const monthSet=new Set(monthsUsed),selected=new Set(page2FilterList_(filters.subdists).map(page2Key_)),dates=new Set(page2FilterList_(filters.dates).map(v=>String(v).padStart(2,'0'))),customers=new Set(page2FilterList_(filters.customers).map(page2Key_)),pics=new Set(page2FilterList_(filters.pics).map(page2Key_)),brands=new Set(page2FilterList_(filters.brands).map(page2Key_)),subbrands=new Set(page2FilterList_(filters.subbrands).map(page2Key_)),categories=new Set(page2FilterList_(filters.categories).map(page2Key_));
 const master=new Map((meta.skuMaster||[]).map(x=>[page2Key_(x.sku),x]));
 const map=new Map();let matchedSourceRows=0,totalValue=0;const c=src.columns;
 for(let i=0;i<src.sourceRows;i++){
  const type=page2Key_(c.salesType[i]);
  if(mode==='SOO'&&type!=='SOO')continue;
  if(Number(c.year[i])!==year||!monthSet.has(page2Month_(c.month[i]))||(dates.size&&!dates.has(String(Number(c.dd[i])||c.dd[i]||'').padStart(2,'0'))))continue;
  const adapter=[null,null,null,null,c.customer[i],c.subdist[i],null,null,null,null,null,c.salesType[i],null,null,null,null,null,c.subDist[i]];
  const entity=page34Entity_(adapter,mode),customer=page2Text_(c.customer[i]),name=page2Text_(c.sku[i]);
  // Page 4 SO table follows the top pivot only: ALL sales types, EPM BAC + EPM LSE. This is isolated to this table and does not alter Page 2/3 or SOO.
  if(mode==='SO'&&!['EPM BAC','EPM LSE'].includes(page2Key_(entity)))continue;
  if(!entity||!name||(selected.size&&!selected.has(page2Key_(entity)))||(customers.size&&!customers.has(page2Key_(customer)))||(pics.size&&!pics.has(page2Key_(c.pic[i])))||(brands.size&&!brands.has(page2Key_(c.brand[i])))||(subbrands.size&&!subbrands.has(page2Key_(c.subbrand[i])))||(categories.size&&!categories.has(page2Key_(c.category[i]))))continue;
  matchedSourceRows+=Number(c.sourceRows[i])||1;
  const key=page2Key_(name),sbu=page2Text_(c.sbu[i])||'LAINNYA',mapKey=sbu+'|'+key,m=page2Month_(c.month[i]),value=Number(c.value[i])||0;
  let x=map.get(mapKey);if(!x){const masterSku=master.get(key)||{};x={sbu:sbu,sku:name,brand:masterSku.brand||page2Text_(c.brand[i]),subbrand:masterSku.subbrand||page2Text_(c.subbrand[i]),category:masterSku.category||page2Text_(c.category[i]),entity:entity,monthly:{},totalValue:0};PAGE4_MONTHS_.forEach(mm=>x.monthly[mm]=0);map.set(mapKey,x);}
  x.monthly[m]+=value;x.totalValue+=value;totalValue+=value;
 }
 const rows=Array.from(map.values()).map(x=>{const vals=monthsUsed.map(m=>x.monthly[m]||0),last=monthsUsed.length?PAGE4_MONTHS_.indexOf(monthsUsed[monthsUsed.length-1]):-1,tail3=last<0?[]:PAGE4_MONTHS_.slice(Math.max(0,last-2),last+1),tail6=last<0?[]:PAGE4_MONTHS_.slice(Math.max(0,last-5),last+1);x.avg3=tail3.length?tail3.reduce((a,m)=>a+x.monthly[m],0)/3:0;x.avg6=tail6.length?tail6.reduce((a,m)=>a+x.monthly[m],0)/6:0;x.months=PAGE4_MONTHS_.map(m=>x.monthly[m]||0);return x;}).sort((a,b)=>b.totalValue-a.totalValue||a.sku.localeCompare(b.sku,'id'));
 return {mode:mode,modeSalesType:mode==='SOO'?'SOO':'ALL',label:mode==='SOO'?'SUBDIST':'SUB_DIST',year:year,monthsUsed:monthsUsed,monthsVisible:monthsUsed,subdist:selected.size?Array.from(selected).join(', '):'Semua Entitas',rows:rows,skuCount:rows.length,totals:{value:totalValue},matchedSourceRows:matchedSourceRows,sourceRows:meta.sourceRows};
}
function getPage4InitialPayload(mode){
 const started=Date.now(),readStarted=Date.now(),src=page4ReadSource_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page4BuildMeta_(src),metaMs=Date.now()-metaStarted;
 const selectedMode=page34Mode_(mode),year=meta.years.includes('2026')?2026:Number(meta.years[meta.years.length-1]);
 const aggregateStarted=Date.now(),data=page4BuildData_(src,meta,{mode:'SOO',year:year}),dataSO=page4BuildData_(src,meta,{mode:'SO',year:year}),aggregateMs=Date.now()-aggregateStarted;
 return {meta:meta,data:data,dataSO:dataSO,diagnostics:{source:src.sourceName,sourceRows:src.sourceRows,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started}};
}
function getPage4SkuData(filters){
 const started=Date.now(),readStarted=Date.now(),src=page4ReadSource_(),readMs=Date.now()-readStarted;
 const metaStarted=Date.now(),meta=page4BuildMeta_(src),metaMs=Date.now()-metaStarted;
 const aggregateStarted=Date.now(),data=page4BuildData_(src,meta,filters),aggregateMs=Date.now()-aggregateStarted;
 data.diagnostics={source:src.sourceName,sourceRows:src.sourceRows,readMs:readMs,metaMs:metaMs,aggregateMs:aggregateMs,totalMs:Date.now()-started};return data;
}
