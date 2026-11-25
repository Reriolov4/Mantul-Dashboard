/**
 * PAGE 4 — SKU KCH
 * RAW_DATA_SALES is resolved by header names to tolerate column movement.
 * Subdist follows the Page 2 key (DISTRIBUTORS); SKU KCH maps to SKU.PRODUCT_DESC.
 * Read-only: no source or derived sheet is written.
 */
const PAGE4_MONTHS_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function page4Month_(v){return page2Month_(v);}
function page4HeaderIndex_(headers,aliases,label){
 const normalized=headers.map(page2HeaderKey_);
 for(let i=0;i<aliases.length;i++){const at=normalized.indexOf(page2HeaderKey_(aliases[i]));if(at>=0)return at;}
 throw new Error('Kolom '+label+' tidak ditemukan. Header: '+headers.join(' | '));
}
function page4ReadSelectedColumns_(sheet,headerMap){
 const n=sheet.getLastRow()-1,keys=Object.keys(headerMap),indices=keys.map(k=>headerMap[k]),groups=[];
 indices.slice().sort((a,b)=>a-b).forEach(col=>{const last=groups[groups.length-1];if(last&&col<=last.end+1)last.end=col;else groups.push({start:col,end:col,values:null});});
 groups.forEach(g=>{g.values=sheet.getRange(2,g.start+1,n,g.end-g.start+1).getValues();});
 const result={};keys.forEach(key=>{const col=headerMap[key],g=groups.find(x=>col>=x.start&&col<=x.end);result[key]=g.values.map(row=>row[col-g.start]);});
 return {rows:n,columns:result};
}
function page4ReadSource_(){
 const ss=getDatabase_(),raw=ss.getSheetByName('RAW_DATA_SALES'),sku=ss.getSheetByName('SKU');
 if(!raw||raw.getLastRow()<2)throw new Error('RAW_DATA_SALES belum berisi data.');
 if(!sku||sku.getLastRow()<2)throw new Error('Sheet SKU belum tersedia.');
 const rh=raw.getRange(1,1,1,raw.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
 const rm={
  year:page4HeaderIndex_(rh,['TAHUN','YEAR'],'TAHUN'),month:page4HeaderIndex_(rh,['BULAN','MONTH'],'BULAN'),
  subbrand:page4HeaderIndex_(rh,['SUB BRAND','SUBBRAND'],'SUBBRAND'),brand:page4HeaderIndex_(rh,['BRAND'],'BRAND'),
  carton:page4HeaderIndex_(rh,['KARTON','CARTON'],'KARTON'),subdist:page4HeaderIndex_(rh,['DISTRIBUTORS','DISTRIBUTOR','SUBDIST'],'SUBDIST'),
  value:page4HeaderIndex_(rh,['VAL','VALUE'],'VAL'),salesType:page4HeaderIndex_(rh,['JENIS SALES','SALES TYPE','SALES_TYPE'],'JENIS SALES'),
  category:page4HeaderIndex_(rh,['CATEGORY'],'CATEGORY'),sbu:page4HeaderIndex_(rh,['SBU'],'SBU'),sku:page4HeaderIndex_(rh,['SKU KCH','SKU_KCH'],'SKU KCH')
 };
 const rawData=page4ReadSelectedColumns_(raw,rm);
 const mh=sku.getRange(1,1,1,sku.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
 const mm={sku:page4HeaderIndex_(mh,['PRODUCT_DESC','PRODUCT DESC'],'SKU master PRODUCT_DESC'),
  subbrand:page4HeaderIndex_(mh,['SUB_BRAND','SUB BRAND','SUBBRAND'],'SKU master SUB_BRAND'),
  brand:page4HeaderIndex_(mh,['BRAND'],'SKU master BRAND'),category:page4HeaderIndex_(mh,['CATEGORY'],'SKU master CATEGORY'),
  entity:page4HeaderIndex_(mh,['ENTITAS','ENTITY'],'SKU master ENTITAS')};
 const masterData=page4ReadSelectedColumns_(sku,mm);
 return {sourceRows:rawData.rows,columns:rawData.columns,master:masterData.columns,masterRows:masterData.rows};
}
function page4BuildMeta_(src){
 const years=new Set(),monthsByYear={},subdists=new Set(),skuMap=new Map();let sooRows=0;
 for(let i=0;i<src.sourceRows;i++){const c=src.columns;if(page2Key_(c.salesType[i])!=='SOO')continue;sooRows++;
  const y=String(Number(c.year[i])||c.year[i]||'').trim(),m=page4Month_(c.month[i]),sd=page2Text_(c.subdist[i]);
  if(y)years.add(y);if(sd)subdists.add(sd);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
 }
 for(let i=0;i<src.masterRows;i++){const name=page2Text_(src.master.sku[i]);if(!name||skuMap.has(name))continue;
  skuMap.set(name,{sku:name,subbrand:page2Text_(src.master.subbrand[i]),brand:page2Text_(src.master.brand[i]),category:page2Text_(src.master.category[i]),entity:page2Text_(src.master.entity[i])});
 }
 const sort=a=>Array.from(a).sort((x,y)=>x.localeCompare(y,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),
  monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE4_MONTHS_.filter(m=>monthsByYear[y].has(m))])),
  subdists:sort(subdists),skuMaster:Array.from(skuMap.values()).sort((a,b)=>a.sku.localeCompare(b.sku,'id')),
  skuMasterCount:skuMap.size,sourceRows:src.sourceRows,sooRows:sooRows};
}
function page4BuildData_(src,meta,filters){
 filters=filters||{};const year=Number(filters.year)||Number(meta.years.slice(-1)[0])||new Date().getFullYear();
 const monthsUsed=meta.monthsByYear[String(year)]||[],monthSet=new Set(monthsUsed),selected=page2FilterList_(filters.subdists),subdist=selected[0]||'';
 const map=new Map(meta.skuMaster.map(s=>[s.sku,{...s,total:{value:0,carton:0}}]));let matchedSourceRows=0;
 for(let i=0;i<src.sourceRows;i++){const c=src.columns;if(page2Key_(c.salesType[i])!=='SOO'||Number(c.year[i])!==year||!monthSet.has(page4Month_(c.month[i])))continue;
  const sd=page2Text_(c.subdist[i]);if(subdist&&page2Key_(sd)!==subdist)continue;
  const sku=page2Text_(c.sku[i]);if(!sku)continue;matchedSourceRows++;
  let x=map.get(sku);if(!x){x={sku:sku,brand:page2Text_(c.brand[i]),subbrand:page2Text_(c.subbrand[i]),category:page2Text_(c.category[i]),entity:page2Text_(c.sbu[i]),total:{value:0,carton:0}};map.set(sku,x);}
  x.total.value+=Number(c.value[i])||0;x.total.carton+=Number(c.carton[i])||0;
 }
 const all=Array.from(map.values()).sort((a,b)=>(b.total.value-a.total.value)||a.sku.localeCompare(b.sku,'id'));
 const totals=all.reduce((a,r)=>({value:a.value+r.total.value,carton:a.carton+r.total.carton}),{value:0,carton:0});
 return {year:year,monthsUsed:monthsUsed,subdist:subdist||'Semua Subdist',rows:all,skuCount:all.length,totals:totals,matchedSourceRows:matchedSourceRows,sourceRows:src.sourceRows};
}
function getPage4InitialPayload(){
 const started=Date.now(),src=page4ReadSource_(),meta=page4BuildMeta_(src),year=meta.years.includes('2026')?'2026':meta.years[meta.years.length-1];
 return {meta:meta,data:page4BuildData_(src,meta,{year:year}),diagnostics:{source:'RAW_DATA_SALES',readAndAggregateMs:Date.now()-started}};
}
function getPage4SkuData(filters){const started=Date.now(),src=page4ReadSource_(),meta=page4BuildMeta_(src),data=page4BuildData_(src,meta,filters);data.diagnostics={elapsedMs:Date.now()-started};return data;}