/**
 * PAGE 4 — SKU KCH. Aggregates SOO only; uses SUBDIST for distributor and SKU master for attributes.
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
function page4ReadSource_(){
 if(PropertiesService.getScriptProperties().getProperty('MANTUL_V2_DATA_STALE')==='TRUE')throw new Error('Database V2 belum diperbarui setelah import. Selesaikan build, finalize, dan publish terlebih dahulu.');
 const ss=getDatabase_(),source=ss.getSheetByName('P3_OUTLET_MONTHLY'),sku=ss.getSheetByName('SKU');
 if(!source||source.getLastRow()<2)throw new Error('P3_OUTLET_MONTHLY belum berisi data. Jalankan builder Database V2 terlebih dahulu.');
 if(!sku||sku.getLastRow()<2)throw new Error('Sheet SKU belum tersedia.');
 const expected=['YEAR','MONTH','DD','CUSTOMER_CODE','CUSTOMER_NAME','SUBDIST','SUB_DIST','SALES_TYPE','PIC','BRAND','SUBBRAND','CATEGORY','SBU','ITEM_CODE','SKU','VAL','CARTON','SOURCE_ROWS'];
 const eh=source.getRange(1,1,1,18).getDisplayValues()[0].map(page2HeaderKey_);
 if(!expected.every((v,i)=>page2HeaderKey_(v)===eh[i]))throw new Error('Header P3_OUTLET_MONTHLY tidak sesuai schema V2: '+eh.join(' | '));
 const em={year:0,month:1,customer:4,subdist:5,brand:9,subbrand:10,category:11,sbu:12,salesType:7,value:15,carton:16,sku:14,sourceRows:17};
 const eventData=page4ReadSelectedColumns_(source,em),mh=sku.getRange(1,1,1,sku.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
 const mm={sku:page4HeaderIndex_(mh,['PRODUCT_DESC','PRODUCT DESC'],'SKU master PRODUCT_DESC'),subbrand:page4HeaderIndex_(mh,['SUB_BRAND','SUB BRAND','SUBBRAND'],'SKU master SUB_BRAND'),
  brand:page4HeaderIndex_(mh,['BRAND'],'SKU master BRAND'),category:page4HeaderIndex_(mh,['CATEGORY'],'SKU master CATEGORY'),entity:page4HeaderIndex_(mh,['ENTITAS','ENTITY'],'SKU master ENTITAS')};
 const masterData=page4ReadSelectedColumns_(sku,mm);return {sourceRows:eventData.rows,columns:eventData.columns,master:masterData.columns,masterRows:masterData.rows};
}
function page4BuildMeta_(src){
 const years=new Set(),monthsByYear={},skuAttrs=new Map(),customers=new Set();let sooRows=0,soModeRows=0;const c=src.columns;
 const add=(map,key,field,value)=>{if(!key||value==null||String(value).trim()==='')return;let obj=map.get(key);if(!obj){obj={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};map.set(key,obj);}obj[field].add(String(value).trim());};
 for(let i=0;i<src.sourceRows;i++){const type=page2Key_(c.salesType[i]);if(type!=='SOO')continue;sooRows++;
  const y=String(Number(c.year[i])||c.year[i]||'').trim(),m=page2Month_(c.month[i]),name=page2Text_(c.sku[i]);if(page2Text_(c.customer[i]))customers.add(page2Text_(c.customer[i]));if(y)years.add(y);if(y&&m){if(!monthsByYear[y])monthsByYear[y]=new Set();monthsByYear[y].add(m);}
  add(skuAttrs,name,'subbrand',c.subbrand[i]);add(skuAttrs,name,'brand',c.brand[i]);add(skuAttrs,name,'category',c.category[i]);add(skuAttrs,name,'entity',c.sbu[i]);
 }
 const masterMap=new Map();for(let i=0;i<src.masterRows;i++){const name=page2Text_(src.master.sku[i]);if(!name)continue;let x=masterMap.get(name);if(!x){x={subbrand:new Set(),brand:new Set(),category:new Set(),entity:new Set()};masterMap.set(name,x);}['subbrand','brand','category','entity'].forEach(k=>{const v=src.master[k][i];if(v!=null&&String(v).trim())x[k].add(String(v).trim());});}
 const choose=(set,fallback)=>set&&set.size===1?Array.from(set)[0]:set&&set.size>1?'MULTI':fallback||'',names=new Set([...masterMap.keys(),...skuAttrs.keys()]);
 const skuMaster=Array.from(names).map(name=>{const raw=skuAttrs.get(name)||{},master=masterMap.get(name)||{};return {sku:name,subbrand:choose(raw.subbrand,choose(master.subbrand,'')),brand:choose(raw.brand,choose(master.brand,'')),category:choose(raw.category,choose(master.category,'')),entity:choose(raw.entity,choose(master.entity,''))};}).sort((a,b)=>a.sku.localeCompare(b.sku,'id'));
 return {years:Array.from(years).sort((a,b)=>Number(a)-Number(b)),monthsByYear:Object.fromEntries(Object.keys(monthsByYear).map(y=>[y,PAGE4_MONTHS_.filter(m=>monthsByYear[y].has(m))])),subdists:PAGE34_ENTITIES_.slice(),customers:Array.from(customers).sort((a,b)=>a.localeCompare(b,'id')),skuMaster:skuMaster,skuMasterCount:skuMaster.length,sourceRows:src.sourceRows,sooRows:sooRows};
}
function page4BuildData_(src,meta,filters){
 filters=filters||{};const mode='SOO',year=Number(meta.years.slice(-1)[0])||new Date().getFullYear(),monthsUsed=meta.monthsByYear[String(year)]||[],monthSet=new Set(monthsUsed),selected=new Set(page2FilterList_(filters.subdists).map(page2Key_)),customers=new Set(page2FilterList_(filters.customers).map(page2Key_));
 const map=new Map(meta.skuMaster.map(s=>[s.sku,{...s,total:{value:0,carton:0}}]));let matchedSourceRows=0;const c=src.columns;
 for(let i=0;i<src.sourceRows;i++){const type=page2Key_(c.salesType[i]);if(type!=='SOO')continue;
  if(Number(c.year[i])!==year||!monthSet.has(page2Month_(c.month[i])))continue;
  const row=[null,null,null,null,null,c.subdist[i],null,null,null,null,null,c.salesType[i],null,null,null,null,c.sku[i]],entity=page34Entity_(row,mode),customer=page2Text_(c.customer[i]);
  if(!entity||(selected.size&&!selected.has(page2Key_(entity)))||(customers.size&&!customers.has(page2Key_(customer))))continue;const name=page2Text_(c.sku[i]);if(!name)continue;matchedSourceRows+=Number(c.sourceRows[i])||1;
  let x=map.get(name);if(!x){x={sku:name,brand:page2Text_(c.brand[i]),subbrand:page2Text_(c.subbrand[i]),category:page2Text_(c.category[i]),entity:page2Text_(c.sbu[i]),total:{value:0,carton:0}};map.set(name,x);}
  x.total.value+=Number(c.value[i])||0;x.total.carton+=Number(c.carton[i])||0;
 }
 const all=Array.from(map.values()).sort((a,b)=>(b.total.value-a.total.value)||a.sku.localeCompare(b.sku,'id')),totals=all.reduce((a,r)=>({value:a.value+r.total.value,carton:a.carton+r.total.carton}),{value:0,carton:0});
 return {mode:mode,modeSalesType:mode==='SOO'?'SOO':'Semua',label:'SUBDIST',year:year,monthsUsed:monthsUsed,subdist:selected.size?Array.from(selected).join(', '):'Semua SUBDIST',rows:all,skuCount:all.length,totals:totals,matchedSourceRows:matchedSourceRows,sourceRows:src.sourceRows};
}
function getPage4InitialPayload(mode){const started=Date.now(),src=page4ReadSource_(),meta=page4BuildMeta_(src),selectedMode=page34Mode_(mode),year=meta.years.includes('2026')?'2026':meta.years[meta.years.length-1];return {meta:meta,data:page4BuildData_(src,meta,{mode:selectedMode,year:year}),diagnostics:{source:'P3_OUTLET_MONTHLY',readAndAggregateMs:Date.now()-started}};}
function getPage4SkuData(filters){const started=Date.now(),src=page4ReadSource_(),meta=page4BuildMeta_(src),data=page4BuildData_(src,meta,filters);data.diagnostics={elapsedMs:Date.now()-started};return data;}
