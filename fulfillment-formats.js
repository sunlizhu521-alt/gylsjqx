(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'),require('./fulfillment-core.js'));else root.FulfillmentFormats=factory(root.XLSX,root.FulfillmentCore);})(typeof globalThis!=='undefined'?globalThis:this,function(X,defaultCore){
'use strict';
const DAY_FORMAT='yyyy"年"m"月"d"日"',MONTH_FORMAT='yyyy"年"m"月"';
const norm=value=>String(value??'').replace(/\s/g,'');
const decimal=value=>value.toLocaleString('en-US',{useGrouping:false,maximumSignificantDigits:21});
const deliveryFields=new Set(['合同约定交期','生产中交付时间','为生产预计交付时间','未生产预计交付时间']);
const identifiers=new Set(['物料编码','SKU','OA备货单号','OA备货流程号','OA流程号','采购订单号','合同编号','合同号']);
const errors={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'};
function numericFormat(value){const precision=(decimal(value).split('.')[1]||'').length;return precision>250?'0.0000000000000000E+00':precision?'0.'+'0'.repeat(precision):'0';}
function dateInput(cell){
 if(cell?.v instanceof Date){const d=cell.v;return Number.isFinite(d.getTime())?{...cell,v:`${d.getUTCFullYear()}-${d.getUTCMonth()+1}-${d.getUTCDate()}`}:{...cell,v:String(d)};}
 return cell;
}
function textValue(cell,field,date1904,core){
 const value=cell?.v;if(value==null)return '';
 if(cell.t==='e')return cell.w||errors[value]||String(value);
 if(cell.t==='b')return value?'TRUE':'FALSE';
 if(!identifiers.has(field)&&(cell.t==='d'||value instanceof Date||(typeof value==='number'&&(deliveryFields.has(field)||X.SSF.is_date(cell.z||''))))){
  const date=core.dateCell(dateInput(cell),date1904);
  if(!date.invalid&&date.t==='n'){
   const monthOnly=X.SSF.is_date(cell.z||'')&&!/[dD日]/.test(cell.z||'');
   return X.SSF.format(monthOnly?MONTH_FORMAT:DAY_FORMAT,date.v);
  }
 }
 if(typeof value!=='number')return String(value);
 // Honor an explicit zero-padded identifier format, never a stale scientific display cache.
 if(/^0{2,}$/.test(cell.z||'')&&Number.isFinite(value))return X.SSF.format(cell.z,value);
 return Number.isFinite(value)?decimal(value):String(value);
}
function normalize(input,core=defaultCore){
 const book={...input,Sheets:{...input.Sheets}},updates=[],issues=[],warnings=[];let normalizedSheetCount=0;
 const quantityFields=new Set(core.QUANTITY_FIELDS),date1904=!!input.Workbook?.WBProps?.date1904;
 const selected=Object.keys(input.Sheets||{}).filter(name=>core.SHEETS.some(person=>new RegExp('^'+person+'(?:[0-9]{4}|[0-9]{8})?$').test(norm(name))));
 for(const name of selected){
  const original=input.Sheets[name],range=X.utils.decode_range(original['!ref']||'A1');let header=-1,fields=new Map();
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){
   const candidate=new Map();for(let c=range.s.c;c<=range.e.c;c++){const key=norm(original[X.utils.encode_cell({r,c})]?.v);if(key)candidate.set(c,key);}
   const names=[...candidate.values()];if(names.includes('事业部')&&names.includes('物料编码')){header=r;fields=candidate;break;}
  }
  if(header<0){issues.push({sheet:name,cell:'',field:'单元格格式',message:'找不到事业部和物料编码所在的明细表头，未转换工作表格式'});continue;}
  const sheet={...original},updateStart=updates.length;book.Sheets[name]=sheet;
  const addresses=new Set(Object.keys(original).filter(address=>/^[A-Z]+[1-9]\d*$/.test(address)));
  for(const merge of original['!merges']||[]){if(merge.s.c===merge.e.c&&quantityFields.has(fields.get(merge.s.c))&&merge.s.r>header)for(let r=merge.s.r+1;r<=merge.e.r;r++)addresses.add(X.utils.encode_cell({r,c:merge.s.c}));}
  for(const address of addresses){
   const position=X.utils.decode_cell(address),source=original[address]||{t:'s',v:''},field=position.r>header?fields.get(position.c):'';
   if(source.f){issues.push({sheet:name,cell:address,field:'单元格格式',message:'公式尚未成功重算，保留原单元格，不能直接使用保存的旧结果'});continue;}
   let value,type='s',format='@';
   if(quantityFields.has(field)){
    const quantity=core.quantityCell(original,position.r,position.c);
    if(quantity.invalid||(quantity.v!==''&&(quantity.t!=='n'||!Number.isFinite(quantity.v)||Math.abs(quantity.v)>Number.MAX_SAFE_INTEGER))){
     issues.push({sheet:name,cell:address,field:'单元格格式',message:`${field}：${quantity.invalid||'不是可精确核验的有效数值'}`});continue;
    }
    value=quantity.v;type=value===''?'s':'n';format=type==='n'?numericFormat(value):'0';
   }else if(field==='下单月份'||field==='下单日期'){
    const candidate=field==='下单月份'&&norm(source.v)==='/'?{v:'2026-04-01'}:dateInput(source);
    const date=core.dateCell(candidate,date1904);format=field==='下单月份'?MONTH_FORMAT:DAY_FORMAT;
    if(date.t==='n'&&!date.invalid){value=date.v-(date1904?1462:0);type='n';}
    else {value=source.v==null?'':String(source.v);if(value.trim()&&value.trim()!=='/')warnings.push(`${name}!${address}：日期“${value}”无法识别，保留原文`);}
   }else value=textValue(source,field,date1904,core);
   const cell={...source,t:type,v:value,z:format};delete cell.w;delete cell.f;delete cell.F;delete cell.D;
   if(source.t===cell.t&&source.v===cell.v&&source.z===cell.z&&!('F'in source)&&!('D'in source))continue;
   sheet[address]=cell;updates.push({sheet:name,cell:address,value,type,format});
  }
  if(updates.length>updateStart)normalizedSheetCount++;else book.Sheets[name]=original;
 }
 return {book,updates,issues,warnings,normalizedSheetCount};
}
return {normalize};
});
