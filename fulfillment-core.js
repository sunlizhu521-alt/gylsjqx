(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FulfillmentCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const SHEETS=['徐英','李奇','操思敏','邢芳芳','孙立柱'];
const COLUMNS=['采购下单人','事业部','销售分渠道','运营','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','未生产预计交付时间','是否正常履约'];
const DATE_FORMAT='yyyy"年"m"月"d"日"';
const norm=v=>String(v??'').replace(/\s/g,'');
function dateCell(cell,date1904){
 const v=cell?.v;if(v==null||v==='')return {t:'s',v:''};if(String(v).trim()==='/')return {t:'s',v:'/'};
 let y,m,d;
 if(typeof v==='number'){const p=X.SSF.parse_date_code(v,{date1904});if(p){y=p.y;m=p.m;d=p.d;}}
 else {const match=String(v).trim().match(/^(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);if(match)[,y,m,d]=match.map(Number);}
 const utc=Date.UTC(y,m-1,d),check=new Date(utc);
 if(!y||y<1900||y>9999||check.getUTCFullYear()!==y||check.getUTCMonth()!==m-1||check.getUTCDate()!==d)return {t:'s',v:String(v),invalid:true};
 let serial=(utc-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:DATE_FORMAT};
}
function merge(book){
 const missing=SHEETS.filter(n=>!book.Sheets[n]);if(missing.length)throw Error('缺少工作表：'+missing.join('、'));
 const rows=[],sources=[],counts=[],warnings=[];
 for(const name of SHEETS){const sheet=book.Sheets[name],range=X.utils.decode_range(sheet['!ref']||'A1');let hr=-1,map;
 for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){const h={};for(let c=range.s.c;c<=range.e.c;c++){let key=norm(sheet[X.utils.encode_cell({r,c})]?.v);if(key==='为生产预计交付时间')key='未生产预计交付时间';if(key)(h[key]??=[]).push(c);}if(h['物料编码']&&h['事业部']){hr=r;map=h;break;}}
 if(hr<0)throw Error(name+'：找不到明细表头');
 for(const key of COLUMNS){if(!map[key])throw Error(name+'：缺少列“'+key+'”');if(map[key].length!==1)throw Error(name+'：重复列“'+key+'”');}
 let count=0;
 for(let r=hr+1;r<=range.e.r;r++){let nonempty=false;for(let c=range.s.c;c<=range.e.c;c++){const cell=sheet[X.utils.encode_cell({r,c})];if(cell?.v!=null&&String(cell.v).trim()!==''){nonempty=true;break;}}if(!nonempty)continue;
 const record=COLUMNS.map((key,i)=>{const address=X.utils.encode_cell({r,c:map[key][0]}),cell=sheet[address];let output;
 if(['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)){output=dateCell(cell,!!book.Workbook?.WBProps?.date1904);if(output.invalid)warnings.push(`${name}!${address}：日期“${cell.v}”无法识别，保留原文`);}
 else if(cell?.t==='e'||(cell?.f&&cell.v==null)){warnings.push(`${name}!${address}：公式结果不可用，请在 Excel 中重算后上传`);output={t:'s',v:cell.w||'#公式结果缺失'};}
 else if(key==='物料编码'||key==='SKU')output={t:'s',v:cell?String(cell.w??cell.v??''):''};
 else output={t:typeof cell?.v==='number'?'n':'s',v:cell?.v??''};
 delete output.invalid;return output;});rows.push(record);sources.push({sheet:name,row:r+1});count++;}
 counts.push({sheet:name,count});}
 return {rows,sources,counts,warnings};
}
function columnWidth(key){return key==='物料名称'?42:['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)?23:key==='SKU'?26:20;}
function workbook(result){const s=X.utils.aoa_to_sheet([COLUMNS]);result.rows.forEach((row,r)=>row.forEach((cell,c)=>s[X.utils.encode_cell({r:r+1,c})]={...cell}));s['!ref']=`A1:${X.utils.encode_col(COLUMNS.length-1)}${result.rows.length+1}`;s['!autofilter']={ref:s['!ref']};s['!cols']=COLUMNS.map((key,col)=>({wch:[key,...result.rows.map(row=>row[col].z?X.SSF.format(row[col].z,row[col].v):String(row[col].v))].reduce((max,text)=>Math.max(max,[...text].reduce((width,ch)=>width+(ch.charCodeAt(0)>255?2:1),0)+4),columnWidth(key))}));const w=X.utils.book_new();X.utils.book_append_sheet(w,s,'汇总表');return w;}
return {SHEETS,COLUMNS,DATE_FORMAT,columnWidth,dateCell,merge,workbook};
});
