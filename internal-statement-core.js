(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.InternalStatement=factory(root.XLSX);})(globalThis,function(X){
'use strict';
const HEADERS=['物料编码','名称','单价','数量','总计','备注'];
const REQUIRED=['供应商','采购组织','入库日期','物料编码','物料名称','实收数量','含税单价','价税合计'];
const text=c=>String(c?.v??'').trim(),norm=s=>String(s??'').normalize('NFKC').replace(/\s+/g,'');
function decimal(value,label='数值',precision=8){
 if(value?.t==='e'||value?.f&&(value.v==null||value.v===''))throw Error(label+'：公式结果不可用，请核实源表');
 let s=String(value&&typeof value==='object'?value.v??'':value??'').normalize('NFKC').trim();
 if(!s)return null;
 s=s.replace(/^[¥￥]\s*/,'').replace(/^\((.*)\)$/,'-$1');
 if(!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(s))throw Error(label+'：不是有效数值“'+s+'”');
 s=s.replace(/,/g,'');let [whole,fraction='']=s.replace(/^[+-]/,'').split('.');fraction=fraction.replace(/0+$/,'');
 if(fraction.length>precision)throw Error(label+'：小数超过 '+precision+' 位，请核实');
 const n=(s.startsWith('-')?-1n:1n)*(BigInt(whole)*10n**BigInt(precision)+BigInt(fraction.padEnd(precision,'0')));
 if(n>9007199254740991n||n< -9007199254740991n)throw Error(label+'：数值过大');
 return n;
}
function decimalText(value,precision=8,min=0){
 const sign=value<0n?'-':'',digits=(value<0n?-value:value).toString().padStart(precision+1,'0');
 const fraction=precision?digits.slice(-precision).replace(/0+$/,'').padEnd(min,'0'):'';
 return sign+(precision?digits.slice(0,-precision):digits)+(fraction?'.'+fraction:'');
}
const amountText=n=>decimalText(BigInt(n),2,2);
function date(c,date1904=false,label='日期'){
 if(c?.t==='e'||c?.f&&(c.v==null||c.v===''))throw Error(label+'：公式结果不可用');
 const v=c?.v??c;let year,month,day;
 if(typeof v==='number'){const d=X.SSF.parse_date_code(v,{date1904});if(d){year=d.y;month=d.m;day=d.d;}}
 else{const m=String(v??'').trim().match(/^(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);if(m)[,year,month,day]=m.map(Number);}
 const d=new Date(Date.UTC(year,month-1,day));
 if(!year||year<1900||d.getUTCFullYear()!==year||d.getUTCMonth()!==month-1||d.getUTCDate()!==day)throw Error(label+'：日期无法识别“'+String(v??'')+'”');
 return `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}
const dateLabel=s=>s?`${Number(s.slice(0,4))}年${Number(s.slice(5,7))}月${Number(s.slice(8,10))}日`:'';
function parse(book,fileName){
 const groups=new Map(),warnings=[];let count=0,grandQuantity=0n,grandAmount=0n,matchedSheets=0;
 for(const name of book.SheetNames){
  const sheet=book.Sheets[name];if(!sheet['!ref'])continue;const range=X.utils.decode_range(sheet['!ref']);
  if(range.e.r>100000||range.e.c>300)throw Error(name+'：工作表范围过大，请删除无用途空行列');
  let header=-1,map;
  for(let r=range.s.r;r<=Math.min(range.e.r,30);r++){const m={};for(let c=range.s.c;c<=range.e.c;c++){const h=norm(text(sheet[X.utils.encode_cell({r,c})]));if(h){if(m[h]!==undefined&&REQUIRED.includes(h))throw Error(name+'：表头重复“'+h+'”');m[h]=c;}}if(REQUIRED.every(h=>m[h]!==undefined)){header=r;map=m;break;}}
  if(header<0){warnings.push(name+'：未识别为采购入库明细，未读取');continue;}matchedSheets++;
  let quantity=0n,amount=0n;const totals=[];
  for(let r=header+1;r<=range.e.r;r++){
   const at=h=>sheet[X.utils.encode_cell({r,c:map[h]})],pos=h=>`${fileName} / ${name}!${X.utils.encode_cell({r,c:map[h]})}（${h}）`;
   if(REQUIRED.every(h=>!text(at(h))))continue;
   if(/^(合计|总计|小计)[：:]?$/.test(norm(text(at('供应商'))))){totals.push({r,quantity:at('实收数量'),amount:at('价税合计')});continue;}
   for(const h of REQUIRED)if(!text(at(h))||at(h)?.t==='e')throw Error(pos(h)+'为空或异常，请核实');
   const supplier=text(at('供应商')),buyer=text(at('采购组织')),day=date(at('入库日期'),!!book.Workbook?.WBProps?.date1904,pos('入库日期')),month=day.slice(0,7);
   const code=text(at('物料编码'));if(at('物料编码').t==='n'&&!Number.isSafeInteger(at('物料编码').v))throw Error(pos('物料编码')+'不是完整整数，请将源编码保存为文本');
   const qty=decimal(at('实收数量'),pos('实收数量')),price=decimal(at('含税单价'),pos('含税单价')),gross=decimal(at('价税合计'),pos('价税合计'),2);
   const key=JSON.stringify([norm(buyer),norm(supplier),month]);if(!groups.has(key))groups.set(key,{key,buyer,supplier,month,items:new Map(),count:0,quantity:0n,amount:0n});
   const group=groups.get(key),itemKey=JSON.stringify([code,price.toString()]);
   if(!group.items.has(itemKey))group.items.set(itemKey,{code,price,names:new Set(),notes:new Set(),quantity:0n,amount:0n,sources:[]});
   const item=group.items.get(itemKey);item.names.add(text(at('物料名称')));if(map.备注!==undefined&&text(at('备注')))item.notes.add(text(at('备注')));
   item.quantity+=qty;item.amount+=gross;item.sources.push({sheet:name,row:r+1});group.count++;group.quantity+=qty;group.amount+=gross;
   quantity+=qty;amount+=gross;count++;grandQuantity+=qty;grandAmount+=gross;
  }
  if(totals.length===1){const total=totals[0];for(const [field,cell,expected,precision]of [['实收数量',total.quantity,quantity,8],['价税合计',total.amount,amount,2]]){const n=decimal(cell,`${name} 第 ${total.r+1} 行 ${field}`,precision);if(n!==null&&n!==expected)throw Error(`${name}：${field}合计不一致，明细 ${decimalText(expected,precision)}，源表合计 ${decimalText(n,precision)}，请核实后上传`);}}
  else if(totals.length>1)throw Error(name+'：存在多个合计或小计行，请提供仅含明细及一个总计的导出表');
 }
 if(!matchedSheets)throw Error('未找到采购入库明细表头，需要：'+REQUIRED.join('、'));
 if(!count)throw Error('文件没有有效入库明细');
 const result=[...groups.values()].map(g=>({...g,quantity:decimalText(g.quantity),amount:g.amount.toString(),items:[...g.items.values()].map(item=>{if(item.names.size>1)warnings.push(g.buyer+' / '+g.supplier+'：物料 '+item.code+' 存在不同名称，已全部保留');return {code:item.code,name:[...item.names].join('\n'),price:decimalText(item.price),quantity:decimalText(item.quantity),amount:item.amount.toString(),note:[...item.notes].join('\n'),sources:item.sources};})}));
 return {groups:result,count,quantity:decimalText(grandQuantity),amount:grandAmount.toString(),warnings};
}
function defaults(month){
 const next=new Date(date(month+'-01')+'T00:00:00Z');next.setUTCMonth(next.getUTCMonth()+1,10);
 return {reconcileDate:next.toISOString().slice(0,10),cutoff:'',terms:'票到90天付款',balance:'',receipts:'',afterSales:'',returns:'',invoiced:'',uninvoiced:'',note:'',phone:'',invoices:Array.from({length:5},()=>({month:'',status:'',amount:''}))};
}
function footer(group,input){
 const data={...input},labels={balance:'上期余额',receipts:'本期收款',afterSales:'售后费用',returns:'退货金额',invoiced:'已开票欠款',uninvoiced:'未开票欠款'};
 for(const [k,label]of Object.entries(labels)){const n=decimal(input[k],label,2);data[k]=n===null?'':amountText(n);}
 if(!input.reconcileDate)throw Error('请选择对账日期');data.reconcileDate=date(input.reconcileDate);if(input.cutoff)data.cutoff=date(input.cutoff);
 const complete=['balance','receipts','afterSales','returns'].every(k=>data[k]!=='');
 data.total=complete?amountText(BigInt(group.amount)+decimal(data.balance,'',2)-decimal(data.receipts,'',2)-decimal(data.afterSales,'',2)-decimal(data.returns,'',2)):'';
 data.invoices=input.invoices.map(row=>{if(row.month&&!/^\d{4}-(0[1-9]|1[0-2])$/.test(row.month))throw Error('货款汇总月份无效');const n=decimal(row.amount,'开具金额',2);return {...row,amount:n===null?'':amountText(n)};});
 return data;
}
const fileName=(buyer,groups)=>'内部交易对账单-'+buyer+'-'+[...new Set(groups.map(g=>g.month))].join('、');
const shortNames=new Map([
 ['迈德斯特',['浙江迈德斯特医疗器械科技有限公司','迈德斯特斯特医疗器械有限公司科技有限公司']],
 ['瑞朗德',['河北瑞朗德医疗器械科技有限公司','河北瑞朗德医疗器械科技集团有限公司']],
 ['迈德斯特（宁波）',['迈德斯特（宁波）医疗器械科技有限公司','迈德斯特（宁波）医疗科技有限公司']],
 ['国源',['杭州国源养老科技有限公司']]
].flatMap(([short,names])=>[short,...names].map(name=>[norm(name),short])));
const shortName=name=>shortNames.get(norm(name))||String(name).trim();
function statementFileNames(groups){
 const used=new Set();return groups.map(g=>{
  const base=`${shortName(g.buyer)}&${shortName(g.supplier)}-${g.month.slice(0,4)}年${Number(g.month.slice(5))}月`.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_');
  let name=base,n=1;while(used.has(name.toLowerCase()))name=base+'（'+(++n)+'）';used.add(name.toLowerCase());return name;
 });
}
return {HEADERS,parse,defaults,footer,decimal,decimalText,amountText,date,dateLabel,fileName,statementFileNames};
});
