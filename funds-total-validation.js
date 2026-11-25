(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory();else root.FundsTotalValidation=factory();})(typeof globalThis!=='undefined'?globalThis:this,function(){
'use strict';
const HEADERS=['来源文件','来源工作表','来源行号','供应商','对公付款OA流程号','谈判明细金额合计','OA-申请付款金额','分批总金额','日期计划金额合计','月度支付金额合计','谈判明细有效条数','日期计划非零金额格数','月度有数值月份数','谈判明细减OA差额','分批总金额减OA差额','日期计划减OA差额','月度支付减OA差额','校验结果','需要处理'];
const LIMIT=BigInt(Number.MAX_SAFE_INTEGER),text=v=>String(v??'').trim(),norm=v=>text(v).normalize('NFKC').replace(/\s/g,''),cell=v=>({t:typeof v==='number'?'n':'s',v:v??''});
const invalid=c=>c?.t==='e'||c?.f&&(c.t==='z'||c.v==null||text(c.v)==='');
const hasValue=c=>c&&(c.f||c.t==='e'||c.v!=null&&text(c.v)!=='');
function column(c){let out='';for(let n=c+1;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
function day(y,m,d){const date=new Date(Date.UTC(y,m-1,d));return y>=1900&&y<=9999&&date.getUTCFullYear()===y&&date.getUTCMonth()+1===m&&date.getUTCDate()===d?{year:y,month:m,day:d}:null;}
function headerDate(c,date1904){
 if(!c||invalid(c))return null;
 const v=c.v;
 if(Object.prototype.toString.call(v)==='[object Date]'&&Number.isFinite(v.getTime()))return day(v.getUTCFullYear(),v.getUTCMonth()+1,v.getUTCDate());
 if(typeof v==='number'){
  const format=text(c.z).replace(/"[^"]*"|\\./g,'').replace(/\[[^\]]*\]/g,'');
  const displayed=/^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?$/.exec(norm(c.w));
  if(!(/y/i.test(format)&&/m/i.test(format)&&/d/i.test(format))&&!(displayed&&day(+displayed[1],+displayed[2],+displayed[3])))return null;
  if(!Number.isFinite(v)||v<(date1904?0:1)||v>=2958466||!date1904&&Math.floor(v)===60)return null;
  const date=new Date((date1904?Date.UTC(1904,0,1):Date.UTC(1899,11,v<60?31:30))+Math.floor(v)*86400000);
  return day(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate());
 }
 const m=/^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?$/.exec(norm(v));
 return m?day(+m[1],+m[2],+m[3]):null;
}
function monthHeader(c){if(!c||invalid(c))return null;const m=/^(\d{4})年(\d{1,2})月(?:支付|付款)$/.exec(norm(c.v));return m&&day(+m[1],+m[2],1)?{year:+m[1],month:+m[2]}:null;}
const monthKey=({year,month})=>year+'-'+String(month).padStart(2,'0');
function moneyText(cents){return (cents<0n?'-':'')+(cents<0n?-cents:cents)/100n+'.'+String((cents<0n?-cents:cents)%100n).padStart(2,'0');}
function selectColumns(table){
 const dateColumns=[],monthColumns=[],last=Math.max((table.headers?.length||0)-1,table.range?.e?.c??-1);
 for(let c=0;c<=last;c++){
  const value=table.sheet[column(c)+(table.headerRow+1)];
  if(monthHeader(value))monthColumns.push(c);else if(headerDate(value,table.date1904))dateColumns.push(c);
 }
 return {dateColumns,monthColumns};
}
function parseDetailLines(raw,yearsForMonth,money,sourceAddress){
 const result={sum:null,count:0,complete:true,months:new Map(),payments:[],entries:[],issues:[]};
 const lines=raw.replace(/\s+(?=\d+\s*[、.)）．]\s*(?:(?:\d{4}年)?\d{1,2}月|\d{4}[-/.]))/g,'\n').split(/[\r\n;；]+/).map(text).filter(Boolean);
 if(!lines.length){result.issues.push(sourceAddress+'谈判明细没有可解析的日期和金额');result.complete=false;}
 for(let i=0;i<lines.length;i++){
  const entry={line:i+1};result.entries.push(entry);
  const line=lines[i].replace(/^\d+\s*[、.)）．]\s*(?=(?:\d{4}年)?\d{1,2}月|\d{4}[-/.])/,'');
  const match=/^(?:(\d{4})年)?(\d{1,2})月(?:(\d{1,2})日?)?\s*[－—–:：-]\s*(.+)$/.exec(line)||/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*[－—–:：-]\s*(.+)$/.exec(line),location=sourceAddress+'第'+(i+1)+'笔';
  if(!match){result.issues.push(location+'无法完整解析“'+lines[i]+'”，请填写日期－金额');result.complete=false;continue;}
  const years=yearsForMonth(+match[2]),year=match[1]?+match[1]:years.size===1?[...years][0]:null;
  Object.assign(entry,{year,month:+match[2],day:match[3]===undefined?null:+match[3],validDate:year!==null&&!!day(year,+match[2],match[3]===undefined?1:+match[3])});
  if(!entry.validDate){result.issues.push(location+(year===null?'日期缺少年份且无唯一表头依据':'日期无效')+'：“'+lines[i]+'”');result.complete=false;}
  let raw=text(match[4]);const note=/[（(]([\p{L}\s，,。！!、；;：:]+)[）)]$/u.exec(raw);
  if(note){raw=raw.slice(0,note.index).trim();result.issues.push(location+'括号说明“'+text(note[1])+'”不参与金额汇总');}
  const value=money(cell(raw.normalize('NFKC').replace(/\s*元$/,'')));
  if(value.cents===undefined){result.issues.push(location+'金额无效：“'+match[4]+'”');result.complete=false;continue;}
  entry.cents=value.cents;
  if(!entry.validDate)continue;
  result.sum=(result.sum??0n)+value.cents;result.count++;result.payments.push({year,month:+match[2],day:match[3]===undefined?null:+match[3],cents:value.cents});
  const key=monthKey({year,month:+match[2]});result.months.set(key,(result.months.get(key)??0n)+value.cents);
 }
 return result;
}
function isMonthBandLabel(value){
 const label=norm(value),match=/^(\d{4})年(\d{1,2})月(?:付款)?$/.exec(label);
 return label==='付款月份待核对'||!!(match&&day(+match[1],+match[2],1));
}
function paymentHeaderBands(table){
 // Preview only: keep the uploaded month labels and merges without deriving new ones.
 if(table.headerTopRow==null)return [];
 const bands=[],seen=new Set();
 for(const c of selectColumns(table).dateColumns){
  const merge=(table.sheet['!merges']||[]).find(m=>m.s.r===table.headerTopRow&&m.e.r===table.headerTopRow&&m.s.c<=c&&m.e.c>=c);
  const startCol=merge?.s.c??c,endCol=merge?.e.c??c;
  if(seen.has(startCol))continue;
  seen.add(startCol);bands.push({startCol,endCol,label:text(table.sheet[column(startCol)+(table.headerTopRow+1)]?.v)});
 }
 return bands;
}
function build(table,money){
 const oaColumn=table.map['对公付款0A流程号'];
 const records=table.records.filter(record=>{
  const source=oaColumn===undefined?null:table.sheet[column(oaColumn)+(record.r+1)];
  return invalid(source)||norm(source?.v)!=='不支付';
 });
 const {dateColumns,monthColumns}=selectColumns(table),rows=[],stats={total:records.length,excludedNotPay:table.records.length-records.length,consistent:0,inconsistent:0,unverified:0};
 const address=(c,r)=>column(c)+(r+1),head=c=>table.sheet[address(c,table.headerRow)],dates=new Map(),months=new Map(),monthlyColumns=new Map();
 const remember=(map,value)=>{if(!map.has(value.month))map.set(value.month,new Set());map.get(value.month).add(value.year);};
 for(const c of dateColumns)remember(dates,headerDate(head(c),table.date1904));
 for(const c of monthColumns){const month=monthHeader(head(c)),key=monthKey(month);remember(months,month);if(!monthlyColumns.has(key))monthlyColumns.set(key,[]);monthlyColumns.get(key).push(c);}
 const fields={detail:'谈判明细',oa:'申请付款金额',batch:'分批总金额'};
 const merged=(c,r)=>(table.sheet['!merges']||[]).find(m=>c>=m.s.c&&c<=m.e.c&&r>=m.s.r&&r<=m.e.r&&(m.s.c!==m.e.c||table.records.filter(record=>record.r>=m.s.r&&record.r<=m.e.r).length>1));
 for(const record of records){
  const issues=[],r=record.r,where=c=>address(c,r),group=()=>({sum:null,count:0,complete:true});
  const read=(c,label)=>{
   const source=table.sheet[where(c)],merge=merged(c,r);
   if(merge){issues.push(where(c)+'属于合并单元格'+address(merge.s.c,merge.s.r)+':'+address(merge.e.c,merge.e.r)+'，'+label+'金额归属不清');return {invalid:true};}
   if(invalid(source)){issues.push(where(c)+'公式结果错误或缺失，请重新计算并保存源表');return {invalid:true};}
   const value=money(source);
   if(value.invalid)issues.push(where(c)+'金额“'+text(source?.v)+'”无效或超出安全整数分范围');
   return value;
  };
  const field=name=>{
   const result=group(),c=table.map[name];
   if(c===undefined){issues.push('缺少“'+name+'”表头，无法核验该组金额');result.complete=false;return result;}
   const value=read(c,name);
   if(value.cents!==undefined){result.sum=value.cents;result.count=1;}else{result.complete=false;if(value.blank)issues.push(where(c)+name+'为空，空白不按零处理');}
   return result;
  };
  const detail=group(),detailMonths=new Map(),dc=table.map[fields.detail],detailCell=dc===undefined?null:table.sheet[where(dc)];
  if(dc===undefined){issues.push('缺少“谈判明细”表头');detail.complete=false;}
  else if(invalid(detailCell)||merged(dc,r)){issues.push(where(dc)+'谈判明细公式结果错误、缺失或跨记录合并，请核实来源');detail.complete=false;}
  else if(typeof detailCell?.v!=='string'||!text(detailCell.v)){issues.push(where(dc)+'谈判明细为空或不是日期金额文本，空白不按零处理');detail.complete=false;}
  else{
   const parsed=parseDetailLines(detailCell.v,month=>new Set([...(months.get(month)||[]),...(dates.get(month)||[])]),money,where(dc));
   Object.assign(detail,{sum:parsed.sum,count:parsed.count,complete:parsed.complete});
   for(const [key,value]of parsed.months)detailMonths.set(key,value);issues.push(...parsed.issues);
  }
  const range=(columns,label,nonzero)=>{
   const result={...group(),values:new Map()};
   if(!columns.length){issues.push('未找到'+label+'表头，无法核验该组金额');result.complete=false;return result;}
   for(const c of columns){const value=read(c,label);result.values.set(c,value);if(value.invalid)result.complete=false;if(value.cents!==undefined){result.sum=(result.sum??0n)+value.cents;if(!nonzero||value.cents!==0n)result.count++;}}
   if(result.sum===null){issues.push(label+'（'+columns.map(where).join('、')+'）没有有效金额，空白不按零处理');result.complete=false;}
   return result;
  };
  const oa=field(fields.oa),batch=field(fields.batch),plan=range(dateColumns,'日期计划',true),monthly=range(monthColumns,'月度支付',false),groups=[detail,oa,batch,plan,monthly];
  const labels=HEADERS.slice(5,10);
  const amount=(value,label)=>{
   if(value===null)return cell('');
   if(value>LIMIT||value< -LIMIT){issues.push(label+'超出安全整数分范围，未输出近似金额');return cell('');}
   return {t:'n',v:Number(value)/100,z:'#,##0.00'};
  };
  const values=groups.map((g,i)=>{const value=amount(g.sum,labels[i]);if(value.v==='')g.complete=false;if(!g.complete&&g.sum!==null)issues.push(labels[i]+'仅保留可解析有效金额小计，仍需核实完整金额');return value;});
  const differences=[detail,batch,plan,monthly].map((g,i)=>g.complete&&oa.complete?amount(g.sum-oa.sum,HEADERS[13+i]):cell(''));
  let monthIncomplete=false,monthMismatch=false;
  if(detail.complete)for(const key of [...new Set([...detailMonths.keys(),...monthlyColumns.keys()])].sort()){
   const [year,month]=key.split('-'),label=year+'年'+Number(month)+'月',columns=monthlyColumns.get(key)||[],detailSum=detailMonths.get(key)??0n;
   if(columns.length>1){issues.push(label+'月度支付表头重复（'+columns.map(c=>address(c,table.headerRow)).join('、')+'），请明确唯一金额来源');monthIncomplete=true;continue;}
   if(detailSum>LIMIT||detailSum< -LIMIT){issues.push(label+'谈判明细金额超出安全整数分范围，无法按月核对');monthIncomplete=true;continue;}
   if(!columns.length){issues.push(label+'谈判明细'+where(dc)+'合计'+moneyText(detailSum)+'，缺少对应月度支付列，无法按月核对');monthIncomplete=true;continue;}
   const c=columns[0],value=monthly.values.get(c);
   if(value.cents===undefined){
    if(!detailMonths.has(key)&&value.blank)continue;
    issues.push(label+'月度支付'+where(c)+(value.blank?'为空白':'金额无效')+'，无法与谈判明细'+where(dc)+'按月核对');monthIncomplete=true;continue;
   }
   const difference=detailSum-value.cents;
   if(difference>LIMIT||difference< -LIMIT){issues.push(label+'月度差额超出安全整数分范围，未输出近似金额');monthIncomplete=true;continue;}
   if(difference!==0n){issues.push(label+'：谈判明细'+where(dc)+'合计'+moneyText(detailSum)+'；月度支付'+where(c)+'金额'+moneyText(value.cents)+'；差额'+moneyText(difference)+'（谈判明细减月度支付）');monthMismatch=true;}
  }
  const totalMismatch=differences.some(c=>typeof c.v==='number'&&c.v!==0),incomplete=groups.some(g=>!g.complete)||differences.some(c=>c.v==='')||monthIncomplete,mismatch=totalMismatch||monthMismatch,status=incomplete?'待核对':mismatch?'不一致':'一致';
  if(totalMismatch)issues.push('差额为对应组金额减OA申请付款金额；请核对各来源值，不自动改动原记录');
  if(status!=='一致')issues.unshift('金额来源：谈判明细'+(dc===undefined?'未找到':where(dc))+'；OA'+(table.map[fields.oa]===undefined?'未找到':where(table.map[fields.oa]))+'；分批总金额'+(table.map[fields.batch]===undefined?'未找到':where(table.map[fields.batch]))+'；日期计划'+dateColumns.map(where).join('、')+'；月度支付'+monthColumns.map(where).join('、'));
  stats[incomplete?'unverified':mismatch?'inconsistent':'consistent']++;
  rows.push([cell(record.source?.file??table.name),cell(record.source?.sheet??table.sheetName),cell(r+1),cell(record.values['供应商名称']?.v),cell(record.values['对公付款0A流程号']?.v),...values,cell(detail.count),cell(plan.count),cell(monthly.count),...differences,cell(status),cell([...new Set(issues)].join('\n'))]);
 }
 return {headers:HEADERS.slice(),rows,stats};
}
return {HEADERS,selectColumns,paymentHeaderBands,isMonthBandLabel,parseDetailLines,headerDate,monthHeader,build};
});
