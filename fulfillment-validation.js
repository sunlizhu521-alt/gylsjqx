(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FulfillmentValidation=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const FIELDS=['未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量'];
const norm=value=>String(value??'').replace(/\s/g,'');
const blank=cell=>!cell||(!cell.f&&(cell.v==null||cell.v===''));
const finite=value=>typeof value==='number'&&Number.isFinite(value)&&Math.abs(value)<=Number.MAX_SAFE_INTEGER;
const equal=(a,b)=>Number.isInteger(a)&&Number.isInteger(b)?a===b:Math.abs(a-b)<=Math.min(1e-9,Number.EPSILON*8*Math.max(1,Math.abs(a),Math.abs(b)));
function sum(values){let total=0,correction=0;for(const value of values){const next=value-correction,t=total+next;correction=(t-total)-next;total=t;}return total;}
function number(cell){
 if(blank(cell))return 0;
 if(cell.t==='e')throw Error('单元格为错误值，请修正源表');
 if(cell.t&&cell.t!=='n')throw Error('数量“'+String(cell.v??'')+'”为文本或其他非数值类型，请在源表转换为数字');
 if(!finite(cell.v))throw Error('数量不是可精确核验的有限数字');
 return cell.v;
}

// Only this small arithmetic grammar is accepted; workbook formulas are never executed as JavaScript.
function parse(formula,valuesOnly=false){
 const text=String(formula).replace(/^=/,'');let pos=0,depth=0;
 const tokens=[];
 while(pos<text.length){const part=text.slice(pos),space=part.match(/^\s+/);if(space){pos+=space[0].length;continue;}
  const match=part.match(valuesOnly?/^(?:"(?:[^"]|"")*"|'(?:[^']|'')*'|[A-Za-z_0-9.\u0080-\uFFFF]+(?=\s*!)|\$?[A-Za-z]{1,3}\$?[1-9]\d*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[A-Za-z_\u0080-\uFFFF][A-Za-z_0-9.\u0080-\uFFFF]*|[+\-*/(),:!])/:/^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[A-Za-z_][A-Za-z_0-9.]*|[+\-*/(),:])/);
  if(!match)throw Error('公式含暂不支持的表达式，请在 WPS/Excel 核对');tokens.push(match[0]);pos+=match[0].length;
  if(tokens.length>10000)throw Error('公式过长，无法可靠核验');
 }
 let index=0;const peek=()=>tokens[index],take=()=>tokens[index++];
 function primary(){
  if(++depth>150)throw Error('公式嵌套过深，无法可靠核验');
  let node,token=take(),sheet;
  if(valuesOnly&&peek()==='!'){sheet=token.startsWith("'")?token.slice(1,-1).replace(/''/g,"'"):token;take();token=take();}
  if(sheet!==undefined&&!/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(token||''))throw Error('跨表引用必须指向单元格');
  if(token==='+'||token==='-')node={op:token==='+'?'positive':'negative',value:primary()};
  else if(token==='('){node=expression();if(take()!==')')throw Error('公式括号不完整');}
  else if(/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(token||'')){
   const first=token.replace(/\$/g,'').toUpperCase(),at=X.utils.decode_cell(first);
   if(at.c>16383||at.r>1048575)throw Error('公式引用超出工作表范围');
   node={op:'ref',first,sheet};
   if(peek()===':'){take();const last=take();if(!/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(last||''))throw Error('公式引用范围无效');node={op:'range',first,last:last.replace(/\$/g,'').toUpperCase(),sheet};}
  }else if(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(token||''))node={op:'number',value:Number(token)};
  else if(valuesOnly&&token?.startsWith('"'))node={op:'literal',value:token.slice(1,-1).replace(/""/g,'"')};
  else if(valuesOnly&&/^(TRUE|FALSE)$/i.test(token||'')){node={op:'literal',value:token.toUpperCase()==='TRUE'};if(peek()==='('){take();if(take()!==')')throw Error('布尔函数不能带参数');}}
  else if(/^[A-Za-z_]/.test(token||'')){
   const name=token.toUpperCase();if(!['SUM','SUBTOTAL'].includes(name))throw Error('暂不支持公式 '+name+'，请在 WPS/Excel 核对');
   if(take()!=='(')throw Error('公式函数格式无效');const args=[];
   if(peek()!==')'){do{args.push(expression());if(peek()!==',')break;take();}while(true);}
   if(take()!==')')throw Error('公式函数参数不完整');node={op:name,args};
  }else throw Error('公式表达式无效');
  depth--;return node;
 }
 function product(){let node=primary();while(peek()==='*'||peek()==='/'){const op=take();node={op,left:node,right:primary()};}return node;}
 function expression(){let node=product();while(peek()==='+'||peek()==='-'){const op=take();node={op,left:node,right:product()};}return node;}
 if(!tokens.length)throw Error('公式为空');const tree=expression();if(index!==tokens.length)throw Error('公式含暂不支持的表达式，请在 WPS/Excel 核对');return tree;
}

function evaluator(sheet,onIssue,onFormula,options={}){
 const cache=new Map(),visiting=new Set();let operations=0;
 const numeric=value=>{if(!finite(value))throw Error('公式运算引用了非数字或超出可精确核验范围');return value;};
 const scalar=source=>{if(!options.valuesOnly)return number(source);if(blank(source))return 0;if(source.t==='e')throw Error('单元格为错误值，请修正源表');if(source.t==='s'&&typeof source.v==='string'||source.t==='b'&&typeof source.v==='boolean')return source.v;return number(source);};
 const activeFilter=sheet['!fulfillmentFilterActive']===true||sheet['!filterMode']===true||Object.keys(sheet['!autofilter']||{}).some(key=>key!=='ref'&&sheet['!autofilter'][key]!=null);
 function cell(address,field){
  if(cache.has(address))return cache.get(address);
  if(visiting.has(address))throw Error('公式存在循环引用（'+address+'）');
  if(visiting.size>=150)throw Error('公式引用层数过多，无法可靠核验');
  const source=sheet[address];if(!source?.f){try{return scalar(source);}catch(error){error.cell=address;if(options.name)error.sheet=options.name;throw error;}}
  onFormula(address);visiting.add(address);
  let value;
  try{
   value=evaluate(parse(source.f,options.valuesOnly),field);
   if(!(options.valuesOnly&&['string','boolean'].includes(typeof value))&&!finite(value))throw Error('公式结果超出可精确核验范围');
   if(!options.valuesOnly){try{const cached=number(source);if(!equal(cached,value))onIssue(address,field,'公式保存值为 '+cached+'，按当前引用计算为 '+value+'；请在 WPS/Excel 重算并保存后重新上传','formula-cache');}
   catch(error){onIssue(address,field,'公式结果不可用：'+error.message,'formula-result');}}
   cache.set(address,value);return value;
  }catch(error){error.cell??=address;if(options.name)error.sheet??=options.name;throw error;
  }finally{visiting.delete(address);}
 }
 function references(node,subtotal,field){
  if(node.sheet!==undefined){if(!options.resolve)throw Error('暂不支持跨表引用');return options.resolve(node.sheet).references({...node,sheet:undefined},subtotal,field);}
  if(subtotal&&activeFilter)throw Error('工作表存在筛选条件，SUBTOTAL 结果需在 WPS/Excel 清除筛选后核验');
  if(subtotal&&sheet['!autofilter']&&sheet['!fulfillmentFilterActive']!==false)throw Error('工作表筛选条件未能确认，SUBTOTAL 结果需重新核验');
  const range=X.utils.decode_range(node.op==='range'?node.first+':'+node.last:node.first);
  if(range.e.c<range.s.c||range.e.r<range.s.r)throw Error('公式引用范围方向无效，请在 WPS/Excel 核对');
  if(range.e.c>16383||range.e.r>1048575||(range.e.c-range.s.c+1)*(range.e.r-range.s.r+1)>100000)throw Error('公式引用范围过大，无法可靠核验');
  const values=[];
  for(let r=range.s.r;r<=range.e.r;r++)for(let c=range.s.c;c<=range.e.c;c++){
   if(++operations>1000000)throw Error('公式引用过多，无法可靠核验');
   const address=X.utils.encode_cell({r,c}),source=sheet[address];
   if(subtotal&&source?.f&&/\bSUBTOTAL\s*\(/i.test(source.f))continue;
   if(subtotal===109&&sheet['!rows']?.[r]?.hidden)continue;
   if(subtotal===9&&sheet['!rows']?.[r]?.hidden&&sheet['!autofilter']&&sheet['!fulfillmentFilterActive']!==false)throw Error('隐藏行的筛选状态无法确认，SUBTOTAL(9) 需在 WPS/Excel 核对');
   // SUM/SUBTOTAL ignore text in referenced ranges, just as spreadsheet SUM does.
   if(!source?.f&&source?.t!=='e'&&(blank(source)||typeof source?.v==='string'||typeof source?.v==='boolean'))continue;
   const value=cell(address,field);if(options.valuesOnly&&['string','boolean'].includes(typeof value))continue;values.push(value);
  }
  return sum(values);
 }
 function evaluate(node,field){
  if(node.op==='number')return node.value;
  if(node.op==='literal')return node.value;
  if(node.op==='ref')return node.sheet===undefined?cell(node.first,field):options.resolve(node.sheet)(node.first,field);
  if(node.op==='range')throw Error('范围引用只能用于 SUM/SUBTOTAL');
  if(node.op==='positive')return numeric(evaluate(node.value,field));
  if(node.op==='negative')return -numeric(evaluate(node.value,field));
  if(node.op==='SUM'||node.op==='SUBTOTAL'){
   let args=node.args,subtotal=0;
   if(node.op==='SUBTOTAL'){
    subtotal=args.length?evaluate(args[0],field):0;args=args.slice(1);
    if(![9,109].includes(subtotal)||!args.length)throw Error('仅支持 SUBTOTAL(9) 和 SUBTOTAL(109) 求和');
    if(!options.valuesOnly&&activeFilter)throw Error('工作表存在筛选条件，SUBTOTAL 结果需在 WPS/Excel 清除筛选后核验');
    if(!options.valuesOnly&&sheet['!autofilter']&&sheet['!fulfillmentFilterActive']!==false)throw Error('工作表筛选条件未能确认，SUBTOTAL 结果需重新核验');
    if(args.some(arg=>!['ref','range'].includes(arg.op)))throw Error('SUBTOTAL 参数必须是单元格引用');
   }
   return sum(args.map(arg=>['ref','range'].includes(arg.op)?references(arg,subtotal,field):numeric(evaluate(arg,field))));
  }
  const left=numeric(evaluate(node.left,field)),right=numeric(evaluate(node.right,field));
  if(node.op==='+')return left+right;if(node.op==='-')return left-right;if(node.op==='*')return left*right;
  if(right===0)throw Error('公式除数为零');return left/right;
 }
 cell.references=references;return cell;
}

function materialize(book){
 const updates=[],issues=[],readers=new Map(),sheets=book.Sheets||{},names=Object.keys(sheets),copy={...book,Sheets:{...sheets}};let formulaCount=0,correctedCount=0;
 const resolve=name=>{const matches=names.filter(item=>item.toLowerCase()===String(name).toLowerCase());if(matches.length!==1)throw Error(matches.length?'引用的工作表名称不唯一：'+name:'引用的工作表不存在：'+name);const actual=matches[0];if(!readers.has(actual))readers.set(actual,evaluator(sheets[actual],()=>{},()=>{},{valuesOnly:true,name:actual,resolve}));return readers.get(actual);};
 for(const name of names){copy.Sheets[name]={...sheets[name]};for(const [address,source] of Object.entries(sheets[name])){
  if(!/^[A-Z]+[1-9]\d*$/.test(address)||!source||source.f==null&&source.F==null&&!source.D)continue;
  formulaCount++;
  try{
   if(!source.f)throw Error('公式内容缺失，无法重新计算');
   if(source.D||source.F&&source.F!==address&&source.F!==address+':'+address)throw Error('数组公式需要先在 WPS/Excel 转为数值，不能仅使用保存结果');
   const value=resolve(name)(address),type=typeof value==='string'?'s':typeof value==='boolean'?'b':'n',cell={...source,t:type,v:value};
   for(const key of ['f','F','D','w','r','h'])delete cell[key];
   if(source.t!==type||(type==='n'?!finite(source.v)||!equal(source.v,value):source.v!==value))correctedCount++;
   copy.Sheets[name][address]=cell;updates.push({sheet:name,cell:address,value,type});
  }catch(error){const reference=error.cell&&(error.cell!==address||error.sheet&&error.sheet!==name)?'引用 '+(error.sheet||name)+'!'+error.cell+'：':'';issues.push({sheet:name,cell:address,field:'',message:reference+error.message,kind:'formula'});}
 }}
 return {book:copy,updates,issues,formulaCount,correctedCount};
}

function validate(book,result,core){
 const issues=[],checks=[],issueKeys=new Set(),formulas=new Set();
 const issue=(sheet,cell,field,message,kind)=>{const key=[sheet,cell,field,kind,message].join('\0');if(!issueKeys.has(key)){issueKeys.add(key);issues.push({sheet,cell,field,message,kind});}};
 const total=(values,sheet,field)=>{const value=sum(values);if(finite(value))return value;issue(sheet,'',field,'合计超出可精确核验范围','total');return null;};
 const selected=[];
 for(const person of core.SHEETS){
  const matches=Object.keys(book.Sheets||{}).filter(name=>new RegExp('^'+person+'(?:[0-9]{4}|[0-9]{8})?$').test(norm(name)));
  if(matches.length!==1){issue(person,'','',matches.length?'匹配到多张来源工作表，无法核验':'缺少来源工作表','sheet');selected.push({name:person,invalid:true});}
  else selected.push({name:matches[0],sheet:book.Sheets[matches[0]]});
 }
 let summaries=[];try{summaries=core.summaryRows(result.rows||[]);}catch(error){issue('汇总表','','','汇总列生成失败：'+error.message,'summary');}
 const output=new Map(),rowCount=(result.rows||[]).length;
 if((result.sources||[]).length!==rowCount||summaries.length!==rowCount)issue('汇总表','','','汇总记录与来源映射的数量不一致','mapping');
 for(let i=0;i<rowCount;i++){
  const source=result.sources?.[i];
  if(!source||!selected.some(item=>!item.invalid&&item.name===source.sheet)||!Number.isInteger(source.row)||source.row<1){issue(source?.sheet||'汇总表','', '', '第 '+(i+2)+' 行来源无效或不属于本次四人工作表','mapping');continue;}
  const key=source.sheet+'\0'+source.row;if(!output.has(key))output.set(key,[]);output.get(key).push(i);
 }
 const grandSource=FIELDS.map(()=>[]),grandSummary=FIELDS.map(()=>[]),grandInvalid=FIELDS.map(()=>false);
 for(const entry of selected){
  const {name,sheet}=entry,sourceValues=FIELDS.map(()=>[]),summaryValues=FIELDS.map(()=>[]),invalid=FIELDS.map(()=>false),expected=new Set();
  let header=-1,map={},range;
  if(!entry.invalid){
   try{range=X.utils.decode_range(sheet['!ref']||'A1');}catch{entry.invalid=true;}
   if(range)for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){
    const found={};for(let c=range.s.c;c<=range.e.c;c++){const key=norm(sheet[X.utils.encode_cell({r,c})]?.v);if(key)(found[key]??=[]).push(c);}
    if(found['物料编码']&&found['事业部']){header=r;map=found;break;}
   }
   if(header<0){issue(name,'','','找不到明细表头，无法核验','header');entry.invalid=true;}
  }
  if(entry.invalid)invalid.fill(true);
  else{
   FIELDS.forEach((field,index)=>{if(map[field]?.length!==1){invalid[index]=true;issue(name,'',field,map[field]?'数量字段重复':'缺少数量字段','header');}});
   const actualField=(address,fallback)=>FIELDS.find(field=>map[field]?.[0]===X.utils.decode_cell(address).c)||fallback;
   const read=evaluator(sheet,(cell,field,message,kind)=>issue(name,cell,actualField(cell,field),message,kind),address=>formulas.add(name+'\0'+address));
   for(let r=header+1;r<=range.e.r;r++){
    let nonempty=false;for(let c=range.s.c;c<=range.e.c;c++){const original=sheet[X.utils.encode_cell({r,c})];if(original?.f||(original?.v!=null&&String(original.v).trim()!=='')){nonempty=true;break;}}
    if(!nonempty)continue;
    const key=name+'\0'+(r+1);expected.add(key);const indexes=output.get(key)||[];
    if(indexes.length!==1)issue(name,String(r+1),'',indexes.length?'源表该行被重复汇总 '+indexes.length+' 次':'源表该行未进入汇总','mapping');
    FIELDS.forEach((field,index)=>{
     if(map[field]?.length!==1)return;const address=X.utils.encode_cell({r,c:map[field][0]});let value=null;
     try{value=read(address,field);sourceValues[index].push(value);}catch(error){invalid[index]=true;issue(name,error.cell||address,actualField(error.cell||address,field),error.message,'source');}
     for(const i of indexes){
      const outputCell=summaries[i]?.[core.SUMMARY_COLUMNS.indexOf(field)];
      try{if(!outputCell)throw Error('汇总数量单元格缺失');const actual=number(outputCell);summaryValues[index].push(actual);if(blank(sheet[address])!==blank(outputCell))issue(name,address,field,'源表的空白与数字状态在汇总第 '+(i+2)+' 行发生变化','value');if(value!==null&&!equal(value,actual))issue(name,address,field,'源表计算值为 '+value+'，汇总第 '+(i+2)+' 行为 '+actual,'value');}
      catch(error){invalid[index]=true;issue(name,address,field,'汇总第 '+(i+2)+' 行：'+error.message,'summary');}
     }
    });
   }
   for(const [key,indexes] of output){if(!key.startsWith(name+'\0')||expected.has(key))continue;issue(name,key.split('\0')[1],'','汇总引用了表头、空行或不存在的源行','mapping');for(const i of indexes)FIELDS.forEach((field,index)=>{try{summaryValues[index].push(number(summaries[i]?.[core.SUMMARY_COLUMNS.indexOf(field)]));}catch{invalid[index]=true;}});}
   const declared=(result.counts||[]).filter(item=>item.sheet===name);if(declared.length!==1||declared[0].count!==expected.size)issue(name,'','','来源记录数应为 '+expected.size+'，汇总计数不一致','mapping');
  }
  FIELDS.forEach((field,index)=>{
   const sourceTotal=invalid[index]?null:total(sourceValues[index],name,field),summaryTotal=invalid[index]?null:total(summaryValues[index],name,field);
   const difference=sourceTotal===null||summaryTotal===null?null:equal(sourceTotal,summaryTotal)?0:summaryTotal-sourceTotal;
   const failed=issues.some(item=>item.sheet===name&&(!item.field||item.field===field));
   checks.push({sheet:name,field,sourceTotal,summaryTotal,difference,status:sourceTotal===null||summaryTotal===null?'待核验':failed||!equal(sourceTotal,summaryTotal)?'不一致':'通过'});
   if(sourceTotal===null||summaryTotal===null){grandInvalid[index]=true;return;}grandSource[index].push(sourceTotal);grandSummary[index].push(summaryTotal);
  });
 }
 FIELDS.forEach((field,index)=>{const sourceTotal=grandInvalid[index]?null:total(grandSource[index],'合计',field),summaryTotal=grandInvalid[index]?null:total(grandSummary[index],'合计',field);checks.push({sheet:'合计',field,sourceTotal,summaryTotal,difference:sourceTotal===null||summaryTotal===null?null:equal(sourceTotal,summaryTotal)?0:summaryTotal-sourceTotal,status:sourceTotal===null||summaryTotal===null?'待核验':issues.some(item=>!item.field||item.field===field)||!equal(sourceTotal,summaryTotal)?'不一致':'通过'});});
 return {passed:issues.length===0&&checks.every(check=>check.status==='通过'),checks,issues,formulaCount:formulas.size};
}
return {validate,materialize,FIELDS};
});
