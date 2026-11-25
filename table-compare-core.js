(function(root,factory){
  if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));
  else root.TableCompareCore=factory(root.XLSX);
})(typeof globalThis!=='undefined'?globalThis:this,function(X){
  'use strict';
  const HEADERS=['分类','采购跟单员','供应商','付款主体','对公付款OA流程号','摘要','旧版申请付款金额','新版申请付款金额','差额','变化','原备注','比对说明','旧版来源','新版来源','备注'];
  const FIELDS={category:['分类','付款分类'],owner:['采购跟单员','创跟单员','跟单员'],supplier:['供应商','供应商名称'],entity:['付款主体'],oa:['对公付款0A流程号','对公付款OA流程号','对公OA流程号','付款OA号','OA'],summary:['摘要'],amount:['OA-申请付款金额','申请付款金额','本次申请付款金额'],remark:['备注']};
  const REQUIRED=['supplier','entity','oa','amount'];
  const LABELS={category:'分类',owner:'采购跟单员',supplier:'供应商',entity:'付款主体',oa:'对公付款OA流程号',summary:'摘要',amount:'申请付款金额',remark:'原备注'};
  const normalized=value=>String(value??'').normalize('NFKC').replace(/\s+/g,'').toUpperCase();
  const aliases=new Map(Object.entries(FIELDS).flatMap(([field,names])=>names.map(name=>[normalized(name),field])));
  const str=value=>({t:'s',v:String(value??''),z:'@'});
  const number=cents=>({t:'n',v:cents/100,z:'#,##0.00'});
  const key=values=>JSON.stringify(values);
  const hasValue=cell=>!!cell&&(cell.f!=null||cell.v!=null&&String(cell.v).trim()!=='');
  function entries(sheet){
    return Object.entries(sheet).flatMap(([address,cell])=>{
      if(!/^[A-Z]+[1-9]\d*$/.test(address)||!hasValue(cell))return [];
      const position=X.utils.decode_cell(address);
      return [{address,cell,...position}];
    });
  }
  function discover(book,filename='台账.xlsx'){
    const result=[];
    for(const name of book.SheetNames||[]){
      const rows=new Map();
      for(const {r,c,cell}of entries(book.Sheets[name]||{})){
        const field=aliases.get(normalized(cell.v));
        if(!field)continue;
        if(!rows.has(r))rows.set(r,{});
        const fields=rows.get(r);(fields[field]??=[]).push(c);
      }
      const candidates=[...rows].filter(([,fields])=>REQUIRED.every(field=>fields[field]));
      if(candidates.length>1)throw Error(filename+' / '+name+'：发现多个业务表头（第'+candidates.map(([r])=>r+1).join('、')+'行），请将需对比的明细放在独立工作表');
      if(!candidates.length)continue;
      const [headerRow,columns]=candidates[0];
      const repeated=Object.keys(columns).filter(field=>columns[field].length>1);
      if(repeated.length)throw Error(filename+' / '+name+' 第'+(headerRow+1)+'行：表头重复 '+repeated.map(field=>LABELS[field]).join('、')+'，请明确保留一列');
      result.push({name,headerRow,fields:Object.fromEntries(Object.entries(columns).map(([field,cols])=>[field,cols[0]]))});
    }
    return result;
  }
  function valid(cell,label){
    if(cell?.t==='e')throw Error(label+'：Excel 错误 '+(cell.w||cell.v)+'，请修正源表后重新上传');
    if(cell?.f!=null&&(cell.v==null||cell.t==='z'||cell.t==='n'&&(typeof cell.v!=='number'||!Number.isFinite(cell.v))))throw Error(label+'：公式缺少有效计算结果，请在 WPS/Excel 重新计算并保存');
  }
  function display(cell,date1904){
    if(!cell||cell.v==null||cell.t==='z')return '';
    if(cell.t==='d'||cell.v instanceof Date){
      const date=new Date(cell.v);return Number.isNaN(date.getTime())?String(cell.v):date.toISOString().slice(0,10);
    }
    if(cell.t==='n'&&cell.z){
      if(X.SSF.is_date(cell.z))return X.SSF.format(cell.z,cell.v,{date1904});
      if(/^[0]+$/.test(cell.z))return X.SSF.format(cell.z,cell.v);
    }
    return String(cell.v);
  }
  function money(cell){
    if(!cell||cell.v==null||cell.t==='z'||String(cell.v).trim()==='')return {cents:null,error:'申请付款金额为空'};
    if(cell.t==='b'||cell.t==='d'||cell.v instanceof Date||typeof cell.v==='number'&&!Number.isFinite(cell.v))return {cents:null,error:'申请付款金额异常'};
    let value=String(cell.v).normalize('NFKC').trim().replace(/^(?:人民币|RMB|CNY|[¥￥])\s*/i,'').replace(/\s*元$/,'');
    if(/^\(.*\)$/.test(value))value='-'+value.slice(1,-1);
    if(typeof cell.v==='number'&&/[eE]/.test(value)){
      const match=/^(-?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/.exec(value);
      if(match){
        const digits=match[2]+(match[3]||''),point=match[2].length+Number(match[4]);
        value=match[1]+(point<=0?'0.'+'0'.repeat(-point)+digits:point>=digits.length?digits+'0'.repeat(point-digits.length):digits.slice(0,point)+'.'+digits.slice(point));
      }
    }
    if(!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(value))return {cents:null,error:'申请付款金额无法识别“'+String(cell.v)+'”'};
    const negative=value.startsWith('-'),parts=value.replace(/^[+-]/,'').replace(/,/g,'').split('.');
    const fraction=parts[1]||'';
    const absolute=BigInt(parts[0])*100n+BigInt((fraction+'00').slice(0,2))+(Number(fraction[2]||'0')>=5?1n:0n);
    if(absolute>BigInt(Number.MAX_SAFE_INTEGER))return {cents:null,error:'申请付款金额超出安全范围'};
    return {cents:Number(negative?-absolute:absolute),error:''};
  }
  function oa(value){
    const raw=String(value).trim(),text=normalized(raw);
    if(!text||/^(?:未提交(?:OA)?|未申请(?:OA)?|未交OA|无|无OA|\/|-)$/.test(text))return {code:null,error:''};
    if(/[;,，；、/\r\n]/.test(raw)||raw.split(/\s+/).filter(Boolean).length>1)return {code:null,error:'OA号包含多个编号或无法明确识别'};
    if(!/^[A-Z0-9][A-Z0-9_-]*$/.test(text)||!/[0-9]/.test(text))return {code:null,error:'OA号无法识别“'+raw+'”'};
    return {code:text,error:''};
  }
  function read(source){
    if(!source?.book)throw Error('请先上传旧版和新版台账');
    const filename=source.filename||'台账.xlsx',book=source.book,candidates=discover(book,filename);
    if(!candidates.length)throw Error(filename+'：找不到付款台账表头，需要供应商、付款主体、对公付款OA流程号、申请付款金额');
    const selected=source.sheetName?candidates.find(candidate=>candidate.name===source.sheetName):candidates.length===1?candidates[0]:null;
    if(!selected)throw Error(filename+'：请选择一个有效工作表（'+candidates.map(candidate=>candidate.name).join('、')+'）');
    const {name,headerRow,fields}=selected,sheet=book.Sheets[name],date1904=!!book.Workbook?.WBProps?.date1904,warnings=[];
    const missing=Object.keys(FIELDS).filter(field=>fields[field]==null);
    if(missing.length)warnings.push(filename+' / '+name+'：缺少 '+missing.map(field=>LABELS[field]).join('、')+'，对应内容留空');
    const sourceRows=[...new Set(entries(sheet).filter(entry=>entry.r>headerRow&&Object.values(fields).includes(entry.c)).map(entry=>entry.r))].sort((a,b)=>a-b);
    const totalLabels=new Set(['合计','小计','总计']);
    const records=sourceRows.flatMap(row=>{
      const values={},cells={};
      for(const field of Object.keys(FIELDS)){
        const address=fields[field]==null?'':X.utils.encode_cell({r:row,c:fields[field]}),cell=address?sheet[address]:null;
        cells[field]=cell;values[field]=display(cell,date1904);
      }
      const location=filename+' / '+name+' 第'+(row+1)+'行';
      const summaryLabel=['category','owner','supplier','entity','oa','summary','remark'].some(field=>totalLabels.has(values[field].trim()));
      const noIdentity=['supplier','entity','oa'].every(field=>!values[field].trim()||totalLabels.has(values[field].trim()));
      if(summaryLabel&&noIdentity){warnings.push(location+'：忽略无供应商、付款主体和OA业务身份的明确合计／小计／总计行，避免重复计入金额');return [];}
      for(const field of Object.keys(FIELDS))if(fields[field]!=null)valid(cells[field],filename+' / '+name+'!'+X.utils.encode_cell({r:row,c:fields[field]})+'（'+LABELS[field]+'）');
      const amount=money(cells.amount),flow=oa(values.oa),issues=[];
      if(!values.supplier.trim())issues.push('供应商为空');
      if(!values.entity.trim())issues.push('付款主体为空');
      if(amount.error)issues.push(amount.error);
      if(flow.error)issues.push(flow.error);
      return [{values,cents:amount.cents,code:flow.code,location,issues,pair:null,review:[],identity:key([values.supplier.trim(),values.entity.trim()])}];
    });
    return {records,warnings};
  }
  function group(records,selector){
    const groups=new Map();
    for(const record of records){const name=selector(record);if(name==null)continue;if(!groups.has(name))groups.set(name,[]);groups.get(name).push(record);}
    return groups;
  }
  function mark(records,message){for(const record of records)if(!record.review.includes(message))record.review.push(message);}
  function connect(oldRecord,newRecord){oldRecord.pair=newRecord;newRecord.pair=oldRecord;}
  const available=record=>!record.pair&&!record.review.length;
  function match(oldRecords,newRecords){
    const codeKey=record=>record.code&&record.values.supplier.trim()&&record.values.entity.trim()?key([record.identity,record.code]):null;
    const oldGroups=group(oldRecords,codeKey),newGroups=group(newRecords,codeKey);
    for(const [name,olds]of oldGroups){
      const news=newGroups.get(name);if(!news)continue;
      // Match equal amounts first; consuming one row at a time preserves duplicates.
      for(const old of olds){if(old.issues.length)continue;const found=news.find(next=>!next.pair&&!next.issues.length&&next.cents===old.cents);if(found)connect(old,found);}
      const a=olds.filter(record=>!record.pair),b=news.filter(record=>!record.pair);
      if(!a.length||!b.length)continue;
      if(a.length===1&&b.length===1&&!a[0].issues.length&&!b[0].issues.length)connect(a[0],b[0]);
      else mark([...a,...b],'同一OA号、供应商和付款主体：旧版剩余'+a.length+'笔，新版剩余'+b.length+'笔，无法唯一对应；请核对本组明细，来源见各行来源列');
    }
    // An invalid identity/amount cannot make its otherwise plausible counterpart look new.
    for(const [records,other]of [[oldRecords,newRecords],[newRecords,oldRecords]])for(const record of records){
      if(!record.issues.length||!record.code||record.pair)continue;
      const candidates=other.filter(next=>available(next)&&next.code===record.code);
      if(candidates.length)mark([record,...candidates],'相同OA号的来源记录有缺失或异常字段，请核对 '+record.location);
    }
    for(const [records,other]of [[oldRecords,newRecords],[newRecords,oldRecords]])for(const record of records){
      if(!record.issues.length||record.pair||!record.values.supplier.trim()||!record.values.entity.trim())continue;
      const candidates=other.filter(next=>available(next)&&next.identity===record.identity&&(!record.code||!next.code)&&(record.cents==null||next.cents==null||record.cents===next.cents));
      if(candidates.length)mark([record,...candidates],'同一供应商和付款主体的来源记录缺有效OA号且有异常，无法确认新增或减少；请核对 '+record.location);
    }
    const fallbackKey=record=>available(record)&&!record.issues.length?key([record.identity,record.cents]):null;
    const fallbackOld=group(oldRecords,fallbackKey),fallbackNew=group(newRecords,fallbackKey);
    for(const [name,olds]of fallbackOld){
      const news=fallbackNew.get(name);if(!news)continue;
      const oldMissing=olds.filter(record=>!record.code),newMissing=news.filter(record=>!record.code);
      for(const old of olds){
        const candidates=old.code?newMissing:news;
        if(candidates.length!==1)continue;
        const next=candidates[0],reverseCount=next.code?oldMissing.length:olds.length;
        if(reverseCount===1)connect(old,next);
      }
      const ambiguous=[...olds.filter(record=>!record.pair&&(record.code?newMissing.length:news.length)),...news.filter(record=>!record.pair&&(record.code?oldMissing.length:olds.length))];
      if(ambiguous.length)mark(ambiguous,'缺OA号时，同一供应商、付款主体和金额存在多个候选；请核对本组'+ambiguous.length+'笔明细，来源见各行来源列');
    }
  }
  function comparison(oldRecord,newRecord){
    const current=newRecord||oldRecord,notes=[],issues=[...(oldRecord?.issues||[]),...(newRecord?.issues||[]),...(oldRecord?.review||[]),...(newRecord?.review||[])];
    let change,difference=null;
    if(issues.length){change='待核对';notes.push(...new Set(issues));}
    else if(!oldRecord){change='增加';difference=newRecord.cents;notes.push('仅新版存在');}
    else if(!newRecord){change='减少';difference=-oldRecord.cents;notes.push('仅旧版存在');}
    else{
      difference=newRecord.cents-oldRecord.cents;
      if(!Number.isSafeInteger(difference))throw Error('两版金额差额超出安全范围，请核实 '+current.location);
      change=difference>0?'增加':difference<0?'减少':'不变';
      for(const field of ['category','owner','supplier','entity','oa','summary'])if(oldRecord.values[field]!==newRecord.values[field])notes.push(LABELS[field]+'：旧版“'+oldRecord.values[field]+'” → 新版“'+newRecord.values[field]+'”');
      if(!oldRecord.code||!newRecord.code)notes.unshift('按供应商、付款主体和金额唯一匹配');
    }
    let remark=current.values.remark;
    if(oldRecord&&newRecord&&oldRecord.values.remark!==newRecord.values.remark)remark='旧版：'+oldRecord.values.remark+'\n新版：'+newRecord.values.remark;
    const amount=record=>!record?str(''):record.cents==null?str(record.values.amount):number(record.cents);
    return [str(current.values.category),str(current.values.owner),str(current.values.supplier),str(current.values.entity),str(current.values.oa),str(current.values.summary),amount(oldRecord),amount(newRecord),difference==null?str(''):number(difference),str(change),str(remark),str(notes.join('\n')),str(oldRecord?.location||''),str(newRecord?.location||''),str('')];
  }
  function build(oldSource,newSource){
    const old=read(oldSource),next=read(newSource);
    match(old.records,next.records);
    const rows=next.records.map(record=>comparison(record.pair,record));
    for(const record of old.records)if(!record.pair)rows.push(comparison(record,null));
    const total=records=>records.reduce((sum,record)=>{const result=sum+(record.cents??0);if(!Number.isSafeInteger(result))throw Error('台账合计金额超出安全范围，请缩小比对范围');return result;},0);
    const oldCents=total(old.records),newCents=total(next.records),delta=newCents-oldCents;
    if(!Number.isSafeInteger(delta))throw Error('两版合计差额超出安全范围，请缩小比对范围');
    const counts={'增加':0,'减少':0,'不变':0,'待核对':0};for(const row of rows)counts[row[9].v]++;
    const warnings=[...old.warnings,...next.warnings];
    for(const row of rows)if(row[9].v==='待核对')warnings.push((row[13].v||row[12].v)+'：'+row[11].v);
    return {sheets:[{name:'比对汇总',headers:[...HEADERS],rows}],summary:{oldAmount:oldCents/100,newAmount:newCents/100,difference:delta/100,complete:[...old.records,...next.records].every(record=>record.cents!=null),counts},warnings};
  }
  function formatCell(cell){if(!cell||cell.v==null)return '';return cell.t==='n'&&cell.z?X.SSF.format(cell.z,cell.v):String(cell.v);}
  function exportFilename(date,version=1){return '付款台账表格比对-'+date+'-V'+version+'.xlsx';}
  return {HEADERS,discover,build,formatCell,exportFilename};
});
