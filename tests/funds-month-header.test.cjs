const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const source=input=>input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];
function sample(){
 const input=fixture(),sheet=source(input);
 X.utils.sheet_add_aoa(sheet,[['2026年12月14日','2027年1月18日','2026年11月支付','2026年12月支付','2027年1月支付'],['','',100,0,0]],{origin:'O1'});
 return input;
}
function shift(sheet,offset){
 const next={};for(const [address,cell]of Object.entries(sheet)){
  if(!/^[A-Z]+[1-9]\d*$/.test(address)){next[address]=cell;continue;}
  const p=X.utils.decode_cell(address);next[X.utils.encode_cell({r:p.r+offset,c:p.c})]={...cell,...(cell.f?{f:cell.f.replace(/(\$?[A-Z]+\$?)(\d+)/g,(_,c,r)=>c+(Number(r)+offset))}:{})};
 }
 const range=X.utils.decode_range(sheet['!ref']);next['!ref']=X.utils.encode_range({s:{r:0,c:range.s.c},e:{r:range.e.r+offset,c:range.e.c}});
 next['!merges']=(sheet['!merges']||[]).map(m=>({s:{r:m.s.r+offset,c:m.s.c},e:{r:m.e.r+offset,c:m.e.c}}));
 return next;
}
function doubleHeader(input,offset=0){
 const original=source(input),bands=C.build(input).negotiated.headerBands,sheet=shift(original,offset+1),last=X.utils.decode_range(original['!ref']).e.c;
 const dateCols=new Set(bands.flatMap(b=>Array.from({length:b.endCol-b.startCol+1},(_,i)=>b.startCol+i)));
 for(let c=0;c<=last;c++)if(!dateCols.has(c)){
  const lower=X.utils.encode_cell({r:offset+1,c}),upper=X.utils.encode_cell({r:offset,c});
  sheet[upper]={...sheet[lower]};delete sheet[lower];sheet['!merges'].push({s:{r:offset,c},e:{r:offset+1,c}});
 }
 for(const band of bands){sheet[X.utils.encode_cell({r:offset,c:band.startCol})]={t:'s',v:band.label};if(band.endCol>band.startCol)sheet['!merges'].push({s:{r:offset,c:band.startCol},e:{r:offset,c:band.endCol}});}
 if(offset)sheet.A1={t:'s',v:'人工说明'};
 input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]]=sheet;
 return sheet;
}

test('单层源按谈判明细付款月份分组，空申请列待核对，不移动或补金额',()=>{
 const input=sample(),before=JSON.stringify(input),result=C.build(input),n=result.negotiated;
 assert.equal(n.headerRow,0);assert.equal(n.headerTopRow,null);
 assert.deepEqual(n.headerBands,[{startCol:12,endCol:13,label:'2026年11月付款'},{startCol:14,endCol:14,label:'付款月份待核对'},{startCol:15,endCol:15,label:'付款月份待核对'}]);
 assert.deepEqual(n.rows[0].slice(12,19).map(c=>c.v),[40,60,'','',100,0,0]);
 assert.equal(n.rows[0][5].f,'D2-E2');assert.equal(n.rows[0][8].c[0].t,'保留原批注');assert.equal(JSON.stringify(input),before);
 assert.deepEqual(n.columnOrder,n.headers.map((_,i)=>i));
});

test('月份相同但日期列被其他字段隔开时不跨字段合并，非法日期不生成月份带',()=>{
 const input=sample(),sheet=source(input);X.utils.sheet_add_aoa(sheet,[['人工保留列','2026年11月30日','2026年2月30日'],['说明',0,0]],{origin:'T1'});
 const result=C.build(input);
 assert.deepEqual(result.negotiated.headerBands.at(-1),{startCol:20,endCol:20,label:'付款月份待核对'});
 assert(!result.negotiated.headerBands.some(b=>b.startCol<=19&&b.endCol>=19));assert(!result.negotiated.headerBands.some(b=>b.endCol===21));
 assert.equal(source(input).T2.v,'说明');
});

test('没有月度支付列或没有完整日期计划列时不启用新表头',()=>{
 const base=fixture();assert.deepEqual(C.build(base).negotiated.headerBands,[]);
 const input=sample(),sheet=source(input);for(const c of ['M','N','O','P'])sheet[c+'1']={t:'s',v:'保留列'+c};
 const result=C.build(input);assert.deepEqual(result.negotiated.headerBands,[]);assert.equal(result.negotiated.headerTopRow,null);
});

test('双层表头重上传按下方日期行识别，原金额、公式、批注与实际源行保留',()=>{
 const input=sample(),sheet=doubleHeader(input),before=JSON.stringify(input),result=C.build(input),n=result.negotiated;
 assert.equal(n.headerTopRow,0);assert.equal(n.headerRow,1);assert.equal(n.rows.length,1);assert.equal(result.stats.negotiatedRows,1);
 assert.equal(n.headers[0],'供应商名称');assert.equal(n.headers[16],'2026年11月支付');assert.equal(n.rows[0][3].v,100);assert.equal(n.rows[0][5].f,'D3-E3');
 assert.equal(n.rows[0][8].c[0].t,'保留原批注');assert.equal(result.totalValidation.rows[0][2].v,3);
 assert.deepEqual(n.batchPlans[0].cells.map(c=>c.address),['M3','N3']);assert.equal(result.normalRows.length,2);
 assert.equal(sheet.A2,undefined);assert.equal(sheet.Q2,undefined);assert.equal(JSON.stringify(input),before);
 assert.deepEqual(n.headerBands,[{startCol:12,endCol:13,label:'2026年11月付款'},{startCol:14,endCol:14,label:'付款月份待核对'},{startCol:15,endCol:15,label:'付款月份待核对'}]);
 assert.equal(C.build(input).negotiated.headerTopRow,0,'重复读取不重复添加行');
});

test('带前置说明的双层表头保留偏移和真实行号，不将上层或下层算作数据',()=>{
 const input=sample();doubleHeader(input,2);const result=C.build(input);
 assert.equal(result.negotiated.headerTopRow,2);assert.equal(result.negotiated.headerRow,3);assert.equal(result.stats.negotiatedRows,1);
 assert.equal(result.totalValidation.rows[0][2].v,5);assert.equal(result.negotiated.rows[0][5].f,'D5-E5');assert.equal(source(input).A1.v,'人工说明');
});

test('普通单行表头及首条业务数据不被识别为月份带，其他输入保持单行识别',()=>{
 const input=sample(),sheet=source(input);sheet.A2.v='2026年11月';const before=JSON.stringify(input),result=C.build(input);
 assert.equal(result.negotiated.headerTopRow,null);assert.equal(result.negotiated.headerRow,0);assert.equal(result.stats.negotiatedRows,1);assert.equal(result.stats.applications,4);
 assert.equal(result.stats.ledgerRows,3);assert.equal(JSON.stringify(input),before);
});

test('纵向合并表头下层存在冲突值时停止，不覆盖原内容',()=>{
 const input=sample(),sheet=doubleHeader(input);sheet.A2={t:'s',v:'冲突内容'};const before=JSON.stringify(input);
 assert.throws(()=>C.build(input),/双层表头.*A2.*不同内容/);assert.equal(JSON.stringify(input),before);
});

test('付款年份取付款月份依据，申请日期的批次年份决定不改变付款月份或原日期',()=>{
 const input=sample(),sheet=source(input),name=input.negotiated.book.SheetNames[0];sheet.Q1.v='2027年11月支付';const before=JSON.stringify(input),initial=C.build(input);
 assert.equal(initial.negotiated.headerBands[0].label,'2027年11月付款');assert.equal(initial.issues.filter(i=>i.kind==='batch-year'&&!i.resolved).length,2);
 const decisions=Object.fromEntries(['M1','N1'].map(a=>['batch-year:'+name+'!'+a,['year:2027']])),changed=C.build(input,decisions);
 assert.equal(changed.negotiated.headerBands[0].label,'2027年11月付款');assert.equal(changed.negotiated.batchYearAdjustments.length,2);assert.equal(JSON.stringify(input),before);
 const kept=C.build(input,Object.fromEntries(['M1','N1'].map(a=>['batch-year:'+name+'!'+a,['keep']])));
 assert.equal(kept.negotiated.headerBands[0].label,'2027年11月付款');assert.equal(kept.negotiated.batchYearAdjustments.length,0);
});

test('双层带月份文字允许已确认年份与原日期不同，并在重上传再次核对',()=>{
 const input=sample(),sheet=doubleHeader(input);sheet.M1.v='2027年11月';sheet.Q1.v='2027年11月支付';
 const result=C.build(input);assert.equal(result.negotiated.headerTopRow,0);assert.equal(result.negotiated.headerRow,1);
 assert.equal(result.negotiated.headerBands[0].label,'2027年11月付款');assert.equal(result.issues.filter(i=>i.kind==='batch-year'&&!i.resolved).length,2);
});

test('已双层源移除月支付列仍重新核对付款月份，不保留过期上层标签',()=>{
 const input=sample(),sheet=doubleHeader(input);for(const c of ['Q','R','S'])sheet[c+'1'].v='人工列'+c;
 const result=C.build(input);assert.equal(result.negotiated.headerTopRow,0);assert.equal(result.negotiated.headerRow,1);assert.equal(result.stats.negotiatedRows,1);assert.equal(result.negotiated.headerBands[0].label,'2026年11月付款');assert.equal(result.negotiated.headerBands.at(-1).label,'付款月份待核对');
});
