const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,LH}=require('./funds-progress-fixtures.cjs');
function merged(flag='是',owner='李奇',reverse=false){
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],grid=X.utils.sheet_to_json(s,{header:1}),row=grid[1],headers=reverse?LH.slice().reverse():LH;
 const records=[0,1,2].map(n=>{const next=row.slice();next[1]='合成合并供应商'+n;next[3]='MDST-DGFK2026111500'+n;next[15]=n===0?flag:'';return reverse?next.reverse():next;});
 for(const person of C.PEOPLE)i.ledger.book.Sheets['对账明细-'+person]=X.utils.aoa_to_sheet([LH]);
 const sheet=X.utils.aoa_to_sheet([headers,...records]),c=headers.indexOf('是否谈判付款');sheet['!merges']=[{s:{r:1,c},e:{r:3,c}}];i.ledger.book.Sheets['对账明细-'+owner]=sheet;
 return {i,sheet,c};
}
test('是否谈判付款纵向合并先展开再分类，五人不同列位置均不把是的从属格当空白正常付款',()=>{
 for(const [n,person]of C.PEOPLE.entries())for(const flag of ['是','否','']){
  const {i}=merged(flag,person,n%2===0),before=JSON.stringify(i),r=C.build(i);
  assert.equal(r.stats.ledgerRows,3);assert.equal(r.normalRows.length,flag==='是'?0:3);assert.equal(r.stats.negotiatedLedgerRows,flag==='是'?3:0);
  assert.equal(r.negotiated.rows.length,1,'不新增谈判明细');assert.equal(JSON.stringify(i),before,'原工作簿和值不修改');
 }
});
test('不向合并区外或其他字段填值，支持重复整理和锚点有效公式结果',()=>{
 const {i,sheet,c}=merged('是');sheet[X.utils.encode_cell({r:1,c})].f='IF(1=1,"是","否")';
 const row=X.utils.sheet_to_json(sheet,{header:1})[3].slice();row[1]='非合并区正常款';row[c]='';X.utils.sheet_add_aoa(sheet,[row],{origin:-1});
 sheet['!merges'].push({s:{r:1,c:6},e:{r:3,c:6}});delete sheet.G3;delete sheet.G4;
 const before=JSON.stringify(i),first=C.build(i),second=C.build(i);assert.equal(first.normalRows.length,1);assert.equal(first.normalRows[0][2].v,'非合并区正常款');assert.deepEqual(second,first);assert.equal(JSON.stringify(i),before);
});
test('合并首格异常不按空白分类，跨列或从属隐藏冲突提示来源且不覆盖',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'n',f:'A2'}, {t:'s',v:'',f:'A2'}]){
  const {i,sheet,c}=merged();sheet[X.utils.encode_cell({r:1,c})]=bad;assert.throws(()=>C.build(i),/ledger.xlsx \/ 对账明细-李奇!P2:P4.*公式结果不可用/);
 }
 for(const bad of [{t:'s',v:'否'},{t:'s',v:'是',f:'A2'},{t:'e',v:23}]){
  const {i,sheet,c}=merged();sheet[X.utils.encode_cell({r:2,c})]=bad;const before=JSON.stringify(i);assert.throws(()=>C.build(i),/P3 含不同内容或公式/);assert.equal(JSON.stringify(i),before);
 }
 const across=merged();across.sheet['!merges'][0].e.c++;assert.throws(()=>C.build(across.i),/合并跨越其他列/);
 const same=merged();same.sheet.P3={t:'s',v:' 是 '};assert.equal(C.build(same.i).normalRows.length,0);
 const wrong=merged('待定');const r=C.build(wrong.i);assert.equal(r.normalRows.length,0);assert.equal(r.issues.filter(q=>q.kind==='classification').length,3,'异常锚点逐行核对');
});
