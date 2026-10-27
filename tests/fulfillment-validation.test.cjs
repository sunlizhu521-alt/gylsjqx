const {test}=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js'),V=require('../fulfillment-validation.js');
function fixture(){
 const book=X.utils.book_new();
 for(const person of C.SHEETS){const row=C.COLUMNS.map(key=>V.FIELDS.includes(key)?V.FIELDS.indexOf(key)+1:({'事业部':'海外事业一部','物料编码':'001','采购下单人':person==='操思敏'?'陈燕':person})[key]??'');X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,row]),person);}
 return book;
}
const validate=book=>V.validate(book,C.merge(book),C);
function formula(book,f,value=0,address='K2'){book.Sheets.李奇[address]={t:'n',f,v:value};return book;}
function issueOf(result,kind){return result.issues.find(item=>item.kind===kind);}

test('六字段逐表及合计，按来源工作表归属而非下单人，保留源文件',()=>{
 const book=fixture();book.Sheets.孙立柱=X.utils.aoa_to_sheet([C.COLUMNS,C.COLUMNS.map(key=>V.FIELDS.includes(key)?999:'')]);
 const before=JSON.stringify(book),result=validate(book);
 assert.equal(result.passed,true);assert.equal(result.checks.length,30);assert.equal(result.formulaCount,0);
 for(const [index,field] of V.FIELDS.entries()){assert.equal(result.checks.find(c=>c.sheet==='操思敏'&&c.field===field).sourceTotal,index+1);assert.equal(result.checks.find(c=>c.sheet==='合计'&&c.field===field).sourceTotal,(index+1)*4);}
 assert.equal(JSON.stringify(book),before);assert.deepEqual(result.issues,[]);
});
test('支持姓名后日期后缀，缺失或多个同人来源不能通过',()=>{
 let book=fixture();book.Sheets.邢芳芳1017=book.Sheets.邢芳芳;delete book.Sheets.邢芳芳;assert.equal(validate(book).passed,true);
 book=fixture();const merged=C.merge(book);delete book.Sheets.李奇;assert.equal(V.validate(book,merged,C).passed,false);
 book=fixture();book.Sheets.李奇1017=book.Sheets.李奇;assert.equal(V.validate(book,merged,C).passed,false);
});
test('缺行和重复行分别发现，不能只用合计相抵判断',()=>{
 const book=fixture(),result=C.merge(book);result.rows.splice(1,1);result.sources.splice(1,1);const missing=V.validate(book,result,C);assert.equal(missing.passed,false);assert.match(issueOf(missing,'mapping').message,/未进入/);
 const full=C.merge(book);full.rows.push(full.rows[0]);full.sources.push(full.sources[0]);const duplicate=V.validate(book,full,C);assert.equal(duplicate.passed,false);assert.match(issueOf(duplicate,'mapping').message,/重复汇总/);
});
test('逐行数量偏差互相抵消仍然阻止',()=>{
 const book=fixture(),sheet=book.Sheets.徐英;sheet.K3={t:'n',v:3};sheet.H3={t:'s',v:'002'};sheet['!ref']='A1:X3';const result=C.merge(book);
 result.rows[0][10].v=2;result.rows[1][10].v=2;
 const checked=V.validate(book,result,C),total=checked.checks.find(x=>x.sheet==='徐英'&&x.field==='未交付数量');assert.equal(total.difference,0);assert.equal(checked.passed,false);assert.equal(checked.issues.filter(x=>x.kind==='value').length,2);
});
test('来源映射数量错误、额外来源和来源计数不一致均阻止',()=>{
 for(const alter of [r=>r.sources.pop(),r=>r.sources[0]={sheet:'孙立柱',row:2},r=>r.sources[0].row=100,r=>r.counts[0].count++]){const book=fixture(),result=C.merge(book);alter(result);assert.equal(V.validate(book,result,C).passed,false);}
});
test('六字段任一字段被改写都会发现',()=>{
 for(const field of V.FIELDS){const book=fixture(),result=C.merge(book);result.rows[0][C.COLUMNS.indexOf(field)].v++;const checked=V.validate(book,result,C);assert.equal(checked.passed,false);assert.ok(checked.issues.some(x=>x.kind==='value'&&x.field===field));}
});
test('源字段缺失重复或汇总列缺失会阻止',()=>{
 const book=fixture(),result=C.merge(book);delete book.Sheets.徐英.K1;assert.equal(V.validate(book,result,C).passed,false);
 book.Sheets.徐英.K1={t:'s',v:'已发货数量'};assert.equal(V.validate(book,result,C).passed,false);
 const clean=fixture();const badCore={...C,summaryRows:rows=>C.summaryRows(rows).map(row=>{delete row[11];return row;})};assert.equal(V.validate(clean,C.merge(clean),badCore).passed,false);
});
test('空白与数值零均可校验，但文本数字、其他文字和错误不能冒充数量',()=>{
 for(const cell of [undefined,{t:'s',v:''},{t:'n',v:0}]){const book=fixture();book.Sheets.李奇.K2=cell;assert.equal(validate(book).passed,true);}
 for(const cell of [{t:'s',v:'0'},{t:'s',v:'物料通用   474套左右'},{t:'e',v:7},{t:'n',v:NaN},{t:'n',v:Infinity},{t:'n',v:1e20},{t:'b',v:false}]){const book=fixture();book.Sheets.李奇.K2=cell;const checked=validate(book);assert.equal(checked.passed,false);assert.equal(checked.checks.find(x=>x.sheet==='李奇'&&x.field==='未交付数量').sourceTotal,null);}
});
test('汇总中的数值文本类型损坏即使合计可相同也阻止',()=>{
 const book=fixture(),result=C.merge(book);result.rows[0][10]={t:'s',v:'1'};assert.equal(V.validate(book,result,C).passed,false);
});
test('公式常量、括号、运算优先级、绝对引用和递归引用',()=>{
 const cases=[['1+2*3',7],['(1+2)*3',9],['-$L$2+3',1],['SUM(L2:M2,N2:P2)',20],['SUBTOTAL(9,L2:P2)',20],['SUBTOTAL(109,L2:P2)',20],['1E2/2',50],['0.1+0.2',0.3]];
 for(const [expression,value] of cases){const checked=validate(formula(fixture(),expression,value));assert.equal(checked.passed,true,expression);assert.equal(checked.formulaCount,1);}
 const book=formula(fixture(),'L2+1',3);formula(book,'M2-1',2,'L2');const checked=validate(book);assert.equal(checked.passed,true);assert.equal(checked.formulaCount,2);
});
test('保存的公式旧结果为零时也必须重算核验，并显示确切差额',()=>{
 const checked=validate(formula(fixture(),'SUBTOTAL(9,L2:P2)',0));assert.equal(checked.passed,false);assert.equal(issueOf(checked,'formula-cache').cell,'K2');
 const total=checked.checks.find(c=>c.sheet==='李奇'&&c.field==='未交付数量');assert.equal(total.sourceTotal,20);assert.equal(total.summaryTotal,0);assert.equal(total.difference,-20);
});
test('公式正确结果为零和空白引用不误报',()=>{
 const book=formula(fixture(),'L2-L2+Z2',0);assert.equal(validate(book).passed,true);
});
test('公式结果缺失、错误或文本均阻止，源公式不改写',()=>{
 for(const saved of [{t:'n',f:'1+1'},{t:'e',f:'1+1',v:7},{t:'s',f:'1+1',v:'2'}]){const book=fixture();book.Sheets.李奇.K2=saved;const before=JSON.stringify(book),checked=validate(book);assert.equal(checked.passed,false);assert.ok(issueOf(checked,'formula-result'));assert.equal(JSON.stringify(book),before);}
});
test('未知函数、跨表引用、循环、除零和无效公式不能通过',()=>{
 for(const f of ['IF(1,2,3)','SUMIF(A1:A2,1)','徐英!K2','K2+1','1/0','SUM(L2:P2','globalThis.TEST=1']){const checked=validate(formula(fixture(),f,0));assert.equal(checked.passed,false,f);assert.ok(issueOf(checked,'source'),f);}
 const book=formula(fixture(),'L2',1);formula(book,'K2',1,'L2');assert.equal(validate(book).passed,false);assert.equal(globalThis.TEST,undefined);
});
test('公式引用的错误、旧结果和不支持公式也会阻止',()=>{
 const book=formula(fixture(),'L2+1',3);formula(book,'1+1',0,'L2');const checked=validate(book);assert.equal(checked.passed,false);assert.equal(issueOf(checked,'formula-cache').cell,'L2');
 for(const cell of [{t:'e',v:7},{t:'n',f:'IF(1,2,3)',v:2}]){const source=formula(fixture(),'L2+1',3);source.Sheets.李奇.L2=cell;assert.equal(validate(source).passed,false);}
});
test('SUBTOTAL忽略嵌套SUBTOTAL，SUM保留嵌套结果',()=>{
 const book=formula(fixture(),'SUBTOTAL(9,L2:P2)',18);formula(book,'SUBTOTAL(9,Q2)',2,'L2');book.Sheets.李奇.Q2={t:'n',v:2};assert.equal(validate(book).passed,true);
 formula(book,'SUM(L2:P2)',20);assert.equal(validate(book).passed,true);
});
test('SUBTOTAL 9包含手工隐藏行，109忽略；筛选条件和未知过滤状态明确待核验',()=>{
 for(const [fn,value] of [[9,20],[109,0]]){const book=formula(fixture(),`SUBTOTAL(${fn},L2:P2)`,value);book.Sheets.李奇['!rows']=[{}, {hidden:true}];assert.equal(validate(book).passed,true);}
 for(const active of [true,undefined]){const book=formula(fixture(),'SUBTOTAL(9,L2:P2)',20);book.Sheets.李奇['!autofilter']={ref:'A1:X2'};if(active!==undefined)book.Sheets.李奇['!fulfillmentFilterActive']=active;const checked=validate(book);assert.equal(checked.passed,false);assert.match(issueOf(checked,'source').message,/筛选/);}
 const book=formula(fixture(),'SUBTOTAL(9,L2:P2)',20);book.Sheets.李奇['!autofilter']={ref:'A1:X2'};book.Sheets.李奇['!fulfillmentFilterActive']=false;assert.equal(validate(book).passed,true);
});
test('筛选不改变普通SUM核验，文本参考按SUM语义忽略',()=>{
 const book=formula(fixture(),'SUM(Z1:Z2)',2);book.Sheets.李奇.Z1={t:'s',v:'数字文本被SUM忽略'};book.Sheets.李奇.Z2={t:'n',v:2};book.Sheets.李奇['!fulfillmentFilterActive']=true;assert.equal(validate(book).passed,true);
});
test('小数求和不产生浮点误报，真正的小差异仍阻止',()=>{
 const book=formula(fixture(),'0.1+0.2',0.30000000000000004);assert.equal(validate(book).passed,true);
 formula(book,'0.1+0.2',0.300001);assert.equal(validate(book).passed,false);
});
test('合计超出安全范围、过大的公式范围和逆向范围均不能通过',()=>{
 const book=fixture();for(const name of C.SHEETS)book.Sheets[name].K2={t:'n',v:Number.MAX_SAFE_INTEGER/2};const checked=validate(book);assert.equal(checked.passed,false);assert.ok(issueOf(checked,'total'));
 for(const f of ['SUM(A1:XFD1048576)','SUM(L3:L2)','SUM(XFE1:XFE2)'])assert.equal(validate(formula(fixture(),f,0)).passed,false,f);
});
test('错误引用位置应指向实际依赖单元格',()=>{
 const book=formula(fixture(),'Z2+1',2);book.Sheets.李奇.Z2={t:'e',v:7};const checked=validate(book);assert.equal(checked.passed,false);assert.equal(issueOf(checked,'source').cell,'Z2');
});
test('原值零和空白不可互换，即使汇总合计都为零',()=>{
 for(const [source,target] of [[{t:'n',v:0},{t:'s',v:''}],[{t:'s',v:''},{t:'n',v:0}]]){const book=fixture();book.Sheets.李奇.K2=source;const result=C.merge(book);result.rows[result.sources.findIndex(s=>s.sheet==='李奇')][10]=target;const checked=V.validate(book,result,C);assert.equal(checked.passed,false);assert.match(issueOf(checked,'value').message,/空白与数字状态/);}
});
test('大整数相差1也必须报差异，依赖公式问题归属实际数量字段',()=>{
 const book=formula(fixture(),'1000000000000000',1000000000000001);assert.equal(validate(book).passed,false);
 const source=formula(fixture(),'L2+1',3);formula(source,'1+1',0,'L2');const checked=validate(source);assert.equal(issueOf(checked,'formula-cache').field,'已下单未备料未生产');assert.equal(checked.checks.find(c=>c.sheet==='李奇'&&c.field==='已下单未备料未生产').status,'不一致');
});
