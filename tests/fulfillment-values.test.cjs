const {test}=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),V=require('../fulfillment-validation.js');
const numeric=v=>({t:'n',v}),formula=(f,v=0,extra={})=>({t:'n',f,v,...extra});
function book(sheets){return {SheetNames:Object.keys(sheets),Sheets:sheets,Props:{Title:'原文件'},Workbook:{Names:[{Name:'原名称',Ref:'明细!A1'}]}};}

test('全部工作表先按原公式重算再转值，修正旧零并保持所有行及源book不变',()=>{
 const source=book({李奇:{A1:formula('SUM(B1:C1)'),B1:numeric(1245),C1:numeric(50),A2:formula('A1+5'),B9:{t:'s',v:'末行备注'},'!ref':'A1:C9'},辅助:{A1:formula('1+2*3',99),'!ref':'A1'}}),before=JSON.stringify(source),result=V.materialize(source);
 assert.deepEqual(result.issues,[]);assert.equal(result.formulaCount,3);assert.equal(result.correctedCount,3);assert.equal(result.book.Sheets.李奇.A1.v,1295);assert.equal(result.book.Sheets.李奇.A2.v,1300);assert.equal(result.book.Sheets.辅助.A1.v,7);
 assert.equal(JSON.stringify(source),before);assert.equal(source.Sheets.李奇.A1.f,'SUM(B1:C1)');assert.deepEqual(result.book.SheetNames,source.SheetNames);assert.deepEqual(result.book.Sheets.李奇.B9,source.Sheets.李奇.B9);assert.equal(result.book.Sheets.李奇['!ref'],'A1:C9');
 assert.deepEqual(result.updates[0],{sheet:'李奇',cell:'A1',value:1295,type:'n'});for(const sheet of Object.values(result.book.Sheets))for(const cell of Object.values(sheet))if(cell&&typeof cell==='object')assert.equal(cell.f,undefined);
});
test('正确零值、空引用、小数、负数以及缺失或错误缓存仍按公式重新计算',()=>{
 const source=book({明细:{A1:formula('B1-B1'),B1:numeric(5),A2:formula('Z9'),A3:formula('0.1+0.2',0.3),A4:formula('-B1',-5),A5:{t:'n',f:'1+1'},A6:{t:'e',f:'2+3',v:7},'!ref':'A1:B6'}}),result=V.materialize(source);
 assert.deepEqual(result.issues,[]);assert.equal(result.correctedCount,2);assert.deepEqual([1,2,3,4,5,6].map(r=>result.book.Sheets.明细['A'+r].v),[0,0,0.30000000000000004,-5,2,5]);
});
test('保留日期金额格式、样式及合并信息，清除公式和旧显示缓存',()=>{
 const style={font:{bold:true}},merges=[{s:{r:2,c:0},e:{r:2,c:2}}],source=book({明细:{A1:formula('46000',0,{z:'yyyy年m月d日',s:style,w:'旧日期',F:'A1',D:false}),A2:formula('1.25',0,{z:'#,##0.00',r:'旧富文本',h:'旧HTML'}),'!ref':'A1:C3','!merges':merges,'!rows':[{hpt:30}],'!cols':[{wch:20}]}}),result=V.materialize(source),sheet=result.book.Sheets.明细;
 assert.deepEqual(result.issues,[]);assert.equal(sheet.A1.z,'yyyy年m月d日');assert.deepEqual(sheet.A1.s,style);assert.deepEqual(sheet['!merges'],merges);assert.deepEqual(sheet['!rows'],source.Sheets.明细['!rows']);assert.deepEqual(sheet['!cols'],source.Sheets.明细['!cols']);for(const key of ['f','F','D','w'])assert.equal(sheet.A1[key],undefined);assert.equal(sheet.A2.r,undefined);assert.equal(sheet.A2.h,undefined);
});
test('同表绝对引用和跨sheet引用按依赖计算，保留文本与布尔结果类型',()=>{
 const source=book({主表:{A1:formula("'副 表'!$A$1+副表!B1"),A2:formula('SUM(副表!A1:B1)'),A3:formula('副表!C1'),A4:formula('TRUE()'),A5:formula('FALSE'),A6:formula('"文本""引号"'),A7:formula('""'),A8:formula('A4')},'副 表':{A1:formula('2+3')},副表:{A1:numeric(7),B1:numeric(8),C1:{t:'s',v:'原文本'}}}),result=V.materialize(source),sheet=result.book.Sheets.主表;
 assert.deepEqual(result.issues,[]);assert.equal(sheet.A1.v,13);assert.equal(sheet.A2.v,15);assert.deepEqual([sheet.A3.t,sheet.A3.v],['s','原文本']);assert.deepEqual([sheet.A4.t,sheet.A4.v],['b',true]);assert.deepEqual([sheet.A5.t,sheet.A5.v],['b',false]);assert.equal(sheet.A6.v,'文本"引号');assert.deepEqual([sheet.A7.t,sheet.A7.v],['s','']);assert.equal(sheet.A8.v,true);
});
test('SUM忽略范围内原文本与布尔以及公式文本结果，不进行数量猜测',()=>{
 const source=book({明细:{A1:formula('SUM(B1:B5)'),B1:numeric(3),B2:{t:'s',v:'474套左右'},B3:{t:'b',v:true},B4:formula('"5"'),B5:formula('FALSE()'),C1:formula('B2+1')}}),result=V.materialize(source);
 assert.equal(result.book.Sheets.明细.A1.v,3);assert.equal(result.issues.length,1);assert.equal(result.issues[0].cell,'C1');assert.ok(result.book.Sheets.明细.C1.f);
});
test('工作表引用支持类似单元格的名字、大小写和转义引号，清除旧富文本结果',()=>{
 const source=book({主表:{A1:formula("A1Data!A1+'O''Brien'!A1+DATA!A1"),A2:{t:'s',f:'"新结果"',v:'旧结果',r:'旧文本',h:'旧HTML',w:'旧显示'}},A1Data:{A1:numeric(2)},"O'Brien":{A1:numeric(3)},Data:{A1:numeric(4)}}),result=V.materialize(source);
 assert.deepEqual(result.issues,[]);assert.equal(result.book.Sheets.主表.A1.v,9);assert.deepEqual(result.book.Sheets.主表.A2,{t:'s',v:'新结果'});
});
test('SUBTOTAL 9/109 保留隐藏行语义且忽略嵌套小计',()=>{
 const source=book({明细:{A1:formula('SUBTOTAL(9,B1:B3)'),A2:formula('SUBTOTAL(109,B1:B3)'),A3:formula('SUM(B1:B3)'),B1:numeric(3),B2:numeric(4),B3:formula('SUBTOTAL(9,C3)'),C3:numeric(5),'!rows':[{}, {hidden:true}]}}),result=V.materialize(source);
 assert.deepEqual(result.issues,[]);assert.deepEqual([result.book.Sheets.明细.A1.v,result.book.Sheets.明细.A2.v,result.book.Sheets.明细.A3.v],[7,3,12]);
});
test('SUBTOTAL无法确认过滤状态或存在筛选条件时不采用旧缓存',()=>{
 for(const active of [true,undefined]){const sheet={A1:formula('SUBTOTAL(9,B1:B2)',99),B1:numeric(1),B2:numeric(2),'!autofilter':{ref:'A1:B3'}};if(active!==undefined)sheet['!fulfillmentFilterActive']=active;const result=V.materialize(book({明细:sheet}));assert.equal(result.issues.length,1);assert.match(result.issues[0].message,/筛选/);assert.equal(result.book.Sheets.明细.A1.f,'SUBTOTAL(9,B1:B2)');assert.equal(result.updates.length,0);}
 const result=V.materialize(book({明细:{A1:formula('SUBTOTAL(9,B1:B2)'),B1:numeric(1),B2:numeric(2),'!autofilter':{ref:'A1:B3'},'!fulfillmentFilterActive':false}}));assert.deepEqual(result.issues,[]);assert.equal(result.book.Sheets.明细.A1.v,3);
});
test('跨表SUBTOTAL按引用表的筛选状态校验',()=>{
 const source=book({主表:{A1:formula('SUBTOTAL(9,副表!B1:B2)')},副表:{B1:numeric(1),B2:numeric(2),'!fulfillmentFilterActive':true}}),result=V.materialize(source);assert.equal(result.issues.length,1);assert.match(result.issues[0].message,/筛选/);
 source.Sheets.主表['!fulfillmentFilterActive']=true;source.Sheets.副表['!fulfillmentFilterActive']=false;const good=V.materialize(source);assert.deepEqual(good.issues,[]);assert.equal(good.book.Sheets.主表.A1.v,3);
});
test('未知函数、除零、错误引用、循环、外链和巨大范围都保留公式并阻止',()=>{
 for(const f of ['IF(1,2,3)','DISPIMG("ID",1)','1/0','A1+1','SUM(A1:XFD1048576)','SUM(B2:B1)','[远程.xlsx]Sheet1!A1','缺失!A1','globalThis.FULFILLMENT_TEST=1']){const source=book({明细:{A1:formula(f,888)}}),result=V.materialize(source);assert.equal(result.issues.length,1,f);assert.equal(result.updates.length,0,f);assert.equal(result.book.Sheets.明细.A1.f,f);assert.equal(result.correctedCount,0);}
 assert.equal(globalThis.FULFILLMENT_TEST,undefined);
 const source=book({一:{A1:formula('二!A1')},二:{A1:formula('一!A1')}}),result=V.materialize(source);assert.equal(result.issues.length,2);assert.ok(result.issues.every(item=>/循环引用/.test(item.message)));
 const erroneous=V.materialize(book({明细:{A1:formula('B1+1'),B1:{t:'e',v:7}}}));assert.match(erroneous.issues[0].message,/明细!B1.*错误值/);
});
test('多单元格数组、动态数组和无公式定义的数组结果不可直接剥离',()=>{
 for(const cell of [formula('1+1',2,{F:'A1:A2'}),formula('1+1',2,{D:true}),{t:'n',v:2,F:'A1:A2'},{t:'n',v:2,f:''}]){const result=V.materialize(book({明细:{A1:cell}}));assert.equal(result.issues.length,1);assert.equal(result.updates.length,0);assert.deepEqual(result.book.Sheets.明细.A1,cell);}
});
test('导出重读后全部成功公式为真实typed values，且不改源工作簿',()=>{
 const source=book({明细:{A1:formula('1245+50'),A2:formula('"供应商"'),A3:formula('TRUE()'),A4:{t:'s',v:''},'!ref':'A1:A4'}}),result=V.materialize(source),bytes=X.write(result.book,{type:'buffer',bookType:'xlsx'}),reopened=X.read(bytes,{type:'buffer'});
 assert.deepEqual(result.issues,[]);assert.equal(reopened.Sheets.明细.A1.v,1295);assert.equal(reopened.Sheets.明细.A2.v,'供应商');assert.equal(reopened.Sheets.明细.A3.v,true);for(const cell of Object.values(reopened.Sheets.明细))if(cell&&typeof cell==='object')assert.equal(cell.f,undefined);assert.equal(source.Sheets.明细.A1.f,'1245+50');
});
