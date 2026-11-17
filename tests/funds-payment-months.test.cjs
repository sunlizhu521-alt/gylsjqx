const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),Core=require('../funds-core.js'),Validation=require('../funds-total-validation.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const UNKNOWN='付款月份待核对';
const band=(startCol,endCol,label)=>({startCol,endCol,label});
const HEADERS=['供应商','谈判明细','OA-申请付款金额','分批总金额','2026年11月30日','2026年12月28日','2026年12月支付','2027年1月支付'];
function table(rows,headers=HEADERS){
 const sheet=X.utils.aoa_to_sheet([headers,...rows]),aliases={'供应商':'供应商名称','OA-申请付款金额':'申请付款金额'},map={};
 headers.forEach((h,c)=>map[aliases[h]||h]=c);
 const records=rows.map((_,i)=>({r:i+1,source:{file:'合成.xlsx',sheet:'谈判付款',row:i+2},values:Object.fromEntries(Object.entries(map).map(([h,c])=>[h,sheet[X.utils.encode_cell({r:i+1,c})]||{t:'s',v:''}]))}));
 return {sheet,map,records,headers,headerRow:0,range:X.utils.decode_range(sheet['!ref']),date1904:false,name:'合成.xlsx',sheetName:'谈判付款'};
}
const row=(detail,a=100,b=200)=>['合成供应商',detail,Number(a||0)+Number(b||0),Number(a||0)+Number(b||0),a,b,0,0];
const bands=t=>Validation.paymentHeaderBands(t,Core.money);
function sample(){
 const input=fixture(),sheet=input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];
 sheet.J2={t:'s',v:'2026年12月10日－40\n2027年1月10日－60'};
 sheet.M1={t:'s',v:'2026年11月30日'};sheet.N1={t:'s',v:'2026年12月28日'};
 X.utils.sheet_add_aoa(sheet,[['2026年12月支付','2027年1月支付'],[40,60]],{origin:'O1'});
 return input;
}
function doubleHeader(input,labels){
 const name=input.negotiated.book.SheetNames[0],original=input.negotiated.book.Sheets[name],sheet={};
 for(const [address,value]of Object.entries(original)){
  if(!/^[A-Z]+[1-9]\d*$/.test(address)){sheet[address]=value;continue;}
  const point=X.utils.decode_cell(address);sheet[X.utils.encode_cell({r:point.r+1,c:point.c})]={...value,...(value.f?{f:value.f.replace(/(\$?[A-Z]+\$?)(\d+)/g,(_,c,r)=>c+(Number(r)+1))}:{})};
 }
 const range=X.utils.decode_range(original['!ref']);sheet['!ref']=X.utils.encode_range({s:{r:0,c:0},e:{r:range.e.r+1,c:range.e.c}});sheet['!merges']=[];
 for(let c=0;c<=range.e.c;c++)if(c!==12&&c!==13){const upper=X.utils.encode_cell({r:0,c}),lower=X.utils.encode_cell({r:1,c});sheet[upper]=sheet[lower];delete sheet[lower];sheet['!merges'].push({s:{r:0,c},e:{r:1,c}});}
 sheet.M1={t:'s',v:labels[0]};sheet.N1={t:'s',v:labels[1]};
 input.negotiated.book.Sheets[name]=sheet;return sheet;
}

test('第一层使用明细付款月：11月底申请12月付款、12月底申请次年1月付款',()=>{
 const input=sample(),before=JSON.stringify(input),result=Core.build(input),n=result.negotiated;
 assert.deepEqual(n.headerBands,[band(12,12,'2026年12月付款'),band(13,13,'2027年1月付款')]);
 assert.equal(n.headerRow,0);assert.equal(n.headerTopRow,null);assert.equal(n.rows[0][5].f,'D2-E2');
 assert.equal(n.rows[0][8].c[0].t,'保留原批注');assert.deepEqual(n.rows[0].slice(12,16).map(c=>c.v),[40,60,40,60]);assert.equal(JSON.stringify(input),before);
});

test('精确金额配对不依赖明细顺序、申请日期先后或最近月份',()=>{
 const source=table([row('2027年1月10日－200\n2026年12月10日－100')]),before=JSON.stringify(source);
 assert.deepEqual(bands(source),[band(4,4,'2026年12月付款'),band(5,5,'2027年1月付款')]);assert.equal(JSON.stringify(source),before);
});

test('同额候选全部在同一付款月时可确认，完全确定相邻同月才合并',()=>{
 const source=table([row('2026年12月10日－100\n2026年12月20日－100',100,100)]);
 assert.deepEqual(bands(source),[band(4,5,'2026年12月付款')]);
});

test('跨月同额无法对应时分别待核对，不按日期或明细顺序强行分配',()=>{
 const source=table([row('2026年12月10日－100\n2027年1月10日－100',100,100)]);
 assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
});

test('同一明细不得在两个计划格重复使用，整行合计不符时有金额的日期列待核对',()=>{
 for(const detail of ['2026年12月10日－100','2026年12月10日－90\n2027年1月10日－100']){
  assert.deepEqual(bands(table([row(detail,100,100)])),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
 }
});

test('同列多行有不同明确付款月时统一待核对，不合并未知月份列',()=>{
 const source=table([row('2026年12月10日－100\n2027年1月10日－200'),row('2027年1月12日－300\n2026年12月20日－400',300,400)]);
 assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
});

test('确认月份与歧义行并存时全列待核对，不能用部分确定结果掩盖歧义',()=>{
 const source=table([row('2026年12月10日－100\n2026年12月20日－200'),row('2026年12月10日－100\n2027年1月10日－100',100,100)]);
 assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
});

test('明细缺年份优先唯一月度支付年份，不被申请日期的同月年份覆盖',()=>{
 const headers=[...HEADERS];headers[4]='2026年1月31日';headers[6]='2027年1月支付';headers[7]='2026年12月支付';
 assert.deepEqual(bands(table([row('1月10日－100',100,0)],headers)),[band(4,4,'2027年1月付款'),band(5,5,UNKNOWN)]);
 headers[7]='2026年1月支付';
 assert.deepEqual(bands(table([row('1月10日－100',100,0)],headers)),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
});

test('显式付款年份无需从申请月猜测，分位金额精确配对',()=>{
 const source=table([row('2027年1月10日－0.2\n2026年12月10日－0.1',0.1,0.2)]);
 assert.deepEqual(bands(source),[band(4,4,'2026年12月付款'),band(5,5,'2027年1月付款')]);
});

test('日期列无非零金额仍保留且待核对，空白和零不假定月份',()=>{
 const source=table([row('2026年12月10日－100',100,null)]);
 assert.deepEqual(bands(source),[band(4,4,'2026年12月付款'),band(5,5,UNKNOWN)]);
 source.sheet.E2={t:'n',v:0};assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
});

test('明细空白、错误或公式缺缓存时不能产生确定月份',()=>{
 for(const detail of [undefined,{t:'s',v:''},{t:'e',v:7},{t:'n',f:'A1'},{t:'z',f:'A1',v:0},{t:'s',v:'2026年12月10日－100\n无法解析'}]){
  const source=table([row('2026年12月10日－100',100,0)]);if(detail)source.sheet.B2=detail;else delete source.sheet.B2;
  assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
 }
});

test('计划金额错误、无效或公式缺缓存时不能用其他金额掩盖该行问题',()=>{
 for(const value of [{t:'e',v:7},{t:'s',v:'金额待定'},{t:'n',f:'100+100'},{t:'z',f:'200',v:0}]){
  const source=table([row('2026年12月10日－100\n2027年1月10日－200')]);source.sheet.F2=value;
  assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
 }
});

test('谈判明细或付款计划合并单元格归属不明时不分配付款月',()=>{
 for(const ref of ['B2:C2','E2:F2']){
  const source=table([row('2026年12月10日－100\n2027年1月10日－200')]);source.sheet['!merges']=[X.utils.decode_range(ref)];
  assert.deepEqual(bands(source),[band(4,4,UNKNOWN),band(5,5,UNKNOWN)]);
 }
});

test('申请日期表头跨列合并时金额归属不清，不能给出确定付款月份',()=>{
 const source=table([row('2026年12月10日－100',100,0)]);delete source.sheet.F1;source.sheet['!merges']=[X.utils.decode_range('E1:F1')];
 assert.deepEqual(bands(source),[band(4,4,UNKNOWN)]);
});

test('同一月份相邻日期中夹有保留列时不跨列合并',()=>{
 const headers=[...HEADERS.slice(0,5),'人工备注',...HEADERS.slice(5)],values=row('2026年12月10日－100\n2026年12月20日－200');values.splice(5,0,'原内容');
 const source=table([values],headers);assert.deepEqual(bands(source),[band(4,4,'2026年12月付款'),band(6,6,'2026年12月付款')]);assert.equal(source.sheet.F2.v,'原内容');
});

test('新付款月份、待核对及旧年月标签重上传均识别下行日期表头',()=>{
 for(const labels of [['2026年12月付款','2027年1月付款'],['2026年12月付款',UNKNOWN],[UNKNOWN,UNKNOWN],['2026年11月','2026年12月']]){
  const input=sample(),sheet=doubleHeader(input,labels),before=JSON.stringify(input),first=Core.build(input),again=Core.build(input);
  for(const result of [first,again]){
   assert.equal(result.negotiated.headerTopRow,0);assert.equal(result.negotiated.headerRow,1);assert.equal(result.stats.negotiatedRows,1);assert.equal(result.negotiated.rows[0][5].f,'D3-E3');
   assert.deepEqual(result.negotiated.headerBands,[band(12,12,'2026年12月付款'),band(13,13,'2027年1月付款')]);
  }
  assert.equal(sheet.A2,undefined);assert.equal(sheet.M2.v,'2026年11月30日');assert.equal(sheet.N2.v,'2026年12月28日');assert.equal(sheet.I3.c[0].t,'保留原批注');
  assert.equal(X.utils.decode_range(sheet['!ref']).e.r,2);assert.equal(JSON.stringify(input),before,'重读不重复新增表头行或改动源内容');
 }
});
