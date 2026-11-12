const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const X=require('../vendor/xlsx.full.min.js');
const Zip=require('../vendor/jszip.min.js');
const Export=require('../funds-budget-summary-export.js');
const NAME='付款预算汇总';
const headers=['付款月份','分类','状态','申请总金额','已付款金额','剩余付款金额','备注'];
const group=(changes={})=>({monthKey:'2026-11',category:'正常付款',state:'未支付',totalCents:12345,paidCents:2345,remainingCents:10000,...changes});
function result(groups=[group()],changes={}) {
  const totals={totalCents:0,paidCents:0,remainingCents:0};
  for(const g of groups)for(const key of Object.keys(totals))totals[key]+=g[key];
  return {groups,totals,recordCount:groups.length,warnings:[],errors:[],canExport:true,date1904:false,...changes};
}
async function fixture({summaryIndex=-1,date1904=false,formula,calcChain=false}={}) {
  const book=X.utils.book_new(),source=X.utils.aoa_to_sheet([['保留表头','金额','人工说明'],['原始值',123.45,'=字面文本'],['公式',100]]);
  source.D2={t:'n',v:date1904?44865:46327,z:'yyyy"年"m"月"',f:'DATE(2026,11,1)'};source['!ref']='A1:D4';
  source.B3={t:'n',v:100,f:formula||'B2-23.45'};source.A2.c=[{a:'人工作者',t:'原批注保持'}];
  source['!cols']=[{wch:27},{wch:19,hidden:true},{wch:35}];source['!rows']=[{hpt:31},{hpt:43,hidden:true}];source['!merges']=[{s:{r:3,c:0},e:{r:3,c:2}}];source['!autofilter']={ref:'A1:C3'};
  const names=['明细甲','明细乙'];if(summaryIndex>=0)names.splice(summaryIndex,0,NAME);
  for(const name of names){const sheet=name===NAME?X.utils.aoa_to_sheet([['旧汇总'],['旧值']]):name==='明细甲'?source:X.utils.aoa_to_sheet([['另一页'],[56]]);if(name===NAME)sheet.A2.c=[{a:'旧汇总作者',t:'应移除的旧批注'}];X.utils.book_append_sheet(book,sheet,name);}
  book.Workbook={WBProps:{date1904},Sheets:names.map(name=>({Hidden:name==='明细乙'?2:0})),Names:[{Name:'明细局部',Sheet:names.indexOf('明细甲'),Ref:'$B$2'},{Name:'跨表名称',Ref:"'明细甲'!$B$3"}]};
  if(summaryIndex>=0)book.Workbook.Names.push({Name:'_xlnm.Print_Area',Sheet:summaryIndex,Ref:"'付款预算汇总'!$A$1:$X$99"},{Name:'自定义汇总名称',Sheet:summaryIndex,Ref:'$A$1'});
  const zip=await Zip.loadAsync(X.write(book,{type:'buffer',bookType:'xlsx'}));
  zip.file('custom/preserve.bin',new Uint8Array([0,255,11,98]));
  let workbook=await zip.file('xl/workbook.xml').async('string');
  if(!workbook.includes('<bookViews>'))workbook=workbook.replace('<sheets>','<bookViews><workbookView activeTab="1" firstSheet="1"/></bookViews><sheets>');
  zip.file('xl/workbook.xml',workbook);
  if(calcChain){zip.file('xl/calcChain.xml','<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B3" i="1"/></calcChain>');zip.file('xl/_rels/workbook.xml.rels',(await zip.file('xl/_rels/workbook.xml.rels').async('string')).replace('</Relationships>','<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'));zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml').async('string')).replace('</Types>','<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));}
  return zip.generateAsync({type:'uint8array'});
}
async function unpack(bytes){return {book:X.read(bytes,{type:'array',cellNF:true,cellStyles:true}),zip:await Zip.loadAsync(bytes)};}
function summaryPath(zip){return zip.file('xl/workbook.xml').async('string').then(async workbook=>{const id=workbook.match(/<(?:[\w.-]+:)?sheet\b[^>]*name="付款预算汇总"[^>]*r:id="([^"]+)"/)[1],rels=await zip.file('xl/_rels/workbook.xml.rels').async('string'),rel=(rels.match(/<Relationship\b[^>]*\/>/g)||[]).find(r=>r.includes(`Id="${id}"`));return 'xl/'+rel.match(/Target="([^"]+)"/)[1];});}
const section=(xml,name)=>xml.match(new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)</${name}>`))[1];

test('新增第一张汇总，固定7列、年月文本、金额和公式缓存可回读',async()=>{
  const source=await fixture(),{book,zip}=await unpack(await Export(source,result()));
  assert.deepEqual(book.SheetNames,[NAME,'明细甲','明细乙']);
  const sheet=book.Sheets[NAME];assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],headers);
  assert.equal(sheet.A2.v,'2026年11月');assert.equal(sheet.A2.t,'s');assert.equal(sheet.A2.z,'@');assert.equal(sheet.A2.w,'2026年11月');assert.equal(sheet.A2.f,undefined);
  assert.equal(sheet.D2.v,123.45);assert.equal(sheet.E2.v,23.45);assert.equal(sheet.F2.v,100);assert.equal(sheet.F2.f,'D2-E2');
  for(const [c,value] of [['D',123.45],['E',23.45],['F',100]]){assert.equal(sheet[c+'3'].v,value);assert.equal(sheet[c+'3'].f,`SUM(${c}2:${c}2)`);}
  assert.equal(sheet.G2.v,'');assert.equal(sheet.G3.v,'');assert.equal(sheet['!autofilter'].ref,'A1:G2');
  const xml=await zip.file(await summaryPath(zip)).async('string');assert.match(xml,/<pane ySplit="1" topLeftCell="A2"[^>]*state="frozen"/);assert.doesNotMatch(xml,/<tableParts|<filterColumn|hidden="/);
  const styles=await zip.file('xl/styles.xml').async('string'),xfs=section(styles,'cellXfs').match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g),originalStyles=await (await Zip.loadAsync(source)).file('xl/styles.xml').async('string');
  for(const name of ['fonts','fills','borders','cellXfs'])assert(section(styles,name).startsWith(section(originalStyles,name)),name+' preserves every original style');
  const borders=section(styles,'borders').match(/<border\b[^>]*(?:\/>|>[\s\S]*?<\/border>)/g);
  for(const c of xml.match(/<c\b[^>]+>/g)){
    const style=xfs[Number(c.match(/s="(\d+)"/)[1])];assert.match(style,/horizontal="center" vertical="center" wrapText="0"/);assert.match(style,/applyBorder="1"/);
    const address=c.match(/\br="([A-Z]+\d+)"/)[1],row=Number(address.match(/\d+$/)[0]),border=borders[Number(style.match(/\bborderId="(\d+)"/)[1])];
    const value=sheet[address]?.v,blank=value==null||(typeof value==='string'&&!value.trim());
    if(row>1&&blank)assert.doesNotMatch(border,/style=/,address);
    else for(const side of ['left','right','top','bottom'])assert.match(border,new RegExp(`<${side} style="${row===1?'thin':'hair'}"><color rgb="FF404040"/></${side}>`),address);
  }
  assert.equal(sheet.A1.s.fgColor.rgb,'315B9D');assert.equal(sheet.A2.s.fgColor.rgb,'FFFFFF');assert(sheet['!cols'][3].width>=14);
});

test('生成表空白正文无边框，零金额与false仍有边框，公式及原表格式不变',async()=>{
  const source=await fixture(),{book,zip}=await unpack(await Export(source,result([group({category:' \n ',state:false,totalCents:0,paidCents:0,remainingCents:0})]))),sheet=book.Sheets[NAME];
  const styles=await zip.file('xl/styles.xml').async('string'),xfs=section(styles,'cellXfs').match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g),borders=section(styles,'borders').match(/<border\b[^>]*(?:\/>|>[\s\S]*?<\/border>)/g),xml=await zip.file(await summaryPath(zip)).async('string');
  for(const [address,hasBorder] of [['B2',false],['C2',true],['D2',true],['F2',true],['G2',false],['B3',false],['D3',true]]){
    const tag=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0],style=xfs[Number(tag.match(/\bs="(\d+)"/)[1])],border=borders[Number(style.match(/\bborderId="(\d+)"/)[1])];
    if(hasBorder)assert.match(border,/style="hair"/,address);else assert.doesNotMatch(border,/style=/,address);
    assert.equal(sheet[address].s.fgColor.rgb,'FFFFFF');
  }
  assert.equal(sheet.B2.v,' \n ');assert.equal(sheet.C2.v,'false');assert.equal(sheet.D2.v,0);assert.equal(sheet.D2.t,'n');assert.equal(sheet.F2.f,'D2-E2');assert.equal(sheet.F3.f,'SUM(F2:F2)');assert.equal(sheet.D2.z,'#,##0.00');
  const original=await Zip.loadAsync(source);assert.equal(await zip.file('xl/worksheets/sheet1.xml').async('string'),await original.file('xl/worksheets/sheet1.xml').async('string'));
});

test('原工作表XML、批注、格式、隐藏和无关ZIP原生部分逐字节保留',async()=>{
  const source=await fixture({calcChain:true}),original=await Zip.loadAsync(source),{book,zip}=await unpack(await Export(source,result()));
  const changed=new Set(['xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/styles.xml','[Content_Types].xml','docProps/app.xml','xl/calcChain.xml']);
  for(const [name,file] of Object.entries(original.files))if(!file.dir&&!changed.has(name)){assert(zip.file(name),name);assert.deepEqual(await zip.file(name).async('uint8array'),await file.async('uint8array'),name);}
  assert.equal(book.Sheets.明细甲.B3.f,'B2-23.45');assert.equal(book.Sheets.明细甲.A2.c[0].t,'原批注保持');assert.equal(book.Sheets.明细甲['!rows'][1].hidden,true);assert.equal(book.Sheets.明细甲['!cols'][1].hidden,true);assert.equal(book.Workbook.Sheets[2].Hidden,2);
  assert.equal(book.Workbook.Names.find(n=>n.Name==='明细局部').Sheet,1);assert.equal(book.Workbook.Names.find(n=>n.Name==='跨表名称').Ref,"'明细甲'!$B$3");
  const workbook=await zip.file('xl/workbook.xml').async('string');assert.match(workbook,/<workbookView[^>]*activeTab="0" firstSheet="0"/);assert(!zip.file('xl/calcChain.xml'));
  assert.doesNotMatch(await zip.file('xl/_rels/workbook.xml.rels').async('string'),/calcChain/);assert.doesNotMatch(await zip.file('[Content_Types].xml').async('string'),/calcChain/);
});

for(const summaryIndex of [0,1,2])test(`原汇总位于${summaryIndex+1}时原位更新并移到第一，重复导出不重复`,async()=>{
  const source=await fixture({summaryIndex}),original=await Zip.loadAsync(source),firstBytes=await Export(source,result()),{book,zip}=await unpack(firstBytes),oldPath=`xl/worksheets/sheet${summaryIndex+1}.xml`;
  assert.deepEqual(book.SheetNames,[NAME,'明细甲','明细乙']);assert.equal(await summaryPath(zip),oldPath);assert(!zip.file(oldPath.replace(/([^/]+)$/,'_rels/$1.rels')));
  assert.equal(book.Workbook.Names.find(n=>n.Name==='明细局部').Sheet,1);assert.equal(book.Workbook.Names.find(n=>n.Name==='自定义汇总名称').Sheet,0);assert(!book.Workbook.Names.some(n=>n.Name==='_xlnm.Print_Area'));
  const comments=Object.keys(zip.files).filter(p=>/comments\d+\.xml$/.test(p));assert.equal(comments.length,1);assert.match(await zip.file(comments[0]).async('string'),/原批注保持/);
  const preserved=Object.keys(original.files).filter(p=>/^xl\/worksheets\/sheet\d+\.xml$/.test(p)&&p!==oldPath);for(const p of preserved)assert.equal(await zip.file(p).async('string'),await original.file(p).async('string'));
  const repeated=await unpack(await Export(firstBytes,result([group({totalCents:1000,paidCents:1000,remainingCents:0,state:'已支付'})])));assert.deepEqual(repeated.book.SheetNames,book.SheetNames);assert.equal(repeated.book.Sheets[NAME].D2.v,10);assert.equal(repeated.book.Sheets[NAME].F2.v,0);assert.equal(await summaryPath(repeated.zip),oldPath);
});

test('空记录、零、负数和未知月份都有正确的公式缓存，不制造反向SUM范围',async()=>{
  const source=await fixture(),empty=(await unpack(await Export(source,result([])))).book.Sheets[NAME];
  assert.equal(empty['!ref'],'A1:G2');assert.equal(empty['!autofilter'].ref,'A1:G1');for(const c of ['D','E','F']){assert.equal(empty[c+'2'].v,0);assert.equal(empty[c+'2'].f,'SUM(0)');}
  const data=result([group({monthKey:'',totalCents:0,paidCents:0,remainingCents:0}),group({monthKey:'2026-12',category:'=字面分类',state:'<未支付>&',totalCents:-123,paidCents:0,remainingCents:-123})]);
  const sheet=(await unpack(await Export(source,data))).book.Sheets[NAME];assert.equal(sheet.A2.v,'时间待定');assert.equal(sheet.D2.v,0);assert.equal(sheet.D3.v,-1.23);assert.equal(sheet.F3.v,-1.23);assert.equal(sheet.B3.t,'s');assert.equal(sheet.B3.v,'=字面分类');assert.equal(sheet.B3.f,undefined);assert.equal(sheet.C3.v,'<未支付>&');assert.equal(sheet.F4.v,-1.23);
});

for(const date1904 of [false,true])test(`${date1904?'1904':'1900'}日期系统的汇总月份同为文本，原表日期及公式不变`,async()=>{
  const {book}=await unpack(await Export(await fixture({date1904}),result(undefined,{date1904})));
  assert.equal(book.Sheets[NAME].A2.v,'2026年11月');assert.equal(book.Sheets[NAME].A2.t,'s');assert.equal(Boolean(book.Workbook.WBProps.date1904),date1904);
  assert.equal(book.Sheets.明细甲.D2.v,date1904?44865:46327);assert.equal(book.Sheets.明细甲.D2.t,'n');assert.equal(book.Sheets.明细甲.D2.f,'DATE(2026,11,1)');
  assert.equal([book.Sheets[NAME].A2,book.Sheets[NAME].D2].filter(c=>c.t==='n').reduce((sum,c)=>sum+c.v,0),123.45);
});

test('1900年闰年兼容及最早月份不偏移',async()=>{
  const groups=['1900-01','1900-02','1900-03'].map(monthKey=>group({monthKey}));
  const sheet=(await unpack(await Export(await fixture(),result(groups)))).book.Sheets[NAME];assert.deepEqual([sheet.A2.v,sheet.A3.v,sheet.A4.v],['1900年1月','1900年2月','1900年3月']);
});

test('受重排影响的三维公式拒绝，保持范围的三维公式及字面值允许',async()=>{
  await assert.rejects(Export(await fixture({summaryIndex:1,formula:"SUM('明细甲:明细乙'!B2)"}),result()),/三维公式引用/);
  await assert.rejects(Export(await fixture({summaryIndex:1,formula:"SUM('明细甲':'付款预算汇总'!B2)"}),result()),/三维公式引用/);
  await assert.rejects(Export(await fixture({summaryIndex:1,formula:'SUM(明细甲:明细乙!B2)'}),result()),/三维公式引用/);
  for(const formula of ["SUM('明细甲:明细乙'!B2)",'"明细甲:明细乙!B2"',"SUM('[外部.xlsx]明细甲:明细乙'!B2)"]){const source=await fixture({formula}),{book}=await unpack(await Export(source,result()));assert.equal(book.Sheets.明细甲.B3.f,formula);}
});

test('非法文件、被阻止结果、错误金额、过期日期系统拒绝导出且输入不变',async()=>{
  const source=await fixture(),copy=source.slice(),data=result(),snapshot=JSON.stringify(data);
  await assert.rejects(Export(new Uint8Array([1,2,3]),data),/\.xlsx/);
  await assert.rejects(Export(source,{...data,canExport:false}),/不能导出/);
  await assert.rejects(Export(source,{...data,errors:['缺少金额']}),/不能导出/);
  await assert.rejects(Export(source,{...data,totals:{...data.totals,totalCents:NaN}}),/金额无效/);
  await assert.rejects(Export(source,{...data,totals:{...data.totals,remainingCents:0}}),/剩余金额/);
  await assert.rejects(Export(source,result([group({monthKey:'2026-13'})])),/月份无效/);
  await assert.rejects(Export(source,{...data,date1904:true}),/日期系统不一致/);
  const zip=await Zip.loadAsync(source);zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml').async('string')).replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml'));
  await assert.rejects(Export(await zip.generateAsync({type:'uint8array'}),data),/仅支持 .xlsx/);
  assert.deepEqual(source,copy);assert.equal(JSON.stringify(data),snapshot);
});

test('浏览器UMD入口与CommonJS相同，无DOM或新依赖',async()=>{
  const context={XLSX:X,JSZip:Zip};vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../funds-budget-summary-export.js'),'utf8'),context);
  assert.equal(typeof context.FundsBudgetSummaryExport,'function');const bytes=await context.FundsBudgetSummaryExport(await fixture(),result());assert(bytes instanceof Uint8Array);assert.equal((await unpack(bytes)).book.Sheets[NAME].F2.f,'D2-E2');
});

test('安全整数分值转换为元会丢分时拒绝分组与合计导出',async()=>{
  const cents=9007199254740982,source=await fixture();
  assert.equal(Number.isSafeInteger(cents),true);assert.equal(Math.round(cents/100*100),9007199254740981);
  for(const key of ['totalCents','paidCents','remainingCents']){
    const data=result();data.groups[0][key]=cents;
    await assert.rejects(Export(source,data),/丢失分位精度/);
    const totalsOnly=result();totalsOnly.totals[key]=cents;
    await assert.rejects(Export(source,totalsOnly),/丢失分位精度/);
  }
});

test('清理旧汇总的图表附件但保留被原明细引用的共享图片',async()=>{
  const zip=await Zip.loadAsync(await fixture({summaryIndex:1})),R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',P='http://schemas.openxmlformats.org/package/2006/relationships';
  const relationships=body=>`<Relationships xmlns="${P}">${body}</Relationships>`;
  const oldRels='xl/worksheets/_rels/sheet2.xml.rels';
  zip.file(oldRels,(await zip.file(oldRels).async('string')).replace('</Relationships>',`<Relationship Id="drawing" Type="${R}/drawing" Target="../drawings/old-summary.xml"/></Relationships>`));
  zip.file('xl/drawings/old-summary.xml','<old-summary/>');
  zip.file('xl/drawings/_rels/old-summary.xml.rels',relationships(`<Relationship Id="chart" Type="${R}/chart" Target="../charts/old-summary.xml"/><Relationship Id="shared" Type="${R}/image" Target="../media/shared.png"/>`));
  zip.file('xl/charts/old-summary.xml','<old-chart/>');zip.file('xl/media/shared.png',new Uint8Array([1,2,3,4]));
  const originalRels='xl/worksheets/_rels/sheet1.xml.rels';zip.file(originalRels,(await zip.file(originalRels).async('string')).replace('</Relationships>',`<Relationship Id="shared" Type="${R}/image" Target="../media/shared.png"/></Relationships>`));
  const before=await zip.generateAsync({type:'uint8array'}),output=await Zip.loadAsync(await Export(before,result()));
  for(const p of ['xl/drawings/old-summary.xml','xl/drawings/_rels/old-summary.xml.rels','xl/charts/old-summary.xml',oldRels])assert(!output.file(p),p);
  assert.deepEqual(await output.file('xl/media/shared.png').async('uint8array'),new Uint8Array([1,2,3,4]));assert.equal(await output.file(originalRels).async('string'),await zip.file(originalRels).async('string'));
});

test('带命名空间前缀的OOXML保持可读，重复导出所有新样式引用有效',async()=>{
  const zip=await Zip.loadAsync(await fixture());
  for(const p of ['xl/workbook.xml','xl/styles.xml']){let xml=await zip.file(p).async('string');xml=xml.replace(/<(\/?)([A-Za-z][\w]*)(?=[\s/>])/g,'<$1s:$2').replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"','xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');zip.file(p,xml);}
  const first=await Export(await zip.generateAsync({type:'uint8array'}),result()),second=await Export(first,result()),{book,zip:out}=await unpack(second);
  assert.equal(book.Sheets[NAME].F2.v,100);assert.equal(book.Sheets[NAME].A2.w,'2026年11月');
  const styles=await out.file('xl/styles.xml').async('string'),xfCount=Number(styles.match(/<s:cellXfs count="(\d+)"/)[1]),sheet=await out.file(await summaryPath(out)).async('string');
  for(const match of sheet.matchAll(/<c\b[^>]*\bs="(\d+)"/g))assert(Number(match[1])<xfCount);
});

test('文字包含Excel转义模式和XML控制字符时原样回读',async()=>{
  const category='_x0041_\u000b<&>"';const {book}=await unpack(await Export(await fixture(),result([group({category})])));assert.equal(book.Sheets[NAME].B2.v,category);
});
