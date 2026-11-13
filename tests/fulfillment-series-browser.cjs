const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const HEADERS=['采购下单人','事业部','销售分渠道','供应商简称','系列','上期未交付数据','本期未交付数据','差异数量','备注','判断'];
function fixture(stage){
 const book=X.utils.book_new();
 for(const [index,name]of [...C.SHEETS,'孙立柱'].entries()){
  const records=index===4?[['排除','其他采购员','其他系列',999]]:stage===1?
   [[name+'-共','采购员甲','共用系列',index+1],...(index===0?[['消失','采购员乙','旧系列',5],['小数1','小数采购员','小数系列',0.1],['小数2','小数采购员','小数系列',0.2]]:[])]:
   [[name+'-共','采购员甲','共用系列',index+2],...(index===0?[['新增','采购员乙','新系列',7],['小数1','小数采购员','小数系列',0.3],['零值','零值采购员','零值系列',0]]:[])];
  const rows=records.map(([code,owner,series,quantity])=>C.COLUMNS.map(key=>({
   '物料编码':code,'SKU':'SKU-'+code,'物料名称':'物料'+code,'采购下单人':owner,
   '事业部':'国内事业部','销售分渠道':'国内事业部*TEMU业务部','系列':series,
   '供应商简称':'合成供应商','运营':'王志豪','未交付数量':stage===3&&index===0?'错误数量':quantity,
   '下单月份':'2026年11月','下单日期':'2026年11月1日'
  })[key]??''));
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...rows]),name);
 }
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
function identityFixture(stage){
 const current=stage===2,book=X.utils.book_new();
 const records=[
  [
   {order:'旧月份订单',code:'OLD',month:'2025年12月',qty:current?70:50},
   {order:'采购订单1',code:'A',month:'2026年11月',qty:current?10:5},
   {order:'采购订单2',code:'B',month:'2026年11月',qty:6},
   ...(current?[{order:'采购订单4',code:'D',month:'2026年11月',qty:4},{order:'采购订单5',code:'E',month:'2026年11月',qty:5}]:[])
  ],
  [{order:'采购订单3',code:'C',month:'2026年11月',qty:current?8:7}],
  [{order:'其他采购员订单',code:'OTHER-OWNER',month:'2027年1月',qty:current?2:1,buyer:'其他采购员'}],
  [
   {order:'另一系列订单',code:'OTHER-SERIES',month:'2026年12月',qty:current?3:2,series:'另一系列'},
   {order:'乙采购订单1',code:'BA',month:'2026年12月',qty:current?5:3,supplier:'供应商乙'},
   {order:'乙采购订单2',code:'BB',month:'2026年12月',qty:4,supplier:'供应商乙'},
   ...(current?[{order:'乙采购订单3',code:'BC',month:'2026年12月',qty:6,supplier:'供应商乙'}]:[])
  ]
 ];
 if(current)[records[0][1],records[0][2]]=[records[0][2],records[0][1]]; // Existing orders can move without becoming new lines.
 C.SHEETS.forEach((name,index)=>{
  const data=records[index].map(record=>C.COLUMNS.map(key=>({
   '采购下单人':record.buyer||'判断采购员','事业部':'国内事业部','销售分渠道':'国内事业部*TEMU业务部',
   '供应商简称':record.supplier||'供应商甲','系列':record.series||'新增定位系列','采购订单号':record.order,'物料编码':record.code,'SKU':'SKU-'+record.code,
   '下单月份':record.month,'下单日期':record.month+'1日','未交付数量':record.qty,
   '合同约定交期':current?'2026年12月31日':'2026年11月30日','是否正常履约':current?'已变更':'原状态'
  })[key]??''));
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...data]),name);
 });
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[],base=process.env.QA_URL||'http://127.0.0.1:8940/';
  page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('fulfillment.html',base).href);await page.waitForFunction(()=>!document.getElementById('file').disabled);
  assert.equal(await page.locator('#download-review').count(),0);
  const upload=async(bytes,name)=>{await page.locator('#file').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});await page.waitForFunction(()=>!document.getElementById('file').disabled);};
  const history=()=>page.evaluate(()=>FulfillmentCache.get());
  const download=async selector=>{const pending=page.waitForEvent('download');await page.locator(selector).click();const file=await pending,bytes=await fs.readFile(await file.path());return {bytes,book:X.read(bytes,{cellNF:true,cellStyles:true}),name:file.suggestedFilename()};};
  const rows=book=>X.utils.sheet_to_json(book.Sheets.系列对比,{header:1,defval:''});
  await upload(fixture(1),'上期更新版.xlsx');assert.equal((await history()).previous,null);
  const first=await download('#download');assert.deepEqual(rows(first.book),[
   HEADERS,['采购员甲','国内事业部','国内事业部*TEMU业务部','合成供应商','共用系列','',10,'','',''],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','合成供应商','旧系列','',5,'','',''],
   ['小数采购员','国内事业部','国内事业部*TEMU业务部','合成供应商','小数系列','',0.3,'','','']
  ]);
  await upload(fixture(2),'本期更新版.xlsx');const expected=[
   HEADERS,['采购员甲','国内事业部','国内事业部*TEMU业务部','合成供应商','共用系列',10,14,4,'',''],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','合成供应商','新系列',0,7,7,'','2026年11月净增1行（上期0行，本期1行）；本期相关行（新增行待确认）：徐英第3行'],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','合成供应商','旧系列',5,0,-5,'','']
  ];
  let saved=await download('#download');assert.deepEqual(rows(saved.book),expected);
  assert.deepEqual(saved.book.SheetNames,['汇总表','合并核对','系列对比',...C.SHEETS,'孙立柱']);
  const baseline=await history();assert.equal(baseline.previous.filename,'上期更新版.xlsx');
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);
  assert.deepEqual(rows((await download('#download')).book),expected);
  await upload(fixture(2),'重复副本.xlsx');const duplicate=await history();assert.deepEqual(duplicate.previous,baseline.previous);assert.deepEqual(duplicate.current.result.rows,baseline.current.result.rows);
  await upload(fixture(3),'异常更新版.xlsx');assert(await page.locator('#download').isDisabled());for(const button of await page.locator('[data-download-division]').all())assert(await button.isDisabled());assert.match(await page.locator('#status').innerText(),/数量校验未通过/);assert.match(await page.locator('#review-preview').textContent(),/数量／转换异常/);assert.deepEqual(await history(),duplicate);
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);saved=await download('#download');assert.deepEqual(rows(saved.book),expected);
  const division=await download('[data-download-division="国内事业部"]');assert.deepEqual(division.book.SheetNames,['汇总表']);
  assert.deepEqual(X.utils.sheet_to_json(saved.book.Sheets.合并核对,{header:1})[0],C.REVIEW_COLUMNS);
  const sheet=saved.book.Sheets.系列对比;assert.equal(sheet['!autofilter'].ref,'A1:J4');assert.equal(sheet.A1.s.fgColor.rgb,'315B9D');
  for(let r=2;r<=4;r++){for(const col of ['F','G','H'])assert.equal(sheet[col+r].t,'n');assert.equal(sheet['D'+r].v,'合成供应商');assert.equal(sheet['I'+r]?.v??'','');assert.equal(sheet['J'+r].t,'s');assert.equal(sheet['J'+r].z,'@');}
  const style=await page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=x=>p.parseFromString(x,'application/xml');
   const wb=parse(await z.file('xl/workbook.xml').async('string')),rels=parse(await z.file('xl/_rels/workbook.xml.rels').async('string')),entry=all(wb,'sheet').find(e=>e.getAttribute('name')==='系列对比');
   const rel=all(rels,'Relationship').find(e=>e.getAttribute('Id')===entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
   const doc=parse(await z.file(new URL(rel.getAttribute('Target'),'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1)).async('string')),styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],borders=all(styles,'borders')[0];
   return {formulaCount:all(doc,'f').length,hiddenRows:all(doc,'row').filter(r=>r.getAttribute('hidden')==='1').length,filterCount:all(doc,'filterColumn').length,protected:all(doc,'sheetProtection').length>0,aligned:all(doc,'c').every(c=>{const xf=xfs.children[Number(c.getAttribute('s')||0)],a=all(xf,'alignment')[0];return a?.getAttribute('horizontal')==='center'&&a.getAttribute('vertical')==='center';}),bordered:all(doc,'c').filter(c=>all(c,'v').length||all(c,'t').length).every(c=>{const xf=xfs.children[Number(c.getAttribute('s')||0)],border=borders.children[Number(xf.getAttribute('borderId')||0)];return ['left','right','top','bottom'].every(edge=>!!all(border,edge)[0]?.getAttribute('style'));})};
  },[...saved.bytes]);
  assert.deepEqual(style,{formulaCount:0,hiddenRows:0,filterCount:0,protected:false,aligned:true,bordered:true});
  const manualNote='合成人工备注：等待供应商反馈',manual=X.read(saved.bytes,{cellNF:true,cellStyles:true});manual.Sheets.系列对比.I2={t:'s',v:manualNote,z:'@'};
  for(const name of ['差异','差异（本次）'])X.utils.book_append_sheet(manual,X.utils.aoa_to_sheet([['旧差异'],[999]]),name);
  const beforeNote=await history();await upload(Buffer.from(X.write(manual,{type:'buffer',bookType:'xlsx'})),'已填写备注.xlsx');
  const afterNote=await history();assert.deepEqual(afterNote.previous,beforeNote.previous);assert.deepEqual(afterNote.current.result.rows,beforeNote.current.result.rows,'只填写备注不推进历史');
  const noteOutput=await download('#download');assert.deepEqual(rows(noteOutput.book),expected,'本次生成备注仍留空');
  assert.equal(noteOutput.book.Sheets['系列对比（原表）'].I2.v,manualNote,'原系列对比表的人工备注保留在备份表');
  for(const name of ['差异','差异（本次）'])assert(!noteOutput.book.SheetNames.includes(name),'再上传时去掉旧独立差异表：'+name);
  const legacyNote='合成旧版9列备注',legacy=X.read(saved.bytes,{cellNF:true,cellStyles:true});
  const legacyRows=rows(saved.book).map(row=>row.filter((_,index)=>index!==3));legacyRows[1][7]=legacyNote;
  legacy.Sheets.系列对比=X.utils.aoa_to_sheet(legacyRows);
  await upload(Buffer.from(X.write(legacy,{type:'buffer',bookType:'xlsx'})),'旧9列填写备注.xlsx');
  const legacyHistory=await history();assert.deepEqual(legacyHistory.previous,beforeNote.previous);assert.deepEqual(legacyHistory.current.result.rows,beforeNote.current.result.rows,'旧9列备注表不推进历史');
  const legacyOutput=await download('#download');assert.deepEqual(rows(legacyOutput.book),expected);
  assert.deepEqual(X.utils.sheet_to_json(legacyOutput.book.Sheets['系列对比（原表）'],{header:1,defval:''}),legacyRows,'旧9列原表及H列人工备注完整保留在备份表');
  await upload(identityFixture(2),'本期判断更新版.xlsx');
  const previousChooser=page.waitForEvent('filechooser');await page.locator('#replace-previous').click();await(await previousChooser).setFiles({name:'上期判断更新版.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:identityFixture(1)});
  await page.waitForFunction(()=>!document.getElementById('previous-file').disabled&&!document.getElementById('previous-file').value);
  const judgmentExpected=[HEADERS,
   ['判断采购员','国内事业部','国内事业部*TEMU业务部','供应商甲','新增定位系列',68,103,35,'','2026年11月净增2行（上期3行，本期5行）；涉及行：徐英第5-6行'],
   ['其他采购员','国内事业部','国内事业部*TEMU业务部','供应商甲','新增定位系列',1,2,1,'',''],
   ['判断采购员','国内事业部','国内事业部*TEMU业务部','供应商甲','另一系列',2,3,1,'',''],
   ['判断采购员','国内事业部','国内事业部*TEMU业务部','供应商乙','新增定位系列',7,15,8,'','2026年12月净增1行（上期2行，本期3行）；涉及行：李奇第5行']
  ];
  const judgmentOutput=await download('#download');assert.deepEqual(rows(judgmentOutput.book),judgmentExpected,'同五维组的数量、最新月份和新增来源行不串供应商，其他月份及数量/交期/状态变化不算新增');
  assert.equal(judgmentOutput.book.Sheets.系列对比['!autofilter'].ref,'A1:J5');
  for(let r=2;r<=5;r++){assert.equal(judgmentOutput.book.Sheets.系列对比['I'+r]?.v??'','');assert.equal(judgmentOutput.book.Sheets.系列对比['J'+r].t,'s');assert.equal(judgmentOutput.book.Sheets.系列对比['J'+r].z,'@');}
  const judgmentHistory=await history();assert.equal(judgmentHistory.previous.filename,'上期判断更新版.xlsx');
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);assert.deepEqual(rows((await download('#download')).book),judgmentExpected,'刷新后仍使用同一上期判断');
  await upload(identityFixture(2),'判断重复副本.xlsx');const judgmentDuplicate=await history();assert.deepEqual(judgmentDuplicate.previous,judgmentHistory.previous);assert.deepEqual(judgmentDuplicate.current.result.rows,judgmentHistory.current.result.rows);assert.deepEqual(rows((await download('#download')).book),judgmentExpected,'重复上传不推进判断基准');
  assert.deepEqual(errors,[]);
  const out=process.env.QA_OUTPUT||'/private/tmp/fulfillment-series-qa';await fs.mkdir(out,{recursive:true});await fs.writeFile(path.join(out,'系列对比-合成验收.xlsx'),saved.bytes);
  console.log('PASS: ten-column supplier-separated series comparison, zero differences omitted, exact current-minus-previous differences, latest-month net row growth with identity/source references and uncertain fallback, blank editable remarks, current and legacy nine-column manual remark backups, four sources, first-period blanks, cache reload/duplicate/failure, division export and review sheet in full workbook, numeric/text types/filter/styles and no formulas');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
