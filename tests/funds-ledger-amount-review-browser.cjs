const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,LH,NH,AH,OH}=require('./funds-progress-fixtures.cjs');
const PREFIX='台账金额差异核对：',SUPPLIER='金额核验合成供应商';
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const sheet=(headers,records)=>X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(header=>record[header]??''))]);
function payment(amount,extra={}){
 return {'登记月份':46300,'供应商名称':SUPPLIER,'付款主体':'迈德斯特','对公付款0A流程号':'','开票时间':46300,'开票金额':amount,'摘要':'合成八月货款','申请付款金额':amount,'应付款时间':46332,'单据号【财务应付单】':'','OA合同编号':'','合同总金额':amount,'已付金额':0,'剩余金额':amount,'剩余付款金额':amount,'是否结清':'否','是否谈判付款':'是','是否提交付款申请单':'未提交','备注':'保留人工备注','谈判明细':'11月2日－'+amount,'分批总金额':amount,'核对':'原人工核对',46328:amount,46335:0,'2026年11月支付':amount,...extra};
}
function inputs(){
 const input=fixture();
 for(const person of C.PEOPLE)input.ledger.book.Sheets['对账明细-'+person]=sheet(LH,person==='徐英'?[
  payment(311234),payment(121200,{'付款主体':'国源'}),payment(312026,{'是否谈判付款':'否'})
 ]:[]);
 input.ledger.book.Sheets['对账明细-徐英'].H2={t:'n',v:311234,f:'305110+1800+447+214+3663'};
 const name=input.negotiated.book.SheetNames[0],source=sheet([...NH,'2026年11月支付'],[
  payment(312026,{'付款主体':'待核对','摘要':'待核对','对公付款0A流程号':'未提交OA','谈判明细':'1、11月2日－200,000.00\n2、11月9日－112,026.00',46328:200000,46335:112026}),
  payment(121200,{'付款主体':'国源','对公付款0A流程号':'未提交OA'})
 ]);
 source.M1.z=source.N1.z=C.DATE_FORMAT;
 source.K2={t:'n',v:312026,f:'SUM(M2:N2)'};source.F2={t:'n',v:312026,f:'D2-E2'};
 source.I2.c=[{a:'人工',t:'不可覆盖的人工批注'}];input.negotiated.book.Sheets[name]=source;
 input.applications.book.Sheets.付款申请汇总=sheet(AH,[]);
 input.oa.book.Sheets[input.oa.book.SheetNames[0]]=sheet(OH,[]);
 return input;
}
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8963/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[],input=inputs();page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  const frame=page.frameLocator('iframe[data-tool="funds"]');
  async function upload(role,bytes=write(input[role].book)){
   await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:'合成-'+role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
  }
  const snapshot=()=>frame.locator('#reviewPreview').evaluate(el=>({headers:[...el.querySelectorAll('th')].map(e=>e.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(td=>td.textContent))}));
  function verifyOriginal(book){
   const source=book.Sheets.谈判付款,addresses=Object.keys(source).filter(address=>/^[A-Z]+[1-9]\d*$/.test(address));
   const column=labels=>{
    const address=addresses.find(address=>X.utils.decode_cell(address).r<2&&labels.includes(source[address]?.v));
    assert(address,'缺少表头 '+labels.join('/'));return X.utils.decode_cell(address).c;
   };
   const amountColumn=column(['申请付款金额','OA-申请付款金额']);
   const amountAddress=addresses.find(address=>X.utils.decode_cell(address).c===amountColumn&&source[address]?.v===312026);
   assert(amountAddress,'谈判登记金额312026仍保留');const r=X.utils.decode_cell(amountAddress).r;
   const at=labels=>source[X.utils.encode_cell({r,c:column(labels)})];
   assert.equal(at(['供应商名称','供应商']).v,SUPPLIER);
   assert.equal(at(['2026年11月2日']).v,200000);assert.equal(at(['2026年11月9日']).v,112026);
   assert.equal(at(['2026年11月支付']).v,312026);assert.equal(at(['分批总金额']).v,312026);assert(at(['分批总金额']).f,'原分批金额公式保留');
   assert.equal(at(['谈判明细']).v,'1、11月2日－200,000.00\n2、11月9日－112,026.00');
   assert.equal(at(['备注']).v,'保留人工备注');assert.equal(at(['备注']).c[0].t,'不可覆盖的人工批注');
  }
  async function verify(expected,version){
   await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:60000});
   assert.equal(await frame.locator('#pendingCount').innerText(),'0','金额差异提示不阻止下载');
   const preview=await snapshot(),noteColumn=preview.headers.indexOf('核对说明');assert(noteColumn>=0);
   const shown=preview.rows.filter(row=>row[noteColumn].includes(PREFIX));assert.equal(shown.length,expected?1:0);
   if(expected)for(const token of [PREFIX,'312026.00','311234.00','792.00','合成-negotiated.xlsx','合成-ledger.xlsx','对账明细-徐英'])assert(shown[0][noteColumn].includes(token),token);
   const event=page.waitForEvent('download');await frame.locator('#download').click();const download=await event;
   assert.match(download.suggestedFilename(),new RegExp('^付款节奏跟进-\\d{4}-\\d{2}-\\d{2}-V'+version+'\\.xlsx$'));
   const bytes=await fs.readFile(await download.path()),book=X.read(bytes,{cellNF:true,cellFormula:true}),review=book.Sheets.付款申请核对;
   const grid=X.utils.sheet_to_json(review,{header:1,defval:''});assert.deepEqual(grid[0],['付款类别','供应商名称','付款主体','申请付款金额','是否提交付款申请单','核对说明','来源文件','来源工作表','来源行号','跟单员']);assert.deepEqual(grid[0],preview.headers);assert.equal(grid.length-1,preview.rows.length);
   for(let r=0;r<preview.rows.length;r++)for(let c=0;c<grid[0].length;c++)assert.equal(preview.rows[r][c],X.utils.format_cell(review[X.utils.encode_cell({r:r+1,c})]),'核对预览与下载一致');
   assert.equal(grid.slice(1).filter(row=>row[noteColumn].includes(PREFIX)).length,expected?1:0);
   assert.deepEqual(await snapshot(),preview);verifyOriginal(book);return bytes;
  }
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  const first=await verify(true,1);
  input.negotiated.book=X.read(first,{cellNF:true,cellFormula:true});await upload('negotiated',first);
  await verify(true,2);
  const ledger=input.ledger.book.Sheets['对账明细-徐英'];for(const address of ['F2','H2','L2','O2'])ledger[address]={t:'n',v:312026};
  await upload('ledger');await verify(false,3);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,uniqueEmptyOA:true,amounts:[312026,311234],difference:792,previewDownloadMatch:true,reuploadPreservesSource:true,equalLedgerClearsReview:true}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
