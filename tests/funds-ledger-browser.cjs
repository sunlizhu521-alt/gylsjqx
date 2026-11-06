const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const NAMES=['已申请OA付款','未申请OA付款'];
const HEADERS=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否提交付款申请单','是否结清','是否谈判付款','开票时间','所属月份','账期'];
const OUTPUT_HEADERS=['供应商','付款主体','对公付款OA流程号','申请付款金额','摘要','应付款时间','是否提交付款审单'];
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const value=cell=>cell&&typeof cell==='object'?cell.v??'':cell??'';

async function replaceSourceCell(bytes,address,cellXml){
 const zip=await Zip.loadAsync(bytes),file='xl/worksheets/sheet1.xml',xml=await zip.file(file).async('string');
 const pattern=new RegExp('<c\\b[^>]*\\br="'+address+'"[^>]*>[\\s\\S]*?<\\/c>');
 assert.match(xml,pattern,'公式边界用例必须修改实际来源单元格');
 zip.file(file,xml.replace(pattern,cellXml));
 return zip.generateAsync({type:'nodebuffer'});
}

function fixture(replacement=false){
 const book=X.utils.book_new(),rows=PEOPLE.map(()=>[]),expected=[[],[]];
 function add(person,changes={},group=0,expectedTime){
  const row=['合成供应商-'+person,'合成付款主体','000123456789012345',1234.5,'合成货款',46333,46340,'已提交','否','否',46300,'2026-10','60天'];
  for(const [header,v]of Object.entries(changes))row[HEADERS.indexOf(header)]=v;
  rows[person].push(row);
  if(group!==null)expected[group].push([row[0],row[1],row[2],row[3],row[4],expectedTime===undefined?row[group===0?5:6]:expectedTime,row[7]].map(value));
 }
 if(replacement)add(0,{'供应商名称':'合成替换唯一记录','申请付款金额':0});
 else{
  add(0);add(0); // Two identical source rows are two payments, not a deduplication opportunity.
  add(0,{'供应商名称':'合成缓存公式','申请付款金额':{t:'n',v:125,f:'100+25'},'应付款时间':{t:'n',v:46334,f:'DATE(2026,11,7)',z:'m/d/yy'}});
  add(0,{'供应商名称':'合成空金额','对公付款0A流程号':'0','申请付款金额':''});
  add(0,{'供应商名称':'合成空应付日期','对公付款0A流程号':'待补流程号','应付款时间':'','预估付款时间':46340});
  add(0,{'供应商名称':'合成零金额未申请','对公付款0A流程号':'','申请付款金额':0,'是否提交付款申请单':'未提交'},1);
  add(0,{'供应商名称':'合成空预估日期','对公付款0A流程号':'未提交OA','申请付款金额':'','预估付款时间':'','是否提交付款申请单':''},1);
  for(const [label,changes]of [
   ['已结清',{'是否结清':'是'}],['谈判付款',{'是否谈判付款':'是'}],
   ['空结清',{'是否结清':''}],['空谈判',{'是否谈判付款':''}],
   ['非精确否',{'是否结清':'否待确认'}],['双空',{'是否结清':'','是否谈判付款':''}]
  ])add(0,{'供应商名称':'必须排除-'+label,...changes},null);
  for(let i=0;i<64;i++)add(0,{'供应商名称':'合成长预览-'+String(i).padStart(2,'0'),'申请付款金额':i,'摘要':i===0?'=1+1':'合成完整预览第 '+i+' 行'});
  for(let p=1;p<PEOPLE.length;p++){
   add(p,{'供应商名称':'合成'+PEOPLE[p]+'已申请','申请付款金额':p+0.25});
   if(p===2){
    add(p,{'供应商名称':'合成邢芳芳五个月','对公付款0A流程号':'','所属月份':'2026-10','账期':'5个月账期(月底付款)50%承兑','是否提交付款申请单':'未提交'},1,46483); // 2027-04-06, using 150 days rather than calendar months.
    add(p,{'供应商名称':'合成邢芳芳60天','对公付款0A流程号':'','所属月份':'2026-10','账期':'60天','是否提交付款申请单':'未提交'},1,46393); // 2027-01-06.
   }else add(p,{'供应商名称':'合成'+PEOPLE[p]+'未申请','对公付款0A流程号':p===1?'  ':'未提交OA','预估付款时间':46340+p,'是否提交付款申请单':'未提交'},1);
  }
 }
 for(let p=0;p<PEOPLE.length;p++){
  const indices=HEADERS.map((_,c)=>c).filter(c=>p!==2||HEADERS[c]!=='预估付款时间'),headers=indices.map(c=>HEADERS[c]);
  const sheet=X.utils.aoa_to_sheet([headers,...rows[p].map(row=>indices.map(c=>row[c]))]);
  for(let r=2;r<=rows[p].length+1;r++)for(const header of ['应付款时间','预估付款时间','开票时间']){
   const c=headers.indexOf(header);if(c<0)continue;const cell=sheet[X.utils.encode_cell({r:r-1,c})];if(cell?.t==='n')cell.z='yyyy-mm-dd';
  }
  sheet['!rows']=[{},...(rows[p].map((_,r)=>r===0?{hidden:true}:{}))];
  sheet['!cols']=headers.map((_,c)=>({wch:8,hidden:c===3}));
  X.utils.book_append_sheet(book,sheet,'对账明细-'+PEOPLE[p]+(p===4?'20261106':''));
 }
 const history=X.utils.aoa_to_sheet([HEADERS,['历史干扰记录','合成主体','HISTORY',999999,'不得纳入',46333,46340,'已提交','否','否']]);
 X.utils.book_append_sheet(book,history,'对账明细-李奇-历史');
 return {book,expected,bytes:write(book)};
}

(async()=>{
 const base=process.env.BASE_URL||process.env.QA_URL||'http://127.0.0.1:8940/';
 const out=process.env.OUT_DIR||process.env.QA_OUT||'/private/tmp/funds-ledger-browser';
 await fs.mkdir(out,{recursive:true});
 const input=fixture(),replacement=fixture(true),browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000},timezoneId:'Asia/Shanghai'}),errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  await page.clock.setFixedTime(new Date('2026-11-06T02:00:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});
  const homepage=page.url();await page.locator('button[data-tool="funds-ledger"]').click();assert.equal(page.url(),homepage);
  const frame=page.frameLocator('iframe[data-tool="funds-ledger"]');
  await frame.locator('#file').waitFor();assert(await frame.locator('#download').isDisabled());
  assert.match(await page.locator('iframe[data-tool="funds-ledger"]').getAttribute('src'),/funds-ledger\.html/);
  async function upload(bytes,name){await frame.locator('#file:not([disabled])').waitFor();await frame.locator('#file').setInputFiles({name,mimeType:MIME,buffer:bytes});}
  async function ready(name){
   await frame.locator('#download:not([disabled])').waitFor();
   assert((await frame.locator('#reference').innerText()).includes(name),'来源文件名必须可见');
   assert.match(await frame.locator('#reference').innerText(),/2026.*11.*0?6/,'参考时间必须可见');
  }
  async function preview(){
   return Promise.all(['applied','unapplied'].map(async id=>({
    headers:await frame.locator('#'+id+'-preview th').allTextContents(),
    rows:await frame.locator('#'+id+'-preview tbody tr').evaluateAll(rows=>rows.filter(tr=>tr.cells.length===7).map(tr=>[...tr.cells].map(c=>c.textContent))),
    count:await frame.locator('#'+id+'-count').innerText()
   })));
  }
  async function download(label){
   const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;
   assert.match(d.suggestedFilename(),/\.xlsx$/);const bytes=await fs.readFile(await d.path());
   await fs.writeFile(path.join(out,label+'.xlsx'),bytes);return bytes;
  }
  async function verify(bytes,expected){
   const saved=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),snap=await preview();
   assert.deepEqual(saved.SheetNames,NAMES);
   for(let i=0;i<NAMES.length;i++){
    const sheet=saved.Sheets[NAMES[i]],headers=OUTPUT_HEADERS.map((h,c)=>c===5&&i===1?'预估付款时间':h);
    const range=X.utils.decode_range(sheet['!ref']);
    const matrix=Array.from({length:range.e.r-range.s.r+1},(_,r)=>Array.from({length:range.e.c-range.s.c+1},(_,c)=>sheet[X.utils.encode_cell({r:r+range.s.r,c:c+range.s.c})]?.v??''));
    assert.deepEqual(matrix,[headers,...expected[i]]);
    assert.deepEqual(snap[i].headers,headers);assert.equal(snap[i].rows.length,expected[i].length,'全部明细都必须进入预览');
    assert.equal(Number(snap[i].count.replace(/[^\d]/g,'')),expected[i].length);
    assert.equal(sheet['!autofilter'].ref,'A1:G'+(expected[i].length+1));
    assert(!saved.Workbook?.Sheets?.[i]?.Hidden,'输出工作表必须可见');
    for(let r=0;r<expected[i].length;r++)for(let c=0;c<7;c++){
     const address=X.utils.encode_cell({r:r+1,c}),cell=sheet[address],v=expected[i][r][c];
     assert.equal(cell?.v??'',v,NAMES[i]+'!'+address);assert(!cell?.f,'提取结果不得残留原公式');
     if(v!=='')assert.equal(cell.t,typeof v==='number'?'n':'s');
     if(c===3&&v!=='')assert.match(cell.z,/0\.00/,'金额采用两位小数');
     if(c===5&&typeof v==='number'){
      assert(X.SSF.is_date(cell.z),'日期必须为 Excel 日期格式');
      assert.equal(X.SSF.format('yyyy-mm-dd',cell.v),X.SSF.format('yyyy-mm-dd',v));
      const fallback={'合成邢芳芳五个月':'2027-04-06','合成邢芳芳60天':'2027-01-06'}[expected[i][r][0]];
      if(fallback)assert.equal(X.SSF.format('yyyy-mm-dd',cell.v),fallback,'缺少预估列时按月底＋账期天数＋7天补充');
     }
     if(c===5&&expected[i][r][0]==='合成空预估日期')assert.equal(cell?.v??'','','已有预估列的空白不得按月份和账期补填');
     assert.equal(snap[i].rows[r][c],cell?X.utils.format_cell(cell):'','预览和导出显示一致');
    }
   }
   const layout=await frame.locator('body').evaluate(async(_,data)=>{
    const z=await JSZip.loadAsync(Uint8Array.from(data)),parse=s=>new DOMParser().parseFromString(s,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
    const styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0].children,fonts=all(styles,'fonts')[0].children,fills=all(styles,'fills')[0].children,borders=all(styles,'borders')[0].children;
    const wb=parse(await z.file('xl/workbook.xml').async('string')),rels=all(parse(await z.file('xl/_rels/workbook.xml.rels').async('string')),'Relationship'),problems=[],fontNames=new Set(),fontSizes=new Set(),widths=[];
    for(const sheet of all(wb,'sheet')){
     const id=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),target=rels.find(r=>r.getAttribute('Id')===id).getAttribute('Target'),p=new URL(target,'https://x.invalid/xl/workbook.xml').pathname.slice(1),doc=parse(await z.file(p).async('string')),pane=all(doc,'pane')[0],columns=all(doc,'col');
     if(pane?.getAttribute('state')!=='frozen'||pane?.getAttribute('ySplit')!=='1')problems.push(p+':未冻结首行');
     if(all(doc,'filterColumn').length||all(doc,'conditionalFormatting').length||all(doc,'f').length)problems.push(p+':存在旧筛选、条件格式或公式');
     if(all(doc,'row').some(r=>r.getAttribute('hidden')==='1')||columns.some(c=>c.getAttribute('hidden')==='1'))problems.push(p+':存在隐藏行列');
     const sizes=Array(7);for(const col of columns)for(let c=Number(col.getAttribute('min'));c<=Number(col.getAttribute('max'));c++)sizes[c-1]=Number(col.getAttribute('width'));widths.push(sizes);
     for(const cell of all(doc,'c')){
      const ref=cell.getAttribute('r'),xf=xfs[Number(cell.getAttribute('s')||0)],font=fonts[Number(xf.getAttribute('fontId')||0)],border=borders[Number(xf.getAttribute('borderId')||0)];
      fontNames.add(all(font,'name')[0]?.getAttribute('val'));fontSizes.add(all(font,'sz')[0]?.getAttribute('val'));
      for(const edge of ['left','right','top','bottom'])if(!all(border,edge)[0]?.getAttribute('style'))problems.push(p+':'+ref+':缺少'+edge+'边框');
      if(/^[A-Z]+1$/.test(ref)){
       const fill=fills[Number(xf.getAttribute('fillId')||0)],color=all(fill,'fgColor')[0]?.getAttribute('rgb')||'',alignment=all(xf,'alignment')[0];
       if(!color.endsWith('315B9D'))problems.push(p+':'+ref+':表头非统一蓝色');
       if(alignment?.getAttribute('wrapText')==='1')problems.push(p+':'+ref+':表头换行');
      }
     }
    }
    return {problems,fontNames:[...fontNames],fontSizes:[...fontSizes],widths};
   },[...bytes]);
   assert.deepEqual(layout.problems,[]);assert.equal(layout.fontNames.length,1);assert.equal(layout.fontSizes.length,1);
   assert(layout.fontNames[0]);assert(Number(layout.fontSizes[0])>0);
   for(const widths of layout.widths){assert.equal(widths.length,7);assert(widths.every(w=>w>=12&&w<=255),'列宽须适合阅读');assert(widths[2]>=20,'OA 列须完整呈现长文本');}
   return snap;
  }
  await upload(input.bytes,'合成资金台账.xlsx');await ready('合成资金台账.xlsx');
  const first=await download('initial'),snap=await verify(first,input.expected);
  assert.equal(snap[0].rows.length,73);assert.equal(snap[1].rows.length,7);
  assert(!JSON.stringify(snap).includes('必须排除-'));assert(!JSON.stringify(snap).includes('历史干扰记录'));
  const api=await frame.locator('body').evaluate(async(_,{bytes,name})=>{
   const result=FundsLedgerCore.build(XLSX.read(Uint8Array.from(bytes),{type:'array',cellNF:true}),name),output=await FundsLedgerExport(result);
   return {names:result.sheets.map(s=>s.name),counts:result.sheets.map(s=>s.rows.length),byteArray:output instanceof Uint8Array};
  },{bytes:[...input.bytes],name:'合成资金台账.xlsx'});
  assert.deepEqual(api,{names:NAMES,counts:[73,7],byteArray:true});
  await page.screenshot({path:path.join(out,'funds-ledger-home.png')});
  await page.locator('button[data-tool="funds"]').click();
  await page.frameLocator('iframe[data-tool="funds"]').locator('#ledgerFile').waitFor();
  await page.locator('button[data-tool="funds-ledger"]').click();assert.deepEqual(await preview(),snap);
  assert(await frame.locator('#download').isEnabled());
  await upload(replacement.bytes,'合成最新替换.xlsx');await ready('合成最新替换.xlsx');
  await verify(await download('replacement'),replacement.expected);
  assert(!(await frame.locator('#result').innerText()).includes('合成长预览-'));
  const invalid=X.read(input.bytes,{type:'buffer'});invalid.Sheets['对账明细-李奇'].D1.v='无效金额字段';
  await upload(write(invalid),'合成无效表头.xlsx');
  await frame.locator('#status').filter({hasText:/缺少|缺失|找不到|未找到|失败|无法/}).waitFor();
  assert(await frame.locator('#download').isDisabled(),'无效上传禁止下载旧结果');
  assert(!(await frame.locator('#result').innerText()).includes('合成替换唯一记录'),'无效上传清除旧预览');
  await upload(input.bytes,'合成恢复.xlsx');await ready('合成恢复.xlsx');
  await verify(await download('recovered'),input.expected);
  for(const [address,cellXml]of [
   ['D2','<c r="D2"><f>1+1</f></c>'],
   ['I2','<c r="I2" t="str"><f>IF(1=1,&quot;否&quot;,&quot;是&quot;)</f></c>']
  ]){
   let invalidFormula=await replaceSourceCell(input.bytes,address,cellXml);
   if(address==='I2')invalidFormula=await replaceSourceCell(invalidFormula,'H2','<c r="H2"/>');
   await upload(invalidFormula,'合成缺缓存-'+address+'.xlsx');
   await frame.locator('#status').filter({hasText:/公式缺少.*缓存/}).waitFor();
   assert(await frame.locator('#download').isDisabled(),address+' 无缓存公式不得下载');
   assert(!(await frame.locator('#result').innerText()).includes('合成供应商-0'),'缺缓存公式必须清除旧结果');
   await upload(input.bytes,'合成公式恢复-'+address+'.xlsx');await ready('合成公式恢复-'+address+'.xlsx');
  }
  const emptyFormula=await replaceSourceCell(input.bytes,'E2','<c r="E2" t="str"><f>IF(1=1,&quot;&quot;,&quot;货款&quot;)</f><v/></c>');
  const emptyExpected=input.expected.map(rows=>rows.map(row=>[...row]));emptyExpected[0][0][4]='';
  await upload(emptyFormula,'合成有效空缓存.xlsx');await ready('合成有效空缓存.xlsx');
  await verify(await download('empty-formula-cache'),emptyExpected);
  const standalone=await browser.newPage({viewport:{width:1440,height:1000}});standalone.on('pageerror',e=>errors.push(e.message));
  await standalone.goto(new URL('funds-ledger.html',base).href,{waitUntil:'networkidle'});
  await standalone.locator('#file:not([disabled])').waitFor();
  await standalone.locator('#file').setInputFiles({name:'合成独立页面.xlsx',mimeType:MIME,buffer:replacement.bytes});
  await standalone.locator('#download:not([disabled])').waitFor();assert((await standalone.locator('#result').innerText()).includes('合成替换唯一记录'));
  await standalone.screenshot({path:path.join(out,'funds-ledger-standalone.png')});await standalone.close();
  assert.deepEqual(errors,[]);
  console.log(JSON.stringify({base,out,sheets:NAMES,rows:[73,7],preview:true,export:true,styles:true,toolSwitch:true,latestUpload:true,invalidRecovery:true,missingEstimateFallback:true,existingBlankPreserved:true,formulaCacheGuard:true,emptyFormulaCache:true,standalone:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
