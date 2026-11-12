const assert=require('node:assert/strict'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const C=require('../funds-core.js'),{fixtures}=require('./funds-fixtures.cjs');

(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const base=process.env.QA_URL||'http://127.0.0.1:8940/',page=await browser.newPage();
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('funds.html',base).href);
  await page.addScriptTag({path:path.resolve('funds-export.js')});
  const exportBytes=async(bytes,result)=>Buffer.from(await page.evaluate(async({bytes,result})=>[...new Uint8Array(await (await FundsExport(Uint8Array.from(bytes),result)).arrayBuffer())],{bytes:[...bytes],result}));
  const xmlCell=async(bytes,address)=>page.evaluate(async({bytes,address})=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml');
   return s.serializeToString([...doc.getElementsByTagNameNS('*','c')].find(c=>c.getAttribute('r')===address));
  },{bytes:[...bytes],address});

  for(const date1904 of [false,true])for(const headerRow of [0,1]){
   const serial=(Date.UTC(2026,10,13)-Date.UTC(1899,11,30))/86400000-(date1904?1462:0),h=headerRow+1,r=h+1;
   const headers=Array.from({length:29},(_,c)=>c===0?'供应商':'字段'+(c+1)),source=X.utils.aoa_to_sheet([...(headerRow?[['谈判付款计划']]:[]),headers]);
   for(const c of ['B','F','AA','AB','AC'])source[c+h]={t:'n',v:serial,z:c==='AA'?'yyyy"年"m"月"':'yyyy-mm-dd'};
   source['F'+h].f='DATE(2026,11,13)';
   source['C'+h]={t:'s',v:'2026年11月13日'};
   source['D'+h]={t:'s',v:''};source['E'+h]={t:'n',v:0};
   source['AB'+h].c=[{a:'原作者',t:'保留表头说明'}];
   source['A'+r]={t:'s',v:'测试供应商'};source['B'+r]={t:'n',v:100.25,z:'#,##0.00'};
   source['B'+(r+1)]={t:'n',v:0,z:'#,##0.00'};source['B'+(r+2)]={t:'n',v:-25,z:'#,##0.00'};
   source['AA'+r]={t:'n',v:serial,z:'yyyy-mm-dd'};source['AB'+r]={t:'n',v:100.25,f:'B'+r};
   source['H'+r]={t:'n',v:75.25,f:'SUM(B'+r+':B'+(r+2)+')'};
   source['!ref']='A1:AC'+(r+2);
   const book=X.utils.book_new();book.Workbook={WBProps:{date1904}};X.utils.book_append_sheet(book,source,'谈判付款');
   const input=Buffer.from(await page.evaluate(async({bytes,address})=>{
    const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
    const doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),styles=p.parseFromString(await zip.file('xl/styles.xml').async('string'),'application/xml'),xfs=all(styles,'cellXfs')[0],cell=all(doc,'c').find(c=>c.getAttribute('r')===address);
    const xf=xfs.children[Number(cell.getAttribute('s')||0)].cloneNode(true);xf.setAttribute('numFmtId','31');cell.setAttribute('s',xfs.children.length);xfs.append(xf);xfs.setAttribute('count',xfs.children.length);
    zip.file('xl/styles.xml',s.serializeToString(styles));zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
   },{bytes:[...X.write(book,{type:'buffer',bookType:'xlsx'})],address:'B'+h}));
   const result={normalHeaders:['供应商名称'],normalRows:[],reviewHeaders:['核对说明'],reviewRows:[],negotiated:{sheetName:'谈判付款',headerRow,headers,rows:[],updates:[],sourceRange:X.utils.decode_range(source['!ref']),managedColumns:Array.from({length:28},(_,c)=>c),columnOrder:headers.map((_,c)=>c)}};
   const originalCustom=await xmlCell(input,'AC'+h);
   const verify=async bytes=>{
    const saved=X.read(bytes,{cellNF:true}),sheet=saved.Sheets.谈判付款;
    assert.equal(saved.Workbook.WBProps.date1904,date1904);
    for(const col of ['B','F','AA','AB']){
     const cell=sheet[col+h];assert.equal(cell.t,'s',col+h+' 日期表头必须是文本，不能参与选区求和');assert.equal(cell.v,'2026年11月13日');assert.equal(cell.z,'@');
    }
    assert.equal(sheet['F'+h].f,'TEXT(DATE(2026,11,13),"yyyy""年""m""月""d""日""")','公式日期表头保留动态日期计算，重复导出不重复包裹');
    assert.equal(sheet['C'+h].t,'s');assert.equal(sheet['C'+h].v,'2026年11月13日');
    assert.equal(sheet['D'+h]?.v??'','');assert.equal(sheet['E'+h].t,'n');assert.equal(sheet['E'+h].v,0);
    assert.equal(sheet['AB'+h].c[0].t,'保留表头说明');
    assert.equal(sheet['AA'+r].t,'n');assert.equal(sheet['AA'+r].v,serial,'正文原日期保留');
    assert.equal(sheet['AB'+r].f,'B'+r);assert.equal(sheet['AB'+r].v,100.25,'金额公式保留');
    assert.equal(sheet['H'+r].f,'SUM(B'+r+':B'+(r+2)+')');assert.equal(sheet['H'+r].t,'n');assert.equal(sheet['H'+r].v,75.25,'正文金额求和公式及结果不变');
    const values=Array.from({length:4},(_,i)=>sheet['B'+(h+i)]);assert.equal(values.filter(c=>c?.t==='n').reduce((sum,c)=>sum+c.v,0),75.25,'包含日期表头的选区只汇总金额');
    assert.deepEqual(values.slice(1).map(c=>[c.t,c.v]),[['n',100.25],['n',0],['n',-25]]);
    assert.equal(await xmlCell(bytes,'AC'+h),originalCustom,'未管理自定义日期表头的原生 XML 不变');
   };
   const output=await exportBytes(input,result);await verify(output);
   await verify(await exportBytes(output,result));
   assert.equal(X.read(input).Sheets.谈判付款['B'+h].t,'n','输入源文件不变');
   if(!date1904&&!headerRow)for(const type of ['shared','array']){
    const complex=Buffer.from(await page.evaluate(async({bytes,type})=>{
     const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml');
     const cell=[...doc.getElementsByTagNameNS('*','c')].find(c=>c.getAttribute('r')==='F1'),f=cell.getElementsByTagNameNS('*','f')[0];
     f.setAttribute('t',type);f.setAttribute('ref','F1:F1');if(type==='shared')f.setAttribute('si','7');
     zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
    },{bytes:[...input],type}));
    await assert.rejects(()=>exportBytes(complex,result),/日期表头 F1.*共享或数组公式/,type+' 日期公式不静默改写');
   }
  }
  const sample=await fixtures(),inputs=Object.fromEntries(['ledger','applications','negotiated','oa'].map(role=>[role,{name:role+'.xlsx',book:X.read(sample[role],{cellNF:true,cellStyles:true})}]));
  inputs.negotiated.sheetName=sample.sheetName;
  const schedule=inputs.negotiated.book.Sheets[sample.sheetName];schedule.P1.z='m/d/yy';schedule.Q1.z='yyyy"年"m"月"';
  const actualInput=Buffer.from(X.write(inputs.negotiated.book,{type:'buffer',bookType:'xlsx'})),actualResult=C.build(inputs);
  const actualOutput=await exportBytes(actualInput,actualResult),reopened=X.read(actualOutput,{cellNF:true,cellStyles:true});
  const rebuilt=C.build({...inputs,negotiated:{name:'再次上传.xlsx',book:reopened,sheetName:'谈判付款'}});
  for(const column of [15,16]){
   const address=X.utils.encode_cell({r:0,c:column});
   assert.equal(reopened.Sheets.谈判付款[address].t,'s');
   assert.equal(reopened.Sheets.谈判付款[address].v,X.SSF.format(C.DATE_FORMAT,schedule[address].v));
   assert(actualResult.negotiated.managedColumns.includes(column));assert(rebuilt.negotiated.managedColumns.includes(column),'重新导入后仍识别为付款日期列：'+address);
  }
  assert.deepEqual(rebuilt.negotiated.managedColumns,actualResult.negotiated.managedColumns,'文本日期回读不会丢失或增加业务列');
  const rebuiltOutput=X.read(await exportBytes(actualOutput,rebuilt),{cellNF:true});
  for(const address of ['P1','Q1'])assert.equal(rebuiltOutput.Sheets.谈判付款[address].v,reopened.Sheets.谈判付款[address].v);
  console.log('PASS: date headers are text and excluded from numeric sums, with WPS 31/custom date formats, DATE formulas, 1900/1904 systems, shifted header row, text/blanks/zero, monetary values/formulas/comments, custom columns and repeated exports preserved');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
