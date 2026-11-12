const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {REVIEW_HEADERS}=require('./funds-fixtures.cjs');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage();await page.goto(new URL('funds.html',process.env.QA_URL||'http://127.0.0.1:8940/').href);
  await page.addScriptTag({path:require('path').resolve('funds-export.js')});
  const book=X.utils.book_new();const source=X.utils.aoa_to_sheet([['供应商','申请付款金额','已付金额','剩余付款金额','付款计划'],['测试甲',100,0,100,'保持备注'],['测试乙',200,20,180,'合并备注'],['测试丙',300,30,270],['合计',600,50,550]]);
  source.D2={t:'n',v:100,f:'B2-C2'};source.E2={t:'n',v:100,f:'D2'};source.D5={t:'n',v:550,f:'SUM(D2:D4)'};source['!merges']=[{s:{r:2,c:4},e:{r:3,c:4}}];source['!cols']=[{wch:30},{wch:20},{wch:20},{wch:20},{wch:35}];source['!rows']=[{hpt:33},{hpt:40,hidden:true},{hpt:27}];source['!autofilter']={ref:'A1:E5'};
  X.utils.book_append_sheet(book,source,'未来10月谈判付款');X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['必须保留'],['secret']]),'其他');
  const input=Buffer.from(await page.evaluate(async bytes=>{
    const z=await JSZip.loadAsync(new Uint8Array(bytes)),doc=new DOMParser().parseFromString(await z.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),filter=doc.getElementsByTagNameNS('*','autoFilter')[0],column=doc.createElementNS(filter.namespaceURI,'filterColumn'),filters=doc.createElementNS(filter.namespaceURI,'filters'),value=doc.createElementNS(filter.namespaceURI,'filter');column.setAttribute('colId','0');value.setAttribute('val','旧筛选');filters.append(value);column.append(filters);filter.append(column);z.file('xl/worksheets/sheet1.xml',new XMLSerializer().serializeToString(doc));return [...await z.generateAsync({type:'uint8array'})];
  },[...X.write(book,{type:'buffer',bookType:'xlsx'})]));
  const result={normalHeaders:['登记月份','供应商名称','申请付款金额','已付金额','剩余金额'],normalRows:[[{t:'n',v:46319,z:'yyyy年m月'},{t:'s',v:'测试正常公司'},{t:'n',v:12.34,z:'#,##0.00'},{t:'n',v:2.34,z:'#,##0.00'},{t:'n',v:10,f:'C2-D2',z:'#,##0.00'}]],reviewHeaders:REVIEW_HEADERS,reviewRows:[['正常付款','测试正常公司','迈德斯特',12.34,'未申请','供应商名称和申请付款金额未匹配到付款申请','资金台账.xlsx','对账明细-李奇',2,'李奇'].map(v=>({t:typeof v==='number'?'n':'s',v,...(v===12.34?{z:'#,##0.00'}:{})}))],negotiated:{sheetName:'未来10月谈判付款',headerRow:0,updates:[{address:'C2',cell:{t:'n',v:40}},{address:'D2',cell:{t:'n',v:60,f:'B2-C2'}},{address:'F1',cell:{t:'s',v:'提交付款申请单日期'}},{address:'F2',cell:{t:'s',v:'2026-10-01\n2026-10-24'}}]}};
  async function exportBytes(bytes,data){return Buffer.from(await page.evaluate(async({bytes,result})=>[...new Uint8Array(await (await FundsExport(new Uint8Array(bytes),result)).arrayBuffer())],{bytes:[...bytes],result:data}));}
  const matrixRows=[['甲',0,'','',false],['','','',' ',''],['乙',42,'','','']].map(row=>row.map(v=>({t:typeof v==='number'?'n':typeof v==='boolean'?'b':'s',v})));
  const matrixBytes=await exportBytes(input,{...result,normalHeaders:['供应商','金额','备注','','布尔值'],normalRows:matrixRows});
  const matrixBorders=await page.evaluate(async bytes=>{
    const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),parse=x=>new DOMParser().parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],doc=parse(await zip.file('xl/worksheets/funds-normal.xml').async('string')),styles=parse(await zip.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0].children,borders=all(styles,'borders')[0].children;
    return all(doc,'c').filter(c=>!/[A-Z]+1$/.test(c.getAttribute('r'))).map(c=>{const border=borders[Number(xfs[Number(c.getAttribute('s')||0)].getAttribute('borderId'))];return {ref:c.getAttribute('r'),bordered:['left','right','top','bottom'].every(edge=>!!all(border,edge)[0]?.getAttribute('style'))};});
  },[...matrixBytes]);
  for(const cell of matrixBorders)assert.equal(cell.bordered,!cell.ref.startsWith('D')&&!cell.ref.endsWith('3'),cell.ref+'：有内容的行列交叉区保留边框，整空行列无边框');
  const matrixBook=X.read(matrixBytes);assert.equal(matrixBook.Sheets.正常付款.B2.v,0);assert.equal(matrixBook.Sheets.正常付款.E2.v,false);assert.equal(matrixBook.Sheets.正常付款.C2.v,'');
  const output=await exportBytes(input,result),saved=X.read(output,{cellStyles:true,cellNF:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对','付款进度整理','其他']);const s=saved.Sheets.谈判付款;
  for(const sheet of saved.SheetNames.slice(0,4).map(name=>saved.Sheets[name])){assert.equal(sheet['!autofilter'].ref,sheet['!ref']);for(const row of sheet['!rows']||[])if(row){if(sheet===saved.Sheets.付款申请核对)assert(row.hpt>=30);else assert.equal(row.hpt,30);assert(!row.hidden);}assert(sheet['!cols'].every(c=>c.width>=12&&c.width<=80));}
  assert.equal(s.C2.v,40);assert.equal(s.D2.v,60);assert.equal(s.E2.v,60);assert.equal(s.D5.v,510);assert.equal(s.D5.f,'SUM(D2:D4)');assert.deepEqual(s['!merges'],source['!merges']);assert.equal(s['!rows'][1].hpt,30);assert.equal(s.E3.v,'合并备注');assert.equal(saved.Sheets.正常付款.A2.t,'s');assert.equal(saved.Sheets.正常付款.A2.v,'2026年10月');assert.equal(saved.Sheets.正常付款.A2.z,'@');assert.equal(saved.Sheets.正常付款.E2.v,10);
  assert.equal(saved.Sheets.正常付款['!autofilter'].ref,'A1:E2');assert.equal(saved.Sheets.正常付款.A1.s.fgColor.rgb,'315B9D');
  for(const sheet of saved.SheetNames.slice(0,4).map(name=>saved.Sheets[name]))for(const address of Object.keys(sheet).filter(a=>/^[A-Z]+1$/.test(a)))assert.equal(sheet[address].s.fgColor.rgb,'315B9D','所有原有及新增表头统一底色');
  const review=saved.Sheets.付款申请核对;assert.deepEqual(X.utils.sheet_to_json(review,{header:1})[0],REVIEW_HEADERS);assert.equal(review.D2.v,12.34);assert.equal(review.D2.t,'n');assert.equal(review.D2.z,'#,##0.00');assert.equal(review.I2.v,2);assert.equal(review.G2.v,'资金台账.xlsx');assert.equal(review.J2.v,'李奇');assert.equal(review['!autofilter'].ref,'A1:J2');assert.deepEqual(review.A1.s,saved.Sheets.正常付款.A1.s);assert.deepEqual(review.A2.s,saved.Sheets.正常付款.B2.s);
  const parts=await page.evaluate(async bytes=>{const z=await JSZip.loadAsync(new Uint8Array(bytes));return {files:Object.keys(z.files),nativeXml:await z.file('xl/worksheets/sheet1.xml').async('string'),xml:await z.file('xl/worksheets/funds-normal.xml').async('string'),reviewXml:await z.file('xl/worksheets/funds-review.xml').async('string'),app:await z.file('docProps/app.xml').async('string')};},[...output]);assert(parts.files.includes('xl/worksheets/sheet2.xml'));assert.doesNotMatch(parts.nativeXml,/<(?:\w+:)?filterColumn\b/);assert.match(parts.nativeXml,/defaultRowHeight="30"/);assert(!parts.files.some(p=>p.startsWith('xl/tables/')));assert.match(parts.xml,/state="frozen"/);assert.match(parts.reviewXml,/state="frozen"/);assert.match(parts.app,/付款申请核对/);assert.match(parts.app,/<(?:\w+:)?i4>5<\//);
  const again=await exportBytes(output,{...result,reviewRows:[],negotiated:{...result.negotiated,sheetName:'谈判付款'}}),againBook=X.read(again);assert.deepEqual(againBook.SheetNames,['谈判付款','正常付款','付款申请核对','付款进度整理','其他']);assert.deepEqual(X.utils.sheet_to_json(againBook.Sheets.付款申请核对,{header:1}),[REVIEW_HEADERS]);assert.equal(againBook.Sheets.付款申请核对['!autofilter'].ref,'A1:J1');
  const submissionHeaders=['供应商','申请付款金额','已付金额','剩余付款金额','付款计划','提交付款申请单日期','提交时间汇总'];
  const withSummary=await exportBytes(output,{...result,negotiated:{sheetName:'谈判付款',headerRow:0,headers:submissionHeaders,updates:[{address:'G1',cell:{t:'s',v:'提交时间汇总'}},{address:'G2',cell:{t:'s',v:'2026-10-01-10.00'}}]}});
  // Simulate a reused workbook with its own unwrapped style and manually increased row height.
  const reusedInput=Buffer.from(await page.evaluate(async bytes=>{
    const z=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),doc=p.parseFromString(await z.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),styles=p.parseFromString(await z.file('xl/styles.xml').async('string'),'application/xml');
    const all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c])),xfs=all(styles,'cellXfs')[0];
    const xf=xfs.children[Number(cells.get('G2').getAttribute('s'))].cloneNode(true),alignment=all(xf,'alignment')[0];alignment.setAttribute('wrapText','0');alignment.setAttribute('horizontal','left');alignment.setAttribute('vertical','top');alignment.setAttribute('indent','2');xf.setAttribute('applyProtection','1');
    const protection=styles.createElementNS(xf.namespaceURI,'protection');protection.setAttribute('locked','0');xf.append(protection);const styleId=xfs.children.length;xfs.append(xf);xfs.setAttribute('count',xfs.children.length);
    for(const address of ['F2','G2','E3'])cells.get(address).setAttribute('s',styleId);
    const rows=new Map(all(doc,'row').map(r=>[r.getAttribute('r'),r]));rows.get('2').setAttribute('ht','40');rows.get('3').setAttribute('ht','180');rows.get('3').setAttribute('customHeight','1');
    z.file('xl/styles.xml',s.serializeToString(styles));z.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await z.generateAsync({type:'uint8array'})];
  },[...withSummary]));
  const batchSummary='2026-10-01-10.00\n2026-10-13-20.00\n2026-10-24-30.00',batchDates='2026-10-01\n2026-10-13\n2026-10-24';
  const reused=await exportBytes(reusedInput,{...result,negotiated:{sheetName:'谈判付款',headerRow:0,headers:submissionHeaders,updates:['F2','F3','G2','G3'].map(address=>({address,cell:{t:'s',v:address.startsWith('F')?batchDates:batchSummary}}))}}),reusedBook=X.read(reused,{cellStyles:true,cellNF:true}),reusedSheet=reusedBook.Sheets.谈判付款;
  assert.equal(reusedSheet.G2.v,batchSummary);assert.equal(reusedSheet.G3.v,batchSummary);assert.equal(reusedSheet.F2.v,batchDates);assert.equal(reusedSheet.F3.v,batchDates);assert.equal(reusedSheet['!rows'][1].hpt,30);assert.equal(reusedSheet['!rows'][2].hpt,30);assert.deepEqual(reusedSheet['!merges'],source['!merges']);assert.equal(reusedSheet.E3.v,'合并备注');assert.equal(reusedSheet.D2.f,'B2-C2');assert.equal(reusedSheet.D2.v,60);assert.equal(reusedSheet.D5.v,510);
  const stylePreservation=await page.evaluate(async({before,after})=>{
    const p=new DOMParser(),s=new XMLSerializer(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
    async function inspect(bytes){const z=await JSZip.loadAsync(new Uint8Array(bytes)),doc=p.parseFromString(await z.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),styles=p.parseFromString(await z.file('xl/styles.xml').async('string'),'application/xml');return {doc,xfs:all(styles,'cellXfs')[0],cells:new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c]))};}
    const a=await inspect(before),b=await inspect(after),id=Number(a.cells.get('G2').getAttribute('s')),clean=xf=>{const copy=xf.cloneNode(true);copy.removeAttribute('applyAlignment');all(copy,'alignment')[0].removeAttribute('wrapText');return s.serializeToString(copy);};
    return {sourceStyle:s.serializeToString(a.xfs.children[id]),unchangedSourceStyle:s.serializeToString(b.xfs.children[id]),sourceDetails:clean(a.xfs.children[id]),written:['F2','G2'].map(address=>{const index=Number(b.cells.get(address).getAttribute('s')),xf=b.xfs.children[index];return {index,wrap:all(xf,'alignment')[0].getAttribute('wrapText'),details:clean(xf)};}),sourceId:id,unrelatedId:Number(b.cells.get('E3').getAttribute('s')),columnsBefore:s.serializeToString(all(a.doc,'cols')[0]),columnsAfter:s.serializeToString(all(b.doc,'cols')[0])};
  },{before:[...reusedInput],after:[...reused]});
  assert.equal(stylePreservation.sourceStyle,stylePreservation.unchangedSourceStyle);assert.equal(stylePreservation.unrelatedId,stylePreservation.sourceId);for(const written of stylePreservation.written){assert.notEqual(written.index,stylePreservation.sourceId);assert.equal(written.wrap,'1');assert.equal(written.details,stylePreservation.sourceDetails);}
  const b1904=X.utils.book_new();b1904.Workbook={WBProps:{date1904:true}};X.utils.book_append_sheet(b1904,source,'未来10月谈判付款');const out1904=await exportBytes(Buffer.from(X.write(b1904,{type:'buffer',bookType:'xlsx'})),result);assert.equal(X.read(out1904).Sheets.正常付款.A2.t,'s');assert.equal(X.read(out1904).Sheets.正常付款.A2.v,'2026年10月');
  const dateBook=X.utils.book_new(),dateSheet=X.utils.aoa_to_sheet([['供应商','原日期','更新日期'],['甲',46319,46319]]),originalSheet=X.utils.aoa_to_sheet([['原样保留'],[46319]]);
  dateSheet.B2.z=dateSheet.C2.z='yyyy"年"m"月"d"日"';originalSheet.A2={t:'n',v:46319,z:'yyyy-mm-dd',f:'DATE(2026,10,24)'};
  X.utils.book_append_sheet(dateBook,dateSheet,'谈判付款');X.utils.book_append_sheet(dateBook,originalSheet,'原工作表');
  const dateInput=Buffer.from(X.write(dateBook,{type:'buffer',bookType:'xlsx'})),dateOutput=await exportBytes(dateInput,{...result,negotiated:{sheetName:'谈判付款',headerRow:0,updates:[{address:'C2',cell:{t:'n',v:46322,z:'yyyy"年"m"月"d"日"'}}]}}),dateSaved=X.read(dateOutput,{cellNF:true});
  assert.equal(dateSaved.Sheets.谈判付款.B2.t,'n');assert.equal(dateSaved.Sheets.谈判付款.B2.v,46319);assert.equal(dateSaved.Sheets.谈判付款.C2.t,'s');assert.equal(dateSaved.Sheets.谈判付款.C2.v,'2026年10月27日');assert.equal(dateSaved.Sheets.谈判付款.C2.z,'@');
  const dateParts=await page.evaluate(async({a,b})=>{const before=await JSZip.loadAsync(Uint8Array.from(a)),after=await JSZip.loadAsync(Uint8Array.from(b));return [await before.file('xl/worksheets/sheet2.xml').async('string'),await after.file('xl/worksheets/sheet2.xml').async('string')];},{a:[...dateInput],b:[...dateOutput]});assert.equal(dateParts[0],dateParts[1],'原工作表日期及公式原样保留');
  // Reused source files may have unrelated columns between the two submission fields.
  const moveHeaders=['供应商','提交付款申请单日期','付款计划','提交时间汇总','原公式备注'];
  const moveBook=X.utils.book_new(),moveSheet=X.utils.aoa_to_sheet([moveHeaders,['测试甲','2026-10-01',20,'2026-10-01-20.00',40],['合并备注','2026-10-02',30,'2026-10-02-30.00',0],['','2026-10-03',40,'2026-10-03-40.00',70]]);
  moveSheet.E2={t:'n',v:40,f:'C2*2'};moveSheet.E3={t:'n',v:0,f:'IF(B2="D2",$C$2,0)'};moveSheet.E4={t:'n',v:70,f:'SUM(C3:C4)'};
  moveSheet.C2.c=[{a:'原作者',t:'跟随付款计划的批注'}];moveSheet.C2.l={Target:'#E2'};
  moveSheet['!merges']=[{s:{r:2,c:0},e:{r:3,c:0}}];moveSheet['!cols']=[{wch:20},{wch:25},{wch:12},{wch:38},{wch:32}];moveSheet['!autofilter']={ref:'A1:E4'};
  X.utils.book_append_sheet(moveBook,moveSheet,'复用谈判付款');moveBook.Workbook={Names:[{Name:'付款计划范围',Ref:"'复用谈判付款'!$C:$C"}]};
  const moveInput=Buffer.from(X.write(moveBook,{type:'buffer',bookType:'xlsx'})),moveResult={...result,reviewRows:[],negotiated:{sheetName:'复用谈判付款',headerRow:0,headers:moveHeaders,rows:X.utils.sheet_to_json(moveSheet,{header:1}).slice(1).map((row,r)=>row.map((_,c)=>moveSheet[X.utils.encode_cell({r:r+1,c})])),updates:[],columnOrder:[0,1,3,2,4]}};
  const moved=await exportBytes(moveInput,moveResult),movedBook=X.read(moved,{cellStyles:true,cellNF:true}),m=movedBook.Sheets.谈判付款;
  assert.deepEqual(X.utils.sheet_to_json(m,{header:1})[0],['供应商','提交付款申请单日期','提交时间汇总','付款计划','原公式备注']);
  assert.equal(m.C2.v,'2026-10-01-20.00');assert.equal(m.D2.v,20);assert.equal(m.E2.f,'D2*2');assert.equal(m.E3.f,'IF(B2="D2",$D$2,0)');assert.equal(m.E4.f,'SUM(D3:D4)');assert.equal(m.E4.v,70);
  assert.equal(m.D2.c[0].t,'跟随付款计划的批注');assert.equal(m.D2.l.Target,'#E2');assert(m['!cols'][2].width>m['!cols'][3].width);assert.deepEqual(m['!merges'],moveSheet['!merges']);assert.equal(m['!autofilter'].ref,'A1:E4');assert.equal(movedBook.Workbook.Names.find(n=>n.Name==='付款计划范围').Ref,"'谈判付款'!$D:$D");
  const sourceAgain=X.read(moveInput,{cellStyles:true,cellNF:true}).Sheets.复用谈判付款;assert.equal(sourceAgain.E2.f,'C2*2');assert.equal(sourceAgain.C2.c[0].t,'跟随付款计划的批注');
  const repeatResult={...moveResult,negotiated:{...moveResult.negotiated,sheetName:'谈判付款',headers:X.utils.sheet_to_json(m,{header:1})[0],columnOrder:[0,1,2,3,4]}};
  const repeat=X.read(await exportBytes(moved,repeatResult),{cellStyles:true,cellNF:true});assert.equal(repeat.Sheets.谈判付款.E2.f,'D2*2');assert.equal(repeat.Sheets.谈判付款.D2.c[0].t,'跟随付款计划的批注');
  const flowIds=Array.from({length:5},(_,i)=>'MDST-DGFK2026102400'+i),longFlow=flowIds.join('\n'),flowResult={...result,normalHeaders:['流程编号','金额'],normalRows:[[{t:'s',v:longFlow},{t:'n',v:0}]],negotiated:{...result.negotiated,updates:[...result.negotiated.updates,{address:'G1',cell:{t:'s',v:'流程编号'}},{address:'G2',cell:{t:'s',v:longFlow}},{address:'G3',cell:{t:'s',v:'LONG-'+ 'X'.repeat(300)}}]}};
  const flowOutput=await exportBytes(input,flowResult),flowBook=X.read(flowOutput,{cellStyles:true,cellNF:true});
  assert.equal(flowBook.Sheets.谈判付款.G2.v,longFlow);assert.equal(flowBook.Sheets.正常付款.A2.v,flowIds.join('；'));assert.equal(flowBook.Sheets.谈判付款.G3.v,'LONG-'+ 'X'.repeat(300));assert.equal(flowBook.Sheets.谈判付款['!cols'][6].width,255);assert(flowBook.Sheets.正常付款['!cols'][0].width>80);assert.equal(flowBook.Sheets.正常付款.B2.v,0);assert.deepEqual(flowBook.Sheets.谈判付款['!merges'],source['!merges']);
  const untouchedFlow=await exportBytes(flowOutput,{...flowResult,negotiated:{...flowResult.negotiated,sheetName:'谈判付款',updates:[]}}),untouchedSheet=X.read(untouchedFlow,{cellNF:true}).Sheets.谈判付款;
  for(const address of ['A2','B2','C2','D2','E2','E3','G2','G3'])for(const key of ['t','v','f'])assert.equal(untouchedSheet[address]?.[key],flowBook.Sheets.谈判付款[address]?.[key],'未更新单元格保留 '+address+' '+key);
  const flowLayout=await page.evaluate(async bytes=>{
    const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=p.parseFromString(await zip.file('xl/styles.xml').async('string'),'application/xml'),xfs=all(styles,'cellXfs')[0],borders=all(styles,'borders')[0],problems=[];
    for(const path of ['xl/worksheets/sheet1.xml','xl/worksheets/funds-normal.xml','xl/worksheets/funds-review.xml']){const doc=p.parseFromString(await zip.file(path).async('string'),'application/xml');for(const cell of all(doc,'c')){const xf=xfs.children[Number(cell.getAttribute('s')||0)],border=borders.children[Number(xf.getAttribute('borderId')||0)],value=all(cell,'v')[0]?.textContent??all(cell,'t').map(t=>t.textContent).join('');if(value!==''||all(cell,'f').length)for(const edge of ['left','right','top','bottom'])if(all(border,edge)[0]?.getAttribute('style')!==(/[A-Z]+1$/.test(cell.getAttribute('r'))?'thin':'hair')||all(all(border,edge)[0],'color')[0]?.getAttribute('rgb')!=='FF404040')problems.push(path+':'+cell.getAttribute('r')+':'+edge);if((path.endsWith('sheet1.xml')&&/^G[1-3]$/.test(cell.getAttribute('r')))||(path.endsWith('funds-normal.xml')&&/^A[12]$/.test(cell.getAttribute('r')))){const alignment=all(xf,'alignment')[0];if(alignment?.getAttribute('wrapText')!=='0'||alignment?.getAttribute('shrinkToFit')!=='1')problems.push('流程编号样式错误');}if(path.endsWith('sheet1.xml')&&cell.getAttribute('r')==='E4')for(const edge of ['left','right','bottom'])if(!all(border,edge)[0]?.getAttribute('style'))problems.push('合并区域缺少外围边框');}}
    return problems;
  },[...flowOutput]);assert.deepEqual(flowLayout,[]);
  const blankBorders=await page.evaluate(async bytes=>{const zip=await JSZip.loadAsync(new Uint8Array(bytes)),parse=x=>new DOMParser().parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],doc=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),styles=parse(await zip.file('xl/styles.xml').async('string')),cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c])),xfs=all(styles,'cellXfs')[0],borders=all(styles,'borders')[0];return ['E5','F3','F4','F5','G4','G5'].map(ref=>{const c=cells.get(ref),border=c&&borders.children[Number(xfs.children[Number(c.getAttribute('s')||0)].getAttribute('borderId')||0)];return {ref,empty:c&&!all(c,'v').length&&!all(c,'t').length,bordered:!!border&&['left','right','top','bottom'].every(edge=>!!all(border,edge)[0]?.getAttribute('style'))};});},[...flowOutput]);assert(blankBorders.every(c=>c.empty&&c.bordered),JSON.stringify(blankBorders));
  const longHeader='完整显示且不得换行的自适应表头'.repeat(4),wideResult={...result,progressHeaders:[longHeader,'短标题'],progressRows:[],normalHeaders:[longHeader,'短标题'],normalRows:[[{t:'s',v:'短内容'},{t:'s',v:'短'}]],reviewHeaders:[longHeader,'短标题'],reviewRows:[],negotiated:{...result.negotiated,updates:[...result.negotiated.updates,{address:'A1',cell:{t:'s',v:longHeader}}]}};
  const wideOutput=await exportBytes(input,wideResult),wideBook=X.read(wideOutput,{cellStyles:true});
  for(const sheet of wideBook.SheetNames.slice(0,4).map(name=>wideBook.Sheets[name])){assert(sheet['!cols'][0].width>80);assert(sheet['!cols'][0].width>sheet['!cols'][1].width);assert.equal(sheet.A1.v,longHeader);}
  const headerProblems=await page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(new Uint8Array(bytes)),parse=x=>new DOMParser().parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],problems=[];
   for(const path of ['xl/worksheets/sheet1.xml','xl/worksheets/funds-normal.xml','xl/worksheets/funds-review.xml'])for(const c of all(parse(await z.file(path).async('string')),'c').filter(c=>/^[A-Z]+1$/.test(c.getAttribute('r')))){const xf=xfs.children[Number(c.getAttribute('s')||0)];if(all(xf,'alignment')[0]?.getAttribute('wrapText')!=='0')problems.push(path+':'+c.getAttribute('r'));}
   return problems;
  },[...wideOutput]);assert.deepEqual(headerProblems,[]);
  source.E4={f:"'其他'!C2",t:'n',v:11};const linked=X.utils.book_new();X.utils.book_append_sheet(linked,source,'未来10月谈判付款');X.utils.book_append_sheet(linked,X.utils.aoa_to_sheet([[1],[null,null,11]]),'其他');const linkedOutput=X.read(await exportBytes(Buffer.from(X.write(linked,{type:'buffer',bookType:'xlsx'})),result));assert.equal(linkedOutput.Sheets.谈判付款.E4.f,"'其他'!C2");assert.equal(linkedOutput.Sheets.其他.C2.v,11);assert.deepEqual(await page.evaluate(()=>FundsExport.lastWarnings),[],"其他工作表同地址单元格未变，不应误报需要重算");
  await fs.writeFile('/private/tmp/funds-export-synthetic.xlsx',output);
  // Preserve the complete original package, including a hidden sheet and a second negotiation version.
  const preserved=X.utils.book_new(),aux=X.utils.aoa_to_sheet([['人工辅助表',10],['保留原备注','合并内容',''],[10]]);
  aux.A2.c=[{a:'人工作者',t:'原批注不能丢失'}];aux.A3.f='B1+0';aux['!cols']=[{wch:17},{wch:31,hidden:true},{wch:12}];aux['!rows']=[{hpt:42},{hpt:51,hidden:true}];aux['!merges']=[{s:{r:1,c:1},e:{r:1,c:2}}];aux['!autofilter']={ref:'A1:C3'};
  X.utils.book_append_sheet(preserved,aux,'辅助说明');
  X.utils.book_append_sheet(preserved,X.utils.aoa_to_sheet([['保留另一谈判版本'],[99]]),'谈判付款');
  X.utils.book_append_sheet(preserved,moveSheet,'复用谈判付款');
  const refs=X.utils.aoa_to_sheet([['跨表引用'],[20,99,3,12]]);
  refs.A2.f="'复用谈判付款'!C2+C2";refs.B2.f="'谈判付款'!A2";refs.D2.f="'正常付款'!C2";
  refs.A3={t:'s',v:'复用谈判付款!C2',f:'"复用谈判付款!C2"'};
  refs.A4={t:'n',v:20,f:"'[外部.xlsx]复用谈判付款'!C2"};refs.A5={t:'n',v:20,f:'[外部.xlsx]复用谈判付款!C2'};refs['!ref']='A1:D5';refs.A1.l={Target:"#'复用谈判付款'!C2"};
  X.utils.book_append_sheet(preserved,refs,'引用核对');
  for(const name of ['正常付款','付款申请核对','付款进度整理'])X.utils.book_append_sheet(preserved,X.utils.aoa_to_sheet([['旧结果不可重复追加'],[999]]),name);
  preserved.Workbook={Sheets:preserved.SheetNames.map((_,i)=>({Hidden:i===0?1:0})),Names:[{Name:'辅助局部',Sheet:0,Ref:'$C$2'},{Name:'谈判局部',Sheet:2,Ref:'$C$2'},{Name:'跨表范围',Ref:"'复用谈判付款'!$C$2"}]};
  const preservedInput=Buffer.from(X.write(preserved,{type:'buffer',bookType:'xlsx'}));
  const preservedOutput=await exportBytes(preservedInput,moveResult),reopened=X.read(preservedOutput,{cellStyles:true,cellNF:true});
  const preservedNames=['谈判付款','正常付款','付款申请核对','付款进度整理','辅助说明','谈判付款（原表）','引用核对'];
  assert.deepEqual(reopened.SheetNames,preservedNames);assert.equal(reopened.Sheets.谈判付款.D2.v,20);
  assert.equal(reopened.Sheets['谈判付款（原表）'].A2.v,99);assert.equal(reopened.Sheets.正常付款.C2.v,12.34);
  assert.equal(reopened.Workbook.Sheets[4].Hidden,1);assert.equal(reopened.Sheets.辅助说明.A2.c[0].t,'原批注不能丢失');
  assert.equal(reopened.Sheets.辅助说明['!rows'][1].hpt,51);assert(reopened.Sheets.辅助说明['!rows'][1].hidden);assert(reopened.Sheets.辅助说明['!cols'][1].hidden);
  assert.deepEqual(reopened.Sheets.辅助说明['!merges'],aux['!merges']);assert.equal(reopened.Sheets.辅助说明.A3.f,'B1+0');
  assert.equal(reopened.Sheets.引用核对.A2.f,"'谈判付款'!D2+C2");assert.equal(reopened.Sheets.引用核对.B2.f,"'谈判付款（原表）'!A2");assert.equal(reopened.Sheets.引用核对.D2.f,"'正常付款'!C2");
  for(const address of ['A3','A4','A5'])assert.equal(reopened.Sheets.引用核对[address].f,refs[address].f);
  assert.equal(reopened.Sheets.引用核对.A1.l.Target,"#'谈判付款'!D2");
  const names=reopened.Workbook.Names;assert.equal(names.find(n=>n.Name==='辅助局部').Sheet,4);assert.equal(names.find(n=>n.Name==='辅助局部').Ref,'$C$2');assert.equal(names.find(n=>n.Name==='谈判局部').Sheet,0);assert.equal(names.find(n=>n.Name==='谈判局部').Ref,'$D$2');assert.equal(names.find(n=>n.Name==='跨表范围').Ref,"'谈判付款'!$D$2");
  const unchanged=await page.evaluate(async({before,after})=>{
    const a=await JSZip.loadAsync(new Uint8Array(before)),b=await JSZip.loadAsync(new Uint8Array(after));
    const paths=Object.keys(a.files).filter(p=>/^xl\/(?:worksheets\/(?:sheet[12]\.xml|_rels\/sheet1\.xml\.rels)|comments1\.xml|drawings\/vmlDrawing1\.vml)/.test(p)&&!a.files[p].dir);
    return Promise.all(paths.map(async p=>({path:p,equal:!!b.file(p)&&await a.file(p).async('base64')===await b.file(p).async('base64')})));
  },{before:[...preservedInput],after:[...preservedOutput]});assert(unchanged.length>=4);assert(unchanged.every(p=>p.equal),JSON.stringify(unchanged));
  const reusedPreserved=await exportBytes(preservedOutput,repeatResult),reusedPreservedBook=X.read(reusedPreserved);assert.deepEqual(reusedPreservedBook.SheetNames,preservedNames);assert.equal(reusedPreservedBook.Sheets.引用核对.A2.f,"'谈判付款'!D2+C2");
  refs.A6={t:'n',v:20,f:"SUM('辅助说明:复用谈判付款'!C2)"};refs['!ref']='A1:D6';
  await assert.rejects(()=>exportBytes(Buffer.from(X.write(preserved,{type:'buffer',bookType:'xlsx'})),moveResult),/三维公式引用/,'不静默改变跨多个工作表的求和范围');
  console.log('PASS: all original sheets and native parts retained, result tabs replaced once, repeated export, hidden states, comments, merges, local names, cross-sheet formulas and hyperlinks, selected-name collision, source unchanged');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
