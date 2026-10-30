const assert=require('node:assert/strict'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const MARKER='N("GYLSJQX_FUNDS_APPLIED_V1")=0';
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const base=process.env.QA_URL||'http://127.0.0.1:8940/',page=await browser.newPage();
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('funds.html',base).href);await page.addScriptTag({path:path.resolve('funds-export.js')});
  const headers=['供应商','付款计划','备注','申请付款金额','是否提交付款申请单'],book=X.utils.book_new(),source=X.utils.aoa_to_sheet([headers,['甲',10,'保留备注',30,'已经提交部分'],['乙',20,'人工红色',40,'未提交'],['丙',30,'未提交',50,'未提交']]);
  source.B2.c=[{a:'原作者',t:'保留原人工批注'}];source.C2.c=[{a:'原作者',t:'未标红的原批注'}];
  source.D2={t:'n',v:30,f:'B2*3',z:'#,##0.00'};source['!merges']=[{s:{r:2,c:2},e:{r:3,c:2}}];X.utils.book_append_sheet(book,source,'本周谈判付款');
  const input=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],make=(doc,name,attrs={})=>{const n=doc.createElementNS(ns,name);for(const [key,value]of Object.entries(attrs))n.setAttribute(key,value);return n;};
   const styles=parse(await zip.file('xl/styles.xml').async('string')),doc=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),fills=all(styles,'fills')[0],xfs=all(styles,'cellXfs')[0],cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c]));
   const fill=(color)=>{const node=make(styles,'fill'),pattern=make(styles,'patternFill',{patternType:'solid'});pattern.append(make(styles,'fgColor',{rgb:color}),make(styles,'bgColor',{rgb:color}));node.append(pattern);return node;};
   for(const [address,color]of [['B2','FF00FF00'],['B3','FFFF0000']]){const fillId=fills.children.length;fills.append(fill(color));const xf=xfs.children[Number(cells.get(address).getAttribute('s')||0)].cloneNode(true);xf.setAttribute('fillId',fillId);xf.setAttribute('applyFill','1');xf.setAttribute('applyProtection','1');xf.append(make(styles,'protection',{locked:'0'}));cells.get(address).setAttribute('s',xfs.children.length);xfs.append(xf);}
   fills.setAttribute('count',fills.children.length);xfs.setAttribute('count',xfs.children.length);
   // SheetJS already writes an empty dxfs container; reuse it so dxfId 0/1 address the source rules.
   let dxfs=all(styles,'dxfs')[0];if(!dxfs){dxfs=make(styles,'dxfs');styles.documentElement.insertBefore(dxfs,[...styles.documentElement.children].find(n=>['tableStyles','colors','extLst'].includes(n.localName))||null);}
   if(all(styles,'dxfs').length!==1||dxfs.children.length)throw Error('Synthetic base must have one empty dxfs container');
   const blue=make(styles,'dxf'),bold=make(styles,'dxf'),font=make(styles,'font');blue.append(fill('FF0000FF'));font.append(make(styles,'b'));bold.append(font);dxfs.append(blue,bold);dxfs.setAttribute('count',dxfs.children.length);
   const cf=make(doc,'conditionalFormatting',{sqref:'B2:B4'});for(const [priority,dxfId,formula]of [['8','0','B2>0'],['17','1','B2>=0']]){const rule=make(doc,'cfRule',{type:'expression',priority,dxfId}),f=make(doc,'formula');f.textContent=formula;rule.append(f);cf.append(rule);}doc.documentElement.insertBefore(cf,[...doc.documentElement.children].find(n=>['dataValidations','hyperlinks','pageMargins','ignoredErrors','extLst'].includes(n.localName))||null);
   // Prefix-only source namespaces exercise red-dxf reuse after XMLSerializer adds local xmlns declarations.
   const prefixed=s.serializeToString(styles).replace(/(<\/?)([A-Za-z][\w.-]*)(?=[\s/>])/g,'$1s:$2').replace('xmlns="'+ns+'"','xmlns:s="'+ns+'"');
   zip.file('xl/styles.xml',prefixed);zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
  },[...X.write(book,{type:'buffer',bookType:'xlsx'})]));
  const normalHeaders=['供应商名称','申请付款金额','单据号【财务应付单】','OA合同编号','合同总金额'],normalRows=[['甲',30,'AP001','HT001',30],['乙',40,'AP002','HT002',40]].map(row=>row.map(v=>({t:typeof v==='number'?'n':'s',v})));
  const result={applicationComments:{B2:'申请日期（资金整理）：2026年10月27日'},normalHeaders,normalRows,reviewHeaders:['核对说明'],reviewRows:[],negotiated:{sheetName:'本周谈判付款',headerRow:0,headers,rows:[1,2,3].map(r=>headers.map((_,c)=>source[X.utils.encode_cell({r,c})])),updates:[],columnOrder:[0,2,1,3,4]},highlights:{negotiatedCells:['B1','B2','B2'],normalRows:[1,2,2]}};
  const exportBytes=async(bytes,data)=>Buffer.from(await page.evaluate(async({bytes,result})=>[...new Uint8Array(await(await FundsExport(new Uint8Array(bytes),result)).arrayBuffer())],{bytes:[...bytes],result:data}));
  const inspect=async bytes=>page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=parse(await zip.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],fills=all(styles,'fills')[0],dxfs=all(styles,'dxfs')[0];
   if(all(styles,'dxfs').length!==1||Number(dxfs.getAttribute('count'))!==dxfs.children.length)throw Error('Export must have one dxfs container with an accurate count');
   const result={dxfCount:dxfs?.children.length||0,dxfPrefix:[...dxfs?.children||[]].slice(0,2).map(n=>s.serializeToString(n)),styleTags:[...styles.documentElement.children].map(n=>n.localName),sheets:{}},workbook=parse(await zip.file('xl/workbook.xml').async('string')),rels=new Map(all(parse(await zip.file('xl/_rels/workbook.xml.rels').async('string')),'Relationship').map(n=>[n.getAttribute('Id'),n.getAttribute('Target')]));
   for(const [name,filename] of [['谈判付款','xl/worksheets/sheet1.xml'],['正常付款','xl/worksheets/funds-normal.xml'],['付款申请核对','xl/worksheets/funds-review.xml']]){
    const entry=all(workbook,'sheet').find(n=>n.getAttribute('name')===name),target=rels.get(entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')),actual=new URL(target,'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1);
    const doc=parse(await zip.file(actual).async('string')),rules=all(doc,'cfRule').map(n=>{const dxf=dxfs?.children[Number(n.getAttribute('dxfId'))],background=dxf&&all(dxf,'bgColor')[0];return {formula:all(n,'formula')[0]?.textContent,priority:Number(n.getAttribute('priority')),stop:n.getAttribute('stopIfTrue'),range:n.parentElement.getAttribute('sqref'),color:dxf&&all(dxf,'fgColor')[0]?.getAttribute('rgb'),background:background?.getAttribute('rgb'),backgroundIndex:background?.getAttribute('indexed'),bold:!!dxf&&all(dxf,'b').length>0,dxfChildren:dxf?[...dxf.children].map(c=>c.localName):[]};}),cells={};
    for(const c of all(doc,'c')){const xf=xfs.children[Number(c.getAttribute('s')||0)],fill=fills.children[Number(xf.getAttribute('fillId')||0)];cells[c.getAttribute('r')]={fill:all(fill,'fgColor')[0]?.getAttribute('rgb'),locked:all(xf,'protection')[0]?.getAttribute('locked')};}
    result.sheets[filename]={rules,own:rules.filter(r=>r.formula===marker),cells,tags:[...doc.documentElement.children].map(n=>n.localName),hidden:all(doc,'col').filter(c=>c.getAttribute('hidden')==='1').map(c=>Number(c.getAttribute('min')))};
   }
   result.notes=[];result.noteShapes=[];
   for(const path of Object.keys(zip.files))if(path.endsWith('.xml')&&/comments/i.test(path)){const doc=parse(await zip.file(path).async('string')),authors=all(doc,'author').map(n=>n.textContent);for(const n of all(doc,'comment'))result.notes.push({ref:n.getAttribute('ref'),author:authors[Number(n.getAttribute('authorId'))],text:all(n,'t').map(t=>t.textContent).join('')});}
   for(const path of Object.keys(zip.files))if(path.endsWith('.vml')){const doc=parse(await zip.file(path).async('string'));for(const n of all(doc,'ClientData').filter(n=>n.getAttribute('ObjectType')==='Note'))result.noteShapes.push({ref:XLSX.utils.encode_cell({r:Number(all(n,'Row')[0].textContent),c:Number(all(n,'Column')[0].textContent)}),visible:all(n,'Visible').length>0});}
   {if(result.notes.length!==result.noteShapes.length)throw Error('批注与 VML 图形数量不一致');if(new Set(result.noteShapes.map(n=>n.ref)).size!==result.noteShapes.length)throw Error('批注图形坐标重复');}
   return result;
  },{bytes:[...bytes],marker:MARKER});
  const output=await exportBytes(input,result),first=await inspect(output),saved=X.read(output,{cellNF:true}),neg=first.sheets['xl/worksheets/sheet1.xml'],normal=first.sheets['xl/worksheets/funds-normal.xml'];
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);assert.equal(neg.own.length,0);assert.equal(normal.rules.length,0);assert.deepEqual(normal.hidden,[3,4,5]);
  for(const col of ['A','B','C','D','E']){assert.equal(normal.cells[col+'2'].fill,'FFFF0000');assert.notEqual(normal.cells[col+'3'].fill,'FFFF0000');assert.notEqual(normal.cells[col+'1'].fill,'FFFF0000');}
  assert.equal(first.notes.find(n=>n.ref==='C2').author,'原作者');assert.equal(first.notes.length,2);
  assert.equal(saved.Sheets.谈判付款.C2.c[0].t,'保留原人工批注\n申请日期（资金整理）：2026年10月27日');assert.equal(saved.Sheets.谈判付款.C2.c[0].a,X.read(input).Sheets.本周谈判付款.B2.c[0].a);assert.equal(saved.Sheets.谈判付款.B2.c[0].t,'未标红的原批注');
  assert.equal(neg.cells.C2.fill,'FFFF0000');assert.equal(neg.cells.C3.fill,'FFFF0000','Unrelated manual red is retained');assert.equal(neg.cells.C2.locked,'0');
  assert.deepEqual(neg.rules.map(r=>[r.formula,r.priority,r.range]),[['C2>0',8,'C2:C4'],['C2>=0',17,'C2:C4']]);assert.equal(neg.rules[0].color,'FF0000FF');assert.equal(neg.rules[1].bold,true);
  assert.equal(saved.Sheets.谈判付款.D2.f,'C2*3');assert.equal(saved.Sheets.谈判付款.D2.v,30);assert.equal(saved.Sheets.谈判付款.D2.z,'#,##0.00');assert.deepEqual(saved.Sheets.谈判付款['!merges'],[{s:{r:2,c:1},e:{r:3,c:1}}]);
  // Reproduce WPS sorting: all row/cell styles move but conditional-format addresses do not.
  const sorted=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),doc=p.parseFromString(await zip.file('xl/worksheets/funds-normal.xml').async('string'),'application/xml'),data=doc.getElementsByTagNameNS('*','sheetData')[0],rows=[...data.children];
   for(const [i,row]of rows.slice(1).reverse().entries()){const r=i+2;row.setAttribute('r',r);for(const cell of row.children)cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,r));data.append(row);}
   zip.file('xl/worksheets/funds-normal.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
  },[...output]));
  const sortedInfo=(await inspect(sorted)).sheets['xl/worksheets/funds-normal.xml'],sortedBook=X.read(sorted,{cellStyles:true});
  assert.equal(sortedBook.Sheets.正常付款.A2.v,'乙');assert.equal(sortedBook.Sheets.正常付款.A3.v,'甲');assert.equal(sortedInfo.rules.length,0);
  for(const col of ['A','B','C','D','E']){assert.notEqual(sortedInfo.cells[col+'2'].fill,'FFFF0000');assert.equal(sortedInfo.cells[col+'3'].fill,'FFFF0000','Red follows the submitted record including hidden columns');}
  const next={...result,applicationComments:{C2:'申请日期（资金整理）：2026年10月28日；2026年10月29日'},negotiated:{...result.negotiated,sheetName:'谈判付款',headers:headers.map((_,i)=>saved.Sheets.谈判付款[X.utils.encode_cell({r:0,c:i})].v),columnOrder:[0,1,2,3,4]},highlights:{negotiatedCells:['C2'],normalRows:[2]}};
  const repeated=await exportBytes(output,next),again=await inspect(repeated);assert.deepEqual(again.sheets['xl/worksheets/sheet1.xml'].rules,neg.rules);assert.deepEqual(again.dxfPrefix,first.dxfPrefix);assert.equal(again.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');
  assert.equal(X.read(repeated).Sheets.谈判付款.C2.c[0].t,'保留原人工批注\n申请日期（资金整理）：2026年10月28日；2026年10月29日');
  const withdrawn=await exportBytes(repeated,{...next,highlights:undefined}),clean=await inspect(withdrawn),cleanNeg=clean.sheets['xl/worksheets/sheet1.xml'];
  assert(Object.values(clean.sheets['xl/worksheets/funds-normal.xml'].cells).every(c=>c.fill!=='FFFF0000'),'Re-upload withdraws obsolete normal fills');
  assert.equal(X.read(withdrawn).Sheets.谈判付款.C2.c[0].t,'保留原人工批注');assert.equal(X.read(withdrawn).Sheets.谈判付款.B2.c[0].t,'未标红的原批注');
  assert.equal(cleanNeg.cells.C2.fill,'FF00FF00','Withdrawn tool fill restores the original green');assert.equal(cleanNeg.cells.C3.fill,'FFFF0000','Manual red survives');assert.equal(cleanNeg.cells.C2.locked,'0');assert.deepEqual(cleanNeg.rules,neg.rules);
  // Native editing changes just the fill; no tool rule may force it back to red.
  const edited=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=t=>p.parseFromString(t,'application/xml'),styles=parse(await zip.file('xl/styles.xml').async('string')),fills=all(styles,'fills')[0],xfs=all(styles,'cellXfs')[0];
   const fill=styles.createElementNS(ns,'fill'),pattern=styles.createElementNS(ns,'patternFill');pattern.setAttribute('patternType','solid');for(const tag of ['fgColor','bgColor']){const c=styles.createElementNS(ns,tag);c.setAttribute('rgb','FFFFFF00');pattern.append(c);}fill.append(pattern);const fillId=fills.children.length;fills.append(fill);fills.setAttribute('count',fills.children.length);
   for(const [filename,address]of [['xl/worksheets/sheet1.xml','C2'],['xl/worksheets/funds-normal.xml','A2']]){const doc=parse(await zip.file(filename).async('string')),c=all(doc,'c').find(n=>n.getAttribute('r')===address),xf=xfs.children[Number(c.getAttribute('s'))].cloneNode(true);xf.setAttribute('fillId',fillId);c.setAttribute('s',xfs.children.length);xfs.append(xf);zip.file(filename,s.serializeToString(doc));}
   xfs.setAttribute('count',xfs.children.length);zip.file('xl/styles.xml',s.serializeToString(styles));return [...await zip.generateAsync({type:'uint8array'})];
  },[...output]));
  const manual=await inspect(edited);assert.equal(manual.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFFFF00');assert.equal(manual.sheets['xl/worksheets/funds-normal.xml'].cells.A2.fill,'FFFFFF00');
  const retained=await inspect(await exportBytes(edited,{...next,highlights:undefined}));assert.equal(retained.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFFFF00','Manual override survives re-upload without a new matching application');
  // Upgrade a prior export carrying the legacy owned rule; retain all unrelated rules.
  const legacy=Buffer.from(await page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),cf=doc.createElementNS(ns,'conditionalFormatting'),rule=doc.createElementNS(ns,'cfRule'),f=doc.createElementNS(ns,'formula');cf.setAttribute('sqref','C2');rule.setAttribute('type','expression');rule.setAttribute('dxfId','0');rule.setAttribute('priority','1');f.textContent=marker;rule.append(f);cf.append(rule);doc.documentElement.append(cf);zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...output],marker:MARKER}));
  const upgraded=await inspect(await exportBytes(legacy,next));assert.equal(upgraded.sheets['xl/worksheets/sheet1.xml'].own.length,0);assert.deepEqual(upgraded.sheets['xl/worksheets/sheet1.xml'].rules,neg.rules);assert.equal(upgraded.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');
  assert.deepEqual(X.read(input).Sheets.本周谈判付款['!merges'],source['!merges']);
  delete source.B2.c;delete source.C2.c;
  const fresh=await exportBytes(X.write(book,{type:'buffer',bookType:'xlsx'}),result);assert.equal(X.read(fresh).Sheets.谈判付款.C2.c[0].t,'申请日期（资金整理）：2026年10月27日');
  assert.deepEqual((await inspect(fresh)).noteShapes,[{ref:'C2',visible:false}]);
  const removed=await exportBytes(fresh,{...next,highlights:undefined});assert.equal(X.read(removed).Sheets.谈判付款.C2.c,undefined);assert.equal((await inspect(removed)).notes.length,0);
  book.Workbook={WBProps:{date1904:true}};const epoch=await exportBytes(X.write(book,{type:'buffer',bookType:'xlsx'}),{...result,normalHeaders:['申请日期'],normalRows:[[{t:'n',v:46322,z:'yyyy"年"m"月"d"日"'}]]});const epochBook=X.read(epoch,{cellNF:true});assert.equal(epochBook.Workbook.WBProps.date1904,true);assert.equal(X.SSF.format(epochBook.Sheets.正常付款.A2.z,epochBook.Sheets.正常付款.A2.v,{date1904:true}),'2026年10月27日');
  console.log('PASS: application-date notes, native VML anchors and hidden display, column reordering, manual-note preservation, repeat/withdrawal, date1904; editable native fills, full-row sorting including hidden cells, legacy rule removal, repeat/withdrawal restores original fills and preserves manual overrides, formulas and unrelated conditional formatting');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exit(1);});
