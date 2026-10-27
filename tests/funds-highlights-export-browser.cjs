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
  const result={normalHeaders,normalRows,reviewHeaders:['核对说明'],reviewRows:[],negotiated:{sheetName:'本周谈判付款',headerRow:0,headers,rows:[1,2,3].map(r=>headers.map((_,c)=>source[X.utils.encode_cell({r,c})])),updates:[],columnOrder:[0,2,1,3,4]},highlights:{negotiatedCells:['B1','B2','B2'],normalRows:[1,2,2]}};
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
   }return result;
  },{bytes:[...bytes],marker:MARKER});
  const output=await exportBytes(input,result),first=await inspect(output),saved=X.read(output,{cellNF:true}),neg=first.sheets['xl/worksheets/sheet1.xml'],normal=first.sheets['xl/worksheets/funds-normal.xml'];
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);assert.equal(neg.own.length,1);assert.equal(neg.own[0].range,'C2');assert.equal(normal.own.length,1);assert.equal(normal.own[0].range,'A2:E2');assert.deepEqual(normal.hidden,[3,4,5]);
  for(const info of [neg,normal]){assert.equal(info.own[0].priority,1);assert.equal(info.own[0].color,'FFFF0000');assert.equal(info.own[0].background,'FFFF0000','WPS conditional-format background must be explicitly red');assert.equal(info.own[0].backgroundIndex,null);assert.equal(info.own[0].stop,null);assert.deepEqual(info.own[0].dxfChildren,['fill']);assert(info.tags.indexOf('conditionalFormatting')>info.tags.indexOf('sheetData'));if(info.tags.includes('mergeCells'))assert(info.tags.indexOf('conditionalFormatting')>info.tags.indexOf('mergeCells'));for(const following of ['dataValidations','hyperlinks','pageMargins','extLst'])if(info.tags.includes(following))assert(info.tags.indexOf('conditionalFormatting')<info.tags.indexOf(following));}
  assert.equal(first.sheets['xl/worksheets/funds-review.xml'].own.length,0);assert.deepEqual(neg.rules.filter(r=>r.formula!==MARKER).map(r=>[r.formula,r.priority,r.range]),[['C2>0',2,'C2:C4'],['C2>=0',3,'C2:C4']]);
  const sourceRules=neg.rules.filter(r=>r.formula!==MARKER);assert.equal(sourceRules[0].color,'FF0000FF','Original blue rule retains its own DXF');assert.equal(sourceRules[0].background,'FF0000FF');assert.equal(sourceRules[0].bold,false);assert.deepEqual(sourceRules[0].dxfChildren,['fill']);assert.equal(sourceRules[1].bold,true,'Original bold rule retains its font DXF');assert.equal(sourceRules[1].color,undefined);assert.deepEqual(sourceRules[1].dxfChildren,['font']);
  assert.equal(neg.cells.C2.fill,'FF00FF00');assert.equal(neg.cells.C3.fill,'FFFF0000');assert.equal(neg.cells.C2.locked,'0');assert.equal(saved.Sheets.谈判付款.D2.f,'C2*3');assert.equal(saved.Sheets.谈判付款.D2.v,30);assert.equal(saved.Sheets.谈判付款.D2.z,'#,##0.00');assert.deepEqual(saved.Sheets.谈判付款['!merges'],[{s:{r:2,c:1},e:{r:3,c:1}}]);
  for(const following of ['tableStyles','colors','extLst'])if(first.styleTags.includes(following))assert(first.styleTags.indexOf('dxfs')<first.styleTags.indexOf(following));
  const next={...result,negotiated:{...result.negotiated,sheetName:'谈判付款',headers:headers.map((_,i)=>saved.Sheets.谈判付款[X.utils.encode_cell({r:0,c:i})].v),columnOrder:[0,1,2,3,4]},highlights:{negotiatedCells:['C2'],normalRows:[2]}};
  // Previous tool exports used an automatic DXF background that WPS renders white.
  // Re-upload that exact legacy shape without changing any source cell fill or unrelated rule.
  const legacy=Buffer.from(await page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=parse(await zip.file('xl/styles.xml').async('string')),dxfs=all(styles,'dxfs')[0],owned=new Set();
   for(const file of zip.file(/^xl\/worksheets\/[^/]+\.xml$/)){const doc=parse(await file.async('string'));for(const rule of all(doc,'cfRule'))if(all(rule,'formula').some(f=>f.textContent===marker))owned.add(Number(rule.getAttribute('dxfId')));}
   if(owned.size!==1)throw Error('Synthetic legacy fixture must share one tool DXF');
   for(const id of owned){const background=all(dxfs.children[id],'bgColor')[0];for(const a of [...background.attributes])background.removeAttribute(a.name);background.setAttribute('indexed','64');}
   zip.file('xl/styles.xml',s.serializeToString(styles));return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...output],marker:MARKER}));
  const old=await inspect(legacy);assert.equal(old.sheets['xl/worksheets/sheet1.xml'].own[0].backgroundIndex,'64');assert.equal(old.sheets['xl/worksheets/funds-normal.xml'].own[0].backgroundIndex,'64');
  const upgradedBytes=await exportBytes(legacy,next),upgraded=await inspect(upgradedBytes),upgradedNeg=upgraded.sheets['xl/worksheets/sheet1.xml'];
  for(const filename of ['xl/worksheets/sheet1.xml','xl/worksheets/funds-normal.xml']){assert.equal(upgraded.sheets[filename].own.length,1);assert.equal(upgraded.sheets[filename].own[0].background,'FFFF0000','Legacy tool DXF must be replaced with a WPS-compatible red background');assert.equal(upgraded.sheets[filename].own[0].backgroundIndex,null);}
  assert.deepEqual(upgradedNeg.rules,neg.rules);assert.deepEqual(upgraded.dxfPrefix,first.dxfPrefix);assert.deepEqual(upgradedNeg.cells,neg.cells,'Legacy upgrade preserves manual fills and protection');
  const upgradedBook=X.read(upgradedBytes,{cellNF:true});assert.equal(upgradedBook.Sheets.谈判付款.D2.f,saved.Sheets.谈判付款.D2.f);assert.equal(upgradedBook.Sheets.谈判付款.D2.v,saved.Sheets.谈判付款.D2.v);assert.deepEqual(upgradedBook.Sheets.谈判付款['!merges'],saved.Sheets.谈判付款['!merges']);
  const upgradedAgain=await inspect(await exportBytes(upgradedBytes,next));assert.equal(upgradedAgain.dxfCount,upgraded.dxfCount,'Repeated legacy upgrades must not accumulate DXF styles');assert.deepEqual(upgradedAgain.sheets['xl/worksheets/sheet1.xml'].rules,upgradedNeg.rules);
  const repeated=await exportBytes(output,next),again=await inspect(repeated);assert.equal(again.dxfCount,first.dxfCount);assert.deepEqual(again.dxfPrefix,first.dxfPrefix);assert.deepEqual(again.sheets['xl/worksheets/sheet1.xml'].rules,neg.rules);
  const withdrawn=await exportBytes(repeated,{...next,highlights:undefined}),clean=await inspect(withdrawn),cleanNeg=clean.sheets['xl/worksheets/sheet1.xml'];
  assert.equal(cleanNeg.own.length,0);assert.equal(clean.sheets['xl/worksheets/funds-normal.xml'].own.length,0);assert.equal(cleanNeg.cells.C2.fill,'FF00FF00');assert.equal(cleanNeg.cells.C3.fill,'FFFF0000');assert.deepEqual(cleanNeg.rules.map(r=>[r.formula,r.priority]),[['C2>0',1],['C2>=0',2]]);assert.equal(clean.dxfCount,first.dxfCount);assert.deepEqual(clean.dxfPrefix,first.dxfPrefix);
  assert.deepEqual(X.read(input).Sheets.本周谈判付款['!merges'],source['!merges']);
  console.log('PASS: WPS-compatible submitted backgrounds and legacy DXF upgrades; source fills and conditional rules survive, remapped cells and hidden normal columns are covered, headers stay clear, reruns withdraw only owned highlighting');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exit(1);});
