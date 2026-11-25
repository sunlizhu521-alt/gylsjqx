const assert=require('node:assert/strict'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const MARKER='N("GYLSJQX_FUNDS_APPLIED_V1")=0';
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const base=process.env.QA_URL||'http://127.0.0.1:8940/',page=await browser.newPage();
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('funds.html',base).href);if(!process.env.QA_REMOTE_ASSETS)await page.addScriptTag({path:path.resolve('funds-export.js')});
  const headers=['供应商','付款计划','备注','申请付款金额','是否提交付款申请单'],book=X.utils.book_new(),source=X.utils.aoa_to_sheet([headers,['甲',10,'保留备注',30,'已经提交部分'],['乙',20,'人工红色',40,'未提交'],['丙',30,'未提交',50,'未提交']]);
  source.B2.c=[{a:'原作者',t:'保留原人工批注\n申请日期（资金整理）：2026年10月1日'}];source.C2.c=[{a:'原作者',t:'未标红的原批注'}];
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
   zip.file('xl/styles.xml',prefixed);zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));
   for(const path of Object.keys(zip.files))if(path.endsWith('.vml')){const vml=parse(await zip.file(path).async('string'));for(const data of all(vml,'ClientData'))if(['1','2'].includes(all(data,'Column')[0]?.textContent)){all(data,'Anchor')[0].textContent='38, 185, 1009, 8, 41, 18, 1014, 27';data.parentElement.setAttribute('style','position:absolute;margin-left:3276pt;margin-top:30909pt;width:302pt;height:169pt;visibility:hidden');}zip.file(path,s.serializeToString(vml));}
   return [...await zip.generateAsync({type:'uint8array'})];
  },[...X.write(book,{type:'buffer',bookType:'xlsx'})]));
  const normalHeaders=['供应商名称','申请付款金额','OA合同编号','合同总金额'],normalRows=[['甲',30,'HT001',30],['乙',40,'HT002',40]].map(row=>row.map(v=>({t:typeof v==='number'?'n':'s',v})));
  const result={applicationComments:{B2:'申请日期（付款节奏跟进）：2026年10月27日'},normalHeaders,normalRows,progressHeaders:['是否已付款','备注'],progressRows:[[{t:'s',v:'已付款'},{t:'s',v:'保留状态红色'}],[{t:'s',v:'待付款'},{t:'s',v:''}]],reviewHeaders:['核对说明'],reviewRows:[],negotiated:{sheetName:'本周谈判付款',headerRow:0,headers,rows:[1,2,3].map(r=>headers.map((_,c)=>source[X.utils.encode_cell({r,c})])),updates:[],columnOrder:[0,2,1,3,4]},highlights:{negotiatedCells:['B1','B2','B2'],normalRows:[1,2,2]}};
  const exportBytes=async(bytes,data)=>Buffer.from(await page.evaluate(async({bytes,result})=>[...new Uint8Array(await(await FundsExport(new Uint8Array(bytes),result)).arrayBuffer())],{bytes:[...bytes],result:data}));
  const inspect=async bytes=>page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=parse(await zip.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],fills=all(styles,'fills')[0],dxfs=all(styles,'dxfs')[0];
   if(all(styles,'dxfs').length!==1||Number(dxfs.getAttribute('count'))!==dxfs.children.length)throw Error('Export must have one dxfs container with an accurate count');
   const result={dxfCount:dxfs?.children.length||0,dxfPrefix:[...dxfs?.children||[]].slice(0,2).map(n=>s.serializeToString(n)),styleTags:[...styles.documentElement.children].map(n=>n.localName),styleNames:all(styles,'cellStyle').map(n=>({name:n.getAttribute('name'),hidden:n.getAttribute('hidden')})),sheets:{}},workbook=parse(await zip.file('xl/workbook.xml').async('string')),rels=new Map(all(parse(await zip.file('xl/_rels/workbook.xml.rels').async('string')),'Relationship').map(n=>[n.getAttribute('Id'),n.getAttribute('Target')]));
   for(const [name,filename] of [['谈判付款','xl/worksheets/sheet1.xml'],['正常付款','xl/worksheets/funds-normal.xml'],['付款申请核对','xl/worksheets/funds-review.xml'],['付款进度整理','xl/worksheets/funds-progress.xml'],['总额校验核对','xl/worksheets/funds-totalvalidation.xml']]){
    const entry=all(workbook,'sheet').find(n=>n.getAttribute('name')===name),target=rels.get(entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')),actual=new URL(target,'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1);
    const doc=parse(await zip.file(actual).async('string')),rules=all(doc,'cfRule').map(n=>{const dxf=dxfs?.children[Number(n.getAttribute('dxfId'))],background=dxf&&all(dxf,'bgColor')[0];return {formula:all(n,'formula')[0]?.textContent,priority:Number(n.getAttribute('priority')),stop:n.getAttribute('stopIfTrue'),range:n.parentElement.getAttribute('sqref'),color:dxf&&all(dxf,'fgColor')[0]?.getAttribute('rgb'),background:background?.getAttribute('rgb'),backgroundIndex:background?.getAttribute('indexed'),bold:!!dxf&&all(dxf,'b').length>0,dxfChildren:dxf?[...dxf.children].map(c=>c.localName):[]};}),cells={};
    for(const c of all(doc,'c')){const xf=xfs.children[Number(c.getAttribute('s')||0)],fill=fills.children[Number(xf.getAttribute('fillId')||0)];cells[c.getAttribute('r')]={fill:all(fill,'fgColor')[0]?.getAttribute('rgb'),locked:all(xf,'protection')[0]?.getAttribute('locked')};}
    result.sheets[filename]={rules,own:rules.filter(r=>r.formula===marker),cells,tags:[...doc.documentElement.children].map(n=>n.localName),hidden:all(doc,'col').filter(c=>c.getAttribute('hidden')==='1').map(c=>Number(c.getAttribute('min')))};
   }
   result.notes=[];result.noteShapes=[];result.notePositions={};
   for(const path of Object.keys(zip.files))if(path.endsWith('.xml')&&/comments/i.test(path)){const doc=parse(await zip.file(path).async('string')),authors=all(doc,'author').map(n=>n.textContent);for(const n of all(doc,'comment'))result.notes.push({ref:n.getAttribute('ref'),author:authors[Number(n.getAttribute('authorId'))],text:all(n,'t').map(t=>t.textContent).join(''),runs:all(n,'r').map(r=>({text:all(r,'t').map(t=>t.textContent).join(''),format:[...all(r,'rPr')[0]?.children||[]].map(p=>p.localName)}))});}
   for(const path of Object.keys(zip.files))if(path.endsWith('.vml')){const doc=parse(await zip.file(path).async('string'));for(const n of all(doc,'ClientData').filter(n=>n.getAttribute('ObjectType')==='Note')){const ref=XLSX.utils.encode_cell({r:Number(all(n,'Row')[0].textContent),c:Number(all(n,'Column')[0].textContent)});result.noteShapes.push({ref,visible:all(n,'Visible').length>0});result.notePositions[ref]={anchor:all(n,'Anchor')[0]?.textContent.split(',').map(Number),style:n.parentElement.getAttribute('style')};}}
   {if(result.notes.length!==result.noteShapes.length)throw Error('批注与 VML 图形数量不一致');if(new Set(result.noteShapes.map(n=>n.ref)).size!==result.noteShapes.length)throw Error('批注图形坐标重复');}
   return result;
  },{bytes:[...bytes],marker:MARKER});
  const output=await exportBytes(input,result),first=await inspect(output),saved=X.read(output,{cellNF:true}),neg=first.sheets['xl/worksheets/sheet1.xml'],normal=first.sheets['xl/worksheets/funds-normal.xml'];
  assert.deepEqual(saved.SheetNames,['谈判付款','谈判付款付款表付款节奏','正常付款','付款申请核对','付款进度整理','总额校验核对']);assert.equal(neg.own.length,0);assert.equal(normal.rules.length,0);assert.deepEqual(normal.hidden,[3,4]);
  for(const col of normalHeaders.map((_,i)=>X.utils.encode_col(i))){assert.equal(normal.cells[col+'2'].fill,'FFFF0000');assert.notEqual(normal.cells[col+'3'].fill,'FFFF0000');assert.notEqual(normal.cells[col+'1'].fill,'FFFF0000');}
  const adjacent=(info,address)=>{const p=X.utils.decode_cell(address),note=info.notePositions[address];assert.deepEqual(note.anchor.slice(0,4),[p.c+1,6,p.r,2]);assert(note.anchor[4]>=p.c+1);assert(note.anchor[6]>=p.r);const style=Object.fromEntries(note.style.split(';').filter(s=>s.includes(':')).map(s=>s.split(':')));assert.equal(style.top,'0pt');assert.equal(style.left,'0pt');assert.equal(parseFloat(style['margin-top']),31.5);assert.equal(style.visibility,address==='C2'?'hidden':undefined);};
  adjacent(first,'C2');
  assert.equal(first.sheets['xl/worksheets/funds-progress.xml'].cells.A2.fill,'FFFF0000');assert.notEqual(first.sheets['xl/worksheets/funds-progress.xml'].cells.A3.fill,'FFFF0000');
  assert.deepEqual(first.notePositions.B2,{anchor:[38,185,1009,8,41,18,1014,27],style:'position:absolute;margin-left:3276pt;margin-top:30909pt;width:302pt;height:169pt;visibility:hidden'},'Unrelated manual notes retain their original position');
  assert.equal(first.notes.find(n=>n.ref==='C2').author,'原作者');assert.equal(first.notes.length,2);
  assert.equal(saved.Sheets.谈判付款.C2.c[0].t,'保留原人工批注\n申请日期（付款节奏跟进）：2026年10月27日');assert.equal(saved.Sheets.谈判付款.C2.c[0].a,X.read(input).Sheets.本周谈判付款.B2.c[0].a);assert.equal(saved.Sheets.谈判付款.B2.c[0].t,'未标红的原批注');
  assert.equal(neg.cells.C2.fill,'FFFFFF00');assert.equal(neg.cells.C3.fill,'FFFF0000','Unrelated manual red is retained');assert.equal(neg.cells.C2.locked,'0');
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
  for(const col of normalHeaders.map((_,i)=>X.utils.encode_col(i))){assert.notEqual(sortedInfo.cells[col+'2'].fill,'FFFF0000');assert.equal(sortedInfo.cells[col+'3'].fill,'FFFF0000','Red follows the submitted record including hidden columns');}
  const next={...result,applicationComments:{C2:'申请日期（付款节奏跟进）：2026年10月28日；2026年10月29日'},negotiated:{...result.negotiated,sheetName:'谈判付款',headers:headers.map((_,i)=>saved.Sheets.谈判付款[X.utils.encode_cell({r:0,c:i})].v),columnOrder:[0,1,2,3,4]},highlights:{negotiatedCells:['C2'],negotiatedPaidCells:['C2'],normalRows:[2]}};
  const repeated=await exportBytes(output,next),again=await inspect(repeated);assert.deepEqual(again.sheets['xl/worksheets/sheet1.xml'].rules,neg.rules);assert.deepEqual(again.dxfPrefix,first.dxfPrefix);assert.equal(again.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');
  assert.equal(X.read(repeated).Sheets.谈判付款.C2.c[0].t,'保留原人工批注\n申请日期（付款节奏跟进）：2026年10月28日；2026年10月29日');
  adjacent(again,'C2');
  const withdrawn=await exportBytes(repeated,{...next,applicationComments:undefined,batchComments:undefined,highlights:undefined}),clean=await inspect(withdrawn),cleanNeg=clean.sheets['xl/worksheets/sheet1.xml'];
  assert.deepEqual(clean.notePositions,again.notePositions,'Without newly generated notes, existing note positions stay unchanged');
  assert(Object.values(clean.sheets['xl/worksheets/funds-normal.xml'].cells).every(c=>c.fill!=='FFFF0000'),'Re-upload withdraws obsolete normal fills');
  assert.equal(X.read(withdrawn).Sheets.谈判付款.C2.c[0].t,X.read(repeated).Sheets.谈判付款.C2.c[0].t,'No update retains the old application date');assert.equal(X.read(withdrawn).Sheets.谈判付款.B2.c[0].t,'未标红的原批注');
  assert.equal(cleanNeg.cells.C2.fill,'FF00FF00','Withdrawn tool fill restores the original green');assert.equal(cleanNeg.cells.C3.fill,'FFFF0000','Manual red survives');assert.equal(cleanNeg.cells.C2.locked,'0');assert.deepEqual(cleanNeg.rules,neg.rules);
  // Excel can split both tool lines and manual sentences across rich-text runs.
  const richInput=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
   for(const name of Object.keys(zip.files).filter(n=>/comments[^/]*\.xml$/i.test(n))){const doc=p.parseFromString(await zip.file(name).async('string'),'application/xml'),note=all(doc,'comment').find(n=>n.getAttribute('ref')==='C2');if(!note)continue;const text=all(note,'text')[0];text.replaceChildren();
    for(const [value,format]of [['人工备注：','b'],['申请日期（付款节奏跟进）：人工说明，请保留\r\n','i'],['申请日期（付款节奏','u'],['跟进）：2026年11月10日\r\n',''],['第一条人工记录\n\n第二条人工记录\n','b'],['申请日期（资金整理）：2026年',''],['10月1日','i']]){const run=doc.createElementNS(ns,'r'),t=doc.createElementNS(ns,'t');if(format){const props=doc.createElementNS(ns,'rPr');props.append(doc.createElementNS(ns,format));run.append(props);}t.setAttribute('xml:space','preserve');t.textContent=value;run.append(t);text.append(run);}zip.file(name,s.serializeToString(doc));
   }return [...await zip.generateAsync({type:'uint8array'})];
  },[...output]));
  const richManual='人工备注：申请日期（付款节奏跟进）：人工说明，请保留\n第一条人工记录\n\n第二条人工记录',richOutput=await exportBytes(richInput,next),rich=await inspect(richOutput),richNote=rich.notes.find(n=>n.ref==='C2');
  assert.equal(richNote.text,richManual+'\n'+next.applicationComments.C2,'Only complete tool lines are replaced, even when prefixes or dates span runs');
  assert.deepEqual(richNote.runs.slice(0,3),[{text:'人工备注：',format:['b']},{text:'申请日期（付款节奏跟进）：人工说明，请保留\n',format:['i']},{text:'第一条人工记录\n\n第二条人工记录',format:['b']}],'Manual text, blank lines, line breaks and rich-text formatting survive');
  adjacent(rich,'C2');assert.deepEqual(rich.notePositions.B2,first.notePositions.B2);
  const richRepeated=await inspect(await exportBytes(richOutput,next));assert.equal(richRepeated.notes.find(n=>n.ref==='C2').text,richNote.text,'Repeated export adds exactly one current tool line');assert.deepEqual(richRepeated.notes.find(n=>n.ref==='C2').runs,richNote.runs,'Repeated export retains manual rich-text runs');
  const richWithdrawn=await inspect(await exportBytes(richOutput,{...next,applicationComments:undefined,batchComments:undefined,highlights:undefined}));assert.equal(richWithdrawn.notes.find(n=>n.ref==='C2').text,richNote.text);assert.deepEqual(richWithdrawn.notes.find(n=>n.ref==='C2').runs,richNote.runs,'No current update retains all old note lines and rich-text runs');assert.deepEqual(richWithdrawn.notePositions,rich.notePositions);
  // Native editing changes just the fill; no tool rule may force it back to red.
  const edited=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=t=>p.parseFromString(t,'application/xml'),styles=parse(await zip.file('xl/styles.xml').async('string')),fills=all(styles,'fills')[0],xfs=all(styles,'cellXfs')[0];
   const fill=styles.createElementNS(ns,'fill'),pattern=styles.createElementNS(ns,'patternFill');pattern.setAttribute('patternType','solid');for(const tag of ['fgColor','bgColor']){const c=styles.createElementNS(ns,tag);c.setAttribute('rgb','FF00FFFF');pattern.append(c);}fill.append(pattern);const fillId=fills.children.length;fills.append(fill);fills.setAttribute('count',fills.children.length);
   for(const [filename,address]of [['xl/worksheets/sheet1.xml','C2'],['xl/worksheets/funds-normal.xml','A2']]){const doc=parse(await zip.file(filename).async('string')),c=all(doc,'c').find(n=>n.getAttribute('r')===address),xf=xfs.children[Number(c.getAttribute('s'))].cloneNode(true);xf.setAttribute('fillId',fillId);c.setAttribute('s',xfs.children.length);xfs.append(xf);zip.file(filename,s.serializeToString(doc));}
   xfs.setAttribute('count',xfs.children.length);zip.file('xl/styles.xml',s.serializeToString(styles));return [...await zip.generateAsync({type:'uint8array'})];
  },[...output]));
  const manual=await inspect(edited);assert.equal(manual.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FF00FFFF');assert.equal(manual.sheets['xl/worksheets/funds-normal.xml'].cells.A2.fill,'FF00FFFF');
  const retained=await inspect(await exportBytes(edited,{...next,applicationComments:undefined,batchComments:undefined,highlights:undefined}));assert.equal(retained.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FF00FFFF','Manual override survives re-upload without a new matching application');
  // Upgrade a prior export carrying the legacy owned rule; retain all unrelated rules.
  const legacy=Buffer.from(await page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),cf=doc.createElementNS(ns,'conditionalFormatting'),rule=doc.createElementNS(ns,'cfRule'),f=doc.createElementNS(ns,'formula');cf.setAttribute('sqref','C2');rule.setAttribute('type','expression');rule.setAttribute('dxfId','0');rule.setAttribute('priority','1');f.textContent=marker;rule.append(f);cf.append(rule);doc.documentElement.append(cf);zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...output],marker:MARKER}));
  const upgraded=await inspect(await exportBytes(legacy,next));assert.equal(upgraded.sheets['xl/worksheets/sheet1.xml'].own.length,0);assert.deepEqual(upgraded.sheets['xl/worksheets/sheet1.xml'].rules,neg.rules);assert.equal(upgraded.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');
  // Four independent fields replace only their own prefixes, including notes on unmatched cells.
  const batchLines=['批次（付款节奏跟进）：共3批，第1批','申请日期（付款节奏跟进）：2026年11月12日','OA备注时间（付款节奏跟进）：2026年11月13日 09:30:00','付款时间（付款节奏跟进）：2026年11月14日 10:00:00'];
  const batchData={...result,batchComments:{B2:batchLines,C2:['批次（付款节奏跟进）：共3批，第2批']},highlights:{...result.highlights,negotiatedPaidCells:['B2']}};
  const batchOutput=await exportBytes(input,batchData),batch=await inspect(batchOutput);
  assert.equal(batch.notes.find(n=>n.ref==='C2').text,'保留原人工批注\n'+batchLines.join('\n'));
  assert.equal(batch.notes.find(n=>n.ref==='B2').text,'未标红的原批注\n批次（付款节奏跟进）：共3批，第2批','An unmatched existing note receives its batch line');
  assert.equal(batch.notes.find(n=>n.ref==='B2').author,'原作者');assert.deepEqual(batch.noteShapes.map(n=>n.visible),first.noteShapes.map(n=>n.visible));
  assert.equal(batch.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');
  assert(batch.styleNames.some(n=>/^GYLSJQX_FUNDS_APPLIED_V3_FFFF0000_/.test(n.name)&&n.hidden==='1'));assert(first.styleNames.some(n=>/^GYLSJQX_FUNDS_APPLIED_V3_FFFFFF00_/.test(n.name)&&n.hidden==='1'));
  adjacent(batch,'C2');assert.deepEqual(batch.notePositions.B2.anchor.slice(0,4),[2,6,1,2]);
  const oversized=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
   for(const file of Object.keys(zip.files).filter(n=>n.endsWith('.vml'))){const doc=p.parseFromString(await zip.file(file).async('string'),'application/xml');for(const data of all(doc,'ClientData'))if(all(data,'Column')[0]?.textContent==='2'){data.parentElement.setAttribute('style','position:absolute;width:1200pt;height:2400pt;visibility:hidden');all(data,'Anchor')[0].textContent='3,6,1,2,25,20,100,2';}zip.file(file,s.serializeToString(doc));}
   return [...await zip.generateAsync({type:'uint8array'})];
  },[...batchOutput]));
  const compactData={...next,applicationComments:undefined,batchComments:{C2:batchLines}},compactBytes=await exportBytes(oversized,compactData),compact=await inspect(compactBytes);
  const noteSize=info=>Object.fromEntries(info.notePositions.C2.style.split(';').filter(s=>s.includes(':')).map(s=>s.split(':'))),compactStyle=noteSize(compact);
  assert.equal(compact.notes.find(n=>n.ref==='C2').text,batch.notes.find(n=>n.ref==='C2').text,'Resizing retains every note line');
  assert(parseFloat(compactStyle.height)<140,'A short note must shrink even when its original box is thousands of points high');
  assert(parseFloat(compactStyle.width)<500,'A short note must not retain the original oversized width');
  assert.deepEqual(compact.notePositions.B2,batch.notePositions.B2,'Unchanged notes retain their geometry');adjacent(compact,'C2');
  const compactRepeat=await inspect(await exportBytes(compactBytes,compactData));assert.deepEqual(compactRepeat.notePositions.C2,compact.notePositions.C2,'Repeated uploads do not grow the note box');
  const batchNext={...next,applicationComments:undefined,batchComments:{C2:['批次（付款节奏跟进）：共4批，第2批']},highlights:{negotiatedCells:['C2'],normalRows:[2]}};
  const batchUpdatedBytes=await exportBytes(batchOutput,batchNext),batchUpdated=await inspect(batchUpdatedBytes),batchText=batchUpdated.notes.find(n=>n.ref==='C2').text;
  assert.equal(batchText,'保留原人工批注\n'+batchLines.slice(1).join('\n')+'\n批次（付款节奏跟进）：共4批，第2批','Updating only batch count retains all previous dates');
  assert.equal(batchUpdated.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFFFF00','A previously paid fill does not prove a current match is paid');
  const batchRepeat=await inspect(await exportBytes(batchUpdatedBytes,batchNext));assert.equal(batchRepeat.notes.find(n=>n.ref==='C2').text,batchText);assert.equal(batchText.split('批次（付款节奏跟进）：').length,2);
  const batchNoUpdate=await inspect(await exportBytes(batchUpdatedBytes,{...batchNext,batchComments:{},highlights:undefined}));assert.equal(batchNoUpdate.notes.find(n=>n.ref==='C2').text,batchText);assert.equal(batchNoUpdate.sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FF00FF00');
  // Rich manual text and visible notes survive; a small box expands for all four fields and long text.
  const smallVisible=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
   for(const filename of Object.keys(zip.files).filter(n=>/comments[^/]*\.xml$/i.test(n))){const doc=p.parseFromString(await zip.file(filename).async('string'),'application/xml'),note=all(doc,'comment').find(n=>n.getAttribute('ref')==='C2');if(note){const text=all(note,'text')[0],run=doc.createElementNS(text.namespaceURI,'r'),t=doc.createElementNS(text.namespaceURI,'t');t.textContent='\n'+('保留人工长备注'.repeat(24));run.append(t);text.append(run);zip.file(filename,s.serializeToString(doc));}}
   for(const filename of Object.keys(zip.files).filter(n=>n.endsWith('.vml'))){const doc=p.parseFromString(await zip.file(filename).async('string'),'application/xml');for(const data of all(doc,'ClientData'))if(all(data,'Column')[0]?.textContent==='2'){data.parentElement.setAttribute('style','position:absolute;width:50pt;height:20pt;visibility:visible');if(!all(data,'Visible').length)data.append(doc.createElementNS(data.namespaceURI,'Visible'));}zip.file(filename,s.serializeToString(doc));}
   return [...await zip.generateAsync({type:'uint8array'})];
  },[...richInput]));
  const tall=await inspect(await exportBytes(smallVisible,{...next,applicationComments:undefined,batchComments:{C2:batchLines}})),tallNote=tall.notes.find(n=>n.ref==='C2'),tallStyle=Object.fromEntries(tall.notePositions.C2.style.split(';').filter(s=>s.includes(':')).map(s=>s.split(':')));
  assert(tallNote.text.includes('保留人工长备注'.repeat(24)));assert(tallNote.text.endsWith(batchLines.join('\n')));assert.deepEqual(tallNote.runs.slice(0,2),richNote.runs.slice(0,2));assert.equal(tallNote.author,'原作者');assert.equal(tall.noteShapes.find(n=>n.ref==='C2').visible,true);assert.equal(tallStyle.visibility,'visible');assert(parseFloat(tallStyle.width)>=216);assert(parseFloat(tallStyle.height)>parseFloat(compactStyle.height),'Long manual text receives more space than a short note');assert.deepEqual(tall.notePositions.C2.anchor.slice(0,4),[3,6,1,2]);
  // V2 marked only red: a yellow fill on a V2 cell is an intentional manual edit.
  const legacyStyle=async color=>Buffer.from(await page.evaluate(async({bytes,color})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),s=new XMLSerializer(),parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],styles=parse(await zip.file('xl/styles.xml').async('string')),doc=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),xfs=all(styles,'cellXfs')[0],fills=all(styles,'fills')[0],cell=all(doc,'c').find(n=>n.getAttribute('r')==='C2'),xf=xfs.children[Number(cell.getAttribute('s'))],base=xf.getAttribute('xfId');
   all(styles,'cellStyle').find(n=>n.getAttribute('xfId')===base).setAttribute('name','GYLSJQX_FUNDS_APPLIED_V2_'+base);
   const fill=fills.children[Number(xf.getAttribute('fillId'))].cloneNode(true);for(const tag of ['fgColor','bgColor'])all(fill,tag)[0].setAttribute('rgb',color);xf.setAttribute('fillId',fills.children.length);fills.append(fill);fills.setAttribute('count',fills.children.length);zip.file('xl/styles.xml',s.serializeToString(styles));return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...output],color}));
  const noMatch={...next,applicationComments:undefined,batchComments:undefined,highlights:undefined},v2Red=await legacyStyle('FFFF0000'),v2Yellow=await legacyStyle('FFFFFF00');
  assert.equal((await inspect(await exportBytes(v2Red,noMatch))).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FF00FF00','V2 red restores its original fill');
  assert.equal((await inspect(await exportBytes(v2Yellow,noMatch))).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFFFF00','V2 manual yellow is retained');
  const v2Matched=await exportBytes(v2Yellow,next);assert.equal((await inspect(v2Matched)).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');assert.equal((await inspect(await exportBytes(v2Matched,noMatch))).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFFFF00','A new matched status remembers the preceding manual yellow');
  const manualMatched=await exportBytes(edited,next);assert.equal((await inspect(manualMatched)).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FFFF0000');assert.equal((await inspect(await exportBytes(manualMatched,noMatch))).sheets['xl/worksheets/sheet1.xml'].cells.C2.fill,'FF00FFFF');
  const withTotals=await exportBytes(input,{...result,totalValidation:{headers:['来源行','核对结果','需要处理'],rows:[[{t:'n',v:2},{t:'s',v:'不一致'},{t:'s',v:'请核对原表金额\n保留原始数值'}]],stats:{}}}),totals=await inspect(withTotals),totalsBook=X.read(withTotals);
  assert.equal(totalsBook.Sheets.总额校验核对.B2.v,'不一致');assert.equal(totalsBook.Sheets.总额校验核对.C2.v,'请核对原表金额\n保留原始数值');assert.notEqual(totals.sheets['xl/worksheets/funds-totalvalidation.xml'].cells.B2.fill,'FFFF0000');assert.equal(totalsBook.Sheets.谈判付款.D2.f,'C2*3');
  await assert.rejects(exportBytes(input,{...result,highlights:{negotiatedCells:['B2'],negotiatedPaidCells:['B3']}}),/已付款位置必须属于/);
  await assert.rejects(exportBytes(input,{...result,batchComments:{B2:['批次（付款节奏跟进）：共2批\n第1批']}}),/批注数据无效/);
  await assert.rejects(exportBytes(input,{...result,batchComments:{B2:['无关前缀：不得覆盖人工批注']}}),/批注数据无效/);
  await assert.rejects(exportBytes(input,{...result,applicationComments:undefined,batchComments:{C2:batchLines},negotiated:{...result.negotiated,columnOrder:[0,1,2,3,4],managedColumns:[0,1,3,4]}}),/批注超出业务列/);
  assert.deepEqual(X.read(input).Sheets.本周谈判付款['!merges'],source['!merges']);
  delete source.B2.c;delete source.C2.c;
  const fresh=await exportBytes(X.write(book,{type:'buffer',bookType:'xlsx'}),result);assert.equal(X.read(fresh).Sheets.谈判付款.C2.c[0].t,'申请日期（付款节奏跟进）：2026年10月27日');
  assert.deepEqual((await inspect(fresh)).noteShapes,[{ref:'C2',visible:false}]);adjacent(await inspect(fresh),'C2');
  const removed=await exportBytes(fresh,{...next,applicationComments:undefined,batchComments:undefined,highlights:undefined});assert.equal(X.read(removed).Sheets.谈判付款.C2.c[0].t,'申请日期（付款节奏跟进）：2026年10月27日');assert.equal((await inspect(removed)).notes.length,1);
  book.Workbook={WBProps:{date1904:true}};const epoch=await exportBytes(X.write(book,{type:'buffer',bookType:'xlsx'}),{...result,normalHeaders:['申请日期'],normalRows:[[{t:'n',v:46322,z:'yyyy"年"m"月"d"日"'}]]});const epochBook=X.read(epoch,{cellNF:true});assert.equal(epochBook.Workbook.WBProps.date1904,true);assert.equal(X.SSF.format(epochBook.Sheets.正常付款.A2.z,epochBook.Sheets.正常付款.A2.v,{date1904:true}),'2026年10月27日');
  console.log('PASS: four-field batch notes, prefix-only updates, retained old dates and manual rich text, native VML sizing/visibility and column relocation; editable yellow/red fills, V2/V3 upgrade and manual overrides, repeat/withdrawal, full-row sorting, date1904, formulas and unrelated conditional formatting');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exit(1);});
