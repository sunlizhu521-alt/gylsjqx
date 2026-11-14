(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.ProductDimensionCore;
  const sources = {dimension:null,material:null};
  let analysis = null, result = null, busy = false, memoryVersion = 0;
  const versionKey = 'gylsjqx-product-dimension-version';
  function controls() {
    for (const kind of ['dimension','material']) $(kind+'File').disabled = busy;
    $('clear').disabled = busy || !sources.dimension && !sources.material;
    $('download').disabled = busy || !result?.canExport;
  }
  function invalidate() {
    analysis=result=null;
    for(const id of ['result','issue-panel']) $(id).hidden=true;
    for(const id of ['preview','issues']) $(id).replaceChildren();
    controls();
  }
  function table(headers,rows,container) {
    const t=document.createElement('table'), h=t.createTHead().insertRow(), b=t.createTBody();
    for(const text of headers) { const th=document.createElement('th');th.scope='col';th.textContent=text;h.append(th); }
    for(const values of rows) {const row=b.insertRow();for(const value of values) row.insertCell().textContent=value ?? '';}
    container.append(t);
  }
  function showIssues() {
    $('issues').replaceChildren();
    for(const issue of result.issues) {const li=document.createElement('li');li.textContent=issue.message;$('issues').append(li);}
    $('issue-panel').hidden=!result.issues.length;
    $('issue-panel').open=result.issues.some(issue=>issue.severity==='error');
  }
  function calculate() {
    result=C.resolve(analysis);
    showIssues();
    $('result').hidden=!result.canExport;
    if(!result.canExport) { $('status').textContent='请修正下方来源问题后重新上传。';controls();return; }
    const original=analysis.groups.reduce((n,group)=>n+group.rows.length,0), unique=analysis.groups.length;
    $('counts').textContent=`原商品 ${original.toLocaleString('zh-CN')} 行；原唯一编码 ${unique.toLocaleString('zh-CN')} 个；重复 ${original-unique} 行；新增 ${analysis.added.length} 个编码；完成后 ${(unique+analysis.added.length).toLocaleString('zh-CN')} 行。`;
    $('preview').replaceChildren();
    const indices=['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组'].map(name=>analysis.headers.indexOf(name));
    table(indices.map(i=>analysis.headers[i]),analysis.added.slice(0,100).map(row=>indices.map(i=>row.values[i])),$('preview'));
    $('preview-note').textContent=analysis.added.length ? `显示前 ${Math.min(100,analysis.added.length)} 个新增编码；下载包含全部保留商品和新增商品。` : '物料表中的编码均已存在，无需追加。';
    $('status').textContent='核对完成，可下载。重复编码统一保留原表第一行。';
    controls();
  }
  async function readBook(bytes,zip) {
    const book=XLSX.read(bytes,{type:'array',cellFormula:true,cellNF:true,cellStyles:true,sheetStubs:true});
    const parse=xml=>new DOMParser().parseFromString(xml,'application/xml');
    const workbook=parse(await zip.file('xl/workbook.xml').async('string'));
    const relFile=zip.file('xl/_rels/workbook.xml.rels');
    if(!relFile)throw Error('工作簿缺少工作表关系');
    const rels=new Map(Array.from(parse(await relFile.async('string')).getElementsByTagNameNS('*','Relationship')).map(rel=>[rel.getAttribute('Id'),rel.getAttribute('Target')]));
    for(const sheet of workbook.getElementsByTagNameNS('*','sheet')) {
      const relationship=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id') || sheet.getAttribute('r:id');
      const target=rels.get(relationship),data=book.Sheets[sheet.getAttribute('name')];
      if(!target || !data)continue;
      const path=new URL(target,'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1),entry=zip.file(path);
      if(!entry)throw Error('工作簿缺少工作表文件');
      const xml=await entry.async('string');
      if(!/<(?:\w+:)?f(?:\s|>|\/)/.test(xml))continue;
      // SheetJS reports string formulas without a saved value as empty strings.
      for(const match of xml.matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)) {
        if(!match[2] || !/<(?:\w+:)?f(?:\s|>|\/)/.test(match[2]) || /<(?:\w+:)?v(?:\s|>|\/)/.test(match[2]))continue;
        const address=match[1].match(/\br=["']([^"']+)["']/)?.[1];
        if(address && data[address])data[address].v=undefined;
      }
    }
    return book;
  }
  async function upload(kind) {
    const file=$(kind+'File').files[0];sources[kind]=null;invalidate();
    $(kind+'Reference').textContent=file ? file.name : `尚未上传${kind==='dimension'?'商品维度表':'物料表'}`;
    if(!file){$('status').textContent='请上传商品维度表和物料表。';return;}
    busy=true;controls();$('status').textContent='正在读取全部数据行…';
    try {
      if(!/\.xlsx$/i.test(file.name))throw Error('请上传 .xlsx 文件');
      if(file.size>60*1024*1024)throw Error('文件超过 60 MB，请缩小文件后上传');
      const bytes=await file.arrayBuffer(),zip=await JSZip.loadAsync(bytes);
      if(!zip.file('xl/workbook.xml'))throw Error('不是有效的 .xlsx 工作簿');
      const book=await readBook(bytes,zip);
      sources[kind]={bytes,book,filename:file.name};
      $(kind+'Reference').textContent=`${file.name}\n引用时间：${new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false})}`;
      if(sources.dimension && sources.material) {
        analysis=C.analyze(sources.dimension.book,sources.material.book);calculate();
      } else $('status').textContent=`请继续上传${kind==='dimension'?'物料表':'商品维度表'}。`;
    } catch(error) {sources[kind]=null;invalidate();$('status').textContent='读取失败：'+error.message;}
    finally {busy=false;controls();}
  }
  for(const kind of ['dimension','material'])$(kind+'File').addEventListener('change',()=>upload(kind));
  $('clear').addEventListener('click',()=>{
    for(const kind of ['dimension','material']) {sources[kind]=null;$(kind+'File').value='';$(kind+'Reference').textContent=`尚未上传${kind==='dimension'?'商品维度表':'物料表'}`;}
    invalidate();$('status').textContent='请上传商品维度表和物料表。';
  });
  function nextDownload() {
    let saved=0;try{saved=Number(localStorage.getItem(versionKey));}catch(_){}
    const stem=sources.dimension.filename.replace(/\.xlsx$/i,''),suffix=stem.match(/[_-]V(\d+)$/i);
    const inputVersion=Number(suffix?.[1] || 0);
    const version=Math.max(Number.isSafeInteger(saved)&&saved>=0?saved:0,memoryVersion,Number.isSafeInteger(inputVersion)?inputVersion:0)+1;
    if(!Number.isSafeInteger(version))throw Error('版本号超出有效范围，请核对文件名或浏览器版本记录');
    return {version,name:`${stem.replace(/[_-]V\d+$/i,'')}_V${version}.xlsx`};
  }
  function verify(bytes) {
    const book=XLSX.read(bytes,{type:'array',cellFormula:true,cellNF:true}),sheet=book.Sheets[analysis.dimensionSheet];
    if(book.SheetNames.join('\0')!==sources.dimension.book.SheetNames.join('\0') || !sheet)throw Error('导出工作表校验未通过');
    const expectedRange=`A1:${XLSX.utils.encode_col(analysis.headers.length-1)}${analysis.headerRow+result.rows.length}`;
    if(sheet['!ref']!==expectedRange)throw Error('导出行数校验未通过');
    const codes=new Set(),codeColumn=analysis.headers.indexOf('物料编码');
    result.rows.forEach((row,index)=>{
      const r=analysis.headerRow+index;
      row.values.forEach((value,c)=>{
        const cell=sheet[XLSX.utils.encode_cell({r,c})],actual=cell?.v ?? '';
        if(cell?.f || actual!==(value ?? ''))throw Error(`导出内容校验未通过：第${r+1}行，${analysis.headers[c]}`);
      });
      const code=sheet[XLSX.utils.encode_cell({r,c:codeColumn})];
      if(code?.t!=='s' || codes.has(code.v))throw Error('导出物料编码唯一性校验未通过');
      codes.add(code.v);
    });
  }
  $('download').addEventListener('click',async()=>{
    if(busy || !result?.canExport)return;
    busy=true;controls();$('status').textContent='正在生成并校验 Excel…';let url;
    try {
      const bytes=await window.ProductDimensionExport(sources.dimension.bytes,analysis,result);verify(bytes);
      const {version,name}=nextDownload();
      url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
      memoryVersion=version;let saved=true;
      try{localStorage.setItem(versionKey,String(version));}catch(_){saved=false;}
      $('status').textContent=`已生成：${name}${saved?'':'。浏览器未能保存版本号，本页内仍会递增。'}`;
    }catch(error){$('status').textContent='导出失败：'+error.message;}
    finally{if(url)setTimeout(()=>URL.revokeObjectURL(url),1000);busy=false;controls();}
  });
  if(!C || !window.ProductDimensionExport || !window.XLSX || !window.JSZip){$('status').textContent='工具加载失败，请刷新页面重试。';return;}
  controls();$('status').textContent='请上传商品维度表和物料表。';
})();
