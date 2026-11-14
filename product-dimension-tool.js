(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.ProductDimensionCore;
  const sources = {dimension:null,material:null};
  let analysis = null, result = null, busy = false, memoryVersion = 0, overrides = Object.create(null);
  const versionKey = 'gylsjqx-product-dimension-version';
  function controls() {
    for (const kind of ['dimension','material']) $(kind+'File').disabled = busy;
    $('clear').disabled = busy || !sources.dimension && !sources.material;
    $('download').disabled = busy || !result?.canExport;
    $('review-filter').disabled=busy;
    $('reviews').querySelectorAll('input').forEach(input=>{input.disabled=busy;});
  }
  function invalidate() {
    analysis=result=null;overrides=Object.create(null);
    for(const id of ['result','issue-panel','review-panel']) $(id).hidden=true;
    for(const id of ['preview','issues','reviews']) $(id).replaceChildren();
    $('review-filter').value='pending';
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
  const dateFields=['金蝶创建时间','当前日期'];
  const computedFields=['当前日期','上架时间（月）','生命周期'];
  const priceFields=['财务加权平均价','不含税结算价','结算价'];
  function displayValue(row,header) {
    const value=row.values[analysis.headers.indexOf(header)];
    return dateFields.includes(header)&&typeof value==='number' ? XLSX.SSF.format('yyyy-mm-dd',value,{date1904:analysis.date1904}) : value ?? '';
  }
  function sourceLabel(rows) {
    return rows?.length ? `${analysis.dimensionSheet}原第${rows.slice(0,8).join('、')}行${rows.length>8?`等${rows.length}行`:''}` : '';
  }
  function showReviews() {
    const expanded=new Set(Array.from($('reviews').querySelectorAll('details.review[open]')).map(item=>item.dataset.code));
    const focused=document.activeElement?.closest('#reviews input'),focusId=focused?.id;
    const selection=focused ? [focused.selectionStart,focused.selectionEnd] : null;
    $('reviews').replaceChildren();
    const rows=result.rows.filter(row=>row.isNew),pending=result.pendingFields || [],pendingCodes=new Set(pending.map(item=>item.code));
    const rowIndices=new Map(rows.map((row,index)=>[row.code,index]));
    const imported=new Map(analysis.added.map(row=>[row.code,row.values]));
    $('review-panel').hidden=!rows.length;
    $('pending-counts').textContent=pending.length ? `待补全 ${pendingCodes.size} 个商品、${pending.length} 个字段。请补齐后下载。` : '11列必填校验全部通过，可以下载。';
    const visible=rows.filter(row=>$('review-filter').value==='all'||pendingCodes.has(row.code));
    const visibleCodes=new Set(visible.map(row=>row.code));
    for(const code of expanded)if(!visibleCodes.has(code))expanded.delete(code);
    if(!visible.length){const p=document.createElement('p');p.className='muted';p.textContent='没有待补全商品。可切换到“全部新增商品”核对自动填写结果。';$('reviews').append(p);}
    for(const [index,row] of visible.entries()) {
      const missing=pending.filter(item=>item.code===row.code),missingFields=new Set(missing.map(item=>item.field));
      const card=document.createElement('details');card.className='review';card.dataset.code=row.code;
      card.open=expanded.has(row.code)||(!expanded.size&&index===0);
      const summary=document.createElement('summary');summary.textContent=`${row.code} · ${displayValue(row,'金蝶名称')}（${missing.length?`待补 ${missing.length} 项`:'必填已齐全'}）`;card.append(summary);
      const provenance=document.createElement('p');provenance.className='muted';provenance.textContent=`物料表：${analysis.materialSheet}第${row.sourceRow}行`;card.append(provenance);
      const grid=document.createElement('div');grid.className='editor-grid';
      const editingFields=C.EDITABLE_HEADERS.map((header,fieldIndex)=>({header,fieldIndex}));
      editingFields.sort((a,b)=>Number(missingFields.has(b.header))-Number(missingFields.has(a.header)));
      for(const {fieldIndex,header} of editingFields) {
        const meta=row.fields?.[header] || {},wrap=document.createElement('div');wrap.className='editor-field';
        const required=C.REQUIRED_HEADERS.includes(header)||header==='金蝶创建时间';
        const label=document.createElement('label'),input=document.createElement('input');
        input.type='text';input.id=`edit-${rowIndices.get(row.code)}-${fieldIndex}`;input.dataset.code=row.code;input.dataset.field=header;
        let shown=displayValue(row,header);
        if(header==='金蝶创建时间'&&missingFields.has(header)){
          const original=imported.get(row.code)?.[analysis.headers.indexOf(header)];
          shown=typeof original==='number'?XLSX.SSF.format('yyyy-mm-dd',original,{date1904:analysis.date1904}):original??shown;
        }
        input.autocomplete='off';input.value=String(Object.hasOwn(overrides[row.code]||{},header)?overrides[row.code][header]:shown);input.setAttribute('aria-invalid',String(missingFields.has(header)));
        if(required)input.setAttribute('aria-required','true');
        if(priceFields.includes(header))input.inputMode='decimal';
        label.htmlFor=input.id;label.textContent=header+(required?' *':'');
        const suggestions=document.createElement('datalist');suggestions.id=input.id+'-options';input.setAttribute('list',suggestions.id);
        for(const candidate of (meta.candidates || []).slice(0,200)) {const option=document.createElement('option');option.value=String(candidate.value);suggestions.append(option);}
        const note=document.createElement('small');note.id=input.id+'-note';note.className=missingFields.has(header)?'required-note':'muted';input.setAttribute('aria-describedby',note.id);
        const reason=missing.find(item=>item.field===header)?.reason || meta.reason || '';
        const candidates=(meta.candidates || []).slice(0,4).map(item=>`${item.value}（${sourceLabel(item.sourceRows)||'历史候选'}）`).join('；');
        const origin=header==='金蝶创建时间'?`${analysis.materialSheet}物料表原第${row.sourceRow}行`:sourceLabel(meta.sourceRows);
        note.textContent=[meta.method==='manual'?'人工填写':'',reason,origin,candidates?'候选：'+candidates:''].filter(Boolean).join('\n');
        input.addEventListener('change',()=>{
          if(busy)return;
          if(!Object.hasOwn(overrides,row.code))overrides[row.code]=Object.create(null);
          overrides[row.code][header]=input.value;
          calculate();
        });
        wrap.append(label,input,suggestions,note);grid.append(wrap);
      }
      card.append(grid);
      const computed=document.createElement('p');computed.className='computed';
      for(const header of computedFields){
        const output=document.createElement('output');output.dataset.code=row.code;output.dataset.field=header;
        output.setAttribute('aria-label',header);output.textContent=String(displayValue(row,header));
        computed.append(header+'：',output,'　');
      }
      card.append(computed);
      $('reviews').append(card);
    }
    const restored=focusId&&document.getElementById(focusId);
    if(restored){restored.focus({preventScroll:true});if(selection)restored.setSelectionRange(...selection);}
  }
  function calculate() {
    result=C.resolve(analysis,overrides);
    showIssues();
    const fatal=result.issues.some(issue=>issue.severity==='error');
    $('result').hidden=fatal;
    if(fatal) { $('review-panel').hidden=true;$('status').textContent='请修正下方来源问题后重新上传。';controls();return; }
    const original=analysis.groups.reduce((n,group)=>n+group.rows.length,0), unique=analysis.groups.length;
    $('counts').textContent=`原商品 ${original.toLocaleString('zh-CN')} 行；原唯一编码 ${unique.toLocaleString('zh-CN')} 个；重复 ${original-unique} 行；新增 ${analysis.added.length} 个编码；完成后 ${(unique+analysis.added.length).toLocaleString('zh-CN')} 行。处理日期：${analysis.asOfDate}。`;
    $('preview').replaceChildren();
    const added=result.rows.filter(row=>row.isNew);
    table(analysis.headers,added.slice(0,100).map(row=>analysis.headers.map(header=>displayValue(row,header))),$('preview'));
    let auto=0,manual=0,unknown=0,emptyPrices=0;
    for(const row of added)for(const [field,meta] of Object.entries(row.fields||{})){
      const value=row.values[analysis.headers.indexOf(field)];
      if(priceFields.includes(field)&&(value===''||value==null))emptyPrices++;
      else if(meta.method==='manual')manual++;
      else if(meta.method==='pending')unknown++;
      else if(!['empty'].includes(meta.method)&&field!=='金蝶创建时间')auto++;
    }
    $('enrichment-counts').textContent=`自动补全 ${auto} 个字段；人工填写 ${manual} 个字段；待确认 ${unknown} 个字段；缺价 ${emptyPrices} 项。`;
    $('preview-note').textContent=analysis.added.length ? `显示前 ${Math.min(100,analysis.added.length)} 个新增编码；下载包含全部保留商品和新增商品。` : '物料表中的编码均已存在，无需追加。';
    showReviews();
    $('status').textContent=result.canExport?'11列必填校验通过，可下载。':`还有 ${result.pendingRows} 个商品需要补全，请填写下方必填项。`;
    controls();
  }
  $('review-filter').addEventListener('change',()=>{if(result)showReviews();});
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
