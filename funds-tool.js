(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const kinds = ['ledger', 'applications', 'negotiated', 'oa'];
  const inputs = {}, revisions = {}, decisions = {};
  let result = null, preparedBlob = null, busy = false;
  const memoryVersions = {};
  const dateNow = () => new Intl.DateTimeFormat('sv-SE', {timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  const timeNow = () => new Intl.DateTimeFormat('zh-CN', {timeZone:'Asia/Shanghai',dateStyle:'short',timeStyle:'medium'}).format(new Date());
  function nextDownload() {
    const date = dateNow();
    let versions;
    try { versions = JSON.parse(localStorage.getItem('funds-export-versions') || '{}'); } catch (_) { versions = {}; }
    const previous = Math.max(Number(versions?.[date]) || 0, memoryVersions[date] || 0);
    const version = Number.isSafeInteger(previous) && previous >= 0 ? previous + 1 : 1;
    return {date, version, name:`资金整理-${date}-V${version}.xlsx`};
  }
  function rememberDownload({date, version}) {
    memoryVersions[date] = version;
    try {
      const versions = JSON.parse(localStorage.getItem('funds-export-versions') || '{}');
      localStorage.setItem('funds-export-versions', JSON.stringify({...versions,[date]:version}));
      return true;
    } catch (_) { return false; }
  }
  function pending() { return result ? result.issues.filter(issue => !issue.resolved).length : 0; }
  function controls() {
    for (const kind of kinds) $(kind+'File').disabled = busy;
    $('negotiatedSheet').disabled = busy || !inputs.negotiated;
    $('analyze').disabled = busy || !kinds.every(kind => inputs[kind]) || !$('negotiatedSheet').value;
    $('download').disabled = busy || !result || !preparedBlob || pending() > 0;
    document.querySelectorAll('[data-decision]').forEach(select => { select.disabled = busy; });
    $('filenameHint').textContent = nextDownload().name;
  }
  function invalidate() {
    result = null; preparedBlob = null;
    for (const id of Object.keys(decisions)) delete decisions[id];
    for (const id of ['results','negotiatedPreviewSection','normalPreviewSection','reviewPreviewSection']) $(id).hidden = true;
    controls();
  }
  function valueText(cell, date1904=false) {
    if (cell == null) return '';
    if (typeof cell !== 'object') return String(cell);
    if (cell.t === 'e') return cell.w || '源单元格错误';
    if (cell.v == null) return cell.f ? '公式结果待核对' : '';
    try { return cell.z ? XLSX.SSF.format(cell.z, cell.v,{date1904}) : String(cell.v); } catch (_) { return String(cell.v); }
  }
  function preview(id, headers, rows, date1904=false) {
    const table = document.createElement('table'), head = table.createTHead().insertRow(), body = table.createTBody();
    const appliedRows = new Set(result?.highlights?.normalRows || []), appliedCells = new Set(result?.highlights?.negotiatedCells || []);
    const singleLineHeaders=['流程编号','付款自然周','申请付款的日期',...(id==='normalPreview'?['对公付款0A流程号','开票时间','OA合同编号']:[])];
    const hiddenColumns=headers.map(title=>id==='normalPreview'&&FundsCore.NORMAL_HIDDEN_HEADERS.includes(title));
    const flowColumn=headers.map(title=>singleLineHeaders.includes(String(title).normalize('NFKC').replace(/\s+/g,'')));
    for (const [i,title] of headers.entries()) { const th = document.createElement('th'); th.textContent = title;th.hidden=hiddenColumns[i];if(flowColumn[i])th.className='flow-code'; head.append(th); }
    for (const [r,row] of rows.slice(0,100).entries()) {
      const tr = body.insertRow();
      for (let i=0;i<headers.length;i++){
        const td=tr.insertCell(),text=valueText(row[i],date1904);td.hidden=hiddenColumns[i];td.textContent=flowColumn[i]?text.replace(/\r\n|[\r\n]/g,'；'):text;if(flowColumn[i])td.className='flow-code';if(headers[i]==='谈判明细')td.className='batch-detail';if(id==='reviewPreview'&&headers[i]==='核对说明')td.className='review-detail';
        const applied=id==='normalPreview'?appliedRows.has(r+2):id==='negotiatedPreview'&&appliedCells.has(XLSX.utils.encode_cell({r:r+result.negotiated.headerRow+1,c:result.negotiated.columnOrder[i]}));
        if(applied){td.classList.add('funds-applied');td.title=id==='negotiatedPreview'?(result.applicationComments?.[XLSX.utils.encode_cell({r:r+result.negotiated.headerRow+1,c:result.negotiated.columnOrder[i]})]||'已匹配本次付款申请'):'已匹配本次付款申请';}
      }
    }
    $(id).replaceChildren(table);
  }
  function showIssues() {
    const list = $('issueList'), scroll = list.scrollTop, fragment = document.createDocumentFragment();
    const issues = result.issues.filter(issue => !$('onlyPending').checked || !issue.resolved);
    for (const issue of issues) {
      const card = document.createElement('div'); card.className = 'issue'+(issue.resolved?' resolved':''); card.dataset.issueId = issue.id;
      const heading = document.createElement('h3'); heading.textContent = (issue.resolved?'已确认 · ':'') + (issue.targetLabel || '记录核对');
      const message = document.createElement('p'); message.textContent = issue.message;
      const source = document.createElement('p'); source.className = 'issue-source';
      source.textContent = [issue.source?.file, issue.source?.sheet, issue.source?.row ? '第 '+issue.source.row+' 行' : ''].filter(Boolean).join(' / ');
      const label = document.createElement('label'), select = document.createElement('select');
      label.textContent = issue.multiple ? '选择处理方式或多笔匹配记录' : '选择处理方式';
      select.dataset.decision = issue.id; select.multiple = Boolean(issue.multiple);
      select.setAttribute('aria-label', (issue.targetLabel || '') + ' ' + issue.message);
      if (!issue.multiple) select.add(new Option('请选择核对结果',''));
      else select.size = Math.min(6,Math.max(3,issue.choices.length));
      for (const choice of issue.choices) {
        const option = new Option(choice.label,choice.value);
        option.selected = (decisions[issue.id] || []).includes(choice.value);
        select.add(option);
      }
      select.addEventListener('change', () => {
        decisions[issue.id] = [...select.selectedOptions].map(option=>option.value).filter(Boolean);
        analyze(true);
      });
      label.append(select); card.append(heading,message,source,label); fragment.append(card);
    }
    if (!issues.length) { const p = document.createElement('p'); p.textContent = '没有待核对项，可以检查预览后下载。'; fragment.append(p); }
    list.replaceChildren(fragment); list.scrollTop = scroll;
  }
  async function analyze(review=false) {
    preparedBlob = null;
    try {
      const source = {...inputs, negotiated:{...inputs.negotiated,sheetName:$('negotiatedSheet').value}};
      result = FundsCore.build(source,decisions);
      $('results').hidden = false;
      $('pendingCount').textContent = pending();
      $('stats').textContent = `谈判付款 ${result.stats.negotiatedRows} 条；正常付款 ${result.normalRows.length} 条；流程编号已填写 ${result.stats.oaMatched} 条；付款申请核对表 ${result.reviewRows.length} 条；待处理 ${pending()} 项。`;
      $('classificationSummary').textContent = `资金台账 ${result.stats.ledgerRows} 条：${result.stats.negotiatedLedgerRows} 条按台账标记、付款申请分类或谈判表识别为谈判付款，${result.normalRows.length} 条正常付款，${result.stats.excludedNormalRows} 条无对公付款 OA 号未纳入正常付款，${result.stats.unclassifiedLedgerRows} 条暂未分类。`;
      showIssues();
      preview('normalPreview',result.normalHeaders,result.normalRows);
      const order=result.negotiated.columnOrder,negotiatedHeaders=order.map(c=>result.negotiated.headers[c]);
      preview('negotiatedPreview',negotiatedHeaders,result.negotiated.rows.map(row=>order.map(c=>row[c])),Boolean(inputs.negotiated.book.Workbook?.WBProps?.date1904));
      preview('reviewPreview',result.reviewHeaders,result.reviewRows.map(row=>row.map((cell,i)=>result.reviewHeaders[i]==='核对说明'?{...cell,v:FundsCore.formatReviewText(cell.v)}:cell)));
      $('negotiatedPreviewSection').hidden = false; $('normalPreviewSection').hidden = false; $('reviewPreviewSection').hidden = false;
      $('status').textContent = pending() ? `已完成整理，请处理下方 ${pending()} 项核对。` : '核对已完成，正在验证导出结果…';
      if (!pending()) {
        busy = true; controls();
        const blob = await FundsExport(inputs.negotiated.bytes,result);
        if (FundsExport.lastWarnings?.length) throw Error('以下公式无法完成结果校验，请调整源公式后重新上传：'+FundsExport.lastWarnings.join('；'));
        const saved = XLSX.read(await blob.arrayBuffer(),{type:'array',cellNF:true});
        if (saved.SheetNames.join('|') !== '谈判付款|正常付款|付款申请核对') throw Error('导出工作表检查未通过');
        const sheet = saved.Sheets['谈判付款'];
        const rows = result.negotiated.rows.map((_,index)=>result.negotiated.headers.map((_,col)=>sheet[XLSX.utils.encode_cell({r:index+result.negotiated.headerRow+1,c:col})]));
        preview('negotiatedPreview',negotiatedHeaders,rows,Boolean(saved.Workbook?.WBProps?.date1904));
        preparedBlob = blob;
        $('status').textContent = '核对及导出检查已完成，预览与下载使用同一份结果。';
      }
    } catch (e) {
      result = null; $('results').hidden = true;
      for (const id of ['normalPreviewSection','negotiatedPreviewSection','reviewPreviewSection']) $(id).hidden = true;
      $('status').textContent = (review?'核对失败：':'整理失败：') + e.message;
    }
    busy = false; controls();
  }
  for (const kind of kinds) $(kind+'File').addEventListener('change', async event => {
    const file = event.target.files[0]; if (!file) return;
    const revision = revisions[kind] = (revisions[kind] || 0)+1;
    delete inputs[kind]; invalidate();
    const info = $(kind+'Info'); info.classList.remove('file-error'); info.textContent = '正在读取 '+file.name+'…';
    if (kind === 'negotiated') $('negotiatedSheet').replaceChildren(new Option('正在读取…',''));
    try {
      if (file.size > 50*1024*1024) throw Error('单个文件不能超过 50 MB');
      if (kind === 'negotiated' && !/\.xlsx$/i.test(file.name)) throw Error('谈判付款汇总请使用 .xlsx，以保留原表结构和格式');
      const bytes = await file.arrayBuffer();
      if (revision !== revisions[kind]) return;
      const book = XLSX.read(bytes,{type:'array',cellStyles:true,cellNF:true,cellDates:false});
      if (!book.SheetNames.length) throw Error('文件没有工作表');
      if (kind === 'negotiated') {
        const candidates = FundsCore.negotiatedSheets(book);
        if (!candidates.length) throw Error('没有找到名称包含“谈判付款”的工作表，请重新选择文件');
        $('negotiatedSheet').replaceChildren(new Option(candidates.length===1?'已识别谈判付款工作表':'存在多个谈判付款版本，请选择',''));
        for (const name of candidates) $('negotiatedSheet').add(new Option(name,name));
        if (candidates.length === 1) $('negotiatedSheet').value = candidates[0];
      }
      inputs[kind] = {book,name:file.name,bytes};
      info.textContent = file.name+'\n引用时间：'+timeNow();
      $('status').textContent = '文件已读取。上传四份文件并确认谈判工作表后，点击“整理并核对”。';
    } catch (e) {
      if (revision !== revisions[kind]) return;
      info.textContent = file.name+'\n读取失败：'+e.message; info.classList.add('file-error');
      $('status').textContent = '请处理文件读取提示后继续。';
      if (kind === 'negotiated') $('negotiatedSheet').replaceChildren(new Option('请重新上传谈判付款汇总',''));
    }
    controls();
  });
  $('negotiatedSheet').addEventListener('change',()=>{invalidate();$('status').textContent='已更换谈判付款工作表，请重新整理并核对。';});
  $('analyze').addEventListener('click',()=>analyze());
  $('onlyPending').addEventListener('change',()=>{if(result)showIssues();});
  $('download').addEventListener('click',async()=>{
    if (busy || !result || !preparedBlob || pending()) return;
    busy = true; controls(); $('status').textContent = '正在生成三张工作表，请稍候…';
    let url;
    try {
      const blob = preparedBlob;
      const download = nextDownload();
      url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = download.name; document.body.append(a); a.click(); a.remove();
      const saved = rememberDownload(download);
      $('status').textContent = '已生成：'+download.name+(saved?'':'。浏览器未能保存版本号，本次页面内仍会递增。');
    } catch (e) { $('status').textContent = '生成失败：'+e.message; }
    finally { if(url)setTimeout(()=>URL.revokeObjectURL(url),1000);busy=false;controls(); }
  });
  if (window.XLSX && window.JSZip && window.FundsCore && window.FundsExport) {
    controls(); $('status').textContent = '请上传四份文件。';
  } else $('status').textContent = '工具未加载完整，请刷新页面重试。';
})();
