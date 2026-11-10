(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.FundsBudgetSummaryCore;
  let result = null, sourceBytes = null, busy = false;
  const money = cents => (cents / 100).toLocaleString('zh-CN', {minimumFractionDigits:2,maximumFractionDigits:2});
  const month = key => key ? key.replace(/^(\d{4})-(\d{2})$/, (_,y,m) => `${y}年${Number(m)}月`) : '时间待定';
  function controls() {
    $('sourceFile').disabled = $('clear').disabled = busy;
    $('download').disabled = busy || !result?.canExport;
  }
  function reset() {
    result = sourceBytes = null;
    $('result').hidden = $('issue-panel').hidden = true;
    $('preview').replaceChildren(); $('issues').replaceChildren(); $('counts').textContent = '';
    controls();
  }
  function showIssues(errors, warnings) {
    $('issues').replaceChildren();
    for (const text of [...errors,...warnings]) { const li = document.createElement('li'); li.textContent = text; $('issues').append(li); }
    $('issue-panel').hidden = !errors.length && !warnings.length;
    $('issue-panel').open = Boolean(errors.length);
    $('issue-title').textContent = errors.length ? `需要修正（${errors.length} 项）` : `需要留意（${warnings.length} 项）`;
  }
  function preview() {
    const table = document.createElement('table'), head = table.createTHead().insertRow(), body = table.createTBody();
    for (const label of C.HEADERS) { const th = document.createElement('th'); th.scope = 'col'; th.textContent = label; head.append(th); }
    const append = (labels, amounts, total) => {
      const row = body.insertRow(); if (total) row.className = 'total';
      for (const value of [...labels,money(amounts.totalCents),money(amounts.paidCents),money(amounts.remainingCents),'']) row.insertCell().textContent = value;
    };
    for (const g of result.groups) append([month(g.monthKey),g.category,g.state],g,false);
    append(['合计','',''],result.totals,true);
    $('preview').replaceChildren(table); $('result').hidden = false;
    $('counts').textContent = `（${result.recordCount} 条明细，${result.groups.length} 个分组）`;
  }
  async function readBook(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    if (!zip.file('xl/workbook.xml')) throw Error('请上传有效的 .xlsx 工作簿');
    // SheetJS treats an uncached string formula as an empty value; retain the missing-result signal for validation.
    let changed = false;
    for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
      const xml = await zip.file(name).async('string');
      const next = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (cell,attrs,content) => {
        if (!content || !/\bt="str"/.test(attrs) || !/<f(?:\s|>|\/)/.test(content) || /<v(?:\s|>|\/)/.test(content)) return cell;
        changed = true; return '<c' + attrs.replace(/\s+t="str"/,'') + '>' + content + '</c>';
      });
      if (next !== xml) zip.file(name,next);
    }
    return XLSX.read(changed ? await zip.generateAsync({type:'uint8array'}) : bytes,{type:'array',cellNF:true,cellFormula:true,sheetStubs:true});
  }
  $('sourceFile').addEventListener('change',async () => {
    const file = $('sourceFile').files[0]; reset();
    if (!file) { $('reference').textContent = '尚未引用文件'; $('clear').hidden = true; $('status').textContent = '请上传文件。'; return; }
    busy = true; controls(); $('clear').hidden = false;
    $('reference').textContent = `${file.name}\n引用时间：${new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false})}`;
    $('status').textContent = '正在读取并汇总…';
    try {
      if (!/\.xlsx$/i.test(file.name)) throw Error('请使用 .xlsx 文件，以保留原工作表');
      if (file.size > 60 * 1024 * 1024) throw Error('文件超过 60 MB，请缩小文件后上传');
      const bytes = await file.arrayBuffer(), book = await readBook(bytes);
      result = C.build({book,filename:file.name});
      showIssues(result.errors,result.warnings);
      if (result.canExport) { sourceBytes = bytes; preview(); $('status').textContent = '汇总完成，可下载。'; }
      else $('status').textContent = '整理未完成，请修正下方问题后重新上传。';
    } catch (error) { reset(); $('status').textContent = '读取失败：' + error.message; }
    finally { busy = false; controls(); }
  });
  $('clear').addEventListener('click',() => {
    reset(); $('sourceFile').value = ''; $('reference').textContent = '尚未引用文件';
    $('clear').hidden = true; $('status').textContent = '请上传文件。';
  });
  $('download').addEventListener('click',async () => {
    if (busy || !result?.canExport || !sourceBytes) return;
    busy = true; controls(); $('status').textContent = '正在生成并校验 Excel…';
    let url;
    try {
      const bytes = await window.FundsBudgetSummaryExport(sourceBytes,result);
      const book = XLSX.read(bytes,{type:'array'}), sheet = book.Sheets['付款预算汇总'];
      const rows = XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});
      if (book.SheetNames[0] !== '付款预算汇总' || JSON.stringify(rows[0]) !== JSON.stringify(C.HEADERS) || rows.length !== result.groups.length + 2) throw Error('导出结构校验未通过');
      const totals = rows.at(-1);
      if ([result.totals.totalCents,result.totals.paidCents,result.totals.remainingCents].some((v,i) => Math.round(totals[i+3]*100) !== v)) throw Error('导出金额校验未通过');
      const date = new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
      const name = `付款预算汇总-${date}.xlsx`;
      url = URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const a = document.createElement('a'); a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
      $('status').textContent = '已生成：' + name;
    } catch (error) { $('status').textContent = '导出失败：' + error.message; }
    finally { if (url) setTimeout(() => URL.revokeObjectURL(url),1000); busy = false; controls(); }
  });
  if (!C || !window.FundsBudgetSummaryExport || !window.XLSX || !window.JSZip) { $('status').textContent = '工具加载失败，请刷新页面重试。'; return; }
  controls(); $('status').textContent = '请上传文件。';
})();
