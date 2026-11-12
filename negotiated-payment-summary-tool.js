(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.NegotiatedPaymentSummaryCore;
  let source = null, result = null, decisions = {}, reviews = [], busy = false;
  const versionKey = 'negotiated-payment-summary-export-versions', memoryVersions = {};
  const money = cents => cents == null ? '' : (cents / 100).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
  function controls() {
    $('sourceFile').disabled = $('clear').disabled = busy;
    $('download').disabled = busy || !result?.canExport;
    $('reviews').querySelectorAll('select').forEach(select => { select.disabled = busy; });
  }
  function reset() {
    source = result = null; decisions = {}; reviews = [];
    $('result').hidden = $('review-panel').hidden = $('issue-panel').hidden = true;
    for (const id of ['preview','issues','reviews']) $(id).replaceChildren();
    $('counts').textContent = $('totals').textContent = '';
    controls();
  }
  function showIssues() {
    $('issues').replaceChildren();
    for (const message of [...result.errors,...result.warnings]) {
      const li = document.createElement('li'); li.textContent = message; $('issues').append(li);
    }
    $('issue-panel').hidden = !result.errors.length && !result.warnings.length;
    $('issue-panel').open = Boolean(result.errors.length);
    $('issue-title').textContent = result.errors.length ? `需要修正（${result.errors.length} 项）` : `需要留意（${result.warnings.length} 项）`;
  }
  function showReviews() {
    $('reviews').replaceChildren(); $('review-panel').hidden = !reviews.length;
    $('review-title').textContent = result.unresolved.length ? `分类核对（待确认 ${result.unresolved.length} 项）` : '分类已确认';
    for (const item of reviews) {
      const card = document.createElement('div'); card.className = 'review';
      const line = document.createElement('p');
      line.textContent = `${item.supplier}\n${item.source}\n原分类：${item.classification || '空白'}；谈判明细金额合计：${item.amountCents == null ? '金额异常，需核对源表' : '￥'+money(item.amountCents)}`;
      const label = document.createElement('label'); label.textContent = '是否谈判分批';
      const select = document.createElement('select'); select.id = 'classification-' + item.id; label.htmlFor = select.id;
      select.dataset.address = item.address;
      for (const [value,text] of [['','请选择'],['是','是 · 谈判付款'],['否','否 · 正常付款']]) {
        const option = document.createElement('option'); option.value = value; option.textContent = text; select.append(option);
      }
      select.value = decisions[item.id] || '';
      select.addEventListener('change',() => {
        if (select.value) decisions[item.id] = select.value; else delete decisions[item.id];
        calculate();
        document.getElementById(select.id)?.focus();
      });
      card.append(line,label,select); $('reviews').append(card);
    }
  }
  function preview() {
    $('preview').replaceChildren(); $('result').hidden = !result.dates.length || !result.rows.length;
    if ($('result').hidden) return;
    const table = document.createElement('table'), head = table.createTHead(), first = head.insertRow(), second = head.insertRow(), body = table.createTBody();
    const add = (row,label,rowSpan=1,colSpan=1) => {
      const th = document.createElement('th'); th.textContent = label; th.rowSpan = rowSpan; th.colSpan = colSpan;
      th.scope = colSpan > 1 ? 'colgroup' : 'col'; row.append(th);
    };
    add(first,'供应商',2); add(first,'谈判总金额',2); add(first,'谈判付款节奏',1,result.dates.length);
    add(first,'正常付款总金额',2); add(first,'正常付款',1,result.dates.length);
    for (let block=0;block<2;block++) for (const date of result.dates) add(second,date.label);
    for (const item of result.rows) {
      const row = body.insertRow(); row.insertCell().textContent = item.supplier;
      for (const value of [item.negotiatedTotalCents,...item.negotiatedCents,item.normalTotalCents,...item.normalCents]) row.insertCell().textContent = money(value);
    }
    $('preview').append(table);
    $('counts').textContent = `（${result.recordCount} 条明细，${result.rows.length} 家供应商）`;
    $('totals').textContent = `${result.canExport ? '' : '已识别金额：'}谈判付款 ￥${money(result.totals.negotiatedCents)}；正常付款 ￥${money(result.totals.normalCents)}` + (result.unresolved.length ? `；待分类 ￥${money(result.totals.unresolvedCents)}` : '');
  }
  function calculate() {
    result = C.build(source,decisions);
    showIssues(); showReviews(); preview(); controls();
    $('status').textContent = result.errors.length ? '整理未完成，请修正来源问题后重新上传。' : result.unresolved.length ? '请先确认下方分类，再下载完整汇总。' : '汇总完成，可下载。';
  }
  async function readBook(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    if (!zip.file('xl/workbook.xml')) throw Error('请上传有效的 .xlsx 工作簿');
    // Preserve the missing-cache signal that SheetJS otherwise drops for string formulas.
    let changed = false;
    for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
      const xml = await zip.file(name).async('string');
      const next = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,(cell,attrs,content) => {
        if (!content || !/\bt="str"/.test(attrs) || !/<f(?:\s|>|\/)/.test(content) || /<v\b[^>]*>[^<]+<\/v>/.test(content)) return cell;
        changed = true; return '<c' + attrs.replace(/\s+t="str"/,'') + '>' + content + '</c>';
      });
      if (next !== xml) zip.file(name,next);
    }
    return XLSX.read(changed ? await zip.generateAsync({type:'uint8array'}) : bytes,{type:'array',sheets:'谈判付款',cellFormula:true,cellNF:true,sheetStubs:true});
  }
  $('sourceFile').addEventListener('change',async () => {
    const file = $('sourceFile').files[0]; reset();
    if (!file) { $('reference').textContent = '尚未引用文件'; $('clear').hidden = true; $('status').textContent = '请上传文件。'; return; }
    busy = true; controls(); $('clear').hidden = false;
    $('reference').textContent = `${file.name}\n引用时间：${new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false})}`;
    $('status').textContent = '正在读取并汇总…';
    try {
      if (!/\.xlsx$/i.test(file.name)) throw Error('请上传 .xlsx 文件');
      if (file.size > 60 * 1024 * 1024) throw Error('文件超过 60 MB，请缩小文件后上传');
      source = {book:await readBook(await file.arrayBuffer()),filename:file.name};
      result = C.build(source); reviews = result.unresolved;
      calculate();
    } catch (error) { reset(); $('status').textContent = '读取失败：' + error.message; }
    finally { busy = false; controls(); }
  });
  $('clear').addEventListener('click',() => {
    reset(); $('sourceFile').value = ''; $('reference').textContent = '尚未引用文件';
    $('clear').hidden = true; $('status').textContent = '请上传文件。';
  });
  function verify(bytes) {
    const book = XLSX.read(bytes,{type:'array',cellNF:true}), sheet = book.Sheets['谈判付款整理'], n = result.dates.length;
    if (book.SheetNames.length !== 1 || !sheet || sheet['!ref'] !== `A1:${XLSX.utils.encode_col(2+2*n)}${result.rows.length+2}`) throw Error('导出结构校验未通过');
    if (sheet.A1?.v !== '供应商' || sheet.B1?.v !== '谈判总金额' || sheet[XLSX.utils.encode_cell({r:0,c:n+2})]?.v !== '正常付款总金额') throw Error('导出表头校验未通过');
    for (let i=0;i<n;i++) for (const col of [i+2,i+n+3]) {
      const cell = sheet[XLSX.utils.encode_cell({r:1,c:col})];
      if (cell?.t !== 's' || cell.v !== result.dates[i].label || cell.z !== '@') throw Error('导出付款日期校验未通过');
    }
    result.rows.forEach((row,index) => {
      const r = index+3, values = [row.negotiatedTotalCents,...row.negotiatedCents,row.normalTotalCents,...row.normalCents];
      if (sheet['A'+r]?.v !== row.supplier) throw Error('导出供应商校验未通过');
      values.forEach((cents,column) => {
        const cell = sheet[XLSX.utils.encode_cell({r:r-1,c:column+1})];
        if (cents == null ? cell?.v != null && cell.v !== '' : cell?.t !== 'n' || Math.round(cell.v*100) !== cents) throw Error('导出金额校验未通过：第'+r+'行');
      });
      if (sheet['B'+r]?.f !== `SUM(C${r}:${XLSX.utils.encode_col(n+1)}${r})` || sheet[XLSX.utils.encode_col(n+2)+r]?.f !== `SUM(${XLSX.utils.encode_col(n+3)}${r}:${XLSX.utils.encode_col(2*n+2)}${r})`) throw Error('导出合计公式校验未通过');
    });
  }
  function nextDownload() {
    const date = new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
    let versions = {};
    try {
      const saved = JSON.parse(localStorage.getItem(versionKey) || '{}');
      if (saved && typeof saved === 'object' && !Array.isArray(saved)) versions = saved;
    } catch (_) {}
    const stored = Number(versions[date]);
    const version = Math.max(Number.isSafeInteger(stored) && stored >= 0 ? stored : 0,memoryVersions[date] || 0) + 1;
    if (!Number.isSafeInteger(version)) throw Error('当天版本号已超出有效范围，请核实浏览器保存的版本记录');
    return {date,version,versions,name:`谈判付款整理-${date}-V${version}.xlsx`};
  }
  $('download').addEventListener('click',async () => {
    if (busy || !result?.canExport) return;
    busy = true; controls(); $('status').textContent = '正在生成并校验 Excel…'; let url;
    try {
      const bytes = await window.NegotiatedPaymentSummaryExport(result); verify(bytes);
      const {date,version,versions,name} = nextDownload();
      url = URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const a = document.createElement('a'); a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
      memoryVersions[date] = version;
      let saved = true;
      try { localStorage.setItem(versionKey,JSON.stringify({...versions,[date]:version})); } catch (_) { saved = false; }
      $('status').textContent = '已生成：' + name + (saved ? '' : '。浏览器未能保存版本号，本页内仍会递增。');
    } catch (error) { $('status').textContent = '导出失败：' + error.message; }
    finally { if (url) setTimeout(() => URL.revokeObjectURL(url),1000); busy = false; controls(); }
  });
  if (!C || !window.NegotiatedPaymentSummaryExport || !window.XLSX || !window.JSZip) { $('status').textContent = '工具加载失败，请刷新页面重试。'; return; }
  controls(); $('status').textContent = '请上传文件。';
})();
