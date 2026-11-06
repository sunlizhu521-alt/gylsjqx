(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.FundsLedgerCore;
  let result = null, exporting = false;
  const sources = {
    ledger: {input:'file', reference:'reference', label:'采购付款台账'},
    applications: {input:'applicationsFile', reference:'applicationsReference', label:'付款申请汇总'},
    progress: {input:'progressFile', reference:'progressReference', label:'资金整理'}
  };
  for (const source of Object.values(sources)) Object.assign(source, {book:null, filename:'', state:'empty', version:0, error:''});
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  async function readBook(bytes) {
    // SheetJS treats a string formula with no <v> as a valid empty string.
    // Keep a missing cache distinguishable from an explicit empty result.
    const signature = new Uint8Array(bytes, 0, Math.min(2, bytes.byteLength));
    if (signature[0] === 0x50 && signature[1] === 0x4b) {
      const zip = await JSZip.loadAsync(bytes);
      let changed = false;
      for (const path of Object.keys(zip.files).filter(p => /^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
        const xml = await zip.file(path).async('string');
        const next = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (cell, attrs, content) => {
          if (content === undefined || !/\bt="str"/.test(attrs) || !/<f(?:\s|>|\/)/.test(content) || /<v(?:\s|>|\/)/.test(content)) return cell;
          changed = true;
          return '<c' + attrs.replace(/\s+t="str"/, '') + '>' + content + '</c>';
        });
        if (next !== xml) zip.file(path, next);
      }
      if (changed) bytes = await zip.generateAsync({type:'uint8array'});
    }
    return XLSX.read(bytes, {type:'array', cellNF:true, cellFormula:true, cellDates:false, sheetStubs:true});
  }

  function clear() {
    result = null;
    $('result').hidden = true;
    $('download').disabled = true;
    $('warning-panel').hidden = true;
    $('warnings').replaceChildren();
    for (const id of ['applied', 'unapplied']) {
      $(id + '-count').textContent = '0';
      $(id + '-preview').replaceChildren();
    }
  }
  function render() {
    result.sheets.forEach((sheet, index) => {
      const id = index === 0 ? 'applied' : 'unapplied';
      $(id + '-count').textContent = sheet.rows.length;
      $(id + '-preview').innerHTML = '<table><thead><tr>' + sheet.headers.map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + sheet.rows.map(row => '<tr>' + row.map(cell => '<td>' + esc(C.formatCell(cell)) + '</td>').join('') + '</tr>').join('') + '</tbody></table>' + (sheet.rows.length ? '' : '<p class="empty">没有符合条件的记录，下载仍保留本表表头。</p>');
    });
    $('warnings').innerHTML = result.warnings.map(w => '<li>' + esc(w) + '</li>').join('');
    $('warning-count').textContent = result.warnings.length;
    $('warning-panel').hidden = !result.warnings.length;
    $('result').hidden = false;
    $('download').disabled = false;
    $('status').textContent = `整理完成：已申请 OA ${result.sheets[0].rows.length} 条，未申请 OA ${result.sheets[1].rows.length} 条。`;
  }
  function controls() {
    for (const [role, source] of Object.entries(sources)) {
      $(source.input).disabled = exporting;
      const button = $('clear-' + role);
      if (button) { button.hidden = source.state === 'empty'; button.disabled = exporting; }
    }
    $('download').disabled = exporting || !result;
  }
  function refresh() {
    clear();
    const reading = Object.values(sources).filter(source => source.state === 'reading');
    const failed = Object.values(sources).filter(source => source.state === 'error');
    if (reading.length) $('status').textContent = '正在读取：' + reading.map(source => source.label).join('、') + '…';
    else if (failed.length) $('status').textContent = failed.map(source => source.label + '读取失败：' + source.error).join('\n');
    else if (!sources.ledger.book) $('status').textContent = '请选择采购付款台账；两个补充文件可选。';
    else {
      try {
        const optional = {};
        for (const role of ['applications', 'progress']) if (sources[role].book) optional[role] = {book:sources[role].book, filename:sources[role].filename};
        result = C.build(sources.ledger.book, sources.ledger.filename, optional);
        render();
      } catch (error) { clear(); $('status').textContent = '整理失败：' + error.message; }
    }
    controls();
  }
  if (!C || !window.XLSX || !window.FundsLedgerExport || !window.JSZip) {
    $('status').textContent = '工具加载失败，请刷新页面后重试。';
    return;
  }
  for (const [role, source] of Object.entries(sources)) {
    $(source.input).addEventListener('change', async () => {
      const file = $(source.input).files[0];
      if (!file || exporting) return;
      const current = ++source.version;
      source.book = null;
      source.filename = file.name;
      source.state = 'reading';
      source.error = '';
      $(source.reference).textContent = `${file.name}\n引用时间：${new Date().toLocaleString('zh-CN', {timeZone:'Asia/Shanghai', hour12:false})}`;
      refresh();
      try {
        if (!/\.xlsx?$/i.test(file.name)) throw Error('请选择 .xlsx 或 .xls 文件');
        if (file.size > 60 * 1024 * 1024) throw Error('文件超过 60 MB，请缩小后重新上传');
        const bytes = await file.arrayBuffer();
        if (current !== source.version) return;
        const book = await readBook(bytes);
        if (current !== source.version) return;
        if (role === 'ledger') C.build(book, file.name);
        else C.validateSupplement(book, role, file.name);
        source.book = book;
        source.state = 'ready';
      } catch (error) {
        if (current !== source.version) return;
        source.book = null;
        source.state = 'error';
        source.error = error.message;
        $(source.reference).textContent += '\n读取失败，请重新选择文件' + (role === 'ledger' ? '。' : '或移除此补充文件。');
      } finally {
        // Each source owns its generation; another upload must not cancel it.
        if (current === source.version) { $(source.input).value = ''; refresh(); }
      }
    });
    const remove = $('clear-' + role);
    if (remove) remove.addEventListener('click', () => {
      if (exporting) return;
      source.version++;
      Object.assign(source, {book:null, filename:'', state:'empty', error:''});
      $(source.input).value = '';
      $(source.reference).textContent = '尚未引用（可选）';
      refresh();
    });
  }
  refresh();
  $('download').addEventListener('click', async () => {
    if (!result || exporting) return;
    exporting = true;
    controls();
    const status = $('status').textContent;
    $('status').textContent = '正在生成 Excel…';
    try {
      const bytes = await FundsLedgerExport(result);
      const url = URL.createObjectURL(new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link = document.createElement('a');
      link.href = url;
      link.download = C.exportFilename(sources.ledger.filename);
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      $('status').textContent = status;
    } catch (error) {
      $('status').textContent = '导出失败：' + error.message + '。可以重试下载。';
    } finally {
      exporting = false;
      controls();
    }
  });
  window.FundsLedgerToolReady = true;
})();
