(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.FundsLedgerCore;
  let result = null, filename = '', generation = 0, exporting = false;
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
  if (!C || !window.XLSX || !window.FundsLedgerExport || !window.JSZip) {
    $('status').textContent = '工具加载失败，请刷新页面后重试。';
    return;
  }
  $('file').disabled = false;
  $('status').textContent = '请选择采购付款台账。';
  $('file').addEventListener('change', async () => {
    const file = $('file').files[0];
    if (!file || exporting) return;
    const current = ++generation;
    clear();
    filename = file.name;
    $('reference').textContent = `${file.name}\n引用时间：${new Date().toLocaleString('zh-CN', {timeZone:'Asia/Shanghai', hour12:false})}`;
    $('status').textContent = '正在读取并整理…';
    try {
      if (!/\.xlsx?$/i.test(file.name)) throw Error('请选择 .xlsx 或 .xls 文件');
      if (file.size > 60 * 1024 * 1024) throw Error('文件超过 60 MB，请缩小后重新上传');
      const bytes = await file.arrayBuffer();
      if (current !== generation) return;
      const book = await readBook(bytes);
      if (current !== generation) return;
      result = C.build(book, file.name);
      render();
    } catch (error) {
      if (current !== generation) return;
      clear();
      $('status').textContent = '整理失败：' + error.message;
    } finally {
      // Reset only this completed upload, never a newer file selection.
      if (current === generation) $('file').value = '';
    }
  });
  $('download').addEventListener('click', async () => {
    if (!result || exporting) return;
    exporting = true;
    $('download').disabled = true;
    $('file').disabled = true;
    const status = $('status').textContent;
    $('status').textContent = '正在生成 Excel…';
    try {
      const bytes = await FundsLedgerExport(result);
      const url = URL.createObjectURL(new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link = document.createElement('a');
      link.href = url;
      link.download = C.exportFilename(filename);
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      $('status').textContent = status;
    } catch (error) {
      $('status').textContent = '导出失败：' + error.message + '。可以重试下载。';
    } finally {
      exporting = false;
      $('download').disabled = !result;
      $('file').disabled = false;
    }
  });
  window.FundsLedgerToolReady = true;
})();
