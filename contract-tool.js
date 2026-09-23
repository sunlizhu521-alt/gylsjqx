(function () {
  'use strict';

  const C = ContractCore;
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const MIME = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  };
  const state = {
    orderFile: null, orderBytes: null, orderBook: null, orderSheet: '', order: null,
    templateFile: null, templateBytes: null, templateType: '', templateBook: null, templateSheet: '', templateModel: null, templateFeatures: [], fingerprint: '',
    mappings: {}, detailRow: '', selectedTarget: '', output: null, outputFileName: '', busy: false,
  };
  const $ = id => document.getElementById(id);
  const esc = value => C.text(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const direct = (element, name, namespace = W) => [...(element?.children || [])].filter(child => child.namespaceURI === namespace && child.localName === name);
  const all = (element, name, namespace = W) => [...(element?.getElementsByTagNameNS(namespace, name) || [])];

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    initTheme();
    $('orderFile').addEventListener('change', event => loadOrder(event.target.files[0]));
    $('templateFile').addEventListener('change', event => loadTemplate(event.target.files[0]));
    $('orderSheet').addEventListener('change', event => { state.orderSheet = event.target.value; analyzeOrder(); invalidateOutput(); updateAll(); });
    $('templateSheet').addEventListener('change', event => { state.templateSheet = event.target.value; buildExcelModel(); state.mappings = {}; state.detailRow = ''; state.selectedTarget = ''; invalidateOutput(); restoreMapping(); updateAll(); });
    $('outputName').addEventListener('input', () => { invalidateOutput(); updateConfirmation(); });
    $('confirmGenerate').addEventListener('change', updateConfirmation);
    $('generateContract').addEventListener('click', generate);
    $('downloadContract').addEventListener('click', downloadOutput);
    $('clearContract').addEventListener('click', () => location.reload());
    $('exportMapping').addEventListener('click', exportMapping);
    $('importMapping').addEventListener('change', event => importMapping(event.target.files[0]));
    for (const [dropId, inputId] of [['orderDrop', 'orderFile'], ['templateDrop', 'templateFile']]) bindDrop($(dropId), $(inputId));
    updateAll();
  }

  function initTheme() {
    if (localStorage.getItem('dpc-theme') === 'dark') document.documentElement.dataset.theme = 'dark';
    $('themeToggle').addEventListener('click', () => {
      const dark = document.documentElement.dataset.theme !== 'dark';
      if (dark) document.documentElement.dataset.theme = 'dark';
      else delete document.documentElement.dataset.theme;
      localStorage.setItem('dpc-theme', dark ? 'dark' : 'light');
    });
  }

  function bindDrop(drop, input) {
    for (const eventName of ['dragenter', 'dragover']) drop.addEventListener(eventName, event => { event.preventDefault(); drop.classList.add('dragging'); });
    for (const eventName of ['dragleave', 'drop']) drop.addEventListener(eventName, event => { event.preventDefault(); drop.classList.remove('dragging'); });
    drop.addEventListener('drop', event => {
      const file = event.dataTransfer.files[0];
      if (!file) return;
      const transfer = new DataTransfer(); transfer.items.add(file); input.files = transfer.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  function status(message, kind = '') {
    $('contractStatus').textContent = message;
    $('contractStatus').className = `contract-status ${kind}`.trim();
  }

  function toast(message, kind = '') {
    const element = document.createElement('div');
    element.className = `toast ${kind}`.trim(); element.textContent = message; $('toastArea').append(element);
    setTimeout(() => element.remove(), 3600);
  }

  function invalidateOutput() {
    state.output = null; state.outputFileName = '';
    $('resultStage').hidden = true; $('downloadContract').disabled = true; $('generatedPreview').replaceChildren();
  }

  async function loadOrder(file) {
    if (!file) return;
    invalidateOutput(); status('正在读取订单明细…');
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('订单明细必须是 .xlsx 或 .xls 文件');
      if (file.size > 20 * 1024 * 1024) throw new Error('订单明细不能超过 20 MB');
      const bytes = await file.arrayBuffer();
      const book = XLSX.read(bytes, { type: 'array', cellDates: false, cellFormula: true });
      if (!book.SheetNames.length) throw new Error('订单文件中没有工作表');
      state.orderFile = file; state.orderBytes = bytes; state.orderBook = book; state.orderSheet = book.SheetNames[0];
      setOptions($('orderSheet'), book.SheetNames, state.orderSheet); $('orderSheet').disabled = false;
      analyzeOrder();
      $('orderFileName').textContent = file.name; $('orderDrop').classList.add('loaded');
      status(`订单读取完成：${state.order.rows.length} 行，${state.order.headers.length} 个字段`, 'success');
      updateAll();
    } catch (error) {
      state.orderFile = null; state.orderBook = null; state.order = null; $('orderDrop').classList.remove('loaded');
      status(error.message, 'error'); updateAll();
    }
  }

  function analyzeOrder() {
    const sheet = state.orderBook?.Sheets[state.orderSheet];
    if (!sheet) throw new Error('找不到选中的订单工作表');
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: true });
    state.order = C.analyzeMatrix(matrix, 1000);
    const allowed = new Set(state.order.headers);
    for (const [target, mapping] of Object.entries(state.mappings)) if (!allowed.has(mapping.field)) delete state.mappings[target];
  }

  async function loadTemplate(file) {
    if (!file) return;
    invalidateOutput(); status('正在检查合同模板…');
    try {
      if (!/\.(docx|xlsx)$/i.test(file.name)) throw new Error('合同模板必须是 .docx 或 .xlsx；旧式 .xls 请先另存为 .xlsx');
      if (file.size > 30 * 1024 * 1024) throw new Error('合同模板不能超过 30 MB');
      const bytes = await file.arrayBuffer();
      const type = file.name.toLowerCase().endsWith('.docx') ? 'docx' : 'xlsx';
      const zip = await validateOoxml(bytes, type);
      state.templateFile = file; state.templateBytes = bytes; state.templateType = type; state.fingerprint = await sha256(bytes);
      state.templateBook = null; state.templateSheet = ''; state.mappings = {}; state.detailRow = ''; state.selectedTarget = '';
      if (type === 'docx') await parseDocx(zip);
      else parseXlsx();
      $('templateFileName').textContent = file.name; $('templateDrop').classList.add('loaded');
      if (!$('outputName').value.trim()) $('outputName').value = file.name.replace(/\.[^.]+$/, '') + '-生成合同';
      restoreMapping();
      const featureNote = state.templateFeatures.length ? `；已识别并保留${state.templateFeatures.join('、')}` : '';
      status(`模板检查通过：${type === 'docx' ? 'Word' : 'Excel'}，指纹 ${state.fingerprint.slice(0, 12)}${featureNote}`, 'success');
      updateAll();
    } catch (error) {
      state.templateFile = null; state.templateBytes = null; state.templateModel = null; state.templateType = ''; state.templateFeatures = []; $('templateDrop').classList.remove('loaded');
      status(error.message, 'error'); updateAll();
    }
  }

  async function validateOoxml(bytes, type) {
    let zip;
    try { zip = await JSZip.loadAsync(bytes); }
    catch (_) { throw new Error('模板文件损坏或受到密码保护，无法读取'); }
    const names = Object.keys(zip.files);
    const expanded = names.reduce((total, name) => total + Number(zip.files[name]._data?.uncompressedSize || 0), 0);
    if (expanded > 120 * 1024 * 1024) throw new Error('模板解压后过大，已停止处理');
    if (names.some(name => /vbaProject\.bin|macrosheets|xl\/externalLinks\//i.test(name))) throw new Error('模板包含宏或外部链接，不能生成');
    state.templateFeatures = [];
    if (type === 'docx' && !zip.file('word/document.xml')) throw new Error('请选择有效的 Word .docx 模板');
    if (type === 'xlsx') {
      if (!zip.file('xl/workbook.xml')) throw new Error('请选择有效的 Excel .xlsx 模板');
      if (names.some(name => /^xl\/drawings\//i.test(name))) state.templateFeatures.push('图形');
      if (names.some(name => /^xl\/pivot/i.test(name))) state.templateFeatures.push('数据透视表');
      if (names.some(name => /^xl\/tables\//i.test(name))) state.templateFeatures.push('结构化表');
      const workbookXml = await zip.file('xl/workbook.xml').async('string');
      if (/<workbookProtection\b/i.test(workbookXml)) throw new Error('Excel模板受到工作簿保护，不能生成');
      for (const name of names.filter(item => /^xl\/worksheets\/.*\.xml$/i.test(item))) {
        if (/<sheetProtection\b/i.test(await zip.file(name).async('string'))) throw new Error('Excel模板包含受保护工作表，不能生成');
      }
    }
    for (const name of names.filter(item => item.endsWith('.rels'))) {
      if (/TargetMode\s*=\s*["']External["']/i.test(await zip.file(name).async('string'))) throw new Error('模板包含外部链接，不能生成');
    }
    return zip;
  }

  async function parseDocx(zip) {
    const xml = await zip.file('word/document.xml').async('string');
    if (xml.length > 5_000_000) throw new Error('Word模板正文结构过大');
    const doc = parseXml(xml), body = all(doc, 'body')[0];
    if (!body) throw new Error('Word模板缺少正文');
    if (['drawing', 'pict', 'altChunk', 'sdt', 'fldChar', 'fldSimple', 'ins', 'del'].some(name => all(body, name).length)) {
      throw new Error('Word模板含文本框、图片、域或修订内容，浏览器版暂不支持');
    }
    if (all(body, 'tbl').some(table => table.parentElement?.localName !== 'body')) throw new Error('Word模板包含嵌套表格，浏览器版暂不支持');
    state.templateModel = buildWordModel(doc); $('templateSheetLabel').hidden = true;
  }

  function buildWordModel(doc) {
    const body = all(doc, 'body')[0], blocks = [], targets = [];
    let paragraphIndex = 0, tableIndex = 0;
    for (const child of body.children) {
      if (child.namespaceURI !== W) continue;
      if (child.localName === 'p') {
        const target = { id: `p:${paragraphIndex}`, kind: 'paragraph', label: `正文段落 ${paragraphIndex + 1}`, value: wordText(child), rowKey: '' };
        targets.push(target); blocks.push({ type: 'paragraph', target }); paragraphIndex += 1;
      } else if (child.localName === 'tbl') {
        const rows = direct(child, 'tr').map((row, rowIndex) => {
          const rowKey = `word:${tableIndex}:${rowIndex}`;
          const cells = direct(row, 'tc').map((cell, cellIndex) => {
            const target = { id: `t:${tableIndex}:r:${rowIndex}:c:${cellIndex}`, kind: 'word-cell', label: `表格 ${tableIndex + 1} 第 ${rowIndex + 1} 行第 ${cellIndex + 1} 格`, value: wordCellText(cell), rowKey, rowIndex, cellIndex, tableIndex };
            targets.push(target); return target;
          });
          return { rowKey, rowIndex, cells };
        });
        blocks.push({ type: 'table', tableIndex, rows }); tableIndex += 1;
      }
      if (targets.length > 1000) throw new Error('Word模板可映射位置超过 1,000 个，请简化模板');
    }
    return { type: 'docx', blocks, targets };
  }

  function parseXlsx() {
    const book = XLSX.read(state.templateBytes, { type: 'array', cellStyles: true, cellFormula: true, cellNF: true });
    if (!book.SheetNames.length) throw new Error('Excel模板中没有工作表');
    state.templateBook = book; state.templateSheet = book.SheetNames[0];
    setOptions($('templateSheet'), book.SheetNames, state.templateSheet); $('templateSheet').disabled = false; $('templateSheetLabel').hidden = false;
    buildExcelModel();
  }

  function buildExcelModel() {
    const sheet = state.templateBook?.Sheets[state.templateSheet];
    const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 20, c: 8 } };
    const rows = Math.max(range.e.r + 1, 12), columns = Math.max(range.e.c + 1, 6);
    if (rows > 100 || columns > 40) throw new Error('Excel模板范围超过 100 行或 40 列，请移除无关区域后再上传');
    const modelRows = [], targets = [];
    for (let r = 0; r < rows; r += 1) {
      const rowKey = `excel:${r + 1}`, cells = [];
      for (let c = 0; c < columns; c += 1) {
        const address = XLSX.utils.encode_cell({ r, c }), cell = sheet?.[address];
        const target = { id: `x:${address}`, kind: 'excel-cell', label: `${state.templateSheet}!${address}`, value: cell?.w ?? C.text(cell?.v), rowKey, rowIndex: r, cellIndex: c, address };
        targets.push(target); cells.push(target);
      }
      modelRows.push({ rowKey, rowIndex: r, cells });
    }
    state.templateModel = { type: 'xlsx', rows: modelRows, targets, columns };
  }

  function setOptions(select, values, selected) {
    select.innerHTML = values.map(value => `<option ${value === selected ? 'selected' : ''}>${esc(value)}</option>`).join('');
  }

  function updateAll() {
    const ready = !!(state.order && state.templateModel);
    $('mappingStage').hidden = !ready; $('confirmStage').hidden = !ready;
    $('orderMeta').textContent = state.order ? `${state.order.rows.length} 行 · ${state.order.headers.length} 个字段` : '尚未读取订单';
    $('orderFields').innerHTML = state.order ? state.order.headers.map(header => {
      const sample = C.distinctValues(state.order.rows.slice(0, 20), header).slice(0, 2).join('、');
      return `<div class="field-item">${esc(header)}<small>${esc(sample || '前20行为空')}</small></div>`;
    }).join('') : '';
    if (ready) { renderTemplate(); renderInspector(); }
    const summary = [];
    if (state.orderFile) summary.push(`订单：${state.orderFile.name}（${state.order.rows.length}行）`);
    if (state.templateFile) summary.push(`模板：${state.templateFile.name}`);
    $('fileSummary').textContent = summary.join('；') || '等待上传订单和模板';
    $('templateMeta').textContent = state.templateFile ? `${state.templateFile.name}${state.templateSheet ? ` · ${state.templateSheet}` : ''}` : '';
    $('mappingStatus').textContent = `${mappedEntries().length} 个映射`;
    updateConfirmation(); updateSteps();
  }

  function renderTemplate() {
    const mapped = new Set(mappedEntries().map(([target]) => target));
    if (state.templateModel.type === 'docx') {
      $('templatePreview').innerHTML = `<div class="word-page">${state.templateModel.blocks.map(block => {
        if (block.type === 'paragraph') return targetButton(block.target, mapped, 'word-paragraph');
        return `<table class="word-table"><tbody>${block.rows.map(row => `<tr class="word-row ${state.detailRow === row.rowKey ? 'detail-row' : ''}">${row.cells.map(cell => `<td>${targetButton(cell, mapped)}</td>`).join('')}<td class="word-row-marker"><button type="button" data-detail-row="${esc(row.rowKey)}">${state.detailRow === row.rowKey ? '取消明细行' : '设为明细行'}</button></td></tr>`).join('')}</tbody></table>`;
      }).join('')}</div>`;
    } else {
      const letters = Array.from({ length: state.templateModel.columns }, (_, index) => XLSX.utils.encode_col(index));
      $('templatePreview').innerHTML = `<table class="excel-preview"><thead><tr><th class="row-number"></th>${letters.map(letter => `<th>${letter}</th>`).join('')}</tr></thead><tbody>${state.templateModel.rows.map(row => `<tr class="${state.detailRow === row.rowKey ? 'detail-row' : ''}"><th><button type="button" data-detail-row="${esc(row.rowKey)}" title="设为明细模板行">${row.rowIndex + 1}</button></th>${row.cells.map(cell => `<td>${targetButton(cell, mapped)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    }
    $('templatePreview').querySelectorAll('[data-target]').forEach(button => button.addEventListener('click', () => { state.selectedTarget = button.dataset.target; renderTemplate(); renderInspector(); }));
    $('templatePreview').querySelectorAll('[data-detail-row]').forEach(button => button.addEventListener('click', () => toggleDetailRow(button.dataset.detailRow)));
  }

  function targetButton(target, mapped, extra = '') {
    const value = target.value ? esc(target.value) : '<span style="color:#98a2b3">空白位置</span>';
    const classes = ['template-target', extra, state.selectedTarget === target.id ? 'selected' : '', mapped.has(target.id) ? 'mapped' : ''].filter(Boolean).join(' ');
    return `<button type="button" class="${classes}" data-target="${esc(target.id)}" title="${esc(target.label)}">${value}</button>`;
  }

  function toggleDetailRow(rowKey) {
    const previous = state.detailRow;
    state.detailRow = previous === rowKey ? '' : rowKey;
    for (const [targetId, mapping] of Object.entries(state.mappings)) {
      const target = findTarget(targetId);
      if (!target) continue;
      if (target.rowKey === previous) mapping.mode = 'single';
      if (target.rowKey === state.detailRow) mapping.mode = 'detail';
    }
    invalidateOutput(); saveMapping(); updateAll();
  }

  function renderInspector() {
    const host = $('mappingInspector'), target = findTarget(state.selectedTarget);
    if (!target) { host.className = 'mapping-empty'; host.textContent = '请点击模板中的段落、表格单元格或Excel单元格。'; return; }
    host.className = '';
    const mapping = state.mappings[target.id] || { field: '', mode: target.rowKey === state.detailRow ? 'detail' : 'single', strategy: '' };
    const values = mapping.field ? C.distinctValues(state.order.rows, mapping.field) : [];
    const conflict = mapping.mode !== 'detail' && values.length > 1;
    host.innerHTML = `<div class="mapping-form">
      <div class="mapping-location"><strong>${esc(target.label)}</strong>${esc(target.value || '当前为空白位置')}</div>
      <label>订单字段<select id="mapField"><option value="">不映射</option>${state.order.headers.map(header => `<option value="${esc(header)}" ${header === mapping.field ? 'selected' : ''}>${esc(header)}</option>`).join('')}</select></label>
      <label>写入方式<select id="mapMode" ${target.rowKey ? '' : 'disabled'}><option value="single" ${mapping.mode !== 'detail' ? 'selected' : ''}>合同单值</option><option value="detail" ${mapping.mode === 'detail' ? 'selected' : ''}>订单明细（需设为明细行）</option></select></label>
      ${conflict ? `<div class="mapping-conflict">${esc(mapping.field)} 有 ${values.length} 个不同值：${esc(values.slice(0, 4).join('、'))}${values.length > 4 ? '…' : ''}</div><label>多值处理<select id="mapStrategy"><option value="">请选择</option><option value="first" ${mapping.strategy === 'first' ? 'selected' : ''}>指定取第一条非空值</option><option value="merge" ${mapping.strategy === 'merge' ? 'selected' : ''}>合并去重值</option><option value="manual" ${mapping.strategy === 'manual' ? 'selected' : ''}>手工输入</option><option value="sum" ${mapping.strategy === 'sum' ? 'selected' : ''}>明确求和</option></select></label>${mapping.strategy === 'manual' ? `<label>手工值<input id="mapManual" value="${esc(mapping.manual || '')}" /></label>` : ''}` : ''}
      ${target.rowKey ? `<button class="secondary-button mapping-row-button" id="detailRowButton" type="button">${state.detailRow === target.rowKey ? '取消当前明细模板行' : '将当前行设为明细模板行'}</button>` : ''}
      ${mapping.field ? '<button class="ghost-button" id="removeMapping" type="button">移除此映射</button>' : ''}
    </div>`;
    $('mapField').addEventListener('change', event => changeMapping(target, { field: event.target.value }));
    $('mapMode').addEventListener('change', event => {
      if (event.target.value === 'detail' && target.rowKey !== state.detailRow) toggleDetailRow(target.rowKey);
      else changeMapping(target, { mode: event.target.value });
    });
    $('mapStrategy')?.addEventListener('change', event => changeMapping(target, { strategy: event.target.value, manual: '' }));
    $('mapManual')?.addEventListener('input', event => changeMapping(target, { manual: event.target.value }, false));
    $('detailRowButton')?.addEventListener('click', () => toggleDetailRow(target.rowKey));
    $('removeMapping')?.addEventListener('click', () => { delete state.mappings[target.id]; invalidateOutput(); saveMapping(); updateAll(); });
  }

  function changeMapping(target, changes, fullRender = true) {
    const current = state.mappings[target.id] || { field: '', mode: target.rowKey === state.detailRow ? 'detail' : 'single', strategy: '' };
    const next = { ...current, ...changes };
    if (!next.field) delete state.mappings[target.id];
    else {
      const values = C.distinctValues(state.order.rows, next.field);
      if (!('strategy' in changes) && values.length <= 1) next.strategy = 'first';
      if (!('strategy' in changes) && values.length > 1 && current.field !== next.field) next.strategy = '';
      if (target.rowKey === state.detailRow) next.mode = 'detail';
      state.mappings[target.id] = next;
    }
    invalidateOutput(); saveMapping();
    if (fullRender) updateAll(); else updateConfirmation();
  }

  function mappedEntries() { return Object.entries(state.mappings).filter(([, mapping]) => mapping?.field); }
  function findTarget(id) { return state.templateModel?.targets.find(target => target.id === id); }

  function updateConfirmation() {
    if (!state.order || !state.templateModel) return;
    const issues = C.mappingIssues(state.order.rows, state.mappings, state.detailRow);
    if (state.detailRow && !mappedEntries().some(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow)) issues.push('明细模板行还没有映射任何订单字段');
    const outputName = C.sanitizeFileName($('outputName').value);
    if (!$('outputName').value.trim()) issues.push('请填写合同文件名称');
    $('confirmSummary').innerHTML = `订单：<strong>${esc(state.orderSheet)}</strong>，${state.order.rows.length} 行；模板：<strong>${esc(state.templateFile.name)}</strong>${state.templateSheet ? `，合同Sheet：<strong>${esc(state.templateSheet)}</strong>` : ''}；输出：<strong>${esc(outputName)}.${state.templateType}</strong>`;
    $('contractIssues').innerHTML = issues.length ? `<ul>${issues.map(issue => `<li>${esc(issue)}</li>`).join('')}</ul>` : '<div class="ready">映射检查通过，可以人工确认并生成。</div>';
    $('generateContract').disabled = state.busy || issues.length > 0 || !$('confirmGenerate').checked;
  }

  function updateSteps() {
    const steps = [...$('contractSteps').children];
    const mapped = mappedEntries().length > 0;
    const values = [!!state.order, !!state.templateModel, mapped, mapped && $('confirmGenerate').checked, !!state.output];
    let active = values.findIndex(value => !value); if (active < 0) active = 4;
    steps.forEach((step, index) => { step.classList.toggle('done', values[index] && index < active); step.classList.toggle('active', index === active); });
  }

  async function generate() {
    updateConfirmation(); if ($('generateContract').disabled) return;
    state.busy = true; $('generateContract').textContent = '正在生成…'; updateConfirmation(); status('正在生成合同副本…');
    try {
      const blob = state.templateType === 'docx' ? await generateDocx() : await generateXlsx();
      state.output = blob; state.outputFileName = `${C.sanitizeFileName($('outputName').value)}.${state.templateType}`;
      renderGeneratedPreview(); $('resultStage').hidden = false; $('downloadContract').disabled = false; $('downloadContract').textContent = `下载 ${state.templateType.toUpperCase()}`;
      saveMapping(); status(`合同已生成：${state.outputFileName}`, 'success'); toast('合同生成完成，请预览后下载', 'success'); updateSteps();
      $('resultStage').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) { status(error.message, 'error'); toast(error.message, 'error'); }
    finally { state.busy = false; $('generateContract').textContent = '确认生成'; updateConfirmation(); }
  }

  async function generateDocx() {
    const zip = await JSZip.loadAsync(state.templateBytes), xml = await zip.file('word/document.xml').async('string'), doc = parseXml(xml);
    for (const [targetId, mapping] of mappedEntries()) {
      if (mapping.mode === 'detail') continue;
      const target = locateWordTarget(doc, targetId); if (!target) throw new Error(`模板位置已变化：${targetId}`);
      putWordText(target, resolveMapping(mapping));
    }
    if (state.detailRow) {
      const match = state.detailRow.match(/^word:(\d+):(\d+)$/); if (!match) throw new Error('Word明细模板行无效');
      const table = direct(all(doc, 'body')[0], 'tbl')[Number(match[1])], row = direct(table, 'tr')[Number(match[2])];
      if (!row) throw new Error('找不到Word明细模板行');
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      for (const record of state.order.rows) {
        const clone = row.cloneNode(true), cells = direct(clone, 'tc');
        for (const [targetId, mapping] of rowMappings) {
          const cellIndex = Number(targetId.match(/:c:(\d+)$/)?.[1]);
          if (cells[cellIndex]) putWordText(cells[cellIndex], C.text(record[mapping.field]));
        }
        row.parentNode.insertBefore(clone, row);
      }
      row.remove();
    }
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    return zip.generateAsync({ type: 'blob', mimeType: MIME.docx, compression: 'DEFLATE' });
  }

  async function generateXlsx() {
    const zip = await JSZip.loadAsync(state.templateBytes);
    const context = await locateXlsxSheet(zip, state.templateSheet);
    for (const [targetId, mapping] of mappedEntries()) {
      if (mapping.mode === 'detail') continue;
      writeSheetCell(context.sheetDoc, targetId.slice(2), resolveMapping(mapping));
    }
    if (state.detailRow) {
      const rowNumber = Number(state.detailRow.split(':')[1]);
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      await expandExcelDetailRowXml(zip, context, rowNumber, state.order.rows, rowMappings);
    }
    await forceWorkbookRecalculation(zip, context);
    zip.file(context.sheetPath, new XMLSerializer().serializeToString(context.sheetDoc));
    return zip.generateAsync({ type: 'blob', mimeType: MIME.xlsx, compression: 'DEFLATE' });
  }

  async function locateXlsxSheet(zip, sheetName) {
    const workbookPath = 'xl/workbook.xml', workbookRelsPath = 'xl/_rels/workbook.xml.rels';
    const workbookDoc = parseXml(await zip.file(workbookPath).async('string'));
    const workbookRelsDoc = parseXml(await zip.file(workbookRelsPath).async('string'));
    const sheet = all(workbookDoc, 'sheet', S).find(item => item.getAttribute('name') === sheetName);
    if (!sheet) throw new Error('找不到选中的合同工作表');
    const relationId = sheet.getAttributeNS(R, 'id') || sheet.getAttribute('r:id');
    const relation = all(workbookRelsDoc, 'Relationship', P).find(item => item.getAttribute('Id') === relationId);
    if (!relation) throw new Error('合同工作表关系损坏');
    const sheetPath = resolveZipPath('xl', relation.getAttribute('Target'));
    const sheetFile = zip.file(sheetPath); if (!sheetFile) throw new Error('合同工作表文件缺失');
    const sheetDoc = parseXml(await sheetFile.async('string'));
    return { zip, workbookPath, workbookRelsPath, workbookDoc, workbookRelsDoc, sheetPath, sheetDoc, sheetIndex: all(workbookDoc, 'sheet', S).indexOf(sheet) };
  }

  function resolveZipPath(base, target) {
    if (!target) return '';
    const parts = (target.startsWith('/') ? target.slice(1) : `${base}/${target}`).split('/'), output = [];
    for (const part of parts) {
      if (!part || part === '.') continue;
      if (part === '..') output.pop(); else output.push(part);
    }
    return output.join('/');
  }

  function sheetData(doc) {
    const data = all(doc, 'sheetData', S)[0];
    if (!data) throw new Error('Excel工作表缺少单元格数据');
    return data;
  }

  function rowNumberOf(row) { return Number(row.getAttribute('r') || 0); }
  function cellAddressOf(cell) { return cell.getAttribute('r') || ''; }

  function ensureSheetRow(doc, rowNumber) {
    const data = sheetData(doc), rows = direct(data, 'row', S);
    let row = rows.find(item => rowNumberOf(item) === rowNumber);
    if (row) return row;
    row = doc.createElementNS(S, 'row'); row.setAttribute('r', String(rowNumber));
    const next = rows.find(item => rowNumberOf(item) > rowNumber); data.insertBefore(row, next || null);
    return row;
  }

  function ensureRowCell(row, address) {
    let cell = direct(row, 'c', S).find(item => cellAddressOf(item) === address);
    if (cell) return cell;
    const point = XLSX.utils.decode_cell(address), doc = row.ownerDocument;
    cell = doc.createElementNS(S, 'c'); cell.setAttribute('r', address);
    const next = direct(row, 'c', S).find(item => XLSX.utils.decode_cell(cellAddressOf(item)).c > point.c);
    row.insertBefore(cell, next || null); return cell;
  }

  function writeSheetCell(doc, address, value) {
    const point = XLSX.utils.decode_cell(address), row = ensureSheetRow(doc, point.r + 1), cell = ensureRowCell(row, address);
    writeCellValue(cell, value);
  }

  function writeCellValue(cell, value) {
    const doc = cell.ownerDocument, originalType = cell.getAttribute('t');
    direct(cell, 'f', S).forEach(item => item.remove()); direct(cell, 'v', S).forEach(item => item.remove()); direct(cell, 'is', S).forEach(item => item.remove());
    const number = C.parseNumber(value), preferText = ['s', 'str', 'inlineStr'].includes(originalType);
    if (number !== null && !preferText) {
      cell.removeAttribute('t'); const node = doc.createElementNS(S, 'v'); node.textContent = String(number); cell.append(node); return;
    }
    cell.setAttribute('t', 'inlineStr');
    const inline = doc.createElementNS(S, 'is'), textNode = doc.createElementNS(S, 't');
    textNode.setAttribute('xml:space', 'preserve'); textNode.textContent = C.text(value); inline.append(textNode); cell.append(inline);
  }

  async function expandExcelDetailRowXml(zip, context, rowNumber, records, rowMappings) {
    const data = sheetData(context.sheetDoc), rows = direct(data, 'row', S), prototype = rows.find(row => rowNumberOf(row) === rowNumber);
    if (!prototype) throw new Error('找不到Excel明细模板行');
    const delta = records.length - 1;
    if (delta > 0) {
      const mergeRefs = all(context.sheetDoc, 'mergeCell', S).map(item => item.getAttribute('ref')).filter(Boolean);
      if (mergeRefs.some(ref => rangeTouchesRow(ref, rowNumber))) throw new Error('Excel明细模板行包含合并单元格，浏览器版不能安全扩展');
      const formulas = all(context.sheetDoc, 'f', S).filter(item => rowNumberFromAddress(cellAddressOf(item.parentElement)) >= rowNumber);
      if (formulas.length) throw new Error('Excel明细行或其下方含公式，扩展可能改变引用，请调整模板后再生成');
      const relationships = await loadSheetRelationships(zip, context.sheetPath);
      if (relationships.some(item => /\/pivotTable$/i.test(item.type))) throw new Error('当前合同Sheet包含数据透视表，不能在该Sheet扩展明细行；可改用单值映射或将明细放到普通Sheet');
      for (const row of rows) {
        const current = rowNumberOf(row); if (current <= rowNumber) continue;
        row.setAttribute('r', String(current + delta));
        for (const cell of direct(row, 'c', S)) cell.setAttribute('r', shiftCellAddress(cellAddressOf(cell), delta));
      }
      shiftSheetRanges(context.sheetDoc, rowNumber, delta);
      await shiftRelatedExcelParts(zip, relationships, rowNumber, delta);
      shiftWorkbookNames(context.workbookDoc, state.templateSheet, rowNumber, delta);
    }
    records.forEach((record, offset) => {
      const clone = prototype.cloneNode(true), targetRow = rowNumber + offset; clone.setAttribute('r', String(targetRow));
      for (const cell of direct(clone, 'c', S)) cell.setAttribute('r', replaceAddressRow(cellAddressOf(cell), targetRow));
      for (const [targetId, mapping] of rowMappings) {
        const source = XLSX.utils.decode_cell(targetId.slice(2)), address = XLSX.utils.encode_cell({ r: targetRow - 1, c: source.c });
        writeCellValue(ensureRowCell(clone, address), C.text(record[mapping.field]));
      }
      data.insertBefore(clone, prototype);
    });
    prototype.remove();
  }

  async function loadSheetRelationships(zip, sheetPath) {
    const slash = sheetPath.lastIndexOf('/'), directory = sheetPath.slice(0, slash), name = sheetPath.slice(slash + 1), relsPath = `${directory}/_rels/${name}.rels`;
    const file = zip.file(relsPath); if (!file) return [];
    const doc = parseXml(await file.async('string'));
    return all(doc, 'Relationship', P).map(item => ({
      id: item.getAttribute('Id'), type: item.getAttribute('Type') || '', path: resolveZipPath(directory, item.getAttribute('Target')), relsPath, doc,
    }));
  }

  async function shiftRelatedExcelParts(zip, relationships, rowNumber, delta) {
    for (const relation of relationships) {
      const file = zip.file(relation.path); if (!file) continue;
      if (/\/table$/i.test(relation.type)) {
        const doc = parseXml(await file.async('string')), table = doc.documentElement;
        if (table.getAttribute('ref')) table.setAttribute('ref', shiftRange(table.getAttribute('ref'), rowNumber, delta));
        for (const filter of all(doc, 'autoFilter', S)) if (filter.getAttribute('ref')) filter.setAttribute('ref', shiftRange(filter.getAttribute('ref'), rowNumber, delta));
        zip.file(relation.path, new XMLSerializer().serializeToString(doc));
      } else if (/\/drawing$/i.test(relation.type)) {
        const doc = parseXml(await file.async('string'));
        for (const marker of [...doc.getElementsByTagNameNS('*', 'row')]) {
          const value = Number(marker.textContent); if (Number.isInteger(value) && value >= rowNumber) marker.textContent = String(value + delta);
        }
        zip.file(relation.path, new XMLSerializer().serializeToString(doc));
      }
    }
  }

  function shiftSheetRanges(doc, rowNumber, delta) {
    for (const element of [...doc.getElementsByTagName('*')]) {
      for (const attribute of ['ref', 'sqref']) {
        const value = element.getAttribute(attribute); if (!value) continue;
        if (['dimension', 'mergeCell', 'autoFilter', 'conditionalFormatting', 'dataValidation', 'hyperlink'].includes(element.localName)) element.setAttribute(attribute, shiftRangeList(value, rowNumber, delta));
      }
    }
    for (const item of all(doc, 'brk', S)) {
      const value = Number(item.getAttribute('id')); if (Number.isInteger(value) && value > rowNumber) item.setAttribute('id', String(value + delta));
    }
  }

  function shiftWorkbookNames(workbookDoc, sheetName, rowNumber, delta) {
    for (const name of all(workbookDoc, 'definedName', S)) {
      const value = name.textContent || '', bang = value.lastIndexOf('!'); if (bang < 0) continue;
      const prefix = value.slice(0, bang).replace(/^'/, '').replace(/'$/, '').replace(/''/g, "'");
      if (prefix !== sheetName) continue;
      name.textContent = `${value.slice(0, bang + 1)}${shiftRangeList(value.slice(bang + 1), rowNumber, delta)}`;
    }
  }

  function shiftRangeList(value, rowNumber, delta) { return value.split(/(\s+)/).map(part => /[A-Z]\$?\d/i.test(part) ? shiftRange(part, rowNumber, delta) : part).join(''); }
  function parseCellRef(value) { const match = value.match(/^(\$?[A-Z]{1,3})(\$?)(\d+)$/i); return match ? { column: match[1], rowAbsolute: match[2], row: Number(match[3]) } : null; }
  function formatCellRef(cell) { return `${cell.column}${cell.rowAbsolute}${cell.row}`; }
  function shiftRange(value, rowNumber, delta) {
    const parts = value.split(':').map(parseCellRef); if (!parts[0] || (parts.length > 1 && !parts[1])) return value;
    if (parts.length === 1) { if (parts[0].row > rowNumber) parts[0].row += delta; return formatCellRef(parts[0]); }
    const [start, end] = parts;
    if (start.row > rowNumber) { start.row += delta; end.row += delta; }
    else if (end.row >= rowNumber) end.row += delta;
    return `${formatCellRef(start)}:${formatCellRef(end)}`;
  }
  function rangeTouchesRow(value, rowNumber) { const parts = value.split(':').map(parseCellRef); return parts[0] && rowNumber >= parts[0].row && rowNumber <= (parts[1]?.row || parts[0].row); }
  function rowNumberFromAddress(value) { return Number(value.match(/(\d+)$/)?.[1] || 0); }
  function shiftCellAddress(value, delta) { return replaceAddressRow(value, rowNumberFromAddress(value) + delta); }
  function replaceAddressRow(value, rowNumber) { return value.replace(/\d+$/, String(rowNumber)); }

  async function forceWorkbookRecalculation(zip, context) {
    let calc = all(context.workbookDoc, 'calcPr', S)[0];
    if (!calc) { calc = context.workbookDoc.createElementNS(S, 'calcPr'); context.workbookDoc.documentElement.append(calc); }
    calc.setAttribute('calcMode', 'auto'); calc.setAttribute('fullCalcOnLoad', '1'); calc.setAttribute('forceFullCalc', '1');
    zip.file(context.workbookPath, new XMLSerializer().serializeToString(context.workbookDoc));
    const calcRelationships = all(context.workbookRelsDoc, 'Relationship', P).filter(item => /\/calcChain$/i.test(item.getAttribute('Type') || ''));
    if (calcRelationships.length || zip.file('xl/calcChain.xml')) {
      calcRelationships.forEach(item => item.remove()); zip.file(context.workbookRelsPath, new XMLSerializer().serializeToString(context.workbookRelsDoc)); zip.remove('xl/calcChain.xml');
      const typesFile = zip.file('[Content_Types].xml');
      if (typesFile) {
        const doc = parseXml(await typesFile.async('string'));
        [...doc.getElementsByTagNameNS('*', 'Override')].filter(item => item.getAttribute('PartName') === '/xl/calcChain.xml').forEach(item => item.remove());
        zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(doc));
      }
    }
  }

  function resolveMapping(mapping) { return C.resolveField(state.order.rows, mapping.field, mapping.strategy || 'first', mapping.manual || ''); }

  function renderGeneratedPreview() {
    if (state.templateModel.type === 'docx') {
      $('generatedPreview').innerHTML = `<div class="generated-page">${state.templateModel.blocks.map(block => {
        if (block.type === 'paragraph') return `<p>${esc(previewValue(block.target))}</p>`;
        return `<table class="generated-table"><tbody>${block.rows.flatMap(row => {
          const records = row.rowKey === state.detailRow ? state.order.rows : [null];
          return records.map(record => `<tr>${row.cells.map(cell => `<td>${esc(previewValue(cell, record))}</td>`).join('')}</tr>`);
        }).join('')}</tbody></table>`;
      }).join('')}</div>`;
    } else {
      $('generatedPreview').innerHTML = `<div class="generated-page"><table class="generated-table"><tbody>${state.templateModel.rows.flatMap(row => {
        const records = row.rowKey === state.detailRow ? state.order.rows : [null];
        return records.map(record => `<tr>${row.cells.map(cell => `<td>${esc(previewValue(cell, record))}</td>`).join('')}</tr>`);
      }).join('')}</tbody></table></div>`;
    }
  }

  function previewValue(target, record = null) {
    const mapping = state.mappings[target.id];
    if (!mapping?.field) return target.value;
    if (mapping.mode === 'detail') return C.text(record?.[mapping.field]);
    return resolveMapping(mapping);
  }

  function downloadOutput() {
    if (!state.output) return;
    const url = URL.createObjectURL(state.output), anchor = document.createElement('a');
    anchor.href = url; anchor.download = state.outputFileName; document.body.append(anchor); anchor.click(); anchor.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function exportMapping() {
    if (!state.templateModel) return toast('请先上传合同模板', 'error');
    const payload = mappingPayload(), blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob), anchor = document.createElement('a'); anchor.href = url; anchor.download = `${C.sanitizeFileName(state.templateFile.name.replace(/\.[^.]+$/, ''))}-映射.json`; anchor.click(); URL.revokeObjectURL(url);
  }

  async function importMapping(file) {
    if (!file || !state.templateModel) return;
    try {
      const payload = JSON.parse(await file.text()); applyMappingPayload(payload); saveMapping(); updateAll(); toast('映射方案导入成功', 'success');
    } catch (error) { toast(error.message || '映射文件无法读取', 'error'); }
    finally { $('importMapping').value = ''; }
  }

  function mappingPayload() {
    return { version: 1, fingerprint: state.fingerprint, templateType: state.templateType, templateSheet: state.templateSheet || '', detailRow: state.detailRow, mappings: state.mappings };
  }

  function applyMappingPayload(payload) {
    if (payload?.version !== 1 || payload.fingerprint !== state.fingerprint || payload.templateType !== state.templateType || (payload.templateSheet || '') !== (state.templateSheet || '')) throw new Error('映射文件与当前合同模板不完全一致');
    const targetIds = new Set(state.templateModel.targets.map(target => target.id)), fields = new Set(state.order?.headers || []), mappings = {};
    for (const [target, mapping] of Object.entries(payload.mappings || {})) if (targetIds.has(target) && fields.has(mapping.field)) mappings[target] = { field: mapping.field, mode: mapping.mode === 'detail' ? 'detail' : 'single', strategy: mapping.strategy || '', manual: mapping.manual || '' };
    state.mappings = mappings; state.detailRow = state.templateModel.targets.some(target => target.rowKey === payload.detailRow) ? payload.detailRow : '';
  }

  function mappingKey() { return state.fingerprint ? `gylsjqx-contract-mapping:${state.fingerprint}:${state.templateSheet || 'docx'}` : ''; }
  function saveMapping() { try { const key = mappingKey(); if (key) localStorage.setItem(key, JSON.stringify(mappingPayload())); } catch (_) {} }
  function restoreMapping() {
    try { const key = mappingKey(), raw = key && localStorage.getItem(key); if (raw) applyMappingPayload(JSON.parse(raw)); }
    catch (_) { state.mappings = {}; state.detailRow = ''; }
  }

  function parseXml(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('模板XML无法读取');
    return doc;
  }

  function wordText(element) {
    let output = '';
    for (const node of element.getElementsByTagName('*')) {
      if (node.namespaceURI !== W) continue;
      if (node.localName === 't') output += node.textContent;
      else if (node.localName === 'br' || node.localName === 'cr') output += '\n';
      else if (node.localName === 'tab') output += '    ';
    }
    return output;
  }
  function wordCellText(cell) { return direct(cell, 'p').map(wordText).join('\n'); }

  function wordNode(doc, name) { return doc.createElementNS(W, `w:${name}`); }
  function putWordText(target, value) {
    const doc = target.ownerDocument;
    let paragraph = target.localName === 'p' ? target : direct(target, 'p')[0];
    if (!paragraph) { paragraph = wordNode(doc, 'p'); target.append(paragraph); }
    if (target !== paragraph) direct(target, 'p').slice(1).forEach(item => item.remove());
    const runProperties = all(paragraph, 'rPr')[0]?.cloneNode(true);
    [...paragraph.children].filter(item => item.localName !== 'pPr').forEach(item => item.remove());
    const run = wordNode(doc, 'r'); if (runProperties) run.append(runProperties);
    C.text(value).split('\n').forEach((line, index) => {
      if (index) run.append(wordNode(doc, 'br'));
      const textNode = wordNode(doc, 't'); textNode.setAttribute('xml:space', 'preserve'); textNode.textContent = line; run.append(textNode);
    });
    paragraph.append(run);
  }

  function locateWordTarget(doc, id) {
    const body = all(doc, 'body')[0];
    let match = id.match(/^p:(\d+)$/); if (match) return direct(body, 'p')[Number(match[1])];
    match = id.match(/^t:(\d+):r:(\d+):c:(\d+)$/); if (!match) return null;
    return direct(direct(direct(body, 'tbl')[Number(match[1])], 'tr')[Number(match[2])], 'tc')[Number(match[3])];
  }

  async function sha256(bytes) {
    const digest = await crypto.subtle.digest('SHA-256', bytes.slice(0));
    return [...new Uint8Array(digest)].map(value => value.toString(16).padStart(2, '0')).join('');
  }
})();
