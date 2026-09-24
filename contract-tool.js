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
    pdf: 'application/pdf',
  };
  const ONLYOFFICE_ORIGIN = 'https://edit.chaxus.com';
  const CONVERSION_TIMEOUT = 240000;
  const NUMERIC_BUSINESS_FIELDS = new Set(['sequence', 'quantity', 'taxUnitPrice', 'taxAmount', 'taxRate', 'taxTotalLower']);
  const BUSINESS_FIELDS = [
    { key: 'sequence', label: '序号', aliases: ['序号', '行号'], kind: 'detail', automatic: true },
    { key: 'materialCode', label: '物料编码', aliases: ['物料编码', '产品编码', '商品编码', '货号'], kind: 'detail' },
    { key: 'materialName', label: '物料名称', aliases: ['物料名称', '物料', '产品名称', '商品名称', '品名'], kind: 'detail' },
    { key: 'specification', label: '规格型号', aliases: ['规格型号', '规格', '型号'], kind: 'detail' },
    { key: 'sku', label: 'SKU', aliases: ['SKU', 'SKU编码'], kind: 'detail' },
    { key: 'unit', label: '单位', aliases: ['单位', '计量单位'], kind: 'detail' },
    { key: 'quantity', label: '数量', aliases: ['数量', '采购数量', '订单数量'], kind: 'detail' },
    { key: 'taxUnitPrice', label: '含税运单价（元）', aliases: ['含税运单价（元）', '含税运单价', '含税单价（元）', '含税单价', '单价'], kind: 'detail' },
    { key: 'taxAmount', label: '含税运总金额（元）', aliases: ['含税运总金额（元）', '含税运总金额', '含税总金额（元）', '含税总金额', '含税金额', '金额'], kind: 'detail' },
    { key: 'taxRate', label: '税率', aliases: ['税率', '增值税率'], kind: 'detail' },
    { key: 'deliveryTime', label: '交货时间', aliases: ['交货时间', '交期', '要求货好时间', '要求交货日期'], kind: 'single' },
    { key: 'remark', label: '备注', aliases: ['备注', '说明'], kind: 'detail' },
    { key: 'taxTotalLower', label: '含税运合计（小写）', aliases: ['含税运合计（小写）', '含税运合计小写', '合计（小写）', '合计小写', '小写合计'], kind: 'single' },
    { key: 'taxTotalUpper', label: '含税运合计（大写）', aliases: ['含税运合计（大写）', '含税运合计大写', '合计（大写）', '合计大写', '大写合计'], kind: 'single' },
    { key: 'contractNumber', label: '合同编号', aliases: ['合同编号', '合同号'], kind: 'single' },
    { key: 'orderNumber', label: '订单编号', aliases: ['订单编号', '采购订单号', '采购单号', '订单号'], kind: 'single' },
    { key: 'buyer', label: '采购方（甲方）', aliases: ['采购方（甲方）', '采购方', '甲方', '买方'], kind: 'single' },
    { key: 'supplier', label: '供应商（乙方）', aliases: ['供应商名称', '供应商（乙方）', '供应商', '乙方', '卖方'], kind: 'single' },
    { key: 'signDate', label: '签订日期', aliases: ['签订日期', '合同日期', '签约日期'], kind: 'single' },
    { key: 'deliveryPlace', label: '交货地点', aliases: ['交货地点', '送货地址', '交付地点'], kind: 'single' },
    { key: 'paymentTerms', label: '付款方式', aliases: ['付款方式', '付款条件', '结算方式', '结算条件'], kind: 'single' },
  ];
  let pdfJsPromise = null;
  let converterFrame = null;
  let converterReadyPromise = null;
  let previewResizeTimer = null;
  const converterRequests = new Map();
  const state = {
    orderFile: null, orderBytes: null, orderBook: null, orderSheet: '', order: null,
    templateFile: null, templateBytes: null, templateType: '', templateBook: null, templateSheet: '', templateModel: null, templateFeatures: [], fingerprint: '',
    previewDocument: null, previewObjectUrl: '', previewMode: '', previewPageCount: 0, previewPage: 0, previewRenderToken: 0,
    mappings: {}, detailRow: '', fieldSelections: {}, fieldStrategies: {}, templateBindings: {}, mappingSignature: '', output: null, outputFileName: '', outputPdf: null, outputPdfFileName: '', busy: false,
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
    $('orderSheet').addEventListener('change', event => { state.orderSheet = event.target.value; state.fieldSelections = {}; state.fieldStrategies = {}; state.mappingSignature = ''; analyzeOrder(); invalidateOutput(); updateAll(); });
    $('templateSheet').addEventListener('change', event => { state.templateSheet = event.target.value; buildExcelModel(); state.mappings = {}; state.detailRow = ''; state.templateBindings = {}; state.mappingSignature = ''; invalidateOutput(); restoreMapping(); updateAll(); });
    $('resultPrevious').addEventListener('click', () => changePreviewPage(-1));
    $('resultNext').addEventListener('click', () => changePreviewPage(1));
    $('outputName').addEventListener('input', () => { invalidateOutput(); updateConfirmation(); });
    $('confirmGenerate').addEventListener('change', updateConfirmation);
    $('generateContract').addEventListener('click', generate);
    $('downloadContract').addEventListener('click', downloadOutput);
    $('downloadPdf').addEventListener('click', downloadPdf);
    $('confirmExport').addEventListener('change', updateExportButtons);
    $('clearContract').addEventListener('click', () => location.reload());
    $('exportMapping').addEventListener('click', exportMapping);
    $('importMapping').addEventListener('change', event => importMapping(event.target.files[0]));
    for (const [dropId, inputId] of [['orderDrop', 'orderFile'], ['templateDrop', 'templateFile']]) bindDrop($(dropId), $(inputId));
    window.addEventListener('message', handleConverterMessage);
    window.addEventListener('resize', () => {
      if (!state.previewDocument) return;
      clearTimeout(previewResizeTimer);
      previewResizeTimer = setTimeout(renderGeneratedPdf, 100);
    });
    window.addEventListener('beforeunload', releasePreviewDocument);
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
    state.output = null; state.outputFileName = ''; state.outputPdf = null; state.outputPdfFileName = '';
    releasePreviewDocument(); state.previewPageCount = 0; state.previewPage = 0;
    $('confirmExport').checked = false; $('resultStage').hidden = true; $('resultPagination').hidden = true;
    $('downloadContract').disabled = true; $('downloadPdf').disabled = true; $('generatedPreview').replaceChildren();
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
      if (state.templateModel) restoreMapping();
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
      state.templateBook = null; state.templateSheet = ''; state.mappings = {}; state.detailRow = ''; state.fieldSelections = {}; state.fieldStrategies = {}; state.templateBindings = {}; state.mappingSignature = '';
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

  function loadPdfJs() {
    if (!pdfJsPromise) {
      pdfJsPromise = import('./vendor/pdf.min.mjs').then(pdfjs => {
        pdfjs.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.mjs', location.href).href;
        return pdfjs;
      });
    }
    return pdfJsPromise;
  }

  function releasePreviewDocument() {
    state.previewRenderToken += 1;
    state.previewDocument?.destroy?.();
    state.previewDocument = null;
    if (state.previewObjectUrl) URL.revokeObjectURL(state.previewObjectUrl);
    state.previewObjectUrl = ''; state.previewMode = '';
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
    if (ready) { ensureBusinessMappings(); renderBusinessMappings(); }
    const summary = [];
    if (state.orderFile) summary.push(`订单：${state.orderFile.name}（${state.order.rows.length}行）`);
    if (state.templateFile) summary.push(`模板：${state.templateFile.name}`);
    $('fileSummary').textContent = summary.join('；') || '等待上传订单和合同模板';
    const selectedCount = BUSINESS_FIELDS.filter(field => field.automatic || state.fieldSelections[field.key]).length;
    $('mappingStatus').textContent = ready ? `${mappedEntries().length} / ${selectedCount} 已识别` : '0 个映射';
    updateConfirmation(); updateSteps();
  }

  function normalizeBusinessLabel(value) {
    return C.text(value).toUpperCase().replace(/[^A-Z0-9\u4e00-\u9fff]/g, '');
  }

  function fieldMatchScore(value, definition) {
    const text = normalizeBusinessLabel(value); if (!text) return 0;
    let score = 0;
    for (const alias of definition.aliases) {
      const candidate = normalizeBusinessLabel(alias); if (!candidate) continue;
      if (text === candidate) score = Math.max(score, 100 + candidate.length);
      else if (candidate.length >= 4 && text.includes(candidate)) score = Math.max(score, 60 + candidate.length);
      else if (text.length >= 4 && candidate.includes(text)) score = Math.max(score, 40 + text.length);
    }
    return score;
  }

  function bestDefinition(value, kind = '') {
    return BUSINESS_FIELDS.filter(field => !kind || field.kind === kind).map(field => ({ field, score: fieldMatchScore(value, field) })).sort((a, b) => b.score - a.score)[0];
  }

  function bestOrderHeader(definition) {
    const best = state.order.headers.map(header => ({ header, score: fieldMatchScore(header, definition) })).sort((a, b) => b.score - a.score)[0];
    return best?.score > 0 ? best.header : '';
  }

  function detailRecords() {
    const rows = state.order?.rows || [];
    const identityFields = ['materialCode', 'materialName', 'sku'].map(key => state.fieldSelections[key]).filter(field => field && field !== '@sequence');
    if (identityFields.length) {
      const matched = rows.filter(record => identityFields.some(field => isDetailIdentity(record[field])));
      if (matched.length) return matched;
    }
    const fallbackFields = ['specification', 'quantity', 'taxUnitPrice', 'taxAmount'].map(key => state.fieldSelections[key]).filter(field => field && field !== '@sequence');
    if (fallbackFields.length) {
      const matched = rows.filter(record => fallbackFields.some(field => C.text(record[field]).trim()));
      if (matched.length) return matched;
    }
    return rows;
  }

  function isDetailIdentity(value) {
    const normalized = C.text(value).replace(/[\s：:]/g, '');
    return !!normalized && !/^(序号|合计|总计|小计|交货时间|交货日期|交期|付款方式|付款条件|备注|说明|签字|盖章)$/.test(normalized);
  }

  function templateRows() {
    if (state.templateModel.type === 'docx') return state.templateModel.blocks.filter(block => block.type === 'table').flatMap(block => block.rows);
    return state.templateModel.rows;
  }

  function locateRow(rowKey) { return templateRows().find(row => row.rowKey === rowKey); }

  function detectTemplateBindings() {
    const bindings = {}, rows = templateRows(), detailFields = BUSINESS_FIELDS.filter(field => field.kind === 'detail');
    let bestHeader = null;
    for (let index = 0; index < rows.length - 1; index += 1) {
      const matches = new Map(); let score = 0;
      for (const cell of rows[index].cells) {
        const found = bestDefinition(cell.value, 'detail');
        if (found?.score > 0 && !matches.has(found.field.key)) { matches.set(found.field.key, cell); score += found.score; }
      }
      if (matches.size && (!bestHeader || matches.size > bestHeader.matches.size || (matches.size === bestHeader.matches.size && score > bestHeader.score))) {
        bestHeader = { index, matches, score };
      }
    }
    state.detailRow = '';
    if (bestHeader) {
      const detailRow = rows[bestHeader.index + 1]; state.detailRow = detailRow.rowKey;
      for (const definition of detailFields) {
        const headerCell = bestHeader.matches.get(definition.key); if (!headerCell) continue;
        const target = detailRow.cells.find(cell => cell.cellIndex === headerCell.cellIndex);
        if (target) bindings[definition.key] = { targetId: target.id, mode: 'detail', templateLabel: headerCell.value };
      }
    }
    const excludedRows = new Set([bestHeader ? rows[bestHeader.index].rowKey : '', state.detailRow].filter(Boolean));
    const usedTargets = new Set(Object.values(bindings).map(binding => binding.targetId));
    for (const definition of BUSINESS_FIELDS.filter(field => field.kind === 'single')) {
      let best = null;
      for (const target of state.templateModel.targets) {
        if (excludedRows.has(target.rowKey)) continue;
        const score = fieldMatchScore(target.value, definition);
        if (score > 0 && (!best || score > best.score)) best = { target, score };
      }
      if (!best) continue;
      const row = locateRow(best.target.rowKey), next = row?.cells.find(cell => cell.cellIndex === best.target.cellIndex + 1);
      const canUseNext = next && !usedTargets.has(next.id) && (!bestDefinition(next.value)?.score || /待填|填写|空白/.test(C.text(next.value)));
      const target = canUseNext ? next : best.target;
      if (usedTargets.has(target.id)) continue;
      bindings[definition.key] = { targetId: target.id, mode: 'single', preserveLabel: target.id === best.target.id, labelText: best.target.value, templateLabel: best.target.value };
      usedTargets.add(target.id);
    }
    state.templateBindings = bindings;
  }

  function ensureBusinessMappings() {
    const signature = `${state.fingerprint}:${state.templateSheet}:${state.orderSheet}:${state.order.headers.join('|')}`;
    if (state.mappingSignature !== signature) {
      for (const definition of BUSINESS_FIELDS) {
        if (definition.automatic) state.fieldSelections[definition.key] = '@sequence';
        else if (!state.fieldSelections[definition.key] || !state.order.headers.includes(state.fieldSelections[definition.key])) state.fieldSelections[definition.key] = bestOrderHeader(definition);
      }
      state.mappingSignature = signature;
    }
    detectTemplateBindings(); rebuildMappings();
  }

  function rebuildMappings() {
    const mappings = {};
    for (const definition of BUSINESS_FIELDS) {
      const binding = state.templateBindings[definition.key], field = state.fieldSelections[definition.key];
      if (!binding || !field) continue;
      const values = field === '@sequence' ? [] : C.distinctValues(state.order.rows, field);
      mappings[binding.targetId] = {
        field, mode: binding.mode, businessKey: definition.key, preserveLabel: !!binding.preserveLabel, labelText: binding.labelText || '',
        strategy: binding.mode === 'single' ? (values.length <= 1 ? 'first' : (state.fieldStrategies?.[definition.key] || '')) : '',
      };
    }
    state.mappings = mappings;
  }

  function renderBusinessMappings() {
    if (!state.fieldStrategies) state.fieldStrategies = {};
    const mapped = new Set(mappedEntries().map(([, mapping]) => mapping.businessKey));
    const detected = Object.keys(state.templateBindings).length;
    const detailRows = detailRecords();
    $('templateDetectionSummary').textContent = `模板自动识别 ${detected} 个可写字段；序号按 ${detailRows.length} 条物料明细自动生成`;
    $('businessMappingRows').innerHTML = BUSINESS_FIELDS.map(definition => {
      const selection = state.fieldSelections[definition.key] || '', binding = state.templateBindings[definition.key];
      const values = selection && selection !== '@sequence' ? C.distinctValues(state.order.rows, selection) : [];
      const conflict = definition.kind === 'single' && values.length > 1;
      const options = definition.automatic
        ? '<option value="@sequence">自动生成 1、2、3…</option>'
        : `<option value="">不填写</option>${state.order.headers.map(header => `<option value="${esc(header)}" ${header === selection ? 'selected' : ''}>${esc(header)}</option>`).join('')}`;
      const strategy = conflict ? `<select class="business-field-select business-strategy-select" data-strategy-key="${definition.key}" aria-label="${esc(definition.label)}多值处理"><option value="">该列有多个值，请选择处理方式</option><option value="first" ${state.fieldStrategies[definition.key] === 'first' ? 'selected' : ''}>取第一条非空值</option><option value="merge" ${state.fieldStrategies[definition.key] === 'merge' ? 'selected' : ''}>合并去重值</option><option value="sum" ${state.fieldStrategies[definition.key] === 'sum' ? 'selected' : ''}>求和</option></select>` : '';
      const success = !!(binding && selection && mapped.has(definition.key));
      const templateField = binding?.templateLabel || '未识别到对应字段';
      return `<div class="business-mapping-row" role="row" data-business-key="${definition.key}">
        <div class="business-field-name" role="cell" data-cell-label="映射字段"><strong>${esc(definition.label)}</strong><small>${definition.automatic ? '序号由系统生成' : esc(definition.aliases.slice(0, 3).join('、'))}</small></div>
        <div class="business-order-field" role="cell" data-cell-label="订单明细"><select class="business-field-select" data-field-key="${definition.key}" ${definition.automatic ? 'disabled' : ''} aria-label="${esc(definition.label)}对应订单列">${options}</select>${strategy}</div>
        <span class="business-template-field" role="cell" data-cell-label="合同模板">${esc(templateField)}</span>
        <span class="template-detection ${success ? 'detected' : ''}" role="cell"><span class="mobile-cell-label" aria-hidden="true">是否识别：</span><span class="detection-text">${success ? '识别成功' : '未识别'}</span></span>
        <span class="business-write-mode" role="cell" data-cell-label="写入方式">${definition.kind === 'detail' ? '按订单逐行写入' : '填充内容'}</span>
      </div>`;
    }).join('');
    $('businessMappingRows').querySelectorAll('[data-field-key]').forEach(select => select.addEventListener('change', event => {
      state.fieldSelections[event.target.dataset.fieldKey] = event.target.value; state.fieldStrategies[event.target.dataset.fieldKey] = '';
      invalidateOutput(); rebuildMappings(); saveMapping(); updateAll();
    }));
    $('businessMappingRows').querySelectorAll('[data-strategy-key]').forEach(select => select.addEventListener('change', event => {
      state.fieldStrategies[event.target.dataset.strategyKey] = event.target.value;
      invalidateOutput(); rebuildMappings(); saveMapping(); updateAll();
    }));
  }

  async function renderGeneratedPdf() {
    if (state.previewMode === 'native' && state.previewObjectUrl) {
      $('resultPagination').hidden = true;
      $('generatedPreview').innerHTML = `<iframe class="native-pdf-preview" src="${esc(state.previewObjectUrl)}#toolbar=0&navpanes=0&view=FitH" title="生成合同PDF预览"></iframe>`;
      return;
    }
    if (!state.previewDocument) return;
    state.previewPage = Math.max(0, Math.min(state.previewPage, state.previewPageCount - 1));
    const pageNumber = state.previewPage + 1, renderToken = ++state.previewRenderToken;
    $('generatedPreview').innerHTML = `<div class="pdf-preview-shell"><canvas class="pdf-preview-canvas" data-page="${pageNumber}" aria-label="生成合同PDF第${pageNumber}页"></canvas><span class="pdf-render-status">正在渲染第 ${pageNumber} 页…</span></div>`;
    $('resultPagination').hidden = false;
    $('resultPageLabel').textContent = `第 ${pageNumber} / ${state.previewPageCount} 页`;
    $('resultPrevious').disabled = state.previewPage === 0;
    $('resultNext').disabled = state.previewPage >= state.previewPageCount - 1;
    $('resultPagination').dataset.totalPages = String(state.previewPageCount);
    $('resultPagination').dataset.currentPage = String(state.previewPage + 1);
    try {
      const pdfPage = await state.previewDocument.getPage(pageNumber);
      if (renderToken !== state.previewRenderToken) return;
      const host = $('generatedPreview'), shell = host.querySelector('.pdf-preview-shell'), canvas = host.querySelector('.pdf-preview-canvas');
      const baseViewport = pdfPage.getViewport({ scale: 1 }), availableWidth = Math.min(794, Math.max(260, host.clientWidth - 48));
      const cssScale = availableWidth / baseViewport.width, pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
      const cssViewport = pdfPage.getViewport({ scale: cssScale }), renderViewport = pdfPage.getViewport({ scale: cssScale * pixelRatio });
      canvas.width = Math.ceil(renderViewport.width); canvas.height = Math.ceil(renderViewport.height);
      canvas.style.width = `${Math.ceil(cssViewport.width)}px`; canvas.style.height = `${Math.ceil(cssViewport.height)}px`;
      shell.style.width = `${Math.ceil(cssViewport.width)}px`; shell.style.height = `${Math.ceil(cssViewport.height)}px`;
      await pdfPage.render({ canvasContext: canvas.getContext('2d', { alpha: false }), viewport: renderViewport }).promise;
      if (renderToken !== state.previewRenderToken) return;
      canvas.dataset.rendered = 'true'; shell.querySelector('.pdf-render-status')?.remove(); pdfPage.cleanup();
    } catch (error) {
      if (renderToken !== state.previewRenderToken) return;
      $('generatedPreview').innerHTML = `<div class="pdf-render-error">生成PDF第 ${pageNumber} 页渲染失败：${esc(error.message || '未知错误')}</div>`;
    }
  }

  function changePreviewPage(offset) {
    if (!state.previewDocument) return;
    const total = state.previewPageCount || 1, next = Math.max(0, Math.min(state.previewPage + offset, total - 1));
    if (next === state.previewPage) return;
    state.previewPage = next; renderGeneratedPdf(); $('generatedPreview').scrollTop = 0; $('generatedPreview').scrollLeft = 0;
  }

  function mappedEntries() { return Object.entries(state.mappings).filter(([, mapping]) => mapping?.field); }
  function findTarget(id) { return state.templateModel?.targets.find(target => target.id === id); }

  function updateConfirmation() {
    if (!state.order || !state.templateModel) { $('generateContract').disabled = true; return; }
    const issues = C.mappingIssues(state.order.rows, state.mappings, state.detailRow);
    if (state.detailRow && !mappedEntries().some(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow)) issues.push('明细模板行还没有映射任何订单字段');
    const notDetected = BUSINESS_FIELDS.filter(field => state.fieldSelections[field.key] && !state.templateBindings[field.key]).map(field => field.label);
    const outputName = C.sanitizeFileName($('outputName').value);
    if (!$('outputName').value.trim()) issues.push('请填写合同文件名称');
    const detailRows = detailRecords();
    $('confirmSummary').innerHTML = `订单：<strong>${esc(state.orderSheet)}</strong>，识别 ${detailRows.length} 条物料明细${detailRows.length !== state.order.rows.length ? `（原表 ${state.order.rows.length} 行）` : ''}；模板：<strong>${esc(state.templateFile.name)}</strong>${state.templateSheet ? `，合同Sheet：<strong>${esc(state.templateSheet)}</strong>` : ''}；输出：<strong>${esc(outputName)}.${state.templateType}</strong>`;
    const warnings = notDetected.length ? `<div class="warnings">模板中未识别：${esc(notDetected.join('、'))}；这些字段本次不会写入。</div>` : '';
    $('contractIssues').innerHTML = `${issues.length ? `<ul>${issues.map(issue => `<li>${esc(issue)}</li>`).join('')}</ul>` : '<div class="ready">映射检查通过，可以生成PDF预览。</div>'}${warnings}`;
    $('generateContract').disabled = state.busy || issues.length > 0 || !$('confirmGenerate').checked;
  }

  function updateSteps() {
    const steps = [...$('contractSteps').children];
    const mapped = mappedEntries().length > 0;
    const values = [!!state.order, !!state.templateModel, mapped, !!state.outputPdf, !!state.outputPdf && $('confirmExport').checked];
    let active = values.findIndex(value => !value); if (active < 0) active = 4;
    steps.forEach((step, index) => { step.classList.toggle('done', values[index] && index < active); step.classList.toggle('active', index === active); });
  }

  async function generate() {
    updateConfirmation(); if ($('generateContract').disabled) return;
    invalidateOutput();
    state.busy = true; $('generateContract').textContent = '正在生成合同…'; updateConfirmation(); status('正在生成合同副本…');
    try {
      const blob = state.templateType === 'docx' ? await generateDocx() : await generateXlsx();
      state.output = blob; state.outputFileName = `${C.sanitizeFileName($('outputName').value)}.${state.templateType}`;
      $('generateContract').textContent = '正在转换PDF…'; status('合同副本已生成，正在浏览器内转换PDF；首次加载可能需要一些时间…');
      let pdfFile = await convertContractToPdf(blob, state.outputFileName);
      try { state.outputPdf = await preparePdfPreview(pdfFile, false); }
      catch (error) {
        if (error.code !== 'PDF_PARSE_FAILED') throw error;
        status('第一次PDF结果无法解析，正在自动重新转换…');
        pdfFile = await convertContractToPdf(blob, state.outputFileName);
        state.outputPdf = await preparePdfPreview(pdfFile, true);
      }
      state.outputPdfFileName = `${C.sanitizeFileName($('outputName').value)}.pdf`;
      $('resultStage').hidden = false;
      $('downloadContract').textContent = `下载 ${state.templateType.toUpperCase()}`;
      $('downloadPdf').textContent = '下载 PDF';
      $('confirmExport').checked = false; updateExportButtons(); renderGeneratedPdf();
      const previewText = state.previewMode === 'native' ? '浏览器原生预览' : `${state.previewPageCount}页`;
      saveMapping(); status(`PDF预览已生成：${state.outputPdfFileName}（${previewText}）`, 'success'); toast('PDF预览生成完成，请核对后确认导出', 'success'); updateSteps();
      $('resultStage').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) { status(error.message, 'error'); toast(error.message, 'error'); }
    finally { state.busy = false; $('generateContract').textContent = '生成PDF预览'; updateConfirmation(); }
  }

  async function preparePdfPreview(file, allowNativeFallback) {
    const normalized = await normalizePdfFile(file, state.outputFileName.replace(/\.[^.]+$/, '.pdf'));
    const bytes = await normalized.arrayBuffer();
    const pdfjs = await loadPdfJs();
    let document;
    try {
      if (window.__CONTRACT_FORCE_NATIVE_PDF_PREVIEW__) throw new Error('测试浏览器原生PDF预览');
      document = await pdfjs.getDocument({
        data: new Uint8Array(bytes.slice(0)),
        cMapUrl: new URL('vendor/pdfjs-cmaps/', location.href).href,
        cMapPacked: true,
        standardFontDataUrl: new URL('vendor/pdfjs-standard-fonts/', location.href).href,
        wasmUrl: new URL('vendor/pdfjs-wasm/', location.href).href,
        useSystemFonts: true,
      }).promise;
    } catch (error) {
      if (!allowNativeFallback && !window.__CONTRACT_FORCE_NATIVE_PDF_PREVIEW__) {
        const parseError = new Error('PDF转换结果无法解析，正在重试'); parseError.code = 'PDF_PARSE_FAILED'; throw parseError;
      }
      releasePreviewDocument();
      state.previewObjectUrl = URL.createObjectURL(normalized); state.previewMode = 'native'; state.previewPageCount = 0; state.previewPage = 0;
      console.warn('PDF.js读取失败，已切换浏览器原生预览：', error?.message || error);
      return normalized;
    }
    if (!document.numPages) { document.destroy(); throw new Error('生成的PDF没有页面'); }
    if (document.numPages > 500) { document.destroy(); throw new Error('生成的PDF超过500页，请拆分订单'); }
    releasePreviewDocument();
    state.previewDocument = document; state.previewMode = 'pdfjs'; state.previewPageCount = document.numPages; state.previewPage = 0;
    return normalized;
  }

  async function convertContractToPdf(blob, fileName) {
    if (typeof window.__CONTRACT_PDF_CONVERTER__ === 'function') {
      let lastError;
      for (let attempt = 0; attempt < 2; attempt += 1) {
        try {
          const result = await window.__CONTRACT_PDF_CONVERTER__(blob, fileName);
          return await normalizePdfFile(result, fileName.replace(/\.[^.]+$/, '.pdf'));
        } catch (error) { lastError = error; }
      }
      throw lastError || new Error('PDF转换测试接口没有返回有效文件');
    }
    await ensureConverterFrame();
    const buffer = await blob.arrayBuffer();
    await requestConverter('document:open-buffer', { fileName, buffer, readonly: false }, 'document:opened', [buffer]);
    let lastError;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      try {
        const result = await requestConverter('document:save', { targetExt: 'PDF' }, 'document:saved');
        return await normalizePdfFile(result.file, fileName.replace(/\.[^.]+$/, '.pdf'));
      } catch (error) {
        lastError = error;
        if (attempt === 0) await new Promise(resolve => setTimeout(resolve, 500));
      }
    }
    throw lastError || new Error('PDF转换组件没有返回有效文件');
  }

  async function normalizePdfFile(value, fileName) {
    let bytes;
    if (value && typeof value.arrayBuffer === 'function') bytes = new Uint8Array(await value.arrayBuffer());
    else if (value instanceof ArrayBuffer) bytes = new Uint8Array(value);
    else if (ArrayBuffer.isView(value)) bytes = new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
    else throw new Error('PDF转换组件没有返回文件');
    if (bytes.length < 100) throw new Error('PDF转换结果不完整，正在重试');
    const prefix = new TextDecoder('latin1').decode(bytes.slice(0, Math.min(bytes.length, 1024))), headerIndex = prefix.indexOf('%PDF-');
    if (headerIndex < 0) throw new Error('PDF转换结果格式无效，正在重试');
    if (headerIndex > 0) bytes = bytes.slice(headerIndex);
    return new File([bytes], fileName, { type: MIME.pdf, lastModified: Date.now() });
  }

  function ensureConverterFrame() {
    if (converterFrame?.dataset.ready === 'true') return Promise.resolve();
    if (converterReadyPromise) return converterReadyPromise;
    converterReadyPromise = new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        converterReadyPromise = null;
        reject(new Error('PDF转换组件加载超时，请检查网络后重试'));
      }, CONVERSION_TIMEOUT);
      converterReadyResolve = () => { clearTimeout(timeout); resolve(); };
      converterFrame = document.createElement('iframe');
      converterFrame.className = 'onlyoffice-converter-frame';
      converterFrame.title = '浏览器本地PDF转换组件'; converterFrame.tabIndex = -1; converterFrame.setAttribute('aria-hidden', 'true');
      converterFrame.src = `${ONLYOFFICE_ORIGIN}/editor?embed=1&locale=zh-CN&embedOrigin=${encodeURIComponent(location.origin)}`;
      converterFrame.addEventListener('error', () => { clearTimeout(timeout); converterReadyPromise = null; reject(new Error('PDF转换组件加载失败，请检查网络后重试')); }, { once: true });
      document.body.append(converterFrame);
    });
    return converterReadyPromise;
  }

  function requestConverter(type, payload, expectedType, transfer = []) {
    const id = `contract-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => { converterRequests.delete(id); reject(new Error('PDF转换超时，请重新生成')); }, CONVERSION_TIMEOUT);
      converterRequests.set(id, { expectedType, resolve, reject, timeout });
      converterFrame.contentWindow.postMessage({ id, type, payload }, ONLYOFFICE_ORIGIN, transfer);
    });
  }

  let converterReadyResolve = null;
  function handleConverterMessage(event) {
    if (event.origin !== ONLYOFFICE_ORIGIN || !event.data?.type?.startsWith('document:')) return;
    if (event.data.type === 'document:ready') {
      if (converterFrame) converterFrame.dataset.ready = 'true';
      converterReadyResolve?.(); converterReadyResolve = null;
      return;
    }
    const request = converterRequests.get(event.data.id);
    if (!request) return;
    if (event.data.type === 'document:error') {
      clearTimeout(request.timeout); converterRequests.delete(event.data.id);
      request.reject(new Error(`PDF转换失败：${event.data.payload?.message || '未知错误'}`));
    } else if (event.data.type === request.expectedType) {
      clearTimeout(request.timeout); converterRequests.delete(event.data.id); request.resolve(event.data.payload || {});
    }
  }

  async function generateDocx() {
    const zip = await JSZip.loadAsync(state.templateBytes), xml = await zip.file('word/document.xml').async('string'), doc = parseXml(xml);
    for (const [targetId, mapping] of mappedEntries()) {
      if (mapping.mode === 'detail') continue;
      const target = locateWordTarget(doc, targetId); if (!target) throw new Error(`模板位置已变化：${targetId}`);
      const value = resolveMapping(mapping);
      putWordText(target, mapping.preserveLabel ? labeledValue(mapping.labelText, value) : value);
    }
    if (state.detailRow) {
      const match = state.detailRow.match(/^word:(\d+):(\d+)$/); if (!match) throw new Error('Word明细模板行无效');
      const table = direct(all(doc, 'body')[0], 'tbl')[Number(match[1])], row = direct(table, 'tr')[Number(match[2])];
      if (!row) throw new Error('找不到Word明细模板行');
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      for (const [recordIndex, record] of detailRecords().entries()) {
        const clone = row.cloneNode(true), cells = direct(clone, 'tc');
        for (const [targetId, mapping] of rowMappings) {
          const cellIndex = Number(targetId.match(/:c:(\d+)$/)?.[1]);
          if (cells[cellIndex]) putWordText(cells[cellIndex], detailValue(mapping, record, recordIndex));
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
      const value = resolveMapping(mapping);
      writeSheetCell(context.sheetDoc, targetId.slice(2), mapping.preserveLabel ? labeledValue(mapping.labelText, value) : value, { forceNumber: NUMERIC_BUSINESS_FIELDS.has(mapping.businessKey) && !mapping.preserveLabel });
    }
    if (state.detailRow) {
      const rowNumber = Number(state.detailRow.split(':')[1]);
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      await expandExcelDetailRowXml(zip, context, rowNumber, detailRecords(), rowMappings);
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

  function writeSheetCell(doc, address, value, options) {
    const point = XLSX.utils.decode_cell(address), row = ensureSheetRow(doc, point.r + 1), cell = ensureRowCell(row, address);
    writeCellValue(cell, value, options);
  }

  function writeCellValue(cell, value, options = {}) {
    const doc = cell.ownerDocument, originalType = cell.getAttribute('t');
    direct(cell, 'f', S).forEach(item => item.remove()); direct(cell, 'v', S).forEach(item => item.remove()); direct(cell, 'is', S).forEach(item => item.remove());
    const number = C.parseNumber(value), preferText = ['s', 'str', 'inlineStr'].includes(originalType);
    if (number !== null && (options.forceNumber || !preferText)) {
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
      validateExpandableFormulas(context.sheetDoc, rowNumber);
      adjustFormulasForInsertedRows(context.sheetDoc, state.templateSheet, rowNumber, delta, prototype);
      await adjustOtherSheetFormulas(zip, context, state.templateSheet, rowNumber, delta);
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
      for (const cell of direct(clone, 'c', S)) {
        cell.setAttribute('r', replaceAddressRow(cellAddressOf(cell), targetRow));
        for (const formula of direct(cell, 'f', S)) {
          if (offset) formula.textContent = adjustCopiedFormula(formula.textContent || '', offset);
          clearFormulaCache(cell);
        }
      }
      for (const [targetId, mapping] of rowMappings) {
        const source = XLSX.utils.decode_cell(targetId.slice(2)), address = XLSX.utils.encode_cell({ r: targetRow - 1, c: source.c });
        writeCellValue(ensureRowCell(clone, address), detailValue(mapping, record, offset), { forceNumber: NUMERIC_BUSINESS_FIELDS.has(mapping.businessKey) });
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

  function validateExpandableFormulas(doc, rowNumber) {
    for (const formula of all(doc, 'f', S)) {
      const type = formula.getAttribute('t') || 'normal';
      const formulaRow = rowNumberFromAddress(cellAddressOf(formula.parentElement));
      const formulaRange = formula.getAttribute('ref') || '';
      if (formulaRow === rowNumber && type !== 'normal') {
        throw new Error(`Excel明细模板行含${formulaTypeLabel(type)}，无法按订单行复制：${cellAddressOf(formula.parentElement)}`);
      }
      if (formulaRange && rangeTouchesRow(formulaRange, rowNumber) && type !== 'normal') {
        throw new Error(`Excel公式区域穿过明细模板行，无法安全扩展：${formulaRange}`);
      }
    }
  }

  function formulaTypeLabel(type) {
    return ({ shared: '共享公式', array: '数组公式', dataTable: '模拟运算表公式' })[type] || `特殊公式（${type}）`;
  }

  function adjustFormulasForInsertedRows(doc, sheetName, rowNumber, delta, prototype) {
    for (const cell of all(doc, 'c', S)) {
      if (cell.parentElement === prototype) continue;
      for (const formula of direct(cell, 'f', S)) {
        formula.textContent = transformFormulaReferences(formula.textContent || '', (range, formulaSheet) => {
          if (formulaSheet && formulaSheet.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()) return range;
          return shiftRange(range, rowNumber, delta);
        });
        const formulaRange = formula.getAttribute('ref');
        if (formulaRange) formula.setAttribute('ref', shiftRange(formulaRange, rowNumber, delta));
        clearFormulaCache(cell);
      }
    }
  }

  async function adjustOtherSheetFormulas(zip, context, targetSheetName, rowNumber, delta) {
    for (const sheet of all(context.workbookDoc, 'sheet', S)) {
      if (sheet.getAttribute('name') === targetSheetName) continue;
      const relationId = sheet.getAttributeNS(R, 'id') || sheet.getAttribute('r:id');
      const relation = all(context.workbookRelsDoc, 'Relationship', P).find(item => item.getAttribute('Id') === relationId);
      if (!relation || !/\/worksheet$/i.test(relation.getAttribute('Type') || '')) continue;
      const sheetPath = resolveZipPath('xl', relation.getAttribute('Target')), file = zip.file(sheetPath);
      if (!file) continue;
      const doc = parseXml(await file.async('string')); let changed = false;
      for (const cell of all(doc, 'c', S)) {
        for (const formula of direct(cell, 'f', S)) {
          const original = formula.textContent || '';
          const updated = transformFormulaReferences(original, (range, formulaSheet) => {
            if (!formulaSheet || formulaSheet.toLocaleLowerCase() !== targetSheetName.toLocaleLowerCase()) return range;
            return shiftRange(range, rowNumber, delta);
          });
          if (updated === original) continue;
          formula.textContent = updated; clearFormulaCache(cell); changed = true;
        }
      }
      if (changed) zip.file(sheetPath, new XMLSerializer().serializeToString(doc));
    }
  }

  function adjustCopiedFormula(formula, offset) {
    return transformFormulaReferences(formula, range => range.split(':').map(reference => {
      const cell = parseCellRef(reference); if (!cell) return reference;
      if (!cell.rowAbsolute) cell.row += offset;
      return formatCellRef(cell);
    }).join(':'));
  }

  function transformFormulaReferences(formula, transform) {
    return C.text(formula).split(/("(?:[^"]|"")*")/g).map((part, index) => {
      if (index % 2) return part;
      return part.replace(/(^|[^A-Z0-9_.])((?:'(?:[^']|'')+'|[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![A-Z0-9_(])/gi, (match, prefix, sheetToken, start, end) => {
        const formulaSheet = normalizeFormulaSheet(sheetToken);
        const range = end ? `${start}:${end}` : start;
        return `${prefix}${sheetToken || ''}${transform(range, formulaSheet)}`;
      });
    }).join('');
  }

  function normalizeFormulaSheet(sheetToken) {
    if (!sheetToken) return '';
    const value = sheetToken.slice(0, -1);
    return value.startsWith("'") && value.endsWith("'") ? value.slice(1, -1).replace(/''/g, "'") : value;
  }

  function clearFormulaCache(cell) { direct(cell, 'v', S).forEach(item => item.remove()); }

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
  function detailValue(mapping, record, index) { return mapping.field === '@sequence' ? String(index + 1) : C.text(record[mapping.field]); }
  function labeledValue(label, value) {
    const source = C.text(label), output = C.text(value);
    if (/\{\{[^{}]+\}\}/.test(source)) return source.replace(/\{\{[^{}]+\}\}/g, output);
    const separator = source.search(/[：:]/);
    if (separator >= 0) return `${source.slice(0, separator + 1)}${output}`;
    return `${source}${source ? '：' : ''}${output}`;
  }

  function downloadOutput() {
    if (!state.output || !$('confirmExport').checked) return;
    downloadBlob(state.output, state.outputFileName);
  }

  function downloadPdf() {
    if (!state.outputPdf || !$('confirmExport').checked) return;
    downloadBlob(state.outputPdf, state.outputPdfFileName);
  }

  function updateExportButtons() {
    const ready = !!(state.output && state.outputPdf && $('confirmExport').checked);
    $('downloadContract').disabled = !ready; $('downloadPdf').disabled = !ready; updateSteps();
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob), anchor = document.createElement('a');
    anchor.href = url; anchor.download = fileName; document.body.append(anchor); anchor.click(); anchor.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
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
    return { version: 2, fingerprint: state.fingerprint, templateType: state.templateType, templateSheet: state.templateSheet || '', fieldSelections: state.fieldSelections, fieldStrategies: state.fieldStrategies };
  }

  function applyMappingPayload(payload) {
    if (payload?.version !== 2 || payload.fingerprint !== state.fingerprint || payload.templateType !== state.templateType || (payload.templateSheet || '') !== (state.templateSheet || '')) throw new Error('映射文件与当前合同模板不完全一致');
    const fields = new Set(state.order?.headers || []), selections = {};
    for (const definition of BUSINESS_FIELDS) {
      const value = payload.fieldSelections?.[definition.key];
      if (definition.automatic) selections[definition.key] = '@sequence';
      else if (fields.has(value)) selections[definition.key] = value;
    }
    state.fieldSelections = selections; state.fieldStrategies = { ...(payload.fieldStrategies || {}) }; state.mappingSignature = '';
  }

  function mappingKey() { return state.fingerprint ? `gylsjqx-contract-mapping:${state.fingerprint}:${state.templateSheet || 'docx'}` : ''; }
  function saveMapping() { try { const key = mappingKey(); if (key) localStorage.setItem(key, JSON.stringify(mappingPayload())); } catch (_) {} }
  function restoreMapping() {
    try { const key = mappingKey(), raw = key && localStorage.getItem(key); if (raw) applyMappingPayload(JSON.parse(raw)); }
    catch (_) { state.mappings = {}; state.detailRow = ''; state.fieldSelections = {}; state.fieldStrategies = {}; state.mappingSignature = ''; }
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
