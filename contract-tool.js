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
  const NUMERIC_BUSINESS_FIELDS = new Set(['sequence', 'quantity', 'taxUnitPrice', 'taxAmount', 'taxRate']);
  const RIGHT_SIDE_VALUE_FIELDS = new Set(['contractNumber', 'deliveryPlace', 'deliveryTime', 'signDate']);
  const ORDER_RIGHT_SIDE_FIELDS = new Set(['contractNumber', 'deliveryTime', 'taxTotalLower', 'taxTotalUpper']);
  const ORDER_FIELD_KEYS = new Set(['contractNumber', 'sequence', 'materialCode', 'materialName', 'specification', 'sku', 'unit', 'quantity', 'taxUnitPrice', 'remark', 'deliveryTime', 'taxTotalLower', 'taxTotalUpper']);
  const CONTRACT_TERMS = [
    '1、采购合同所述价格为甲方在本合同项下应向乙方支付的最终价格，其中已经包括所有的安装费、售后服务费和税费等，除合同金额外，甲方不再支付任何其他费用。',
    '2、乙方应根据甲方要求进行包装并确保产品交付给甲方时包装完好无损，按照甲方要求进行必要的标识贴附工作并承担相关费用。',
    '3、乙方应按时将产品及时足额送货至甲方指定地点，同时附带产品合格证及出厂检验报告及其他应提供的相应材料，否则甲方有权拒绝收货。因乙方原因造成延迟交货、不能交货的，乙方应承担相应责任。甲方在交货地依据出厂标准或双方约定对合同产品进行检验并签署验收单，若产品与本合同约定不符，乙方应免费更换或退货退款并承担相应费用。',
    '4、乙方保证产品质量符合相关标准及双方相关约定。乙方产品保修期为 / 年，保修期内产生产品质量问题的，乙方应于甲方指定时间及地点接收产品，并免费进行产品部件或整体保修。产品出货后因人为损坏原因或超出保修期限，有偿维修期为 / 年，乙方提供成本价有偿维修，对修复产品提供与新品相同的质量保证。',
    '5、乙方向甲方交付产品并经甲方检验合格后，产品所有权及风险转移到甲方。乙方保证对交付给甲方的产品拥有合法的所有权、知识产权及其他权益，保证不侵犯任何第三方的合法权利。否则，由此产生的一切责任由乙方承担。',
    '6、乙方应对所接触到的甲方的保密信息承担保密义务，未经甲方事先书面同意不得将甲方保密信息对外提供、披露。',
    '7、任何一方违反本合同约定的，应按照本合同总价款30%的标准向守约方支付违约金。如违约金不足以弥补守约方的实际损失，违约方仍应向守约方赔偿其全部损失。',
    '8、因本合同发生争议的，应友好协商解决；协商不成的，任何一方均有权向甲方所在地有管辖权的人民法院提起诉讼。',
    '9、本合同一式二份，双方各执一份，具有同等法律效力。本合同自双方加盖公章或合同专用章之日起生效，扫描件、复印件与本合同原件具有同等法律效力。',
  ];
  const CONTRACT_TERMS_TEXT = CONTRACT_TERMS.join('\n');
  const CONTRACT_TERMS_MARKERS = ['乙方应根据甲方要求进行包装', '产品所有权及风险转移到甲方', '有权向甲方所在地有管辖权的人民法院提起诉讼', '本合同一式二份'];
  const BUSINESS_FIELDS = [
    { key: 'sequence', label: '序号', aliases: ['序号', '行号'], kind: 'detail', automatic: '@sequence', automaticLabel: '自动生成 1、2、3…', writeMode: '按订单逐行写入' },
    { key: 'materialCode', label: '物料编码', aliases: ['物料编码', '产品编码', '商品编码', '货号'], kind: 'detail' },
    { key: 'materialName', label: '物料名称', aliases: ['名称', '物料名称', '物料', '产品名称', '商品名称', '品名'], kind: 'detail' },
    { key: 'specification', label: '规格型号', aliases: ['规格型号', '规格', '型号'], kind: 'detail' },
    { key: 'sku', label: 'SKU', aliases: ['SKU', 'SKU编码'], kind: 'detail' },
    { key: 'unit', label: '单位', aliases: ['单位', '计量单位'], kind: 'detail' },
    { key: 'quantity', label: '数量', aliases: ['数量', '采购数量', '订单数量'], kind: 'detail' },
    { key: 'taxUnitPrice', label: '含税运单价（元）', aliases: ['含税运单价（元）', '含税运单价', '含税单价（元）', '含税单价', '单价'], kind: 'detail' },
    { key: 'taxAmount', label: '含税运总金额（元）', aliases: ['含税运总金额（元）', '含税运总金额', '含税总金额（元）', '含税总金额', '含税金额', '金额'], kind: 'detail', automatic: '@line-amount', automaticLabel: '数量 × 含税单价', writeMode: '自动计算（两位小数）' },
    { key: 'taxRate', label: '税率', aliases: ['税率', '增值税率'], kind: 'detail' },
    { key: 'deliveryTime', label: '交货时间', aliases: ['交货时间', '交期', '要求货好时间', '要求交货日期'], kind: 'single' },
    { key: 'remark', label: '备注', aliases: ['备注', '说明'], kind: 'detail' },
    { key: 'taxTotalLower', label: '含税运合计（小写）', aliases: ['含税运合计（小写）', '含税运合计小写', '合计（小写）', '合计小写', '小写合计', '人民币小写', '人民币小写金额'], kind: 'single', automatic: '@tax-total-lower', automaticLabel: '逐行计算数量 × 含税单价后汇总', writeMode: '自动汇总' },
    { key: 'taxTotalUpper', label: '含税运合计（大写）', aliases: ['含税运合计（大写）', '含税运合计大写', '合计（大写）', '合计大写', '大写合计', '人民币大写', '人民币大写金额'], kind: 'single', automatic: '@tax-total-upper', automaticLabel: '由小写合计自动转人民币大写', writeMode: '自动转大写' },
    { key: 'contractNumber', label: '合同编号', aliases: ['合同编号', '合同编码', '合同号'], kind: 'single' },
    { key: 'orderNumber', label: '订单编号', aliases: ['订单编号', '采购订单号', '采购单号', '订单号'], kind: 'single' },
    { key: 'buyer', label: '采购方（甲方）', aliases: ['采购方（甲方）', '采购方', '甲方', '买方'], kind: 'single' },
    { key: 'supplier', label: '供应商（乙方）', aliases: ['供应方（乙方）', '供应方', '供应商名称', '供应商（乙方）', '供应商', '乙方', '卖方'], kind: 'single' },
    { key: 'signDate', label: '甲方签署日期', aliases: ['甲方签署日期', '甲方签署时间', '签署时间', '签署日期', '签订日期', '签订时间', '签字日期', '合同日期', '签约日期'], kind: 'single', automatic: '@sign-date', automaticLabel: '请选择日期', manualInput: 'date', writeMode: '选择后填充' },
    { key: 'deliveryPlace', label: '交货地点', aliases: ['交货地点', '送货地址', '交付地点'], kind: 'single' },
    { key: 'paymentTerms', label: '付款方式', aliases: ['付款方式', '付款条件', '结算方式', '结算条件'], kind: 'single' },
  ];
  const ACTIVE_FIELDS = BUSINESS_FIELDS.filter(field => ORDER_FIELD_KEYS.has(field.key) || ['taxAmount', 'taxTotalLower', 'taxTotalUpper', 'signDate'].includes(field.key));
  let pdfJsPromise = null;
  let legacyPdfJsPromise = null;
  let pdfiumPromise = null;
  let converterFrame = null;
  let converterReadyPromise = null;
  let previewResizeTimer = null;
  const converterRequests = new Map();
  const state = {
    orderFile: null, orderBytes: null, orderBook: null, orderSheet: '', order: null,
    templateFile: null, templateBytes: null, templateType: '', templateBook: null, templateSheet: '', templateModel: null, templateFeatures: [], fingerprint: '',
    previewDocument: null, previewPdfiumDocument: null, previewMode: '', previewPageCount: 0, previewPage: 0, previewRenderToken: 0,
    mappings: {}, detailRow: '', fieldSelections: {}, fieldStrategies: {}, manualValues: {}, templateBindings: {}, mappingSignature: '', customOutputName: false, output: null, outputFileName: '', outputPdf: null, outputPdfFileName: '', busy: false,
  };
  const $ = id => document.getElementById(id);
  const esc = value => C.text(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const direct = (element, name, namespace = W) => [...(element?.children || [])].filter(child => child.namespaceURI === namespace && child.localName === name);
  const all = (element, name, namespace = W) => [...(element?.getElementsByTagNameNS(namespace, name) || [])];

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    initTheme();
    $('orderFile').addEventListener('change', event => { const file = event.target.files[0]; event.target.value = ''; loadOrder(file); });
    $('templateFile').addEventListener('change', event => { const file = event.target.files[0]; event.target.value = ''; loadTemplate(file); });
    $('orderSheet').addEventListener('change', event => { state.orderSheet = event.target.value; state.fieldSelections = {}; state.fieldStrategies = {}; state.mappingSignature = ''; analyzeOrder(); invalidateOutput(); updateAll(); });
    $('templateSheet').addEventListener('change', event => { state.templateSheet = event.target.value; buildExcelModel(); state.mappings = {}; state.detailRow = ''; state.templateBindings = {}; state.mappingSignature = ''; invalidateOutput(); restoreMapping(); updateAll(); });
    $('resultPrevious').addEventListener('click', () => changePreviewPage(-1));
    $('resultNext').addEventListener('click', () => changePreviewPage(1));
    $('outputName').addEventListener('input', () => { state.customOutputName = true; invalidateOutput(); updateConfirmation(); });
    $('outputParties').addEventListener('input', () => { state.customOutputName = false; invalidateOutput(); updateConfirmation(); });
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
      if (!state.previewDocument && !state.previewPdfiumDocument) return;
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
      state.orderFile = file; state.orderBytes = bytes; state.orderBook = book; state.orderSheet = book.SheetNames[0]; state.manualValues = {};
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
    // Anchor at A1 so matrix indexes agree with absolute merged-cell coordinates.
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: true, range: { s: { r: 0, c: 0 }, e: range.e } });
    state.order = C.analyzeMatrix(matrix, 1000);
    const rightSideValues = C.extractAdjacentLabelValues(
      matrix,
      Object.fromEntries(BUSINESS_FIELDS.map(field => [field.key, field.key === 'taxTotalLower' ? [...field.aliases, '含税合计'] : field.aliases])),
      sheet['!merges'] || [],
      [state.order.headerIndex],
    );
    state.order.sourceTotals = { lower: rightSideValues.taxTotalLower || '', upper: rightSideValues.taxTotalUpper || '' };
    state.order.rightSideFields = {};
    for (const definition of BUSINESS_FIELDS.filter(field => ORDER_RIGHT_SIDE_FIELDS.has(field.key) && !field.manualInput)) {
      const value = rightSideValues[definition.key];
      if (!C.text(value).trim()) continue;
      const header = `${definition.label}（右侧内容）`;
      state.order.headers.push(header);
      state.order.rows.forEach(record => { record[header] = value; });
      state.order.rightSideFields[definition.key] = header;
    }
    const allowedHeaders = new Set(BUSINESS_FIELDS.filter(field => ORDER_FIELD_KEYS.has(field.key)).flatMap(field => field.aliases.map(normalizeBusinessLabel)));
    const adjacentHeaders = new Set(Object.values(state.order.rightSideFields));
    state.order.headers = state.order.headers.filter(header => allowedHeaders.has(normalizeBusinessLabel(header)) || adjacentHeaders.has(header));
    const selections = Object.fromEntries(BUSINESS_FIELDS.filter(field => ORDER_FIELD_KEYS.has(field.key)).map(field => [field.key, bestOrderHeader(field)]));
    state.order.rows = C.selectDetailRows(state.order.rows, state.order.headers, selections).map(row => Object.fromEntries([['_row', row._row], ...state.order.headers.map(header => [header, row[header]])]));
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
      state.templateBook = null; state.templateSheet = ''; state.mappings = {}; state.detailRow = ''; state.fieldSelections = {}; state.fieldStrategies = {}; state.manualValues = {}; state.templateBindings = {}; state.mappingSignature = '';
      if (type === 'docx') await parseDocx(zip);
      else parseXlsx();
      $('templateFileName').textContent = file.name; $('templateDrop').classList.add('loaded');
      $('outputParties').value = file.name.replace(/\.[^.]+$/, '').replace(/[（(][^）)]*[）)]/g, '').replace(/-生成合同$/, '').trim();
      state.customOutputName = false;
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

  function loadLegacyPdfJs() {
    if (!legacyPdfJsPromise) {
      legacyPdfJsPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = new URL('vendor/pdf-legacy-3.11.174.min.js', location.href).href;
        script.onload = () => {
          const pdfjs = window.pdfjsLib;
          if (!pdfjs?.getDocument || !C.text(pdfjs.version).startsWith('3.')) {
            reject(new Error('兼容PDF解析器加载异常'));
            return;
          }
          pdfjs.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf-legacy-3.11.174.worker.min.js', location.href).href;
          resolve(pdfjs);
        };
        script.onerror = () => reject(new Error('兼容PDF解析器加载失败'));
        document.head.append(script);
      });
    }
    return legacyPdfJsPromise;
  }

  function loadPdfium() {
    if (!pdfiumPromise) {
      pdfiumPromise = Promise.all([
        import('./vendor/pdfium-2.15.1.mjs'),
        fetch(new URL('vendor/pdfium-2.15.1.wasm', location.href)).then(response => {
          if (!response.ok) throw new Error('PDFium预览引擎加载失败');
          return response.arrayBuffer();
        }),
      ]).then(async ([module, wasmBinary]) => {
        const pdfium = await module.init({ wasmBinary });
        pdfium.PDFiumExt_Init();
        return pdfium;
      });
    }
    return pdfiumPromise;
  }

  function releasePreviewDocument() {
    state.previewRenderToken += 1;
    state.previewDocument?.destroy?.();
    state.previewDocument = null;
    state.previewPdfiumDocument?.close?.();
    state.previewPdfiumDocument = null;
    state.previewMode = '';
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
    const selectedCount = ACTIVE_FIELDS.filter(field => field.automatic || state.fieldSelections[field.key]).length;
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

  function inlineTotalTemplate(definition, value) {
    const source = C.text(value);
    const pattern = definition.key === 'taxTotalLower'
      ? /^([\s\S]*?人民币小写\s*[：:]?)[\s\u00a0]*(元)\s*$/
      : definition.key === 'taxTotalUpper'
        ? /^([\s\S]*?人民币大写\s*[：:]?)[\s\u00a0]*([圆元]整)\s*$/
        : null;
    const match = pattern?.exec(source);
    return match ? { prefix: match[1].trimEnd(), suffix: match[2] } : null;
  }

  function inlineSingleTemplate(definition, value) {
    if (!['contractNumber', 'signDate'].includes(definition.key)) return null;
    const source = C.text(value);
    // Only replace a recognizable value slot; preserve surrounding signature text.
    for (const alias of [...definition.aliases].sort((a, b) => b.length - a.length)) {
      const label = [...alias].join('\\s*');
      const match = new RegExp(`(${label}[ \t]*[：:]?)([^\\n\\r]*)`).exec(source);
      if (!match) continue;
      const tail = match[2].trim();
      const placeholder = /^[\s_＿—－.·…/年月日{}【】\[\]（）()0-9-]*$/.test(tail) || /^(待填写?|请选择|填写日期|填写时间)$/.test(tail);
      const identifier = definition.key === 'contractNumber' && /^[A-Za-z0-9][A-Za-z0-9_./-]*$/.test(tail);
      if (!placeholder && !identifier) continue;
      return { prefix: source.slice(0, match.index) + match[1].trimEnd() + (/[：:]$/.test(match[1].trimEnd()) ? '' : '：'), suffix: source.slice(match.index + match[0].length), hasValue: !!tail };
    }
    return null;
  }

  function bestDefinition(value, kind = '') {
    return BUSINESS_FIELDS.filter(field => !kind || field.kind === kind).map(field => ({ field, score: fieldMatchScore(value, field) })).sort((a, b) => b.score - a.score)[0];
  }

  function bestOrderHeader(definition) {
    const rightSideHeader = state.order.rightSideFields?.[definition.key];
    if (rightSideHeader) return rightSideHeader;
    const best = state.order.headers.map(header => ({ header, score: fieldMatchScore(header, definition) })).sort((a, b) => b.score - a.score)[0];
    return best?.score > 0 ? best.header : '';
  }

  function detailRecords() {
    return C.selectDetailRows(state.order?.rows || [], state.order?.headers || [], state.fieldSelections);
  }

  function detailTemplateRowCount() {
    if (!state.detailRow) return 1;
    const sequenceTargetId = mappedEntries().find(([, mapping]) => mapping.mode === 'detail' && mapping.businessKey === 'sequence')?.[0]
      || state.templateBindings?.sequence?.targetId;
    const sequenceTarget = sequenceTargetId ? findTarget(sequenceTargetId) : null;
    return C.reservedDetailRowCount(templateRows(), state.detailRow, sequenceTarget?.cellIndex || 0);
  }

  function templateRows() {
    if (state.templateModel.type === 'docx') return state.templateModel.blocks.filter(block => block.type === 'table').flatMap(block => block.rows);
    return state.templateModel.rows;
  }

  function locateRow(rowKey) { return templateRows().find(row => row.rowKey === rowKey); }

  function nextWritableTemplateCell(row, labelCell) {
    if (state.templateModel.type !== 'xlsx') return row?.cells.find(cell => cell.cellIndex === labelCell.cellIndex + 1);
    const sheet = state.templateBook?.Sheets[state.templateSheet], rowIndex = labelCell.rowIndex;
    const labelMerge = (sheet?.['!merges'] || []).find(range => rowIndex >= range.s.r && rowIndex <= range.e.r && labelCell.cellIndex >= range.s.c && labelCell.cellIndex <= range.e.c);
    const nextColumn = (labelMerge?.e.c ?? labelCell.cellIndex) + 1;
    const candidateMerge = (sheet?.['!merges'] || []).find(range => rowIndex >= range.s.r && rowIndex <= range.e.r && nextColumn >= range.s.c && nextColumn <= range.e.c);
    const targetColumn = candidateMerge?.s.c ?? nextColumn;
    return row?.cells.find(cell => cell.cellIndex === targetColumn);
  }

  function templatePartyForTarget(target) {
    const ownText = normalizeBusinessLabel(target?.value);
    if (ownText.includes('甲方') && ownText.includes('乙方')) return 'ambiguous';
    if (ownText.includes('甲方')) return 'buyer';
    if (ownText.includes('乙方')) return 'supplier';
    const rows = templateRows(), rowIndex = rows.findIndex(row => row.rowKey === target?.rowKey);
    const scope = C.text(target?.rowKey).replace(/:\d+$/, '');
    for (let index = rowIndex - 1, distance = 1; index >= 0 && distance <= 3; index -= 1, distance += 1) {
      const row = rows[index];
      if (!C.text(row?.rowKey).startsWith(`${scope}:`)) break;
      const nearby = [...(row.cells || [])].sort((left, right) => Math.abs(left.cellIndex - target.cellIndex) - Math.abs(right.cellIndex - target.cellIndex));
      for (const cell of nearby) {
        if (Math.abs(cell.cellIndex - target.cellIndex) > 1) continue;
        const text = normalizeBusinessLabel(cell.value);
        if (text.includes('甲方')) return 'buyer';
        if (text.includes('乙方')) return 'supplier';
      }
    }
    return '';
  }

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
    for (const definition of ACTIVE_FIELDS.filter(field => field.kind === 'single')) {
      let best = null;
      for (const labelTarget of state.templateModel.targets) {
        if (excludedRows.has(labelTarget.rowKey)) continue;
        let score = fieldMatchScore(labelTarget.value, definition);
        if (score <= 0) continue;
        if (definition.key === 'signDate') {
          const party = templatePartyForTarget(labelTarget);
          if (party === 'supplier' || party === 'ambiguous') continue;
          if (party === 'buyer') score += 1000;
        }
        let inline = inlineTotalTemplate(definition, labelTarget.value);
        const row = locateRow(labelTarget.rowKey), next = nextWritableTemplateCell(row, labelTarget);
        const canUseNext = next && !usedTargets.has(next.id) && (!bestDefinition(next.value)?.score || /待填|填写|空白/.test(C.text(next.value)));
        const inlineSingle = inlineSingleTemplate(definition, labelTarget.value);
        if (inlineSingle && (!canUseNext || inlineSingle.hasValue)) inline = inlineSingle;
        if (RIGHT_SIDE_VALUE_FIELDS.has(definition.key) && !canUseNext && !inline) continue;
        if (canUseNext && !inline) score += 10;
        const target = inline ? labelTarget : (canUseNext ? next : labelTarget);
        if (usedTargets.has(target.id)) continue;
        if (!best || score > best.score) best = { labelTarget, target, score, inline };
      }
      if (!best) continue;
      const { labelTarget, target } = best;
      bindings[definition.key] = {
        targetId: target.id,
        mode: 'single',
        preserveLabel: target.id === labelTarget.id,
        labelText: labelTarget.value,
        inlinePrefix: best.inline?.prefix || '',
        inlineSuffix: best.inline?.suffix || '',
        templateLabel: best.inline ? `${labelTarget.value} → 标签后填写` : (RIGHT_SIDE_VALUE_FIELDS.has(definition.key) ? `${labelTarget.value} → 右侧填写位置` : labelTarget.value),
      };
      usedTargets.add(target.id);
    }
    state.templateBindings = bindings;
  }

  function ensureBusinessMappings() {
    const signature = `${state.fingerprint}:${state.templateSheet}:${state.orderSheet}:${state.order.headers.join('|')}`;
    if (state.mappingSignature !== signature) {
      for (const definition of ACTIVE_FIELDS) {
        if (definition.automatic) state.fieldSelections[definition.key] = definition.automatic;
        else if (!state.fieldSelections[definition.key] || !state.order.headers.includes(state.fieldSelections[definition.key])) state.fieldSelections[definition.key] = bestOrderHeader(definition);
      }
      state.mappingSignature = signature;
    }
    detectTemplateBindings(); rebuildMappings();
  }

  function rebuildMappings() {
    const mappings = {};
    for (const definition of ACTIVE_FIELDS) {
      const binding = state.templateBindings[definition.key], field = state.fieldSelections[definition.key];
      if (!binding || !field) continue;
      const values = field.startsWith('@') ? [] : C.distinctValues(state.order.rows, field);
      mappings[binding.targetId] = {
        field, mode: binding.mode, businessKey: definition.key, preserveLabel: !!binding.preserveLabel, labelText: binding.labelText || '',
        inlinePrefix: binding.inlinePrefix || '', inlineSuffix: binding.inlineSuffix || '',
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
    const termsSummary = hasStandardContractTermsTemplate() ? '；标准合同条款 1—9 已识别，生成时将完整写入' : '';
    $('templateDetectionSummary').textContent = `模板自动识别 ${detected} 个可写字段；序号按 ${detailRows.length} 条物料明细自动生成${termsSummary}`;
    $('businessMappingRows').innerHTML = ACTIVE_FIELDS.map(definition => {
      const selection = state.fieldSelections[definition.key] || '', binding = state.templateBindings[definition.key];
      const values = selection && selection !== '@sequence' ? C.distinctValues(state.order.rows, selection) : [];
      const conflict = definition.kind === 'single' && values.length > 1;
      const options = definition.automatic
        ? `<option value="${definition.automatic}">${esc(definition.automaticLabel)}</option>`
        : `<option value="">不填写</option>${state.order.headers.map(header => `<option value="${esc(header)}" ${header === selection ? 'selected' : ''}>${esc(header)}</option>`).join('')}`;
      const inputControl = definition.manualInput === 'date'
        ? `<input class="business-field-select business-date-input" type="date" data-manual-key="${definition.key}" value="${esc(state.manualValues?.[definition.key] || '')}" aria-label="请选择${esc(definition.label)}" />`
        : `<select class="business-field-select" data-field-key="${definition.key}" ${definition.automatic ? 'disabled' : ''} aria-label="${esc(definition.label)}对应订单列">${options}</select>`;
      const strategy = conflict ? `<select class="business-field-select business-strategy-select" data-strategy-key="${definition.key}" aria-label="${esc(definition.label)}多值处理"><option value="">该列有多个值，请选择处理方式</option><option value="first" ${state.fieldStrategies[definition.key] === 'first' ? 'selected' : ''}>取第一条非空值</option><option value="merge" ${state.fieldStrategies[definition.key] === 'merge' ? 'selected' : ''}>合并去重值</option><option value="sum" ${state.fieldStrategies[definition.key] === 'sum' ? 'selected' : ''}>求和</option></select>` : '';
      const success = !!(binding && selection && mapped.has(definition.key));
      const sourceTotal = definition.key === 'taxTotalLower' ? state.order.sourceTotals?.lower : definition.key === 'taxTotalUpper' ? state.order.sourceTotals?.upper : undefined;
      const sourceNote = sourceTotal !== undefined ? `<small class="business-source-total">原表核对值：${esc(sourceTotal || '未提供')}</small>` : '';
      const templateField = binding?.templateLabel || (definition.manualInput === 'date' ? '未识别到日期填写位置，所选日期暂不写入模板' : '未识别到对应字段');
      return `<div class="business-mapping-row" role="row" data-business-key="${definition.key}">
        <div class="business-field-name" role="cell" data-cell-label="映射字段"><strong>${esc(definition.label)}</strong><small>${definition.automatic ? esc(definition.automaticLabel) : esc(definition.aliases.slice(0, 3).join('、'))}</small></div>
        <div class="business-order-field" role="cell" data-cell-label="订单明细">${inputControl}${strategy}${sourceNote}</div>
        <span class="business-template-field" role="cell" data-cell-label="合同模板">${esc(templateField)}</span>
        <span class="template-detection ${success ? 'detected' : ''}" role="cell"><span class="mobile-cell-label" aria-hidden="true">是否识别：</span><span class="detection-text">${success ? '识别成功' : '未识别'}</span></span>
        <span class="business-write-mode" role="cell" data-cell-label="写入方式">${esc(definition.writeMode || (definition.kind === 'detail' ? '按订单逐行写入' : '填充内容'))}</span>
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
    $('businessMappingRows').querySelectorAll('[data-manual-key]').forEach(input => input.addEventListener('input', event => {
      state.manualValues[event.target.dataset.manualKey] = event.target.value;
      // Keep the native date input mounted while the calendar/keyboard is active.
      invalidateOutput(); rebuildMappings(); updateConfirmation(); updateSteps();
    }));
  }

  async function renderGeneratedPdf() {
    if (!state.previewDocument && !state.previewPdfiumDocument) return;
    state.previewPage = Math.max(0, Math.min(state.previewPage, state.previewPageCount - 1));
    const pageNumber = state.previewPage + 1, renderToken = ++state.previewRenderToken;
    $('generatedPreview').innerHTML = `<div class="pdf-preview-shell"><canvas class="pdf-preview-canvas" data-page="${pageNumber}" aria-label="生成合同PDF第${pageNumber}页"></canvas><span class="pdf-render-status">正在渲染第 ${pageNumber} 页…</span></div>`;
    $('resultPagination').hidden = state.previewPageCount <= 1;
    $('resultPageLabel').textContent = `第 ${pageNumber} / ${state.previewPageCount} 页`;
    $('resultPrevious').disabled = state.previewPage === 0;
    $('resultNext').disabled = state.previewPage >= state.previewPageCount - 1;
    $('resultPagination').dataset.totalPages = String(state.previewPageCount);
    $('resultPagination').dataset.currentPage = String(state.previewPage + 1);
    try {
      const host = $('generatedPreview'), shell = host.querySelector('.pdf-preview-shell'), canvas = host.querySelector('.pdf-preview-canvas');
      const availableWidth = Math.min(794, Math.max(260, host.clientWidth - 48));
      if (state.previewMode === 'pdfium') {
        const dimensions = await state.previewPdfiumDocument.renderPage(state.previewPage, availableWidth, canvas);
        canvas.dataset.renderer = 'pdfium';
        shell.style.width = `${Math.ceil(dimensions.width)}px`; shell.style.height = `${Math.ceil(dimensions.height)}px`;
      } else {
        const pdfPage = await state.previewDocument.getPage(pageNumber);
        if (renderToken !== state.previewRenderToken) return;
        const baseViewport = pdfPage.getViewport({ scale: 1 });
        const cssScale = availableWidth / baseViewport.width, pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
        const cssViewport = pdfPage.getViewport({ scale: cssScale }), renderViewport = pdfPage.getViewport({ scale: cssScale * pixelRatio });
        canvas.width = Math.ceil(renderViewport.width); canvas.height = Math.ceil(renderViewport.height);
        canvas.style.width = `${Math.ceil(cssViewport.width)}px`; canvas.style.height = `${Math.ceil(cssViewport.height)}px`;
        shell.style.width = `${Math.ceil(cssViewport.width)}px`; shell.style.height = `${Math.ceil(cssViewport.height)}px`;
        await pdfPage.render({ canvasContext: canvas.getContext('2d', { alpha: false }), viewport: renderViewport }).promise;
        pdfPage.cleanup();
      }
      if (renderToken !== state.previewRenderToken) return;
      canvas.dataset.rendered = 'true'; shell.querySelector('.pdf-render-status')?.remove();
    } catch (error) {
      if (renderToken !== state.previewRenderToken) return;
      $('generatedPreview').innerHTML = `<div class="pdf-render-error">生成PDF第 ${pageNumber} 页渲染失败：${esc(error.message || '未知错误')}</div>`;
    }
  }

  function changePreviewPage(offset) {
    if (!state.previewDocument && !state.previewPdfiumDocument) return;
    const total = state.previewPageCount || 1, next = Math.max(0, Math.min(state.previewPage + offset, total - 1));
    if (next === state.previewPage) return;
    state.previewPage = next; renderGeneratedPdf(); $('generatedPreview').scrollTop = 0; $('generatedPreview').scrollLeft = 0;
  }

  function mappedEntries() { return Object.entries(state.mappings).filter(([, mapping]) => mapping?.field); }
  function findTarget(id) { return state.templateModel?.targets.find(target => target.id === id); }

  function updateAutomaticOutputName() {
    if (state.customOutputName) return;
    let amount = '', number = '';
    try { amount = contractTotals().lower; } catch (_) { /* Invalid amounts are reported by confirmation validation. */ }
    const field = state.fieldSelections.contractNumber;
    if (field) {
      try { number = C.resolveField(state.order.rows, field, state.fieldStrategies.contractNumber || 'first', ''); } catch (_) { /* No fabricated number on conflicting records. */ }
    }
    const parts = [$('outputParties').value.trim(), state.manualValues.signDate || '', amount, number];
    $('outputName').value = parts.map(value => C.text(value).trim()).filter(Boolean).join('-');
  }

  function updateConfirmation() {
    if (!state.order || !state.templateModel) { $('generateContract').disabled = true; return; }
    updateAutomaticOutputName();
    const issues = C.mappingIssues(state.order.rows, state.mappings, state.detailRow);
    if (state.detailRow && !mappedEntries().some(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow)) issues.push('明细模板行还没有映射任何订单字段');
    const notDetected = ACTIVE_FIELDS.filter(field => state.fieldSelections[field.key] && !state.templateBindings[field.key]).map(field => field.label);
    const outputName = C.sanitizeFileName($('outputName').value);
    if (!$('outputName').value.trim()) issues.push('请填写合同文件名称');
    if (state.templateBindings.signDate && !C.text(state.manualValues?.signDate).trim()) issues.push('请选择甲方签署日期，选择后才能生成PDF预览');
    const detailRows = detailRecords();
    if (state.templateBindings.taxAmount || state.templateBindings.taxTotalLower || state.templateBindings.taxTotalUpper) {
      try { contractTotals(); } catch (error) { issues.push(error.message); }
    }
    const summaryRows = [
      ['订单工作表', state.orderSheet],
      ['物料明细', `识别 ${detailRows.length} 条物料明细${detailRows.length !== state.order.rows.length ? `（原表 ${state.order.rows.length} 行）` : ''}`],
      ['合同模板', state.templateFile.name],
      ...(state.templateSheet ? [['合同工作表', state.templateSheet]] : []),
      ['输出文件', `${outputName}.${state.templateType}`],
    ];
    $('confirmSummary').innerHTML = summaryRows.map(([label, value]) => `<div class="confirm-summary-row"><dt>${esc(label)}</dt><dd>${esc(value)}</dd></div>`).join('');
    let totalWarnings = [];
    try { totalWarnings = C.compareSourceTotals(state.order.sourceTotals, contractTotals()); } catch (_) {}
    const sourceWarnings = totalWarnings.length ? `<div class="warnings">${totalWarnings.map(esc).join('<br>')}；合同仍按数量 × 含税单价计算。</div>` : '';
    const warnings = notDetected.length ? `<div class="warnings">模板中未识别：${esc(notDetected.join('、'))}；这些字段本次不会写入。</div>` : '';
    $('contractIssues').innerHTML = `${issues.length ? `<ul>${issues.map(issue => `<li>${esc(issue)}</li>`).join('')}</ul>` : '<div class="ready">映射检查通过，可以生成PDF预览。</div>'}${warnings}${sourceWarnings}`;
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
      const previewText = `${state.previewPageCount}页，${state.previewMode === 'legacy' ? '兼容渲染，' : state.previewMode === 'pdfium' ? 'PDFium完整渲染，' : ''}整页显示`;
      saveMapping(); status(`PDF预览已生成：${state.outputPdfFileName}（${previewText}）`, 'success'); toast('PDF预览生成完成，请核对后确认导出', 'success'); updateSteps();
      $('resultStage').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) { status(error.message, 'error'); toast(error.message, 'error'); }
    finally { state.busy = false; $('generateContract').textContent = '生成PDF预览'; updateConfirmation(); }
  }

  async function preparePdfPreview(file, allowCompatPreview) {
    let normalized = await normalizePdfFile(file, state.outputFileName.replace(/\.[^.]+$/, '.pdf'));
    const pdfjs = await loadPdfJs();
    const forceLegacyPreview = !!window.__CONTRACT_FORCE_LEGACY_PDF_PREVIEW__;
    const forcePdfiumPreview = !!window.__CONTRACT_FORCE_PDFIUM_PREVIEW__;
    let usedLegacyRenderer = false;
    let result = forceLegacyPreview || forcePdfiumPreview
      ? { document: null, lastError: new Error('测试兼容预览') }
      : await parsePdfDocument(pdfjs, await normalized.arrayBuffer());
    if (!result.document && window.PDFLib?.PDFDocument) {
      try {
        const repaired = await window.PDFLib.PDFDocument.load(await normalized.arrayBuffer(), { ignoreEncryption: false, updateMetadata: false });
        const repairedBytes = await repaired.save({ useObjectStreams: false, addDefaultPage: false, updateFieldAppearances: false });
        normalized = await normalizePdfFile(repairedBytes, normalized.name);
        if (!forceLegacyPreview && !forcePdfiumPreview) result = await parsePdfDocument(pdfjs, await normalized.arrayBuffer());
      } catch (error) {
        result.lastError = error;
      }
    }
    if (!result.document && !forcePdfiumPreview) {
      try {
        const legacyPdfJs = await loadLegacyPdfJs();
        result = await parsePdfDocument(legacyPdfJs, await normalized.arrayBuffer());
        usedLegacyRenderer = !!result.document;
      } catch (error) {
        result.lastError = error;
      }
    }
    const document = result.document;
    if (!document) {
      if (allowCompatPreview) {
        try {
          const pdfiumDocument = await openPdfiumDocument(new Uint8Array(await normalized.arrayBuffer()));
          releasePreviewDocument();
          state.previewPdfiumDocument = pdfiumDocument;
          state.previewMode = 'pdfium';
          state.previewPageCount = pdfiumDocument.pageCount;
          state.previewPage = 0;
          return normalized;
        } catch (error) {
          result.lastError = error;
        }
      }
      const parseError = new Error('生成的PDF无法完整读取，请重新生成');
      parseError.code = 'PDF_PARSE_FAILED';
      parseError.cause = result.lastError;
      throw parseError;
    }
    if (!document.numPages) { document.destroy(); throw new Error('生成的PDF没有页面'); }
    if (document.numPages > 500) { document.destroy(); throw new Error('生成的PDF超过500页，请拆分订单'); }
    releasePreviewDocument();
    state.previewDocument = document; state.previewMode = usedLegacyRenderer ? 'legacy' : 'pdfjs'; state.previewPageCount = document.numPages; state.previewPage = 0;
    return normalized;
  }

  async function openPdfiumDocument(bytes) {
    const pdfium = await loadPdfium();
    const filePtr = pdfium.pdfium.wasmExports.malloc(bytes.length);
    pdfium.pdfium.HEAPU8.set(bytes, filePtr);
    const docPtr = pdfium.FPDF_LoadMemDocument(filePtr, bytes.length, 0);
    if (!docPtr) {
      const errorCode = pdfium.FPDF_GetLastError();
      pdfium.pdfium.wasmExports.free(filePtr);
      throw new Error(`PDFium无法读取PDF（错误 ${errorCode}）`);
    }
    const pageCount = pdfium.FPDF_GetPageCount(docPtr);
    if (!pageCount) {
      pdfium.FPDF_CloseDocument(docPtr); pdfium.pdfium.wasmExports.free(filePtr);
      throw new Error('生成的PDF没有页面');
    }
    if (pageCount > 500) {
      pdfium.FPDF_CloseDocument(docPtr); pdfium.pdfium.wasmExports.free(filePtr);
      throw new Error('生成的PDF超过500页，请拆分订单');
    }
    let closed = false;
    return {
      pageCount,
      close() {
        if (closed) return;
        closed = true; pdfium.FPDF_CloseDocument(docPtr); pdfium.pdfium.wasmExports.free(filePtr);
      },
      async renderPage(pageIndex, availableWidth, canvas) {
        if (closed) throw new Error('PDFium预览已关闭');
        const pagePtr = pdfium.FPDF_LoadPage(docPtr, pageIndex);
        if (!pagePtr) throw new Error(`PDFium无法读取第 ${pageIndex + 1} 页`);
        try {
          const pageWidth = pdfium.FPDF_GetPageWidthF(pagePtr), pageHeight = pdfium.FPDF_GetPageHeightF(pagePtr);
          if (!(pageWidth > 0 && pageHeight > 0)) throw new Error('PDF页面尺寸无效');
          const cssScale = availableWidth / pageWidth, pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
          const width = Math.max(1, Math.round(pageWidth * cssScale * pixelRatio));
          const height = Math.max(1, Math.round(pageHeight * cssScale * pixelRatio));
          const bitmapPtr = pdfium.FPDFBitmap_Create(width, height, 0);
          if (!bitmapPtr) throw new Error('PDFium预览内存不足');
          try {
            pdfium.FPDFBitmap_FillRect(bitmapPtr, 0, 0, width, height, 0xFFFFFFFF);
            pdfium.FPDF_RenderPageBitmap(bitmapPtr, pagePtr, 0, 0, width, height, 0, 16);
            const bufferPtr = pdfium.FPDFBitmap_GetBuffer(bitmapPtr);
            if (!bufferPtr) throw new Error('PDFium无法读取页面图像');
            const rgba = new Uint8ClampedArray(pdfium.pdfium.HEAPU8.buffer, pdfium.pdfium.HEAPU8.byteOffset + bufferPtr, width * height * 4).slice();
            canvas.width = width; canvas.height = height;
            canvas.style.width = `${Math.ceil(width / pixelRatio)}px`; canvas.style.height = `${Math.ceil(height / pixelRatio)}px`;
            const context = canvas.getContext('2d', { alpha: false });
            if (!context) throw new Error('浏览器无法创建PDF预览画布');
            context.putImageData(new ImageData(rgba, width, height), 0, 0);
            return { width: width / pixelRatio, height: height / pixelRatio };
          } finally {
            pdfium.FPDFBitmap_Destroy(bitmapPtr);
          }
        } finally {
          pdfium.FPDF_ClosePage(pagePtr);
        }
      },
    };
  }

  async function parsePdfDocument(pdfjs, bytes) {
    let lastError;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      let loadingTask;
      try {
        loadingTask = pdfjs.getDocument({
          data: new Uint8Array(bytes.slice(0)),
          cMapUrl: new URL('vendor/pdfjs-cmaps/', location.href).href,
          cMapPacked: true,
          standardFontDataUrl: new URL('vendor/pdfjs-standard-fonts/', location.href).href,
          wasmUrl: new URL('vendor/pdfjs-wasm/', location.href).href,
          useSystemFonts: true,
        });
        return { document: await loadingTask.promise, lastError: null };
      } catch (error) {
        lastError = error;
        try { await loadingTask?.destroy?.(); } catch (_) { /* 保留原解析错误 */ }
      }
    }
    return { document: null, lastError };
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

  const SINGLE_LINE_DETAIL_FIELDS = new Set(['materialCode', 'sku', 'taxUnitPrice', 'taxAmount']);

  async function generateDocx() {
    const zip = await JSZip.loadAsync(state.templateBytes), xml = await zip.file('word/document.xml').async('string'), doc = parseXml(xml);
    for (const [targetId, mapping] of mappedEntries()) {
      if (mapping.mode === 'detail') continue;
      const target = locateWordTarget(doc, targetId); if (!target) throw new Error(`模板位置已变化：${targetId}`);
      const value = resolveMapping(mapping);
      putWordText(target, templateValue(mapping, value));
    }
    if (state.detailRow) {
      const match = state.detailRow.match(/^word:(\d+):(\d+)$/); if (!match) throw new Error('Word明细模板行无效');
      const table = direct(all(doc, 'body')[0], 'tbl')[Number(match[1])], tableRows = direct(table, 'tr'), row = tableRows[Number(match[2])];
      if (!row) throw new Error('找不到Word明细模板行');
      const reservedRows = tableRows.slice(Number(match[2]), Number(match[2]) + detailTemplateRowCount());
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      widenWordSkuColumn(table, row, rowMappings, detailRecords());
      for (const [recordIndex, record] of detailRecords().entries()) {
        const clone = row.cloneNode(true), cells = direct(clone, 'tc');
        for (const [targetId, mapping] of rowMappings) {
          const cellIndex = Number(targetId.match(/:c:(\d+)$/)?.[1]);
          if (cells[cellIndex]) {
            putWordText(cells[cellIndex], detailValue(mapping, record, recordIndex));
            if (SINGLE_LINE_DETAIL_FIELDS.has(mapping.businessKey)) fitWordDetailCell(cells[cellIndex], table, mapping.businessKey);
          }
        }
        row.parentNode.insertBefore(clone, row);
      }
      reservedRows.forEach(item => item.remove());
    }
    normalizeWordContractTerms(doc);
    compactWordContractLayout(doc);
    fitWordContractMainPage(doc);
    preventWordTextClipping(doc);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    return zip.generateAsync({ type: 'blob', mimeType: MIME.docx, compression: 'DEFLATE' });
  }

  function compactWordContractLayout(doc) {
    const body = all(doc, 'body')[0];
    let inIntegrityAppendix = false;
    for (const paragraph of all(body, 'p')) {
      let properties = direct(paragraph, 'pPr')[0];
      if (!properties) { properties = wordNode(doc, 'pPr'); paragraph.prepend(properties); }
      const appendixText = normalizeBusinessLabel(wordText(paragraph));
      if (paragraph.parentNode === body && appendixText === '通用条款') {
        let pageBreak = direct(properties, 'pageBreakBefore')[0];
        if (!pageBreak) { pageBreak = wordNode(doc, 'pageBreakBefore'); properties.append(pageBreak); }
        pageBreak.setAttributeNS(W, 'w:val', '1');
        let keep = direct(properties, 'keepNext')[0];
        if (!keep) { keep = wordNode(doc, 'keepNext'); properties.append(keep); }
        keep.setAttributeNS(W, 'w:val', '1');
      }
      const appendixHeading = /^(?:附件[一1])?供应方廉洁诚信承诺书$/.test(appendixText);
      if (paragraph.parentNode === body && appendixHeading) {
        let pageBreak = direct(properties, 'pageBreakBefore')[0];
        if (!pageBreak) { pageBreak = wordNode(doc, 'pageBreakBefore'); properties.append(pageBreak); }
        // The attachment label and its title belong on the same new page.
        pageBreak.setAttributeNS(W, 'w:val', inIntegrityAppendix ? '0' : '1');
        let keep = direct(properties, 'keepNext')[0];
        if (!keep) { keep = wordNode(doc, 'keepNext'); properties.append(keep); }
        keep.setAttributeNS(W, 'w:val', '1');
        inIntegrityAppendix = true;
      }
      // Preserve the appendix's original paragraph spacing and indentation.
      if (inIntegrityAppendix) continue;
      let snap = direct(properties, 'snapToGrid')[0];
      if (!snap) { snap = wordNode(doc, 'snapToGrid'); properties.append(snap); }
      snap.setAttributeNS(W, 'w:val', '0');
      let spacing = direct(properties, 'spacing')[0];
      if (!spacing) { spacing = wordNode(doc, 'spacing'); properties.append(spacing); }
      for (const key of ['beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing']) spacing.removeAttributeNS(W, key);
      spacing.setAttributeNS(W, 'w:before', '0'); spacing.setAttributeNS(W, 'w:after', '0');
      const sizes = all(paragraph, 'sz').map(node => Number(node.getAttributeNS(W, 'val')) / 2).filter(size => size > 0);
      const fontSize = sizes.length ? Math.max(...sizes) : 11;
      const text = wordText(paragraph).trim();
      const alignment = direct(properties, 'jc')[0]?.getAttributeNS(W, 'val');
      if (paragraph.parentNode === body && text && !['center', 'right', 'end'].includes(alignment)
          && !direct(properties, 'numPr').length && !/盖章|签署日期|授权代表|^年\s*月\s*日$/.test(text)) {
        // Reclaim horizontal space in prose without changing list numbering or signature layout.
        let indent = direct(properties, 'ind')[0];
        if (!indent) { indent = wordNode(doc, 'ind'); properties.append(indent); }
        for (const attribute of [...indent.attributes]) indent.removeAttributeNode(attribute);
        for (const key of ['left', 'right', 'start', 'end', 'leftChars', 'rightChars', 'hanging', 'hangingChars']) indent.setAttributeNS(W, `w:${key}`, '0');
        const numbered = /^(?:[一二三四五六七八九十百]+[、．.]|[（(]?\d+[）)、．.])/.test(text);
        indent.setAttributeNS(W, 'w:firstLine', numbered ? '0' : String(Math.round(fontSize * 20)));
        indent.setAttributeNS(W, 'w:firstLineChars', numbered ? '0' : '100');
      }
      const line = Math.max(paragraph.parentNode.localName === 'tc' ? 220 : 240, Math.ceil(fontSize * 1.3 * 20));
      spacing.setAttributeNS(W, 'w:line', String(line));
      spacing.setAttributeNS(W, 'w:lineRule', 'atLeast');
      // Move standalone page breaks onto the following heading; an overflowing empty
      // break paragraph otherwise creates an entirely blank page after a full table.
      if (!wordText(paragraph).trim() && !all(paragraph, 'sectPr').length && all(paragraph, 'br').some(node => node.getAttributeNS(W, 'type') === 'page')) {
        const next = paragraph.nextElementSibling;
        if (next?.localName === 'p' && wordText(next).trim()) {
          let nextProperties = direct(next, 'pPr')[0];
          if (!nextProperties) { nextProperties = wordNode(doc, 'pPr'); next.prepend(nextProperties); }
          let pageBreak = direct(nextProperties, 'pageBreakBefore')[0];
          if (!pageBreak) { pageBreak = wordNode(doc, 'pageBreakBefore'); nextProperties.append(pageBreak); }
          pageBreak.setAttributeNS(W, 'w:val', '1'); paragraph.remove(); continue;
        }
      }
      // Empty spacers may carry section breaks or signature space: retain those.
      if (!wordText(paragraph).trim() && !all(paragraph, 'sectPr').length && !all(paragraph, 'br').length && paragraph.parentNode === body) {
        const previous = paragraph.previousElementSibling;
        if (previous?.localName === 'p' && !wordText(previous).trim() && !all(previous, 'sectPr').length && !all(previous, 'br').length) paragraph.remove();
      }
    }
    const sections = all(body, 'sectPr');
    const pageShapes = sections.map(section => {
      const page = direct(section, 'pgSz')[0];
      return ['w', 'h', 'orient'].map(key => page?.getAttributeNS(W, key) || '').join(':');
    });
    if (new Set(pageShapes).size === 1) {
      for (const section of sections) {
        let type = direct(section, 'type')[0];
        if (!type) { type = wordNode(doc, 'type'); section.prepend(type); }
        type.setAttributeNS(W, 'w:val', 'continuous');
      }
    }
    for (const table of direct(body, 'tbl')) {
      // Keep stamping/signature rows spacious; remove minimum heights only from detail forms.
      const isDetail = direct(table, 'tr').some(row => direct(row, 'tc').some(cell => normalizeBusinessLabel(wordCellText(cell)) === 'SKU'));
      if (!isDetail) continue;
      for (const height of all(table, 'trHeight')) height.remove();
    }
  }

  function isSeparateContractHeading(text) {
    return /^(?:通用条款|(?:附件[一1])?供应方廉洁诚信承诺书)$/.test(normalizeBusinessLabel(text));
  }

  function fitWordContractMainPage(doc) {
    const body = all(doc, 'body')[0], blocks = [];
    for (const block of [...body.children]) {
      if (all(block, 'p').some(p => isSeparateContractHeading(wordText(p)))
          || (block.localName === 'p' && isSeparateContractHeading(wordText(block)))) break;
      if (block.localName !== 'sectPr') blocks.push(block);
    }
    if (!blocks.length) return;
    const ensure = (parent, name) => {
      let node = direct(parent, name)[0];
      if (!node) { node = wordNode(doc, name); parent.append(node); }
      return node;
    };
    const section = all(body, 'sectPr')[0];
    if (!section) return;
    const page = ensure(section, 'pgSz'), landscape = page.getAttributeNS(W, 'orient') === 'landscape';
    page.setAttributeNS(W, 'w:w', landscape ? '16838' : '11906');
    page.setAttributeNS(W, 'w:h', landscape ? '11906' : '16838');
    const margins = ensure(section, 'pgMar');
    for (const side of ['top', 'bottom', 'left', 'right']) {
      margins.setAttributeNS(W, `w:${side}`, String(Math.min(720, Number(margins.getAttributeNS(W, side) || 720))));
    }
    const availableWidth = (Number(page.getAttributeNS(W, 'w')) - Number(margins.getAttributeNS(W, 'left')) - Number(margins.getAttributeNS(W, 'right'))) / 20;
    const availableHeight = (Number(page.getAttributeNS(W, 'h')) - Number(margins.getAttributeNS(W, 'top')) - Number(margins.getAttributeNS(W, 'bottom'))) / 20 - 24;
    const canvas = document.createElement('canvas').getContext('2d');
    const paragraphs = block => block.localName === 'p' ? [block] : all(block, 'p');
    const originalSizes = new Map();
    for (const block of blocks) for (const p of paragraphs(block)) {
      for (const run of all(p, 'r')) originalSizes.set(run, Number(all(run, 'sz')[0]?.getAttributeNS(W, 'val') || 22) / 2);
    }
    const sizeFor = (p, cap) => Math.min(cap, Math.max(8, ...all(p, 'r').map(r => originalSizes.get(r) || 11)));
    const paragraphHeight = (p, width, cap) => {
      const size = sizeFor(p, cap);
      canvas.font = `${size}px "宋体", serif`;
      const lines = wordText(p).split(/\r?\n/).reduce((sum, text) => sum + Math.max(1, Math.ceil(canvas.measureText(text).width * 1.18 / Math.max(12, width - 12))), 0);
      const imageHeight = Math.max(0, ...[...p.getElementsByTagNameNS('*', 'extent')].map(n => Number(n.getAttribute('cy') || 0) / 12700));
      return Math.max(lines * size * 1.22, imageHeight);
    };
    const estimate = cap => blocks.reduce((sum, block) => {
      if (block.localName !== 'tbl') return sum + paragraphs(block).reduce((h, p) => h + paragraphHeight(p, availableWidth, cap), 0);
      const grid = all(direct(block, 'tblGrid')[0], 'gridCol').map(n => Number(n.getAttributeNS(W, 'w')) / 20);
      return sum + direct(block, 'tr').reduce((height, row) => {
        let column = 0;
        const heights = direct(row, 'tc').map(cell => {
          const span = Number(all(cell, 'gridSpan')[0]?.getAttributeNS(W, 'val') || 1);
          const width = grid.slice(column, column + span).reduce((a, b) => a + b, 0) || availableWidth / Math.max(1, direct(row, 'tc').length);
          column += span;
          return all(cell, 'p').reduce((h, p) => h + paragraphHeight(p, width, cap), 0) + 4;
        });
        const minimum = Math.max(0, ...all(direct(row, 'trPr')[0], 'trHeight').map(n => Number(n.getAttributeNS(W, 'val')) / 20));
        return height + Math.max(minimum, ...heights);
      }, 0);
    }, 0);
    let cap = 10;
    while (cap > 8 && estimate(cap) > availableHeight) cap -= 0.5;
    for (const block of blocks) for (const p of paragraphs(block)) {
      let pp = direct(p, 'pPr')[0];
      if (!pp) { pp = wordNode(doc, 'pPr'); p.prepend(pp); }
      const spacing = ensure(pp, 'spacing');
      spacing.setAttributeNS(W, 'w:line', String(Math.ceil(sizeFor(p, cap) * 1.22 * 20)));
      spacing.setAttributeNS(W, 'w:lineRule', 'atLeast');
      spacing.setAttributeNS(W, 'w:before', '0'); spacing.setAttributeNS(W, 'w:after', '0');
      // Keep all text in normal flow; never use fixed heights or clip overflowing content.
      for (const run of all(p, 'r')) {
        let rp = direct(run, 'rPr')[0];
        if (!rp) { rp = wordNode(doc, 'rPr'); run.prepend(rp); }
        const size = Math.min(originalSizes.get(run) || 11, cap);
        for (const name of ['sz', 'szCs']) ensure(rp, name).setAttributeNS(W, 'w:val', String(size * 2));
      }
    }
  }

  function preventWordTextClipping(doc) {
    const body = all(doc, 'body')[0];
    // An exact template row height clips wrapped headings; retain it only as a minimum.
    for (const height of all(body, 'trHeight')) height.setAttributeNS(W, 'w:hRule', 'atLeast');
    for (const paragraph of all(body, 'p')) {
      let properties = direct(paragraph, 'pPr')[0];
      if (!properties) { properties = wordNode(doc, 'pPr'); paragraph.prepend(properties); }
      let spacing = direct(properties, 'spacing')[0];
      if (!spacing) { spacing = wordNode(doc, 'spacing'); properties.append(spacing); }
      const rule = spacing.getAttributeNS(W, 'lineRule');
      if (rule === 'exact') spacing.setAttributeNS(W, 'w:lineRule', 'atLeast');
      else if (!rule) { spacing.setAttributeNS(W, 'w:lineRule', 'auto'); spacing.setAttributeNS(W, 'w:line', '240'); }
    }
  }

  async function fitExcelRowHeights(zip, context) {
    const doc = context.sheetDoc;
    const styles = zip.file('xl/styles.xml') ? parseXml(await zip.file('xl/styles.xml').async('string')) : null;
    const shared = zip.file('xl/sharedStrings.xml') ? parseXml(await zip.file('xl/sharedStrings.xml').async('string')) : null;
    const strings = all(shared, 'si', S).map(node => all(node, 't', S).map(item => item.textContent).join(''));
    const xfs = direct(all(styles, 'cellXfs', S)[0], 'xf', S), fonts = direct(all(styles, 'fonts', S)[0], 'font', S);
    const columns = all(doc, 'col', S), merges = all(doc, 'mergeCell', S).map(node => XLSX.utils.decode_range(node.getAttribute('ref')));
    const canvas = document.createElement('canvas').getContext('2d'), clonedStyles = new Map();
    const columnWidth = column => {
      const item = columns.find(node => column + 1 >= Number(node.getAttribute('min')) && column + 1 <= Number(node.getAttribute('max')));
      return Number(item?.getAttribute('width') || 8.43) * 7 + 5;
    };
    for (const row of direct(sheetData(doc), 'row', S)) {
      if (row.getAttribute('hidden') === '1') continue;
      for (const cell of direct(row, 'c', S)) {
        const type = cell.getAttribute('t'), raw = direct(cell, 'v', S)[0]?.textContent || '';
        const value = type === 's' ? strings[Number(raw)] : type === 'inlineStr' ? all(cell, 't', S).map(node => node.textContent).join('') : raw;
        if (!value) continue;
        const point = XLSX.utils.decode_cell(cellAddressOf(cell));
        const merge = merges.find(range => range.s.r === point.r && range.s.c === point.c);
        const styleIndex = Number(cell.getAttribute('s') || 0), xf = xfs[styleIndex] || xfs[0];
        const font = fonts[Number(xf?.getAttribute('fontId') || 0)], size = Number(direct(font, 'sz', S)[0]?.getAttribute('val') || 11);
        const alignment = direct(xf, 'alignment', S)[0], singleLine = alignment?.getAttribute('shrinkToFit') === '1';
        const name = direct(font, 'name', S)[0]?.getAttribute('val') || 'Arial';
        canvas.font = `${direct(font, 'b', S).length ? 'bold ' : ''}${size * 4 / 3}px "${name}"`;
        let width = 0;
        for (let c = point.c; c <= (merge?.e.c ?? point.c); c += 1) width += columnWidth(c);
        width = Math.max(4, width - 10 - Number(alignment?.getAttribute('indent') || 0) * size * 4);
        const lines = value.split(/\r?\n/).reduce((sum, text) => sum + (singleLine ? 1 : Math.max(1, Math.ceil(canvas.measureText(text).width * 1.15 / width))), 0);
        const needed = Math.ceil(lines * size * 1.5 + 6);
        const endRow = (merge?.e.r ?? point.r) + 1;
        let existing = 0;
        for (let r = point.r + 1; r <= endRow; r += 1) existing += Number(ensureSheetRow(doc, r).getAttribute('ht') || 15);
        const target = ensureSheetRow(doc, endRow), height = Number(target.getAttribute('ht') || 15) + Math.max(0, needed - existing);
        if (height > 409.5) throw new Error(`模板 ${cellAddressOf(cell)} 内容过多，无法完整显示；请加宽该列或拆分该单元格`);
        target.setAttribute('ht', String(height)); target.setAttribute('customHeight', '1');
        if (xf && !singleLine && ['s', 'str', 'inlineStr'].includes(type)) {
          if (!clonedStyles.has(styleIndex)) {
            const clone = xf.cloneNode(true); let align = direct(clone, 'alignment', S)[0];
            if (!align) { align = styles.createElementNS(S, 'alignment'); clone.append(align); }
            align.setAttribute('wrapText', '1'); clone.setAttribute('applyAlignment', '1');
            const list = all(styles, 'cellXfs', S)[0]; clonedStyles.set(styleIndex, direct(list, 'xf', S).length); list.append(clone); list.setAttribute('count', String(direct(list, 'xf', S).length));
          }
          cell.setAttribute('s', String(clonedStyles.get(styleIndex)));
        }
      }
    }
    if (styles) zip.file('xl/styles.xml', new XMLSerializer().serializeToString(styles));
  }

  function widenWordSkuColumn(table, prototype, mappings, records) {
    const skuMapping = mappings.find(([, mapping]) => mapping.businessKey === 'sku');
    if (!skuMapping) return;
    const cells = direct(prototype, 'tc'), indexOf = id => Number(id.match(/:c:(\d+)$/)?.[1]);
    const skuIndex = indexOf(skuMapping[0]), skuCell = cells[skuIndex];
    if (!skuCell) return;
    const grid = direct(direct(table, 'tblGrid')[0], 'gridCol');
    if (!grid.length) return;
    const spans = cell => Number(all(cell, 'gridSpan')[0]?.getAttributeNS(W, 'val') || 1);
    const offset = (rowCells, index) => rowCells.slice(0, index).reduce((sum, cell) => sum + spans(cell), 0);
    const widths = grid.map(node => Number(node.getAttributeNS(W, 'w') || 0));
    const start = offset(cells, skuIndex), count = spans(skuCell);
    const current = widths.slice(start, start + count).reduce((sum, width) => sum + width, 0);
    const context = document.createElement('canvas').getContext('2d');
    const rp = all(skuCell, 'rPr')[0], fonts = direct(rp, 'rFonts')[0];
    const font = fonts?.getAttributeNS(W, 'ascii') || 'Times New Roman';
    context.font = `${direct(rp, 'b').length ? 'bold ' : ''}8px "${font}"`;
    const tableMargins = all(direct(table, 'tblPr')[0], 'tblCellMar')[0];
    const margins = all(skuCell, 'tcMar')[0];
    const padding = ['left', 'right'].reduce((sum, side) => sum + Number((direct(margins, side)[0] || direct(tableMargins, side)[0])?.getAttributeNS(W, 'w') ?? 108), 0);
    const textWidth = Math.max(0, ...records.map((record, index) => context.measureText(detailValue(skuMapping[1], record, index)).width));
    const needed = Math.ceil(textWidth / 0.95 * 20 + padding + 60);
    let remaining = Math.max(0, needed - current);
    if (!remaining) return;
    // Borrow only from descriptive columns; preserve the total table width and merged-cell boundaries.
    for (const [key, minimum] of [['specification', 600], ['remark', 500], ['materialName', 1600]]) {
      const mapping = mappings.find(([, item]) => item.businessKey === key);
      if (!mapping) continue;
      const index = indexOf(mapping[0]), cell = cells[index];
      if (!cell) continue;
      const donorStart = offset(cells, index), donorCount = spans(cell);
      const total = widths.slice(donorStart, donorStart + donorCount).reduce((sum, width) => sum + width, 0);
      const take = Math.min(remaining, Math.max(0, total - minimum));
      if (!take) continue;
      let deducted = 0;
      for (let i = donorStart; i < donorStart + donorCount; i += 1) {
        const amount = i === donorStart + donorCount - 1 ? take - deducted : Math.floor(take * widths[i] / total);
        widths[i] -= amount; deducted += amount;
      }
      widths[start + count - 1] += take; remaining -= take;
      if (!remaining) break;
    }
    if (remaining) throw new Error('SKU列过窄，无法在保持清晰字号的同时单行显示，请加宽模板中的SKU列');
    grid.forEach((node, index) => node.setAttributeNS(W, 'w:w', String(widths[index])));
    for (const row of direct(table, 'tr')) {
      let column = Number(all(direct(row, 'trPr')[0], 'gridBefore')[0]?.getAttributeNS(W, 'val') || 0);
      for (const cell of direct(row, 'tc')) {
        const count = spans(cell), width = widths.slice(column, column + count).reduce((sum, width) => sum + width, 0);
        column += count;
        if (!width) continue;
        let properties = direct(cell, 'tcPr')[0];
        if (!properties) { properties = wordNode(cell.ownerDocument, 'tcPr'); cell.prepend(properties); }
        let node = direct(properties, 'tcW')[0];
        if (!node) { node = wordNode(cell.ownerDocument, 'tcW'); properties.prepend(node); }
        node.setAttributeNS(W, 'w:type', 'dxa'); node.setAttributeNS(W, 'w:w', String(width));
      }
    }
  }

  // Keep identifiers and money on one line without widening the template table.
  function fitWordDetailCell(cell, table, businessKey) {
    const doc = cell.ownerDocument;
    const ensure = (parent, name) => {
      let node = direct(parent, name)[0];
      if (!node) { node = wordNode(doc, name); if (name === 'pPr' || name === 'rPr') parent.prepend(node); else parent.append(node); }
      return node;
    };
    let properties = direct(cell, 'tcPr')[0];
    if (!properties) { properties = wordNode(doc, 'tcPr'); cell.prepend(properties); }
    ensure(properties, 'noWrap').setAttributeNS(W, 'w:val', '1');
    // Font size is fitted explicitly; native fit-text can compress it again in PDF conversion.
    direct(properties, 'tcFitText').forEach(node => node.remove());
    const widthNode = direct(properties, 'tcW')[0];
    let width = widthNode?.getAttributeNS(W, 'type') === 'dxa' ? Number(widthNode.getAttributeNS(W, 'w')) : 0;
    if (!width) {
      const siblings = direct(cell.parentNode, 'tc');
      const span = item => Number(all(item, 'gridSpan')[0]?.getAttributeNS(W, 'val') || 1);
      const start = siblings.slice(0, siblings.indexOf(cell)).reduce((sum, item) => sum + span(item), 0);
      width = direct(direct(table, 'tblGrid')[0], 'gridCol').slice(start, start + span(cell))
        .reduce((sum, item) => sum + Number(item.getAttributeNS(W, 'w') || 0), 0);
    }
    const tableMargins = all(direct(table, 'tblPr')[0], 'tblCellMar')[0];
    const cellMargins = direct(properties, 'tcMar')[0];
    const margin = side => Number((direct(cellMargins, side)[0] || direct(tableMargins, side)[0])?.getAttributeNS(W, 'w') ?? 108);
    const available = (width - margin('left') - margin('right')) / 20 - 2;
    const canvas = document.createElement('canvas'), context = canvas.getContext('2d');
    for (const paragraph of direct(cell, 'p')) {
      const paragraphProperties = ensure(paragraph, 'pPr');
      // Template first-line/hanging indents must not steal space from a short field.
      direct(paragraphProperties, 'ind').forEach(node => node.remove());
      ensure(paragraphProperties, 'wordWrap').setAttributeNS(W, 'w:val', '0');
      for (const run of direct(paragraph, 'r')) {
        const rp = ensure(run, 'rPr');
        for (const name of ['fitText', 'w', 'spacing']) direct(rp, name).forEach(node => node.remove());
        const fonts = direct(rp, 'rFonts')[0];
        const size = Math.max(businessKey === 'sku' ? 8 : 0, Number(direct(rp, 'sz')[0]?.getAttributeNS(W, 'val') || 22) / 2);
        const font = fonts?.getAttributeNS(W, 'ascii') || 'Times New Roman';
        context.font = `${direct(rp, 'b').length ? 'bold ' : ''}${size}px "${font}"`;
        const measured = context.measureText(all(run, 't').map(node => node.textContent).join('')).width;
        for (const name of ['sz', 'szCs']) ensure(rp, name).setAttributeNS(W, 'w:val', String(size * 2));
        if (available > 0 && measured > available) {
          const fitted = Math.max(1, Math.floor(size * available / measured * 0.95 * 2)) / 2;
          if (businessKey === 'sku' && fitted < 8) throw new Error('SKU列宽不足以清晰单行显示，请加宽模板中的SKU列后重新生成');
          for (const name of ['sz', 'szCs']) ensure(rp, name).setAttributeNS(W, 'w:val', String(fitted * 2));
        }
      }
    }
  }

  async function fitExcelDetailCells(zip, cells) {
    if (!cells.length || !zip.file('xl/styles.xml')) return;
    const doc = parseXml(await zip.file('xl/styles.xml').async('string'));
    const xfs = all(doc, 'cellXfs', S)[0];
    if (!xfs) return;
    const originals = direct(xfs, 'xf', S), styles = new Map();
    for (const cell of cells) {
      const index = Number(cell.getAttribute('s') || 0);
      if (!styles.has(index)) {
        const xf = (originals[index] || originals[0]).cloneNode(true);
        let alignment = direct(xf, 'alignment', S)[0];
        if (!alignment) { alignment = doc.createElementNS(S, 'alignment'); xf.append(alignment); }
        alignment.setAttribute('wrapText', '0'); alignment.setAttribute('shrinkToFit', '1');
        xf.setAttribute('applyAlignment', '1');
        styles.set(index, direct(xfs, 'xf', S).length); xfs.append(xf);
      }
      cell.setAttribute('s', String(styles.get(index)));
    }
    xfs.setAttribute('count', String(direct(xfs, 'xf', S).length));
    zip.file('xl/styles.xml', new XMLSerializer().serializeToString(doc));
  }

  async function generateXlsx() {
    const zip = await JSZip.loadAsync(state.templateBytes);
    const context = await locateXlsxSheet(zip, state.templateSheet);
    for (const [targetId, mapping] of mappedEntries()) {
      if (mapping.mode === 'detail') continue;
      const value = resolveMapping(mapping);
      writeSheetCell(context.sheetDoc, targetId.slice(2), templateValue(mapping, value), { forceNumber: NUMERIC_BUSINESS_FIELDS.has(mapping.businessKey) && !mapping.preserveLabel });
    }
    if (state.detailRow) {
      const rowNumber = Number(state.detailRow.split(':')[1]);
      const rowMappings = mappedEntries().filter(([id, mapping]) => mapping.mode === 'detail' && findTarget(id)?.rowKey === state.detailRow);
      await expandExcelDetailRowXml(zip, context, rowNumber, detailRecords(), rowMappings);
    }
    await normalizeExcelContractTerms(zip, context);
    await fitExcelRowHeights(zip, context);
    await fitExcelContractMainPage(zip, context);
    await forceWorkbookRecalculation(zip, context);
    zip.file(context.sheetPath, new XMLSerializer().serializeToString(context.sheetDoc));
    return zip.generateAsync({ type: 'blob', mimeType: MIME.xlsx, compression: 'DEFLATE' });
  }

  function compactText(value) { return C.text(value).replace(/\s+/g, ''); }

  function looksLikeStandardContractTerms(value) {
    const text = compactText(value);
    if (text.length < 160) return false;
    return CONTRACT_TERMS_MARKERS.filter(marker => text.includes(compactText(marker))).length >= 2;
  }

  function hasStandardContractTermsTemplate() {
    return (state.templateModel?.rows || []).some(row => {
      const labelIndex = row.cells.findIndex(cell => compactText(cell.value) === '合同条款');
      return labelIndex >= 0 && row.cells.slice(labelIndex + 1).some(cell => looksLikeStandardContractTerms(cell.value));
    });
  }

  function normalizeWordContractTerms(doc) {
    for (const table of all(all(doc, 'body')[0], 'tbl')) {
      for (const row of direct(table, 'tr')) {
        const cells = direct(row, 'tc'), labelIndex = cells.findIndex(cell => compactText(wordCellText(cell)) === '合同条款');
        if (labelIndex < 0) continue;
        const contentCell = cells.slice(labelIndex + 1).find(cell => looksLikeStandardContractTerms(wordCellText(cell)));
        if (!contentCell) continue;
        putWordText(contentCell, CONTRACT_TERMS_TEXT);
        const rowProperties = direct(row, 'trPr')[0];
        if (rowProperties) {
          direct(rowProperties, 'trHeight').forEach(item => item.remove());
          direct(rowProperties, 'cantSplit').forEach(item => item.remove());
        }
        let cellProperties = direct(contentCell, 'tcPr')[0];
        if (!cellProperties) { cellProperties = wordNode(doc, 'tcPr'); contentCell.insertBefore(cellProperties, contentCell.firstChild); }
        direct(cellProperties, 'noWrap').forEach(item => item.remove());
        let vertical = direct(cellProperties, 'vAlign')[0];
        if (!vertical) { vertical = wordNode(doc, 'vAlign'); cellProperties.append(vertical); }
        vertical.setAttributeNS(W, 'w:val', 'top');
        const paragraphProperties = direct(direct(contentCell, 'p')[0], 'pPr')[0];
        if (paragraphProperties) ['keepLines', 'keepNext', 'pageBreakBefore'].forEach(name => direct(paragraphProperties, name).forEach(item => item.remove()));
        return true;
      }
    }
    return false;
  }

  function findExcelContractTermsTarget() {
    for (const row of state.templateModel?.rows || []) {
      const label = row.cells.find(cell => compactText(cell.value) === '合同条款');
      if (!label) continue;
      const content = row.cells.find(cell => cell.cellIndex > label.cellIndex && looksLikeStandardContractTerms(cell.value));
      if (content) return content;
    }
    return null;
  }

  function shiftedTemplateAddress(address) {
    const point = XLSX.utils.decode_cell(address), detailRow = Number(state.detailRow?.split(':')[1] || 0), delta = detailRecords().length - detailTemplateRowCount();
    if (detailRow && delta !== 0 && point.r + 1 > detailRow) point.r += delta;
    return XLSX.utils.encode_cell(point);
  }

  function excelContractTermsLayout(address) {
    const sheet = state.templateBook?.Sheets[state.templateSheet], point = XLSX.utils.decode_cell(address);
    const merge = (sheet?.['!merges'] || []).find(range => point.r >= range.s.r && point.r <= range.e.r && point.c >= range.s.c && point.c <= range.e.c);
    const startColumn = merge?.s.c ?? point.c, endColumn = merge?.e.c ?? point.c;
    let width = 0;
    for (let column = startColumn; column <= endColumn; column += 1) width += Number(sheet?.['!cols']?.[column]?.wch || 8.43);
    const usableWidth = Math.max(18, width - 2), weightedLength = line => [...line].reduce((total, character) => total + (/[^\u0000-\u00ff]/.test(character) ? 2 : 1), 0);
    const lines = CONTRACT_TERMS.reduce((total, clause) => total + Math.max(1, Math.ceil(weightedLength(clause) / usableWidth)), 0);
    let fontSize = 10, height = Math.ceil(lines * fontSize * 1.38 + 8);
    if (height > 409.5) { fontSize = Math.max(8, Number((fontSize * 409.5 / height).toFixed(1))); height = Math.ceil(lines * fontSize * 1.38 + 8); }
    if (height > 409.5) throw new Error('合同条款区域过窄，无法完整显示1—9条；请扩大模板中“合同条款”右侧内容区域');
    return { fontSize, height: Math.min(409.5, Math.max(15, height)) };
  }

  async function normalizeExcelContractTerms(zip, context) {
    const target = findExcelContractTermsTarget();
    if (!target) return false;
    const address = shiftedTemplateAddress(target.address), point = XLSX.utils.decode_cell(address), row = ensureSheetRow(context.sheetDoc, point.r + 1), cell = ensureRowCell(row, address);
    const layout = excelContractTermsLayout(target.address);
    writeCellValue(cell, CONTRACT_TERMS_TEXT);
    row.setAttribute('ht', String(layout.height)); row.setAttribute('customHeight', '1');
    await ensureExcelContractTermsStyle(zip, cell, layout.fontSize);
    configureExcelContractPrint(context.sheetDoc, detailRecords().length <= 8);
    return true;
  }

  async function fitExcelContractMainPage(zip, context) {
    const doc = context.sheetDoc, worksheet = doc.documentElement;
    const shared = zip.file('xl/sharedStrings.xml') ? parseXml(await zip.file('xl/sharedStrings.xml').async('string')) : null;
    const strings = all(shared, 'si', S).map(n => all(n, 't', S).map(t => t.textContent).join(''));
    const rows = direct(sheetData(doc), 'row', S), boundaries = [];
    let appendixStarted = false;
    for (const row of rows) for (const cell of direct(row, 'c', S)) {
      const value = cell.getAttribute('t') === 's' ? strings[Number(direct(cell, 'v', S)[0]?.textContent)] : cell.getAttribute('t') === 'inlineStr' ? all(cell, 't', S).map(n => n.textContent).join('') : direct(cell, 'v', S)[0]?.textContent || '';
      if (!isSeparateContractHeading(value)) continue;
      const appendix = normalizeBusinessLabel(value) !== '通用条款';
      if (!appendix || !appendixStarted) boundaries.push(Number(row.getAttribute('r')) - 1);
      if (appendix) appendixStarted = true;
    }
    const firstBoundary = Math.min(...boundaries.filter(n => n > 0), Infinity);
    const height = rows.filter(row => Number(row.getAttribute('r')) <= firstBoundary && row.getAttribute('hidden') !== '1')
      .reduce((sum, row) => sum + Number(row.getAttribute('ht') || 15), 0);
    configureExcelContractPrint(doc, false);
    const setup = direct(worksheet, 'pageSetup', S)[0];
    const landscape = setup.getAttribute('orientation') === 'landscape';
    let margins = direct(worksheet, 'pageMargins', S)[0];
    if (!margins) { margins = doc.createElementNS(S, 'pageMargins'); worksheet.insertBefore(margins, setup); }
    for (const side of ['left', 'right', 'top', 'bottom']) margins.setAttribute(side, String(Math.min(0.4, Number(margins.getAttribute(side) || 0.4))));
    for (const side of ['header', 'footer']) if (!margins.hasAttribute(side)) margins.setAttribute(side, '0.2');
    const usableHeight = (landscape ? 595.28 : 841.89) - (Number(margins.getAttribute('top')) + Number(margins.getAttribute('bottom'))) * 72 - 12;
    const usableWidth = (landscape ? 841.89 : 595.28) - (Number(margins.getAttribute('left')) + Number(margins.getAttribute('right'))) * 72 - 12;
    if (!boundaries.length) {
      // Fit a modest main form to one page; very large orders retain readable pagination.
      setup.setAttribute('fitToHeight', height <= usableHeight / 0.65 ? '1' : '0');
      return;
    }
    let breaks = direct(worksheet, 'rowBreaks', S)[0];
    if (!breaks) { breaks = doc.createElementNS(S, 'rowBreaks'); worksheet.insertBefore(breaks, [...worksheet.children].find(n => ['colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'].includes(n.localName)) || null); }
    const ids = new Set([...direct(breaks, 'brk', S).map(n => Number(n.getAttribute('id'))), ...boundaries.filter(n => n > 0)]);
    breaks.replaceChildren();
    for (const id of [...ids].sort((a, b) => a - b)) {
      const node = doc.createElementNS(S, 'brk');
      for (const [key, value] of Object.entries({ id, min: 0, max: 16383, man: 1 })) node.setAttribute(key, String(value));
      breaks.append(node);
    }
    breaks.setAttribute('count', String(ids.size)); breaks.setAttribute('manualBreakCount', String(ids.size));
    // Excel ignores manual section breaks in fit-to-page mode; use explicit print scale.
    direct(direct(worksheet, 'sheetPr', S)[0], 'pageSetUpPr', S)[0].setAttribute('fitToPage', '0');
    const cells = rows.flatMap(row => direct(row, 'c', S));
    const lastColumn = Math.max(0, ...cells.map(cell => XLSX.utils.decode_cell(cellAddressOf(cell)).c));
    const columns = all(doc, 'col', S);
    const styleDoc = zip.file('xl/styles.xml') ? parseXml(await zip.file('xl/styles.xml').async('string')) : null;
    const defaultFont = all(styleDoc, 'font', S)[0];
    // Column units depend on the workbook's normal font, not a fixed 11-point font.
    const digitWidth = Math.max(7, Number(direct(defaultFont, 'sz', S)[0]?.getAttribute('val') || 11) * 0.75);
    let width = 0;
    for (let c = 1; c <= lastColumn + 1; c++) {
      const column = columns.find(n => c >= Number(n.getAttribute('min')) && c <= Number(n.getAttribute('max')));
      width += (Number(column?.getAttribute('width') || 8.43) * digitWidth + 5) * 0.75;
    }
    setup.removeAttribute('fitToWidth'); setup.removeAttribute('fitToHeight');
    setup.setAttribute('scale', String(Math.max(10, Math.floor(Math.min(1, usableWidth / width * 0.96, Math.max(0.65, usableHeight / Math.max(1, height))) * 100))));
  }

  function configureExcelContractPrint(doc, fitOnePage) {
    const worksheet = doc.documentElement;
    let sheetProperties = direct(worksheet, 'sheetPr', S)[0];
    if (!sheetProperties) { sheetProperties = doc.createElementNS(S, 'sheetPr'); worksheet.insertBefore(sheetProperties, worksheet.firstChild); }
    let pageSetupProperties = direct(sheetProperties, 'pageSetUpPr', S)[0];
    if (!pageSetupProperties) { pageSetupProperties = doc.createElementNS(S, 'pageSetUpPr'); sheetProperties.append(pageSetupProperties); }
    pageSetupProperties.setAttribute('fitToPage', '1'); pageSetupProperties.setAttribute('autoPageBreaks', '0');
    let pageSetup = direct(worksheet, 'pageSetup', S)[0];
    if (!pageSetup) {
      pageSetup = doc.createElementNS(S, 'pageSetup');
      const margins = direct(worksheet, 'pageMargins', S)[0];
      worksheet.insertBefore(pageSetup, margins?.nextSibling || [...worksheet.children].find(n => ['headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'].includes(n.localName)) || null);
    }
    pageSetup.setAttribute('paperSize', '9'); pageSetup.setAttribute('fitToWidth', '1'); pageSetup.setAttribute('fitToHeight', fitOnePage ? '1' : '0'); pageSetup.removeAttribute('scale');
  }

  async function ensureExcelContractTermsStyle(zip, cell, fontSize) {
    const stylesFile = zip.file('xl/styles.xml');
    if (!stylesFile) return;
    const doc = parseXml(await stylesFile.async('string')), fonts = all(doc, 'fonts', S)[0], cellXfs = all(doc, 'cellXfs', S)[0];
    if (!fonts || !cellXfs) return;
    const xfs = direct(cellXfs, 'xf', S), sourceXf = xfs[Number(cell.getAttribute('s') || 0)] || xfs[0];
    if (!sourceXf) return;
    const fontId = Number(sourceXf.getAttribute('fontId') || 0), sourceFont = direct(fonts, 'font', S)[fontId] || direct(fonts, 'font', S)[0];
    const font = sourceFont?.cloneNode(true) || doc.createElementNS(S, 'font');
    let size = direct(font, 'sz', S)[0];
    if (!size) { size = doc.createElementNS(S, 'sz'); font.append(size); }
    const sourceSize = Number(size.getAttribute('val') || fontSize); size.setAttribute('val', String(Math.min(sourceSize || fontSize, fontSize)));
    fonts.append(font); fonts.setAttribute('count', String(direct(fonts, 'font', S).length));
    const xf = sourceXf.cloneNode(true); xf.setAttribute('fontId', String(direct(fonts, 'font', S).length - 1)); xf.setAttribute('applyFont', '1'); xf.setAttribute('applyAlignment', '1');
    let alignment = direct(xf, 'alignment', S)[0];
    if (!alignment) { alignment = doc.createElementNS(S, 'alignment'); xf.append(alignment); }
    alignment.setAttribute('wrapText', '1'); alignment.setAttribute('vertical', 'top'); alignment.removeAttribute('shrinkToFit');
    cellXfs.append(xf); cellXfs.setAttribute('count', String(direct(cellXfs, 'xf', S).length)); cell.setAttribute('s', String(direct(cellXfs, 'xf', S).length - 1));
    zip.file('xl/styles.xml', new XMLSerializer().serializeToString(doc));
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
    const reservedCount = detailTemplateRowCount();
    const reservedRows = rows.filter(row => rowNumberOf(row) >= rowNumber && rowNumberOf(row) < rowNumber + reservedCount);
    const delta = records.length - reservedCount;
    if (delta !== 0) {
      const mergeRefs = all(context.sheetDoc, 'mergeCell', S).map(item => item.getAttribute('ref')).filter(Boolean);
      if (mergeRefs.some(ref => Array.from({ length: reservedCount }, (_, offset) => rowNumber + offset).some(targetRow => rangeTouchesRow(ref, targetRow)))) throw new Error('Excel明细模板行包含合并单元格，浏览器版不能安全扩展');
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
    const singleLineCells = [];
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
        const cell = ensureRowCell(clone, address);
        writeCellValue(cell, detailValue(mapping, record, offset), { forceNumber: NUMERIC_BUSINESS_FIELDS.has(mapping.businessKey) });
        if (SINGLE_LINE_DETAIL_FIELDS.has(mapping.businessKey)) singleLineCells.push(cell);
      }
      data.insertBefore(clone, prototype);
    });
    reservedRows.forEach(row => row.remove());
    await fitExcelDetailCells(zip, singleLineCells);
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

  function calculatedAmount(record) {
    try { return C.formatAmountCents(C.lineAmountCents(record[state.fieldSelections.quantity], record[state.fieldSelections.taxUnitPrice])); }
    catch (e) { throw new Error(`订单第 ${record._row} 行：${e.message}`); }
  }
  function contractTotals() {
    return C.sumAmountField(detailRecords().map(record => ({ _row: record._row, calculated: calculatedAmount(record) })), 'calculated');
  }
  function resolveMapping(mapping) {
    if (mapping.businessKey === 'taxTotalLower') return contractTotals().lower;
    if (mapping.businessKey === 'taxTotalUpper') return contractTotals().upper;
    if (mapping.businessKey === 'signDate') return formatChineseDate(state.manualValues?.signDate);
    return C.resolveField(state.order.rows, mapping.field, mapping.strategy || 'first', mapping.manual || '');
  }
  function formatChineseDate(value) {
    const match = C.text(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return match ? `${match[1]}年${Number(match[2])}月${Number(match[3])}日` : C.text(value);
  }
  function detailValue(mapping, record, index) { if (mapping.field === '@line-amount') return calculatedAmount(record); return mapping.field === '@sequence' ? String(index + 1) : C.text(record[mapping.field]); }
  function templateValue(mapping, value) {
    if (mapping.inlinePrefix || mapping.inlineSuffix) {
      let output = C.text(value);
      if (mapping.businessKey === 'taxTotalUpper' && /[圆元]整$/.test(mapping.inlineSuffix)) {
        if (/[圆元]整$/.test(output)) output = output.replace(/[圆元]整$/, '');
        else return `${mapping.inlinePrefix}${output}`;
      }
      return `${mapping.inlinePrefix}${output}${mapping.inlineSuffix}`;
    }
    return mapping.preserveLabel ? labeledValue(mapping.labelText, value) : value;
  }
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
    for (const definition of ACTIVE_FIELDS) {
      const value = payload.fieldSelections?.[definition.key];
      if (definition.automatic) selections[definition.key] = definition.automatic;
      else if (fields.has(value)) selections[definition.key] = value;
    }
    state.fieldSelections = selections; state.fieldStrategies = { ...(payload.fieldStrategies || {}) }; state.mappingSignature = '';
  }

  function mappingKey() { return state.fingerprint ? `gylsjqx-contract-mapping:${state.fingerprint}:${state.templateSheet || 'docx'}` : ''; }
  function saveMapping() { try { const key = mappingKey(); if (key) localStorage.setItem(key, JSON.stringify(mappingPayload())); } catch (_) {} }
  function restoreMapping() {
    try { const key = mappingKey(), raw = key && localStorage.getItem(key); if (raw) applyMappingPayload(JSON.parse(raw)); }
    catch (_) { state.mappings = {}; state.detailRow = ''; state.fieldSelections = {}; state.fieldStrategies = {}; state.manualValues = {}; state.mappingSignature = ''; }
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
