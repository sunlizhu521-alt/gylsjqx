const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const C = require('../funds-core.js');

// Only synthetic suppliers and files are used, including when QA_URL is the live site.
const FILE_IDS = { ledger: 'ledgerFile', applications: 'applicationsFile', negotiated: 'negotiatedFile', oa: 'oaFile' };
const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MARKER = 'N("GYLSJQX_FUNDS_APPLIED_V1")=0';
const SHEET = '10月谈判付款批次';
const ledgerHeaders = ['登记月份', '供应商名称', '付款主体', '对公付款0A流程号', '开票时间', '开票金额', '摘要', '申请付款金额', '应付款时间', '单据号【财务应付单】', 'OA合同编号', '合同总金额', '已付金额', '是否结清', '剩余金额', '是否谈判付款', '是否提交付款申请单'];
const negotiatedHeaders = ['供应商名称', '付款主体', '对公付款0A流程号', '申请付款金额', '已付金额', '剩余付款金额', '是否结清', '应付款时间', '单据号【财务应付单】', '备注', '谈判明细', '分批总金额', '核对', '与张培培沟通记录', '申请人', '当前节点', '未操作者', '流程状态', '流程编号', '是否提交付款申请单', 46321, 46328];
const records = [
  { supplier: '唯一分批合成供应商', id: 'MDST-DGFK20261027001', amount: 100, paid: 10, type: '谈判付款', plan: [40, 60] },
  { supplier: '同额分批合成供应商', id: 'MDST-DGFK20261027002', amount: 80, paid: 0, type: '谈判付款', plan: [40, 40] },
  { supplier: '正常已申请合成供应商', id: 'MDST-DGFK20261027003', amount: 123.45, paid: 23.45, type: '正常付款' },
  { supplier: '正常未申请合成供应商', id: 'MDST-DGFK20261027004', amount: 200, paid: 0, type: '正常付款' }
];
const write = book => Buffer.from(X.write(book, { type: 'buffer', bookType: 'xlsx' }));
const sheetBook = (name, rows) => { const b = X.utils.book_new(); X.utils.book_append_sheet(b, X.utils.aoa_to_sheet(rows), name); return b; };

function fixture({ removeFirstApplication = false } = {}) {
  const ledger = X.utils.book_new();
  for (const person of C.PEOPLE) {
    const rows = [ledgerHeaders];
    if (person === '李奇') for (const [i, r] of records.entries()) rows.push([46300, r.supplier, '迈德斯特', r.id, 46300, r.amount, '合成货款', r.amount, 46328, 'AP2026102700' + i, 'HT-SYN-' + i, r.amount, r.paid, '否', r.amount - r.paid, r.type === '谈判付款' ? '是' : '否', '未申请']);
    const sheet = X.utils.aoa_to_sheet(rows);
    for (let row = 2; row <= rows.length; row++) for (const col of ['A', 'E', 'I']) sheet[col + row].z = 'yyyy"年"m"月"d"日"';
    X.utils.book_append_sheet(ledger, sheet, '对账明细-' + person);
  }
  const negotiated = sheetBook(SHEET, [negotiatedHeaders, ...records.slice(0, 2).map((r, i) => [r.supplier, '迈德斯特', r.id, r.amount, r.paid, r.amount - r.paid, '否', 46328, 'AP2026102700' + i, '保留人工备注', '1、10月26日-' + r.plan[0] + '\n2、11月2日-' + r.plan[1], r.amount, '一致', '保留沟通记录', '李奇', '财务审核', '合成审核人', '未归档', r.id, '未提交', ...r.plan])]);
  const plan = negotiated.Sheets[SHEET];
  for (const col of ['U', 'V']) plan[col + '1'].z = 'yyyy"年"m"月"d"日"';
  for (const row of [2, 3]) {
    plan['F' + row] = { t: 'n', v: records[row - 2].amount - records[row - 2].paid, f: 'D' + row + '-E' + row, z: '#,##0.00' };
    plan['L' + row] = { t: 'n', v: records[row - 2].amount, f: 'SUM(U' + row + ':V' + row + ')', z: '#,##0.00' };
    plan['H' + row].z = 'yyyy"年"m"月"d"日"';
  }
  const applications = [['申请日期', '供应商全称', '付款主体', '流程编码', '本次申请付款金额', '付款分类']];
  records.slice(0, 3).forEach((r, i) => { if (!removeFirstApplication || i !== 0) applications.push(['2026-10-27', r.supplier, '迈德斯特', r.id, r.plan ? 40 : r.amount, r.type]); });
  const oa = [['流程编号', '当前节点', '未操作者', '流程状态', '申请人', '收款单位', '付款金额'], ...records.map(r => [r.id, '财务审核', '合成审核人', '未归档', '李奇', r.supplier, r.amount])];
  return { ledger: write(ledger), negotiated: write(negotiated), applications: write(sheetBook('付款申请汇总', applications)), oa: write(sheetBook('Sheet0', oa)) };
}

(async () => {
  const base = process.env.QA_URL || 'http://127.0.0.1:8940/';
  const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } }), errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.clock.setFixedTime(new Date('2026-10-27T01:00:00Z'));
    await page.goto(base, { waitUntil: 'networkidle' });
    const homepage = page.url();
    await page.locator('.tool-tab[data-tool="funds"]').click();
    assert.equal(page.url(), homepage);
    const frame = page.frameLocator('iframe[data-tool="funds"]');
    await frame.locator('#ledgerFile').waitFor();
    const data = fixture();
    async function upload(key, bytes = data[key]) {
      await frame.locator('#' + FILE_IDS[key] + ':not([disabled])').waitFor();
      await frame.locator('#' + FILE_IDS[key]).setInputFiles({ name: key + '.xlsx', mimeType: MIME, buffer: bytes });
    }
    for (const key of ['ledger', 'applications', 'oa', 'negotiated']) await upload(key);
    await frame.locator('#analyze:not([disabled])').waitFor();

    async function snapshot() {
      return frame.locator('body').evaluate(() => Object.fromEntries(['negotiatedPreview', 'normalPreview', 'reviewPreview'].map(id => [id, {
        headers: [...document.querySelectorAll('#' + id + ' th')].map(cell => ({ text: cell.textContent, applied: !!cell.closest('.funds-applied') })),
        rows: [...document.querySelectorAll('#' + id + ' tbody tr')].map(row => [...row.cells].map(cell => ({ text: cell.textContent, applied: !!cell.closest('.funds-applied'), hidden: cell.hidden, color: getComputedStyle(cell.closest('.funds-applied') || cell).backgroundColor })))
      }])));
    }
    function checkPreview(actual, highlighted) {
      for (const preview of Object.values(actual)) assert(preview.headers.every(cell => !cell.applied), '表头不标红');
      const n = actual.negotiatedPreview, normal = actual.normalPreview;
      const firstDate = n.headers.findIndex(h => h.text === '2026年10月26日'), secondDate = n.headers.findIndex(h => h.text === '2026年11月2日');
      assert(firstDate >= 0 && secondDate >= 0, '批次日期表头完整');
      assert.equal(n.rows.length, 2); assert.equal(normal.rows.length, 2);
      assert.equal(n.rows[0][firstDate].applied, highlighted);
      if (highlighted) assert.equal(n.rows[0][firstDate].color, 'rgb(255, 0, 0)');
      assert.equal(n.rows[0][secondDate].applied, false, '未申请60不标红');
      assert.equal(n.rows[1][firstDate].applied, true, '两批相同40按日期先后标记最早批次');
      assert.equal(n.rows[1][firstDate].color, 'rgb(255, 0, 0)');
      assert.equal(n.rows[1][secondDate].applied, false, '只有一笔申请，后续同额批次不能重复标红');
      assert.equal(n.rows[1].filter(cell => cell.applied).length, 1);
      assert.equal(n.rows[0].filter(cell => cell.applied).length, highlighted ? 1 : 0);
      assert(normal.rows[0].every(cell => cell.applied && cell.color === 'rgb(255, 0, 0)'), '正常付款已申请整行红色背景');
      assert(normal.rows[0].some(cell => cell.hidden && cell.applied), '隐藏列也属于整行标红范围');
      assert(normal.rows[1].every(cell => !cell.applied), '正常未申请不标红');
      assert(actual.reviewPreview.rows.every(row => row.every(cell => !cell.applied)), '核对表不套用申请标红');
      assert.equal(actual.reviewPreview.rows.length, 0, '同额批次已按日期确定，不生成同额核对');
    }
    async function analyze(highlighted) {
      // Pause export once: inspect the initial preview before the production export/re-read refresh.
      await frame.locator('body').evaluate(() => {
        const original = window.FundsExport;
        window.FundsExport = async (...args) => {
          window.FundsExport = original;
          window.__highlightResult = args[1].highlights;
          await new Promise(resolve => { window.__releaseHighlightExport = resolve; });
          return original(...args);
        };
      });
      await frame.locator('#analyze:not([disabled])').click();
      await frame.locator('body').evaluate(() => new Promise((resolve, reject) => {
        let attempts = 0;
        const check = () => window.__releaseHighlightExport ? resolve() : attempts++ > 1500 ? reject(Error(document.getElementById('status').textContent)) : setTimeout(check, 10); check();
      }));
      assert.equal(await frame.locator('#pendingCount').innerText(), '0', '同额批次按日期分配，不产生待核对事项');
      assert.deepEqual(await frame.locator('body').evaluate(() => window.__highlightResult), { negotiatedCells: highlighted ? ['U2', 'U3'] : ['U3'], normalRows: [2] });
      const initial = await snapshot(); checkPreview(initial, highlighted);
      await frame.locator('body').evaluate(() => { const release = window.__releaseHighlightExport; delete window.__releaseHighlightExport; release(); });
      await frame.locator('#download:not([disabled])').waitFor();
      const final = await snapshot(); checkPreview(final, highlighted);
      assert.deepEqual(final, initial, '首次预览和导出重读后的文字及标红一致');
      return final;
    }
    async function download(version) {
      const event = page.waitForEvent('download'); await frame.locator('#download').click();
      const result = await event;
      assert.equal(result.suggestedFilename(), '资金整理-2026-10-27-V' + version + '.xlsx');
      return fs.readFile(await result.path());
    }
    async function checkExport(bytes, highlighted) {
      const book = X.read(bytes, { type: 'buffer', cellNF: true, cellStyles: true });
      assert.deepEqual(book.SheetNames, ['谈判付款', '正常付款', '付款申请核对']);
      const n = book.Sheets.谈判付款, normal = book.Sheets.正常付款;
      for (const row of [2, 3]) {
        const r = records[row - 2];
        assert.deepEqual([n['U' + row].v, n['V' + row].v], r.plan, '原分批金额不变');
        assert.deepEqual([n['D' + row].v, n['E' + row].v, n['F' + row].v], [r.amount, r.paid, r.amount - r.paid]);
        assert.equal(n['F' + row].f, 'D' + row + '-E' + row); assert.equal(n['L' + row].f, 'SUM(U' + row + ':V' + row + ')');
        assert.equal(n['L' + row].v, r.amount); assert.equal(n['G' + row].v, '否');
      }
      assert.deepEqual([normal.I2.v, normal.N2.v, normal.P2.v, normal.I3.v, normal.N3.v, normal.P3.v], [123.45, 23.45, 100, 200, 0, 200]);
      assert.equal(normal.P2.f, 'I2-N2'); assert.equal(normal.P3.f, 'I3-N3');
      for (const header of C.NORMAL_HIDDEN_HEADERS) assert.equal(normal['!cols'][C.NORMAL_HEADERS.indexOf(header)].hidden, true);
      const metadata = await frame.locator('body').evaluate(async (_, { bytes, marker }) => {
        const zip = await JSZip.loadAsync(Uint8Array.from(bytes)), parse = text => new DOMParser().parseFromString(text, 'application/xml'), all = (doc, name) => [...doc.getElementsByTagNameNS('*', name)];
        const workbook = parse(await zip.file('xl/workbook.xml').async('string')), rels = parse(await zip.file('xl/_rels/workbook.xml.rels').async('string')), styles = parse(await zip.file('xl/styles.xml').async('string')), dxfs = all(styles, 'dxfs')[0];
        const result = {};
        for (const entry of all(workbook, 'sheet')) {
          const id = entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'), rel = all(rels, 'Relationship').find(e => e.getAttribute('Id') === id), path = new URL(rel.getAttribute('Target'), 'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1), doc = parse(await zip.file(path).async('string'));
          const cells = new Set(), rules = [];
          for (const cf of all(doc, 'conditionalFormatting')) for (const rule of all(cf, 'cfRule')) {
            if (!all(rule, 'formula').some(f => f.textContent === marker)) continue;
            const dxf = dxfs?.children[Number(rule.getAttribute('dxfId'))];
            rules.push({ type: rule.getAttribute('type'), color: dxf && all(dxf, 'fgColor')[0]?.getAttribute('rgb'), background: dxf && all(dxf, 'bgColor')[0]?.getAttribute('rgb'), backgroundIndex: dxf && all(dxf, 'bgColor')[0]?.getAttribute('indexed') });
            for (const ref of cf.getAttribute('sqref').split(/\s+/)) { const range = XLSX.utils.decode_range(ref); for (let r = range.s.r; r <= range.e.r; r++) for (let c = range.s.c; c <= range.e.c; c++) cells.add(XLSX.utils.encode_cell({ r, c })); }
          }
          result[entry.getAttribute('name')] = { cells: [...cells].sort(), rules };
        }
        return result;
      }, { bytes: [...bytes], marker: MARKER });
      assert.deepEqual(metadata.谈判付款.cells, highlighted ? ['U2', 'U3'] : ['U3']);
      assert.deepEqual(metadata.正常付款.cells, C.NORMAL_HEADERS.map((_, c) => X.utils.encode_cell({ r: 1, c })).sort());
      assert.deepEqual(metadata.付款申请核对.cells, []);
      for (const sheet of Object.values(metadata)) for (const rule of sheet.rules) { assert.equal(rule.type, 'expression'); assert.match(rule.color, /^(?:FF)?FF0000$/i); assert.equal(rule.background, 'FFFF0000', 'WPS conditional-format background must be explicitly red'); assert.equal(rule.backgroundIndex, null); }
      assert.equal(X.utils.sheet_to_json(book.Sheets.付款申请核对).length, 0, '同额批次按日期确定后不再产生核对记录');
    }

    const initial = await analyze(true);
    await page.locator('.tool-tab[data-tool="payment"]').click(); await page.locator('.tool-tab[data-tool="funds"]').click();
    assert.equal(page.url(), homepage); assert.deepEqual(await snapshot(), initial, '首页切换保留本次标红');
    const first = await download(1); await checkExport(first, true);
    if (process.env.QA_OUT) {
      await fs.mkdir(process.env.QA_OUT, { recursive: true });
      await frame.locator('#negotiatedPreview td.funds-applied').first().scrollIntoViewIfNeeded();
      await frame.locator('#negotiatedPreview').screenshot({ path: process.env.QA_OUT + '/negotiated.png' });
      await frame.locator('#normalPreview td.funds-applied').first().scrollIntoViewIfNeeded();
      await frame.locator('#normalPreview').screenshot({ path: process.env.QA_OUT + '/normal.png' });
      await fs.writeFile(process.env.QA_OUT + '/synthetic-highlight.xlsx', first);
    }

    // Feed the previous export back as the latest negotiated table, then remove one application.
    // Remove the withdrawn U2 highlight while retaining the earliest equal-amount batch at U3.
    await upload('negotiated', first);
    await upload('applications', fixture({ removeFirstApplication: true }).applications);
    await analyze(false);
    const second = await download(2); await checkExport(second, false);
    assert.deepEqual(errors, []);
    console.log('PASS: homepage upload, first/re-read preview, earliest equal-amount installment and full normal-row red fill, no unnecessary review, original formulas, hidden columns, fresh-input CF replacement');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
