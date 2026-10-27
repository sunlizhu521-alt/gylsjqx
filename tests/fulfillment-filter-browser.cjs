const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../fulfillment-core.js');

// Synthetic data only: the same checks can run against localhost or GitHub Pages.
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const mainNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const digest = value => crypto.createHash('sha256').update(value).digest('hex');
const sheetPath = name => 'xl/worksheets/sheet' + (C.SHEETS.indexOf(name) + 1) + '.xml';

async function fixture({ formulas = true, broken = false } = {}) {
  const book = X.utils.book_new();
  for (const [person, name] of C.SHEETS.entries()) {
    const row = index => C.COLUMNS.map(field => C.QUANTITY_FIELDS.includes(field) ? 0 : ({
      '采购下单人': name, '事业部': '国内事业部', '物料编码': 'FILTER-' + person + '-' + index,
      '供应商简称': '筛选测试供应商', '物料名称': '合成物料'
    })[field] ?? '');
    const sheet = X.utils.aoa_to_sheet([C.COLUMNS, row(0), row(1), row(2)]);
    if (name === '李奇') {
      sheet.L2.v = 1245; sheet.L3.v = 1050;
      sheet.K2 = formulas ? { t: 'n', v: 0, f: broken ? '#REF!+1' : 'SUBTOTAL(9,L2:O2)' } : { t: 'n', v: 1245 };
      sheet.K3 = formulas ? { t: 'n', v: 0, f: 'SUBTOTAL(109,L3:O3)' } : { t: 'n', v: 1050 };
      X.utils.sheet_add_aoa(sheet, [row(3)], { origin: 'A6' });
      sheet['!rows'] = [null, { hidden: true }, { hidden: true }, null, null, { hidden: true }];
      sheet['!autofilter'] = { ref: 'A1:X4' };
    }
    if (name === '操思敏') {
      sheet.L2.v = 7;
      sheet.K2 = formulas ? { t: 'n', v: 0, f: 'SUBTOTAL(109,L2:O2)' } : { t: 'n', v: 7 };
      sheet['!rows'] = [null, { hidden: true }];
    }
    if (name === '徐英') {
      sheet.L2.v = 3; sheet.L3.v = 5;
      sheet.K2 = formulas ? { t: 'n', v: 0, f: 'SUBTOTAL(9,L2:L2)' } : { t: 'n', v: 3 };
      sheet.K3 = formulas ? { t: 'n', v: 0, f: 'SUBTOTAL(109,L2:L3)' } : { t: 'n', v: 5 };
      sheet['!rows'] = [null, { hidden: true }];
      sheet['!autofilter'] = { ref: 'A1:X4' };
    }
    X.utils.book_append_sheet(book, sheet, name);
  }
  const zip = await Zip.loadAsync(X.write(book, { type: 'buffer', bookType: 'xlsx' }));
  const li = sheetPath('李奇');
  zip.file(li, (await zip.file(li).async('string'))
    .replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr filterMode="1"/>')
    .replace(/<autoFilter\b[^>]*\/>/, '<autoFilter ref="A1:X4"><filterColumn colId="2"><customFilters><customFilter operator="equal" val="/"/></customFilters></filterColumn></autoFilter>'));
  const xu = sheetPath('徐英');
  zip.file(xu, (await zip.file(xu).async('string')).replace(/<autoFilter\b[^>]*\/>/, '<autoFilter ref="A1:X4"><filterColumn colId="0" hiddenButton="1" showButton="0"/></autoFilter>'));
  const cao = sheetPath('操思敏');
  zip.file(cao, (await zip.file(cao).async('string')).replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rIdFilterTable"/></tableParts></worksheet>'));
  zip.file(cao.replace('/sheet', '/_rels/sheet') + '.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFilterTable" Type="' + relNS + '/table" Target="../tables/table1.xml"/></Relationships>');
  zip.file('xl/tables/table1.xml', '<table xmlns="' + mainNS + '" id="1" name="FilterFixture" displayName="FilterFixture" ref="A1:X4" totalsRowShown="0"><autoFilter ref="A1:X4"><filterColumn colId="0"><filters><filter val="徐英"/></filters></filterColumn></autoFilter><tableColumns count="24">' + C.COLUMNS.map((name, i) => '<tableColumn id="' + (i + 1) + '" name="' + name + '"/>').join('') + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>');
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

function rowTag(xml, row) { return xml.match(new RegExp('<row\\b[^>]*\\br="' + row + '"[^>]*>'))?.[0] || ''; }
async function assertCleared(bytes, { formulas = true } = {}) {
  const zip = await Zip.loadAsync(bytes);
  const [li, cao, xu, table] = await Promise.all([sheetPath('李奇'), sheetPath('操思敏'), sheetPath('徐英'), 'xl/tables/table1.xml'].map(path => zip.file(path).async('string')));
  for (const xml of [li, cao, table]) {
    assert.doesNotMatch(xml, /<(?:\w+:)?filterColumn\b/);
    assert.doesNotMatch(xml, /filterMode="(?:1|true)"/);
  }
  assert.match(li, /<autoFilter\b[^>]*ref="A1:X4"/);
  assert.match(table, /<autoFilter\b[^>]*ref="A1:X4"/);
  assert.match(table, /TableStyleMedium2/);
  for (const row of [2, 3]) assert.doesNotMatch(rowTag(li, row), /hidden="(?:1|true)"/, '筛选范围内的隐藏行恢复显示');
  assert.doesNotMatch(rowTag(cao, 2), /hidden="(?:1|true)"/, 'Table 筛选隐藏行恢复显示');
  assert.match(rowTag(li, 6), /hidden="1"/, '范围外手工隐藏行不改动');
  assert.match(rowTag(xu, 2), /hidden="1"/, '没有活动筛选条件的手工隐藏行不改动');
  assert.match(xu, /<filterColumn\b[^>]*hiddenButton="1"/, '仅按钮设置不视为活动筛选，保留按钮属性');
  const book = X.read(bytes, { type: 'buffer', cellNF: true });
  assert.deepEqual([book.Sheets.李奇.K2.v, book.Sheets.李奇.K3.v, book.Sheets.操思敏.K2.v], [1245, 1050, 7]);
  assert.deepEqual([book.Sheets.徐英.K2.v, book.Sheets.徐英.K3.v], [3, 5], '9 包括手工隐藏行，109 排除手工隐藏行');
  if (formulas) for (const name of book.SheetNames) for (const [address, cell] of Object.entries(book.Sheets[name])) if (!address.startsWith('!')) assert.equal(cell.f, undefined, name + '!' + address);
  return book;
}

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: true });
  try {
    const page = await browser.newPage(), errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8940/').href);
    const ready = () => page.waitForFunction(() => window.FulfillmentToolReady && window.FulfillmentValues && !document.getElementById('file').disabled);
    await ready();
    const prepare = buffer => page.evaluate(async bytes => {
      const input = Uint8Array.from(bytes).buffer, before = [...new Uint8Array(input)];
      const result = await FulfillmentValues.prepare(input);
      return { bytes: [...new Uint8Array(result.bytes)], inputUnchanged: JSON.stringify(before) === JSON.stringify([...new Uint8Array(input)]), issues: result.issues, formulaCount: result.formulaCount, clearedFilterCount: result.clearedFilterCount };
    }, [...buffer]);
    const upload = async (buffer, name) => {
      await page.locator('#file').setInputFiles({ name, mimeType, buffer });
      await page.waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
    };
    const history = () => page.evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(), (_, value) => value instanceof ArrayBuffer ? [...new Uint8Array(value)] : value)));

    const input = await fixture(), originalHash = digest(input), prepared = await prepare(input);
    assert.deepEqual(prepared.issues, []);
    assert.equal(prepared.clearedFilterCount, 2);
    assert.equal(prepared.inputUnchanged, true);
    await assertCleared(Buffer.from(prepared.bytes));
    assert.equal(digest(input), originalHash);

    await upload(input, '自动清除筛选更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#quantityIssues').innerText());
    assert.match(await page.locator('#formulaStatus').innerText(), /已自动清除\s*2\s*个工作表的筛选条件/);
    const downloadWait = page.waitForEvent('download');
    await page.locator('#download').click();
    const download = await downloadWait, exported = await fs.readFile(await download.path());
    const outputBook = X.read(exported, { type: 'buffer' });
    assert.equal(outputBook.SheetNames[0], '汇总表');
    assert.equal(X.utils.sheet_to_json(outputBook.Sheets.汇总表, { header: 1 }).length - 1, 13);
    // Export renumbers source sheet parts; the cached prepared source keeps original part names.
    await assertCleared(Buffer.from((await history()).current.sourceBytes));
    assert.deepEqual([outputBook.Sheets.李奇.K2.v, outputBook.Sheets.李奇.K3.v], [1245, 1050]);

    const noFormula = await fixture({ formulas: false }), preparedValues = await prepare(noFormula);
    assert.equal(preparedValues.formulaCount, 0);
    assert.equal(preparedValues.clearedFilterCount, 2);
    assert.deepEqual(preparedValues.issues, []);
    await assertCleared(Buffer.from(preparedValues.bytes), { formulas: false });

    // A legacy filtered snapshot must be re-merged even when there are no formulas to convert.
    await page.evaluate(async bytes => {
      const sourceBytes = Uint8Array.from(bytes).buffer;
      const result = FulfillmentCore.merge(XLSX.read(sourceBytes, { type: 'array', cellNF: true }), '旧筛选更新版.xlsx');
      const previous = { filename: '原历史.xlsx', savedAt: '2026-10-20T00:00:00.000Z', result: JSON.parse(JSON.stringify(result)) };
      result.rows.pop(); result.sources.pop();
      const current = { filename: '旧筛选更新版.xlsx', savedAt: '2026-10-23T00:00:00.000Z', sourceBytes, result };
      await FulfillmentCache.replace({ version: 1, previous, current });
    }, [...noFormula]);
    const oldHistory = await history();
    await page.reload(); await ready();
    const restored = await history();
    assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#quantityIssues').innerText());
    assert.equal(restored.current.result.rows.length, 13);
    assert.deepEqual(restored.previous, oldHistory.previous, '只迁移当前快照，不推进历史基准');
    assert.equal(restored.current.savedAt, oldHistory.current.savedAt);
    await assertCleared(Buffer.from(restored.current.sourceBytes), { formulas: false });
    assert.match(await page.locator('#formulaStatus').innerText(), /已自动清除\s*2\s*个工作表的筛选条件/);

    const broken = await fixture({ broken: true }), failed = await prepare(broken);
    assert(failed.issues.some(issue => /REF/.test(issue.message)), '清除筛选不绕过真实公式错误');
    assert.equal(digest(Buffer.from(failed.bytes)), digest(broken), '失败时不交付部分修改文件');
    assert.equal(failed.inputUnchanged, true);
    const beforeFailure = await history();
    await upload(broken, '错误公式更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), true);
    assert.deepEqual(await history(), beforeFailure, '公式失败不替换缓存');
    assert.deepEqual(errors, []);
    console.log('PASS: worksheet/table filters, SUBTOTAL 9/109, preserved manual hiding, filter-only cache migration, source integrity and formula failure');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
