const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../fulfillment-core.js');

// All fixtures are synthetic and safe to run against the deployed page.
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const quantities = ['未交付数量', '已下单未备料未生产', '已备料未生产', '生产中产品', '完工未发产品', '已发货数量'];
const divisions = ['海外事业一部', '海外事业二部', '国内事业部'];
const materialCodes = ['1007010122', '1007010122', '1007010123', '1007010123', '1007010123'];
const detailValues = [1245, 50, 1050, 50, 140];
const hash = bytes => crypto.createHash('sha256').update(bytes).digest('hex');

async function fixture({ unsupported = false, formulas = true, brokenAuxiliary = false, quantityDependency = false } = {}) {
  const book = X.utils.book_new();
  for (const [index, name] of [...C.SHEETS, '孙立柱'].entries()) {
    const rowCount = name === '李奇' ? 5 : 1;
    const rows = Array.from({ length: rowCount }, (_, row) => C.COLUMNS.map(field => {
      if (quantities.includes(field)) return 0;
      return ({
        '采购下单人': name === '操思敏' ? '陈燕' : name,
        '事业部': divisions[index % divisions.length],
        '供应商简称': '合成供应商',
        '物料编码': name === '李奇' ? materialCodes[row] : 'SYN-' + index,
        'SKU': 'SKU-' + index + '-' + row,
        '物料名称': '合成物料',
        '下单日期': '2026年10月23日'
      })[field] ?? '';
    }));
    const sheet = X.utils.aoa_to_sheet([C.COLUMNS, ...rows]);
    if (name === '李奇') {
      for (const [index, value] of detailValues.entries()) {
        const row = index + 2;
        sheet['AA' + row] = { t: 'n', v: value };
        sheet['K' + row] = formulas
          ? { t: 'n', v: 0, f: 'SUBTOTAL(9,AA' + row + ':AD' + row + ')' }
          : { t: 'n', v: value };
      }
      sheet['!ref'] = 'A1:AD6';
    }
    if (name === '孙立柱') sheet.K2 = formulas ? { t: 'n', v: 0, f: "'辅助'!A2" } : { t: 'n', v: 11 };
    if (name === '操思敏' && brokenAuxiliary) {
      for (const address of Object.keys(sheet).filter(key => /^[A-Z]+2$/.test(key))) {
        sheet[address.replace(/2$/, '159')] = sheet[address]; delete sheet[address];
      }
      sheet.AI1 = { t: 's', v: '原表数量校验' }; sheet.AJ1 = { t: 's', v: '原表检查结果' };
      sheet.AI159 = { t: 'e', v: 23, f: 'AG159+AH159-R159-#REF!' };
      sheet.AJ159 = { t: 'e', v: 23, f: 'IF(AI159=0,"一致","不一致")' };
      if (quantityDependency) sheet.K159 = { t: 'n', v: 0, f: 'AI159' };
      sheet['!ref'] = 'A1:AJ159';
    }
    X.utils.book_append_sheet(book, sheet, name === '邢芳芳' ? '邢芳芳1017' : name);
  }
  const helper = X.utils.aoa_to_sheet([['计算辅助', '', '', '', '判断', '空白结果', '数值判断'], [11, 1256, 0, 0.3, '一致', '', 11], ['保留合并内容']]);
  if (formulas) {
    helper.A2 = { t: 'n', v: 0, f: '5+6', z: '0.000' };
    helper.B2 = { t: 'n', v: 0, f: "'李奇'!K2+'孙立柱'!K2" };
    helper.C2 = { t: 'n', v: 100, f: '0' };
    helper.D2 = { t: 'n', v: 0, f: unsupported ? 'UNKNOWNFUNCTION(A2)' : '0.1+0.2' };
    helper.E2 = { t: 's', v: '不一致', f: 'IF(C2=0,"一致","不一致")' };
    helper.F2 = { t: 's', v: '旧值', f: 'IF(C2<>0,"有差异","")' };
    helper.G2 = { t: 'n', v: 0, f: 'IF(C2=0,A2,1/0)' };
  }
  helper['!merges'] = [{ s: { r: 2, c: 0 }, e: { r: 2, c: 2 } }];
  helper['!cols'] = [{ wch: 26 }, { wch: 18 }, { wch: 14 }];
  helper['!rows'] = [{ hpt: 25 }, { hpt: 31 }, { hpt: 35 }];
  X.utils.book_append_sheet(book, helper, '辅助');
  const bytes = Buffer.from(X.write(book, { type: 'buffer', bookType: 'xlsx' }));
  if (!formulas) return bytes;
  const zip = await Zip.loadAsync(bytes);
  zip.file('xl/calcChain.xml', '<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="K2" i="4"/></calcChain>');
  const rels = 'xl/_rels/workbook.xml.rels';
  zip.file(rels, (await zip.file(rels).async('string')).replace('</Relationships>', '<Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

function shiftedHeadersFixture() {
  const book = X.utils.book_new(), expected = [];
  for (const [person, name] of C.SHEETS.entries()) {
    // Four independent permutations: each quantity field occupies a different column in every sheet.
    const order = C.COLUMNS.map((_, column) => C.COLUMNS[(column * (person % 2 ? 7 : 5) + person * 3 + 1) % C.COLUMNS.length]);
    const headers = order.map((field, column) => column % 2 ? ' ' + field + ' ' : field.slice(0, 2) + '\n' + field.slice(2));
    const values = Object.fromEntries(quantities.map((field, q) => [field, (person + 1) * 100 + (q + 1) * 7]));
    Object.assign(values, { '采购下单人': name, '事业部': divisions[person % divisions.length], '物料编码': 'SHIFT-' + person, 'SKU': 'SHIFT-SKU-' + person, '供应商简称': '移位合成供应商', '物料名称': '按表头取值物料' });
    const dataRow = person + 3;
    const sheet = X.utils.aoa_to_sheet([
      ...Array.from({ length: person + 1 }, (_, row) => row ? [] : ['各人员独立列顺序']),
      headers,
      order.map(field => values[field] ?? '')
    ]);
    const address = field => X.utils.encode_cell({ r: dataRow - 1, c: order.indexOf(field) });
    if (person === 0) sheet[address('未交付数量')] = { t: 'n', v: 0, f: address('已发货数量') + '-35' };
    if (person === 3) sheet[address('生产中产品')] = { t: 'n', v: 0, f: address('已备料未生产') + '+7' };
    X.utils.book_append_sheet(book, sheet, name);
    expected.push({ sheet: name, code: values.物料编码, values, columns: quantities.map(field => order.indexOf(field)) });
  }
  for (let q = 0; q < quantities.length; q++) assert.equal(new Set(expected.map(row => row.columns[q])).size, C.SHEETS.length);
  return { bytes: Buffer.from(X.write(book, { type: 'buffer', bookType: 'xlsx' })), expected };
}

function assertNoFormulas(book) {
  for (const name of book.SheetNames) for (const [address, cell] of Object.entries(book.Sheets[name])) {
    if (address.startsWith('!')) continue;
    for (const key of ['f', 'F', 'D']) assert.equal(cell[key], undefined, name + '!' + address + '不应保留' + key);
  }
}

function mergedTextFixture() {
  const book=X.utils.book_new();
  for(const [person,name] of C.SHEETS.entries()){
    const columns=C.COLUMNS.slice(person*4).concat(C.COLUMNS.slice(0,person*4));if(person%2)columns.reverse();
    const rows=Array.from({length:3},(_,row)=>columns.map(field=>quantities.includes(field)?(row?'':field==='生产中产品'?'物料通用   474套左右':'约'+(100+quantities.indexOf(field))+'件'):({'物料编码':'MERGED-'+person+'-'+row,'事业部':divisions[person%divisions.length],'物料名称':'合并数量测试'})[field]??''));
    const sheet=X.utils.aoa_to_sheet([columns,...rows]);sheet['!merges']=quantities.map(field=>({s:{r:1,c:columns.indexOf(field)},e:{r:3,c:columns.indexOf(field)}}));X.utils.book_append_sheet(book,sheet,name);
  }
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}

(async () => {
  const browser = await chromium.launch({
    executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true
  });
  try {
    const page = await browser.newPage(), errors = [];
    page.on('pageerror', error => errors.push(error.message));
    async function xmlSnapshot(text, styles = false) {
      return page.evaluate(({ text, styles }) => {
        const doc = new DOMParser().parseFromString(text, 'application/xml');
        const attrs = element => Object.fromEntries([...element.attributes].filter(a => a.prefix !== 'xmlns' && a.name !== 'xmlns').map(a => [a.name, a.value]).sort(([a], [b]) => a.localeCompare(b)));
        const node = element => ({ name: element.localName, attrs: attrs(element), children: [...element.children].map(node), text: element.children.length ? '' : element.textContent });
        if (styles) return Object.fromEntries([...doc.documentElement.children].map(element => [element.localName, node(element)]));
        return {
          structure: [...doc.documentElement.children].filter(element => element.localName !== 'sheetData').map(node),
          rows: [...doc.getElementsByTagNameNS('*', 'row')].map(attrs),
          cellStyles: [...doc.getElementsByTagNameNS('*', 'c')].map(cell => [cell.getAttribute('r'), cell.getAttribute('s') || '0'])
        };
      }, { text, styles });
    }
    const url = new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8940/').href;
    await page.goto(url);
    const ready = () => page.waitForFunction(() => window.FulfillmentToolReady && window.FulfillmentValues && !document.getElementById('file').disabled);
    await ready();
    async function upload(buffer, name) {
      await page.locator('#file').setInputFiles({ name, mimeType, buffer });
      await page.waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
    }
    async function history() {
      return page.evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(),
        (_, value) => value instanceof ArrayBuffer ? Array.from(new Uint8Array(value)) : value)));
    }
    async function getDownload(selector = '#download') {
      const wait = page.waitForEvent('download');
      await page.locator(selector).click();
      const dl = await wait;
      return { name: dl.suggestedFilename(), bytes: await fs.readFile(await dl.path()) };
    }
    async function assertBlocked() {
      assert.equal(await page.locator('#download').isDisabled(), true);
      for (const division of divisions) assert.equal(await page.locator('[data-download-division="' + division + '"]').isDisabled(), true);
    }
    async function putLegacy(buffer, { corrupt = false } = {}) {
      await page.evaluate(async ({ bytes, corrupt }) => {
        const sourceBytes = Uint8Array.from(bytes).buffer;
        const result = FulfillmentCore.merge(XLSX.read(sourceBytes, { type: 'array', cellNF: true }), '旧版更新版.xlsx');
        if (corrupt) result.rows[0][FulfillmentCore.COLUMNS.indexOf('未交付数量')] = { t: 'n', v: 999999 };
        const previous = { filename: '保留历史.xlsx', savedAt: '2026-10-20T00:00:00.000Z', result: { ...result, rows: [] } };
        const current = { filename: '旧版更新版.xlsx', savedAt: '2026-10-23T00:00:00.000Z', sourceBytes, result };
        await new Promise((resolve, reject) => {
          const open = indexedDB.open('gylsjqx-fulfillment-history', 1);
          open.onerror = () => reject(open.error);
          open.onsuccess = () => {
            const db = open.result, tx = db.transaction('snapshots', 'readwrite');
            tx.objectStore('snapshots').put({ version: 1, previous, current }, 'history');
            tx.oncomplete = () => { db.close(); resolve(); };
            tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
          };
        });
      }, { bytes: [...buffer], corrupt });
      const before = await history();
      await page.reload(); await ready();
      return before;
    }

    const input = await fixture(), inputHash = hash(input);
    await upload(input, '合成订单2026.10.23更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#status').innerText());
    assert.match(await page.locator('#quantityStatus').innerText(), /通过/);
    assert.equal(await page.locator('#quantityChecks tbody tr').count(), 30);
    const cached = await history();
    const preparedBytes = Buffer.from(cached.current.sourceBytes);
    const prepared = X.read(preparedBytes, { type: 'buffer', cellNF: true, cellStyles: true });
    assertNoFormulas(prepared);
    assert.deepEqual(detailValues.map((_, index) => prepared.Sheets.李奇['K' + (index + 2)].v), detailValues);
    assert.equal(prepared.Sheets.孙立柱.K2.v, 11, '不参与汇总的工作表也计算并转值');
    assert.equal(prepared.Sheets.辅助.B2.v, 1256, '跨工作表公式读取重算后的依赖');
    assert.equal(prepared.Sheets.辅助.C2.v, 0, '公式得到的零必须保留');
    assert(Math.abs(prepared.Sheets.辅助.D2.v - 0.3) < 1e-12);
    assert.deepEqual([prepared.Sheets.辅助.E2.t, prepared.Sheets.辅助.E2.v], ['s', '一致']);
    assert.deepEqual([prepared.Sheets.辅助.F2.t, prepared.Sheets.辅助.F2.v], ['s', '']);
    assert.equal(prepared.Sheets.辅助.G2.v, 11, 'IF只计算所选分支，不读取旧缓存或计算未选分支的除零');
    assert.equal(prepared.Sheets.辅助.A2.z, '0.000');
    assert.equal(prepared.Sheets.辅助['!merges'].length, 1);
    const [sourceZip, preparedZip] = await Promise.all([Zip.loadAsync(input), Zip.loadAsync(preparedBytes)]);
    const oldStyles = await xmlSnapshot(await sourceZip.file('xl/styles.xml').async('string'), true);
    const newStyles = await xmlSnapshot(await preparedZip.file('xl/styles.xml').async('string'), true);
    for (const [name, before] of Object.entries(oldStyles)) {
      const after = newStyles[name];
      assert.ok(after, '原样式分组仍存在：' + name);
      if (['numFmts', 'cellXfs'].includes(name)) {
        const { count: oldCount, ...oldAttrs } = before.attrs, { count: newCount, ...newAttrs } = after.attrs;
        assert.deepEqual(newAttrs, oldAttrs, '样式分组属性保持：' + name);
        assert.deepEqual(after.children.slice(0, before.children.length), before.children, '保留旧编号和原样式定义，仅追加新格式：' + name);
        assert.equal(Number(newCount), after.children.length);
      } else assert.deepEqual(after, before, '字体、填充、边框及其他原样式保持：' + name);
    }
    for (const name of Object.keys(newStyles)) assert.ok(oldStyles[name] || name === 'numFmts', '只允许新增数字格式分组');
    for (let index = 1; index <= 6; index++) {
      const file = 'xl/worksheets/sheet' + index + '.xml';
      const before = await xmlSnapshot(await sourceZip.file(file).async('string'));
      const after = await xmlSnapshot(await preparedZip.file(file).async('string'));
      assert.deepEqual(after.structure, before.structure, '原工作表合并、列宽及其他结构保持：' + file);
      assert.deepEqual(after.rows, before.rows, '原工作表行高等属性保持：' + file);
      if (index >= 5) assert.deepEqual(after.cellStyles, before.cellStyles, '孙立柱及辅助表格式不变：' + file);
    }
    for (const name of ['徐英', '操思敏', '邢芳芳1017', '李奇']) {
      const sheet = prepared.Sheets[name];
      for (const field of quantities) {
        const cell = sheet[X.utils.encode_cell({ r: 1, c: C.COLUMNS.indexOf(field) })];
        assert.equal(cell.t, 'n', name + field + '使用真实数值');assert.equal(cell.z, '0');
      }
      assert.equal(sheet.X2.t, 'n');assert.equal(sheet.X2.z, C.DATE_FORMAT, '下单日期使用日期格式');
      assert.equal(sheet.H2.t, 's');assert.equal(sheet.H2.z, '@', '物料编码使用文本格式');
      assert.equal(sheet.I2.t, 's');assert.equal(sheet.I2.z, '@', 'SKU 使用文本格式');
    }
    assert.equal(preparedZip.file('xl/calcChain.xml'), null);
    assert.doesNotMatch(await preparedZip.file('xl/_rels/workbook.xml.rels').async('string'), /calcChain/);
    assert.doesNotMatch(await preparedZip.file('[Content_Types].xml').async('string'), /calcChain/);

    const download = await getDownload();
    assert.equal(download.name, '合成订单2026.10.23汇总版.xlsx');
    const exported = X.read(download.bytes, { type: 'buffer', cellNF: true });
    assertNoFormulas(exported);
    assert.deepEqual([exported.Sheets.辅助.E2.v, exported.Sheets.辅助.F2.v, exported.Sheets.辅助.G2.v], ['一致', '', 11]);
    assert.deepEqual(exported.SheetNames, ['汇总表', '合并核对', '差异', '徐英', '操思敏', '邢芳芳1017', '李奇', '孙立柱', '辅助']);
    const rows = X.utils.sheet_to_json(exported.Sheets.汇总表, { header: 1, defval: '' });
    assert.deepEqual(rows[0], C.SUMMARY_COLUMNS);
    assert.equal(rows.length - 1, 8, '保留全部明细，不因公式删行或合并物料');
    assert.deepEqual(rows.slice(-5).map(row => [row[C.SUMMARY_COLUMNS.indexOf('物料编码')], row[C.SUMMARY_COLUMNS.indexOf('未交付数量')]]), materialCodes.map((code, index) => [code, detailValues[index]]));
    assert.equal(rows.filter(row => row[C.SUMMARY_COLUMNS.indexOf('物料编码')] === materialCodes[0]).reduce((sum, row) => sum + row[C.SUMMARY_COLUMNS.indexOf('未交付数量')], 0), 1295);
    assert.equal(rows.filter(row => row[C.SUMMARY_COLUMNS.indexOf('物料编码')] === materialCodes[2]).reduce((sum, row) => sum + row[C.SUMMARY_COLUMNS.indexOf('未交付数量')], 0), 1240);
    const outputZip = await Zip.loadAsync(download.bytes);
    for (const name of Object.keys(preparedZip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
      assert.equal(await outputZip.file(name).async('string'), await preparedZip.file(name).async('string'), '值版原表导出时完整保留：' + name);
    }
    for (const name of Object.keys(outputZip.files).filter(name => /^xl\/worksheets\/.*\.xml$/.test(name))) assert.doesNotMatch(await outputZip.file(name).async('string'), /<(?:\w+:)?f(?:\s|>)/, name);
    const divided = X.read((await getDownload('[data-download-division="海外事业一部"]')).bytes, { type: 'buffer' });
    assert.deepEqual(divided.SheetNames, ['汇总表']);
    assertNoFormulas(divided);
    assert.equal(hash(input), inputHash, '上传原文件不修改');
    assert.equal(X.read(input, { type: 'buffer' }).Sheets.李奇.K2.v, 0);
    assert.equal(X.read(input, { type: 'buffer' }).Sheets.李奇.K2.f, 'SUBTOTAL(9,AA2:AD2)');

    const goodHistory = await history();
    await upload(await fixture({ unsupported: true }), '无法计算的辅助表.xlsx');
    await assertBlocked();
    const message = (await page.locator('#status').innerText()) + '\n' + (await page.locator('#quantityIssues').innerText());
    assert.match(message, /辅助/); assert.match(message, /D2/); assert.match(message, /公式|函数|计算/);
    assert.deepEqual(await history(), goodHistory, '失败不覆盖历史或允许旧结果下载');

    const brokenInput = await fixture({ brokenAuxiliary: true }), brokenHash = hash(brokenInput);
    await upload(brokenInput, '辅助校验失效更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#quantityIssues').innerText());
    assert.match(await page.locator('#formulaStatus').innerText(), /2 个不参与汇总的辅助列公式已移除/);
    assert.equal(await page.locator('#quantityIssues').innerText(), '');
    const repaired = X.read((await getDownload()).bytes, { type: 'buffer', cellNF: true });
    assertNoFormulas(repaired);
    assert.equal(repaired.Sheets.操思敏.AI159.v, '#REF!'); assert.equal(repaired.Sheets.操思敏.AJ159.v, '#REF!');
    assert.equal(repaired.Sheets.操思敏.AI159.t, 's'); assert.equal(repaired.Sheets.操思敏.AI159.z, '@');
    assert.equal(repaired.Sheets.李奇.K2.v, 1245, '数量公式仍重算，不能采用旧零');
    assert.equal(X.utils.sheet_to_json(repaired.Sheets.汇总表, { header: 1 }).length, 9, '移除公式不删除数据行');
    for (const division of divisions) {
      const divided = X.read((await getDownload('[data-download-division="' + division + '"]')).bytes, { type: 'buffer' });
      assert.deepEqual(divided.SheetNames, ['汇总表']); assertNoFormulas(divided);
    }
    assert.equal(hash(brokenInput), brokenHash);
    const repairedHistory = await history();
    await page.reload(); await ready();
    assert.equal(await page.locator('#download').isDisabled(), false);
    assert.match(await page.locator('#formulaStatus').innerText(), /2 个不参与汇总的辅助列公式已移除/);
    assert.deepEqual(await history(), repairedHistory, '重新打开不推进历史基准');
    await upload(await fixture({ brokenAuxiliary: true, quantityDependency: true }), '数量依赖失效更新版.xlsx');
    await assertBlocked();
    assert.match(await page.locator('#quantityIssues').innerText(), /AI159/);
    assert.deepEqual(await history(), repairedHistory, '数量依赖坏公式仍阻止，不可用辅助列旧值绕过');

    const legacyBefore = await putLegacy(input);
    assert.equal(await page.locator('#download').isDisabled(), false, '旧缓存有公式时须迁移而非沿用旧值');
    const legacyAfter = await history();
    assert.equal(legacyAfter.current.filename, legacyBefore.current.filename);
    assert.equal(legacyAfter.current.savedAt, legacyBefore.current.savedAt);
    assert.deepEqual(legacyAfter.previous, legacyBefore.previous, '迁移当前公式不推进或改写上次历史');
    assert.deepEqual(legacyAfter.current.result.rows.slice(-5).map(row => row[C.COLUMNS.indexOf('未交付数量')].v), detailValues);
    assertNoFormulas(X.read(Buffer.from(legacyAfter.current.sourceBytes), { type: 'buffer' }));
    await page.reload(); await ready();
    assert.equal(await page.locator('#download').isDisabled(), false);
    assert.deepEqual(await history(), legacyAfter, '迁移后重复打开不再改历史');

    const normalizedWithoutFormulas = Buffer.from(await page.evaluate(async bytes => Array.from(new Uint8Array((await FulfillmentValues.prepare(Uint8Array.from(bytes).buffer)).bytes)), [...await fixture({ formulas: false })]));
    const corruptBefore = await putLegacy(normalizedWithoutFormulas, { corrupt: true });
    await assertBlocked();
    assert.deepEqual(await history(), corruptBefore, '无公式缓存结果不一致仍阻止，不能静默重建');
    await upload(input, '恢复更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false);

    const shifted = shiftedHeadersFixture();
    await upload(shifted.bytes, '四人独立列位置更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#status').innerText());
    assert.match(await page.locator('#quantityStatus').innerText(), /通过/);
    assert.equal(await page.locator('#quantityChecks tbody tr').count(), 30);
    const shiftedExport = X.read((await getDownload()).bytes, { type: 'buffer' });
    assertNoFormulas(shiftedExport);
    const shiftedRows = X.utils.sheet_to_json(shiftedExport.Sheets.汇总表, { header: 1, defval: '' });
    assert.deepEqual(shiftedRows[0], C.SUMMARY_COLUMNS);
    assert.equal(shiftedRows.length - 1, C.SHEETS.length);
    for (const [index, expected] of shifted.expected.entries()) {
      const row = shiftedRows[index + 1];
      assert.equal(row[C.SUMMARY_COLUMNS.indexOf('物料编码')], expected.code);
      for (const field of quantities) assert.equal(row[C.SUMMARY_COLUMNS.indexOf(field)], expected.values[field], expected.sheet + '按表头读取“' + field + '”');
    }

    const mergedInput=mergedTextFixture(),mergedHash=hash(mergedInput);await upload(mergedInput,'合并文本数量更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(),false,await page.locator('#quantityIssues').innerText());
    const mergedOutput=X.read((await getDownload()).bytes,{type:'buffer'}),mergedRows=X.utils.sheet_to_json(mergedOutput.Sheets.汇总表,{header:1});
    assert.equal(mergedRows.length,13);
    for(let i=0;i<12;i++)for(const field of quantities)assert.equal(mergedRows[i+1][C.SUMMARY_COLUMNS.indexOf(field)],i%3?0:field==='生产中产品'?474:100+quantities.indexOf(field));
    const mergedSourceZip=await Zip.loadAsync(mergedInput),mergedOutputZip=await Zip.loadAsync((await getDownload()).bytes);
    for(const path of Object.keys(mergedSourceZip.files).filter(p=>/^xl\/worksheets\/sheet\d+\.xml$/.test(p))){
      const before=await xmlSnapshot(await mergedSourceZip.file(path).async('string')),after=await xmlSnapshot(await mergedOutputZip.file(path).async('string'));
      assert.deepEqual(after.structure,before.structure,'数量转值保留原合并、列宽和其他结构');assert.deepEqual(after.rows,before.rows,'数量转值保留原行属性');
    }
    for(const name of C.SHEETS){const sheet=mergedOutput.Sheets[name];for(const field of quantities){const column=Object.keys(sheet).find(address=>/^[A-Z]+1$/.test(address)&&sheet[address].v===field).replace(/1$/,'');for(let row=2;row<=4;row++){const cell=sheet[column+row];assert.equal(cell.t,'n');assert.equal(cell.v,row===2?(field==='生产中产品'?474:100+quantities.indexOf(field)):0,'源人员表合并数量只首行计入');}}}
    assert.equal(hash(mergedInput),mergedHash,'上传原文件不改写');assert.match(await mergedSourceZip.file('xl/worksheets/sheet1.xml').async('string'),/物料通用/,'原输入仍保留文本数量');
    await page.evaluate(async bytes=>{const h=await FulfillmentCache.get(),C=FulfillmentCore;h.current.sourceBytes=Uint8Array.from(bytes).buffer;for(let i=0;i<h.current.result.rows.length;i++)for(const field of C.QUANTITY_FIELDS)h.current.result.rows[i][C.COLUMNS.indexOf(field)]={t:'s',v:i%3?'':field==='生产中产品'?'物料通用   474套左右':'约'+(100+C.QUANTITY_FIELDS.indexOf(field))+'件'};await FulfillmentCache.replace(h);},[...mergedInput]);
    const oldQuantityHistory=await history();await page.reload();await ready();assert.equal(await page.locator('#download').isDisabled(),false);const normalizedHistory=await history();assert.equal(normalizedHistory.current.savedAt,oldQuantityHistory.current.savedAt);assert.deepEqual(normalizedHistory.previous,oldQuantityHistory.previous);
    assert.deepEqual(normalizedHistory.current.result.rows.slice(0,3).map(row=>row[C.COLUMNS.indexOf('生产中产品')].v),[474,0,0]);
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ url, detailRows: 8, codes: { '1007010122': 1295, '1007010123': 1240 }, allSheetsValuesOnly: 'passed', crossSheet: 'passed', preservedStructure: 'passed', legacyMigration: 'passed', blockedUnsupported: 'passed', independentHeaderPositions: 'passed', ifResults:'passed', mergedTextQuantities:'passed' }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
