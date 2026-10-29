const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js'), C = require('../fulfillment-core.js');
const names = ['徐英', '操思敏', '邢芳芳1017', '李奇'];
const divisions = ['海外事业一部', '海外事业二部', '国内事业部'];
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const out = process.env.QA_OUT || '/private/tmp/fulfillment-review-qa';
function fixture({delta = 0, badQuantity = false, badFormula = false} = {}) {
  const book = X.utils.book_new();
  for (const [person, name] of names.entries()) {
    const columns = C.COLUMNS.slice(person * 4).concat(C.COLUMNS.slice(0, person * 4));
    const records = [0, 1].map(index => ({
      '采购下单人': '表内其他人', '事业部': divisions[person % 3], '销售分渠道': '原渠道',
      '运营': ['倪芷涵', '王志豪', '王志豪', '薛文乐'][person],
      '下单日期': index ? '2026-09-30' : '2026-10-01', '下单月份': index ? '2026-09' : '2026-10',
      '物料编码': '000' + person + index, 'SKU': 'SKU-' + person + index, 'OA备货单号': '000' + person,
      '未交付数量': 1.25 + (person === 0 && index === 0 ? delta : 0), '已下单未备料未生产': 0,
      '已备料未生产': '', '生产中产品': 2.5, '完工未发产品': 3, '已发货数量': 4,
      ...(person === 1 && index === 1 ? {'事业部': ' ', '销售分渠道': ''} : {})
    }));
    const sheet = X.utils.aoa_to_sheet([columns.map(h => h === '销售分渠道' ? '销售分渠道部' : h), ...records.map(row => columns.map(key => row[key] ?? ''))]);
    if (person === 0 && badQuantity) sheet[X.utils.encode_cell({r: 1, c: columns.indexOf('未交付数量')})] = {t: 's', v: '100或200'};
    if (person === 2 && badFormula) sheet[X.utils.encode_cell({r: 1, c: columns.indexOf('销售分渠道')})] = {t: 'n', v: 0, f: '#REF!'};
    X.utils.book_append_sheet(book, sheet, name);
  }
  X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([['原说明'], ['此表不汇总']]), '孙立柱');
  return Buffer.from(X.write(book, {type: 'buffer', bookType: 'xlsx'}));
}
(async () => {
  await fs.mkdir(out, {recursive: true});
  const browser = await chromium.launch({headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
  try {
    const page = await browser.newPage({viewport: {width: 1440, height: 1000}}), errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.addInitScript(() => localStorage.setItem('gylsjqx-last-tool', 'fulfillment'));
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8940/');
    const tool = () => page.frames().find(frame => /\/fulfillment\.html/.test(frame.url()));
    await page.locator('iframe[data-tool="fulfillment"]').waitFor();
    const ready = async () => {await page.locator('iframe[data-tool="fulfillment"]').waitFor(); await tool().waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);};
    await ready();
    const upload = async (buffer, name = '核对2026.10.29更新版.xlsx') => {
      await tool().locator('#file').setInputFiles({name, mimeType, buffer});
      await tool().waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
    };
    const history = () => tool().evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(), (_, value) => value instanceof ArrayBuffer ? Array.from(new Uint8Array(value)) : value)));
    const download = async (selector, saveAs) => {
      const event = page.waitForEvent('download'); await tool().locator(selector).click();
      const file = await event, bytes = await fs.readFile(await file.path());
      if (saveAs) await fs.writeFile(path.join(out, saveAs), bytes);
      return {bytes, name: file.suggestedFilename(), book: X.read(bytes, {type: 'buffer', cellNF: true, cellStyles: true})};
    };
    await upload(fixture({delta: -1}), '基准更新版.xlsx');
    await upload(fixture());
    assert.equal(await tool().locator('#download').isDisabled(), false);
    assert.equal(await tool().locator('#download-review').isDisabled(), false);
    assert.equal(await tool().locator('#preview thead th').count(), 22);
    const owners = names.flatMap(name => [name, name]);
    assert.deepEqual(await tool().locator('#preview tbody tr td:first-child').allTextContents(), owners);
    const before = await history();
    assert(!before.current.result.review, '核对结果只在显示时生成，不写入历史数据');
    const full = await download('#download', '完整汇总.xlsx');
    assert.deepEqual(full.book.SheetNames, ['汇总表', '合并核对', '差异', ...names, '孙立柱']);
    const summary = X.utils.sheet_to_json(full.book.Sheets.汇总表, {header: 1, defval: ''});
    assert.deepEqual(summary[0], C.SUMMARY_COLUMNS); assert.deepEqual(summary.slice(1).map(row => row[0]), owners);
    assert(summary.slice(1).every(row => row[2] === '原渠道' || row[2] === ''), '仅提示，不自动修改渠道');
    assert.equal(full.book.Sheets.差异.F2.v, 1, '新增输出列不影响数量差异');
    const audit = X.utils.sheet_to_json(full.book.Sheets.合并核对, {header: 1, defval: ''});
    assert.deepEqual(audit[0], C.REVIEW_COLUMNS); assert.equal(audit.filter(row => row[0] === '数量合计').length, 30);
    assert(audit.some(row => row[8] === '海外事业一部*欧洲部*亚马逊英国组'));
    assert(audit.some(row => row[8] === '国内事业部*TEMU业务部'));
    assert(audit.some(row => row[0] === '事业部／渠道' && row[1] === '操思敏' && row[2] === '第 3 行'));
    assert.equal(full.book.Sheets.合并核对.A1.s.fgColor.rgb, '315B9D');
    assert.equal(full.book.Sheets.合并核对['!autofilter'].ref, `A1:N${audit.length}`);
    assert(full.book.Sheets.合并核对['!cols'][8].wch >= 36, '建议渠道按内容加宽');
    const totals = audit.findIndex(row => row[0] === '数量合计') + 1;
    assert.equal(full.book.Sheets.合并核对['J' + totals].t, 'n');
    assert.equal(full.book.Sheets.合并核对['J' + totals].z, '0.0'); // 1.25 + 1.25 = 2.5
    const standalone = await download('#download-review', '合并核对.xlsx');
    assert.equal(standalone.name, '核对2026.10.29更新版－合并核对.xlsx');
    assert.deepEqual(standalone.book.SheetNames, ['合并核对']);
    assert.deepEqual(X.utils.sheet_to_json(standalone.book.Sheets.合并核对, {header: 1, defval: ''}), audit);
    const preview = await tool().locator('#review-preview tbody tr').evaluateAll(rows => rows.map(row => [...row.cells].map(cell => cell.textContent)));
    assert.deepEqual(preview, audit.slice(1).map(row => row.map(String)), '预览和两种下载一致');
    for (const division of divisions) {
      const split = await download('[data-download-division="' + division + '"]');
      assert.deepEqual(split.book.SheetNames, ['汇总表']);
      const rows = X.utils.sheet_to_json(split.book.Sheets.汇总表, {header: 1, defval: ''});
      assert.deepEqual(rows[0], C.SUMMARY_COLUMNS);
      assert.deepEqual(rows.slice(1), summary.slice(1).filter(row => row[1] === division), '事业部下载保留对应来源及全部字段');
    }
    // The existing source-preserving exporter must not alter original worksheet XML.
    const source = await Zip.loadAsync(Buffer.from(before.current.sourceBytes)), output = await Zip.loadAsync(full.bytes);
    for (const entry of Object.keys(source.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) assert.equal(await source.file(entry).async('string'), await output.file(entry).async('string'));
    const xml = await output.file('xl/worksheets/fulfillment2.xml').async('string'), styles = await output.file('xl/styles.xml').async('string');
    assert.match(xml, /autoFilter/); assert.match(styles, /style="thin"/); assert.match(styles, /wrapText="0"/);
    await page.locator('[data-tool="payment"]').first().click();
    await page.locator('button[data-tool="fulfillment"]').click();
    assert.deepEqual(await history(), before, '切换页面不丢失结果或推进基准');
    await tool().locator('#mergeReview summary').click();
    await page.screenshot({path: path.join(out, 'homepage-review.png'), fullPage: true});
    // Older cache: missing provenance must be rebuilt from the saved file, not purchaser names.
    await tool().evaluate(async () => {const history = await FulfillmentCache.get(); delete history.current.result.sources; history.current.result.review = [[{t:'s',v:'旧的错误核对'}]]; await FulfillmentCache.replace(history);});
    await page.reload(); await ready();
    const restored = await history();
    assert.deepEqual(restored.previous, before.previous); assert.equal(restored.current.savedAt, before.current.savedAt);
    assert.deepEqual(restored.current.result.rows, before.current.result.rows);
    assert.deepEqual(restored.current.result.sources, before.current.result.sources);
    assert.deepEqual(await tool().locator('#preview tbody tr td:first-child').allTextContents(), owners);
    assert.deepEqual((await download('#download-review')).book.Sheets.合并核对['!ref'], standalone.book.Sheets.合并核对['!ref']);
    await upload(fixture()); const stable = await history(); assert.deepEqual(stable.previous, restored.previous, '相同内容重传不推进历史'); assert.deepEqual(stable.current.result.rows, restored.current.result.rows);
    for (const options of [{badQuantity: true}, {badFormula: true}]) {
      await upload(fixture(options), options.badQuantity ? '数量异常.xlsx' : '必要公式异常.xlsx');
      assert.equal(await tool().locator('#download').isDisabled(), true);
      for (const division of divisions) assert.equal(await tool().locator('[data-download-division="' + division + '"]').isDisabled(), true);
      assert.equal(await tool().locator('#download-review').isDisabled(), false);
      const check = await download('#download-review', options.badQuantity ? '数量异常核对.xlsx' : '公式异常核对.xlsx');
      assert.deepEqual(check.book.SheetNames, ['合并核对']);
      const problems = X.utils.sheet_to_json(check.book.Sheets.合并核对);
      assert(problems.some(row => row.核对类型 === '数量／转换异常' && row.需要处理.includes(options.badQuantity ? '多个数字' : '#REF!')));
      assert.deepEqual(await history(), stable, '失败不覆盖历史');
    }
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({status: 'PASS', url: page.url(), summaryColumns: 22, quantityChecks: 30, downloads: 'full + 3 divisions + review', sourcePreserved: true, cache: 'rebuilt without advancing baseline', blockedQuantityAndFormula: true, output: out}));
  } finally {await browser.close();}
})().catch(error => {console.error(error); process.exitCode = 1;});
