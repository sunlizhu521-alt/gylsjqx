const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const X = require('../vendor/xlsx.full.min.js');
const Fulfillment = require('../fulfillment-core.js');
const LedgerExport = require('../funds-ledger-export.js');
const BudgetExport = require('../quarter-budget-export.js');

const values = [1378, 0, -60, 0.25, -0.125, 1234.56789];
function check(cell, expected) {
  assert.equal(cell.t, 'n');
  assert.equal(cell.v, expected);
  assert(!/\.\s*$/.test(cell.w), `${expected} 显示为 ${cell.w}`);
}
function reopen(book) {
  return X.read(X.write(book, {type:'buffer', bookType:'xlsx'}), {type:'buffer', cellNF:true});
}

test('全局导出不声明会让整数显示末尾小数点的可选小数格式', () => {
  const root = path.resolve(__dirname, '..');
  function scan(dir) {
    for (const entry of fs.readdirSync(dir, {withFileTypes:true})) {
      if (entry.name.startsWith('.') || ['vendor','node_modules','tests'].includes(entry.name)) continue;
      const file = path.join(dir, entry.name);
      if (entry.isDirectory()) scan(file);
      else if (/\.(?:js|html)$/.test(entry.name)) {
        const source = fs.readFileSync(file, 'utf8');
        for (const match of source.matchAll(/(["'])((?:#,##)?0\.[#?]+)\1/g)) {
          // A mandatory decimal digit (money: 0.00) and value-specific precision are safe.
          assert(!/\.\s*$/.test(X.SSF.format(match[2], 0)), `${path.relative(root,file)} 使用了 ${match[2]}；整数应使用 0，小数应按实际精度设置格式`);
        }
      }
    }
  }
  scan(root);
});

test('订单履约汇总和合并核对保留小数，整数不带末尾点', () => {
  const row = Fulfillment.COLUMNS.map(key => ({t:'s',v:key === '物料编码' ? '000123' : ''}));
  Fulfillment.QUANTITY_FIELDS.forEach((key, i) => {row[Fulfillment.COLUMNS.indexOf(key)] = {t:'n',v:values[i]};});
  const output = reopen(Fulfillment.workbook({rows:[row],sources:[{sheet:'李奇'}]})).Sheets.汇总表;
  Fulfillment.QUANTITY_FIELDS.forEach((key, i) => check(output[X.utils.encode_cell({r:1,c:Fulfillment.SUMMARY_COLUMNS.indexOf(key)})], values[i]));
  assert.equal(output.K2.v, '000123');
  const reviewRows = Fulfillment.reviewRows({rows:[],sources:[]}, {checks:[{sheet:'李奇',field:'未交付数量',sourceTotal:1378,summaryTotal:1378,difference:0,status:'通过'}],issues:[]});
  const review = reopen({SheetNames:['核对'],Sheets:{核对:Fulfillment.reviewSheet(reviewRows)}});
  assert.equal(review.Sheets.核对.J2.v, 1378);
  for (const cell of Object.values(review.Sheets.核对)) if (cell?.t === 'n') assert(!/\.\s*$/.test(cell.w));
});

test('资金台账导出保留金额小数位和空白，不出现孤立末尾点', async () => {
  const headers = ['金额','是否已付款'];
  const rows = values.map(v => [{t:'n',v,z:'#,##0.00########'}, {t:'s',v:''}]);
  const result = {sheets:[{name:'正常付款',headers,rows}]};
  const output = X.read(await LedgerExport(result), {type:'array',cellNF:true}).Sheets.正常付款;
  values.forEach((v,i) => check(output['A'+(i+2)], v));
  assert.equal(output.A2.w, '1,378.00');
  assert.equal(output.B2.v, '');
});

test('季度预算导出的明细、汇总和合计保留两位金额，不出现孤立末尾点', async () => {
  const amounts = [1378,0,-60,0.25,-0.12];
  const cents = amounts.reduce((sum,v) => sum + Math.round(v*100), 0);
  const totals = {totalCents:cents,paidCents:0,remainingCents:cents};
  const result = {records:amounts.map(amount => ({monthKey:'2026-11',supplier:'供应商',amount,cents:Math.round(amount*100)})),summary:{classified:[{monthKey:'2026-11',category:'正常付款',...totals}],monthly:[{monthKey:'2026-11',...totals}],totals}};
  const output = X.read(await BudgetExport(result), {type:'array',cellNF:true});
  amounts.forEach((v,i) => check(output.Sheets.付款汇总明细['F'+(i+2)], v));
  for (const name of output.SheetNames) for (const cell of Object.values(output.Sheets[name])) if (cell?.t === 'n') assert(!/\.\s*$/.test(cell.w));
  assert.equal(output.Sheets.付款汇总明细.F2.w, '1,378.00');
});
