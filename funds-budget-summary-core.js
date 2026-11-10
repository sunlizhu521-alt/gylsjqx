(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FundsBudgetSummaryCore = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const HEADERS = ['付款月份','分类','状态','申请总金额','已付款金额','剩余付款金额','备注'];
  const INPUT_HEADERS = ['预估付款月份','分类','申请付款金额','是否已支付'];
  const CATEGORIES = new Map([
    ['谈判付款表-未暂停', {category:'谈判付款表',state:'未暂停'}],
    ['谈判付款表-暂停付款', {category:'谈判付款表',state:'暂停付款'}],
    ['正常付款', {category:'正常付款',state:'正常付款'}]
  ]);
  const STATE_ORDER = ['未暂停','暂停付款','正常付款'];
  const LIMIT = BigInt(Number.MAX_SAFE_INTEGER);
  const text = value => String(value ?? '').trim();
  const hasValue = cell => cell && (cell.f || cell.t === 'e' || cell.v != null && text(cell.v) !== '');
  const header = value => text(value).replace(/\s/g, '');

  function decimal(value) {
    if (!['number','string'].includes(typeof value) || typeof value === 'number' && !Number.isFinite(value)) return null;
    let str = text(value).replace(/^[￥¥]\s*/, '').replace(/\s*元$/, '');
    if (str.includes(',') && !/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) return null;
    str = str.replace(/,/g, '');
    const match = /^([+-]?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(str);
    if (!match) return null;
    const exponent = Number(match[4] || 0), scale = (match[3] || '').length - exponent;
    if (!Number.isInteger(exponent) || Math.abs(exponent) > 100) return null;
    const units = BigInt(match[2] + (match[3] || ''));
    let cents;
    if (scale <= 2) cents = units * 10n ** BigInt(2 - scale);
    else {
      const divisor = 10n ** BigInt(scale - 2);
      cents = units / divisor + (units % divisor * 2n >= divisor ? 1n : 0n);
    }
    if (match[1] === '-') cents = -cents;
    return cents > LIMIT || cents < -LIMIT ? null : {cents, precision:scale > 2};
  }

  function monthKey(value, date1904) {
    let year, month, day;
    if (value instanceof Date) {
      if (!Number.isFinite(value.getTime())) return '';
      [year,month,day] = [value.getUTCFullYear(),value.getUTCMonth()+1,value.getUTCDate()];
    } else if (typeof value === 'number') {
      // Small integers also mean bare month numbers; do not invent their year.
      if (!Number.isFinite(value) || value <= 12 || value >= 2958466 || !date1904 && Math.floor(value) === 60) return '';
      const serial = Math.floor(value) + (date1904 ? 1462 : 0);
      const date = new Date((serial < 60 ? Date.UTC(1899,11,31) : Date.UTC(1899,11,30)) + serial * 86400000);
      [year,month,day] = [date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate()];
    } else {
      const match = /^(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})(?:\s*(?:月|[-/.])\s*(\d{1,2}))?(?:月|日)?(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?)?$/.exec(text(value));
      if (!match) return '';
      [year,month,day] = [Number(match[1]),Number(match[2]),Number(match[3] || 1)];
    }
    const date = new Date(0);
    date.setUTCFullYear(year,month-1,day);
    if (year < 1 || year > 9999 || date.getUTCFullYear() !== year || date.getUTCMonth()+1 !== month || date.getUTCDate() !== day) return '';
    return String(year).padStart(4,'0') + '-' + String(month).padStart(2,'0');
  }

  function build(source) {
    const book = source?.book, filename = source?.filename || '来源文件', sheetName = '付款汇总明细';
    const result = {groups:[],totals:{totalCents:0,paidCents:0,remainingCents:0},recordCount:0,warnings:[],errors:[],canExport:false,
      date1904:!!(book?.Workbook?.WBProps?.date1904)};
    const at = coordinate => filename + ' / ' + sheetName + (coordinate ? ' / ' + coordinate : '');
    const sheet = book?.Sheets?.[sheetName];
    if (!sheet) { result.errors.push(at('') + '：没有找到指定工作表'); return result; }
    const rows = new Map();
    for (const [address,cell] of Object.entries(sheet)) {
      const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      if (!match || !hasValue(cell)) continue;
      const row = Number(match[2]);
      if (!rows.has(row)) rows.set(row,[]);
      rows.get(row).push({column:match[1],address,cell});
    }
    const rowNumbers = [...rows.keys()].sort((a,b) => a-b);
    let headerRow, columns;
    for (const row of rowNumbers.slice(0,30)) {
      const found = INPUT_HEADERS.map(label => rows.get(row).filter(entry => header(entry.cell.v) === label));
      if (!found.every(entries => entries.length)) continue;
      for (let i=0;i<found.length;i++) if (found[i].length > 1) result.errors.push(at('第'+row+'行') + '：表头重复“' + INPUT_HEADERS[i] + '”');
      headerRow = row; columns = found.map(entries => entries[0].column); break;
    }
    if (!headerRow) { result.errors.push(at('') + '：未找到完整表头，需要“' + INPUT_HEADERS.join('、') + '”'); return result; }
    if (result.errors.length) return result;
    const groups = new Map(), totals = {total:0n,paid:0n};
    for (const row of rowNumbers.filter(row => row > headerRow)) {
      result.recordCount++;
      const addresses = columns.map(column => column+row), cells = addresses.map(address => sheet[address]);
      let invalid = false;
      cells.forEach((cell,index) => {
        // SheetJS can expose uncached formulas as stub cells with a fabricated zero.
        if (cell?.t === 'e' || cell?.f && (cell.t === 'z' || cell.v == null || text(cell.v) === '')) {
          result.errors.push(at(addresses[index]) + '：“' + INPUT_HEADERS[index] + '”公式无有效结果或包含错误'); invalid = true;
        }
      });
      if (invalid) continue;
      const category = CATEGORIES.get(text(cells[1]?.v)), amount = decimal(cells[2]?.v);
      if (!category) { result.errors.push(at(addresses[1]) + '：分类无法识别“' + text(cells[1]?.v) + '”'); invalid = true; }
      if (!amount) { result.errors.push(at(addresses[2]) + '：申请付款金额为空、无效或超出安全整数分范围'); invalid = true; }
      if (invalid) continue;
      const month = monthKey(cells[0]?.v,result.date1904);
      if (!month) result.warnings.push(at(addresses[0]) + '：付款年月无法确定，已计入“时间待定”；原值“' + text(cells[0]?.v) + '”');
      if (amount.precision) result.warnings.push(at(addresses[2]) + '：金额超过两位小数，统计按分四舍五入');
      const key = JSON.stringify([month,category.category,category.state]);
      if (!groups.has(key)) groups.set(key,{monthKey:month,...category,total:0n,paid:0n});
      const group = groups.get(key), paid = text(cells[3]?.v) === '已支付' ? amount.cents : 0n;
      group.total += amount.cents; group.paid += paid;
      totals.total += amount.cents; totals.paid += paid;
    }
    function finish(value, label) {
      const amounts = [value.total,value.paid,value.total-value.paid];
      if (amounts.some(amount => amount > LIMIT || amount < -LIMIT)) {
        result.errors.push(at('') + '：' + label + '金额超出安全整数分范围');
        return {totalCents:null,paidCents:null,remainingCents:null};
      }
      return {totalCents:Number(amounts[0]),paidCents:Number(amounts[1]),remainingCents:Number(amounts[2])};
    }
    result.groups = [...groups.values()].sort((a,b) => (a.monthKey || '9999-99').localeCompare(b.monthKey || '9999-99') || STATE_ORDER.indexOf(a.state)-STATE_ORDER.indexOf(b.state))
      .map(group => ({monthKey:group.monthKey,category:group.category,state:group.state,...finish(group,(group.monthKey || '时间待定')+' / '+group.category+' / '+group.state)}));
    result.totals = finish(totals,'总计');
    result.canExport = result.errors.length === 0;
    return result;
  }
  return {HEADERS,build};
});
