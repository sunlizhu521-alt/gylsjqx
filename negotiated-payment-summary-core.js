(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NegotiatedPaymentSummaryCore = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const LIMIT = BigInt(Number.MAX_SAFE_INTEGER);
  const text = value => String(value ?? '').trim();
  const header = value => text(value).replace(/\s/g, '');
  const hasValue = cell => cell && (cell.f || cell.t === 'e' || cell.v != null && text(cell.v) !== '');
  const invalidResult = cell => cell?.t === 'e' || cell?.f && (cell.t === 'z' || cell.v == null || text(cell.v) === '');

  function amount(value) {
    if (!['number','string'].includes(typeof value) || typeof value === 'number' && !Number.isFinite(value)) return null;
    let str = text(value).replace(/^[￥¥]\s*/, '').replace(/\s*元$/, '');
    if (str.includes(',') && !/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) return null;
    str = str.replace(/,/g, '');
    const match = /^([+-]?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(str);
    if (!match || str.length > 200) return null;
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

  function build(source, decisions = {}) {
    const sheetName = '谈判付款', filename = source?.filename || '来源文件', sheet = source?.book?.Sheets?.[sheetName];
    const result = {months:[],rows:[],recordCount:0,counts:{negotiated:0,normal:0,unresolved:0},unresolved:[],
      totals:{negotiatedCents:0,normalCents:0,unresolvedCents:0},errors:[],warnings:[],canExport:false};
    const at = address => `${filename} / ${sheetName}${address ? ' / ' + address : ''}`;
    if (!sheet) { result.errors.push(at('') + '：没有找到“谈判付款”工作表'); return result; }
    // The source can contain formatting through XFD; inspect stored values rather than its declared rectangle.
    const rowMap = new Map();
    for (const [address, cell] of Object.entries(sheet)) {
      const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      if (!match || !hasValue(cell)) continue;
      const row = Number(match[2]);
      if (!rowMap.has(row)) rowMap.set(row, []);
      rowMap.get(row).push({column:match[1], address, cell});
    }
    const rowNumbers = [...rowMap.keys()].sort((a,b) => a-b);
    let headerRow, supplierColumn, classColumn, monthColumns;
    for (const row of rowNumbers.slice(0,30)) {
      const entries = rowMap.get(row);
      const suppliers = entries.filter(e => ['供应商','供应商名称'].includes(header(e.cell.v)));
      const classes = entries.filter(e => header(e.cell.v) === '是否谈判分批');
      if (!suppliers.length || !classes.length) continue;
      headerRow = row; supplierColumn = suppliers[0].column; classColumn = classes[0].column; monthColumns = [];
      if (suppliers.length > 1) result.errors.push(at('第'+row+'行') + '：供应商表头重复或存在多个别名列');
      if (classes.length > 1) result.errors.push(at('第'+row+'行') + '：“是否谈判分批”表头重复');
      for (const entry of entries) {
        const label = header(entry.cell.v), match = /^(\d{4})年(\d{1,2})月支付$/.exec(label);
        if (!match) {
          if (/月支付$/.test(label)) result.errors.push(at(entry.address) + '：月度支付表头缺少年份或格式无法识别“' + text(entry.cell.v) + '”');
          continue;
        }
        const year = Number(match[1]), month = Number(match[2]);
        if (year < 1900 || year > 9999 || month < 1 || month > 12) {
          result.errors.push(at(entry.address) + '：月度支付表头日期无效“' + label + '”'); continue;
        }
        if (invalidResult(entry.cell)) result.errors.push(at(entry.address) + '：月份表头公式无有效结果或包含错误');
        const key = match[1] + '-' + String(month).padStart(2,'0');
        if (monthColumns.some(m => m.key === key)) result.errors.push(at(entry.address) + '：月份重复“' + key + '”');
        monthColumns.push({key,label:year+'年'+month+'月',column:entry.column});
      }
      monthColumns.sort((a,b) => a.key.localeCompare(b.key));
      break;
    }
    if (!headerRow) { result.errors.push(at('') + '：找不到表头，需要“供应商／供应商名称”和“是否谈判分批”'); return result; }
    if (!monthColumns.length) result.errors.push(at('第'+headerRow+'行') + '：未找到“2026年9月支付”格式的月度支付列');
    result.months = monthColumns.map(({key,label}) => ({key,label}));
    if (result.errors.length) return result;
    const groups = new Map(), totals = {negotiated:0n,normal:0n,unresolved:0n};
    const checked = (value,label) => {
      if (value > LIMIT || value < -LIMIT) { result.errors.push(at('') + '：' + label + '超出安全整数分范围'); return null; }
      return Number(value);
    };
    for (const row of rowNumbers.filter(r => r > headerRow)) {
      const supplierAddress = supplierColumn+row, classAddress = classColumn+row;
      const supplierCell = sheet[supplierAddress], classCell = sheet[classAddress];
      // Notes or styles outside the business columns do not create an extra payment record.
      if (![supplierCell,classCell,...monthColumns.map(m => sheet[m.column+row])].some(hasValue)) continue;
      result.recordCount++;
      let invalid = false;
      for (const address of [supplierAddress,classAddress,...monthColumns.map(m => m.column+row)]) {
        if (invalidResult(sheet[address])) {
          result.errors.push(at(address) + '：公式无有效结果或包含错误，请在原表核实并重新保存'); invalid = true;
        }
      }
      const supplier = text(supplierCell?.v);
      if (!supplier || !['s','str','inlineStr'].includes(supplierCell?.t)) {
        result.errors.push(at(supplierAddress) + '：供应商名称为空或不是文本，无法汇总'); invalid = true;
      }
      if (supplier && !groups.has(supplier)) groups.set(supplier,{supplier,negotiated:Array(monthColumns.length).fill(null),normal:Array(monthColumns.length).fill(null)});
      const amounts = monthColumns.map(m => {
        const address = m.column+row, cell = sheet[address];
        if (invalidResult(cell)) return null;
        if (!hasValue(cell)) return null;
        const parsed = amount(cell.v);
        if (!parsed) { result.errors.push(at(address) + '：月度金额无效或超出安全整数分范围“' + text(cell.v) + '”'); invalid = true; return null; }
        if (parsed.precision) result.warnings.push(at(address) + '：金额超过两位小数，按分四舍五入');
        return parsed.cents;
      });
      const originalClass = text(classCell?.v);
      const override = Object.prototype.hasOwnProperty.call(decisions || {},classAddress) ? decisions[classAddress] : '';
      const classification = ['是','否'].includes(originalClass) ? originalClass : ['是','否'].includes(override) ? override : '';
      if (!classification) {
        result.counts.unresolved++;
        const rowTotal = amounts.reduce((sum,v) => sum+(v ?? 0n),0n);
        result.unresolved.push({id:classAddress,address:classAddress,row,supplier,classification:originalClass,
          amountCents:invalid ? null : checked(rowTotal,`第${row}行待分类金额`),source:at('第'+row+'行')});
        if (!invalid) totals.unresolved += rowTotal;
        continue;
      }
      const type = classification === '是' ? 'negotiated' : 'normal';
      result.counts[type]++;
      if (invalid) continue;
      const group = groups.get(supplier);
      amounts.forEach((value,index) => {
        if (value === null) return;
        group[type][index] = (group[type][index] ?? 0n) + value;
        totals[type] += value;
      });
    }
    result.rows = [...groups.values()].map(group => {
      const cells = type => group[type].map((v,index) => v === null ? null : checked(v,group.supplier+' / '+result.months[index].label+'金额'));
      const total = type => checked(group[type].reduce((sum,v) => sum+(v ?? 0n),0n),group.supplier+' / '+(type === 'negotiated' ? '谈判' : '正常')+'总金额');
      return {supplier:group.supplier,negotiatedCents:cells('negotiated'),normalCents:cells('normal'),
        negotiatedTotalCents:total('negotiated'),normalTotalCents:total('normal')};
    });
    for (const type of ['negotiated','normal','unresolved']) result.totals[type+'Cents'] = checked(totals[type],type === 'unresolved' ? '待分类合计' : type === 'negotiated' ? '谈判总金额' : '正常付款总金额');
    if (!result.recordCount) result.errors.push(at('') + '：没有可整理的付款记录');
    result.canExport = result.errors.length === 0 && result.unresolved.length === 0;
    return result;
  }
  return {build};
});
