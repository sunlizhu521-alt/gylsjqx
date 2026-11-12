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

  function day(year, month, date) {
    const value = new Date(Date.UTC(year, month - 1, date));
    if (year < 1900 || year > 9999 || value.getUTCFullYear() !== year || value.getUTCMonth() + 1 !== month || value.getUTCDate() !== date) return null;
    return {key:String(year).padStart(4,'0')+'-'+String(month).padStart(2,'0')+'-'+String(date).padStart(2,'0'),label:year+'年'+month+'月'+date+'日'};
  }
  function headerDay(cell, date1904) {
    if (!cell || invalidResult(cell)) return null;
    const value = cell.v;
    if (typeof value === 'number') {
      const format = text(cell.z).replace(/"[^"]*"|\\./g,'').replace(/\[[^\]]*\]/g,'');
      if (!/[yd]/i.test(format)) {
        const displayed = /^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?$/.exec(header(cell.w));
        return displayed ? day(+displayed[1],+displayed[2],+displayed[3]) : null;
      }
      if (!Number.isFinite(value) || value < (date1904 ? 0 : 1) || value >= 2958466 || !date1904 && Math.floor(value) === 60) return null;
      const base = date1904 ? Date.UTC(1904,0,1) : Date.UTC(1899,11,value < 60 ? 31 : 30);
      const d = new Date(base + Math.floor(value)*86400000);
      return day(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate());
    }
    if (Object.prototype.toString.call(value) === '[object Date]' && Number.isFinite(value.getTime())) return day(value.getUTCFullYear(),value.getUTCMonth()+1,value.getUTCDate());
    const m = /^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?$/.exec(header(value));
    return m ? day(+m[1],+m[2],+m[3]) : null;
  }
  function columnIndex(column) {
    return [...column].reduce((sum,char) => sum*26+char.charCodeAt(0)-64,0)-1;
  }
  function paymentComment(cell, cents) {
    const lines = [...new Set((Array.isArray(cell?.c) ? cell.c : []).flatMap(c => text(c.t).normalize('NFKC').split(/[\r\n]+/)).map(text).filter(Boolean))];
    let confirmed = false, unclear = false, relevant = false;
    for (const line of lines) {
      const normalized = line.replace(/\s/g,'');
      if (/^申请日期\((?:付款节奏跟进|资金整理)\):/.test(normalized)) continue;
      if (!/支付|付款|付清|到账|退汇|撤回|退款|退回|冲正|作废/.test(normalized)) continue;
      relevant = true;
      const value = normalized.replace(/^(?:(?:\d{4}[年/.-])?\d{1,2}[月/.-]\d{1,2}日?)[-—:：,，]*/,'').replace(/^(?:经)?(?:财务|出纳)?(?:核实|确认)[-—:：,，]*/,'');
      if (/未支付|未付款|未付清|未到账|支付失败|付款失败|退汇|撤回|退款|退回|冲正|作废|部分(?:已)?(?:支付|付款)/.test(value)) { unclear = true; continue; }
      const match = /^(?:本批|本次)?(?:已支付|已付款|已付清)(?:完成|成功)?(?:[:：]?(?:人民币|[￥¥])?(-?[\d,]+(?:\.\d+)?)元?)?[。.!！]*$/.exec(value);
      const paidAmount = match?.[1] ? amount(match[1]) : null;
      if (!match || match[1] && (!paidAmount || paidAmount.precision || paidAmount.cents !== cents)) unclear = true;
      else confirmed = true;
    }
    return {paid:confirmed && !unclear,relevant};
  }

  function build(source, decisions = {}) {
    const sheetName = '谈判付款', filename = source?.filename || '来源文件', sheet = source?.book?.Sheets?.[sheetName];
    const result = {dates:[],rows:[],paid:{negotiated:[],normal:[]},recordCount:0,counts:{negotiated:0,normal:0,unresolved:0},unresolved:[],
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
    let headerRow, supplierColumn, classColumn, detailColumn;
    for (const row of rowNumbers.slice(0,30)) {
      const entries = rowMap.get(row);
      const suppliers = entries.filter(e => ['供应商','供应商名称'].includes(header(e.cell.v)));
      const classes = entries.filter(e => header(e.cell.v) === '是否谈判分批');
      const details = entries.filter(e => header(e.cell.v) === '谈判明细');
      if (!suppliers.length || !classes.length || !details.length) continue;
      headerRow = row; supplierColumn = suppliers[0].column; classColumn = classes[0].column; detailColumn = details[0].column;
      for (const [label,list] of [['供应商',suppliers],['是否谈判分批',classes],['谈判明细',details]]) {
        if (list.length > 1) result.errors.push(at('第'+row+'行') + '：“'+label+'”表头重复或存在多个别名列');
        for (const entry of list) if (invalidResult(entry.cell)) result.errors.push(at(entry.address) + '：必要表头公式无有效结果或包含错误');
      }
      break;
    }
    if (!headerRow) { result.errors.push(at('') + '：找不到表头，需要“供应商／供应商名称”“是否谈判分批”和“谈判明细”'); return result; }
    if (result.errors.length) return result;
    const monthlyYears = new Map(), datedYears = new Map(), scheduleColumns = [];
    const remember = (map,year,month,address) => {
      if (!map.has(month)) map.set(month,new Map());
      const years = map.get(month);
      if (!years.has(year)) years.set(year,[]);
      years.get(year).push(address);
    };
    for (const entry of rowMap.get(headerRow)) {
      if (invalidResult(entry.cell)) continue;
      const m = /^(\d{4})年(\d{1,2})月支付$/.exec(header(entry.cell.v));
      if (m && day(+m[1],+m[2],1)) remember(monthlyYears,+m[1],+m[2],entry.address);
      const d = headerDay(entry.cell,!!source.book.Workbook?.WBProps?.date1904);
      if (d) {
        remember(datedYears,+d.key.slice(0,4),+d.key.slice(5,7),entry.address);
        scheduleColumns.push({column:entry.column,date:d.key});
      }
    }
    const parseDetail = (value,address) => {
      if (!text(value)) { result.errors.push(at(address) + '：谈判明细为空，请补充具体付款日期和金额'); return null; }
      if (typeof value !== 'string') { result.errors.push(at(address) + '：谈判明细不是日期和金额文本，请核实原值'); return null; }
      const parts = [], before = result.errors.length;
      const lines = value.replace(/\s+(?=\d+\s*[、.)）．]\s*(?:(?:\d{4}年)?\d{1,2}月|\d{4}[-/.]))/g,'\n').split(/[\r\n;；]+/).map(text).filter(Boolean);
      for (let index=0;index<lines.length;index++) {
        const line = lines[index].replace(/^\d+\s*[、.)）．]\s*(?=(?:\d{4}年)?\d{1,2}月|\d{4}[-/.])/,'');
        const location = at(address)+' 第'+(index+1)+'笔';
        let m = /^(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日?\s*[－—–:：-]\s*(.+)$/.exec(line);
        if (!m) m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*[－—–:：-]\s*(.+)$/.exec(line);
        if (!m) { result.errors.push(location+'：无法完整解析“'+lines[index]+'”，请填写“年月日－金额”'); continue; }
        let year = m[1] ? +m[1] : null;
        const month = +m[2], date = +m[3];
        if (year === null) {
          const years = monthlyYears.get(month) || datedYears.get(month);
          if (!years || years.size !== 1) {
            const references = years ? [...years].map(([y,addresses]) => y+'年（'+addresses.join('、')+'）').join('、') : '无同月年月表头或完整日期表头';
            result.errors.push(location+'：日期“'+month+'月'+date+'日”缺少年份，'+(years?'同月年份不唯一：':'没有唯一年份依据：')+references+'；请补充完整年月日'); continue;
          }
          year = years.keys().next().value;
        }
        const parsedDate = day(year,month,date);
        if (!parsedDate) { result.errors.push(location+'：日期无效“'+year+'年'+month+'月'+date+'日”'); continue; }
        let rawAmount = text(m[4]);
        const annotation = /[（(]([\p{L}\s，,。！!、；;：:]+)[）)]$/u.exec(rawAmount);
        if (annotation) {
          rawAmount = rawAmount.slice(0,annotation.index).trim();
          result.warnings.push(location+'：括号说明“'+text(annotation[1])+'”不参与金额汇总');
        }
        const parsed = amount(rawAmount);
        if (!parsed) { result.errors.push(location+'：金额无效、超出安全整数分范围或含未明确的附加内容“'+text(m[4])+'”，请规范为单笔金额'); continue; }
        if (parsed.precision) result.warnings.push(location+'：金额超过两位小数，按分四舍五入');
        parts.push({...parsedDate,cents:parsed.cents});
      }
      if (!lines.length) result.errors.push(at(address)+'：谈判明细没有可解析的具体付款日期和金额');
      return result.errors.length === before ? parts : null;
    };
    const recordRows = rowNumbers.filter(r => r > headerRow && [supplierColumn,classColumn,detailColumn].some(col => hasValue(sheet[col+r])));
    const mergedRecords = new Set(), detailIndex = columnIndex(detailColumn);
    for (const merge of sheet['!merges'] || []) {
      if (merge.s.c > detailIndex || merge.e.c < detailIndex || merge.s.r === merge.e.r) continue;
      const affected = recordRows.filter(row => row >= merge.s.r+1 && row <= merge.e.r+1);
      if (affected.length > 1) for (const row of affected) mergedRecords.add(row);
    }
    const paidParts = (row, parts) => {
      const paid = parts.map(() => false), batches = [], warned = new Set();
      const warn = (address,reason) => {
        if (warned.has(address)) return;
        warned.add(address);
        result.warnings.push(at(address)+'：支付批注未用于标红，'+reason+'；请核对本行谈判明细的日期、金额及对应支付批注');
      };
      for (const item of scheduleColumns) {
        const address = item.column+row, cell = sheet[address];
        const parsed = invalidResult(cell) ? null : amount(cell?.v), payment = paymentComment(cell,parsed?.cents);
        if (!parsed) {
          if (payment.relevant) warn(address,'批次金额为空、异常或公式结果不可用');
          continue;
        }
        const column = columnIndex(item.column);
        const merged = (sheet['!merges'] || []).some(m => m.s.c <= column && m.e.c >= column && m.s.r < m.e.r && row >= m.s.r+1 && row <= m.e.r+1 && recordRows.filter(r => r >= m.s.r+1 && r <= m.e.r+1).length > 1);
        if (merged && payment.relevant) warn(address,'金额单元格跨多条记录合并，不能据此确认本批已支付');
        else if (payment.relevant && !payment.paid) warn(address,'批注未明确全额已支付，或含疑问、否定、部分支付、相互矛盾或支付金额不一致');
        batches.push({...item,address,cents:parsed.cents,payment:{...payment,paid:payment.paid && !merged},used:false});
      }
      const matched = new Set();
      const match = (indices,candidates) => {
        indices.forEach(index => { paid[index] = candidates.every(batch => batch.payment.paid); matched.add(index); });
        candidates.forEach(batch => { batch.used = true; });
      };
      const exactKeys = new Set(parts.map(part => part.key+'|'+part.cents));
      for (const key of exactKeys) {
        const indices = parts.map((part,index) => part.key+'|'+part.cents === key ? index : -1).filter(index => index !== -1);
        const candidates = batches.filter(batch => batch.date+'|'+batch.cents === key);
        if (indices.length === 1 && candidates.length === 1) match(indices,candidates);
      }
      const remainingKeys = new Set(parts.filter((_,index) => !matched.has(index)).map(part => part.key.slice(0,7)+'|'+part.cents));
      for (const key of remainingKeys) {
        const indices = parts.map((part,index) => !matched.has(index) && part.key.slice(0,7)+'|'+part.cents === key ? index : -1).filter(index => index !== -1);
        const candidates = batches.filter(batch => !batch.used && batch.date.slice(0,7)+'|'+batch.cents === key);
        if (indices.length === 1 && candidates.length === 1 || indices.length === candidates.length && candidates.every(batch => batch.payment.paid)) match(indices,candidates);
        else for (const batch of candidates) if (batch.payment.relevant) warn(batch.address,'同月同金额的多批归属不唯一，不能将一个支付批注用于其他批次');
      }
      for (const batch of batches) if (!batch.used && batch.payment.relevant) warn(batch.address,'未找到同源行、同年月及同金额的唯一谈判明细，不跨月配对');
      return paid;
    };
    const groups = new Map(), dates = new Map(), totals = {negotiated:0n,normal:0n,unresolved:0n};
    const checked = (value,label) => {
      if (value > LIMIT || value < -LIMIT) { result.errors.push(at('') + '：' + label + '超出安全整数分范围'); return null; }
      return Number(value);
    };
    for (const row of recordRows) {
      const supplierAddress = supplierColumn+row, classAddress = classColumn+row, detailAddress = detailColumn+row;
      const supplierCell = sheet[supplierAddress], classCell = sheet[classAddress], detailCell = sheet[detailAddress];
      result.recordCount++;
      let invalid = false;
      for (const address of [supplierAddress,classAddress,detailAddress]) {
        if (invalidResult(sheet[address])) { result.errors.push(at(address) + '：公式无有效结果或包含错误，请在原表核实并重新保存'); invalid = true; }
      }
      const supplier = text(supplierCell?.v);
      if (!supplier || !['s','str','inlineStr'].includes(supplierCell?.t)) { result.errors.push(at(supplierAddress) + '：供应商名称为空或不是文本，无法汇总'); invalid = true; }
      if (supplier && !groups.has(supplier)) groups.set(supplier,{supplier,negotiated:new Map(),normal:new Map(),paid:{negotiated:new Map(),normal:new Map()}});
      let parts = null;
      if (mergedRecords.has(row)) { result.errors.push(at(detailAddress)+'：谈判明细纵向合并跨多条记录，不能重复分配金额；请明确各记录的日期和金额'); invalid = true; }
      else if (!invalidResult(detailCell)) parts = parseDetail(detailCell?.v,detailAddress);
      if (!parts) invalid = true;
      if (parts) for (const part of parts) dates.set(part.key,{key:part.key,label:part.label});
      const originalClass = text(classCell?.v);
      const override = Object.prototype.hasOwnProperty.call(decisions || {},classAddress) ? decisions[classAddress] : '';
      const classification = ['是','否'].includes(originalClass) ? originalClass : ['是','否'].includes(override) ? override : '';
      if (!classification) {
        result.counts.unresolved++;
        const rowTotal = (parts || []).reduce((sum,v) => sum+v.cents,0n);
        result.unresolved.push({id:classAddress,address:classAddress,row,supplier,classification:originalClass,
          amountCents:invalid ? null : checked(rowTotal,`第${row}行待分类金额`),source:at('第'+row+'行')});
        if (!invalid) totals.unresolved += rowTotal;
        continue;
      }
      const type = classification === '是' ? 'negotiated' : 'normal';
      result.counts[type]++;
      if (invalid) continue;
      const group = groups.get(supplier);
      const confirmed = paidParts(row,parts);
      for (const [index,part] of parts.entries()) {
        group[type].set(part.key,(group[type].get(part.key) ?? 0n)+part.cents);
        group.paid[type].set(part.key,(group.paid[type].get(part.key) ?? true) && confirmed[index]);
        totals[type] += part.cents;
      }
    }
    result.dates = [...dates.values()].sort((a,b) => a.key.localeCompare(b.key));
    result.rows = [...groups.values()].map(group => {
      const cells = type => result.dates.map(date => group[type].has(date.key) ? checked(group[type].get(date.key),group.supplier+' / '+date.label+'金额') : null);
      const total = type => checked([...group[type].values()].reduce((sum,v) => sum+v,0n),group.supplier+' / '+(type === 'negotiated' ? '谈判' : '正常')+'总金额');
      return {supplier:group.supplier,negotiatedCents:cells('negotiated'),normalCents:cells('normal'),
        negotiatedTotalCents:total('negotiated'),normalTotalCents:total('normal')};
    });
    for (const type of ['negotiated','normal']) result.paid[type] = [...groups.values()].map(group => result.dates.map(date => group.paid[type].get(date.key) === true));
    for (const type of ['negotiated','normal','unresolved']) result.totals[type+'Cents'] = checked(totals[type],type === 'unresolved' ? '待分类合计' : type === 'negotiated' ? '谈判总金额' : '正常付款总金额');
    if (!result.recordCount) result.errors.push(at('') + '：没有可整理的付款记录');
    result.canExport = result.errors.length === 0 && result.unresolved.length === 0;
    return result;
  }
  function dateColumns(result) {
    const columns = {};
    for (const type of ['negotiated','normal']) {
      columns[type] = result.dates.map((_,index) => index).filter(index =>
        result.rows.some(row => Number.isFinite(row[type+'Cents']?.[index])));
    }
    return columns;
  }
  return {build,dateColumns};
});
