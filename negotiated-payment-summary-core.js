(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NegotiatedPaymentSummaryCore = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const LIMIT = BigInt(Number.MAX_SAFE_INTEGER);
  const VALIDATION_HEADERS = ['来源文件','来源工作表','来源行号','供应商','付款主体','对公付款OA流程号','是否谈判分批','OA-申请付款金额','谈判明细金额合计','Q-AB列金额合计','AI-AP列金额合计','谈判明细差额','Q-AB差额','AI-AP差额','校验结果','需要处理'];
  const VALIDATION_MONEY_COLUMNS = [7,8,9,10,11,12,13];
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
  function columnName(index) {
    let name = '';
    for (let n=index+1;n>0;n=Math.floor((n-1)/26)) name=String.fromCharCode(65+(n-1)%26)+name;
    return name;
  }

  function build(source, decisions = {}) {
    const sheetName = '谈判付款', filename = source?.filename || '来源文件', sheet = source?.book?.Sheets?.[sheetName];
    const result = {dates:[],rows:[],validationRows:[],validationCounts:{consistent:0,inconsistent:0,unverified:0},recordCount:0,counts:{negotiated:0,normal:0,unresolved:0},unresolved:[],
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
    const monthlyYears = new Map(), datedYears = new Map();
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
    const fieldEntries = names => rowMap.get(headerRow).filter(e => names.includes(header(e.cell.v).replace(/0A/gi,'OA')));
    const validationFields = {
      amount:fieldEntries(['OA-申请付款金额','申请付款金额']),
      subject:fieldEntries(['付款主体']),
      oa:fieldEntries(['对公付款OA流程号','对公付款OA号','付款OA号'])
    };
    const ranges = [{start:16,end:27,label:'Q-AB'},{start:34,end:41,label:'AI-AP'}];
    for (const range of ranges) {
      range.headerIssues = [];
      for (let c=range.start;c<=range.end;c++) {
        const address=columnName(c)+headerRow, cell=sheet[address];
        if (!hasValue(cell)) continue;
        const monthly=/^(\d{4})年(\d{1,2})月支付$/.exec(header(cell.v));
        if (invalidResult(cell) || !(monthly && day(+monthly[1],+monthly[2],1)) && !headerDay(cell,!!source.book.Workbook?.WBProps?.date1904)) {
          range.headerIssues.push(address+'表头“'+text(cell.v)+'”不是付款日期或年月支付；请核实'+range.label+'固定列范围是否移位');
        }
      }
    }
    const ambiguousMerge = (column,row) => (sheet['!merges'] || []).find(merge => {
      if (column<merge.s.c || column>merge.e.c || row<merge.s.r+1 || row>merge.e.r+1) return false;
      const anchor=sheet[columnName(merge.s.c)+(merge.s.r+1)];
      return hasValue(anchor) && (merge.s.c!==merge.e.c || recordRows.filter(r => r>=merge.s.r+1 && r<=merge.e.r+1).length>1);
    });
    const readMoney = (column,row,issues) => {
      const address=columnName(column)+row, cell=sheet[address], merge=ambiguousMerge(column,row);
      if (merge) { issues.push(address+'属于合并单元格'+columnName(merge.s.c)+(merge.s.r+1)+':'+columnName(merge.e.c)+(merge.e.r+1)+'，金额归属不清；请明确每条记录、每个日期的金额'); return null; }
      if (invalidResult(cell)) { issues.push(address+'公式结果错误或缺失；请核实公式、重新计算并保存源表'); return null; }
      if (!hasValue(cell)) return null;
      const parsed=amount(cell.v);
      if (!parsed) { issues.push(address+'金额“'+text(cell.v)+'”无效或超出安全整数分范围；请核实并填写有效金额'); return null; }
      return parsed.cents;
    };
    const validationNumber = (value,label,issues) => {
      if (value===null) return null;
      if (value>LIMIT || value< -LIMIT) { issues.push(label+'超出安全整数分范围，未输出近似金额；请核实该金额'); return null; }
      return Number(value);
    };
    const validateRow = (row,supplier,classification,parts) => {
      const issues=[], values=[];
      const amountEntries=validationFields.amount;
      if (amountEntries.length!==1) {
        values.push(null); issues.push(amountEntries.length ? 'OA申请金额表头重复（'+amountEntries.map(e=>e.address).join('、')+'）；请保留唯一金额来源列' : '未找到OA-申请付款金额／申请付款金额表头；请补充对应列');
      } else {
        const entry=amountEntries[0], before=issues.length;
        if (invalidResult(entry.cell)) { issues.push(entry.address+'金额表头公式无有效结果；请核实表头'); values.push(null); }
        else {
          const value=readMoney(columnIndex(entry.column),row,issues); values.push(value);
          if (value===null && before===issues.length) issues.push(entry.column+row+'申请付款金额为空；请补充金额，空白不按零处理');
        }
      }
      values.push(parts ? parts.reduce((sum,part)=>sum+part.cents,0n) : null);
      if (!parts) issues.push(detailColumn+row+'谈判明细金额无法完整核验；请处理该单元格的日期、金额或公式问题');
      for (const range of ranges) {
        const before=issues.length; issues.push(...range.headerIssues);
        let total=0n, populated=false;
        for (let c=range.start;c<=range.end;c++) {
          const column=columnName(c), value=readMoney(c,row,issues);
          if (value!==null) {
            populated=true; total+=value;
            if (!hasValue(sheet[column+headerRow])) issues.push(column+row+'有金额但'+column+headerRow+'表头为空；请核实'+range.label+'固定列范围');
          }
        }
        values.push(issues.length===before && populated ? total : null);
        if (issues.length===before && !populated) issues.push(range.label+'第'+row+'行金额全空；请补充分批／月度金额，空白不按零处理');
      }
      const numbers=values.map((value,i)=>validationNumber(value,VALIDATION_HEADERS[i+7]+'（第'+row+'行）',issues));
      const differences=values.slice(1).map((value,i)=>value===null || values[0]===null ? null : validationNumber(value-values[0],VALIDATION_HEADERS[i+11]+'（第'+row+'行）',issues));
      const incomplete=numbers.includes(null) || differences.includes(null), mismatch=differences.some(value=>value!==null && value!==0);
      if (mismatch) issues.push('请核对本行OA申请金额与谈判明细、Q-AB分批金额、AI-AP月度金额；差额为对应组金额减OA申请金额，确认并修正不一致的来源值');
      const status=incomplete ? '待核对' : mismatch ? '不一致' : '一致';
      result.validationCounts[incomplete ? 'unverified' : mismatch ? 'inconsistent' : 'consistent']++;
      const sourceField = entries => entries.length===1 && !invalidResult(sheet[entries[0].column+row]) ? text(sheet[entries[0].column+row]?.v) : '';
      result.validationRows.push([filename,sheetName,row,supplier,sourceField(validationFields.subject),sourceField(validationFields.oa),classification,...numbers,...differences,status,[...new Set(issues)].join('\n')]);
    };
    const finishValidation = () => {
      if (!result.validationRows.length) return;
      const total=Array(VALIDATION_HEADERS.length).fill(''), issues=[];
      total[0]=filename; total[1]=sheetName; total[3]='合计';
      for (const c of VALIDATION_MONEY_COLUMNS) {
        if (result.validationRows.some(row=>row[c]===null)) { total[c]=null; issues.push(VALIDATION_HEADERS[c]+'存在未核验记录，合计留空，不以部分金额代替总额'); }
        else total[c]=validationNumber(result.validationRows.reduce((sum,row)=>sum+BigInt(row[c]),0n),VALIDATION_HEADERS[c]+'合计',issues);
      }
      const counts=result.validationCounts;
      total[14]=issues.length || counts.unverified ? '待核对' : counts.inconsistent ? '不一致' : '一致';
      total[15]='逐行结果：一致'+counts.consistent+'条，不一致'+counts.inconsistent+'条，待核对'+counts.unverified+'条。'+(counts.inconsistent || counts.unverified ? '请按各来源行处理，合计抵消不代表逐行一致。' : '')+issues.join('\n');
      result.validationRows.push(total);
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
      if (supplier && !groups.has(supplier)) groups.set(supplier,{supplier,negotiated:new Map(),normal:new Map()});
      let parts = null;
      if (mergedRecords.has(row)) { result.errors.push(at(detailAddress)+'：谈判明细纵向合并跨多条记录，不能重复分配金额；请明确各记录的日期和金额'); invalid = true; }
      else if (!invalidResult(detailCell)) parts = parseDetail(detailCell?.v,detailAddress);
      if (!parts) invalid = true;
      if (parts) for (const part of parts) dates.set(part.key,{key:part.key,label:part.label});
      const originalClass = text(classCell?.v);
      const override = Object.prototype.hasOwnProperty.call(decisions || {},classAddress) ? decisions[classAddress] : '';
      const classification = ['是','否'].includes(originalClass) ? originalClass : ['是','否'].includes(override) ? override : '';
      validateRow(row,supplier,classification || originalClass,parts);
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
      for (const part of parts) {
        group[type].set(part.key,(group[type].get(part.key) ?? 0n)+part.cents);
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
    for (const type of ['negotiated','normal','unresolved']) result.totals[type+'Cents'] = checked(totals[type],type === 'unresolved' ? '待分类合计' : type === 'negotiated' ? '谈判总金额' : '正常付款总金额');
    finishValidation();
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
  return {build,dateColumns,VALIDATION_HEADERS,VALIDATION_MONEY_COLUMNS};
});
