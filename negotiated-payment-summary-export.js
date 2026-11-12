(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'), require('./negotiated-payment-summary-core.js'));
  else root.NegotiatedPaymentSummaryExport = factory(root.XLSX, root.JSZip, root.NegotiatedPaymentSummaryCore);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip, Core) {
  'use strict';
  const AMOUNT_FORMAT = '#,##0.00';
  const width = text => [...String(text)].reduce((n, c) => n + (c.codePointAt(0) > 255 ? 2 : 1), 0);
  const cell = value => ({t:'s',v:value == null ? '' : String(value)});

  function amount(cents) {
    if (cents === null) return cell('');
    if (!Number.isSafeInteger(cents)) throw Error('金额不是有效的整数分，请重新整理。');
    const value = cents / 100;
    if (Math.round(value * 100) !== cents) throw Error('金额超过可精确保存到分的范围，请核对金额。');
    return {t:'n',v:value,z:AMOUNT_FORMAT};
  }
  function dateCell(key) {
    const match = /^(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.exec(key);
    if (!match || Number(match[1]) < 1900) throw Error('付款日期无效，请重新整理。');
    const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]), date = new Date(Date.UTC(year,month-1,day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month-1 || date.getUTCDate() !== day) throw Error('付款日期无效，请重新整理。');
    return {t:'s',v:`${year}年${month}月${day}日`,z:'@'};
  }
  function checkAmounts(values, total, count) {
    if (!Array.isArray(values) || values.length !== count || !Number.isSafeInteger(total)) throw Error('日期金额或合计不完整，请重新整理。');
    let sum = 0n;
    for (const value of values) {
      if (value === null) continue;
      if (!Number.isSafeInteger(value)) throw Error('金额不是有效的整数分，请重新整理。');
      sum += BigInt(value);
    }
    if (sum !== BigInt(total)) throw Error('日期金额与合计不一致，请重新整理。');
  }

  return async function(result) {
    if (!result?.canExport || result.errors?.length || result.unresolved?.length) throw Error('请先完成分类核对并修正金额问题，再下载。');
    if (!Array.isArray(result.dates) || !Array.isArray(result.rows) || !result.rows.length) throw Error('缺少谈判付款整理结果。');
    const keys = result.dates.map(item => item?.key);
    if (keys.some((key, index) => typeof key !== 'string' || index > 0 && key <= keys[index-1])) throw Error('付款日期重复或未按日期升序排列。');
    const dates = keys.map(dateCell);
    for (const entry of result.rows) {
      if (typeof entry.supplier !== 'string' || !entry.supplier.trim()) throw Error('供应商名称为空，请重新整理。');
      checkAmounts(entry.negotiatedCents,entry.negotiatedTotalCents,dates.length);
      checkAmounts(entry.normalCents,entry.normalTotalCents,dates.length);
    }
    const selected = Core.dateColumns(result), negotiatedDates = selected.negotiated, normalDates = selected.normal;
    const normalTotal = negotiatedDates.length + 2, normalStart = normalTotal + 1, columns = normalStart + normalDates.length;
    if (columns > 16384 || result.rows.length + 2 > 1048576) throw Error('结果超过 Excel 的行列数量限制。');
    const top = Array.from({length:columns},() => cell(''));
    const second = Array.from({length:columns},() => cell(''));
    top[0] = cell('供应商'); top[1] = cell('谈判总金额'); top[normalTotal] = cell('正常付款总金额');
    if (negotiatedDates.length) top[2] = cell('谈判付款节奏');
    if (normalDates.length) top[normalStart] = cell('正常付款');
    negotiatedDates.forEach((source,index) => { second[index+2] = {...dates[source]}; });
    normalDates.forEach((source,index) => { second[index+normalStart] = {...dates[source]}; });
    const rows = [top,second], paidCells = new Set();
    result.rows.forEach((entry,index) => {
      const row = index + 3;
      const negotiated = {...amount(entry.negotiatedTotalCents),f:negotiatedDates.length?`SUM(C${row}:${X.utils.encode_col(normalTotal-1)}${row})`:'SUM(0)'};
      const normal = {...amount(entry.normalTotalCents),f:normalDates.length?`SUM(${X.utils.encode_col(normalStart)}${row}:${X.utils.encode_col(columns-1)}${row})`:'SUM(0)'};
      rows.push([cell(entry.supplier),negotiated,...negotiatedDates.map(source=>amount(entry.negotiatedCents[source])),normal,...normalDates.map(source=>amount(entry.normalCents[source]))]);
      negotiatedDates.forEach((source,c) => { if (entry.negotiatedCents[source] !== null && result.paid?.negotiated?.[index]?.[source] === true) paidCells.add(X.utils.encode_col(c+2)+row); });
      normalDates.forEach((source,c) => { if (entry.normalCents[source] !== null && result.paid?.normal?.[index]?.[source] === true) paidCells.add(X.utils.encode_col(c+normalStart)+row); });
    });
    const widths = Array.from({length:columns},() => 12);
    const heights = [];
    rows.forEach((row,r) => {
      let lines = 1;
      row.forEach((value,c) => {
        if (r === 0 && (c === 2 && negotiatedDates.length > 1 || c === normalStart && normalDates.length > 1)) return;
        const text = value.z ? X.SSF.format(value.z,value.v) : value.v;
        const parts = String(text).split(/\r\n|\r|\n/);
        lines = Math.max(lines,parts.length);
        for (const part of parts) widths[c] = Math.max(widths[c],width(part)+4);
      });
      const height = Math.max(26,lines*16+10);
      if (height > 409) throw Error(`第${r+1}行内容超过 Excel 行高限制，请核对供应商名称。`);
      heights.push({hpt:height});
    });
    if (widths.some(value => value > 255)) throw Error('内容超过 Excel 列宽限制，请核对供应商名称或金额。');
    const sheet = X.utils.aoa_to_sheet(rows), book = X.utils.book_new();
    const activeRows = new Set(), activeColumns = new Set();
    rows.forEach((row,r) => row.forEach((item,c) => {
      if (item.v != null && (typeof item.v !== 'string' || item.v.trim())) { activeRows.add(r); activeColumns.add(c); }
    }));
    sheet['!cols'] = widths.map(wch => ({wch})); sheet['!rows'] = heights;
    sheet['!autofilter'] = {ref:X.utils.encode_range({s:{r:1,c:0},e:{r:rows.length-1,c:columns-1}})};
    sheet['!merges'] = [0,1,normalTotal].map(c => ({s:{r:0,c},e:{r:1,c}}));
    if (negotiatedDates.length > 1) sheet['!merges'].push({s:{r:0,c:2},e:{r:0,c:normalTotal-1}});
    if (normalDates.length > 1) sheet['!merges'].push({s:{r:0,c:normalStart},e:{r:0,c:columns-1}});
    X.utils.book_append_sheet(book,sheet,'谈判付款整理');
    const zip = await Zip.loadAsync(X.write(book,{type:'array',bookType:'xlsx'}));
    let styles = await zip.file('xl/styles.xml').async('string');
    const original = styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
    const formats = original.map(xf => xf.match(/\bnumFmtId="(\d+)"/)?.[1] || '0'), styleCount = formats.length;
    const fonts = '<fonts count="2"><font><sz val="11"/><color rgb="FF172033"/><name val="宋体"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="宋体"/></font></fonts>';
    const fills = '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315B9D"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></fills>';
    const borders = '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' + ['thin','hair'].map(style => '<border>' + ['left','right','top','bottom'].map(side => `<${side} style="${style}"><color rgb="FF404040"/></${side}>`).join('') + '<diagonal/></border>').join('') + '</borders>';
    const xf = (format,header,unframed=false,paid=false) => `<xf numFmtId="${format}" fontId="${header?1:0}" fillId="${header?3:paid?4:2}" borderId="${header?1:unframed?0:2}" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="0" shrinkToFit="0"/></xf>`;
    styles = styles.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/,fonts).replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/,fills).replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/,borders)
      .replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,`<cellXfs count="${styleCount*4}">${formats.map(format => xf(format,false)).join('')}${formats.map(format => xf(format,true)).join('')}${formats.map(format => xf(format,false,true)).join('')}${formats.map(format => xf(format,false,false,true)).join('')}</cellXfs>`);
    zip.file('xl/styles.xml',styles);
    let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    xml = xml.replace(/<c\b([^>]*?)(\/?)>/g,(tag,attrs,end) => {
      const address = attrs.match(/\br="([A-Z]+(\d+))"/);
      if (!address) return tag;
      const point = X.utils.decode_cell(address[1]), framed = activeRows.has(point.r) && activeColumns.has(point.c);
      const style = Number(attrs.match(/\bs="(\d+)"/)?.[1] || 0) + (Number(address[2]) <= 2 ? styleCount : paidCells.has(address[1]) ? styleCount*3 : framed ? 0 : styleCount*2);
      return '<c' + attrs.replace(/\s+s="\d+"/,'') + ` s="${style}"${end}>`;
    });
    const views = '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/></sheetView></sheetViews>';
    xml = /<sheetViews\b/.test(xml) ? xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/,views) : xml.replace(/<(?:sheetFormatPr|cols|sheetData)\b/,match => views+match);
    zip.file('xl/worksheets/sheet1.xml',xml);
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  };
});
