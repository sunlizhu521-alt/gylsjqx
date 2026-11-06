(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'), require('./funds-ledger-core.js'));
  else root.FundsLedgerExport = factory(root.XLSX, root.JSZip, root.FundsLedgerCore);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip, Core) {
  'use strict';
  const width = text => [...String(text)].reduce((n, c) => n + (c.codePointAt(0) > 255 ? 2 : 1), 0);
  return async function(result) {
    const book = Core.workbook(result);
    const zip = await Zip.loadAsync(X.write(book, {type:'array', bookType:'xlsx'}));
    let styles = await zip.file('xl/styles.xml').async('string');
    const original = styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
    const count = original.length;
    const xf = (format, wrap, header = false) => `<xf numFmtId="${format}" fontId="${header ? 1 : 0}" fillId="${header ? 3 : 2}" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="${wrap ? 1 : 0}" shrinkToFit="0"/></xf>`;
    const formats = original.map(s => s.match(/\bnumFmtId="(\d+)"/)?.[1] || '0');
    const fonts = '<fonts count="2"><font><sz val="11"/><color rgb="FF172033"/><name val="等线"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font></fonts>';
    const fills = '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315B9D"/><bgColor indexed="64"/></patternFill></fill></fills>';
    const borders = '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>' + ['left', 'right', 'top', 'bottom'].map(side => `<${side} style="thin"><color rgb="FFD6DEE8"/></${side}>`).join('') + '<diagonal/></border></borders>';
    styles = styles.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/, fonts).replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, fills).replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, borders).replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${count * 2 + 1}">${formats.map(f => xf(f, false)).join('')}${formats.map(f => xf(f, true)).join('')}${xf(0, false, true)}</cellXfs>`);
    zip.file('xl/styles.xml', styles);
    for (let index = 0; index < book.SheetNames.length; index++) {
      const sheet = book.Sheets[book.SheetNames[index]], rows = result.sheets[index].rows;
      const noWrap = new Set(['对公付款OA流程号','申请付款金额','应付款时间','预估付款时间']);
      const wrapColumns = result.sheets[index].headers.map((name,c) => noWrap.has(name) ? -1 : c).filter(c => c >= 0);
      const path = `xl/worksheets/sheet${index + 1}.xml`;
      let xml = await zip.file(path).async('string');
      xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (tag, attrs, end) => {
        const address = attrs.match(/\br="([A-Z]+)(\d+)"/);
        if (!address) return tag;
        const old = Number(attrs.match(/\bs="(\d+)"/)?.[1] || 0);
        const style = address[2] === '1' ? count * 2 : old + (wrapColumns.includes(X.utils.decode_col(address[1])) ? count : 0);
        return '<c' + attrs.replace(/\s+s="\d+"/, '') + ` s="${style}"${end}>`;
      });
      xml = xml.replace(/<row\b([^>]*)>/g, (tag, attrs) => {
        const row = Number(attrs.match(/\br="(\d+)"/)?.[1]);
        let height = 26;
        if (row > 1) {
          const lines = Math.max(1, ...wrapColumns.map(c => Core.formatCell(rows[row - 2]?.[c]).split(/\r?\n/).reduce((n, line) => n + Math.max(1, Math.ceil(width(line) / Math.max(1, (sheet['!cols'][c].wch || 20) - 2))), 0)));
          height = Math.min(409, Math.max(24, lines * 15 + 8));
        }
        return '<row' + attrs.replace(/\s+(?:ht|customHeight)="[^"]*"/g, '') + ` ht="${height}" customHeight="1">`;
      });
      const views = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
      xml = /<sheetViews\b/.test(xml) ? xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, views) : xml.replace(/<sheetData\b/, views + '<sheetData');
      zip.file(path, xml);
    }
    return zip.generateAsync({type:'uint8array', compression:'DEFLATE'});
  };
});
