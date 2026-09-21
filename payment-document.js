/* Template-based DOCX editing and a shared, measured one-page layout. */
(function () {
  'use strict';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const C = PaymentCore;
  const FONT = '"SimSun", "Songti SC", "STSong", serif';
  const PAGE = { width: 841.9, height: 595.3 };
  const children = (e, name) => [...e.children].filter(c => c.namespaceURI === W && c.localName === name);
  const all = (e, name) => [...e.getElementsByTagNameNS(W, name)];
  const attr = (e, name) => e?.getAttributeNS(W, name);
  function node(doc, name, attrs = {}) {
    const e = doc.createElementNS(W, 'w:' + name);
    for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, 'w:' + k, String(v));
    return e;
  }
  function prop(parent, name, attrs = {}) {
    let e = children(parent, name)[0];
    if (!e) { e = node(parent.ownerDocument, name); parent.prepend(e); }
    for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, 'w:' + k, String(v));
    return e;
  }
  function remove(parent, names) { for (const name of names) children(parent, name).forEach(e => e.remove()); }
  function paragraphText(p) {
    let out = '';
    for (const e of p.getElementsByTagName('*')) {
      if (e.namespaceURI !== W) continue;
      if (e.localName === 't') out += e.textContent;
      if (e.localName === 'br' || e.localName === 'cr') out += '\n';
      if (e.localName === 'tab') out += '    ';
    }
    return out;
  }
  const cellText = cell => children(cell, 'p').map(paragraphText).join('\n');
  function putText(target, value) {
    const doc = target.ownerDocument;
    let p = target.localName === 'p' ? target : children(target, 'p')[0];
    if (!p) { p = node(doc, 'p'); target.append(p); }
    if (target !== p) children(target, 'p').slice(1).forEach(e => e.remove());
    const rp = all(p, 'rPr')[0]?.cloneNode(true);
    [...p.children].filter(e => e.localName !== 'pPr').forEach(e => e.remove());
    const r = node(doc, 'r'); if (rp) r.append(rp);
    String(value).split('\n').forEach((s, i) => {
      if (i) r.append(node(doc, 'br'));
      const t = node(doc, 't'); t.setAttribute('xml:space', 'preserve'); t.textContent = s; r.append(t);
    });
    p.append(r);
  }
  function parseXml(s) {
    const doc = new DOMParser().parseFromString(s, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('模板 XML 无法读取');
    return doc;
  }
  async function readTemplate(bytes) {
    if (bytes.byteLength > 10 * 1024 * 1024) throw new Error('模板请控制在 10 MB 以内');
    const zip = await JSZip.loadAsync(bytes);
    const file = zip.file('word/document.xml');
    if (!file) throw new Error('请选择有效的 .docx 模板');
    const xml = await file.async('string');
    if (xml.length > 2000000) throw new Error('模板结构过大');
    const doc = parseXml(xml), body = all(doc, 'body')[0];
    if (!body) throw new Error('模板缺少正文');
    const tables = children(body, 'tbl');
    const paragraphs = children(body, 'p').filter(p => paragraphText(p).trim());
    const tableRows = tables.map(t => children(t, 'tr'));
    const counts = tableRows.map(rs => rs.length);
    const labels = ['申请部门', '经办人', '付款类型', '汇总付款总金额', '计划付款日期'];
    if (tables.length !== 4 || paragraphs.length !== 5 || counts[0] !== 3 || counts[1] < 3 || counts[2] !== 3 || counts[3] !== 2 ||
        !paragraphText(paragraphs[0]).includes('付款申请单') || !paragraphText(paragraphs[1]).includes('申请日期') ||
        !paragraphText(paragraphs[2]).includes('付款明细') || !paragraphText(paragraphs[3]).includes('审批意见') || !paragraphText(paragraphs[4]).includes('财务办结记录') ||
        !labels.every(s => tables[0].textContent.includes(s)) ||
        children(tableRows[1][0], 'tc').map(cellText).join('|') !== ['序号', ...C.fields].join('|') ||
        children(tableRows[0][0], 'tc').length !== 4 || children(tableRows[0][1], 'tc').length !== 2 || children(tableRows[0][2], 'tc').length !== 4 ||
        !tableRows[1].slice(1, -1).every(r => children(r, 'tc').length === 10) ||
        children(tableRows[1].at(-1), 'tc').length !== 8 || !cellText(children(tableRows[1].at(-1), 'tc')[0]).includes('合计') ||
        ![...tableRows[2], ...tableRows[3]].every(r => children(r, 'tc').length === 2)) {
      throw new Error('模板结构不兼容：请使用与公用模板相同的四张表、明细列及章节结构');
    }
    const metadata = tableRows[0].map(r => children(r, 'tc'));
    if ([cellText(metadata[0][0]), cellText(metadata[0][2]), cellText(metadata[1][0]), cellText(metadata[2][0]), cellText(metadata[2][2])].map(s => s.trim()).join('|') !== labels.join('|') ||
        [...body.children].some(e => !['p', 'tbl', 'sectPr'].includes(e.localName))) {
      throw new Error('模板字段位置不兼容，请保留公用模板的标签及排列顺序');
    }
    const unsupported = ['drawing', 'pict', 'altChunk', 'sdt', 'fldChar', 'fldSimple', 'ins', 'del', 'vMerge', 'hyperlink', 'footnoteReference', 'endnoteReference'];
    if (unsupported.some(n => all(body, n).length) || all(doc, 'sectPr').length !== 1 || children(body, 'sectPr').length !== 1 ||
        all(doc, 'headerReference').length || all(doc, 'footerReference').length || all(body, 'tbl').length !== 4) {
      throw new Error('模板含不支持的图片、嵌套表格、修订、域或页眉页脚，请使用同结构的普通表格模板');
    }
    for (const t of tables) {
      const grid = children(children(t, 'tblGrid')[0] || t, 'gridCol');
      if (!grid.length || grid.some(e => !(Number(attr(e, 'w')) > 0))) throw new Error('模板表格列宽无效');
      for (const r of children(t, 'tr')) {
        if (children(r, 'tc').reduce((n, c) => n + Number(attr(children(c, 'tcPr')[0]?.getElementsByTagNameNS(W, 'gridSpan')[0], 'val') || 1), 0) !== grid.length) throw new Error('模板合并单元格结构不兼容');
      }
    }
    return { bytes, xml };
  }
  function fill(template, group) {
    const doc = parseXml(template.xml), body = all(doc, 'body')[0];
    for (const p of children(body, 'p')) if (!paragraphText(p).trim()) p.remove();
    const ps = children(body, 'p'), ts = children(body, 'tbl');
    putText(ps[0], group.subject + '付款申请单');
    putText(ps[1], '申请日期：' + displayDate(group.info['申请日期']));
    const rs = children(ts[0], 'tr');
    putText(children(rs[0], 'tc')[1], group.info['申请部门']);
    putText(children(rs[0], 'tc')[3], group.info['经办人']);
    const typeCell = children(rs[1], 'tc')[1];
    putText(typeCell, '☑ 采购货款    □ 服务费用    □ 模具款    □ 尾款    □ 预付款    □ 其他＿＿＿＿＿＿＿＿');
    putText(children(rs[2], 'tc')[1], `小写：￥${C.money(group.total)}\n大写：${C.upper(group.total)}`);
    const payment = children(rs[2], 'tc')[3];
    putText(payment, `${displayDate(group.info['计划付款日期'])}\n付款方式：☑ 对公转账  □ 其他＿＿＿`);
    const rows = children(ts[1], 'tr'), proto = rows[1].cloneNode(true), total = rows.at(-1);
    rows.slice(1, -1).forEach(r => r.remove());
    group.rows.forEach((r, i) => {
      const tr = proto.cloneNode(true);
      const values = [String(i + 1), ...C.fields.map(f => f === '本次申请付款金额' ? C.money(r._cents) : C.text(r[f]))];
      children(tr, 'tc').forEach((cell, j) => putText(cell, values[j]));
      ts[1].insertBefore(tr, total);
    });
    putText(children(total, 'tc')[1], `本次申请付款合计（小写）￥${C.money(group.total)}`);
    return { doc, body };
  }
  function displayDate(value) {
    return value ? value.replace(/^(\d+)-(\d+)-(\d+)$/, '$1年$2月$3日') : '＿＿＿＿年＿＿月＿＿日';
  }
  function wrap(value, width, measure) {
    const lines = [];
    for (const part of value.split('\n')) {
      let line = '';
      for (const ch of part) {
        if (line && measure(line + ch) > width) { lines.push(line); line = ''; }
        if (measure(ch) > width) throw new Error('列宽不足，无法完整排入文字');
        line += ch;
      }
      lines.push(line);
    }
    return lines;
  }
  function paragraphModel(p, size, line, width, context, title = false) {
    size = Math.max(1, Math.floor(size * 2) / 2);
    line = Math.ceil(line * 20) / 20;
    const value = paragraphText(p);
    const bold = title || all(p, 'b').some(e => attr(e, 'val') !== '0');
    context.font = `${bold ? 'bold ' : ''}${size}px ${FONT}`;
    const measure = s => context.measureText(s).width;
    const lines = wrap(value, width * .92, measure);
    const align = title ? 'center' : attr(all(p, 'jc')[0], 'val') || 'left';
    return { element: p, value, size, line, bold, align, lines, height: lines.length * line };
  }
  function measureLayout(filled, config) {
    const ctx = document.createElement('canvas').getContext('2d');
    const width = PAGE.width - config.mx * 2, blocks = [];
    let height = 0, tableIndex = 0, paragraphIndex = 0;
    for (const el of filled.body.children) {
      if (el.localName === 'p') {
        const title = paragraphIndex++ === 0;
        const size = title ? Math.max(config.body, config.body * 1.5) : config.body;
        const p = paragraphModel(el, size, title ? config.bodyLine * 1.5 : config.bodyLine, width, ctx, title);
        const gap = config.bodyLine * .4;
        blocks.push({ type: 'p', ...p, gap }); height += p.height + gap;
      } else if (el.localName === 'tbl') {
        const raw = children(children(el, 'tblGrid')[0], 'gridCol').map(e => Number(attr(e, 'w')));
        const sum = raw.reduce((a, b) => a + b, 0), grid = raw.map(w => width * w / sum);
        const rows = children(el, 'tr').map(row => {
          let col = 0;
          const cells = children(row, 'tc').map(cell => {
            const span = Number(attr(all(cell, 'gridSpan')[0], 'val') || 1);
            const cw = grid.slice(col, col + span).reduce((a, b) => a + b, 0); col += span;
            const ps = children(cell, 'p').map(p => paragraphModel(p, config.table, config.tableLine, cw - config.pad * 2, ctx));
            const fill = attr(all(cell, 'shd')[0], 'fill');
            return { element: cell, span, width: cw, ps, height: ps.reduce((a, p) => a + p.height, 0), fill: fill && /^[0-9a-f]{6}$/i.test(fill) ? '#' + fill : '#ffffff' };
          });
          // Space for handwriting remains even on a one-row application.
          const minimum = tableIndex >= 2 ? config.tableLine * 1.7 : config.tableLine;
          const contentHeight = Math.max(minimum, ...cells.map(c => c.height)) + config.pad * 2 + config.tableLine * .15;
          // Word adds border geometry outside the minimum trHeight. Account for it
          // in the page budget, rather than applying the allowance twice to the row.
          return { element: row, cells, contentHeight, height: contentHeight + 1.2 };
        });
        const th = rows.reduce((a, r) => a + r.height, 0);
        blocks.push({ type: 'table', element: el, grid, rows, height: th }); height += th;
        tableIndex++;
      }
    }
    return { doc: filled.doc, domBody: filled.body, ...config, blocks, width, height };
  }
  async function layout(template, group) {
    await document.fonts.ready;
    const filled = fill(template, group);
    let config = { body: 8, table: 7, bodyLine: 10, tableLine: 9, mx: 17.01, my: 14.17, pad: 1.4 };
    for (let attempt = 0; attempt < 60; attempt++) {
      const result = measureLayout(filled, config);
      // Reserve Word's required terminal paragraph and rounding/font differences.
      if (result.height <= PAGE.height - config.my * 2 - 24) return result;
      if (attempt === 0) config = { body: 7, table: 6.5, bodyLine: 10, tableLine: 9, mx: 11.34, my: 11.34, pad: 1 };
      else {
        const factor = Math.min(.94, (PAGE.height - config.my * 2 - 30) / result.height);
        config = { ...config, body: config.body * factor, table: config.table * factor, bodyLine: config.bodyLine * factor, tableLine: config.tableLine * factor, pad: config.pad * factor };
      }
      if (config.table < 1) throw new Error(`${group.subject}：内容过多，低于 Word 可表示的字号仍无法完整排入一页，生成失败`);
    }
    throw new Error(`${group.subject}：一页布局计算失败`);
  }
  function styleParagraph(p) {
    const el = p.element, doc = el.ownerDocument;
    putText(el, p.lines.join('\n'));
    const pp = prop(el, 'pPr');
    remove(pp, ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'sectPr', 'tabs', 'ind', 'contextualSpacing']);
    prop(pp, 'spacing', { before: 0, after: 0, line: Math.ceil(p.line * 20), lineRule: 'exact' });
    prop(pp, 'snapToGrid', { val: 0 }); prop(pp, 'widowControl', { val: 0 });
    prop(pp, 'jc', { val: ['center', 'right'].includes(p.align) ? p.align : 'left' });
    // OOXML size is half-points. Round down so the measured line remains safe.
    const half = Math.max(2, Math.floor(p.size * 2));
    for (const r of children(el, 'r')) {
      const rp = prop(r, 'rPr');
      remove(rp, ['rStyle', 'sz', 'szCs', 'rFonts', 'spacing', 'position', 'w', 'color', 'b', 'bCs']);
      rp.append(node(doc, 'rFonts', { ascii: 'SimSun', hAnsi: 'SimSun', eastAsia: '宋体', cs: 'SimSun' }));
      rp.append(node(doc, 'sz', { val: half }), node(doc, 'szCs', { val: half }), node(doc, 'color', { val: '000000' }));
      if (p.bold) rp.append(node(doc, 'b'));
    }
  }
  function applyLayout(result) {
    const doc = result.doc;
    for (const b of result.blocks) {
      if (b.type === 'p') {
        styleParagraph(b);
        prop(prop(b.element, 'pPr'), 'spacing', { after: Math.ceil(b.gap * 20) });
        continue;
      }
      const tp = prop(b.element, 'tblPr');
      remove(tp, ['tblpPr', 'tblInd', 'tblCellSpacing']);
      prop(tp, 'tblW', { w: Math.floor(result.width * 20), type: 'dxa' });
      prop(tp, 'tblLayout', { type: 'fixed' }); prop(tp, 'jc', { val: 'left' });
      const margins = prop(tp, 'tblCellMar');
      for (const side of ['top', 'left', 'bottom', 'right']) prop(margins, side, { w: Math.floor(result.pad * 20), type: 'dxa' });
      children(children(b.element, 'tblGrid')[0], 'gridCol').forEach((c, i) => c.setAttributeNS(W, 'w:w', Math.floor(b.grid[i] * 20)));
      for (const row of b.rows) {
        const rp = prop(row.element, 'trPr'); remove(rp, ['trHeight']);
        prop(rp, 'cantSplit');
        prop(rp, 'trHeight', { val: Math.ceil(row.contentHeight * 20), hRule: 'atLeast' });
        for (const cell of row.cells) {
          const cp = prop(cell.element, 'tcPr'); remove(cp, ['tcMar', 'noWrap', 'tcFitText']);
          prop(cp, 'tcW', { w: Math.floor(cell.width * 20), type: 'dxa' });
          prop(cp, 'vAlign', { val: 'center' });
          cell.ps.forEach(styleParagraph);
        }
      }
    }
    const section = children(result.domBody, 'sectPr')[0];
    prop(section, 'pgSz', { w: 16838, h: 11906, orient: 'landscape' });
    prop(section, 'pgMar', { top: Math.round(result.my * 20), bottom: Math.round(result.my * 20), left: Math.round(result.mx * 20), right: Math.round(result.mx * 20), header: 0, footer: 0, gutter: 0 });
    remove(section, ['docGrid']); prop(section, 'type', { val: 'nextPage' });
    return section;
  }
  async function docx(template, layouts) {
    const doc = parseXml(template.xml), body = all(doc, 'body')[0];
    body.replaceChildren();
    layouts.forEach((l, i) => {
      const section = applyLayout(l);
      for (const block of l.blocks) body.append(doc.importNode(block.element, true));
      if (i < layouts.length - 1) {
        const p = node(doc, 'p'), pp = node(doc, 'pPr');
        pp.append(node(doc, 'spacing', { before: 0, after: 0, line: 20, lineRule: 'exact' }), doc.importNode(section, true));
        p.append(pp); body.append(p);
      } else body.append(doc.importNode(section, true));
    });
    const zip = await JSZip.loadAsync(template.bytes);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }
  function canvas(result, resolution = 3) {
    const c = document.createElement('canvas');
    c.width = Math.ceil(PAGE.width * resolution); c.height = Math.ceil(PAGE.height * resolution);
    const ctx = c.getContext('2d'); ctx.scale(resolution, resolution);
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, PAGE.width, PAGE.height);
    ctx.textBaseline = 'alphabetic';
    function drawP(p, x, y, width) {
      ctx.fillStyle = '#000000'; ctx.font = `${p.bold ? 'bold ' : ''}${p.size}px ${FONT}`;
      p.lines.forEach((s, i) => {
        const w = ctx.measureText(s).width;
        const shift = p.align === 'center' ? (width - w) / 2 : p.align === 'right' ? width - w : 0;
        ctx.fillText(s, x + shift, y + i * p.line + (p.line - p.size) / 2 + p.size * .85);
      });
    }
    let y = result.my;
    for (const b of result.blocks) {
      if (b.type === 'p') { drawP(b, result.mx, y, result.width); y += b.height + b.gap; continue; }
      for (const row of b.rows) {
        let x = result.mx;
        for (const cell of row.cells) {
          ctx.fillStyle = cell.fill; ctx.fillRect(x, y, cell.width, row.height);
          ctx.strokeStyle = '#000000'; ctx.lineWidth = .45; ctx.strokeRect(x, y, cell.width, row.height);
          let py = y + (row.height - cell.height) / 2;
          for (const p of cell.ps) { drawP(p, x + result.pad, py, cell.width - 2 * result.pad); py += p.height; }
          x += cell.width;
        }
        y += row.height;
      }
    }
    if (y > PAGE.height - result.my) throw new Error('页面溢出，已停止生成');
    return c;
  }
  async function pdf(layouts) {
    const pdf = await PDFLib.PDFDocument.create();
    pdf.setTitle('付款申请单'); pdf.setCreator('供应链数据清洗');
    for (const l of layouts) {
      const image = await pdf.embedPng(canvas(l).toDataURL('image/png'));
      const page = pdf.addPage([PAGE.width, PAGE.height]);
      page.drawImage(image, { x: PAGE.width * .075, y: PAGE.height * .075, width: PAGE.width * .85, height: PAGE.height * .85 });
    }
    return new Blob([await pdf.save()], { type: 'application/pdf' });
  }
  window.PaymentDocument = { readTemplate, layout, docx, pdf, canvas, paragraphText, cellText };
})();
