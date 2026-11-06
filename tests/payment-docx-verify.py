"""After rendering a browser-test DOCX, verify real pagination and content.

Usage: python tests/payment-docx-verify.py assets/payment-template.docx /tmp/qa/combined-76-rows.docx /tmp/render/combined-76-rows.pdf
Requires pypdf in the selected verification runtime; does not change any files.
"""
import re
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as E
from pypdf import PdfReader

reference, result, rendered_pdf = sys.argv[1:]
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def text(element):
    return re.sub(r'\s+', '', ''.join(e.text or '' for e in element.findall('.//w:t', ns)))


with ZipFile(reference) as source, ZipFile(result) as output:
    changed = [n for n in source.namelist() if not n.endswith('/') and source.read(n) != output.read(n)]
    assert changed == ['word/document.xml'], changed
    original = E.fromstring(source.read('word/document.xml')).findall('w:body/w:tbl', ns)
    doc = E.fromstring(output.read('word/document.xml'))
    expected_pages, blocks = [], []
    for block in doc.find('w:body', ns):
        if block.tag == '{%s}sectPr' % ns['w'] or block.find('w:pPr/w:sectPr', ns) is not None:
            expected_pages.append(blocks)
            blocks = []
        else:
            blocks.append(block)
    assert not blocks and expected_pages, 'Missing final page section'

pages = PdfReader(rendered_pdf).pages
assert len(pages) == len(expected_pages), f'Expected {len(expected_pages)} actual Word pages, got {len(pages)}'
subject, previous_number, detail_count = None, 0, 0
for i, (page, blocks) in enumerate(zip(pages, expected_pages)):
    tables = [block for block in blocks if block.tag == '{%s}tbl' % ns['w']]
    assert len(tables) in (2, 4), f'Unexpected table structure on page {i + 1}'
    current_subject = text(blocks[0])
    if current_subject != subject:
        subject, previous_number = current_subject, 0
    final = i == len(pages) - 1 or text(expected_pages[i + 1][0]) != subject
    assert len(tables) == (4 if final else 2), 'Approval and settlement belong on the last subject page only'
    if final:
        assert text(tables[2]) == text(original[2]), 'Approval text changed'
        assert text(tables[3]) == text(original[3]), 'Settlement text changed'
    content = re.sub(r'\s+', '', page.extract_text())
    # PDF fallback-font runs can extract check marks after the surrounding text.
    assert content.count('☑') == sum(text(block).count('☑') for block in blocks)
    for block in blocks:
        elements = block.findall('.//w:tc', ns) if block.tag == '{%s}tbl' % ns['w'] else [block]
        for element in elements:
            assert text(element).replace('☑', '') in content.replace('☑', ''), f'Content missing on actual page {i + 1}: {text(element)}'
    for row in tables[1].findall('w:tr', ns)[1:]:
        cells = row.findall('w:tc', ns)
        if len(cells) == 10:
            previous_number += 1
            detail_count += 1
            assert text(cells[0]) == str(previous_number), 'Detail numbering must continue across pages'
print(f'PASS: {len(pages)} actual Word pages, {detail_count} complete details, original approval and settlement on each subject final page, only document.xml changed')
