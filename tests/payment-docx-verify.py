"""After rendering browser-test stress.docx, verify real pagination and content.

Usage: python tests/payment-docx-verify.py assets/payment-template.docx /tmp/qa/stress.docx /tmp/render/stress.pdf
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
    tables = doc.findall('w:body/w:tbl', ns)
    assert len(tables) == 12
    assert [len(tables[i].findall('w:tr', ns)) - 2 for i in (1, 5, 9)] == [1, 30, 45]
    for group in range(3):
        assert text(tables[group * 4 + 2]) == text(original[2]), 'Approval text changed'
        assert text(tables[group * 4 + 3]) == text(original[3]), 'Settlement text changed'
    assert len(doc.findall('.//w:sectPr', ns)) == 3

pages = PdfReader(rendered_pdf).pages
assert len(pages) == 3, f'Expected 3 actual Word pages, got {len(pages)}'
for i, page in enumerate(pages):
    content = re.sub(r'\s+', '', page.extract_text())
    assert '出纳签字' in content, f'Settlement missing on page {i + 1}'
    assert '三、审批意见' in content and '四、财务办结记录' in content
    if i in (1, 2):
        for j in range(1, (30 if i == 1 else 45) + 1):
            assert f'记录{j}' in content, f'Detail missing: page {i + 1}, row {j}'
    if i == 1:
        assert content.count('0012345678901234') == 30, 'Bank account content lost'
print('PASS: 3 actual Word pages, all 76 details retained, bank data intact, original approval and settlement retained, only document.xml changed')
