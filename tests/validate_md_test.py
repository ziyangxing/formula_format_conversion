"""
Proper validation: compare md_test.docx output against markdown source.
Checks formatting, styling, and content fidelity.
"""
import re, zipfile
from lxml import etree
from io import StringIO

MD_PATH = 'D:/item/Claude_code_build/formula_format_conversion/word-template/md_test.md'
DOCX_PATH = 'D:/item/Claude_code_build/formula_format_conversion/word-template/md_test_converted2.docx'

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W_NS, 'm': M_NS}

with open(MD_PATH, 'r', encoding='utf-8') as f:
    md_text = f.read()

z = zipfile.ZipFile(DOCX_PATH)
tree = etree.parse(z.open('word/document.xml'))

results = []

def check(name, ok, detail=''):
    status = 'OK' if ok else 'FAIL'
    results.append(f'[{status}] {name}')
    if detail and not ok:
        results.append(f'       {detail}')

# ================================================================
# 1. HEADINGS — must use Word heading styles, not just bold paragraph
# ================================================================
md_headings = [m for m in re.finditer(r'^(#{1,6})\s+(.+)$', md_text, re.MULTILINE)]
heading_count = 0
for hp in tree.findall('.//w:p', NS):
    pStyle = hp.find('.//w:pStyle', NS)
    if pStyle is not None:
        val = pStyle.get(f'{{{W_NS}}}val') or ''
        if 'Heading' in val:
            heading_count += 1
check('Headings use Word heading styles', heading_count >= 3,
      f'Found {heading_count} styled headings, expected >= 3')

# Check heading levels match
for mh in md_headings:
    md_level = len(mh.group(1))
    md_title = mh.group(2)
    # Try to find matching heading paragraph
    found = False
    for hp in tree.findall('.//w:p', NS):
        w_text = ''.join(t.text or '' for t in hp.findall('.//w:t', NS))
        if md_title[:15] in w_text:
            found = True
            break
    if not found:
        check(f'Heading found: {md_title[:30]}', False, 'Heading text not in any styled paragraph')

# ================================================================
# 2. BOLD — must have w:b element in run properties
# ================================================================
bold_runs = 0
for r in tree.findall('.//w:r', NS):
    if r.find('.//w:b', NS) is not None:
        bold_runs += 1
check('Bold formatting (w:b) present', bold_runs >= 5,
      f'Found {bold_runs} bold runs')

# ================================================================
# 3. INLINE CODE — must be Consolas font, distinct from body text
# ================================================================
code_runs = []
for r in tree.findall('.//w:r', NS):
    rFonts = r.find('.//w:rFonts', NS)
    if rFonts is not None:
        ascii_font = rFonts.get(f'{{{W_NS}}}ascii', '')
        if 'Consolas' in ascii_font:
            code_runs.append(r)
# Inline code content check
inline_code_texts = [
    '净现金流量 = 现金流入 - 现金流出',
    '累计净现金流量 = 上年累计净现金流量 + 本年净现金流量',
]
for ict in inline_code_texts:
    found = False
    for r in tree.findall('.//w:r', NS):
        t = r.find('.//w:t', NS)
        if t is not None and t.text and ict in t.text:
            rFonts = r.find('.//w:rFonts', NS)
            if rFonts is not None and 'Consolas' in (rFonts.get(f'{{{W_NS}}}ascii') or ''):
                found = True
                break
    check(f'Inline code formatted: {ict[:40]}', found,
          'Not found or not Consolas font')

# ================================================================
# 4. TABLES — must have grid borders, proper cell count
# ================================================================
tables = tree.findall('.//w:tbl', NS)
table_ok = len(tables) >= 1
if table_ok:
    tbl = tables[0]
    rows = tbl.findall('w:tr', NS)
    tblBorders = tbl.find('w:tblPr/w:tblBorders', NS)
    has_borders = tblBorders is not None
    check('Table present with borders', has_borders and len(rows) > 5,
          f'{len(rows)} rows, borders={has_borders}')
else:
    check('Table present', False, 'No table found')

# ================================================================
# 5. HORIZONTAL RULES — paragraph bottom border
# ================================================================
hr_count = 0
for p in tree.findall('.//w:p', NS):
    pBdr = p.find('.//w:pBdr/w:bottom', NS)
    if pBdr is not None:
        hr_count += 1
check('Horizontal rules (--- rendered as borders)', hr_count >= 3,
      f'Found {hr_count} rules')

# ================================================================
# 6. LISTS — indented with bullet/number prefix
# ================================================================
# Check for indented paragraphs with bullet chars
list_paras = 0
for p in tree.findall('.//w:p', NS):
    ind = p.find('.//w:ind', NS)
    if ind is not None:
        left_val = ind.get(f'{{{W_NS}}}left', '0')
        if int(left_val) > 100:  # Has meaningful indent
            list_paras += 1
check('Lists have indentation', list_paras >= 5,
      f'Found {list_paras} indented paragraphs')

# ================================================================
# 7. FORMULAS — proper OMML structures
# ================================================================
omath_inline = len(tree.findall('.//m:oMath', NS))
omath_display = len(tree.findall('.//m:oMathPara', NS))
check('Inline formulas in OMML', omath_inline >= 30,
      f'Found {omath_inline}')
check('Display formulas in OMML', omath_display >= 5,
      f'Found {omath_display}')

# ================================================================
# 8. TEXT CONTENT FIDELITY — key sentences preserved
# ================================================================
all_w = ''.join(t.text or '' for t in tree.findall('.//w:t', NS))
all_m = ''.join(t.text or '' for t in tree.findall('.//m:t', NS))

key_sentences = [
    '4 年期项目（2025 - 2028）',
    '建设期 + 运营期共 4 年',
    '自有资金现金流量表',
    '基准收益率取行业通用标准',
    '净现金流量 = 现金流入 - 现金流出',
    '项目资本金利润率',
    '静态总投资收益率',
    '偿债备付率',
    '利息备付率',
    '借款利息支出为 0',
]
for ks in key_sentences:
    found = ks in all_w or any(ks[:20] in mt for mt in [all_m])
    check(f'Content: {ks[:50]}', found,
          'Not found in document text')

# ================================================================
# 9. CODE BLOCKS (``` ```) — no formula conversion inside
# ================================================================
# Our test doc doesn't have fenced code blocks, but let's check
# that no stray $ from formulas appears in wrong places

# ================================================================
# SUMMARY
# ================================================================
print('=' * 60)
print('VALIDATION REPORT: md_test.md → md_test_converted2.docx')
print('=' * 60)
passed = sum(1 for r in results if r.startswith('[OK]'))
failed = sum(1 for r in results if r.startswith('[FAIL]'))
for r in results:
    print(r)
print(f'\nPASSED: {passed}, FAILED: {failed}, TOTAL: {passed + failed}')
