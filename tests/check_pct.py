import zipfile, re
z = zipfile.ZipFile('D:/item/Claude_code_build/formula_format_conversion/word-template/md_test_fixed.docx')
content = z.read('word/document.xml').decode('utf-8')
w_texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', content)
backslash = chr(92)
pct = [(i, t) for i, t in enumerate(w_texts) if '%' in t and backslash in t]
print(f'w:t with backslash+percent: {len(pct)}')
for i, t in pct[:5]:
    print(f'  [{i}] {repr(t[:120])}')

# Also check for any control/escape issues in math text
m_texts = re.findall(r'<m:t[^>]*>(.*?)</m:t>', content)
m_pct = [(i, t) for i, t in enumerate(m_texts) if backslash in t]
print(f'\nm:t with backslash: {len(m_pct)}')
for i, t in m_pct[:5]:
    print(f'  [{i}] {repr(t[:120])}')
