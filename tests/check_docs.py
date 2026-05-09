"""Check what formula patterns exist in user documents."""
import zipfile, re, os, sys

for fname in ['format.docx', 'WPS.docx']:
    path = f'D:/item/Claude_code_build/formula_format_conversion/word-template/{fname}'
    if not os.path.exists(path):
        print(f'{fname}: NOT FOUND')
        continue

    z = zipfile.ZipFile(path)
    from lxml import etree
    tree = etree.parse(z.open('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    texts = [t.text or '' for t in tree.findall('.//w:t', ns)]
    full = ''.join(texts)

    print(f'=== {fname} ===')
    print(f'Total text: {len(full)} chars')
    print(f'Has $$: {"$$" in full}')
    print(f'Has $: {"$" in full}')
    print(f'Has \\\\(: {"\\\\(" in full}')
    print(f'Has \\\\[: {"\\\\[" in full}')
    print(f'Has \\\\frac: {"\\\\frac" in full}')
    print(f'Has \\\\dfrac: {"\\\\dfrac" in full}')
    print(f'Has ^: {"^" in full}')
    print(f'Has _: {"_" in full}')
    bs_count = full.count(chr(92))
    print(f'Backslash count: {bs_count}')

    # Show text content (use repr to handle encoding)
    print(f'First 500 chars (repr):')
    print(repr(full[:500]))
    print()
