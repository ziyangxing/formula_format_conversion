"""List all LaTeX commands used in md_test.md."""
import re
with open('D:/item/Claude_code_build/formula_format_conversion/word-template/md_test.md', 'r', encoding='utf-8') as f:
    content = f.read()

cmds = set()
for m in re.finditer(r'\\', content):
    # Get the command name after backslash
    rest = content[m.start()+1:]
    name = ''
    for c in rest:
        if c.isalpha():
            name += c
        else:
            break
    if name:
        cmds.add(name)

print('LaTeX commands used in md_test.md:')
for c in sorted(cmds):
    print(f'  \\{c}')
