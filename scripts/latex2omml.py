#!/usr/bin/env python3
"""
latex2omml.py - LaTeX → OMML Professional 编译器

将 LaTeX 数学公式解析为 OMML (Office Math Markup Language) Professional 格式 XML。
生成结构化元素: m:f(分数), m:sSup(上标), m:sSub(下标), m:rad(根号),
m:nary(求和/积分), m:acc(重音符) 等，使公式在 Word/WPS 中直接显示为专业格式。
"""

import re
from lxml import etree
from docx.oxml.ns import qn

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = f"{{{MATH_NS}}}"

# ============================================================
# Greek letter mappings
# ============================================================
GREEK_LETTERS = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
    'epsilon': 'ε', 'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η',
    'theta': 'θ', 'vartheta': 'ϑ', 'iota': 'ι', 'kappa': 'κ',
    'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ',
    'pi': 'π', 'varpi': 'ϖ', 'rho': 'ρ', 'varrho': 'ϱ',
    'sigma': 'σ', 'varsigma': 'ς', 'tau': 'τ', 'upsilon': 'υ',
    'phi': 'φ', 'varphi': 'ϕ', 'chi': 'χ', 'psi': 'ψ',
    'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ',
    'Lambda': 'Λ', 'Xi': 'Ξ', 'Pi': 'Π',
    'Sigma': 'Σ', 'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
}

# Math operators that should be rendered as text
MATH_OPERATORS = set('+-*/=<>|!')

# Functions that get special treatment
MATH_FUNCTIONS = {'sin', 'cos', 'tan', 'log', 'ln', 'lim', 'exp', 'max', 'min', 'det', 'dim', 'gcd', 'hom', 'ker', 'Pr', 'sup', 'inf', 'arg', 'deg'}


# ============================================================
# OMML XML Builders
# ============================================================

def make_elements():
    """Create a namespace-aware element factory."""
    return {}

def _m(tag):
    """Create an m: namespace element."""
    return etree.Element(f"{M}{tag}")

def make_mr(text=""):
    """<m:r><m:t>text</m:t></m:r>"""
    mr = _m('r')
    mt = _m('t')
    mt.text = text if text else ''
    mr.append(mt)
    return mr


def make_text_run(text):
    """Create a math run from text, handling special characters."""
    if not text:
        return make_mr('')
    mr = _m('r')
    mt = _m('t')
    mt.text = text
    # Handle spaces specially in OMML
    if ' ' in text:
        mt.set(f'{{{MATH_NS}}}xml:space', 'preserve')
    mr.append(mt)
    return mr


def make_fraction(num_elements, den_elements):
    """<m:f><m:fPr/><m:num>...</m:num><m:den>...</m:den></m:f>"""
    mf = _m('f')
    mf.append(_m('fPr'))
    mnum = _m('num')
    for e in num_elements:
        mnum.append(e)
    mden = _m('den')
    for e in den_elements:
        mden.append(e)
    mf.append(mnum)
    mf.append(mden)
    return mf


def make_superscript(base_elements, sup_elements):
    """<m:sSup><m:e>...</m:e><m:sup>...</m:sup></m:sSup>"""
    mss = _m('sSup')
    mss.append(_m('sSupPr'))
    me = _m('e')
    for e in base_elements:
        me.append(e)
    ms = _m('sup')
    for e in sup_elements:
        ms.append(e)
    mss.append(me)
    mss.append(ms)
    return mss, [mss]  # Return as single-element list wrapper


def make_subscript(base_elements, sub_elements):
    """<m:sSub><m:e>...</m:e><m:sub>...</m:sub></m:sSub>"""
    mss = _m('sSub')
    mss.append(_m('sSubPr'))
    me = _m('e')
    for e in base_elements:
        me.append(e)
    ms = _m('sub')
    for e in sub_elements:
        ms.append(e)
    mss.append(me)
    mss.append(ms)
    return [mss]


def make_subsup(base_elements, sub_elements, sup_elements):
    """<m:sSubSup><m:e>...</m:e><m:sub>...</m:sub><m:sup>...</m:sup></m:sSubSup>"""
    mss = _m('sSubSup')
    mss.append(_m('sSubSupPr'))
    me = _m('e')
    for e in base_elements:
        me.append(e)
    msub = _m('sub')
    for e in sub_elements:
        msub.append(e)
    msup = _m('sup')
    for e in sup_elements:
        msup.append(e)
    mss.append(me)
    mss.append(msub)
    mss.append(msup)
    return [mss]


def make_radical(deg_elements, content_elements):
    """<m:rad><m:radPr/><m:deg>...</m:deg><m:e>...</m:e></m:rad>"""
    mrad = _m('rad')
    mrad.append(_m('radPr'))
    mdeg = _m('deg')
    if deg_elements:
        for e in deg_elements:
            mdeg.append(e)
    me = _m('e')
    for e in content_elements:
        me.append(e)
    mrad.append(mdeg)
    mrad.append(me)
    return [mrad]


def make_nary(operator_char, sub_elements, sup_elements, base_elements=None):
    """<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr>...</m:nary>"""
    mnary = _m('nary')
    mnarypr = _m('naryPr')
    mchr = _m('chr')
    mchr.set(f'{M}val', operator_char)
    mnarypr.append(mchr)
    mlimloc = _m('limLoc')
    mlimloc.set(f'{M}val', 'undOvr')
    mnarypr.append(mlimloc)
    mnary.append(mnarypr)

    if sub_elements:
        msub = _m('sub')
        for e in sub_elements:
            msub.append(e)
        mnary.append(msub)
    if sup_elements:
        msup = _m('sup')
        for e in sup_elements:
            msup.append(e)
        mnary.append(msup)

    me = _m('e')
    if base_elements:
        for e in base_elements:
            me.append(e)
    else:
        me.append(make_mr(''))
    mnary.append(me)
    return [mnary]


def make_accent(accent_char, base_elements):
    """<m:acc><m:accPr><m:chr m:val="̅"/></m:accPr><m:e>...</m:e></m:acc>"""
    macc = _m('acc')
    maccpr = _m('accPr')
    mchr = _m('chr')
    mchr.set(f'{M}val', accent_char)
    maccpr.append(mchr)
    macc.append(maccpr)
    me = _m('e')
    for e in base_elements:
        me.append(e)
    macc.append(me)
    return [macc]


def make_delimiter(left_char, right_char, content_elements):
    """<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>...</m:e></m:d>"""
    md = _m('d')
    mdpr = _m('dPr')
    if left_char:
        beg = _m('begChr')
        beg.set(f'{M}val', left_char)
        mdpr.append(beg)
    if right_char:
        end = _m('endChr')
        end.set(f'{M}val', right_char)
        mdpr.append(end)
    md.append(mdpr)
    me = _m('e')
    for e in content_elements:
        me.append(e)
    md.append(me)
    return [md]


def make_function(func_name, arg_elements):
    """<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e>...</m:e></m:func>"""
    mfunc = _m('func')
    mfname = _m('fName')
    mfname.append(make_mr(func_name))
    mfunc.append(mfname)
    me = _m('e')
    for e in arg_elements:
        me.append(e)
    mfunc.append(me)
    return [mfunc]


def make_group_chr(chr_val, content_elements, pos='bottom'):
    """<m:groupChr><m:groupChrPr><m:chr m:val="︷"/><m:pos m:val="bot"/></m:groupChrPr><m:e>...</m:e></m:groupChr>"""
    mgc = _m('groupChr')
    mgcpr = _m('groupChrPr')
    mchr = _m('chr')
    mchr.set(f'{M}val', chr_val)
    mgcpr.append(mchr)
    mpos = _m('pos')
    mpos.set(f'{M}val', 'bot')
    mgcpr.append(mpos)
    mgc.append(mgcpr)
    me = _m('e')
    for e in content_elements:
        me.append(e)
    mgc.append(me)
    return [mgc]


# ============================================================
# LaTeX Tokenizer
# ============================================================

TOKEN_TEXT = 'TEXT'
TOKEN_COMMAND = 'CMD'
TOKEN_BRACE_OPEN = '{'
TOKEN_BRACE_CLOSE = '}'
TOKEN_SUB = '_'
TOKEN_SUP = '^'
TOKEN_AMPERSAND = '&'
TOKEN_NEWLINE = '\\\\'


def tokenize(latex):
    """Split LaTeX string into tokens."""
    tokens = []
    i = 0
    while i < len(latex):
        c = latex[i]
        if c == '\\':
            # Check for \begin{...} or \end{...}
            rest = latex[i:]
            env_match = re.match(r'\\(begin|end)\{([^}]+)\}', rest)
            if env_match:
                tokens.append(('CMD', env_match.group(0)[1:]))  # Without leading backslash
                i += len(env_match.group(0))
                continue
            # Regular command: \name
            cmd_match = re.match(r'\\([a-zA-Z]+|.)', rest)
            if cmd_match:
                tokens.append(('CMD', cmd_match.group(1)))
                i += len(cmd_match.group(0))
                continue
            # Just a backslash
            tokens.append(('TEXT', '\\'))
            i += 1
        elif c == '{':
            tokens.append(('{', '{'))
            i += 1
        elif c == '}':
            tokens.append(('}', '}'))
            i += 1
        elif c == '_':
            tokens.append(('_', '_'))
            i += 1
        elif c == '^':
            tokens.append(('^', '^'))
            i += 1
        elif c == '&':
            tokens.append(('&', '&'))
            i += 1
        elif c == '[':
            tokens.append(('[', '['))
            i += 1
        elif c == ']':
            tokens.append((']', ']'))
            i += 1
        elif c in ' \t\n\r':
            i += 1  # skip whitespace between tokens
        else:
            # Collect text run
            j = i
            while j < len(latex) and latex[j] not in '\\{}_^&[]' and latex[j] not in ' \t\n\r':
                j += 1
            tokens.append(('TEXT', latex[i:j]))
            i = j

    return tokens


# ============================================================
# LaTeX → OMML Parser
# ============================================================

class LaTeXParser:
    """Recursive descent parser for LaTeX math → OMML elements."""

    def __init__(self, tokens):
        self.tokens = tokens
        self.pos = 0

    def peek(self):
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        return (None, None)

    def consume(self):
        t = self.peek()
        self.pos += 1
        return t

    def parse(self):
        """Parse the full token stream into OMML elements."""
        elements = self.parse_expression(until_end=True)
        return elements

    def parse_expression(self, until_end=False):
        """Parse a sequence of math elements."""
        elements = []
        while self.pos < len(self.tokens):
            tok_type, tok_val = self.peek()

            if tok_type == '}':
                if not until_end:
                    break
                self.consume()  # Consume unexpected closing brace
                break
            elif tok_type == '{':
                self.consume()  # {
                group_elems = self.parse_expression(until_end=False)
                if self.peek()[0] == '}':
                    self.consume()  # }
                # Wrap group in a math run if it's just text
                if len(group_elems) == 1:
                    elements.extend(group_elems)
                else:
                    elements.extend(group_elems)
            elif tok_type == 'CMD':
                cmd_elems = self.parse_command()
                if cmd_elems:
                    elements.extend(cmd_elems)
            elif tok_type == '_':
                self.consume()
                sub_elems = self.parse_script_content()
                elements = self.attach_script(elements, sub_elems, 'sub')
            elif tok_type == '^':
                self.consume()
                sup_elems = self.parse_script_content()
                elements = self.attach_script(elements, sup_elems, 'sup')
            elif tok_type == '&':
                # Alignment separator in arrays - treat as separator
                self.consume()
                elements.append(make_mr('&'))
            elif tok_type == 'TEXT':
                self.consume()
                elements.append(make_text_run(tok_val))
            else:
                self.consume()
        return elements

    def parse_command(self):
        """Parse a LaTeX command and return OMML elements."""
        tok_type, cmd = self.consume()

        # Greek letters
        if cmd in GREEK_LETTERS:
            return [make_text_run(GREEK_LETTERS[cmd])]

        # Fractions
        if cmd in ('frac', 'dfrac'):
            num = self.parse_braced_group()
            den = self.parse_braced_group()
            return [make_fraction(num, den)]

        # Square root
        if cmd == 'sqrt':
            # Check for optional [n]
            if self.peek()[0] == '[':
                self.consume()  # [
                deg = self.parse_expression()
                if self.peek()[0] == ']':
                    self.consume()  # ]
            else:
                deg = []
            content = self.parse_braced_group()
            return make_radical(deg, content)

        # Accents
        if cmd == 'bar':
            arg = self.parse_braced_or_single()
            return make_accent('̅', arg)  # Combining overline
        if cmd == 'vec':
            arg = self.parse_braced_or_single()
            return make_accent('⃗', arg)  # Combining right arrow above
        if cmd == 'hat':
            arg = self.parse_braced_or_single()
            return make_accent('̂', arg)  # Combining circumflex
        if cmd == 'dot':
            arg = self.parse_braced_or_single()
            return make_accent('̇', arg)  # Combining dot above
        if cmd == 'ddot':
            arg = self.parse_braced_or_single()
            return make_accent('̈', arg)  # Combining diaeresis

        # Summation and products
        if cmd == 'sum':
            return self.parse_nary('∑')  # ∑
        if cmd == 'prod':
            return self.parse_nary('∏')  # ∏
        if cmd == 'int':
            return self.parse_nary('∫')  # ∫
        if cmd == 'oint':
            return self.parse_nary('∮')  # ∮
        if cmd == 'iint':
            return self.parse_nary('∬')  # ∬
        if cmd == 'iiint':
            return self.parse_nary('∭')  # ∭

        # Limits
        if cmd == 'lim':
            return self.parse_nary('lim')

        # Functions
        if cmd in MATH_FUNCTIONS:
            arg = self.parse_braced_or_single()
            return make_function(cmd, arg)

        # \text{...}, \boldsymbol{...}, \mathbf{...}, \mathit{...}, \mathrm{...}
        if cmd in ('text', 'boldsymbol', 'mathbf', 'mathit', 'mathrm', 'textrm'):
            content = self.parse_braced_group()
            text_val = ''
            for e in content:
                mt = e.find(f'{M}t')
                if mt is not None and mt.text:
                    text_val += mt.text
            if cmd in ('boldsymbol', 'mathbf'):
                # Bold math
                mr = _m('r')
                mrPr = _m('rPr')
                sty = _m('sty')
                sty.set(f'{M}val', 'b')
                mrPr.append(sty)
                mr.append(mrPr)
                mt = _m('t')
                mt.text = text_val
                mr.append(mt)
                return [mr]
            return [make_text_run(text_val)]

        # \left, \right delimiters
        if cmd == 'left':
            delim = self.consume()[1] if self.peek()[0] in ('CMD', 'TEXT') else '.'
            if isinstance(delim, str) and len(delim) == 1:
                pass  # Actually need to handle this with matching \right
            # For now, just skip \left and \right — OMML auto-sizes
            return []

        if cmd == 'right':
            # Skip the delimiter
            if self.peek()[0] in ('CMD', 'TEXT'):
                self.consume()
            return []

        # Binary operators and relations
        if cmd == 'cdot':
            return [make_text_run('⋅')]  # ⋅
        if cmd == 'times':
            return [make_text_run('×')]  # ×
        if cmd == 'div':
            return [make_text_run('÷')]  # ÷
        if cmd == 'pm':
            return [make_text_run('±')]  # ±
        if cmd == 'mp':
            return [make_text_run('∓')]  # ∓
        if cmd == 'le' or cmd == 'leq':
            return [make_text_run('≤')]  # ≤
        if cmd == 'ge' or cmd == 'geq':
            return [make_text_run('≥')]  # ≥
        if cmd == 'ne' or cmd == 'neq':
            return [make_text_run('≠')]  # ≠
        if cmd == 'approx':
            return [make_text_run('≈')]  # ≈
        if cmd == 'equiv':
            return [make_text_run('≡')]  # ≡
        if cmd == 'propto':
            return [make_text_run('∝')]  # ∝
        if cmd == 'sim':
            return [make_text_run('∼')]  # ∼
        if cmd == 'simeq':
            return [make_text_run('≃')]  # ≃
        if cmd == 'to' or cmd == 'rightarrow':
            return [make_text_run('→')]  # →
        if cmd == 'leftarrow':
            return [make_text_run('←')]  # ←
        if cmd == 'leftrightarrow':
            return [make_text_run('↔')]  # ↔
        if cmd == 'Rightarrow':
            return [make_text_run('⇒')]  # ⇒
        if cmd == 'Leftrightarrow':
            return [make_text_run('⇔')]  # ⇔
        if cmd == 'infty':
            return [make_text_run('∞')]  # ∞
        if cmd == 'partial':
            return [make_text_run('∂')]  # ∂
        if cmd == 'nabla':
            return [make_text_run('∇')]  # ∇
        if cmd == 'forall':
            return [make_text_run('∀')]  # ∀
        if cmd == 'exists':
            return [make_text_run('∃')]  # ∃
        if cmd == 'neg':
            return [make_text_run('¬')]  # ¬
        if cmd == 'emptyset':
            return [make_text_run('∅')]  # ∅
        if cmd == 'in':
            return [make_text_run('∈')]  # ∈
        if cmd == 'notin':
            return [make_text_run('∉')]  # ∉
        if cmd == 'subset':
            return [make_text_run('⊂')]  # ⊂
        if cmd == 'subseteq':
            return [make_text_run('⊆')]  # ⊆
        if cmd == 'cup':
            return [make_text_run('∪')]  # ∪
        if cmd == 'cap':
            return [make_text_run('∩')]  # ∩
        if cmd == 'angle':
            return [make_text_run('∠')]  # ∠
        if cmd == 'triangle':
            return [make_text_run('△')]  # △
        if cmd == 'perp':
            return [make_text_run('⟂')]  # ⟂
        if cmd == 'parallel':
            return [make_text_run('∥')]  # ∥
        if cmd == 'circ':
            return [make_text_run('∘')]  # ∘
        if cmd == 'cdot':
            return [make_text_run('⋅')]  # ⋅
        if cmd == 'ldots':
            return [make_text_run('…')]  # …
        if cmd == 'cdots':
            return [make_text_run('⋯')]  # ⋯
        if cmd == 'vdots':
            return [make_text_run('⋮')]  # ⋮
        if cmd == 'ddots':
            return [make_text_run('⋱')]  # ⋱
        if cmd == 'prime':
            return [make_text_run('′')]  # ′

        # \begin/\end environments — handled elsewhere, skip here
        if cmd.startswith('begin{') or cmd.startswith('end{'):
            return []

        # LaTeX escape sequences: \% \$ \{ \} \& \# \_ \~ \^
        if cmd == '%':
            return [make_text_run('%')]
        if cmd == '$':
            return [make_text_run('$')]
        if cmd == '{':
            return [make_text_run('{')]
        if cmd == '}':
            return [make_text_run('}')]
        if cmd == '&':
            return [make_text_run('&')]
        if cmd == '#':
            return [make_text_run('#')]
        if cmd == '_':
            return [make_text_run('_')]
        if cmd == '~':
            return [make_text_run('~')]
        if cmd == '^':
            return [make_text_run('^')]

        # Unknown command - keep as text
        return [make_text_run('\\' + cmd)]

    def parse_script_content(self):
        """Parse content for a subscript or superscript."""
        if self.pos >= len(self.tokens):
            return [make_mr('')]

        tok_type, tok_val = self.peek()
        if tok_type == '{':
            self.consume()  # {
            elems = self.parse_expression()
            if self.peek()[0] == '}':
                self.consume()
            return elems if elems else [make_mr('')]
        elif tok_type in ('CMD', 'TEXT'):
            self.consume()
            if tok_type == 'CMD':
                if tok_val in GREEK_LETTERS:
                    return [make_text_run(GREEK_LETTERS[tok_val])]
                return [make_text_run('\\' + tok_val)]
            return [make_text_run(tok_val)]
        else:
            return [make_mr('')]

    def parse_braced_group(self):
        """Parse { ... } and return elements."""
        if self.peek()[0] == '{':
            self.consume()  # {
            elems = self.parse_expression()
            if self.peek()[0] == '}':
                self.consume()
            return elems
        # Single token as group
        return self.parse_expression()

    def parse_braced_or_single(self):
        """Parse {group} or single token and return elements."""
        if self.peek()[0] == '{':
            return self.parse_braced_group()
        elif self.peek()[0] in ('CMD', 'TEXT'):
            tok_type, tok_val = self.consume()
            if tok_type == 'CMD':
                if tok_val in GREEK_LETTERS:
                    return [make_text_run(GREEK_LETTERS[tok_val])]
                return [make_text_run('\\' + tok_val)]
            return [make_text_run(tok_val)]
        return [make_mr('')]

    def parse_nary(self, operator_char):
        """Parse n-ary operator like \sum_{}^{} or \int_{}^{}."""
        sub_elems = []
        sup_elems = []

        if self.peek()[0] == '_':
            self.consume()
            sub_elems = self.parse_script_content()

        if self.peek()[0] == '^':
            self.consume()
            sup_elems = self.parse_script_content()

        # Also handle ^ then _ (reverse order)
        if self.peek()[0] == '_' and not sub_elems:
            self.consume()
            sub_elems = self.parse_script_content()

        return make_nary(operator_char, sub_elems, sup_elems)

    def attach_script(self, elements, script_elems, script_type):
        """Attach subscript or superscript to the last element.
        If the last element is a multi-char text run, split off the last char as the base.
        E.g., 'mc^2' → 'm' + 'c^2'.
        """
        if not elements:
            return script_elems

        # Get the last element as the base
        base_elem = elements.pop()

        # If base is a text run with >1 char, split it: last char → base, rest → keep as text
        if base_elem.tag == f'{M}r':
            mt = base_elem.find(f'{M}t')
            if mt is not None and mt.text and len(mt.text) > 1:
                base_text = mt.text[-1]
                rest_text = mt.text[:-1]
                if rest_text:
                    elements.append(make_text_run(rest_text))
                base = [make_text_run(base_text)]
            else:
                base = [base_elem]
        else:
            base = [base_elem]

        if script_type == 'sub':
            if self.peek()[0] == '^':
                self.consume()
                sup_elems = self.parse_script_content()
                elements.extend(make_subsup(base, script_elems, sup_elems))
            else:
                elements.extend(make_subscript(base, script_elems))
        elif script_type == 'sup':
            if self.peek()[0] == '_':
                self.consume()
                sub_elems = self.parse_script_content()
                elements.extend(make_subsup(base, sub_elems, script_elems))
            else:
                mss, wrapper = make_superscript(base, script_elems)
                elements.extend(wrapper)

        return elements


# ============================================================
# Main Entry Point
# ============================================================

def latex_to_omml_professional(latex: str, is_display: bool = False):
    """
    将 LaTeX 公式转换为 OMML Professional 格式 XML 元素。

    返回 OMML 元素 (<m:oMath> 或 <m:oMathPara>)，可直接插入文档 XML。
    """
    # Pre-process: translate certain LaTeX constructs
    latex = preprocess_latex(latex)

    # Tokenize and parse
    tokens = tokenize(latex)
    parser = LaTeXParser(tokens)
    elements = parser.parse()

    # Wrap in OMML container
    if is_display:
        omath_para = etree.Element(f"{M}oMathPara", nsmap={'m': MATH_NS})
        omath = etree.SubElement(omath_para, f"{M}oMath")
        for elem in elements:
            omath.append(elem)
        return omath_para
    else:
        omath = etree.Element(f"{M}oMath", nsmap={'m': MATH_NS})
        for elem in elements:
            omath.append(elem)
        return omath


def preprocess_latex(latex: str) -> str:
    """Pre-process LaTeX for parsing."""
    # Handle \left( ... \right) -> (...), etc.
    # The parser already handles this, but we can simplify
    # Remove \left and \right since OMML handles sizing automatically
    result = latex
    result = re.sub(r'\\left\s*([({[|.])', r'\1', result)
    result = re.sub(r'\\right\s*([)}\]|.])', r'\1', result)
    # \left\{ -> {
    result = re.sub(r'\\left\\\{', '{', result)
    # \right. -> empty
    result = re.sub(r'\\right\.', '', result)
    return result


# ============================================================
# Test
# ============================================================

if __name__ == '__main__':
    test_formulas = [
        (r'E=mc^2', 'Simple'),
        (r'\frac{a}{b}', 'Fraction'),
        (r'\dfrac{1}{2}mv^2', 'Complex fraction'),
        (r'x_i', 'Subscript'),
        (r'x^2', 'Superscript'),
        (r'x_i^2', 'Sub+Sup'),
        (r'\sqrt{x}', 'Square root'),
        (r'\sqrt[3]{8}', 'Cube root'),
        (r'\sum_{i=1}^{n} x_i', 'Summation'),
        (r'\int_{0}^{\infty} f(x)dx', 'Integral'),
        (r'\alpha + \beta = \gamma', 'Greek letters'),
        (r'\bar{x}', 'Accent bar'),
        (r'\vec{F} = m\vec{a}', 'Accent vec'),
        (r'\frac{1}{2}gt^2', 'Physics'),
        (r'\sin\theta', 'Function'),
        (r'\lim_{x\to\infty} f(x)', 'Limit'),
    ]

    for formula, desc in test_formulas:
        elem = latex_to_omml_professional(formula)
        xml_str = etree.tostring(elem, encoding='unicode')
        print(f'{desc}: {formula}')
        print(f'  OMML: {xml_str[:150]}')
        print()
