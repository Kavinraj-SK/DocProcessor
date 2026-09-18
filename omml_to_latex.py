"""
omml_to_latex.py

Converts OOXML Math (OMML, the <m:oMath> element format Word uses for
its Equation Editor) into LaTeX strings suitable for MathJax/KaTeX
rendering on a web frontend.

Handles: text runs, fractions, superscript/subscript (incl. combined),
delimiters (parens/brackets/abs-value with custom begin/end chars),
radicals (sqrt / nth-root), n-ary operators (sum/int/prod with limits),
named functions (sin/cos/log/lim/...), limit expressions, accents
(bar/hat/dot/tilde/vec), box (pass-through), and equation arrays.
Does NOT handle matrices (m:m) -- none were found in the source corpus.
"""

from lxml import etree

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _mtag(el):
    return etree.QName(el).localname


def _find(el, tag):
    return el.find(f"{{{M_NS}}}{tag}")


def _findall(el, tag):
    return el.findall(f"{{{M_NS}}}{tag}")


ACCENT_MAP = {
    "\u0304": "bar", "\u0305": "bar", "\u00af": "bar",
    "\u0307": "dot", "\u0302": "hat", "\u0303": "tilde",
    "\u20d7": "vec", "\u0331": "underline",
}

NARY_CHR_MAP = {
    "\u2211": "\\sum", "\u220f": "\\prod", "\u222b": "\\int",
    "\u222c": "\\iint", "\u222d": "\\iiint", "\u222e": "\\oint",
    "\u22c3": "\\bigcup", "\u22c2": "\\bigcap",
}

KNOWN_FUNCS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "lim",
    "sinh", "cosh", "tanh", "arcsin", "arccos", "arctan", "min", "max",
    "det", "gcd",
}

TEXT_REPLACEMENTS = {
    "\u2212": "-", "\u00d7": "\\times ", "\u00f7": "\\div ",
    "\u2264": "\\leq ", "\u2265": "\\geq ", "\u2260": "\\neq ",
    "\u221e": "\\infty ", "\u03c0": "\\pi ", "\u03b8": "\\theta ",
    "\u03b1": "\\alpha ", "\u03b2": "\\beta ", "\u03b3": "\\gamma ",
    "\u03c3": "\\sigma ", "\u03bc": "\\mu ", "\u03bb": "\\lambda ",
    "\u03b4": "\\delta ", "\u0394": "\\Delta ", "\u03a3": "\\Sigma ",
    "\u03c6": "\\phi ", "\u03c9": "\\omega ", "\u03c1": "\\rho ",
    "\u2192": "\\to ", "\u00b1": "\\pm ", "\u2202": "\\partial ",
    "\u2211": "\\sum ", "\u221a": "\\sqrt ", "\u2208": "\\in ",
    "\u2286": "\\subseteq ", "\u222a": "\\cup ", "\u2229": "\\cap ",
}

DELIM_LATEX = {
    "(": "(", ")": ")", "[": "[", "]": "]",
    "{": "\\{", "}": "\\}", "|": "|", "": "", "\u2016": "\\|",
}


def _escape_text(text):
    if not text:
        return ""
    return "".join(TEXT_REPLACEMENTS.get(ch, ch) for ch in text)


def _wrap(s):
    s = s.strip()
    if len(s) <= 1:
        return s
    return f"{{{s}}}"


def _val(el, tag):
    """Read w:val from a property child, e.g. <m:begChr m:val="|"/>."""
    child = _find(el, tag) if el is not None else None
    if child is None:
        return None
    for attr in (f"{{{M_NS}}}val", f"{{{W_NS}}}val", "val"):
        v = child.get(attr)
        if v is not None:
            return v
    return None


def _convert_children(el):
    if el is None:
        return ""
    return "".join(_convert_node(c) for c in el)


def _convert_node(el):
    tag = _mtag(el)

    if tag in ("oMath", "oMathPara"):
        return "".join(_convert_node(c) for c in el if _mtag(c) != "oMathParaPr")

    if tag == "r":
        text = "".join(t.text or "" for t in _findall(el, "t"))
        return _escape_text(text)

    if tag == "t":
        return _escape_text(el.text or "")

    if tag == "f":
        num = _convert_children(_find(el, "num"))
        den = _convert_children(_find(el, "den"))
        return f"\\frac{{{num}}}{{{den}}}"

    if tag == "sSup":
        base = _convert_children(_find(el, "e"))
        exp = _convert_children(_find(el, "sup"))
        return f"{_wrap(base)}^{{{exp}}}"

    if tag == "sSub":
        base = _convert_children(_find(el, "e"))
        sub = _convert_children(_find(el, "sub"))
        return f"{_wrap(base)}_{{{sub}}}"

    if tag == "sSubSup":
        base = _convert_children(_find(el, "e"))
        sub = _convert_children(_find(el, "sub"))
        sup = _convert_children(_find(el, "sup"))
        return f"{_wrap(base)}_{{{sub}}}^{{{sup}}}"

    if tag == "d":
        dpr = _find(el, "dPr")
        beg = _val(dpr, "begChr")
        end = _val(dpr, "endChr")
        beg = "(" if beg is None else beg
        end = ")" if end is None else end
        parts = [_convert_children(e) for e in _findall(el, "e")]
        inner = ", ".join(parts)
        return f"{DELIM_LATEX.get(beg, beg)}{inner}{DELIM_LATEX.get(end, end)}"

    if tag == "rad":
        radpr = _find(el, "radPr")
        deg_hide_val = _val(radpr, "degHide")
        deg_hidden = deg_hide_val in ("1", "true", "on", None) and deg_hide_val is not None
        degree = _convert_children(_find(el, "deg"))
        radicand = _convert_children(_find(el, "e"))
        if deg_hidden or not degree.strip():
            return f"\\sqrt{{{radicand}}}"
        return f"\\sqrt[{degree}]{{{radicand}}}"

    if tag == "nary":
        narypr = _find(el, "naryPr")
        # OOXML's own default nary symbol (when m:chr is omitted entirely,
        # as it commonly is for a plain integral inserted from Word's
        # basic Integral gallery) is the integral sign, not summation.
        chr_val = _val(narypr, "chr") or "\u222b"
        op = NARY_CHR_MAP.get(chr_val, chr_val)
        sub = _convert_children(_find(el, "sub"))
        sup = _convert_children(_find(el, "sup"))
        body = _convert_children(_find(el, "e"))
        out = op
        if sub.strip():
            out += f"_{{{sub}}}"
        if sup.strip():
            out += f"^{{{sup}}}"
        return f"{out} {body}"

    if tag == "func":
        name = _convert_children(_find(el, "fName")).strip()
        arg = _convert_children(_find(el, "e"))
        base_name = name.split("_")[0].split("^")[0].replace("\\", "")
        if base_name in KNOWN_FUNCS:
            return f"\\{name} {arg}"
        return f"{name} {arg}"

    if tag == "limLow":
        base = _convert_children(_find(el, "e"))
        lim_text = _convert_children(_find(el, "lim"))
        return f"{base}_{{{lim_text}}}"

    if tag == "acc":
        accpr = _find(el, "accPr")
        chr_val = _val(accpr, "chr")
        base = _convert_children(_find(el, "e"))
        cmd = ACCENT_MAP.get(chr_val, "bar")
        return f"\\{cmd}{{{base}}}"

    if tag == "groupChr":
        grouppr = _find(el, "groupChrPr")
        chr_val = _val(grouppr, "chr")
        base = _convert_children(_find(el, "e"))
        # An arrow character here (commonly seen in a hand-built "lim"
        # construct: "Lim" grouped with a "->" arrow, followed by a
        # sibling target value) reads best as "<base> \to " so the
        # following sibling text flows naturally after it -- wrapping it
        # in \overbrace (the old unconditional behavior) was wrong.
        if chr_val in ("\u2192", "\u2190", "\u2194"):
            arrow = {"\u2192": "\\to", "\u2190": "\\leftarrow", "\u2194": "\\leftrightarrow"}[chr_val]
            return f"{base} {arrow} "
        if chr_val in ("\ufe37", "\u23df"):
            return f"\\underbrace{{{base}}}"
        if chr_val in ("\ufe38", "\u23de"):
            return f"\\overbrace{{{base}}}"
        return base

    if tag == "box":
        return _convert_children(_find(el, "e"))

    if tag == "eqArr":
        lines = [_convert_children(e) for e in _findall(el, "e")]
        return "; ".join(l for l in lines if l.strip())

    if tag == "e":
        return _convert_children(el)

    if tag in ("num", "den", "sup", "sub", "deg", "lim", "fName"):
        return _convert_children(el)

    if tag.endswith("Pr"):
        return ""

    return _convert_children(el)


def omml_to_latex(om_element):
    """Convert one <m:oMath> lxml element to a LaTeX string (no $ delimiters)."""
    return _convert_node(om_element).strip()
