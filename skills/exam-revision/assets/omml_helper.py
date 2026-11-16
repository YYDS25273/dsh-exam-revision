# -*- coding: utf-8 -*-
"""极简 OMML 生成器：把 $...$ 中的公式渲染成 Word 原生公式（可选中、可编辑）。"""

from docx.oxml import parse_xml

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SYM = {
    "times": "×", "cdot": "·", "div": "÷", "pm": "±",
    "le": "≤", "ge": "≥", "ne": "≠", "approx": "≈", "equiv": "≡",
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "Delta": "Δ", "lambda": "λ", "mu": "μ", "phi": "φ", "varphi": "φ",
    "omega": "ω", "Omega": "Ω", "rho": "ρ", "sigma": "σ", "eta": "η",
    "Sigma": "Σ", "Gamma": "Γ", "Phi": "Φ", "Theta": "Θ", "Lambda": "Λ",
    "infty": "∞", "sum": "∑", "prime": "′", "circ": "°",
    "sim": "～", "approx": "≈", "leq": "≤", "geq": "≥", "neq": "≠",
    "propto": "∝", "equiv": "≡", "mp": "∓", "ll": "≪", "gg": "≫",
    "cdots": "⋯", "ldots": "…", "vdots": "⋮",
    "rightarrow": "→", "to": "→", "Rightarrow": "⇒",
}

_MATH_FONT = ('<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
              '</w:rPr>')


def _esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _has_cjk(t):
    return any("\u4e00" <= ch <= "\u9fff" for ch in t)


def _run(text):
    if not text:
        return ""
    nor = "<m:nor/>" if _has_cjk(text) else ""
    return ('<m:r><m:rPr>%s</m:rPr>%s<m:t xml:space="preserve">%s</m:t></m:r>'
            % (nor, _MATH_FONT, _esc(text)))


def _parse_arg(s, i):
    """取一个参数：{...} 整组，或命令，或**单个**字符（与 LaTeX 的 x^2 规则一致）。"""
    if i < len(s) and s[i] == "{":
        seq, i = _parse_seq(s, i + 1, "}")
        return seq, i
    if i >= len(s):
        return [], i
    if s[i] == "\\":
        atom, i = _parse_atom(s, i)
        return [atom], i
    return [("r", s[i])], i + 1


def _parse_atom(s, i):
    c = s[i]
    if c == "{":
        seq, i = _parse_seq(s, i + 1, "}")
        return ("seq", seq), i
    if c == "\\":
        j = i + 1
        while j < len(s) and s[j].isalpha():
            j += 1
        name = s[i + 1:j]
        if name == "frac":
            a, i = _parse_arg(s, j)
            b, i = _parse_arg(s, i)
            return ("frac", a, b), i
        if name == "sqrt":
            a, i = _parse_arg(s, j)
            return ("sqrt", a), i
        return ("r", SYM.get(name, name)), j
    j = i
    while j < len(s) and s[j] not in "\\{}_^":
        j += 1
    if j == i:
        return ("r", s[i]), i + 1
    return ("r", s[i:j]), j


def _parse_seq(s, i, closer):
    out = []
    while i < len(s):
        if closer is not None and s[i] == closer:
            i += 1
            break
        atom, i = _parse_atom(s, i)
        sub = sup = None
        while i < len(s) and s[i] in "_^":
            if s[i] == "_":
                sub, i = _parse_arg(s, i + 1)
            else:
                sup, i = _parse_arg(s, i + 1)
        if sub is not None or sup is not None:
            atom = ("subsup", atom, sub, sup)
        out.append(atom)
    return out, i


def _render(seq):
    return "".join(_render_atom(a) for a in seq)


def _render_atom(a):
    k = a[0]
    if k == "r":
        return _run(a[1])
    if k == "seq":
        return _render(a[1])
    if k == "frac":
        return ('<m:f><m:fPr><m:ctrlPr>%s</m:ctrlPr></m:fPr>'
                '<m:num>%s</m:num><m:den>%s</m:den></m:f>'
                % (_MATH_FONT, _render(a[1]), _render(a[2])))
    if k == "sqrt":
        return ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
                '<m:e>%s</m:e></m:rad>' % _render(a[1]))
    if k == "subsup":
        _, base, sub, sup = a
        b = _render([base])
        if sub is not None and sup is not None:
            return ('<m:sSubSup><m:e>%s</m:e><m:sub>%s</m:sub><m:sup>%s</m:sup>'
                    '</m:sSubSup>' % (b, _render(sub), _render(sup)))
        if sub is not None:
            return ('<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>'
                    % (b, _render(sub)))
        return ('<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>'
                % (b, _render(sup)))
    return ""


def omath(expr):
    """返回 <m:oMath> 元素，用于插入到段落。"""
    seq, _ = _parse_seq(expr, 0, None)
    xml = '<m:oMath xmlns:m="%s" xmlns:w="%s">%s</m:oMath>' % (M, W, _render(seq))
    return parse_xml(xml)


def omath_para(expr, align="center"):
    """返回 <m:oMathPara> 元素，用于独立成行的居中显示公式。"""
    seq, _ = _parse_seq(expr, 0, None)
    jc = {'center': 'center', 'left': 'left', 'right': 'right'}[align]
    xml = ('<m:oMathPara xmlns:m="%s" xmlns:w="%s">'
           '<m:oMathParaPr><m:jc m:val="%s"/></m:oMathParaPr><m:oMath>%s</m:oMath>'
           '</m:oMathPara>' % (M, W, jc, _render(seq)))
    return parse_xml(xml)
