from __future__ import annotations

import os
import re
from functools import lru_cache
from pathlib import Path

from latex2mathml.converter import convert
from lxml import etree


_GREEK = {
    "α": r"\alpha ", "β": r"\beta ", "γ": r"\gamma ", "δ": r"\delta ",
    "ε": r"\epsilon ", "θ": r"\theta ", "λ": r"\lambda ", "μ": r"\mu ",
    "π": r"\pi ", "ρ": r"\rho ", "σ": r"\sigma ", "τ": r"\tau ",
    "φ": r"\phi ", "ω": r"\omega ", "Γ": r"\Gamma ", "Δ": r"\Delta ",
    "Θ": r"\Theta ", "Λ": r"\Lambda ", "Π": r"\Pi ", "Σ": r"\Sigma ",
    "Φ": r"\Phi ", "Ω": r"\Omega ",
}
_SYMBOLS = {
    "∑": r"\sum ", "∏": r"\prod ", "∫": r"\int ", "∞": r"\infty ",
    "∂": r"\partial ", "∇": r"\nabla ", "≤": r"\le ", "≥": r"\ge ",
    "≈": r"\approx ", "≠": r"\ne ", "×": r"\times ", "÷": r"\div ",
    "→": r"\to ", "−": "-", "·": r"\cdot ",
}
_SUBSCRIPTS = {
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
}
_SUPERSCRIPTS = {
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
}
_SUBSCRIPT_LETTERS = {
    "ᵢ": "i", "ⱼ": "j", "ₖ": "k", "ₗ": "l", "ₘ": "m", "ₙ": "n",
    "ₚ": "p", "ₛ": "s", "ₜ": "t", "ₐ": "a", "ₑ": "e", "ₕ": "h",
    "ₒ": "o", "ᵣ": "r", "ᵤ": "u", "ᵥ": "v", "ₓ": "x",
}
_SUBSCRIPT_CHARS = {**_SUBSCRIPTS, **_SUBSCRIPT_LETTERS}


@lru_cache(maxsize=1)
def _find_mml2omml_xsl() -> Path | None:
    configured = os.getenv("PDF_MML2OMML_XSL", "").strip()
    candidates = []
    if configured:
        candidates.append(Path(configured))
    candidates.extend(
        [
            Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
            Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
            Path(r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL"),
        ]
    )
    for office_root in (
        Path(r"C:\Program Files\Microsoft Office"),
        Path(r"C:\Program Files (x86)\Microsoft Office"),
    ):
        if office_root.is_dir():
            candidates.extend(office_root.rglob("MML2OMML.XSL"))
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


@lru_cache(maxsize=1)
def _mml2omml_transform():
    xsl_path = _find_mml2omml_xsl()
    if xsl_path is None:
        return None
    return etree.XSLT(etree.parse(str(xsl_path)))


def formula_text_to_latex(text: str) -> str:
    """把常见 Unicode 数学文本归一化为 LaTeX。"""
    value = text.strip()
    for symbol, replacement in {**_GREEK, **_SYMBOLS}.items():
        value = value.replace(symbol, replacement)
    value = re.sub(r"√\s*([A-Za-z0-9]+)", r"\\sqrt{\1}", value)
    output: list[str] = []
    index = 0
    while index < len(value):
        character = value[index]
        if character in _SUBSCRIPT_CHARS:
            end = index
            collected = []
            while end < len(value) and value[end] in _SUBSCRIPT_CHARS:
                collected.append(_SUBSCRIPT_CHARS[value[end]])
                end += 1
            output.append("_{" + "".join(collected) + "}")
            index = end
            continue
        if character in _SUPERSCRIPTS:
            end = index
            collected = []
            while end < len(value) and value[end] in _SUPERSCRIPTS:
                collected.append(_SUPERSCRIPTS[value[end]])
                end += 1
            output.append("^{" + "".join(collected) + "}")
            index = end
            continue
        output.append(character)
        index += 1
    return "".join(output).strip()


def formula_text_to_omml(text: str) -> tuple[bool, str | None, str | None]:
    """把公式文本转换为 Word OMML；失败时返回原因。"""
    latex = formula_text_to_latex(text)
    if not latex:
        return False, None, "empty formula"
    transform = _mml2omml_transform()
    if transform is None:
        return False, None, "MML2OMML.XSL not found"
    try:
        mathml = convert(latex)
        mathml_tree = etree.fromstring(mathml.encode("utf-8"))
        result = transform(mathml_tree)
        omml = str(result)
    except Exception as error:
        return False, None, f"{type(error).__name__}: {error}"
    if "oMath" not in omml:
        return False, None, "OMML result is empty"
    if omml.startswith("<?xml"):
        omml = omml.split("?>", 1)[-1].strip()
    return True, omml, latex