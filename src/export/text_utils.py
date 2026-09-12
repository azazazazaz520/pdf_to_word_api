"""导出过程中的纯文本工具：表格 HTML 解析、代码行识别与文本压缩。

不依赖本包其他模块，也不依赖本目录的其他导出模块。
"""


import re
from collections.abc import Mapping
from typing import Any


_CODE_KEYWORD = re.compile(
    r"^(?:void|int|char|short|long|float|double|unsigned|signed|static|const|"
    r"struct|typedef|enum|union|volatile|if|else|for|while|switch|case|default|"
    r"return|break|continue|goto|sizeof|do)\b"
)


_HTML_TAG = re.compile(r"<[^>]+>")


def _looks_like_code_line(value: str) -> bool:
    """用 C 语言常见语法特征识别代码行，避免把普通正文误判为公式或列表。"""
    text = value.strip()
    if not text:
        return False
    if text.startswith(
        ("#include", "#define", "#if", "#ifdef", "#ifndef", "#else", "#endif",
         "//", "/*", "*/", "*", "*(", "**")
    ):
        return True
    if text.endswith((";", "{", "}")):
        return True
    if _CODE_KEYWORD.match(text):
        return True
    if re.search(r"[A-Za-z_]\w*\s*->\s*[A-Za-z_]\w*", text):
        return True
    if "::" in text:
        return True
    if ";" in text and "=" in text:
        return True
    if re.search(r"0x[0-9A-Fa-f]", text) and (
        text.startswith(
            ("*", "0x", "GPIO", "RCC_", "USART", "TIM", "DMA", "NVIC", "SPI", "I2C")
        )
        or "=" in text
        or ";" in text
        or re.search(r"\((?:unsigned|volatile|uint|int|char)", text)
    ):
        return True
    return False


def _plain_text(value: str) -> str:
    value = _HTML_TAG.sub("", value)
    return re.sub(r"[ \t]+", " ", value).strip()


_FORMULA_PREFIX = re.compile(r"^\s*(?:max|min|s\.\s*t\.)(?:\s|$)", re.IGNORECASE)


_FORMULA_SYMBOL_PREFIX = re.compile(r"^\s*[∑∏∫√∞∂∇∀∃]+")


def _field(value: Any, name: str, default: Any = None) -> Any:
    if isinstance(value, Mapping):
        return value.get(name, default)
    attribute = getattr(value, name, None)
    if attribute is not None:
        return attribute
    aliases = {
        "block_label": "label",
        "block_content": "content",
        "block_bbox": "bbox",
        "block_order": "order_index",
    }
    return getattr(value, aliases.get(name, name), default)
