"""PDF 字体识别与系统字体匹配。

目标：把 PDF 里的字体（含子集字体、CID 字体、Type1 标准字体）尽量
准确地还原成 Word 可用的字体族 + 粗体/斜体，并报告替换情况。

识别依据按优先级：

1. PDF 字体描述符（``/FontDescriptor``）：``/FontWeight``、``/Flags``、
   ``/ItalicAngle``、``/StemV``、``/FontName``；
2. 字体名称语义：``TimesNewRomanPS-BoldMT``、``ABCDEF+SimSun``、
   ``MicrosoftYaHei-0`` 等；
3. pdfium 解析出的 family name（对 CID 子集字体尤其有用）。

最终输出统一的字体族名（Word 字体名）和 bold/italic 标记，并在字体
未安装时给出替换后的字体和原因。
"""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
import re
from typing import Any, Iterable


_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")
_TRAILING_INDEX = re.compile(r"-\d+$")
_NON_NAME = re.compile(r"[^0-9a-z\u4e00-\u9fff]+")
_CJK_RANGE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")


@dataclass(frozen=True)
class PdfFontDescriptor:
    """PDF 字体描述符中与识别相关的信息。"""

    base_font: str
    family: str = ""
    subtype: str = ""
    bold: bool = False
    italic: bool = False
    weight: int | None = None
    flags: int | None = None
    stem_v: float | None = None
    italic_angle: float | None = None
    embedded: bool = False

    @property
    def has_descriptor(self) -> bool:
        return any(
            value is not None
            for value in (self.weight, self.flags, self.stem_v, self.italic_angle)
        )


@dataclass(frozen=True)
class FontMatch:
    """一次字体识别的结果。"""

    requested: str
    family: str
    bold: bool = False
    italic: bool = False
    matched: bool = False
    substituted: bool = False
    fallback_reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "requested": self.requested,
            "family": self.family,
            "bold": self.bold,
            "italic": self.italic,
            "matched": self.matched,
            "substituted": self.substituted,
            "fallback_reason": self.fallback_reason,
        }


# --------------------------------------------------------------------- 名称处理
_ALIASES: dict[str, str] = {
    # 中文
    "simsun": "SimSun",
    "nsimsun": "SimSun",
    "songti": "SimSun",
    "宋体": "SimSun",
    "simhei": "SimHei",
    "heiti": "SimHei",
    "黑体": "SimHei",
    "kaiti": "KaiTi",
    "stkaiti": "KaiTi",
    "楷体": "KaiTi",
    "fangsong": "FangSong",
    "stfangsong": "FangSong",
    "仿宋": "FangSong",
    "microsoftyahei": "Microsoft YaHei",
    "microsoftyaheimono": "Microsoft YaHei",
    "msyh": "Microsoft YaHei",
    "微软雅黑": "Microsoft YaHei",
    "dengxian": "DengXian",
    "dengxianlight": "DengXian Light",
    "等线": "DengXian",
    "yuanti": "Microsoft YaHei",
    "lixu": "LiSu",
    "lisu": "LiSu",
    "youyuan": "YouYuan",
    "msmincho": "MS Mincho",
    "msgothic": "MS Gothic",
    "malgungothic": "Malgun Gothic",
    # 拉丁
    "timesnewroman": "Times New Roman",
    "timesnewromanps": "Times New Roman",
    "timesnewromanpsmt": "Times New Roman",
    "timesnewromanpsboldmt": "Times New Roman",
    "timesnewromanpsitalicmt": "Times New Roman",
    "timesnewromanpsbolditalicmt": "Times New Roman",
    "times": "Times New Roman",
    "arial": "Arial",
    "arialmt": "Arial",
    "arialboldmt": "Arial",
    "arialitalicmt": "Arial",
    "arialbolditalicmt": "Arial",
    "arialunicode": "Arial Unicode MS",
    "arialunicodems": "Arial Unicode MS",
    "helvetica": "Arial",
    "helveticaneue": "Arial",
    "helveticabold": "Arial",
    "helveticaoblique": "Arial",
    "helveticaboldoblique": "Arial",
    "couriernew": "Courier New",
    "couriernewpsmt": "Courier New",
    "couriernewboldmt": "Courier New",
    "couriernewitalicmt": "Courier New",
    "courier": "Courier New",
    "cambria": "Cambria",
    "cambriamath": "Cambria Math",
    "calibri": "Calibri",
    "calibrilight": "Calibri Light",
    "verdana": "Verdana",
    "tahoma": "Tahoma",
    "georgia": "Georgia",
    "consolas": "Consolas",
    "consola": "Consolas",
    "segoeui": "Segoe UI",
    "segoeuibold": "Segoe UI",
    "symbol": "Symbol",
    "wingdings": "Wingdings",
    "wingdings2": "Wingdings 2",
    "wingdings3": "Wingdings 3",
    "zapfdingbats": "Wingdings",
    "dejavusans": "Arial",
    "dejavuserif": "Times New Roman",
    "liberationserif": "Times New Roman",
    "liberationsans": "Arial",
    "nimbusroman": "Times New Roman",
    "nimbussans": "Arial",
    "cmr": "Times New Roman",
    "cmsy": "Symbol",
    "cmmi": "Cambria Math",
}

_BOLD_TOKENS = ("bold", "black", "heavy", "semibold", "demibold", "extrabold")
_ITALIC_TOKENS = ("italic", "oblique", "kursiv")


def strip_subset_prefix(name: str) -> str:
    """去掉 ``ABCDEF+`` 子集前缀和 ``-0`` 之类的序号后缀。"""
    value = str(name or "").strip()
    if value.startswith("/"):
        value = value[1:]
    value = _SUBSET_PREFIX.sub("", value)
    value = _TRAILING_INDEX.sub("", value)
    return value


def _compact(value: str) -> str:
    return _NON_NAME.sub("", strip_subset_prefix(value).lower())


def _style_from_name(name: str) -> tuple[bool, bool]:
    lowered = strip_subset_prefix(name).lower()
    bold = any(token in lowered for token in _BOLD_TOKENS)
    italic = any(token in lowered for token in _ITALIC_TOKENS)
    return bold, italic


def normalize_family(name: str) -> str:
    """把任意 PDF 字体名归一化成字体族名。"""
    cleaned = strip_subset_prefix(name)
    if not cleaned:
        return ""
    compact = _compact(cleaned)
    if not compact:
        return cleaned
    # 依次尝试：完整压缩名、去掉常见后缀、包含匹配
    if compact in _ALIASES:
        return _ALIASES[compact]
    base = compact
    for suffix in (
        "bolditalic", "boldoblique", "italic", "oblique", "bold",
        "regular", "roman", "mt", "ps", "psmt", "std", "w3", "0",
    ):
        if base.endswith(suffix) and len(base) > len(suffix) + 1:
            base = base[: -len(suffix)]
            if base in _ALIASES:
                return _ALIASES[base]
    if base in _ALIASES:
        return _ALIASES[base]
    for key, family in _ALIASES.items():
        if len(key) >= 4 and key in compact:
            return family
    # 去掉样式词后返回可读名称
    readable = re.sub(
        r"[-,]?\s*(bold|italic|oblique|regular|roman|semibold|light|mt|ps)",
        " ",
        cleaned,
        flags=re.IGNORECASE,
    )
    readable = re.sub(r"\s+", " ", readable).strip(" -,")
    return readable or cleaned


# ----------------------------------------------------------------- 描述符解析
def _descriptor_style(
    *,
    name: str = "",
    weight: int | None = None,
    flags: int | None = None,
    stem_v: float | None = None,
    italic_angle: float | None = None,
) -> tuple[bool, bool]:
    name_bold, name_italic = _style_from_name(name)
    bold = name_bold
    italic = name_italic
    if weight is not None and weight >= 600:
        bold = True
    if flags is not None:
        if flags & (1 << 18):  # ForceBold
            bold = True
        if flags & (1 << 6):  # Italic
            italic = True
    if italic_angle is not None and abs(float(italic_angle)) > 4.0:
        italic = True
    # StemV 只对 TrueType/Type1 有意义；CID 字体常填 1000 之类的占位值。
    if stem_v is not None and weight is None and name_bold is False:
        value = float(stem_v)
        if 0 < value < 400 and value >= 140:
            bold = True
    return bold, italic


def descriptor_from_mapping(
    base_font: str,
    mapping: dict[str, Any] | None,
) -> PdfFontDescriptor:
    mapping = mapping or {}
    font_name = str(mapping.get("font_name") or "")
    bold, italic = _descriptor_style(
        name=f"{base_font} {font_name}",
        weight=mapping.get("weight"),
        flags=mapping.get("flags"),
        stem_v=mapping.get("stem_v"),
        italic_angle=mapping.get("italic_angle"),
    )
    family = normalize_family(font_name or base_font)
    return PdfFontDescriptor(
        base_font=strip_subset_prefix(base_font),
        family=family,
        subtype=str(mapping.get("subtype") or ""),
        bold=bold,
        italic=italic,
        weight=mapping.get("weight"),
        flags=mapping.get("flags"),
        stem_v=mapping.get("stem_v"),
        italic_angle=mapping.get("italic_angle"),
        embedded=bool(mapping.get("embedded", True)),
    )


def extract_pdf_font_descriptors(
    source_pdf: Path,
) -> dict[str, PdfFontDescriptor]:
    """扫描 PDF 资源字典，按 BaseFont 名称收集字体描述符。"""
    try:
        from pypdf import PdfReader
    except Exception:
        return {}
    descriptors: dict[str, PdfFontDescriptor] = {}
    try:
        reader = PdfReader(str(source_pdf))
    except Exception:
        return descriptors
    for page in reader.pages:
        try:
            resources = page.get("/Resources")
            if resources is None:
                continue
            fonts = resources.get("/Font")
            if fonts is None:
                continue
        except Exception:
            continue
        for _, reference in fonts.items():
            try:
                font = reference.get_object()
            except Exception:
                continue
            try:
                base_font = str(font.get("/BaseFont", "") or "")
                subtype = str(font.get("/Subtype", "") or "")
                descriptor = font.get("/FontDescriptor")
                if descriptor is None:
                    descendants = font.get("/DescendantFonts")
                    if descendants:
                        descriptor = (
                            descendants[0]
                            .get_object()
                            .get("/FontDescriptor")
                        )
                if descriptor is not None:
                    descriptor = descriptor.get_object()
                payload: dict[str, Any] = {
                    "subtype": subtype,
                    "embedded": descriptor is not None,
                }
                if isinstance(descriptor, dict):
                    payload.update(
                        {
                            "font_name": str(
                                descriptor.get("/FontName", "") or ""
                            ),
                            "flags": _safe_int(descriptor.get("/Flags")),
                            "weight": _safe_int(descriptor.get("/FontWeight")),
                            "stem_v": _safe_float(descriptor.get("/StemV")),
                            "italic_angle": _safe_float(
                                descriptor.get("/ItalicAngle")
                            ),
                        }
                    )
                info = descriptor_from_mapping(base_font, payload)
            except Exception:
                continue
            for key in {base_font, strip_subset_prefix(base_font)}:
                if key and key not in descriptors:
                    descriptors[key] = info
    return descriptors


def _safe_int(value: Any) -> int | None:
    try:
        return int(value)
    except Exception:
        return None


def _safe_float(value: Any) -> float | None:
    try:
        return float(value)
    except Exception:
        return None


# ------------------------------------------------------------- 系统字体匹配
@lru_cache(maxsize=1)
def installed_font_names() -> frozenset[str]:
    """返回本机已安装字体族名（小写、去空格）。"""
    names: set[str] = set()

    def add(value: str) -> None:
        cleaned = re.sub(
            r"\s*\((?:TrueType|OpenType|Type ?1|All res)\)\s*$",
            "",
            str(value or ""),
            flags=re.IGNORECASE,
        )
        # 注册表里常见 "Cambria & Cambria Math (TrueType)" 这类组合名
        for part in re.split(r"[&,]", cleaned):
            part = part.strip()
            if not part:
                continue
            names.add(re.sub(r"\s+", "", part).lower())
            names.add(re.sub(r"\s+", " ", part).strip().lower())

    try:
        import winreg  # type: ignore

        for root, subkey in (
            (
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
            ),
            (
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
            ),
        ):
            try:
                with winreg.OpenKey(root, subkey) as key:
                    index = 0
                    while True:
                        try:
                            value_name, _, _ = winreg.EnumValue(key, index)
                        except OSError:
                            break
                        add(value_name)
                        index += 1
            except OSError:
                continue
    except Exception:
        pass
    fonts_dir = Path(r"C:\Windows\Fonts")
    if fonts_dir.is_dir():
        for path in fonts_dir.glob("*"):
            if path.suffix.lower() in {".ttf", ".ttc", ".otf", ".fon"}:
                add(path.stem)
    return frozenset(names)


def _is_installed(family: str) -> bool:
    if not family:
        return False
    installed = installed_font_names()
    compact = re.sub(r"\s+", "", family).lower()
    spaced = re.sub(r"\s+", " ", family).strip().lower()
    return compact in installed or spaced in installed


_CJK_FALLBACKS = ("Microsoft YaHei", "SimSun", "SimHei", "DengXian")
_LATIN_FALLBACKS = ("Times New Roman", "Arial", "Calibri")


def resolve_font(
    *,
    raw_name: str = "",
    family_hint: str = "",
    descriptor: PdfFontDescriptor | None = None,
    text: str = "",
    font_weight: int | None = None,
    font_flags: int | None = None,
    italic_angle: float | None = None,
) -> FontMatch:
    """识别字体并映射到本机字体，返回替换信息。"""
    requested = strip_subset_prefix(raw_name or family_hint or "")
    family = normalize_family(family_hint or raw_name)
    if not family and descriptor is not None:
        family = descriptor.family
    if not family:
        family = normalize_family(raw_name) or raw_name or "Arial"

    bold, italic = _style_from_name(f"{raw_name} {family_hint}")
    if descriptor is not None:
        descriptor_bold, descriptor_italic = _descriptor_style(
            name=f"{descriptor.base_font} {descriptor.family}",
            weight=descriptor.weight,
            flags=descriptor.flags,
            stem_v=descriptor.stem_v,
            italic_angle=descriptor.italic_angle,
        )
        bold = bold or descriptor.bold or descriptor_bold
        italic = italic or descriptor.italic or descriptor_italic
    if font_weight is not None and font_weight >= 600:
        bold = True
    if font_flags is not None:
        if font_flags & (1 << 18):
            bold = True
        if font_flags & (1 << 6):
            italic = True
    if italic_angle is not None and abs(float(italic_angle)) > 4.0:
        italic = True

    if _is_installed(family):
        return FontMatch(
            requested=requested or family,
            family=family,
            bold=bold,
            italic=italic,
            matched=True,
        )

    cjk = bool(_CJK_RANGE.search(text)) or bool(_CJK_RANGE.search(family))
    chain = _CJK_FALLBACKS if cjk else _LATIN_FALLBACKS
    for candidate in chain:
        if _is_installed(candidate):
            return FontMatch(
                requested=requested or family,
                family=candidate,
                bold=bold,
                italic=italic,
                matched=False,
                substituted=True,
                fallback_reason=f"font_not_installed:{family}",
            )
    return FontMatch(
        requested=requested or family,
        family=family,
        bold=bold,
        italic=italic,
        matched=False,
        substituted=True,
        fallback_reason=f"no_fallback_font:{family}",
    )


def style_key(
    font_name: str,
    font_size: float,
    color: tuple[int, int, int] | None,
    bold: bool,
    italic: bool,
    rotation: float = 0.0,
) -> tuple[Any, ...]:
    """用于把同一行的字符切分成 font run 的比较键。"""
    return (
        font_name,
        round(float(font_size) * 4) / 4,
        tuple(color) if color else None,
        bool(bold),
        bool(italic),
        round(float(rotation) / 2.0) * 2.0,
    )


def summarize_font_usage(entries: Iterable[dict[str, Any]]) -> dict[str, Any]:
    """把逐行/逐 span 的字体信息聚合成字体使用报告。"""
    usage: dict[tuple[str, str, bool, bool], dict[str, Any]] = {}
    substituted: dict[str, dict[str, Any]] = {}
    for entry in entries:
        requested = str(entry.get("pdf_font_name") or entry.get("font_name") or "")
        family = str(entry.get("font_name") or "")
        bold = bool(entry.get("bold"))
        italic = bool(entry.get("italic"))
        key = (requested, family, bold, italic)
        record = usage.setdefault(
            key,
            {
                "pdf_font": requested,
                "word_font": family,
                "bold": bold,
                "italic": italic,
                "count": 0,
            },
        )
        record["count"] += 1
        if entry.get("font_substituted"):
            info = substituted.setdefault(
                requested or family,
                {
                    "pdf_font": requested,
                    "word_font": family,
                    "reason": entry.get("font_fallback_reason", ""),
                    "count": 0,
                },
            )
            info["count"] += 1
    return {
        "usage": sorted(
            usage.values(), key=lambda item: item["count"], reverse=True
        ),
        "substituted": sorted(
            substituted.values(), key=lambda item: item["count"], reverse=True
        ),
        "substituted_font_count": len(substituted),
        "resolved_font_count": len(usage),
    }
