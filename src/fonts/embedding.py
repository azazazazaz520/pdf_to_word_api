"""把 PDF 内嵌字体抽取并以 ODTTF 形式嵌入 DOCX。

Word 只能用系统字体渲染，遇到 PDF 子集字体（如 ``BCDEEE+SimSun``）时字形
会与原文不一致。本模块从 PDF 的 ``/FontFile2`` 等流中抽取字体程序，按
ECMA-376 §17.8.1 混淆后写入 ``word/fonts/*.odttf``，并在 ``fontTable.xml``
里登记 ``w:embedRegular``，让 Word 直接用原字体渲染。

需要的 OOXML 结构：

* ``word/fonts/fontN.odttf``：前 32 字节与 fontKey 异或后的字体数据；
* ``fontTable.xml``：``<w:font w:name="..."><w:embedRegular r:id="rIdN"
  w:fontKey="{GUID}" w:subsetted="true"/></w:font>``；
* ``fontTable.xml.rels``：指向字体部件的 ``.../font`` 关系；
* ``settings.xml``：``<w:embedTrueTypeFonts/>``。
"""

from __future__ import annotations

import hashlib
import struct
import uuid
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable

from lxml import etree

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .resolver import (
    normalize_family,
    strip_subset_prefix,
    _descriptor_style,
)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OBFUSCATED_FONT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.obfuscatedFont"
)
_SFNT_VERSIONS = {b"\x00\x01\x00\x00", b"true", b"OTTO", b"ttcf"}
_FONT_STREAM_TAGS = ("/FontFile2", "/FontFile3", "/FontFile")


@dataclass(frozen=True)
class EmbeddedFontProgram:
    """一个可嵌入 DOCX 的 PDF 内嵌字体程序。"""

    raw_name: str
    family: str
    word_name: str
    data: bytes
    guid: str
    bold: bool = False
    italic: bool = False
    source_tag: str = "/FontFile2"
    part_name: str = ""
    relationship_id: str = ""

    @property
    def size(self) -> int:
        return len(self.data)

    @property
    def obfuscated(self) -> bytes:
        return obfuscate_font_data(self.data, self.guid)


@dataclass(frozen=True)
class FontPlanEntry:
    """单个 PDF 字体最终在 Word 中使用的字体。"""

    raw_name: str
    word_name: str
    embedded: bool
    bold: bool = False
    italic: bool = False
    reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "pdf_font": self.raw_name,
            "word_font": self.word_name,
            "embedded": self.embedded,
            "bold": self.bold,
            "italic": self.italic,
            "reason": self.reason,
        }


@dataclass
class FontPlan:
    """PDF 字体 -> Word 字体（含内嵌字体部件）的映射。"""

    entries: dict[str, FontPlanEntry] = field(default_factory=dict)
    programs: list[EmbeddedFontProgram] = field(default_factory=list)

    def entry(self, raw_name: str) -> FontPlanEntry | None:
        if not raw_name:
            return None
        if raw_name in self.entries:
            return self.entries[raw_name]
        stripped = strip_subset_prefix(raw_name)
        return self.entries.get(stripped)

    def word_name_for(self, raw_name: str, fallback: str) -> str:
        entry = self.entry(raw_name)
        if entry is not None and entry.word_name:
            return entry.word_name
        return fallback

    def style_for(
        self,
        raw_name: str,
        bold: bool,
        italic: bool,
    ) -> tuple[bool, bool]:
        entry = self.entry(raw_name)
        if entry is not None and entry.embedded:
            # 内嵌字体程序本身已经是粗体/斜体字形，避免 Word 再次加粗
            return False, False
        return bool(bold), bool(italic)

    def report(self) -> dict[str, Any]:
        unique: dict[tuple[str, str], FontPlanEntry] = {}
        for entry in self.entries.values():
            unique.setdefault((entry.raw_name, entry.word_name), entry)
        entries = list(unique.values())
        embedded = [entry for entry in entries if entry.embedded]
        return {
            "entries": [entry.to_dict() for entry in entries],
            "embedded_font_count": len(self.programs),
            "embedded_font_bytes": sum(program.size for program in self.programs),
            "embedded_families": sorted({entry.word_name for entry in embedded}),
            "fallback_fonts": sorted(
                {
                    entry.raw_name
                    for entry in entries
                    if not entry.embedded
                }
            ),
        }


# ------------------------------------------------------------------ 混淆算法
SYMBOL_FONT_FAMILIES = {
    "wingdings",
    "wingdings 2",
    "wingdings 3",
    "webdings",
    "symbol",
    "zapfdingbats",
    "marlett",
    "bookshelf symbol 7",
}
"""符号字体：Word 会拒绝嵌入的符号子集，统一改用系统字体。"""


def is_symbol_font(family: str, raw_name: str = "") -> bool:
    combined = f"{family} {raw_name}".strip().lower()
    for name in SYMBOL_FONT_FAMILIES:
        if name in combined:
            return True
    return False


_CJK_FAMILY_TOKENS = (
    "simsun",
    "simhei",
    "kaiti",
    "fangsong",
    "yahei",
    "dengxian",
    "song",
    "hei",
    "ming",
    "gothic",
    "宋",
    "黑",
    "楷",
    "仿",
    "雅黑",
    "等线",
)


def guid_bytes(guid: str) -> bytes:
    """把 ``{XXXXXXXX-....}`` GUID 转成 ECMA-376 使用的 16 字节 key。

    实测（用 Word 自己生成的嵌入字体反推）key 是 GUID 十六进制字节的
    **整体反转**，而不是 .NET ``Guid.ToByteArray`` 的 bytes_le 顺序。
    """
    digits = str(guid).strip("{}").replace("-", "")
    return bytes.fromhex(digits)[::-1]


def is_cjk_font(family: str, raw_name: str = "") -> bool:
    combined = f"{family} {raw_name}".lower()
    return any(token in combined for token in _CJK_FAMILY_TOKENS)


def obfuscate_font_data(data: bytes, guid: str) -> bytes:
    """按 ECMA-376 §17.8.1 混淆字体数据（前 32 字节与 key 异或两次）。"""
    key = guid_bytes(guid)
    payload = bytearray(data)
    limit = min(32, len(payload))
    for index in range(limit):
        payload[index] ^= key[index % 16]
    return bytes(payload)


def _looks_like_sfnt(data: bytes) -> bool:
    return bool(data) and data[:4] in _SFNT_VERSIONS


_MEASURE_SIZE_PX = 256.0


@lru_cache(maxsize=512)
def _measure_font(data: bytes) -> Any:
    from io import BytesIO

    from PIL import ImageFont

    try:
        return ImageFont.truetype(BytesIO(data), int(_MEASURE_SIZE_PX))
    except Exception:
        return None


def measure_text_width(data: bytes, text: str, font_size: float) -> float:
    """用字体自身度量文本宽度（单位 point）。"""
    if not data or not text:
        return 0.0
    font = _measure_font(data)
    if font is None:
        return 0.0
    try:
        width_px = font.getlength(text)
    except Exception:
        try:
            width_px = font.getsize(text)[0]
        except Exception:
            return 0.0
    return float(width_px) * float(font_size) / _MEASURE_SIZE_PX


# ------------------------------------------------------------- PDF 字体抽取
def extract_embedded_font_programs(
    source_pdf: Path,
) -> dict[str, EmbeddedFontProgram]:
    """扫描 PDF 页面字体资源，抽取可嵌入的字体程序。"""
    try:
        from pypdf import PdfReader
    except Exception:
        return {}
    try:
        reader = PdfReader(str(source_pdf))
    except Exception:
        return {}
    programs: dict[str, EmbeddedFontProgram] = {}
    seen_raw: set[str] = set()
    for page in reader.pages:
        try:
            resources = page.get("/Resources")
            fonts = resources.get("/Font") if resources is not None else None
        except Exception:
            continue
        if fonts is None:
            continue
        for _, reference in fonts.items():
            try:
                font = reference.get_object()
            except Exception:
                continue
            raw_name = str(font.get("/BaseFont", "") or "")
            if not raw_name or raw_name in seen_raw:
                continue
            descriptor = font.get("/FontDescriptor")
            if descriptor is None:
                descendants = font.get("/DescendantFonts")
                if descendants:
                    try:
                        descriptor = (
                            descendants[0].get_object().get("/FontDescriptor")
                        )
                    except Exception:
                        descriptor = None
            descriptor = _resolve_object(descriptor)
            if not isinstance(descriptor, dict):
                seen_raw.add(raw_name)
                continue
            data = None
            source_tag = ""
            for tag in _FONT_STREAM_TAGS:
                stream = descriptor.get(tag)
                if stream is None:
                    continue
                try:
                    candidate = stream.get_object().get_data()
                except Exception:
                    continue
                if _looks_like_sfnt(candidate):
                    data = candidate
                    source_tag = tag
                    break
            seen_raw.add(raw_name)
            if data is None:
                continue
            family = normalize_family(raw_name)
            weight = _safe_int(descriptor.get("/FontWeight"))
            flags = _safe_int(descriptor.get("/Flags"))
            stem_v = _safe_float(descriptor.get("/StemV"))
            italic_angle = _safe_float(descriptor.get("/ItalicAngle"))
            font_name = str(descriptor.get("/FontName", "") or "")
            bold, italic = _descriptor_style(
                name=f"{raw_name} {font_name}",
                weight=weight,
                flags=flags,
                stem_v=stem_v,
                italic_angle=italic_angle,
            )
            subset = strip_subset_prefix(raw_name)
            tag = ""
            if raw_name.startswith(subset) and raw_name != subset:
                tag = raw_name[: len(raw_name) - len(subset)].strip("+")
            if not tag:
                tag = hashlib.sha1(raw_name.encode("utf-8")).hexdigest()[:6].upper()
            word_name = f"{family} {tag}".strip()
            guid = "{" + str(
                uuid.uuid5(uuid.NAMESPACE_URL, f"pdf-font:{raw_name}")
            ).upper() + "}"
            program = EmbeddedFontProgram(
                raw_name=raw_name,
                family=family,
                word_name=word_name,
                data=data,
                guid=guid,
                bold=bold,
                italic=italic,
                source_tag=source_tag,
            )
            programs[raw_name] = program
            programs[strip_subset_prefix(raw_name)] = program
    return programs


def _resolve_object(value: Any) -> Any:
    """把 pypdf 的 IndirectObject 解析成实际对象。"""
    if value is None:
        return None
    if isinstance(value, dict):
        return value
    resolver = getattr(value, "get_object", None)
    if resolver is not None:
        try:
            return resolver()
        except Exception:
            return value
    return value


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


# ---------------------------------------------------------------- 字体计划
def build_font_plan(
    source_pdf: Path,
    font_requests: Iterable[dict[str, Any]],
    *,
    programs: dict[str, EmbeddedFontProgram] | None = None,
) -> FontPlan:
    """根据 IR 中出现的字体，生成"识别 + 内嵌"计划。"""
    if programs is None:
        programs = extract_embedded_font_programs(source_pdf)
    plan = FontPlan()
    by_digest: dict[str, FontPlanEntry] = {}
    for request in font_requests:
        raw_name = str(request.get("pdf_font_name") or request.get("font_name") or "")
        family = str(request.get("font_name") or "") or normalize_family(raw_name)
        bold = bool(request.get("bold"))
        italic = bool(request.get("italic"))
        if raw_name in plan.entries:
            continue
        program = programs.get(raw_name) or programs.get(
            strip_subset_prefix(raw_name)
        )
        if program is not None and is_symbol_font(family, raw_name):
            # 符号字体的嵌入子集在 Word 中不可用，改用系统字体
            plan.entries[raw_name] = FontPlanEntry(
                raw_name=raw_name,
                word_name=family or raw_name,
                embedded=False,
                bold=bold,
                italic=italic,
                reason="symbol_font_uses_system",
            )
            continue
        if program is not None:
            digest = hashlib.sha1(program.data).hexdigest()
            entry = by_digest.get(digest)
            if entry is None:
                entry = FontPlanEntry(
                    raw_name=raw_name,
                    word_name=program.word_name,
                    embedded=True,
                )
                by_digest[digest] = entry
                plan.programs.append(program)
            plan.entries[raw_name] = entry
            plan.entries[strip_subset_prefix(raw_name)] = entry
            continue
        plan.entries[raw_name] = FontPlanEntry(
            raw_name=raw_name,
            word_name=family or raw_name,
            embedded=False,
            bold=bold,
            italic=italic,
            reason="font_not_embedded_in_pdf",
        )
    return plan


# ------------------------------------------------------------- 写入 DOCX
def _font_table_part(document: Any) -> Any:
    return document.part.part_related_by(RT.FONT_TABLE)


def _settings_part(document: Any) -> Any:
    return document.part.part_related_by(RT.SETTINGS)


def attach_embedded_fonts(
    document: Any,
    programs: list[EmbeddedFontProgram],
) -> dict[str, str]:
    """把字体部件写入 DOCX，返回 ``word_name -> rId``。"""
    if not programs:
        return {}
    package = document.part.package
    font_table = _font_table_part(document)
    table_root = etree.fromstring(font_table.blob)
    index = 0
    relationship_ids: dict[str, str] = {}
    for program in programs:
        index += 1
        part_name = f"/word/fonts/font{index}.odttf"
        part = Part(
            PackURI(part_name),
            OBFUSCATED_FONT_CONTENT_TYPE,
            program.obfuscated,
            package,
        )
        relationship_id = font_table.relate_to(part, RT.FONT)
        relationship_ids[program.word_name] = relationship_id
        font_element = _ensure_font_element(
            table_root,
            program.word_name,
            cjk=is_cjk_font(program.family, program.raw_name),
        )
        embed = OxmlElement("w:embedRegular")
        embed.set(qn("r:id"), relationship_id)
        embed.set(qn("w:fontKey"), program.guid)
        embed.set(qn("w:subsetted"), "true")
        font_element.append(embed)
    font_table._blob = etree.tostring(
        table_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    _enable_embedded_fonts(document)
    return relationship_ids


def _ensure_font_element(
    table_root: Any,
    word_name: str,
    *,
    cjk: bool = False,
) -> Any:
    for element in table_root.findall(qn("w:font")):
        if element.get(qn("w:name")) == word_name:
            return element
    element = OxmlElement("w:font")
    element.set(qn("w:name"), word_name)
    charset = OxmlElement("w:charset")
    charset.set(qn("w:val"), "86" if cjk else "00")
    element.append(charset)
    family = OxmlElement("w:family")
    family.set(qn("w:val"), "auto")
    element.append(family)
    pitch = OxmlElement("w:pitch")
    pitch.set(qn("w:val"), "fixed" if cjk else "variable")
    element.append(pitch)
    table_root.append(element)
    return element


def _enable_embedded_fonts(document: Any) -> None:
    """在 settings.xml 打开嵌入字体开关（SettingsPart 是 XmlPart，需改 element）。"""
    settings = _settings_part(document)
    if settings is None:
        return
    root = getattr(settings, "element", None)
    if root is None:
        return
    if root.find(qn("w:embedTrueTypeFonts")) is None:
        element = OxmlElement("w:embedTrueTypeFonts")
        anchor = root.find(qn("w:proofState"))
        if anchor is None:
            anchor = root.find(qn("w:compat"))
        if anchor is not None:
            anchor.addprevious(element)
        else:
            root.append(element)
    if root.find(qn("w:saveSubsetFonts")) is None:
        element = OxmlElement("w:saveSubsetFonts")
        anchor = root.find(qn("w:embedTrueTypeFonts"))
        if anchor is not None:
            anchor.addnext(element)
        else:
            root.append(element)
