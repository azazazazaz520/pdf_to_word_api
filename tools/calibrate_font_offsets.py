"""用 Microsoft Word 实测各字体的 framePr 定位偏移，生成字体度量表。

用法：

    python tools/calibrate_font_offsets.py [--output src/font_metrics_table.json]

原理：把同一段文本以不同字体/字号写成 `w:framePr` 绝对定位段落，
用 Word 导出 PDF，再用 pdfium 读取实际字形框顶部，得到

    k(font, size) = (glyph_top - frame_top) / font_size

导出时用 `frame_y = source_glyph_top - k * font_size` 即可对齐。
"""

from __future__ import annotations

import argparse
import json
import math
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pypdf import PdfReader

FONTS = [
    "Arial",
    "Times New Roman",
    "Courier New",
    "Calibri",
    "Cambria",
    "SimSun",
    "SimHei",
    "Microsoft YaHei",
    "DengXian",
    "KaiTi",
    "FangSong",
    "Consolas",
    "Georgia",
    "Verdana",
]
SIZES = [8.0, 9.0, 10.5, 12.0, 14.0, 16.0, 20.0, 24.0]
SAMPLE = "字体偏移标定 FontOffset 0123456789"

TWIPS_PER_POINT = 20


def _frame_paragraph(document, text: str, *, x: float, y: float, width: float, font: str, size: float):
    paragraph = document.add_paragraph()
    properties = paragraph._p.get_or_add_pPr()
    frame = OxmlElement("w:framePr")
    frame.set(qn("w:wrap"), "none")
    frame.set(qn("w:hAnchor"), "page")
    frame.set(qn("w:vAnchor"), "page")
    frame.set(qn("w:x"), str(int(round(x * TWIPS_PER_POINT))))
    frame.set(qn("w:y"), str(int(round(y * TWIPS_PER_POINT))))
    frame.set(qn("w:w"), str(int(round(width * TWIPS_PER_POINT))))
    frame.set(qn("w:hRule"), "auto")
    properties.insert(0, frame)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(size)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run_properties = run._element.get_or_add_rPr()
    fonts = run_properties.get_or_add_rFonts()
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attribute), font)
    return paragraph


def build_calibration_docx(path: Path) -> list[dict]:
    document = Document()
    section = document.sections[0]
    section.page_width = Pt(595)
    section.page_height = Pt(842)
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, margin, Pt(0))
    entries: list[dict] = []
    index = 0
    row_y = 20.0
    column_x = 20.0
    for font in FONTS:
        for size in SIZES:
            if row_y > 800.0:
                document.add_section()
                new_section = document.sections[-1]
                new_section.page_width = Pt(595)
                new_section.page_height = Pt(842)
                for margin in (
                    "top_margin",
                    "bottom_margin",
                    "left_margin",
                    "right_margin",
                ):
                    setattr(new_section, margin, Pt(0))
                row_y = 20.0
                column_x = 20.0
            marker = f"{index:03d}"
            sample = f"{marker} 偏移标定 FontOffset 0123456789"
            entry = {
                "index": index,
                "marker": marker,
                "font": font,
                "size": size,
                "y": row_y,
                "x": column_x,
                "text": sample,
                "page": len(document.sections) - 1,
            }
            _frame_paragraph(
                document,
                sample,
                x=column_x,
                y=row_y,
                width=540.0,
                font=font,
                size=size,
            )
            entries.append(entry)
            index += 1
            row_y += max(size * 1.6, 22.0) + 6.0
    document.save(str(path))
    return entries


def render_docx(docx_path: Path, pdf_path: Path) -> None:
    script = (
        "$ErrorActionPreference='Stop';"
        "$word=New-Object -ComObject Word.Application;"
        "$word.Visible=$false; $word.DisplayAlerts=0;"
        "try{"
        f"$doc=$word.Documents.Open('{docx_path}', $false, $true);"
        f"$doc.ExportAsFixedFormat('{pdf_path}', 17);"
        "$doc.Close($false)"
        "}finally{$word.Quit()}"
    )
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=600,
        check=False,
    )
    if completed.returncode != 0 or not pdf_path.is_file():
        raise RuntimeError(f"Word 渲染失败：{completed.stderr or completed.stdout}")


def measure(pdf_path: Path) -> list[tuple[float, float]]:
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(pdf_path))
    results: list[tuple[float, float]] = []
    try:
        for page_index in range(len(document)):
            page = document[page_index]
            text_page = page.get_textpage()
            try:
                text = text_page.get_text_range()
                tops: dict[float, float] = {}
                for index, character in enumerate(text):
                    if character in "\r\n":
                        continue
                    x0, y0, x1, y1 = text_page.get_charbox(index)
                    if x1 - x0 <= 0 and y1 - y0 <= 0:
                        continue
                    key = round(max(y1) / 6.0)
                    top = float(min(y0, y1))
                    if key not in tops or top > tops[key]:
                        tops[key] = top
                for key in sorted(tops):
                    results.append((key * 6.0, tops[key]))
            finally:
                text_page.close()
                page.close()
    finally:
        document.close()
    return results


def measure_entries(pdf_path: Path, entries: list[dict]) -> dict[str, float]:
    """测量每个条目的字形框顶部，返回 font -> 偏移系数。"""
    import pypdfium2 as pdfium

    by_marker = {entry["marker"]: entry for entry in entries}
    samples: dict[str, list[float]] = {}
    document = pdfium.PdfDocument(str(pdf_path))
    try:
        for page_index in range(len(document)):
            page = document[page_index]
            text_page = page.get_textpage()
            try:
                text = text_page.get_text_range()
                characters: list[tuple[str, float, float, float, float]] = []
                for index, character in enumerate(text):
                    if character in "\r\n":
                        continue
                    x0, y0, x1, y1 = text_page.get_charbox(index)
                    if x1 - x0 <= 0 and y1 - y0 <= 0:
                        continue
                    characters.append(
                        (character, float(x0), float(y0), float(x1), float(y1))
                    )
            finally:
                text_page.close()
                page.close()

            # 按 y 聚合成行
            lines: list[list[tuple[str, float, float, float, float]]] = []
            for item in sorted(characters, key=lambda value: (-value[3], value[1])):
                placed = False
                for line in lines:
                    reference = line[0]
                    if abs(item[3] - reference[3]) <= 4.0:
                        line.append(item)
                        placed = True
                        break
                if not placed:
                    lines.append([item])
            for line in lines:
                ordered = sorted(line, key=lambda value: value[1])
                content = "".join(item[0] for item in ordered)
                compact = "".join(content.split())
                marker = compact[:3]
                if marker not in by_marker:
                    continue
                entry = by_marker[marker]
                if page_index != entry["page"]:
                    continue
                top = 842.0 - max(item[4] for item in ordered)
                delta = (top - entry["y"]) / entry["size"]
                if -1.0 < delta < 1.0:
                    samples.setdefault(entry["font"], []).append(delta)
    finally:
        document.close()
    return {
        font: round(sum(values) / len(values), 4)
        for font, values in samples.items()
        if values
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output",
        default=str(ROOT / "src" / "font_metrics_table.json"),
    )
    arguments = parser.parse_args()

    with tempfile.TemporaryDirectory() as temporary_directory:
        workdir = Path(temporary_directory)
        docx_path = workdir / "font-calibration.docx"
        pdf_path = workdir / "font-calibration.pdf"
        entries = build_calibration_docx(docx_path)
        render_docx(docx_path, pdf_path)
        table = measure_entries(pdf_path, entries)
        output_path = Path(arguments.output)
        output_path.write_text(
            json.dumps(table, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(json.dumps(table, ensure_ascii=False, indent=2))
        print("written:", output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
