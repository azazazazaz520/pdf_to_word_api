from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from reportlab.pdfgen.canvas import Canvas

from src.fonts.resolver import FontMatch, PdfFontDescriptor, extract_pdf_font_descriptors, normalize_family, resolve_font, strip_subset_prefix, summarize_font_usage


class FontResolverTest(unittest.TestCase):
    def test_subset_prefix_and_index_are_stripped(self) -> None:
        self.assertEqual(strip_subset_prefix("BCDEEE+SimSun"), "SimSun")
        self.assertEqual(strip_subset_prefix("/AAAAAA+MicrosoftYaHei-0"), "MicrosoftYaHei")
        self.assertEqual(strip_subset_prefix("SimSun"), "SimSun")

    def test_family_aliases(self) -> None:
        cases = {
            "BCDEEE+SimSun": "SimSun",
            "SimHei": "SimHei",
            "MicrosoftYaHei-0": "Microsoft YaHei",
            "TimesNewRomanPS-BoldMT": "Times New Roman",
            "TimesNewRomanPSMT": "Times New Roman",
            "ArialMT": "Arial",
            "Helvetica-Oblique": "Arial",
            "CourierNewPSMT": "Courier New",
            "CambriaMath": "Cambria Math",
            "DengXian": "DengXian",
        }
        for raw, expected in cases.items():
            self.assertEqual(normalize_family(raw), expected, raw)

    def test_name_based_style_detection(self) -> None:
        bold = resolve_font(raw_name="TimesNewRomanPS-BoldMT", text="STM32F103")
        self.assertTrue(bold.bold)
        italic = resolve_font(raw_name="Helvetica-Oblique", text="hello")
        self.assertTrue(italic.italic)
        regular = resolve_font(raw_name="BCDEEE+SimSun", text="正文")
        self.assertFalse(regular.bold)
        self.assertFalse(regular.italic)

    def test_descriptor_style_detection(self) -> None:
        descriptor = PdfFontDescriptor(
            base_font="F1",
            family="CustomFont",
            weight=700,
            flags=1 << 6,
            italic_angle=12.0,
        )
        match = resolve_font(
            raw_name="F1",
            family_hint="",
            descriptor=descriptor,
            text="text",
        )
        self.assertTrue(match.bold)
        self.assertTrue(match.italic)

    def test_unknown_font_is_substituted_with_reason(self) -> None:
        match = resolve_font(
            raw_name="TotallyMissingFont-Regular",
            family_hint="TotallyMissingFont",
            text="hello world",
        )
        self.assertTrue(match.substituted)
        self.assertFalse(match.matched)
        self.assertTrue(match.family)
        self.assertIn("font_not_installed", match.fallback_reason)

    def test_cjk_font_falls_back_to_cjk_chain(self) -> None:
        match = resolve_font(
            raw_name="NotExistingHeiTi",
            family_hint="不存在的黑体",
            text="中文内容",
        )
        self.assertTrue(match.substituted)
        self.assertIn(
            match.family,
            {"Microsoft YaHei", "SimSun", "SimHei", "DengXian"},
        )

    def test_extract_descriptors_from_real_pdf(self) -> None:
        fixture = Path(__file__).resolve().parents[1] / "fixtures" / "synthetic_text_table.pdf"
        if not fixture.is_file():
            self.skipTest("缺少合成表格样本")
        descriptors = extract_pdf_font_descriptors(fixture)
        families = {info.family for info in descriptors.values()}
        self.assertTrue(any("yahei" in family.lower() for family in families), families)

    def test_summarize_font_usage_counts_and_substitutions(self) -> None:
        report = summarize_font_usage(
            [
                {
                    "font_name": "SimSun",
                    "pdf_font_name": "BCDEEE+SimSun",
                    "bold": False,
                    "italic": False,
                    "font_substituted": False,
                },
                {
                    "font_name": "SimSun",
                    "pdf_font_name": "BCDEEE+SimSun",
                    "bold": False,
                    "italic": False,
                    "font_substituted": False,
                },
                {
                    "font_name": "Arial",
                    "pdf_font_name": "MissingFont",
                    "bold": True,
                    "italic": False,
                    "font_substituted": True,
                    "font_fallback_reason": "font_not_installed:MissingFont",
                },
            ]
        )
        self.assertEqual(report["resolved_font_count"], 2)
        self.assertEqual(report["substituted_font_count"], 1)
        self.assertEqual(report["usage"][0]["pdf_font"], "BCDEEE+SimSun")
        self.assertEqual(report["usage"][0]["count"], 2)
        self.assertEqual(
            report["substituted"][0]["reason"],
            "font_not_installed:MissingFont",
        )


class MixedFontSpanTest(unittest.TestCase):
    def test_mixed_font_line_is_split_into_spans(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "mixed-font.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(400, 200))
            canvas.setFont("Helvetica", 12)
            canvas.drawString(40, 150, "中文")
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawString(64, 150, "BOLD123")
            canvas.setFont("Helvetica", 12)
            canvas.drawString(145, 150, "tail")
            canvas.save()

            from src.layout.layout import extract_pdf_layout

            page = extract_pdf_layout(pdf_path, include_fidelity=True).pages[0]
            spans = [
                span for line in page.lines for span in line.spans
            ]
            self.assertGreaterEqual(len(spans), 2)
            fonts = {span.font_name for span in spans}
            self.assertIn("Arial", fonts)
            self.assertTrue(any(span.bold for span in spans))
            self.assertTrue(all(span.pdf_font_name for span in spans), spans)
            span_text = "".join(span.text for span in spans)
            self.assertIn("BOLD123", span_text)


if __name__ == "__main__":
    unittest.main()
