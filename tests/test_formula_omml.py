from __future__ import annotations

import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from src.ir.model import IRBlock, IRDocument, IRPage
from src.formula_omml import formula_text_to_latex, formula_text_to_omml


class FormulaOmmlTest(unittest.TestCase):
    def test_unicode_formula_is_normalized_to_latex(self) -> None:
        latex = formula_text_to_latex("∑ᵢ xᵢ = 1")

        self.assertIn(r"\sum", latex)
        self.assertIn("_{i}", latex)

    def test_formula_converts_to_omml_structures(self) -> None:
        ok, omml, _ = formula_text_to_omml(r"x = \frac{a}{b} + 1")
        self.assertTrue(ok)
        self.assertIn("m:f", omml or "")

        ok, omml, _ = formula_text_to_omml("∑ᵢ xᵢ = 1")
        self.assertTrue(ok)
        self.assertIn("m:nary", omml or "")

        ok, omml, _ = formula_text_to_omml("√x = 2")
        self.assertTrue(ok)
        self.assertIn("m:rad", omml or "")

if __name__ == "__main__":
    unittest.main()