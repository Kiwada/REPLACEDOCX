from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from replacedocx.engine_lib.common import BADGE_TAG  # noqa: E402
from replacedocx.engine_lib.docx_utils import replace_question_prefix_marker_in_paragraph  # noqa: E402


class ReplaceQuestionPrefixMarkerTests(unittest.TestCase):
    def _render_document_xml(self, with_math: bool) -> str:
        with TemporaryDirectory() as td:
            doc_path = Path(td) / "sample.docx"

            doc = Document()
            p = doc.add_paragraph()
            p.add_run("1. (MÉDIA) Energia ")
            if with_math:
                omml = parse_xml(
                    rf"""
                    <m:oMath {nsdecls('m')}>
                      <m:r><m:t>E=mc2</m:t></m:r>
                    </m:oMath>
                    """
                )
                p._element.append(omml)
            p.add_run(" final")
            doc.save(doc_path)

            reopened = Document(doc_path)
            replaced = replace_question_prefix_marker_in_paragraph(
                reopened.paragraphs[0],
                {"(MÉDIA)": "Assets/areas/fisica/capsulas/media.png"},
                1.3,
                "Arial",
                11,
            )
            self.assertTrue(replaced)

            reopened.save(doc_path)
            return ZipFile(doc_path).read("word/document.xml").decode("utf-8")

    def test_replaces_prefix_marker_in_plain_paragraph(self) -> None:
        xml = self._render_document_xml(with_math=False)

        self.assertIn(BADGE_TAG, xml)
        self.assertNotIn("(MÉDIA)", xml)
        self.assertIn("Energia", xml)

    def test_preserves_office_math_when_replacing_prefix_marker(self) -> None:
        xml = self._render_document_xml(with_math=True)

        self.assertIn(BADGE_TAG, xml)
        self.assertIn("m:oMath", xml)
        self.assertIn("Energia", xml)


if __name__ == "__main__":
    unittest.main()
