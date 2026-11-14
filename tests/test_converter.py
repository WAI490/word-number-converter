import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from docx import Document
from docx.oxml import parse_xml

NSDECLS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)

from converter import (
    convert_text,
    convert_to_english,
    convert_to_persian,
    output_path_for,
    process_document,
)


class ConvertTextTests(unittest.TestCase):
    def test_english_to_persian(self):
        self.assertEqual(convert_to_persian("سال 2024، عدد 15"), "سال ۲۰۲۴، عدد ۱۵")

    def test_persian_to_english(self):
        self.assertEqual(convert_to_english("سال ۲۰۲۴، عدد ۱۵"), "سال 2024، عدد 15")

    def test_arabic_to_english(self):
        self.assertEqual(convert_to_english("العدد ١٢٣"), "العدد 123")

    def test_non_digits_untouched(self):
        self.assertEqual(convert_text("abc XYZ!?", "en2fa"), "abc XYZ!?")
        self.assertEqual(convert_text("متن بدون عدد", "fa2en"), "متن بدون عدد")

    def test_idempotent(self):
        once = convert_text("42", "en2fa")
        self.assertEqual(convert_text(once, "en2fa"), once)
        once = convert_text("۴۲", "fa2en")
        self.assertEqual(convert_text(once, "fa2en"), once)

    def test_unknown_direction_raises(self):
        with self.assertRaises(ValueError):
            convert_text("1", "sideways")

    def test_output_path_suffixes(self):
        self.assertEqual(output_path_for("/a/b.docx", "en2fa"), "/a/b_persian.docx")
        self.assertEqual(output_path_for("/a/b.docx", "fa2en"), "/a/b_english.docx")


def _add_textbox(paragraph, text):
    """Attach a minimal text box (w:txbxContent) holding one run of text."""
    drawing = parse_xml(
        '<w:drawing %s>'
        "  <wp:inline>"
        "    <a:graphic>"
        "      <a:graphicData>"
        "        <wps:wsp>"
        "          <wps:txbx>"
        "            <w:txbxContent>"
        "              <w:p><w:r><w:t>%s</w:t></w:r></w:p>"
        "            </w:txbxContent>"
        "          </wps:txbx>"
        "        </wps:wsp>"
        "      </a:graphicData>"
        "    </a:graphic>"
        "  </wp:inline>"
        "</w:drawing>"
        % (
            NSDECLS,
            text,
        )
    )
    run = paragraph.add_run()
    run._r.append(drawing)


class ProcessDocumentTests(unittest.TestCase):
    def _build_doc(self, path):
        doc = Document()
        doc.add_paragraph("سال 2024 و عدد 15")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "مبلغ 42 تومان"
        header_p = doc.sections[0].header.paragraphs[0]
        header_p.text = "سند شماره 7"
        _add_textbox(doc.add_paragraph(), "داخل جعبه 77")
        doc.save(path)

    def test_process_document_en2fa(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "in.docx")
            out = os.path.join(tmp, "out.docx")
            self._build_doc(src)

            process_document(src, out, "en2fa")

            doc = Document(out)
            self.assertEqual(doc.paragraphs[0].text, "سال ۲۰۲۴ و عدد ۱۵")
            self.assertEqual(doc.tables[0].cell(0, 0).text, "مبلغ ۴۲ تومان")
            self.assertEqual(
                doc.sections[0].header.paragraphs[0].text, "سند شماره ۷"
            )
            body_text = doc.element.body.xml
            self.assertIn("داخل جعبه ۷۷", body_text)
            self.assertNotIn("77", body_text)

    def test_process_document_fa2en_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "in.docx")
            mid = os.path.join(tmp, "mid.docx")
            out = os.path.join(tmp, "out.docx")
            self._build_doc(src)

            process_document(src, mid, "en2fa")
            process_document(mid, out, "fa2en")

            doc = Document(out)
            self.assertEqual(doc.paragraphs[0].text, "سال 2024 و عدد 15")
            self.assertEqual(doc.tables[0].cell(0, 0).text, "مبلغ 42 تومان")


if __name__ == "__main__":
    unittest.main()
