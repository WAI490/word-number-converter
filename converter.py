"""Core conversion logic for Word Number Converter.

Converts digits inside .docx documents between English and Persian
(both Persian and Arabic digit sets), preserving run formatting.

This module is shared by both conversion directions and has no GUI
dependencies, so it can be unit-tested and reused from a CLI.
"""

import os

from docx import Document
from docx.text.paragraph import Paragraph

# --- Digit maps -------------------------------------------------------

EN_TO_FA = {
    "0": "۰", "1": "۱", "2": "۲", "3": "۳", "4": "۴",
    "5": "۵", "6": "۶", "7": "۷", "8": "۸", "9": "۹",
}

FA_TO_EN = {
    # Persian digits
    "۰": "0", "۱": "1", "۲": "2", "۳": "3", "۴": "4",
    "۵": "5", "۶": "6", "۷": "7", "۸": "8", "۹": "9",
    # Arabic digits
    "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4",
    "٥": "5", "٦": "6", "٧": "7", "٨": "8", "٩": "9",
}

DIRECTIONS = {
    "en2fa": {
        "map": EN_TO_FA,
        "suffix": "_persian",
        "label": "English → Persian",
    },
    "fa2en": {
        "map": FA_TO_EN,
        "suffix": "_english",
        "label": "Persian → English",
    },
}


def _check_direction(direction):
    if direction not in DIRECTIONS:
        raise ValueError(
            f"Unknown direction {direction!r}; expected one of {sorted(DIRECTIONS)}"
        )


def convert_text(text, direction):
    """Convert the digits in a single string. Pure function."""
    _check_direction(direction)
    digit_map = DIRECTIONS[direction]["map"]
    return "".join(digit_map.get(ch, ch) for ch in text)


def convert_to_persian(text):
    return convert_text(text, "en2fa")


def convert_to_english(text):
    return convert_text(text, "fa2en")


def output_path_for(input_path, direction):
    """Default output path: original name + direction suffix."""
    _check_direction(direction)
    base, ext = os.path.splitext(input_path)
    return f"{base}{DIRECTIONS[direction]['suffix']}{ext}"


# --- Document processing ----------------------------------------------

# WordprocessingML namespace, for finding text-box content.
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_TXBX_CONTENT = f"{{{_W_NS}}}txbxContent"
_P = f"{{{_W_NS}}}p"

# Package parts that python-docx does not expose as paragraphs but that
# can still hold user-visible numbers.
_EXTRA_PART_NAMES = ("footnotes", "endnotes", "comments")


def _process_paragraph(paragraph, direction):
    digit_map = DIRECTIONS[direction]["map"]
    for run in paragraph.runs:
        if any(ch in digit_map for ch in run.text):
            run.text = convert_text(run.text, direction)


def _process_table(table, direction):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _process_paragraph(paragraph, direction)
            for nested_table in cell.tables:
                _process_table(nested_table, direction)


def _process_textboxes(parent_element, direction):
    """Convert paragraphs inside text boxes (w:txbxContent).

    Text boxes are not part of Document.paragraphs / tables, so they
    need this extra pass over the raw XML. Conversion is idempotent,
    so a paragraph already handled elsewhere is not harmed.
    """
    for txbx in parent_element.iter(_TXBX_CONTENT):
        for p_element in txbx.iter(_P):
            _process_paragraph(Paragraph(p_element, None), direction)


def _process_header_or_footer(container, direction):
    for paragraph in container.paragraphs:
        _process_paragraph(paragraph, direction)
    for table in container.tables:
        _process_table(table, direction)
    _process_textboxes(container._element, direction)


def _process_extra_parts(doc, direction):
    """Convert footnotes, endnotes and comments when those parts exist.

    Best effort: a document without such a part, or a part python-docx
    cannot walk, is skipped instead of failing the whole conversion.
    """
    try:
        parts = list(doc.part.package.iter_parts())
    except Exception:
        return
    for part in parts:
        try:
            name = str(part.partname).lower()
            if not any(key in name for key in _EXTRA_PART_NAMES):
                continue
            for p_element in part._element.iter(_P):
                _process_paragraph(Paragraph(p_element, None), direction)
        except Exception:
            continue


def process_document(input_path, output_path, direction):
    """Convert every supported part of a .docx file and save the result.

    Covers body paragraphs, tables (including nested tables), headers
    and footers (default, first-page and even-page), text boxes, and —
    when present — footnotes, endnotes and comments. Only the text of
    each run is changed, so fonts, colors and other formatting are
    preserved.

    Not covered: field results such as automatic page numbers and
    tables of contents (Word generates those itself), and digits
    inside embedded objects or images.
    """
    _check_direction(direction)
    doc = Document(input_path)

    for paragraph in doc.paragraphs:
        _process_paragraph(paragraph, direction)

    for table in doc.tables:
        _process_table(table, direction)

    _process_textboxes(doc.element.body, direction)

    for section in doc.sections:
        headers = (section.header, section.first_page_header, section.even_page_header)
        footers = (section.footer, section.first_page_footer, section.even_page_footer)
        for container in headers + footers:
            if container:
                _process_header_or_footer(container, direction)

    _process_extra_parts(doc, direction)

    doc.save(output_path)
