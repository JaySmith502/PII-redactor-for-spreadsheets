"""Office parts must never be parsed with a document type declaration.

``xml.etree.ElementTree`` expands internal entities. A hand-written package
member can therefore turn a few hundred uploaded bytes into hundreds of
megabytes of text inside the parser, and the package-level limits cannot see it
because the zip entry stays tiny. These tests pin the guard that refuses such a
part, and check that ordinary parts still parse.
"""

import secrets
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import document_processor
from processor import (
    CellCipher,
    ScrubberError,
    parse_xml,
    read_shared_strings,
    transform_workbook,
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ENTITY_NAMES = "abcdefghijkl"
# The chain is 11 doublings of a 10-character root value.
EXPANSION_FACTOR = 10 * 2 ** (len(ENTITY_NAMES) - 1)


def _entity_chain() -> tuple[str, str]:
    """Return a nested entity declaration block plus its innermost name.

    Each level doubles the previous one, so the declaration text stays tiny
    while the expanded value grows exponentially.
    """
    declarations = "".join(
        f"<!ENTITY {ENTITY_NAMES[index]} "
        f"'&{ENTITY_NAMES[index - 1]};&{ENTITY_NAMES[index - 1]};'>"
        for index in range(1, len(ENTITY_NAMES))
    )
    return f"<!ENTITY {ENTITY_NAMES[0]} 'aaaaaaaaaa'>{declarations}", ENTITY_NAMES[-1]


def _entity_reference() -> str:
    return f"&{_entity_chain()[1]};"


def _worksheet_xml() -> str:
    declarations, innermost = _entity_chain()
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<!DOCTYPE worksheet [{declarations}]>"
        f'<worksheet xmlns="{MAIN_NS}"><sheetData><row r="1">'
        f'<c r="A1" t="inlineStr"><is><t>&{innermost};</t></is></c>'
        "</row></sheetData></worksheet>"
    )


def _minimal_xlsx(path: Path, sheet_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def _minimal_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        archive.writestr("word/document.xml", document_xml)


class EntityExpansionIsRealTest(unittest.TestCase):
    """The payload is only dangerous because the stock parser expands it."""

    def test_payload_is_tiny_but_expands_hugely_with_the_stock_parser(self) -> None:
        payload = _worksheet_xml().encode("utf-8")

        self.assertLess(len(payload), 900)
        expanded = ET.fromstring(payload)
        text = expanded.find(f".//{{{MAIN_NS}}}t")
        self.assertEqual(len(text.text or ""), EXPANSION_FACTOR)


class ParseXmlGuardTest(unittest.TestCase):
    def test_doctype_is_refused(self) -> None:
        payload = _worksheet_xml().encode("utf-8")
        with self.assertRaises(ScrubberError):
            parse_xml(payload)

    def test_lowercase_doctype_is_also_refused(self) -> None:
        payload = _worksheet_xml().replace("<!DOCTYPE", "<!doctype").encode("utf-8")
        with self.assertRaises(ScrubberError):
            parse_xml(payload)

    def test_plain_parts_still_parse(self) -> None:
        root = parse_xml(f'<worksheet xmlns="{MAIN_NS}"><sheetData/></worksheet>'.encode())
        self.assertEqual(root.tag, f"{{{MAIN_NS}}}worksheet")

    def test_an_entity_reference_without_a_doctype_raises_scrubber_error(self) -> None:
        # Undefined entities are a parse error, not an expansion, but callers
        # must still see a ScrubberError rather than a bare ElementTree error.
        with self.assertRaises(ScrubberError):
            parse_xml(f'<r>{_entity_reference()}</r>'.encode("utf-8"))


class WorkbookEntityGuardTest(unittest.TestCase):
    def test_sheet_with_doctype_is_reported_not_expanded(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "hostile.xlsx"
            _minimal_xlsx(source, _worksheet_xml())

            with self.assertRaises(ScrubberError):
                transform_workbook(
                    source,
                    root / "out.xlsx",
                    CellCipher(secrets.token_bytes(64)),
                    "encrypt",
                    selected_columns={"0|A"},
                )

    def test_shared_strings_with_doctype_is_refused(self) -> None:
        declarations, innermost = _entity_chain()
        data = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<!DOCTYPE sst [{declarations}]>"
            f'<sst xmlns="{MAIN_NS}"><si><t>&{innermost};</t></si></sst>'
        ).encode("utf-8")

        with self.assertRaises(ScrubberError):
            read_shared_strings({"xl/sharedStrings.xml": data})


class DocxEntityGuardTest(unittest.TestCase):
    def _document_xml(self) -> str:
        declarations, innermost = _entity_chain()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<!DOCTYPE w:document [{declarations}]>"
            f'<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:r>'
            f"<w:t>&{innermost};</w:t>"
            "</w:r></w:p></w:body></w:document>"
        )

    def test_docx_with_doctype_is_reported_not_expanded(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "hostile.docx"
            _minimal_docx(source, self._document_xml())

            with self.assertRaises(ScrubberError):
                document_processor.replace_docx(
                    source,
                    root / "out.docx",
                    [document_processor.ReplacementRule(find="a", replace="B")],
                )


if __name__ == "__main__":
    unittest.main()
