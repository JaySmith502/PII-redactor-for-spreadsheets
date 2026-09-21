"""Regression tests for the pre-deployment hardening pass.

These lock in the behaviours that the optimizations depend on, so a later
refactor cannot quietly reintroduce a correctness problem:

* Worksheet parts that were not modified must come back byte-identical, and
  every value must still round-trip through encrypt/decrypt.
* ``mc:Ignorable`` must never name a namespace prefix that the part does not
  bind, for both XLSX and DOCX, because Word and Excel reject such a part as
  unreadable content.
"""

import secrets
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

import document_processor
from processor import (
    CellCipher,
    ReplacementRule,
    ScrubberError,
    cell_value,
    load_package,
    qn,
    read_shared_strings,
    transform_workbook,
    workbook_sheets,
    xml_bytes,
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
IGNORABLE = f"{{{MC_NS}}}Ignorable"
SHEET_COLUMNS = ("A", "B", "C")


def _sheet_xml(rows: int) -> str:
    header = "".join(
        f'<c r="{column}1" t="inlineStr"><is><t>Header {column}</t></is></c>'
        for column in SHEET_COLUMNS
    )
    body = []
    for row in range(2, rows + 2):
        cells = "".join(
            f'<c r="{column}{row}" t="inlineStr"><is><t>'
            f"{column} value {row}</t></is></c>"
            for column in SHEET_COLUMNS
        )
        body.append(f'<row r="{row}">{cells}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}"><sheetData>'
        f'<row r="1">{header}</row>{"".join(body)}</sheetData></worksheet>'
    )


def _workbook(path: Path, sheet_count: int = 3, rows: int = 3) -> None:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    sheets = "".join(
        f'<sheet name="Sheet {index}" sheetId="{index}" r:id="rId{index}"/>'
        for index in range(1, sheet_count + 1)
    )
    relationships = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, sheet_count + 1)
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f"{overrides}</Types>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets}</sheets></workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{relationships}</Relationships>",
        )
        for index in range(1, sheet_count + 1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _sheet_xml(rows))


def _all_values(path: Path) -> dict[tuple[str, str], str]:
    parts, _ = load_package(path)
    _, shared_strings = read_shared_strings(parts)
    sheets, _ = workbook_sheets(parts)
    values: dict[tuple[str, str], str] = {}
    for sheet_name, root in sheets.items():
        for cell in root.iter(qn("c")):
            values[(sheet_name, cell.get("r"))] = cell_value(cell, shared_strings)
    return values


def _worksheet_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if name.startswith("xl/worksheets/")
        }


class UnchangedSheetTest(unittest.TestCase):
    def test_untouched_sheets_are_byte_identical_and_values_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "multi.xlsx"
            encrypted = root / "multi.encrypted.xlsx"
            restored = root / "multi.restored.xlsx"
            _workbook(source, sheet_count=3, rows=4)

            cipher = CellCipher(secrets.token_bytes(64))
            stats = transform_workbook(
                source,
                encrypted,
                cipher,
                "encrypt",
                selected_columns={"0|A", "0|B"},
            )

            self.assertEqual(stats.total, 8)
            self.assertEqual(
                [sheet for sheet, _header in stats.encrypted_columns],
                ["Sheet 1", "Sheet 1"],
            )

            before = _worksheet_parts(source)
            after = _worksheet_parts(encrypted)
            self.assertEqual(sorted(before), sorted(after))
            # Sheet 1 was encrypted and must be rewritten; sheets 2 and 3 hold
            # nothing selected, so they must be preserved byte for byte.
            self.assertNotEqual(before["xl/worksheets/sheet1.xml"], after["xl/worksheets/sheet1.xml"])
            self.assertEqual(before["xl/worksheets/sheet2.xml"], after["xl/worksheets/sheet2.xml"])
            self.assertEqual(before["xl/worksheets/sheet3.xml"], after["xl/worksheets/sheet3.xml"])

            encrypted_values = _all_values(encrypted)
            self.assertTrue(encrypted_values[("Sheet 1", "A2")].startswith("header-a:"))
            self.assertEqual(encrypted_values[("Sheet 2", "A2")], "A value 2")

            transform_workbook(encrypted, restored, cipher, "decrypt")
            self.assertEqual(_all_values(restored), _all_values(source))

    def test_replace_only_does_not_rewrite_sheets_holding_shared_strings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "replace.xlsx"
            output = root / "replace.out.xlsx"
            _workbook(source, sheet_count=2, rows=2)

            transform_workbook(
                source,
                output,
                None,
                "replace",
                replacements=[ReplacementRule(find="A value 2", replace="CLIENT_001")],
            )

            values = _all_values(output)
            self.assertEqual(values[("Sheet 1", "A2")], "CLIENT_001")
            self.assertEqual(values[("Sheet 2", "A2")], "CLIENT_001")
            self.assertEqual(values[("Sheet 1", "B2")], "B value 2")


class IgnorableNamespaceTest(unittest.TestCase):
    @staticmethod
    def _unbound(xml: bytes) -> list[str]:
        root = ET.fromstring(xml)
        text = xml.decode("utf-8", "replace")
        return [
            prefix
            for prefix in (root.get(IGNORABLE) or "").split()
            if f"xmlns:{prefix}=" not in text
        ]

    def test_known_extension_prefix_is_redeclared(self) -> None:
        root = ET.fromstring(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:mc="{MC_NS}"'
            ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
            ' mc:Ignorable="x14ac"><sheets/></workbook>'
        )
        data = xml_bytes(root)
        self.assertIn(b"xmlns:x14ac=", data)
        self.assertEqual(self._unbound(data), [])

    def test_unknown_prefix_is_dropped_from_ignorable(self) -> None:
        root = ET.fromstring(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:mc="{MC_NS}"'
            ' xmlns:vend="http://example.invalid/vendor"'
            ' mc:Ignorable="vend"><sheets/></workbook>'
        )
        data = xml_bytes(root)
        self.assertEqual(ET.fromstring(data).get(IGNORABLE), None)
        self.assertEqual(self._unbound(data), [])

    def test_docx_rewrite_keeps_word_prefixes_bound(self) -> None:
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            f' xmlns:mc="{MC_NS}"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
            ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
            ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
            ' mc:Ignorable="w14 w15 wp14">'
            "<w:body><w:p><w:r><w:t>John Smith</w:t></w:r></w:p></w:body></w:document>"
        ).encode("utf-8")

        rewritten, count = document_processor.replace_docx_xml(
            document, [ReplacementRule(find="John Smith", replace="CLIENT_001")]
        )

        self.assertEqual(count, 1)
        self.assertIn(b"CLIENT_001", rewritten)
        self.assertEqual(self._unbound(rewritten), [])
        self.assertIn(b"xmlns:w14=", rewritten)
        self.assertIn(b"xmlns:w15=", rewritten)
        self.assertIn(b"xmlns:wp14=", rewritten)


class _Box:
    """Stand-in for a PyMuPDF rectangle."""

    __slots__ = ("x0", "y0", "x1", "y1")

    def __init__(self, x0: float, y0: float, x1: float, y1: float) -> None:
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1

    @property
    def height(self) -> float:
        return self.y1 - self.y0


class _FakePage:
    """Page stub that counts how often the page text is parsed."""

    def __init__(self, spans: list[_Box], matches: dict[str, list[_Box]]) -> None:
        self._spans = spans
        self._matches = matches
        self.get_text_calls = 0
        self.redactions = 0
        self.inserted: list[str] = []

    def get_text(self, kind: str) -> dict:
        self.get_text_calls += 1
        return {
            "blocks": [
                {
                    "lines": [
                        {
                            "spans": [
                                {
                                    "bbox": (box.x0, box.y0, box.x1, box.y1),
                                    "size": 11.0,
                                    "color": 0,
                                    "font": "Helvetica",
                                    "origin": (box.x0, box.y1),
                                }
                                for box in self._spans
                            ]
                        }
                    ]
                }
            ]
        }

    def search_for(self, needle: str) -> list[_Box]:
        return list(self._matches.get(needle, ()))

    def add_redact_annot(self, rectangle, **kwargs) -> None:
        self.redactions += 1

    def apply_redactions(self) -> None:
        pass

    def insert_text(self, point, text, **kwargs) -> None:
        self.inserted.append(text)


class _FakeDocument:
    def __init__(self, pages: list[_FakePage]) -> None:
        self.pages = pages
        self.needs_pass = False

    def __iter__(self):
        return iter(self.pages)

    def save(self, destination, **kwargs) -> None:
        Path(destination).write_bytes(b"%PDF-1.4 stub")

    def close(self) -> None:
        pass


def _grid(count: int) -> list[_Box]:
    return [_Box(0.0, 10.0 * index, 100.0, 10.0 * index + 8.0) for index in range(count)]


class PdfStyleLookupTest(unittest.TestCase):
    """PDF replacement must not re-parse the page for every match.

    ``page.get_text("dict")`` returns the whole page, so calling it from the
    style lookup made a replacement cost grow with (matches x page size). These
    tests pin the count to at most one parse per page, and none at all when the
    page holds no match.
    """

    def _replace(self, pages: list[_FakePage], rules: list[ReplacementRule]) -> int:
        """Run ``replace_pdf`` against stub pages and return the match count."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.pdf"
            source.write_bytes(b"%PDF-1.4 stub")
            fake_fitz = types.SimpleNamespace(
                Rect=lambda bbox: _Box(*bbox),
                open=lambda _source: _FakeDocument(pages),
            )
            with mock.patch.object(document_processor, "fitz", fake_fitz):
                return document_processor.replace_pdf(
                    source, root / "out.pdf", rules, force=True
                )

    def test_page_text_is_parsed_once_per_page(self) -> None:
        matches = {
            "Ada Lovelace": _grid(40),
            "Grace Hopper": _grid(41),
            "Alan Turing": _grid(39),
        }
        page = _FakePage(_grid(3), matches)
        rules = [
            ReplacementRule(find=find, replace=f"CLIENT_{index:03d}")
            for index, find in enumerate(matches, start=1)
        ]

        total = self._replace([page], rules)

        self.assertEqual(total, 120)
        self.assertEqual(page.get_text_calls, 1)
        self.assertEqual(page.redactions, 120)
        self.assertEqual(len(page.inserted), 120)

    def test_page_without_matches_is_never_parsed(self) -> None:
        page = _FakePage(_grid(3), {})
        rules = [ReplacementRule(find="Nobody", replace="CLIENT_001")]

        # No match means nothing to replace, which is reported as an error, but
        # the page must not have been parsed to discover that.
        with self.assertRaises(ScrubberError):
            self._replace([page], rules)

        self.assertEqual(page.get_text_calls, 0)
        self.assertEqual(page.redactions, 0)


if __name__ == "__main__":
    unittest.main()
