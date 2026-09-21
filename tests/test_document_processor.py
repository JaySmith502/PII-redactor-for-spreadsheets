import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import document_processor
from document_processor import replace_text_document
from processor import ReplacementRule


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def wqn(local_name: str) -> str:
    return f"{{{WORD_NS}}}{local_name}"


def _minimal_docx(path: Path) -> None:
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>John </w:t></w:r>
      <w:r><w:t>Smith</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>ABC Retirement Portfolio</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>""",
        )
        archive.writestr("word/document.xml", document)


def _docx_text(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return "".join(node.text or "" for node in root.iter(wqn("t")))


@unittest.skipIf(document_processor.fitz is None, "PyMuPDF is not installed")
class PdfReplacementTest(unittest.TestCase):
    def test_pdf_replacement(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "sample.pdf"
            output = root / "sample.replaced.pdf"

            document = document_processor.fitz.open()
            page = document.new_page()
            page.insert_text(
                (72, 72),
                "John Smith owns ABC Retirement Portfolio",
                fontsize=18,
                color=(0, 0, 1),
            )
            document.save(source)
            document.close()

            count = replace_text_document(
                source,
                output,
                [ReplacementRule("John Smith", "CLIENT_001")],
            )

            self.assertEqual(count, 1)
            replaced = document_processor.fitz.open(output)
            try:
                text = "".join(page.get_text() for page in replaced)
                spans = [
                    span
                    for page in replaced
                    for block in page.get_text("dict").get("blocks", [])
                    for line in block.get("lines", [])
                    for span in line.get("spans", [])
                    if "CLIENT_001" in span.get("text", "")
                ]
            finally:
                replaced.close()
            self.assertNotIn("John Smith", text)
            self.assertTrue(spans)
            self.assertGreaterEqual(spans[0]["size"], 17)


class DocxReplacementTest(unittest.TestCase):
    def test_docx_replacement_handles_google_docs_exports(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "google-doc-export.docx"
            output = root / "google-doc-export.replaced.docx"
            _minimal_docx(source)

            count = replace_text_document(
                source,
                output,
                [
                    ReplacementRule("John Smith", "CLIENT_001"),
                    ReplacementRule("ABC Retirement Portfolio", "PORTFOLIO_001"),
                ],
            )

            self.assertEqual(count, 2)
            self.assertEqual(_docx_text(output), "CLIENT_001PORTFOLIO_001")
            self.assertIn("John Smith", _docx_text(source))


if __name__ == "__main__":
    unittest.main()
