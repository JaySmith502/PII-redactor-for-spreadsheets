import secrets
import tempfile
import unittest
import zipfile
from pathlib import Path

from processor import (
    CellCipher,
    ReplacementRule,
    cell_value,
    load_package,
    peek_workbook_columns,
    qn,
    read_shared_strings,
    token_namespace_from_header,
    transform_workbook,
    workbook_sheets,
)


def _minimal_workbook(path: Path) -> None:
    worksheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Client Export</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Client Name</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Account ID</t></is></c>
      <c r="C2" t="inlineStr"><is><t>Portfolio Label</t></is></c>
      <c r="D2" t="inlineStr"><is><t>Notes</t></is></c>
      <c r="E2" t="inlineStr"><is><t>Formula Cache</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>John Smith</t></is></c>
      <c r="B3" t="inlineStr"><is><t>ACC-100</t></is></c>
      <c r="C3" t="inlineStr"><is><t>ABC Retirement Portfolio</t></is></c>
      <c r="D3" t="inlineStr"><is><t>John Smith review note</t></is></c>
      <c r="E3"><f>CONCAT(A3,&quot; - &quot;,C3)</f><v>John Smith - ABC Retirement Portfolio</v></c>
    </row>
  </sheetData>
</worksheet>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Ad Hoc Clients" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)


def _pivot_like_workbook(path: Path) -> None:
    worksheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Portfolio Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Account Name</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Total</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>A</t></is></c>
      <c r="B2" t="inlineStr"><is><t>A</t></is></c>
      <c r="C2" t="inlineStr"><is><t>(2.180)</t></is></c>
      <c r="D2" t="inlineStr"><is><t>Internal</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>B</t></is></c>
      <c r="B3" t="inlineStr"><is><t>B</t></is></c>
      <c r="C3" t="inlineStr"><is><t>(652)</t></is></c>
    </row>
  </sheetData>
</worksheet>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Pivot Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)


def _sheet_values(path: Path) -> dict[str, str]:
    parts, _ = load_package(path)
    _, shared_strings = read_shared_strings(parts)
    sheets, _ = workbook_sheets(parts)
    root = next(iter(sheets.values()))
    values = {}
    for cell in root.iter(qn("c")):
        values[cell.get("r")] = cell_value(cell, shared_strings)
    return values


class ProcessorSampleWorkbookTest(unittest.TestCase):
    def test_token_namespace_from_header_slugifies_column_headers(self) -> None:
        self.assertEqual(token_namespace_from_header("Client Name", "A"), "client-name")
        self.assertEqual(
            token_namespace_from_header("Portfolio / Account #", "B"),
            "portfolio-account",
        )
        self.assertEqual(token_namespace_from_header("", "A"), "column-a")
        self.assertEqual(
            token_namespace_from_header("Client: Internal ID", "C"),
            "client-internal-id",
        )

    def test_legacy_custom_tokens_still_decrypt(self) -> None:
        cipher = CellCipher(secrets.token_bytes(64))
        token = cipher.encrypt("custom", "legacy value")
        self.assertTrue(token.startswith("custom:"))
        self.assertEqual(cipher.decrypt(token), "legacy value")

    def test_pivot_like_summary_uses_text_header_not_first_data_row(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "pivot-like.xlsx"
            _pivot_like_workbook(source)

            columns = peek_workbook_columns(source)

            self.assertEqual(
                [column["header"] for column in columns["Pivot Summary"]],
                ["Portfolio Name", "Account Name", "Total", "Column D"],
            )
            self.assertEqual(
                [column["namespace"] for column in columns["Pivot Summary"][:3]],
                ["portfolio-name", "account-name", "total"],
            )

    def test_dynamic_columns_replacements_and_decrypt_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "sample.xlsx"
            encrypted = root / "sample.encrypted.xlsx"
            restored = root / "sample.restored.xlsx"
            _minimal_workbook(source)

            columns = peek_workbook_columns(source)
            self.assertEqual(
                [column["header"] for column in columns["Ad Hoc Clients"]],
                ["Client Name", "Account ID", "Portfolio Label", "Notes", "Formula Cache"],
            )

            cipher = CellCipher(secrets.token_bytes(64))
            stats = transform_workbook(
                source,
                encrypted,
                cipher,
                "encrypt",
                selected_columns={"0|A", "0|C"},
                replacements=[
                    ReplacementRule("John Smith", "CLIENT_001"),
                    ReplacementRule("ABC Retirement Portfolio", "PORTFOLIO_001"),
                ],
            )

            self.assertEqual(stats.replacements, 5)
            self.assertEqual(
                stats.encrypted_columns,
                [("Ad Hoc Clients", "Client Name"), ("Ad Hoc Clients", "Portfolio Label")],
            )

            encrypted_values = _sheet_values(encrypted)
            self.assertTrue(encrypted_values["A3"].startswith("client-name:"))
            self.assertEqual(encrypted_values["B3"], "ACC-100")
            self.assertTrue(encrypted_values["C3"].startswith("portfolio-label:"))
            self.assertEqual(encrypted_values["D3"], "CLIENT_001 review note")
            self.assertEqual(encrypted_values["E3"], "CLIENT_001 - PORTFOLIO_001")

            transform_workbook(encrypted, restored, cipher, "decrypt")
            restored_values = _sheet_values(restored)
            self.assertEqual(restored_values["A3"], "CLIENT_001")
            self.assertEqual(restored_values["B3"], "ACC-100")
            self.assertEqual(restored_values["C3"], "PORTFOLIO_001")
            self.assertEqual(restored_values["D3"], "CLIENT_001 review note")

            source_values = _sheet_values(source)
            self.assertEqual(source_values["A3"], "John Smith")
            self.assertEqual(source_values["C3"], "ABC Retirement Portfolio")


if __name__ == "__main__":
    unittest.main()
