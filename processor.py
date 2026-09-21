import base64
import binascii
import copy
import os
import posixpath
import re
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

try:
    from cryptography.exceptions import InvalidTag
    from cryptography.hazmat.primitives.ciphers.aead import AESSIV
except ImportError:
    AESSIV = None
    InvalidTag = Exception


LEGACY_TOKEN_NAMESPACES = {"account-number", "account-name", "portfolio", "custom"}
MAX_TOKEN_NAMESPACE_CHARS = 64
TOKEN_NAMESPACE_RE = re.compile(r"^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$")
TOKEN_CIPHERTEXT_RE = re.compile(r"^[A-Za-z0-9_-]+$")
KEY_BYTES = 64
ASSOCIATED_DATA_CONTEXT = b"pii-xlsx-scrubber:v1"
MAX_PACKAGE_MEMBERS = 10_000
MAX_UNCOMPRESSED_BYTES = 250 * 1024 * 1024


class ScrubberError(RuntimeError):
    """Expected workbook-processing failure safe to handle at the UI boundary."""


def encode_urlsafe(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def decode_urlsafe(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    try:
        return base64.urlsafe_b64decode(value + padding)
    except (ValueError, binascii.Error) as exc:
        raise ScrubberError("Malformed encrypted value") from exc


def valid_token_namespace(namespace: str) -> bool:
    return namespace in LEGACY_TOKEN_NAMESPACES or TOKEN_NAMESPACE_RE.fullmatch(namespace) is not None


def token_namespace_from_header(header: str, column: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "-", header.casefold()).strip("-")
    if not normalized:
        normalized = f"column-{column.casefold()}"
    normalized = normalized[:MAX_TOKEN_NAMESPACE_CHARS].strip("-")
    return normalized or f"column-{column.casefold()}"


class CellCipher:
    def __init__(self, key: bytes):
        if AESSIV is None:
            raise ScrubberError("The cryptography package is required")
        if len(key) != KEY_BYTES:
            raise ScrubberError(f"SECRET_KEY must decode to exactly {KEY_BYTES} bytes")
        self._cipher = AESSIV(key)

    @staticmethod
    def token_parts(value: str) -> tuple[str, str] | None:
        namespace, separator, ciphertext = value.partition(":")
        try:
            decoded = decode_urlsafe(ciphertext) if ciphertext else b""
        except ScrubberError:
            return None
        if (
            separator
            and valid_token_namespace(namespace)
            and TOKEN_CIPHERTEXT_RE.fullmatch(ciphertext)
            and len(decoded) >= 16
        ):
            return namespace, ciphertext
        return None

    @classmethod
    def is_encrypted(cls, value: str) -> bool:
        return cls.token_parts(value) is not None

    @classmethod
    def namespace(cls, value: str) -> str:
        parts = cls.token_parts(value)
        if parts is None:
            raise ScrubberError("Malformed encrypted value")
        return parts[0]

    def encrypt(self, namespace: str, plaintext: str) -> str:
        if self.is_encrypted(plaintext):
            raise ScrubberError("Input already contains an encrypted value")
        if not valid_token_namespace(namespace):
            raise ScrubberError(f"Unsupported encryption namespace: {namespace}")
        associated_data = [ASSOCIATED_DATA_CONTEXT, namespace.encode("utf-8")]
        ciphertext = self._cipher.encrypt(plaintext.encode("utf-8"), associated_data)
        return f"{namespace}:{encode_urlsafe(ciphertext)}"

    def decrypt(self, token: str) -> str:
        parts = self.token_parts(token)
        if parts is None:
            return token
        namespace, encoded = parts
        associated_data = [ASSOCIATED_DATA_CONTEXT, namespace.encode("utf-8")]
        try:
            plaintext = self._cipher.decrypt(decode_urlsafe(encoded), associated_data)
            return plaintext.decode("utf-8")
        except (InvalidTag, UnicodeDecodeError) as exc:
            raise ScrubberError(
                "Unable to decrypt a value: the key is wrong or the ciphertext was modified"
            ) from exc


GENERIC_PIVOT_VALUES = {"", "(blank)", "grand total", "total"}
COLUMN_KEY_RE = re.compile(r"^(\d+)\|([A-Z]+)$")
HEADER_SCAN_ROWS = 10


@dataclass
class TransformStats:
    values: Counter[str] = field(default_factory=Counter)
    excluded_sheets: list[str] = field(default_factory=list)
    encrypted_columns: list[tuple[str, str]] = field(default_factory=list)  # (sheet_name, col_header)
    replacements: int = 0

    @property
    def total(self) -> int:
        return sum(self.values.values())

    def record(self, namespace: str) -> None:
        self.values[namespace] += 1


@dataclass(frozen=True)
class ColumnDescriptor:
    sheet_index: int
    sheet_name: str
    column: str
    header: str
    start_row: int
    key: str


@dataclass(frozen=True)
class ReplacementRule:
    find: str
    replace: str


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
CELL_REFERENCE_RE = re.compile(r"^([A-Z]+)(\d+)$")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", OFFICE_REL_NS)

OFFICE_EXTENSION_NAMESPACES = {
    # SpreadsheetML extensions.
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
    "xcalcf": "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
    "xlrd2": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "xpdl": "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    # WordprocessingML extensions. Word documents list these in mc:Ignorable and
    # document_processor shares :func:`xml_bytes` with this module, so the Word
    # prefixes have to be restorable here as well.
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp15": "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
}
for _prefix, _namespace in OFFICE_EXTENSION_NAMESPACES.items():
    ET.register_namespace(_prefix, _namespace)


def qn(local_name: str, namespace: str = MAIN_NS) -> str:
    return f"{{{namespace}}}{local_name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalize_label(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def column_sort_key(column: str) -> int:
    value = 0
    for character in column:
        value = (value * 26) + (ord(character) - ord("A") + 1)
    return value


def cell_position(reference: str) -> tuple[str, int] | None:
    match = CELL_REFERENCE_RE.match(reference)
    if not match:
        return None
    return match.group(1), int(match.group(2))


def selected_column_parts(key: str) -> tuple[int, str] | None:
    match = COLUMN_KEY_RE.match(key)
    if match is None:
        return None
    return int(match.group(1)), match.group(2)


def looks_numeric(value: str) -> bool:
    normalized = value.strip()
    if normalized.startswith("(") and normalized.endswith(")"):
        normalized = normalized[1:-1]
    normalized = normalized.replace(",", "").replace(".", "").replace(" ", "")
    return bool(normalized) and normalized.lstrip("+-").isdigit()


def header_candidate_score(row_values: dict[str, str]) -> int:
    """Rank how much a row looks like a header rather than a row of data.

    Only ever used to choose between rows near the top of a sheet, so the score
    has to separate a title/metadata row from the header row. A cell that parses
    as a number is the useful signal there, because totals and years look
    numeric while headers do not.

    The score deliberately ignores cell length and embedded spaces. Both were
    tried and both are anti-signals: a data value such as "ABC Retirement
    Portfolio" is exactly as wordy as the header "Portfolio Name", so rewarding
    wordiness ranks the first *data* row above the header. Because the caller
    then starts encrypting at ``header_row + 1``, that mistake silently leaves
    the first row of real PII unencrypted.
    """
    score = 0
    for value in row_values.values():
        text = value.strip()
        if not text:
            continue
        if looks_numeric(text):
            score -= 4
            continue
        if len(text) <= 2:
            score -= 1
        if re.search(r"[A-Za-z]{3,}", text):
            score += 4
    return score + len(row_values)


MARKUP_COMPATIBILITY_NS = OFFICE_EXTENSION_NAMESPACES["mc"]
IGNORABLE_ATTRIBUTE = qn("Ignorable", MARKUP_COMPATIBILITY_NS)


def _declared_prefixes(data: bytes, prefixes: list[str]) -> set[str]:
    return {prefix for prefix in prefixes if f"xmlns:{prefix}=".encode("ascii") in data}


def xml_bytes(root: ET.Element) -> bytes:
    """Serialize Office XML with a self-consistent ``mc:Ignorable``.

    ``ElementTree`` drops a namespace declaration when no element or attribute
    in the tree uses that namespace. ``mc:Ignorable`` still names those
    prefixes afterwards, and Word and Excel reject a part whose ``mc:Ignorable``
    names a prefix with no binding ("unreadable content"). Two cases:

    * Known Excel/Word extension prefixes are re-declared from
      :data:`OFFICE_EXTENSION_NAMESPACES`, which preserves the original meaning.
    * Any prefix we cannot bind is removed from ``mc:Ignorable``. That is safe
      because ``ElementTree`` only drops a declaration for a namespace that
      nothing in the part uses, so there is nothing left to ignore.
    """
    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    prefixes = (root.get(IGNORABLE_ATTRIBUTE) or "").split()
    if not prefixes:
        return data

    declared = _declared_prefixes(data, prefixes)
    unbound = [
        prefix
        for prefix in prefixes
        if prefix not in declared and prefix not in OFFICE_EXTENSION_NAMESPACES
    ]
    if unbound:
        kept = [prefix for prefix in prefixes if prefix not in unbound]
        if kept:
            root.set(IGNORABLE_ATTRIBUTE, " ".join(kept))
        else:
            del root.attrib[IGNORABLE_ATTRIBUTE]
        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        prefixes = kept
        declared = _declared_prefixes(data, prefixes)

    missing_declarations = [
        f' xmlns:{prefix}="{OFFICE_EXTENSION_NAMESPACES[prefix]}"'.encode("utf-8")
        for prefix in prefixes
        if prefix not in declared
    ]
    if not missing_declarations:
        return data

    declaration_end = data.find(b"?>")
    root_start = data.find(b"<", declaration_end + 2)
    root_end = data.find(b">", root_start)
    if root_start < 0 or root_end < 0:
        raise ScrubberError("Unable to serialize Office XML root element")
    return data[:root_end] + b"".join(missing_declarations) + data[root_end:]


DTD_MARKERS = (b"<!DOCTYPE", b"<!doctype")


def parse_xml(data: bytes) -> ET.Element:
    """Parse an Office XML part, refusing any part that declares a DTD.

    ``ElementTree`` expands internal entities, so a hand-written part can turn a
    few hundred uploaded bytes into hundreds of megabytes of text *inside the
    parser*, before any limit on the package can see it: the zip entry stays
    tiny, so the member count and uncompressed-size checks both pass. That is
    the "billion laughs" attack.

    No Office part uses a DTD and ``ElementTree`` has no external-entity support
    to lose, so a ``DOCTYPE`` declaration is refused outright rather than
    expanded. Both spellings are checked because the scan is a fast byte search
    while ``xml.etree`` itself is case sensitive.

    Parse failures are reported as :class:`ScrubberError` so every caller can
    handle malformed input the same way as the other rejections, instead of each
    one needing to know about :class:`xml.etree.ElementTree.ParseError`.
    """
    if any(marker in data for marker in DTD_MARKERS):
        raise ScrubberError("XML part declares a DOCTYPE, which Office files never use")
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise ScrubberError("XML part could not be parsed") from exc


def resolve_relationship_target(source_part: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relationship_part(source_part: str) -> str:
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def shared_string_values(root: ET.Element) -> list[str]:
    """Flatten a sharedStrings root into a positional list of plain values."""
    return [
        "".join(node.text or "" for node in item.iter(qn("t")))
        for item in root.findall(qn("si"))
    ]


def read_shared_strings(parts: dict[str, bytes]) -> tuple[ET.Element | None, list[str]]:
    data = parts.get("xl/sharedStrings.xml")
    if data is None:
        return None, []
    root = parse_xml(data)
    return root, shared_string_values(root)


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "s":
        value = cell.find(qn("v"))
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError) as exc:
            raise ScrubberError(f"Invalid shared-string index in cell {cell.get('r', '?')}") from exc
    if cell_type == "inlineStr":
        inline = cell.find(qn("is"))
        if inline is None:
            return ""
        return "".join(node.text or "" for node in inline.iter(qn("t")))
    value = cell.find(qn("v"))
    return value.text if value is not None and value.text is not None else ""


def set_inline_string(cell: ET.Element, value: str) -> None:
    for child in list(cell):
        if local_name(child.tag) in {"v", "is"}:
            cell.remove(child)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, qn("is"))
    text = ET.SubElement(inline, qn("t"))
    if value != value.strip():
        text.set(qn("space", XML_NS), "preserve")
    text.text = value


def set_formula_cache(cell: ET.Element, value: str) -> None:
    inline = cell.find(qn("is"))
    if inline is not None:
        cell.remove(inline)
    cached = cell.find(qn("v"))
    if cached is None:
        cached = ET.SubElement(cell, qn("v"))
    cached.text = value
    cell.set("t", "str")


def set_restored_value(cell: ET.Element, value: str) -> None:
    """Restore a decrypted value, preserving numeric types (like dates)."""
    try:
        float(value)
        is_number = True
    except ValueError:
        is_number = False

    if is_number:
        for child in list(cell):
            if local_name(child.tag) in {"v", "is"}:
                cell.remove(child)
        cell.set("t", "n")
        v = ET.SubElement(cell, qn("v"))
        v.text = value
    else:
        set_inline_string(cell, value)


def decrypt_shared_string_table(
    root: ET.Element | None, cipher: CellCipher, stats: TransformStats
) -> None:
    if root is None:
        return
    for text in root.iter(qn("t")):
        if text.text and CellCipher.is_encrypted(text.text):
            namespace = CellCipher.namespace(text.text)
            text.text = cipher.decrypt(text.text)
            stats.record(namespace)


def rebuild_shared_string_table(
    root: ET.Element | None, sheets: Iterable[ET.Element]
) -> bool:
    """Drop unused shared strings and compact the table.

    Returns ``True`` when a cell's shared-string index was rewritten, which is
    the signal :func:`transform_workbook` uses to decide whether the worksheet
    parts that reference the table need to be serialized again.
    """
    if root is None:
        return False
    old_items = list(root.findall(qn("si")))
    references: list[tuple[ET.Element, int]] = []
    used_indices: set[int] = set()
    for sheet in sheets:
        for cell in sheet.iter(qn("c")):
            if cell.get("t") != "s":
                continue
            value = cell.find(qn("v"))
            if value is None or value.text is None:
                continue
            try:
                index = int(value.text)
            except ValueError as exc:
                raise ScrubberError("Invalid shared-string cell index") from exc
            if index < 0 or index >= len(old_items):
                raise ScrubberError("Shared-string cell index is out of range")
            references.append((value, index))
            used_indices.add(index)
    ordered_indices = sorted(used_indices)
    index_map = {old: new for new, old in enumerate(ordered_indices)}
    references_changed = False
    for reference, old_index in references:
        new_index = str(index_map[old_index])
        if reference.text != new_index:
            reference.text = new_index
            references_changed = True
    non_items = [child for child in root if local_name(child.tag) != "si"]
    root[:] = [copy.deepcopy(old_items[index]) for index in ordered_indices] + non_items
    root.set("count", str(len(references)))
    root.set("uniqueCount", str(len(ordered_indices)))
    return references_changed


def should_skip_direct_value(value: str) -> bool:
    return value.casefold().strip() in GENERIC_PIVOT_VALUES


def workbook_sheets(parts: dict[str, bytes]) -> tuple[dict[str, ET.Element], dict[str, str]]:
    try:
        workbook = parse_xml(parts["xl/workbook.xml"])
        relationships = parse_xml(parts["xl/_rels/workbook.xml.rels"])
    except (KeyError, ET.ParseError, ScrubberError) as exc:
        raise ScrubberError("Input is not a valid XLSX workbook package") from exc
    targets = {
        relationship.get("Id"): relationship.get("Target") for relationship in relationships
    }
    trees: dict[str, ET.Element] = {}
    paths: dict[str, str] = {}
    sheets = workbook.find(qn("sheets"))
    if sheets is None:
        raise ScrubberError("Workbook does not contain any worksheets")
    for sheet in sheets:
        name = sheet.get("name", "")
        relationship_id = sheet.get(qn("id", OFFICE_REL_NS))
        target = targets.get(relationship_id)
        if not target:
            raise ScrubberError(f"Missing relationship for worksheet {name!r}")
        path = resolve_relationship_target("xl/workbook.xml", target)
        try:
            trees[name] = parse_xml(parts[path])
        except (KeyError, ET.ParseError, ScrubberError) as exc:
            raise ScrubberError(f"Missing or invalid worksheet part for {name!r}") from exc
        paths[name] = path
    return trees, paths


def worksheet_column_descriptors(
    sheet_index: int,
    sheet_name: str,
    root: ET.Element,
    shared_strings: list[str],
) -> list[ColumnDescriptor]:
    values_by_row: dict[int, dict[str, str]] = {}
    used_columns: set[str] = set()
    for cell in root.findall(f".//{qn('sheetData')}/{qn('row')}/{qn('c')}"):
        position = cell_position(cell.get("r", ""))
        if position is None:
            continue
        column, row = position
        value = cell_value(cell, shared_strings)
        if value == "":
            continue
        values_by_row.setdefault(row, {})[column] = value
        used_columns.add(column)

    if not values_by_row:
        return []

    candidate_rows = [
        (row, columns)
        for row, columns in values_by_row.items()
        if row <= HEADER_SCAN_ROWS
    ]
    multi_cell_candidates = [
        (row, columns) for row, columns in candidate_rows if len(columns) > 1
    ]
    if multi_cell_candidates:
        # Ties break towards the *earliest* row (``-item[0]``), which is the safe
        # direction: adopting a data row as the header row pushes ``start_row``
        # past it and leaves that row's values unencrypted, whereas adopting a
        # title row as the header only over-encrypts the real header row.
        header_row = max(
            multi_cell_candidates,
            key=lambda item: (
                header_candidate_score(item[1]),
                len(item[1]),
                -item[0],
            ),
        )[0]
    else:
        header_row = min(values_by_row)
    header_values = values_by_row.get(header_row, {})
    descriptors: list[ColumnDescriptor] = []
    for column in sorted(used_columns, key=column_sort_key):
        header = header_values.get(column) or f"Column {column}"
        descriptors.append(
            ColumnDescriptor(
                sheet_index=sheet_index,
                sheet_name=sheet_name,
                column=column,
                header=header,
                start_row=header_row + 1,
                key=f"{sheet_index}|{column}",
            )
        )
    return descriptors


def workbook_column_descriptors(
    sheets: dict[str, ET.Element], shared_strings: list[str]
) -> dict[tuple[int, str], ColumnDescriptor]:
    descriptors: dict[tuple[int, str], ColumnDescriptor] = {}
    for sheet_index, (sheet_name, root) in enumerate(sheets.items()):
        for descriptor in worksheet_column_descriptors(
            sheet_index, sheet_name, root, shared_strings
        ):
            descriptors[(descriptor.sheet_index, descriptor.column)] = descriptor
    return descriptors


def selected_column_descriptors(
    sheets: dict[str, ET.Element],
    shared_strings: list[str],
    selected_columns: set[str] | None,
) -> dict[tuple[int, str], ColumnDescriptor]:
    available = workbook_column_descriptors(sheets, shared_strings)
    if selected_columns is None:
        return available

    selected: dict[tuple[int, str], ColumnDescriptor] = {}
    for key in selected_columns:
        parts = selected_column_parts(key)
        if parts is None:
            continue
        descriptor = available.get(parts)
        if descriptor is not None:
            selected[parts] = descriptor
    return selected


def replacement_value(value: str, replacements: Iterable[ReplacementRule]) -> tuple[str, int]:
    updated = value
    count = 0
    for rule in replacements:
        if not rule.find:
            continue
        count += updated.count(rule.find)
        updated = updated.replace(rule.find, rule.replace)
    return updated, count


def replace_shared_string_table(
    root: ET.Element | None,
    replacements: list[ReplacementRule],
    stats: TransformStats,
) -> None:
    if root is None or not replacements:
        return
    for text in root.iter(qn("t")):
        if text.text is None:
            continue
        updated, count = replacement_value(text.text, replacements)
        if count:
            text.text = updated
            stats.replacements += count


def replace_worksheet_cells(
    sheets: dict[str, ET.Element],
    shared_strings: list[str],
    replacements: list[ReplacementRule],
    stats: TransformStats,
) -> set[str]:
    """Apply replacement rules, returning the names of sheets that changed."""
    changed: set[str] = set()
    if not replacements:
        return changed
    for sheet_name, root in sheets.items():
        for cell in root.findall(f".//{qn('sheetData')}/{qn('row')}/{qn('c')}"):
            formula = cell.find(qn("f"))
            if formula is not None and formula.text:
                updated_formula, formula_count = replacement_value(
                    formula.text, replacements
                )
                if formula_count:
                    formula.text = updated_formula
                    stats.replacements += formula_count
                    changed.add(sheet_name)
            if cell.get("t") == "s":
                continue
            value = cell_value(cell, shared_strings)
            if not value:
                continue
            updated, count = replacement_value(value, replacements)
            if not count:
                continue
            if cell.find(qn("f")) is not None:
                set_formula_cache(cell, updated)
            else:
                set_inline_string(cell, updated)
            stats.replacements += count
            changed.add(sheet_name)
    return changed


def replace_pivot_cache_values(
    parts: dict[str, bytes],
    replacements: list[ReplacementRule],
    stats: TransformStats,
) -> None:
    if not replacements:
        return
    paths = sorted(
        path
        for path in parts
        if re.fullmatch(
            r"xl/pivotCache/(?:pivotCacheDefinition|pivotCacheRecords)\d+\.xml", path
        )
    )
    for path in paths:
        root = parse_xml(parts[path])
        changed = False
        for element in root.iter():
            value = element.get("v")
            if not value:
                continue
            updated, count = replacement_value(value, replacements)
            if not count:
                continue
            element.set("v", updated)
            stats.replacements += count
            changed = True
        if changed:
            parts[path] = xml_bytes(root)


def peek_workbook_columns(source: Path) -> dict[str, list[dict]]:
    """Return per-sheet column info for the workbook without encrypting anything.

    Returns a dict keyed by sheet name. Each value is a list of column dicts.
    The key is stable for the current uploaded workbook and is posted back by
    the UI when the user chooses columns to encrypt.
    """
    parts, _ = load_package(source)
    _, shared_strings = read_shared_strings(parts)
    sheets, _ = workbook_sheets(parts)
    result: dict[str, list[dict]] = {}
    for sheet_index, (sheet_name, root) in enumerate(sheets.items()):
        cols: list[dict] = []
        for descriptor in worksheet_column_descriptors(
            sheet_index, sheet_name, root, shared_strings
        ):
            cols.append(
                {
                    "col": descriptor.column,
                    "header": descriptor.header,
                    "namespace": token_namespace_from_header(
                        descriptor.header, descriptor.column
                    ),
                    "ns_label": f"{sheet_name}: {descriptor.header}",
                    "key": descriptor.key,
                    "start_row": descriptor.start_row,
                }
            )
        if cols:
            result[sheet_name] = cols
    return result


def encrypt_worksheet_cells(
    sheets: dict[str, ET.Element],
    shared_strings: list[str],
    cipher: CellCipher,
    stats: TransformStats,
    selected: dict[tuple[int, str], ColumnDescriptor],
) -> set[str]:
    """Encrypt PII cells in worksheet data, returning changed sheet names.

    *selected* maps ``(sheet_index, column)`` to a :class:`ColumnDescriptor` for
    every column the caller wants encrypted.  It is resolved once by
    :func:`transform_workbook` and passed in, because rebuilding it here walks
    every cell in every sheet a second time (see the profiling note on
    :func:`transform_workbook`).  Pass the full descriptor map from
    :func:`selected_column_descriptors` to encrypt every discovered column.
    """
    changed: set[str] = set()
    selected_headers_seen: set[tuple[str, str]] = set()

    for sheet_index, (sheet_name, root) in enumerate(sheets.items()):
        for cell in root.findall(f".//{qn('sheetData')}/{qn('row')}/{qn('c')}"):
            position = cell_position(cell.get("r", ""))
            if position is None:
                continue
            column, row = position
            descriptor = selected.get((sheet_index, column))
            if descriptor is None or row < descriptor.start_row:
                continue
            value = cell_value(cell, shared_strings)
            if not value:
                continue
            if should_skip_direct_value(value):
                continue
            namespace = token_namespace_from_header(descriptor.header, descriptor.column)
            encrypted = cipher.encrypt(namespace, value)
            if cell.find(qn("f")) is not None:
                set_formula_cache(cell, encrypted)
            else:
                set_inline_string(cell, encrypted)
            stats.record(namespace)
            changed.add(sheet_name)

            seen_key = (sheet_name, column)
            if seen_key not in selected_headers_seen:
                selected_headers_seen.add(seen_key)
                stats.encrypted_columns.append((sheet_name, descriptor.header))

    return changed


def decrypt_worksheet_cells(
    sheets: dict[str, ET.Element],
    shared_strings: list[str],
    cipher: CellCipher,
    stats: TransformStats,
) -> set[str]:
    """Decrypt encrypted cells, returning the names of sheets that changed."""
    changed: set[str] = set()
    for sheet_name, root in sheets.items():
        for cell in root.findall(f".//{qn('sheetData')}/{qn('row')}/{qn('c')}"):
            if cell.get("t") == "s":
                continue
            value = cell_value(cell, shared_strings)
            if not CellCipher.is_encrypted(value):
                continue
            namespace = CellCipher.namespace(value)
            plaintext = cipher.decrypt(value)
            if cell.find(qn("f")) is not None:
                set_formula_cache(cell, plaintext)
            else:
                set_restored_value(cell, plaintext)
            stats.record(namespace)
            changed.add(sheet_name)

    return changed


def pivot_record_path(parts: dict[str, bytes], definition_path: str) -> str | None:
    rel_data = parts.get(relationship_part(definition_path))
    if rel_data is None:
        return None
    relationships = parse_xml(rel_data)
    for relationship in relationships:
        rel_type = relationship.get("Type", "")
        target = relationship.get("Target")
        if target and rel_type.endswith("/pivotCacheRecords"):
            return resolve_relationship_target(definition_path, target)
    return None


def encrypt_pivot_caches(
    parts: dict[str, bytes],
    cipher: CellCipher,
    stats: TransformStats,
    selected_header_namespaces: dict[str, str],
) -> None:
    if not selected_header_namespaces:
        return
    definition_paths = sorted(
        path
        for path in parts
        if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", path)
    )
    for definition_path in definition_paths:
        root = parse_xml(parts[definition_path])
        fields = root.find(qn("cacheFields"))
        target_fields: dict[int, str] = {}
        if fields is not None:
            for index, field_node in enumerate(fields):
                namespace = selected_header_namespaces.get(
                    normalize_label(field_node.get("name", ""))
                )
                if namespace is None:
                    continue
                target_fields[index] = namespace
                shared_items = field_node.find(qn("sharedItems"))
                if shared_items is None:
                    continue
                for item in shared_items:
                    value = item.get("v")
                    if not value or should_skip_direct_value(value):
                        continue
                    item.set("v", cipher.encrypt(namespace, value))
                    if local_name(item.tag) != "s":
                        item.tag = qn("s")
                    stats.record(f"pivot-{namespace}")
                shared_items.set("containsString", "1")
                for attribute in (
                    "containsNumber",
                    "containsInteger",
                    "containsMixedTypes",
                    "containsSemiMixedTypes",
                    "minValue",
                    "maxValue",
                ):
                    shared_items.attrib.pop(attribute, None)
        parts[definition_path] = xml_bytes(root)
        records_path = pivot_record_path(parts, definition_path)
        if not records_path or records_path not in parts or not target_fields:
            continue
        records = parse_xml(parts[records_path])
        for row in records:
            for index, item in enumerate(list(row)):
                namespace = target_fields.get(index)
                if namespace is None or local_name(item.tag) == "x":
                    continue
                value = item.get("v")
                if not value or should_skip_direct_value(value):
                    continue
                item.set("v", cipher.encrypt(namespace, value))
                item.tag = qn("s")
                stats.record(f"pivot-record-{namespace}")
        parts[records_path] = xml_bytes(records)


def decrypt_pivot_caches(
    parts: dict[str, bytes], cipher: CellCipher, stats: TransformStats
) -> None:
    paths = sorted(
        path
        for path in parts
        if re.fullmatch(
            r"xl/pivotCache/(?:pivotCacheDefinition|pivotCacheRecords)\d+\.xml", path
        )
    )
    for path in paths:
        root = parse_xml(parts[path])
        changed = False
        for element in root.iter():
            value = element.get("v")
            if value and CellCipher.is_encrypted(value):
                namespace = CellCipher.namespace(value)
                element.set("v", cipher.decrypt(value))
                stats.record(f"pivot-{namespace}")
                changed = True
        if changed:
            parts[path] = xml_bytes(root)


def validate_package_limits(path: Path) -> None:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            if len(infos) > MAX_PACKAGE_MEMBERS:
                raise ScrubberError("Workbook contains too many package members")
            if any(info.flag_bits & 0x1 for info in infos):
                raise ScrubberError("Password-protected XLSX packages are not supported")
            total_size = sum(info.file_size for info in infos)
            if total_size > MAX_UNCOMPRESSED_BYTES:
                raise ScrubberError("Workbook expands beyond the safe processing limit")
            names = {info.filename for info in infos}
            required = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
            if not required.issubset(names):
                raise ScrubberError("Input is not a valid XLSX workbook package")
    except zipfile.BadZipFile as exc:
        raise ScrubberError("Input is not a readable XLSX workbook") from exc


def load_package(path: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo]]:
    """Read every package member into memory.

    ``zipfile.testzip()`` used to run here first. It decompresses every member
    to verify its CRC, and each member was then read again, which doubled the
    decompression work in this function. ``archive.read`` already validates the
    CRC of each member it returns and raises :class:`zipfile.BadZipFile` on a
    corrupt member, so the extra pass added nothing but latency.
    """
    validate_package_limits(path)
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            parts = {info.filename: archive.read(info.filename) for info in infos}
    except (zipfile.BadZipFile, KeyError, RuntimeError) as exc:
        raise ScrubberError("Input is not a readable XLSX workbook") from exc
    return parts, infos


def write_package_atomic(
    destination: Path,
    parts: dict[str, bytes],
    infos: list[zipfile.ZipInfo],
    force: bool = False,
) -> None:
    if destination.exists() and not force:
        raise ScrubberError("Output already exists")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{destination.stem}.",
            suffix=".tmp.xlsx",
            dir=destination.parent,
            delete=False,
        ) as temporary:
            temporary_name = temporary.name
        with zipfile.ZipFile(temporary_name, "w") as archive:
            written = set()
            for info in infos:
                archive.writestr(info, parts[info.filename])
                written.add(info.filename)
            for name in sorted(set(parts) - written):
                archive.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
        os.replace(temporary_name, destination)
        temporary_name = None
    finally:
        if temporary_name:
            try:
                os.unlink(temporary_name)
            except FileNotFoundError:
                pass


def transform_workbook(
    source: Path,
    destination: Path,
    cipher: CellCipher | None,
    mode: str,
    force: bool = False,
    selected_columns: set[str] | None = None,
    replacements: list[ReplacementRule] | None = None,
) -> TransformStats:
    """Process *source* workbook into *destination*.

    *selected_columns* is forwarded to :func:`encrypt_worksheet_cells` when
    *mode* is ``"encrypt"``.  Pass ``None`` only when the caller intentionally
    wants every discovered worksheet column encrypted.

    Two deliberate performance properties, both verified by profiling before
    and after:

    * The column-descriptor map is resolved **once** and reused for both the
      pivot-cache namespace lookup and the worksheet encryption pass.  Each
      resolution walks every cell of every sheet, so doing it twice cost about
      30% of total runtime on a single-sheet 120k-cell workbook.
    * Only worksheet parts that were actually modified are serialized again.
      ``ET.tostring`` was the single largest cost at ~44% of runtime, and a
      multi-sheet workbook where the user encrypts one sheet was paying it for
      every sheet.  Leaving untouched parts byte-identical is also better for
      downstream diffing and keeps the output stable.
    """
    if source.suffix.casefold() != ".xlsx" or destination.suffix.casefold() != ".xlsx":
        raise ScrubberError("Both input and output must use the .xlsx extension")
    if not source.is_file():
        raise ScrubberError("Input workbook does not exist")
    if source.resolve() == destination.resolve():
        raise ScrubberError("Input and output paths must be different")

    parts, infos = load_package(source)
    shared_root, shared_strings = read_shared_strings(parts)
    sheets, sheet_paths = workbook_sheets(parts)
    stats = TransformStats()
    changed_sheets: set[str] = set()

    def refresh_shared_strings() -> None:
        """Re-read the in-memory string list after the table was edited."""
        nonlocal shared_strings
        if shared_root is not None:
            shared_strings = shared_string_values(shared_root)

    if mode == "encrypt":
        if cipher is None:
            raise ScrubberError("Encryption key is not available")
        active_replacements = replacements or []
        if active_replacements:
            replace_shared_string_table(shared_root, active_replacements, stats)
            refresh_shared_strings()
            changed_sheets |= replace_worksheet_cells(
                sheets, shared_strings, active_replacements, stats
            )
            replace_pivot_cache_values(parts, active_replacements, stats)

        # Resolved once and shared by the pivot-cache lookup and the cell pass.
        descriptors = selected_column_descriptors(
            sheets, shared_strings, selected_columns
        )
        selected_header_namespaces = {
            normalize_label(descriptor.header): token_namespace_from_header(
                descriptor.header, descriptor.column
            )
            for descriptor in descriptors.values()
        }
        changed_sheets |= encrypt_worksheet_cells(
            sheets, shared_strings, cipher, stats, descriptors
        )
        encrypt_pivot_caches(parts, cipher, stats, selected_header_namespaces)

        # Compacting the table can renumber the indices that worksheets store,
        # so any sheet holding shared-string cells has to be rewritten too.
        if rebuild_shared_string_table(shared_root, sheets.values()):
            changed_sheets.update(sheets)
    elif mode == "decrypt":
        if cipher is None:
            raise ScrubberError("Encryption key is not available")
        decrypt_shared_string_table(shared_root, cipher, stats)
        refresh_shared_strings()
        changed_sheets |= decrypt_worksheet_cells(sheets, shared_strings, cipher, stats)
        decrypt_pivot_caches(parts, cipher, stats)
    elif mode == "replace":
        active_replacements = replacements or []
        if not active_replacements:
            raise ScrubberError("At least one replacement rule is required")
        # Shared-string cells store only an index, so editing the table does not
        # require rewriting the worksheet parts that point at it.
        replace_shared_string_table(shared_root, active_replacements, stats)
        refresh_shared_strings()
        changed_sheets |= replace_worksheet_cells(
            sheets, shared_strings, active_replacements, stats
        )
        replace_pivot_cache_values(parts, active_replacements, stats)
    else:
        raise ScrubberError("Unsupported processing mode")

    if mode == "replace":
        if stats.replacements == 0:
            raise ScrubberError("No replacement values were found in the workbook")
    elif stats.total == 0:
        action = "encrypt" if mode == "encrypt" else "decrypt"
        raise ScrubberError(f"No values were found to {action}")

    for name, path in sheet_paths.items():
        if name in changed_sheets:
            parts[path] = xml_bytes(sheets[name])
    if shared_root is not None:
        parts["xl/sharedStrings.xml"] = xml_bytes(shared_root)
    write_package_atomic(destination, parts, infos, force)
    return stats
