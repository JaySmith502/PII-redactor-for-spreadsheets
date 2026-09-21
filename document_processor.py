import os
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

try:
    from .processor import ReplacementRule, ScrubberError, replacement_value
except ImportError:
    from processor import ReplacementRule, ScrubberError, replacement_value


MAX_PACKAGE_MEMBERS = 10_000
MAX_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PDF_MIN_FONT_SIZE = 4.0
PDF_MAX_FONT_SIZE = 72.0
PDF_DEFAULT_FONT_SIZE = 10.0


def qn(local_name: str, namespace: str = WORD_NS) -> str:
    return f"{{{namespace}}}{local_name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def validate_docx_package(path: Path) -> None:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            if len(infos) > MAX_PACKAGE_MEMBERS:
                raise ScrubberError("Document contains too many package members")
            if any(info.flag_bits & 0x1 for info in infos):
                raise ScrubberError("Password-protected DOCX packages are not supported")
            if sum(info.file_size for info in infos) > MAX_UNCOMPRESSED_BYTES:
                raise ScrubberError("Document expands beyond the safe processing limit")
            names = {info.filename for info in infos}
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise ScrubberError("Input is not a valid DOCX document")
    except zipfile.BadZipFile as exc:
        raise ScrubberError("Input is not a readable DOCX document") from exc


def load_docx_package(path: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo]]:
    validate_docx_package(path)
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad_member = archive.testzip()
            if bad_member:
                raise ScrubberError("Document contains a corrupt package member")
            infos = archive.infolist()
            parts = {info.filename: archive.read(info.filename) for info in infos}
    except (zipfile.BadZipFile, KeyError, RuntimeError) as exc:
        raise ScrubberError("Input is not a readable DOCX document") from exc
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
            suffix=".tmp.docx",
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


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def paragraph_text_nodes(root: ET.Element) -> list[list[ET.Element]]:
    groups: list[list[ET.Element]] = []
    for paragraph in root.iter(qn("p")):
        nodes = [node for node in paragraph.iter(qn("t"))]
        if nodes:
            groups.append(nodes)
    return groups


def replace_docx_xml(data: bytes, replacements: list[ReplacementRule]) -> tuple[bytes, int]:
    root = ET.fromstring(data)
    replacements_applied = 0

    for text in root.iter(qn("t")):
        if text.text is None:
            continue
        updated, count = replacement_value(text.text, replacements)
        if count:
            text.text = updated
            replacements_applied += count

    for nodes in paragraph_text_nodes(root):
        combined = "".join(node.text or "" for node in nodes)
        updated, count = replacement_value(combined, replacements)
        if not count or updated == combined:
            continue
        nodes[0].text = updated
        for node in nodes[1:]:
            node.text = ""
        replacements_applied += count

    return xml_bytes(root), replacements_applied


def is_docx_xml_part(path: str) -> bool:
    normalized = path.replace("\\", "/")
    if not normalized.startswith("word/") or not normalized.endswith(".xml"):
        return False
    name = posixpath.basename(normalized)
    return (
        name == "document.xml"
        or name.startswith("header")
        or name.startswith("footer")
        or name.startswith("footnotes")
        or name.startswith("endnotes")
        or name.startswith("comments")
    )


def replace_docx(
    source: Path,
    destination: Path,
    replacements: list[ReplacementRule],
    force: bool = False,
) -> int:
    if source.suffix.casefold() != ".docx" or destination.suffix.casefold() != ".docx":
        raise ScrubberError("Both input and output must use the .docx extension")
    if not source.is_file():
        raise ScrubberError("Input document does not exist")
    if source.resolve() == destination.resolve():
        raise ScrubberError("Input and output paths must be different")

    parts, infos = load_docx_package(source)
    total = 0
    for path, data in list(parts.items()):
        if not is_docx_xml_part(path):
            continue
        try:
            updated, count = replace_docx_xml(data, replacements)
        except ET.ParseError as exc:
            raise ScrubberError("DOCX document XML could not be read") from exc
        if count:
            parts[path] = updated
            total += count

    if total == 0:
        raise ScrubberError("No replacement values were found in the document")
    write_package_atomic(destination, parts, infos, force)
    return total


def pdf_color_tuple(color: int | None) -> tuple[float, float, float]:
    if not isinstance(color, int):
        return (0, 0, 0)
    red = ((color >> 16) & 0xFF) / 255
    green = ((color >> 8) & 0xFF) / 255
    blue = (color & 0xFF) / 255
    return (red, green, blue)


def pdf_base14_font_name(font_name: str | None) -> str:
    normalized = (font_name or "").casefold()
    if "courier" in normalized:
        return "cour"
    if "times" in normalized:
        return "tiro"
    if "symbol" in normalized:
        return "symb"
    if "zapf" in normalized or "ding" in normalized:
        return "zadb"
    return "helv"


def pdf_intersection_area(first, second) -> float:
    x0 = max(first.x0, second.x0)
    y0 = max(first.y0, second.y0)
    x1 = min(first.x1, second.x1)
    y1 = min(first.y1, second.y1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    return float((x1 - x0) * (y1 - y0))


def pdf_text_style(page, rectangle) -> tuple[float, tuple[float, float, float], str, float]:
    fallback_size = max(
        PDF_MIN_FONT_SIZE,
        min(PDF_MAX_FONT_SIZE, float(rectangle.height) * 0.78),
    )
    fallback_baseline = rectangle.y1 - max(0.0, (float(rectangle.height) - fallback_size) / 2)
    best_span = None
    best_score = 0.0
    text_data = page.get_text("dict")
    for block in text_data.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                span_rect = fitz.Rect(span.get("bbox"))
                score = pdf_intersection_area(rectangle, span_rect)
                if score > best_score:
                    best_span = span
                    best_score = score

    if best_span is None:
        return (fallback_size, (0, 0, 0), "helv", fallback_baseline)

    size = float(best_span.get("size") or fallback_size)
    size = max(PDF_MIN_FONT_SIZE, min(PDF_MAX_FONT_SIZE, size))
    color = pdf_color_tuple(best_span.get("color"))
    font_name = pdf_base14_font_name(best_span.get("font"))
    origin = best_span.get("origin")
    if isinstance(origin, (list, tuple)) and len(origin) >= 2:
        baseline = float(origin[1])
    else:
        baseline = rectangle.y1 - max(0.0, (float(rectangle.height) - size) / 2)
    return (size, color, font_name, baseline)


def replace_pdf(
    source: Path,
    destination: Path,
    replacements: list[ReplacementRule],
    force: bool = False,
) -> int:
    if fitz is None:
        raise ScrubberError("PDF replacement requires the local PyMuPDF package")
    if source.suffix.casefold() != ".pdf" or destination.suffix.casefold() != ".pdf":
        raise ScrubberError("Both input and output must use the .pdf extension")
    if not source.is_file():
        raise ScrubberError("Input PDF does not exist")
    if source.resolve() == destination.resolve():
        raise ScrubberError("Input and output paths must be different")
    if destination.exists() and not force:
        raise ScrubberError("Output already exists")

    total = 0
    try:
        document = fitz.open(source)
    except Exception as exc:
        raise ScrubberError("Input is not a readable PDF document") from exc

    try:
        if document.needs_pass:
            raise ScrubberError("Password-protected PDF documents are not supported")

        for page in document:
            page_changed = False
            overlays = []
            for rule in replacements:
                rectangles = page.search_for(rule.find)
                if not rectangles:
                    continue
                total += len(rectangles)
                page_changed = True
                for rectangle in rectangles:
                    font_size, text_color, font_name, baseline = pdf_text_style(
                        page, rectangle
                    )
                    page.add_redact_annot(
                        rectangle,
                        fill=(1, 1, 1),
                        cross_out=False,
                    )
                    overlays.append(
                        (rectangle.x0, baseline, rule.replace, font_size, text_color, font_name)
                    )
            if page_changed:
                page.apply_redactions()
                for x, baseline, text, font_size, text_color, font_name in overlays:
                    page.insert_text(
                        (x, baseline),
                        text,
                        fontsize=font_size,
                        fontname=font_name,
                        color=text_color,
                    )

        if total == 0:
            raise ScrubberError("No replacement values were found in the PDF")

        destination.parent.mkdir(parents=True, exist_ok=True)
        document.save(destination, garbage=4, deflate=True)
    finally:
        document.close()

    return total


def replace_text_document(
    source: Path,
    destination: Path,
    replacements: list[ReplacementRule],
    force: bool = False,
) -> int:
    suffix = source.suffix.casefold()
    if suffix == ".docx":
        return replace_docx(source, destination, replacements, force)
    if suffix == ".pdf":
        return replace_pdf(source, destination, replacements, force)
    raise ScrubberError("Only .docx and .pdf documents are supported for replacement")
