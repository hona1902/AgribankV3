from __future__ import annotations

from pathlib import Path
from io import BytesIO
import re
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


PLACEHOLDER_PATTERN = re.compile(r"\[[^\[\]]+\]")
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
TEXT_TAG = f"{{{WORD_NAMESPACE}}}t"
PARAGRAPH_TAG = f"{{{WORD_NAMESPACE}}}p"

ET.register_namespace("w", WORD_NAMESPACE)


class WordTemplateError(RuntimeError):
    pass


def scan_word_placeholders(template_path: Path) -> set[str]:
    """Return placeholders found in Word paragraphs, including split runs."""
    if not Path(template_path).is_file():
        raise WordTemplateError(f"Không tìm thấy file mẫu Word: {template_path}")
    placeholders: set[str] = set()
    with ZipFile(template_path) as archive:
        for name in _word_xml_names(archive):
            root = ET.fromstring(archive.read(name))
            for paragraph in root.iter(PARAGRAPH_TAG):
                text = _paragraph_text(paragraph)
                placeholders.update(PLACEHOLDER_PATTERN.findall(text))
    return placeholders


def replace_word_placeholders(
    template_path: Path,
    output_path: Path,
    replacements: dict[str, object],
    *,
    clear_unmapped: bool = True,
) -> set[str]:
    """Create a docx copy with placeholders replaced and return unmapped placeholders."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    if not template_path.is_file():
        raise WordTemplateError(f"Không tìm thấy file mẫu Word: {template_path}")

    placeholder_values = {
        _normalize_placeholder_key(key): "" if value is None else str(value)
        for key, value in replacements.items()
    }
    unmapped: set[str] = set()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with ZipFile(template_path, "r") as source, ZipFile(
        output_path, "w", compression=ZIP_DEFLATED
    ) as target:
        xml_names = set(_word_xml_names(source))
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in xml_names:
                data, missing = _replace_placeholders_in_xml(
                    data,
                    placeholder_values,
                    clear_unmapped=clear_unmapped,
                )
                unmapped.update(missing)
            target.writestr(item, data)
    return unmapped


def extract_docx_text(path: Path) -> str:
    """Return concatenated Word XML paragraph text for tests and diagnostics."""
    chunks: list[str] = []
    with ZipFile(path) as archive:
        for name in _word_xml_names(archive):
            root = ET.fromstring(archive.read(name))
            for paragraph in root.iter(PARAGRAPH_TAG):
                text = _paragraph_text(paragraph)
                if text:
                    chunks.append(text)
    return "\n".join(chunks)


def _word_xml_names(archive: ZipFile) -> list[str]:
    prefixes = (
        "word/document.xml",
        "word/header",
        "word/footer",
        "word/footnotes.xml",
        "word/endnotes.xml",
    )
    return [
        name
        for name in archive.namelist()
        if name.endswith(".xml") and any(name.startswith(prefix) for prefix in prefixes)
    ]


def _replace_placeholders_in_xml(
    data: bytes,
    replacements: dict[str, str],
    *,
    clear_unmapped: bool,
) -> tuple[bytes, set[str]]:
    namespaces = _register_source_namespaces(data)
    root = ET.fromstring(data)
    unmapped: set[str] = set()
    changed = False
    for paragraph in root.iter(PARAGRAPH_TAG):
        text_nodes = [node for node in paragraph.iter(TEXT_TAG)]
        if not text_nodes:
            continue
        original_text = "".join(node.text or "" for node in text_nodes)
        replaced_text, missing = _replace_text(original_text, replacements, clear_unmapped)
        unmapped.update(missing)
        if replaced_text != original_text:
            text_nodes[0].text = replaced_text
            for node in text_nodes[1:]:
                node.text = ""
            changed = True
    if not changed:
        return data, unmapped
    serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return _restore_root_namespace_declarations(serialized, namespaces), unmapped


def _register_source_namespaces(data: bytes) -> dict[str, str]:
    # Word relies on prefixes named in mc:Ignorable. Keep source prefixes stable
    # when ElementTree serializes modified XML back into the docx package.
    namespaces: dict[str, str] = {}
    for _, item in ET.iterparse(BytesIO(data), events=("start-ns",)):
        prefix, uri = item
        namespaces[prefix] = uri
        ET.register_namespace(prefix, uri)
    return namespaces


def _restore_root_namespace_declarations(data: bytes, namespaces: dict[str, str]) -> bytes:
    text = data.decode("utf-8")
    document_start = text.find("<w:")
    if document_start < 0:
        document_start = text.find("<")
    tag_end = text.find(">", document_start)
    if document_start < 0 or tag_end < 0:
        return data
    root_tag = text[document_start:tag_end]
    additions = []
    for prefix, uri in namespaces.items():
        if not prefix:
            continue
        declaration = f"xmlns:{prefix}="
        if declaration not in root_tag:
            additions.append(f' xmlns:{prefix}="{uri}"')
    if not additions:
        return data
    return (text[:tag_end] + "".join(additions) + text[tag_end:]).encode("utf-8")


def _replace_text(
    text: str,
    replacements: dict[str, str],
    clear_unmapped: bool,
) -> tuple[str, set[str]]:
    missing: set[str] = set()

    def repl(match: re.Match[str]) -> str:
        placeholder = match.group(0)
        key = _normalize_placeholder_key(placeholder)
        if key in replacements:
            return replacements[key]
        missing.add(placeholder)
        return "" if clear_unmapped else placeholder

    result = PLACEHOLDER_PATTERN.sub(repl, text)
    den_ngay = replacements.get("DenNgay", "")
    if den_ngay:
        result = re.sub(r"Đến ngày(?:\s|\.|…){2,}", f"Đến ngày {den_ngay}", result)
    return result, missing


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(TEXT_TAG))


def _normalize_placeholder_key(key: str) -> str:
    return key.strip().removeprefix("[").removesuffix("]")
