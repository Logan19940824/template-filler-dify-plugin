from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from .common import array_fields, expand_contexts, placeholder_names, render_context, replace_text, value_defaults


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W_T = f"{{{W_NS}}}t"
W_P = f"{{{W_NS}}}p"
W_TBL = f"{{{W_NS}}}tbl"
W_TR = f"{{{W_NS}}}tr"
W_BODY = f"{{{W_NS}}}body"

ET.register_namespace("w", W_NS)


def _load_document(path: Path) -> tuple[dict[str, bytes], ET.Element]:
    try:
        with ZipFile(path) as archive:
            entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        document_xml = entries["word/document.xml"]
        for _, namespace in ET.iterparse(BytesIO(document_xml), events=("start-ns",)):
            prefix, uri = namespace
            if not prefix.startswith("ns") or not prefix[2:].isdigit():
                ET.register_namespace(prefix, uri)
        root = ET.fromstring(document_xml)
    except (BadZipFile, KeyError, ET.ParseError, OSError, ValueError) as exc:
        raise ValueError("template is not a valid Word DOCX file or is corrupted") from exc
    return entries, root


def _text_nodes(element: ET.Element) -> list[ET.Element]:
    return [node for node in element.iter(W_T)]


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in _text_nodes(paragraph))


def _set_text(node: ET.Element, text: str) -> None:
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set(f"{{{XML_NS}}}space", "preserve")
    else:
        node.attrib.pop(f"{{{XML_NS}}}space", None)


def _set_xml_text(element: ET.Element, text: str) -> None:
    nodes = _text_nodes(element)
    if nodes:
        _set_text(nodes[0], text)
        for node in nodes[1:]:
            _set_text(node, "")


def _replace_paragraph(paragraph: ET.Element, values: dict[str, Any]) -> None:
    text = _paragraph_text(paragraph)
    replaced = replace_text(text, values)
    if replaced == text:
        return
    nodes = _text_nodes(paragraph)
    if nodes:
        _set_text(nodes[0], replaced)
        for node in nodes[1:]:
            _set_text(node, "")


def _iter_body_paragraphs(body: ET.Element):
    paragraph_index = 0
    for child in list(body):
        if child.tag == W_P:
            yield child, {"paragraph": paragraph_index}
            paragraph_index += 1


def _iter_table_paragraphs(table: ET.Element, table_index: int):
    for row_index, row in enumerate(table.findall(f"./{W_TR}")):
        cells = list(row.findall(f"./{{{W_NS}}}tc"))
        for cell_index, cell in enumerate(cells):
            for paragraph_index, paragraph in enumerate(cell.findall(f"./{W_P}")):
                yield paragraph, {"table": table_index, "row": row_index, "cell": cell_index, "paragraph": paragraph_index}


def parse_word(path: Path) -> dict[str, Any]:
    _, root = _load_document(path)
    body = root.find(f"./{W_BODY}")
    if body is None:
        raise ValueError("template is not a valid Word DOCX file or is corrupted")
    found: dict[str, list[dict[str, Any]]] = {}
    tables = [child for child in list(body) if child.tag == W_TBL]
    for paragraph, location in _iter_body_paragraphs(body):
        for name in placeholder_names(_paragraph_text(paragraph)):
            found.setdefault(name, []).append(location)
    for table_index, table in enumerate(tables):
        for paragraph, location in _iter_table_paragraphs(table, table_index):
            for name in placeholder_names(_paragraph_text(paragraph)):
                found.setdefault(name, []).append(location)
    return {"file_type": "docx", "placeholders": [{"name": n, "occurrences": o} for n, o in found.items()], "values": value_defaults(list(found))}


def _fill_body(body: ET.Element, values: dict[str, Any]) -> None:
    for child in list(body):
        if child.tag != W_P:
            continue
        fields = array_fields(_paragraph_text(child))
        if not fields:
            _replace_paragraph(child, values)
            continue
        contexts = expand_contexts(values, fields)
        parent = body
        anchor = child
        template_element = deepcopy(child)
        if not contexts:
            parent.remove(child)
            continue
        for index, context in enumerate(contexts):
            target = child if index == 0 else deepcopy(template_element)
            if index:
                parent.insert(list(parent).index(anchor) + 1, target)
                anchor = target
            text = _paragraph_text(child) if index == 0 else _paragraph_text(target)
            _set_xml_text(target, render_context(text, values, context))


def _fill_tables(body: ET.Element, values: dict[str, Any]) -> None:
    for table in [child for child in list(body) if child.tag == W_TBL]:
        for row in list(table.findall(f"./{W_TR}")):
            cells = list(row.findall(f"./{{{W_NS}}}tc"))
            paragraphs = [paragraph for cell in cells for paragraph in cell.findall(f"./{W_P}")]
            fields = list(dict.fromkeys(field for paragraph in paragraphs for field in array_fields(_paragraph_text(paragraph))))
            if not fields:
                for paragraph in paragraphs:
                    _replace_paragraph(paragraph, values)
                continue
            contexts = expand_contexts(values, fields)
            parent = table
            anchor = row
            template_row = deepcopy(row)
            if not contexts:
                parent.remove(row)
                continue
            for index, context in enumerate(contexts):
                target = row if index == 0 else deepcopy(template_row)
                if index:
                    parent.insert(list(parent).index(anchor) + 1, target)
                    anchor = target
                for node in _text_nodes(target):
                    _set_text(node, render_context(node.text or "", values, context))


def fill_word(path: Path, output: Path, values: dict[str, Any]) -> None:
    entries, root = _load_document(path)
    body = root.find(f"./{W_BODY}")
    if body is None:
        raise ValueError("template is not a valid Word DOCX file or is corrupted")
    _fill_body(body, values)
    _fill_tables(body, values)
    entries["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)
