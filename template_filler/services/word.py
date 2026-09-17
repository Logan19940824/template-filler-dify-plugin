from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
from pathlib import Path
import re
from typing import Any
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from lxml import etree as ET

from .common import TOKEN_RE, array_fields, expand_contexts, placeholder_names, render_context, value_defaults


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W_T = f"{{{W_NS}}}t"
W_P = f"{{{W_NS}}}p"
W_TBL = f"{{{W_NS}}}tbl"
W_TR = f"{{{W_NS}}}tr"
W_BODY = f"{{{W_NS}}}body"
W_TC = f"{{{W_NS}}}tc"
W_R = f"{{{W_NS}}}r"
W_RPR = f"{{{W_NS}}}rPr"

def _load_document(path: Path) -> tuple[dict[str, bytes], ET.Element]:
    try:
        with ZipFile(path) as archive:
            entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        document_xml = entries["word/document.xml"]
        root = ET.fromstring(document_xml, parser=ET.XMLParser(resolve_entities=False, no_network=True))
        if root.getroottree().docinfo.doctype:
            raise ValueError("DOCX XML must not contain a DTD")
    except (BadZipFile, KeyError, ET.ParseError, OSError, ValueError) as exc:
        raise ValueError("template is not a valid Word DOCX file or is corrupted") from exc
    return entries, root


def _text_nodes(element: ET.Element) -> list[ET.Element]:
    return [node for node in element.iter(W_T)
            if element.tag != W_P or next(node.iterancestors(W_P), None) is element]


def _paragraph_text(paragraph: ET.Element) -> str:
    controls = {f"{{{W_NS}}}tab": "\t", f"{{{W_NS}}}br": "\n", f"{{{W_NS}}}cr": "\n"}
    return "".join(
        (node.text or "") if node.tag == W_T else controls.get(node.tag, "")
        for node in paragraph.iter()
        if paragraph.tag != W_P or next(node.iterancestors(W_P), None) is paragraph
    )


def _editable_paragraph(paragraph: ET.Element) -> bool:
    allowed = {W_R, f"{{{W_NS}}}pPr", f"{{{W_NS}}}bookmarkStart", f"{{{W_NS}}}bookmarkEnd", f"{{{W_NS}}}proofErr"}
    return all(child.tag in allowed for child in paragraph) and all(
        child.tag in {W_RPR, W_T} for run in paragraph.findall(W_R) for child in run
    ) and not any(
        node.tag in {f"{{{W_NS}}}rPrChange", f"{{{W_NS}}}pPrChange"} for node in paragraph.iter()
    )


def _empty_row(row: ET.Element) -> bool:
    # Only plain whitespace is disposable; fields, drawings, bookmarks and controls are content.
    if _paragraph_text(row).strip() or any(child.tag not in {W_TC, f"{{{W_NS}}}trPr", f"{{{W_NS}}}tblPrEx"} for child in row):
        return False
    for cell in row.findall(W_TC):
        if any(child.tag not in {W_P, f"{{{W_NS}}}tcPr"} for child in cell):
            return False
        for paragraph in cell.findall(W_P):
            if not _editable_paragraph(paragraph) or any(child.tag not in {W_R, f"{{{W_NS}}}pPr"} for child in paragraph):
                return False
    return not any(node.tag in {f"{{{W_NS}}}{tag}" for tag in ("trPrChange", "tcPrChange", "cellIns", "cellDel", "cellMerge")} for node in row.iter())


def _save_document(entries: dict[str, bytes], root: ET.Element, output: Path) -> None:
    entries["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)


def _set_text(node: ET.Element, text: str) -> None:
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set(f"{{{XML_NS}}}space", "preserve")
    else:
        node.attrib.pop(f"{{{XML_NS}}}space", None)


def _replace_text_range(paragraph: ET.Element, start: int, end: int, replacement: str) -> None:
    if not _editable_paragraph(paragraph):
        raise ValueError("cannot edit a paragraph containing unsupported content")
    nodes = _text_nodes(paragraph)
    text = "".join(node.text or "" for node in nodes)
    if not 0 <= start <= end <= len(text):
        raise ValueError("text range is outside the paragraph")
    if not nodes:
        run = ET.SubElement(paragraph, W_R)
        node = ET.SubElement(run, W_T)
        _set_text(node, replacement)
        return

    positions = []
    offset = 0
    for node in nodes:
        value = node.text or ""
        positions.append((node, offset, offset + len(value)))
        offset += len(value)
    if start == end:
        for node, node_start, node_end in positions:
            if node_start < start <= node_end or (start == 0 and node_end > 0):
                relative = start - node_start
                value = node.text or ""
                _set_text(node, value[:relative] + replacement + value[relative:])
                return
        _set_text(nodes[-1], (nodes[-1].text or "") + replacement)
        return

    first = True
    for node, node_start, node_end in positions:
        if node_end <= start or node_start >= end:
            continue
        value = node.text or ""
        left = value[: max(0, start - node_start)]
        right = value[max(0, end - node_start) :] if end < node_end else ""
        _set_text(node, left + (replacement if first else "") + right)
        first = False


def _replace_matches(paragraph: ET.Element, values: dict[str, Any], context: dict[str, Any]) -> None:
    text = _paragraph_text(paragraph)
    for match in reversed(list(TOKEN_RE.finditer(text))):
        replacement = render_context(match.group(0), values, context)
        if replacement != match.group(0):
            _replace_text_range(paragraph, match.start(), match.end(), replacement)


def _replace_paragraph(paragraph: ET.Element, values: dict[str, Any]) -> None:
    _replace_matches(paragraph, values, {})


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


def _paragraph_details(paragraph: ET.Element, location: dict[str, int]) -> dict[str, Any]:
    text = _paragraph_text(paragraph)
    runs = []
    offset = 0
    for run in paragraph.findall(W_R) if _editable_paragraph(paragraph) else []:
        run_text = "".join(node.text or "" for node in run.iter(W_T))
        if not run_text:
            continue
        properties = run.find(f"./{W_RPR}")
        item: dict[str, Any] = {"text": run_text, "start": offset, "end": offset + len(run_text)}
        if properties is not None:
            bold = properties.find(f"./{{{W_NS}}}b")
            underline = properties.find(f"./{{{W_NS}}}u")
            if bold is not None:
                item["bold"] = bold.get(f"{{{W_NS}}}val", "true") not in {"0", "false", "none"}
            if underline is not None:
                item["underline"] = underline.get(f"{{{W_NS}}}val", "single")
        runs.append(item)
        offset += len(run_text)
    return {
        "location": location,
        "text": text,
        "editable": _editable_paragraph(paragraph),
        "runs": runs,
        "blank_spans": [{"start": match.start(), "end": match.end()} for match in re.finditer(r"\s+", text)],
    }


def inspect_word_document(path: Path) -> dict[str, Any]:
    entries, root = _load_document(path)
    body = root.find(f"./{W_BODY}")
    if body is None:
        raise ValueError("template is not a valid Word DOCX file or is corrupted")

    paragraphs = [_paragraph_details(paragraph, location) for paragraph, location in _iter_body_paragraphs(body)]
    tables = []
    for table_index, table in enumerate(child for child in list(body) if child.tag == W_TBL):
        rows = []
        for row_index, row in enumerate(table.findall(f"./{W_TR}")):
            cells = []
            for cell_index, cell in enumerate(row.findall(f"./{W_TC}")):
                properties = cell.find(f"./{{{W_NS}}}tcPr")
                grid_span = properties.find(f"./{{{W_NS}}}gridSpan") if properties is not None else None
                vertical_merge = properties.find(f"./{{{W_NS}}}vMerge") if properties is not None else None
                cell_paragraphs = [
                    _paragraph_details(
                        paragraph,
                        {"table": table_index, "row": row_index, "cell": cell_index, "paragraph": paragraph_index},
                    )
                    for paragraph_index, paragraph in enumerate(cell.findall(f"./{W_P}"))
                ]
                cells.append({
                    "cell": cell_index,
                    "grid_span": int(grid_span.get(f"{{{W_NS}}}val", "1")) if grid_span is not None else 1,
                    "vertical_merge": vertical_merge.get(f"{{{W_NS}}}val", "continue") if vertical_merge is not None else None,
                    "paragraphs": cell_paragraphs,
                })
            before = row.find(f"./{{{W_NS}}}trPr/{{{W_NS}}}gridBefore")
            grid_column = int(before.get(f"{{{W_NS}}}val", "0")) if before is not None else 0
            for cell in cells:
                cell["grid_column"] = grid_column
                grid_column += cell["grid_span"]
                if cell["vertical_merge"] == "continue":
                    for paragraph in cell["paragraphs"]:
                        paragraph["editable"] = False
            rows.append({"row": row_index, "is_empty": _empty_row(row), "cells": cells})
        tables.append({"table": table_index, "grid_widths_twips": [
            int(column.get(f"{{{W_NS}}}w", "0"))
            for column in table.findall(f"./{{{W_NS}}}tblGrid/{{{W_NS}}}gridCol")
        ], "rows": rows})

    unsupported = []
    tags = {node.tag for node in body.iter()}
    checks = {
        "text_boxes": ("txbxContent",),
        "drawings": ("drawing", "pict", "object"),
        "fields": ("fldChar", "fldSimple", "instrText"),
        "tracked_changes": ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "trPrChange", "tcPrChange", "cellIns", "cellDel", "cellMerge"),
        "content_controls": ("sdt",),
        "hyperlinks": ("hyperlink",),
        "tabs_or_breaks": ("tab", "br", "cr"),
        "notes": ("footnoteReference", "endnoteReference"),
    }
    for name, local_names in checks.items():
        if any(f"{{{W_NS}}}{tag}" in tags for tag in local_names):
            unsupported.append(name)
    if any(next(table.iterancestors(W_TBL), None) is not None for table in body.iter(W_TBL)):
        unsupported.append("nested_tables")
    if any(child.tag not in {W_P, W_TBL, f"{{{W_NS}}}sectPr"} for child in body):
        unsupported.append("other_body_blocks")
    for part, name in (("word/header", "headers"), ("word/footer", "footers"), ("word/comments.xml", "comments")):
        if any(entry.startswith(part) for entry in entries):
            unsupported.append(name)
    blocks = []
    paragraph_index = table_index = 0
    for child in body:
        if child.tag == W_P:
            blocks.append({"type": "paragraph", "paragraph": paragraph_index})
            paragraph_index += 1
        elif child.tag == W_TBL:
            blocks.append({"type": "table", "table": table_index})
            table_index += 1
    return {
        "file_type": "docx",
        "document_sha256": sha256(path.read_bytes()).hexdigest(),
        "blocks": blocks,
        "paragraphs": paragraphs,
        "tables": tables,
        "unsupported_features": unsupported,
    }


def _paragraph_at(body: ET.Element, location: dict[str, Any]) -> ET.Element:
    if not all(type(value) is int and value >= 0 for value in location.values()):
        raise ValueError(f"paragraph location not found: {location}")
    if set(location) == {"paragraph"}:
        paragraphs = [child for child in list(body) if child.tag == W_P]
        index = location["paragraph"]
        if index < len(paragraphs):
            return paragraphs[index]
    elif set(location) == {"table", "row", "cell", "paragraph"}:
        tables = [child for child in list(body) if child.tag == W_TBL]
        try:
            row = tables[location["table"]].findall(f"./{W_TR}")[location["row"]]
            cell = row.findall(f"./{W_TC}")[location["cell"]]
            merge = cell.find(f"./{{{W_NS}}}tcPr/{{{W_NS}}}vMerge")
            if merge is not None and merge.get(f"{{{W_NS}}}val", "continue") != "restart":
                raise ValueError("cannot edit a vertical-merge continuation cell")
            return cell.findall(f"./{W_P}")[location["paragraph"]]
        except (IndexError, TypeError):
            pass
    raise ValueError(f"paragraph location not found: {location}")


def insert_word_placeholders(path: Path, output: Path, payload: dict[str, Any]) -> None:
    if not isinstance(payload, dict) or not isinstance(payload.get("operations"), list) or not payload["operations"]:
        raise ValueError("operations must be a non-empty JSON array")
    if payload.get("document_sha256") != sha256(path.read_bytes()).hexdigest():
        raise ValueError("document_sha256 does not match the inspected document")

    entries, root = _load_document(path)
    body = root.find(f"./{W_BODY}")
    if body is None:
        raise ValueError("template is not a valid Word DOCX file or is corrupted")
    replacements: dict[int, tuple[ET.Element, list[tuple[int, int, str]]]] = {}
    deletions: list[tuple[ET.Element, ET.Element]] = []
    replaced_rows: set[tuple[int, int]] = set()
    deleted_rows: set[tuple[int, int]] = set()
    for operation in payload["operations"]:
        if not isinstance(operation, dict):
            raise ValueError("each operation must be a JSON object")
        operation_type = operation.get("type")
        if operation_type == "replace_text_range":
            location = operation.get("location")
            if not isinstance(location, dict):
                raise ValueError("replace_text_range requires a location object")
            paragraph = _paragraph_at(body, location)
            if not _editable_paragraph(paragraph):
                raise ValueError(f"cannot edit unsupported paragraph at {location}")
            text = _paragraph_text(paragraph)
            start, end, placeholder = operation.get("start"), operation.get("end"), operation.get("placeholder")
            if operation.get("expected_text") != text:
                raise ValueError(f"expected text mismatch at {location}")
            if type(start) is not int or type(end) is not int or not 0 <= start <= end <= len(text):
                raise ValueError(f"invalid text range at {location}")
            match = TOKEN_RE.fullmatch(placeholder) if isinstance(placeholder, str) else None
            if match is None or not match.group(1).strip() or any(char in placeholder for char in "\r\n\t"):
                raise ValueError(f"invalid placeholder: {placeholder}")
            replacements.setdefault(id(paragraph), (paragraph, []))[1].append((start, end, placeholder))
            if "table" in location:
                replaced_rows.add((location["table"], location["row"]))
        elif operation_type == "delete_empty_rows":
            table_index, row_indexes = operation.get("table"), operation.get("rows")
            tables = [child for child in list(body) if child.tag == W_TBL]
            if type(table_index) is not int or not 0 <= table_index < len(tables) or not isinstance(row_indexes, list) or not row_indexes:
                raise ValueError("delete_empty_rows requires a valid table and rows array")
            table = tables[table_index]
            rows = table.findall(f"./{W_TR}")
            for row_index in row_indexes:
                if type(row_index) is not int or not 0 <= row_index < len(rows):
                    raise ValueError(f"row not found: table {table_index}, row {row_index}")
                row = rows[row_index]
                if not _empty_row(row):
                    raise ValueError(f"cannot delete non-empty or complex row: table {table_index}, row {row_index}")
                if row.find(f".//{{{W_NS}}}gridSpan") is not None or row.find(f".//{{{W_NS}}}vMerge") is not None:
                    raise ValueError(f"cannot delete merged row: table {table_index}, row {row_index}")
                if (table_index, row_index) in deleted_rows:
                    raise ValueError(f"duplicate row deletion: table {table_index}, row {row_index}")
                deleted_rows.add((table_index, row_index))
                deletions.append((table, row))
        else:
            raise ValueError(f"unsupported operation type: {operation_type}")

    if replaced_rows & deleted_rows:
        raise ValueError("cannot replace text in a row selected for deletion")
    for table_index in {table for table, _ in deleted_rows}:
        table = [child for child in body if child.tag == W_TBL][table_index]
        if sum(index == table_index for index, _ in deleted_rows) == len(table.findall(W_TR)):
            raise ValueError("cannot delete every row of a table")
    for paragraph, ranges in replacements.values():
        ascending = sorted(ranges)
        for left, right in zip(ascending, ascending[1:]):
            if left[1] > right[0] or left[0] == right[0]:
                raise ValueError("overlapping text ranges are not allowed")
    for paragraph, ranges in replacements.values():
        ascending = sorted(ranges)
        for start, end, placeholder in reversed(ascending):
            _replace_text_range(paragraph, start, end, placeholder)
    for table, row in reversed(deletions):
        table.remove(row)

    _save_document(entries, root, output)


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
            _replace_matches(target, values, context)


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
                for paragraph in [p for cell in target.findall(f"./{W_TC}") for p in cell.findall(f"./{W_P}")]:
                    _replace_matches(paragraph, values, context)


def fill_word(path: Path, output: Path, values: dict[str, Any]) -> None:
    entries, root = _load_document(path)
    body = root.find(f"./{W_BODY}")
    if body is None:
        raise ValueError("template is not a valid Word DOCX file or is corrupted")
    _fill_body(body, values)
    _fill_tables(body, values)
    _save_document(entries, root, output)
