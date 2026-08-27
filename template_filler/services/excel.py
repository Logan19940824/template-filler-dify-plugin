from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .common import array_fields, expand_contexts, placeholder_names, render_context, replace_text, value_defaults


def _open_excel(path: Path, *, read_only: bool = False):
    try:
        return load_workbook(path, read_only=read_only, data_only=False)
    except Exception as exc:
        raise ValueError("template is not a valid Excel XLSX file or is corrupted") from exc


def parse_excel(path: Path) -> dict[str, Any]:
    workbook = _open_excel(path, read_only=True)
    found: dict[str, list[dict[str, Any]]] = {}
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for name in placeholder_names(cell.value):
                        found.setdefault(name, []).append({"sheet": sheet.title, "cells": [cell.coordinate]})
    return {"file_type": "xlsx", "placeholders": [{"name": n, "occurrences": o} for n, o in found.items()], "values": value_defaults(list(found))}


def _copy_row_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.alignment = copy(source.alignment)
    target.fill = copy(source.fill)
    target.font = copy(source.font)
    target.border = copy(source.border)
    target.protection = copy(source.protection)


def fill_excel(path: Path, output: Path, values: dict[str, Any]) -> None:
    workbook = _open_excel(path)
    for sheet in workbook.worksheets:
        for row_index in range(sheet.max_row, 0, -1):
            row = list(sheet[row_index])
            fields = list(dict.fromkeys(f for cell in row if isinstance(cell.value, str) for f in array_fields(cell.value)))
            if fields:
                contexts = expand_contexts(values, fields)
                snapshots = [(cell.value, cell) for cell in row]
                if not contexts:
                    sheet.delete_rows(row_index)
                    continue
                for offset, context in enumerate(contexts):
                    target_index = row_index + offset
                    if offset:
                        sheet.insert_rows(target_index)
                    for column, (value, source) in enumerate(snapshots, 1):
                        target = sheet.cell(target_index, column)
                        target.value = render_context(value, values, context) if isinstance(value, str) else value
                        _copy_row_style(source, target)
            else:
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.value = replace_text(cell.value, values)
    workbook.save(output)
