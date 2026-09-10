from datetime import date
from pathlib import Path

from openpyxl import Workbook

from template_filler.services.excel import inspect_excel_workbook


def test_inspect_excel_workbook_returns_layout_and_non_empty_cells(tmp_path: Path):
    path = tmp_path / "workbook.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Quote"
    sheet["A1"] = "Customer"
    sheet["B1"] = "Example Co"
    sheet["A2"] = date(2026, 9, 10)
    sheet["B2"] = "=SUM(1,2)"
    sheet.merge_cells("C1:D1")
    sheet.freeze_panes = "A2"
    workbook.save(path)

    result = inspect_excel_workbook(path)

    assert result["file_type"] == "xlsx"
    assert result["sheets"][0]["name"] == "Quote"
    assert result["sheets"][0]["freeze_panes"] == "A2"
    assert result["sheets"][0]["merged_cells"] == ["C1:D1"]
    assert result["sheets"][0]["cells"] == [
        {"coordinate": "A1", "value": "Customer", "data_type": "s", "style_id": 0},
        {"coordinate": "B1", "value": "Example Co", "data_type": "s", "style_id": 0},
        {"coordinate": "A2", "value": "2026-09-10T00:00:00", "data_type": "d", "style_id": 1},
        {"coordinate": "B2", "value": "=SUM(1,2)", "data_type": "f", "style_id": 0, "formula": "=SUM(1,2)"},
    ]
