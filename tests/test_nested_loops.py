from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

from template_filler.services import fill_excel, fill_word
from template_filler.services.common import array_fields, expand_contexts, render_context


VALUES = {
    "modules": [
        {"name": "用户端", "features": [{"name": "登录"}, {"name": "注册"}]},
        {"name": "管理端", "features": [{"name": "审核"}]},
    ]
}


def test_context_and_indexes():
    text = "{{modules[]._index}}.{{modules[].name}}-{{modules[].features[]._index}}.{{modules[].features[].name}}"
    contexts = expand_contexts(VALUES, array_fields(text))
    assert [render_context(text, VALUES, context) for context in contexts] == [
        "1.用户端-1.登录",
        "1.用户端-2.注册",
        "2.管理端-1.审核",
    ]


def test_excel_nested_loop(tmp_path: Path):
    source, output = tmp_path / "template.xlsx", tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["模块序号", "模块", "功能序号", "功能点"])
    sheet.append(["{{modules[]._index}}", "{{modules[].name}}", "{{modules[].features[]._index}}", "{{modules[].features[].name}}"])
    workbook.save(source)
    fill_excel(source, output, VALUES)
    rows = list(load_workbook(output).active.values)
    assert rows[1:] == [("1", "用户端", "1", "登录"), ("1", "用户端", "2", "注册"), ("2", "管理端", "1", "审核")]


def test_word_nested_loop(tmp_path: Path):
    source, output = tmp_path / "template.docx", tmp_path / "output.docx"
    document = Document()
    document.add_paragraph("{{modules[]._index}}.{{modules[].name}}-{{modules[].features[]._index}}.{{modules[].features[].name}}")
    document.save(source)
    fill_word(source, output, VALUES)
    assert [paragraph.text for paragraph in Document(output).paragraphs] == [
        "1.用户端-1.登录",
        "1.用户端-2.注册",
        "2.管理端-1.审核",
    ]
