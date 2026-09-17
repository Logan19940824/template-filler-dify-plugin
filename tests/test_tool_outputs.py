from types import SimpleNamespace

import pytest
from dify_plugin.entities.invoke_message import InvokeMessage
from dify_plugin.entities.tool import ToolInvokeMessage

import tools.template_filler as template_tools
from tools.template_filler import (
    FillExcelTemplateTool,
    FillWordTemplateTool,
    InsertWordPlaceholdersTool,
    MarkdownToExcelTool,
    MarkdownToWordTool,
    TemplateFillerUsageGuideTool,
)


class UploadStub:
    def __init__(self, *, preview_url: str | None = "https://dify.example/files/1") -> None:
        self.preview_url = preview_url
        self.calls: list[tuple[str, bytes, str]] = []

    def upload(self, filename: str, content: bytes, mime_type: str) -> SimpleNamespace:
        self.calls.append((filename, content, mime_type))
        return SimpleNamespace(
            id="file-1",
            name=filename,
            size=len(content),
            extension=filename.rsplit(".", 1)[-1],
            mime_type=mime_type,
            preview_url=self.preview_url,
        )


def _tool(tool_type: type, upload: UploadStub):
    tool = object.__new__(tool_type)
    tool.response_type = ToolInvokeMessage
    tool.session = SimpleNamespace(file=upload)
    return tool


def _assert_uploaded_result(messages: list[ToolInvokeMessage], upload: UploadStub, filename: str, mime_type: str) -> None:
    assert len(messages) == 1
    assert messages[0].type == InvokeMessage.MessageType.LINK
    assert messages[0].message.text == "https://dify.example/files/1"
    assert upload.calls[0][0] == filename
    assert upload.calls[0][1] == b"doc"
    assert upload.calls[0][2] == mime_type


def test_markdown_to_word_uploads_result(monkeypatch: pytest.MonkeyPatch) -> None:
    upload = UploadStub()
    monkeypatch.setattr(template_tools, "markdown_to_word", lambda content: b"doc")

    messages = list(
        _tool(MarkdownToWordTool, upload)._invoke(
            {"markdown_content": "# Title", "output_filename": "report"}
        )
    )

    _assert_uploaded_result(messages, upload, "report.docx", template_tools.WORD_MIME_TYPE)


def test_markdown_to_excel_uploads_result(monkeypatch: pytest.MonkeyPatch) -> None:
    upload = UploadStub()
    monkeypatch.setattr(template_tools, "markdown_to_excel", lambda content: b"doc")

    messages = list(
        _tool(MarkdownToExcelTool, upload)._invoke(
            {"markdown_content": "| A |\n| - |\n| 1 |"}
        )
    )

    _assert_uploaded_result(messages, upload, "markdown.xlsx", template_tools.EXCEL_MIME_TYPE)


@pytest.mark.parametrize(
    ("tool_type", "suffix", "filename", "mime_type"),
    [
        (FillWordTemplateTool, ".docx", "filled.docx", template_tools.WORD_MIME_TYPE),
        (FillExcelTemplateTool, ".xlsx", "filled.xlsx", template_tools.EXCEL_MIME_TYPE),
    ],
)
def test_template_fill_tools_upload_result(
    monkeypatch: pytest.MonkeyPatch,
    tool_type: type,
    suffix: str,
    filename: str,
    mime_type: str,
) -> None:
    upload = UploadStub()
    monkeypatch.setattr(template_tools, "_fill", lambda url, file_suffix, handler, values: b"doc")
    monkeypatch.setattr(template_tools, "parse_values", lambda values: {})

    messages = list(
        _tool(tool_type, upload)._invoke(
            {"template_url": "https://example.test/template" + suffix, "values": "{}"}
        )
    )

    _assert_uploaded_result(messages, upload, filename, mime_type)


def test_insert_word_placeholders_uploads_result(monkeypatch: pytest.MonkeyPatch) -> None:
    upload = UploadStub()
    monkeypatch.setattr(template_tools, "_fill", lambda url, suffix, handler, values: b"doc")
    monkeypatch.setattr(template_tools, "parse_values", lambda values: {"document_sha256": "hash", "operations": [{}]})

    messages = list(_tool(InsertWordPlaceholdersTool, upload)._invoke({
        "template_url": "https://example.test/template.docx",
        "operations": "{}",
    }))

    _assert_uploaded_result(messages, upload, "placeholder_template.docx", template_tools.WORD_MIME_TYPE)


def test_upload_requires_download_url(monkeypatch: pytest.MonkeyPatch) -> None:
    upload = UploadStub(preview_url=None)
    monkeypatch.setattr(template_tools, "markdown_to_word", lambda content: b"doc")

    with pytest.raises(RuntimeError, match="did not return a download URL"):
        list(
            _tool(MarkdownToWordTool, upload)._invoke(
                {"markdown_content": "# Title"}
            )
        )


def test_usage_guide_returns_markdown_text():
    tool = object.__new__(TemplateFillerUsageGuideTool)
    tool.response_type = ToolInvokeMessage
    message = list(tool._invoke({}))[0]

    assert message.type == InvokeMessage.MessageType.TEXT
    assert message.message.text.startswith("# Template Filler 使用指南")
    for heading in ("如何解读 Excel 原始单元格数据", "如何构建占位符规则", "如何生成占位符数据", "如何调用工具"):
        assert heading in message.message.text
    assert '"operations": "{\\"operations\\":[' in message.message.text
    assert '{"operations":[{"type":"replace_cell"' in message.message.text
    assert '`{"values":{...}}`' in message.message.text
