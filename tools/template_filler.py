from __future__ import annotations

from collections.abc import Generator
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from template_filler.services import fill_excel, fill_word, parse_excel, parse_values, parse_word


def _download(url: str, suffix: str) -> Path:
    parsed = urlparse(str(url).strip())
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("template_url must be a valid HTTP(S) URL")
    with urlopen(Request(url, headers={"User-Agent": "dify-template-filler/1.0"}), timeout=30) as response:
        data = response.read(25 * 1024 * 1024 + 1)
    if len(data) > 25 * 1024 * 1024:
        raise ValueError("template file exceeds the 25 MB limit")
    temp = NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        temp.write(data)
        return Path(temp.name)
    finally:
        temp.close()


def _parse(url: str, suffix: str, handler) -> dict[str, Any]:
    source = _download(url, suffix)
    try:
        return handler(source)
    finally:
        source.unlink(missing_ok=True)


def _fill(url: str, suffix: str, handler, values: dict[str, Any]) -> bytes:
    source = _download(url, suffix)
    output = Path(source).with_name("filled_" + source.name)
    try:
        handler(source, output, values)
        return output.read_bytes()
    finally:
        source.unlink(missing_ok=True)
        output.unlink(missing_ok=True)


class ParseWordTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".docx", parse_word))


class FillWordTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = _fill(tool_parameters["template_url"], ".docx", fill_word, parse_values(tool_parameters["values"]))
        yield self.create_blob_message(blob=blob, meta={"mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "filename": "filled.docx"})


class ParseExcelTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".xlsx", parse_excel))


class FillExcelTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = _fill(tool_parameters["template_url"], ".xlsx", fill_excel, parse_values(tool_parameters["values"]))
        yield self.create_blob_message(blob=blob, meta={"mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "filename": "filled.xlsx"})
