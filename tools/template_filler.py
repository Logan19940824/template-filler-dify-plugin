from __future__ import annotations

from collections.abc import Generator
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from template_filler.services import (
    fill_excel,
    fill_word,
    inspect_excel_workbook,
    inspect_word_document,
    insert_excel_placeholders,
    insert_word_placeholders,
    markdown_to_excel,
    markdown_to_word,
    parse_excel,
    parse_values,
    parse_word,
)


WORD_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _usage_guide() -> str:
    return """# Template Filler 使用指南

调用任何模板工具前，先确定文件是 DOCX 还是 XLSX，并使用原始文件的公开 HTTP(S) URL。生成文件的工具会返回 Dify 下载链接；后续工具应使用该链接，而不是原始文件链接。

## Excel 模板工作流

1. 调用 `inspect_excel_workbook`，读取原始 XLSX 的工作表、合并区域、冻结窗格和所有非空单元格。
2. 依据业务需求和检查结果，生成 `insert_excel_placeholders` 的 `operations` JSON。
3. 调用 `insert_excel_placeholders`，得到含占位符的 XLSX 模板链接。
4. 调用 `parse_excel_template`，确认占位符、工作表和单元格坐标符合预期。
5. 准备 `values` JSON，调用 `fill_excel_template` 生成最终文件。

## Word 模板工作流

1. 调用 `inspect_word_document`，读取正文、表格、精确文本范围和空白位置。
2. 根据业务要求生成 `replace_text_range` 和可选的 `delete_empty_rows` 操作。
3. 调用 `insert_word_placeholders`，得到含占位符的 DOCX 模板链接。
4. 调用 `parse_word_template` 复核占位符，再调用 `fill_word_template` 填充数据。

检查结果中 `blocks` 保留正文和表格的原文顺序；`paragraphs` 是正文段落，`tables[].rows[].cells[].paragraphs` 是表格段落。空段落、空单元格不会省略。使用返回的 `location` 定位，不能按页码定位，也不能把 Word 单元格当作 Excel 坐标。

`insert_word_placeholders.operations` 的根对象是 `{"document_sha256":"检查结果中的指纹","operations":[...]}`，与 Excel 不同，必须带指纹。`replace_text_range` 示例：

```json
{"document_sha256":"<原文件指纹>","operations":[{"type":"replace_text_range","location":{"table":0,"row":1,"cell":1,"paragraph":0},"expected_text":"","start":0,"end":0,"placeholder":"{{租户[].室号}}"}]}
```

- `expected_text` 必须原样复制目标段落的完整 `text`，包括空格；`start/end` 是原段落中的 Python Unicode 字符下标，从 0 开始、左闭右开。两者相同表示插入，`0,0` 可写入空段落。
- 优先复制 `blank_spans` 或 `runs` 中的范围。替换空白会继承该范围第一个字符的格式；仅在标签后插入会继承标签最后一个字符的格式。要保留下划线，替换带下划线的空白范围。
- 所有操作都基于原文件坐标，不能重叠或重复。文档指纹或原文不一致时应重新检查，不能猜测新坐标。
- 删除样例空行使用 `{"type":"delete_empty_rows","table":0,"rows":[2,3]}`。仅删除明确指定的纯空白、无合并行，不能删除正在写入占位符的行，也不能删除表格全部行。
- 明细表保留一个模板行，字段使用同一个对象数组，例如 `{{租户[]._index}}` 和 `{{租户[].公司名}}`。后续填充按数组长度复制行；空数组删除模板行。
- 查看 `unsupported_features`；`editable:false` 的段落和纵向合并延续单元格不可修改。页眉页脚、文本框、嵌套表格、域、修订、超链接、制表符和换行等复杂段落不在首版编辑范围内。不要声称已覆盖这些区域。

## 工具参数的严格格式

`operations` 和 `values` 在 Dify 中都是字符串参数，但字符串内容必须是合法 JSON 对象。不要传 Markdown、JSON 数组，也不要在参数内容中再嵌套同名字段。

- `insert_excel_placeholders.operations` 的参数原文必须以 `{"operations":` 开头：`{"operations":[{...}]}`。
- `fill_excel_template.values` 的参数原文必须直接以业务字段开头：`{"customer_name":"Alice",...}`，不能写成 `{"values":{...}}`。

当工具调用被表示为完整 JSON 时，字符串参数中的双引号必须转义。下面是可直接参考的完整调用结构：

```json
{
  "template_url": "https://example.com/source.xlsx",
  "operations": "{\\\"operations\\\":[{\\\"type\\\":\\\"replace_cell\\\",\\\"sheet\\\":\\\"Quote\\\",\\\"cell\\\":\\\"B3\\\",\\\"expected_value\\\":\\\"Example Co\\\",\\\"placeholder\\\":\\\"{{customer_name}}\\\"}]}"
}
```

`operations` 的实际字符串内容是：

```json
{"operations":[{"type":"replace_cell","sheet":"Quote","cell":"B3","expected_value":"Example Co","placeholder":"{{customer_name}}"}]}
```

## 如何解读 Excel 原始单元格数据

- `sheets[].cells[]` 仅包含非空单元格；`coordinate` 是坐标，`value` 是原值，`data_type` 为 `s` 表示文本、`n` 表示数字、`d` 表示日期时间、`f` 表示公式、`e` 表示错误。
- 返回数据的最小结构如下；替换时必须同时引用 `name`、`coordinate` 与 `value`，其中 `value` 原样写入 `expected_value`：

```json
{"file_type":"xlsx","sheets":[{"name":"Quote","merged_cells":[],"cells":[{"coordinate":"A3","value":"客户名称","data_type":"s","style_id":2},{"coordinate":"B3","value":"Example Co","data_type":"s","style_id":3},{"coordinate":"C3","value":"=SUM(C5:C10)","data_type":"f","style_id":4,"formula":"=SUM(C5:C10)"}]}]}
```
- 通过相邻标签和取值判断字段含义。例如 `A3` 为“客户名称”、`B3` 为示例客户，则 `B3` 是客户名称字段的候选位置。
- 连续表头和重复的明细值通常表示明细表。标题、标签、说明、合计行和公式行不是填充值候选。
- `merged_cells` 表示合并范围；只允许操作合并范围的左上角单元格。`style_id` 仅用于辅助识别同一视觉区域，不代表固定业务含义。
- 证据不足时不要猜测字段。保留原值并请求用户说明。

## 如何构建占位符规则

单值字段使用英文小写 snake_case，例如 `{{customer_name}}`、`{{contract_no}}`。明细行使用同一数组根路径，例如 `{{items[]._index}}`、`{{items[].name}}`、`{{items[].quantity}}`、`{{items[].unit_price}}`。

使用 `replace_cell` 替换已有示例值。单条操作必须放在外层 `operations` 数组中：

```json
{"operations":[{"type":"replace_cell","sheet":"Quote","cell":"B3","expected_value":"Example Co","placeholder":"{{customer_name}}"}]}
```

`expected_value` 必须等于检查结果中的原值，用于阻止坐标偏移或文件更新后的误写。不得替换公式、错误单元格、非左上角合并单元格、标题、表头、说明或合计单元格。

仅当没有可复用的明细示例行时使用 `insert_row`。它每个工作表最多一次，会在 `before_row` 前插入一行，复制 `copy_style_from_row` 的样式、行高和公式，并在 `cells` 指定列写入占位符：

```json
{"operations":[{"type":"insert_row","sheet":"Quote","before_row":20,"expected_before":{"cell":"A20","expected_value":"Subtotal"},"copy_style_from_row":19,"cells":{"A":"{{items[]._index}}","B":"{{items[].name}}","C":"{{items[].quantity}}"}}]}
```

插入位置和参考行不能涉及合并单元格。`expected_before` 必须指向 `before_row` 上的已知值，用于确认该行确实是预期的合计或锚点行。

## 如何生成占位符数据

先调用 `parse_excel_template` 或 `parse_word_template` 获取完整字段列表及 `values` 骨架。`values` 参数原文是 JSON 对象，而不是 Markdown、JSON 数组或 `{"values":...}`：

```json
{"customer_name":"Alice","contract_no":"CT-2026-001","items":[{"name":"Service A","quantity":2,"unit_price":100}]}
```

完整的 `fill_excel_template` 调用结构如下：

```json
{
  "template_url": "https://example.com/placeholder_template.xlsx",
  "values": "{\\\"customer_name\\\":\\\"Alice\\\",\\\"contract_no\\\":\\\"CT-2026-001\\\",\\\"items\\\":[{\\\"name\\\":\\\"Service A\\\",\\\"quantity\\\":2,\\\"unit_price\\\":100}]}"
}
```

普通字段缺失时会保留原占位符。数组字段必须是对象数组；数组为空时，填充器会删除其模板行。日期、金额和文本应按模板预期的显示格式提供；不要向公式字段传值。

## 如何调用工具

- `inspect_excel_workbook(template_url)`：只读检查 XLSX。
- `insert_excel_placeholders(template_url, operations)`：`operations` 必须是 JSON 字符串，根对象固定为 `{"operations":[...]}`；在原 XLSX 副本中插入占位符，返回模板链接。
- `parse_excel_template(template_url)`：验证 XLSX 模板的占位符。
- `fill_excel_template(template_url, values)`：`values` 必须是 JSON 字符串，根对象直接为业务字段对象；填充 XLSX 并返回下载链接。
- `parse_word_template(template_url)` 与 `fill_word_template(template_url, values)`：解析和填充 DOCX。
- `inspect_word_document(template_url)`：返回 DOCX 正文与表格结构、文档指纹及精确文字范围。
- `insert_word_placeholders(template_url, operations)`：校验指纹和原文后插入占位符，并可删除明确指定的空白表格行。
- `markdown_to_word(markdown_content, output_filename)` 与 `markdown_to_excel(markdown_content, output_filename)`：仅用于将 Markdown 导出为新文件，不能用于保留既有 Excel 样式的模板改造。

完成模板改造后，必须用对应的 `parse_excel_template` 或 `parse_word_template` 复核生成文件。"""


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


def _output_filename(parameters: dict[str, Any], default: str, suffix: str) -> str:
    requested = str(parameters.get("output_filename") or default).strip()
    filename = Path(requested).name
    if not filename or filename in {".", ".."}:
        filename = default
    if not filename.lower().endswith(suffix):
        filename += suffix
    return filename


def _upload_result(
    tool: Tool,
    filename: str,
    blob: bytes,
    mime_type: str,
) -> ToolInvokeMessage:
    uploaded = tool.session.file.upload(filename, blob, mime_type)
    download_url = str(uploaded.preview_url or "").strip()
    if not download_url:
        raise RuntimeError("Dify uploaded the file but did not return a download URL")

    return tool.create_link_message(download_url)


class ParseWordTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".docx", parse_word))


class FillWordTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = _fill(tool_parameters["template_url"], ".docx", fill_word, parse_values(tool_parameters["values"]))
        yield _upload_result(self, "filled.docx", blob, WORD_MIME_TYPE)


class InspectWordDocumentTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".docx", inspect_word_document))


class InsertWordPlaceholdersTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        operations = parse_values(tool_parameters["operations"])
        blob = _fill(tool_parameters["template_url"], ".docx", insert_word_placeholders, operations)
        yield _upload_result(self, "placeholder_template.docx", blob, WORD_MIME_TYPE)


class ParseExcelTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".xlsx", parse_excel))


class InspectExcelWorkbookTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_json_message(_parse(tool_parameters["template_url"], ".xlsx", inspect_excel_workbook))


class InsertExcelPlaceholdersTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        operations = parse_values(tool_parameters["operations"])
        if not isinstance(operations.get("operations"), list):
            raise ValueError("operations must be a JSON object containing an operations array")
        blob = _fill(
            tool_parameters["template_url"],
            ".xlsx",
            insert_excel_placeholders,
            operations["operations"],
        )
        yield _upload_result(self, "placeholder_template.xlsx", blob, EXCEL_MIME_TYPE)


class TemplateFillerUsageGuideTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        yield self.create_text_message(_usage_guide())


class FillExcelTemplateTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = _fill(tool_parameters["template_url"], ".xlsx", fill_excel, parse_values(tool_parameters["values"]))
        yield _upload_result(self, "filled.xlsx", blob, EXCEL_MIME_TYPE)


class MarkdownToWordTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = markdown_to_word(tool_parameters["markdown_content"])
        yield _upload_result(
            self,
            _output_filename(tool_parameters, "markdown.docx", ".docx"),
            blob,
            WORD_MIME_TYPE,
        )


class MarkdownToExcelTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        blob = markdown_to_excel(tool_parameters["markdown_content"])
        yield _upload_result(
            self,
            _output_filename(tool_parameters, "markdown.xlsx", ".xlsx"),
            blob,
            EXCEL_MIME_TYPE,
        )
