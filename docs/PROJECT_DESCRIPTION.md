# Template Filler 项目描述

## 1. 项目概览

`template_filler` 是一个运行在 Dify 插件运行时中的 Python 工具插件。它面向需要由工作流或 Agent 自动生成交付文档的场景，提供两类能力：

1. 以 JSON 数据填充 Word（DOCX）或 Excel（XLSX）模板中的占位符；
2. 将 Markdown 内容转换为可下载的 DOCX 或 XLSX 文件。

插件不需要配置第三方账号或凭据。模板由调用方提供公开的 HTTP(S) 地址，生成的文件上传至 Dify 文件服务，再以下载链接形式返回。因此，Dify Agent 节点即使不能直接消费二进制文件，也可以继续向用户提供生成结果。

当前版本为 `0.1.4`，要求 Python 3.12，插件类型为 Dify `tool`。

## 2. 目标场景

- 根据客户、订单、报价、项目清单等结构化数据批量生成 Word 报告或 Excel 表格；
- 先分析客户提供的模板，取得待补全字段，再由工作流或大模型组织对应 JSON；
- 将大模型生成的 Markdown 报告、会议纪要、清单或数据表导出为 Office 文档；
- 在本地命令行调试模板和输出文件，而不依赖 Dify 环境。

## 3. 对外功能

插件注册了 11 个 Dify 工具，均由 `provider/template_filler.yaml` 声明，并由 `tools/template_filler.py` 实现。

| 工具 | 输入 | 输出 | 用途 |
| --- | --- | --- | --- |
| `template_filler_usage_guide` | 无 | 推荐流程、输入结构、操作 JSON 与安全规则 | 指导智能体使用其他工具 |
| `parse_word_template` | DOCX 模板的公开 URL | 占位符、出现位置、默认 `values` 对象 | 识别 Word 模板需要的数据字段 |
| `fill_word_template` | DOCX 模板 URL、占位符值 JSON | 生成文件的 Dify 下载链接 | 填充 Word 模板 |
| `inspect_word_document` | DOCX 原文件的公开 URL | 文档指纹、正文和表格结构、文字和空白范围 | 为模型规划占位符提供准确位置 |
| `insert_word_placeholders` | DOCX 原文件 URL、操作 JSON | 占位符模板的 DOCX 下载链接 | 校验指纹和原文后写入占位符、删除指定空白行 |
| `parse_excel_template` | XLSX 模板的公开 URL | 占位符、工作表/单元格位置、默认 `values` 对象 | 识别 Excel 模板需要的数据字段 |
| `inspect_excel_workbook` | XLSX 工作簿的公开 URL | 工作表布局与非空单元格数据 | 为智能体理解既有工作簿并规划占位符位置提供结构化上下文 |
| `insert_excel_placeholders` | XLSX 工作簿 URL、固定格式操作 JSON | 插入占位符后的 XLSX 下载链接 | 校验并写入指定现有单元格的占位符 |
| `fill_excel_template` | XLSX 模板 URL、占位符值 JSON | 生成文件的 Dify 下载链接 | 填充 Excel 模板 |
| `markdown_to_word` | Markdown 文本、可选输出名 | DOCX 下载链接 | 将 Markdown 转为 Word 文档 |
| `markdown_to_excel` | Markdown 文本、可选输出名 | XLSX 下载链接 | 将 Markdown 表格或正文转为 Excel 工作簿 |

模板填充工具的结果文件名固定为 `filled.docx` 或 `filled.xlsx`。Markdown 转换工具可通过 `output_filename` 指定文件名；未写扩展名时，插件会补上对应的 `.docx` 或 `.xlsx` 扩展名，并会移除路径部分以避免由参数决定输出路径。

`inspect_excel_workbook` 只读取原工作簿，不写入或上传文件。它为每个工作表返回名称、可见状态、已用范围、冻结窗格、合并区域及非空单元格。每个单元格包含坐标、JSON 可序列化的值、Excel 数据类型、样式 ID；公式单元格额外返回公式文本。日期和时间使用 ISO 8601 字符串表示。

`insert_excel_placeholders` 支持 `replace_cell` 和 `insert_row`。后者每个工作表最多执行一次，要求插入位置的原始锚点和值以及参考行；它在锚点行前插入新行，复制参考行样式、行高和公式，并只在 `cells` 指定列写入占位符。合并行、公式单元格和错误单元格会被拒绝。

## 4. 模板占位符与填充规则

### 4.1 基础字段

占位符采用双花括号语法：`{{field_name}}`。调用方通过 JSON 对象提供值，例如：

```json
{
  "customer_name": "张三",
  "order_id": "ORD-2026-001"
}
```

对应模板中的 `{{customer_name}}` 和 `{{order_id}}` 会被替换。值会转换为文本：`null` 变为空字符串，布尔值变为 `true` 或 `false`，其他值使用其字符串表示。未提供的普通字段会保留原始占位符，便于识别尚未完成的数据。

解析工具会返回：

- 文档类型：`docx` 或 `xlsx`；
- 去重后的占位符名称；
- 每个占位符的出现位置；
- 可直接作为填充工具输入起点的 `values` 对象。普通字段的默认值为空字符串，数组字段的默认值为空数组。

### 4.2 数组循环

数组字段以 `[]` 表示，例如 `{{items[].name}}`。同一段落、Word 表格行或 Excel 行内只要出现数组字段，该元素就会被作为重复模板：

```json
{
  "items": [
    {"name": "产品 A", "price": 100},
    {"name": "产品 B", "price": 80}
  ]
}
```

```text
{{items[]._index}}. {{items[].name}}，价格：{{items[].price}}
```

上例生成两条内容，`_index` 为从 1 开始的数组序号。空数组会删除包含该数组占位符的模板段落或模板行。

支持嵌套数组，例如 `{{modules[].features[].name}}`。实现会从外到内展开数组上下文；内层数组的每一项形成一条输出记录，外层字段会随之重复。各层都能使用自身的 `._index`，例如：

```text
{{modules[]._index}}. {{modules[].name}}
{{modules[].features[]._index}}. {{modules[].features[].name}}
```

数组值必须是“对象数组”。例如 `items` 必须是 `[{...}, {...}]`，而不是字符串数组或数字数组；不符合时会报错，避免生成不确定的结果。

## 5. Word 模板处理

### 5.1 解析

Word 解析直接读取 DOCX 压缩包中的 `word/document.xml`，扫描正文段落和正文表格单元格中的文本节点。输出中：

- 普通正文段落以 `paragraph` 表示从 0 开始的段落序号；
- 表格内容以 `table`、`row`、`cell`、`paragraph` 表示从 0 开始的位置。

这使上游工作流能够知道字段存在于正文还是表格中。

### 5.2 填充

填充时，普通字段会替换正文段落和表格单元格段落中的文本。对于 Word 常见的“一个占位符被拆成多个文本 run”的情况，插件会先合并段落文本后替换，因此 `{{order_` 与 `id}}` 分置于两个 run 时仍可识别。

数组字段会复制包含字段的正文段落或整张表格行。复制过程基于原始 XML 节点，保留已有样式。普通字段和循环字段均按占位符范围跨文字片段替换，新值继承占位符首字符格式，前后文本的格式保持不变。

当前处理范围限于 `word/document.xml` 中的正文段落和正文表格。页眉、页脚、批注、文本框、脚注、尾注及其他 DOCX 部件不在扫描和替换范围内。

`inspect_word_document` 会进一步返回正文和表格中的完整文本、文字片段范围、连续空白范围、空行及文档 SHA-256 指纹。上游模型据此生成精确的字符范围操作；`insert_word_placeholders` 在核对指纹和原文后写入占位符，并可删除明确指定的纯空白表格行。所有字符下标均从 0 开始，范围采用左闭右开形式。

检查结果通过 `blocks` 表达原文顺序，`paragraphs` 存储正文段落，`tables[].rows[].cells[].paragraphs` 存储表格段落。每段包含 `location`、原样 `text`、`runs`、`blank_spans` 和 `editable`。表格还返回网格宽度、物理单元格索引、网格列号与合并信息。`unsupported_features` 标记未完整处理的复杂结构，`editable:false` 的目标会拒绝编辑。具体接口与台账示例见 [Word 模板工作流](WORD_TEMPLATE_WORKFLOW.md)。

## 6. Excel 模板处理

### 6.1 解析

Excel 解析使用 `openpyxl` 遍历所有工作表的单元格。每个占位符的出现位置包含工作表名称及单元格坐标，例如：

```json
{
  "name": "customer_name",
  "occurrences": [{"sheet": "Orders", "cells": ["B2"]}]
}
```

### 6.2 填充

普通字段会在所有工作表中替换。对于数组字段，插件以整行为单位扩展：

- 每个数组上下文生成一行；
- 从末行向首行处理，避免插入行时影响尚未处理的模板行；
- 复制行的单元格值和主要单元格样式，包括字体、填充、边框、对齐、数字格式与保护属性。

因此，推荐把一条明细记录所需的数组占位符放在同一模板行中，例如产品名称、数量、单价各放在不同列。空数组会删除该模板行。

## 7. Markdown 导出

### 7.1 Markdown 转 Word

解析器使用 CommonMark 规则并启用表格扩展，禁止直接解释 HTML。DOCX 导出支持：

- 标题；
- 段落；
- 粗体、斜体、删除线、行内代码和链接样式；
- 有序与无序列表，最多使用 Word 的三级列表样式；
- 引用；
- 围栏代码块或缩进代码块；
- 分隔线；
- Markdown 表格；
- 图片的替代文本与链接地址（以文本形式写入，不下载或嵌入图片）。

表格会生成 Word 的 `Table Grid` 表格，首行加粗。HTML 块不会作为 HTML 渲染，而是作为文本段落写入。

### 7.2 Markdown 转 Excel

若 Markdown 中有表格，每个表格会生成一个工作表：

- 表格前最近的标题作为工作表名称；无标题时使用 `Table N`；
- 非法工作表字符会替换为下划线，名称截断至 31 个字符；重名会追加序号；
- 首行使用深蓝色表头、白色粗体字，冻结首行并启用筛选；
- 内容单元格自动换行，列宽按内容估算并限制在 10 至 60 个字符宽度。

若 Markdown 没有表格，插件会创建名为 `Content` 的工作表，以 `Type` 与 `Content` 两列输出标题、段落、列表项、引用和代码块等文档结构。

两种 Markdown 转换均拒绝空白内容。

## 8. Dify 调用与文件处理流程

模板型工具遵循以下流程：

1. 校验 `template_url` 为带主机名的 `http` 或 `https` URL；
2. 以 `dify-template-filler/1.0` User-Agent 下载模板，下载超时为 30 秒；
3. 最多读取 25 MB；超过限制即报错；
4. 保存至临时文件，并尝试作为目标类型的 DOCX 或 XLSX 打开；
5. 解析占位符或生成填充文件；
6. 将生成文件上传至 Dify 文件 API；
7. 从 Dify 返回的 `preview_url` 创建链接消息；若上传结果没有下载 URL，则明确报错；
8. 删除本次调用创建的临时输入和输出文件。

文件类型以内容是否能被对应的文档库正确打开为准，而非 URL 后缀或 HTTP 响应头。因此，预览 HTML 页面、损坏文件或格式不匹配的文件会被拒绝。

Markdown 转换不访问外部 URL：它只处理传入的文本，在内存中生成 Office 文件，再上传到 Dify。

## 9. 本地命令行

项目提供 `template-filler` 命令，与 Dify 插件复用 `template_filler.services` 服务层：

```powershell
uv sync

uv run template-filler parse-word template.docx
uv run template-filler fill-word template.docx values.json output.docx

uv run template-filler parse-excel template.xlsx
uv run template-filler fill-excel template.xlsx values.json output.xlsx

uv run template-filler markdown-to-word document.md output.docx
uv run template-filler markdown-to-excel tables.md output.xlsx
```

`parse-word` 与 `parse-excel` 将 JSON 输出到标准输出。填充命令的 `values.json` 必须包含 JSON 对象。Markdown 命令读取 UTF-8 编码文件并直接写出 Office 文件。

仓库中的 `fixtures/` 提供示例 DOCX、XLSX、值 JSON 以及生成模板的脚本，便于手工验证基础字段、跨 run 占位符、循环字段与多工作表场景。

## 10. 技术结构与依赖

```text
Dify plugin runtime
  main.py
    provider/template_filler.yaml
      tools/*.yaml
        tools/template_filler.py
          template_filler/services/
            common.py   占位符解析、JSON 校验、数组上下文展开
            word.py     DOCX 解析与填充
            excel.py    XLSX 解析与填充
            markdown.py Markdown 到 DOCX/XLSX 的转换
          template_filler/cli.py
            本地命令行入口
```

主要依赖：

| 依赖 | 职责 |
| --- | --- |
| `dify-plugin` 0.9.1 | Dify 插件、工具消息和文件上传接口 |
| `python-docx` 1.2.0 | 创建 Markdown 导出的 DOCX |
| `lxml` 5.3.0 | DOCX 相关依赖声明 |
| `openpyxl` 3.1.5 | 读取、填充和生成 XLSX |
| `markdown-it-py` 3.0.0 | 解析 Markdown |

Word 模板填充没有使用 `python-docx` 重写文件，而是保留 DOCX 压缩包的其他条目，仅替换 `word/document.xml`。这样可避免仅为填充正文而重建完整的 Word 文件。

## 11. 数据、安全与隐私边界

- 插件不维护自己的数据库或持久化存储；
- 模板型调用只向调用方提供的公开 HTTP(S) 地址发起请求；
- 模板内容、填充值、Markdown 内容及生成文件仅用于本次解析或生成；
- 结果文件由 Dify 的文件服务管理，其保留、访问控制和删除规则取决于部署的 Dify 实例；
- 插件不请求 API Key、密码、账号凭据、分析标识或联系信息；
- 调用方需要自行确保有权访问模板 URL 并处理其中的数据。

需要注意：模板 URL 可以指向任意满足校验条件的 HTTP(S) 地址。部署方应结合自身网络策略和 Dify 运行环境，决定是否需要额外限制可访问的主机或地址范围。

## 12. 已覆盖验证

项目的 pytest 用例覆盖了核心可观察行为：

- 嵌套数组上下文与各层从 1 开始的索引；
- Word 与 Excel 的嵌套循环填充；
- Word 占位符跨多个 run 时的替换；
- 非 DOCX/XLSX 内容的拒绝；
- Markdown 到 DOCX 的常用块级元素和表格；
- Markdown 表格到多工作表 XLSX、无表格时的结构化内容导出；
- 空 Markdown 的拒绝；
- Dify 文件上传、下载链接返回及缺少下载 URL 时的错误处理。

可在项目根目录执行：

```powershell
uv run pytest
```

## 13. 当前边界总结

该项目的核心定位是“基于公开模板 URL 的正文级 Word/Excel 占位符填充，以及 Markdown 的轻量 Office 导出”。它不是通用 Office 编辑器：Word 模板处理不覆盖正文以外的部件；数组循环按段落、表格行或 Excel 行扩展；Markdown 图片不下载嵌入；模板仅支持 DOCX 与 XLSX，且模板下载大小上限为 25 MB。
