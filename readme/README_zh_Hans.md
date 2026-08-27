# 模板填充器

插件提供四个工具：`parse_word_template`、`fill_word_template`、`parse_excel_template`、`fill_excel_template`。

模板文件必须是公开可访问的 `HTTP(S)` 地址，支持 `.docx` 和 `.xlsx`。占位符格式为 `{{name}}`。

解析工具返回占位符列表、位置元数据和可直接填写的 `values` 对象。生成工具接收同一模板地址和 JSON 字符串形式的键值映射，例如：

```json
{"customer_name":"张三","amount":123}
```

循环占位符使用 `{{items[].field}}`，支持嵌套数组，例如 `{{modules[].features[].name}}`。每层循环都提供从 1 开始的 `._index`，例如 `{{modules[]._index}}` 和 `{{modules[].features[]._index}}`。Excel 会按最内层数组复制所在行，Word 会复制所在段落或表格行：

```json
{"items":[{"name":"商品 A"},{"name":"商品 B"}]}
```

未提供的普通值会保留原占位符；生成结果以可下载文件消息返回。

模板 URL 不依赖后缀名或响应头判断类型。Word 工具会尝试按 DOCX 打开下载内容，Excel 工具会尝试按 XLSX 打开；预览 HTML 页面和损坏文件会被拒绝。

## 本地运行

```powershell
uv sync --no-group dify
uv run template-filler parse-word template.docx
uv run template-filler fill-word template.docx values.json output.docx
uv run template-filler parse-excel template.xlsx
uv run template-filler fill-excel template.xlsx values.json output.xlsx
```
