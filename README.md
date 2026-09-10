## template_filler

**Author:** modsdom
**Version:** 0.1.3
**Type:** tool

### Description

Fill placeholders in Word/Excel templates, or convert Markdown to downloadable DOCX and XLSX files.

### Source Repository

https://github.com/Logan19940824/template-filler-dify-plugin

### Setup

Requires Python 3.12. Install local dependencies with `uv sync`; Dify packaging uses the locked runtime dependencies from `pyproject.toml`.

Template URLs are validated by file content, not by URL suffix or response headers. A Word tool must be able to open the downloaded bytes as DOCX, and an Excel tool must be able to open them as XLSX; preview HTML pages and corrupted files are rejected.

### Usage

The plugin provides nine tools. Call `template_filler_usage_guide` first for the recommended workflow and JSON schemas. The template tools use a public `http` or `https` URL. `inspect_excel_workbook` returns sheet layout metadata and non-empty cell data for an XLSX workbook. `insert_excel_placeholders` validates and inserts placeholders into existing XLSX cells or one styled detail row per worksheet. Placeholders use the `{{name}}` syntax. The parse tools return a `values` object that can be completed and passed to the corresponding fill tool as JSON, for example `{"customer_name":"Alice"}`. Generated files are uploaded to Dify and returned as download links, which also work in Agent nodes that cannot consume inline binary messages.

`markdown_to_word` converts headings, paragraphs, emphasis, lists, quotes, code blocks, and tables into DOCX. `markdown_to_excel` creates one formatted worksheet per Markdown table; Markdown without tables is exported as structured content rows.

The first release supports DOCX body paragraphs and tables, and all XLSX worksheets. Missing values are left as their original placeholders.

Array loops use `{{items[].field}}`; nested arrays are supported, for example `{{modules[].features[].name}}`. Each loop level provides a 1-based `._index`, such as `{{modules[]._index}}` and `{{modules[].features[]._index}}`.

### Local CLI

```bash
uv run template-filler parse-word template.docx
uv run template-filler fill-word template.docx values.json output.docx
uv run template-filler parse-excel template.xlsx
uv run template-filler fill-excel template.xlsx values.json output.xlsx
uv run template-filler markdown-to-word document.md output.docx
uv run template-filler markdown-to-excel tables.md output.xlsx
```
