## template_filler

**Author:** modsdom
**Version:** 0.0.1
**Type:** tool

### Description

Pass a Word/Excel template URL and a placeholder-value map; the plugin downloads the template, fills placeholders, and returns a file that keeps the original layout and styling.

### Setup

Requires Python 3.12. Install local dependencies with `uv sync --no-group dify`; Dify packaging uses the dependencies listed in `requirements.txt`.

Template URLs are validated by file content, not by URL suffix or response headers. A Word tool must be able to open the downloaded bytes as DOCX, and an Excel tool must be able to open them as XLSX; preview HTML pages and corrupted files are rejected.

### Usage

Use one of the four tools with a public `http` or `https` URL. Placeholders use the `{{name}}` syntax. The parse tools return a `values` object that can be completed and passed to the corresponding fill tool as JSON, for example `{"customer_name":"Alice"}`. Generated files are returned as downloadable DOCX or XLSX file messages.

The first release supports DOCX body paragraphs and tables, and all XLSX worksheets. Missing values are left as their original placeholders.

Array loops use `{{items[].field}}`; nested arrays are supported, for example `{{modules[].features[].name}}`. Each loop level provides a 1-based `._index`, such as `{{modules[]._index}}` and `{{modules[].features[]._index}}`.

### Local CLI

```bash
uv run template-filler parse-word template.docx
uv run template-filler fill-word template.docx values.json output.docx
uv run template-filler parse-excel template.xlsx
uv run template-filler fill-excel template.xlsx values.json output.xlsx
```
