# Privacy Policy for Template Filler

Effective date: August 28, 2026

Template Filler runs inside the Dify plugin runtime. It fills DOCX and XLSX templates and converts Markdown into DOCX or XLSX files.

## Data Processed

The plugin processes only the data supplied to its tools:

- Public HTTP(S) template URLs. The plugin downloads the file at the URL to read or fill a DOCX or XLSX template.
- Template contents, placeholder values, and Markdown content supplied in a Dify workflow or chat.
- Generated DOCX and XLSX files.

The plugin does not request, collect, or store API keys, passwords, account credentials, analytics identifiers, or contact information.

## How Data Is Used

Data is used only to parse templates, replace placeholders, or create the requested document. For template-based tools, the plugin makes an outbound request only to the public URL supplied by the user. The resulting file is uploaded through Dify's file API so that Dify can return a download link.

The plugin does not send data to an operator-controlled service, advertising service, analytics service, or any other third party. Users are responsible for ensuring that they are authorized to provide the template URL and its contents.

## Retention and Deletion

The plugin does not maintain its own database or persistent storage. Input and generated files are handled by the Dify deployment and any server hosting a user-supplied template URL. Their retention, access controls, and deletion procedures are governed by the applicable Dify deployment and hosting provider policies.

## Security

Template URLs must use HTTP or HTTPS. The plugin validates downloaded files by attempting to open them as the expected DOCX or XLSX format and rejects invalid files.

## Contact

For privacy questions or deletion requests related to this plugin, open an issue in the source repository: https://github.com/Logan19940824/template-filler-dify-plugin/issues
