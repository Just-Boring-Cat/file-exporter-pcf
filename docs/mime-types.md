# Common MIME Types

Use the MIME type that matches the file you want to export.

| Category | Extension | MIME type | Example use |
| --- | --- | --- | --- |
| Plain text | `txt` | `text/plain` | notes, logs, simple exports |
| CSV | `csv` | `text/csv` | table exports |
| TSV | `tsv` | `text/tab-separated-values` | spreadsheet-style exports |
| JSON | `json` | `application/json` | config or integration payloads |
| XML | `xml` | `application/xml` | XML integrations |
| HTML | `html` | `text/html` | markup exports |
| Markdown | `md` | `text/markdown` | documentation-like exports |
| JavaScript | `js` | `application/javascript` | script file exports |
| CSS | `css` | `text/css` | stylesheet exports |
| SVG | `svg` | `image/svg+xml` | vector markup |
| YAML | `yaml` | `application/yaml` | YAML config |
| PDF | `pdf` | `application/pdf` | only if the content is already valid PDF data |

## Practical rule

- If the content is text, use a text-based MIME type
- If another system will consume the file, use the MIME type that system expects
- For Power Automate or SharePoint flows, the MIME type is useful metadata but the actual file bytes usually come from the `base64` output
