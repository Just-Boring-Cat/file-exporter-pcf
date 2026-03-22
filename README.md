# File Exporter Component PCF

<p align="center">
  <img alt="GitHub release" src="https://img.shields.io/github/v/release/Just-Boring-Cat/file-exporter-pcf?label=release">
  <img alt="License" src="https://img.shields.io/github/license/Just-Boring-Cat/file-exporter-pcf">
  <img alt="Last commit" src="https://img.shields.io/github/last-commit/Just-Boring-Cat/file-exporter-pcf">
  <img alt="Repo size" src="https://img.shields.io/github/repo-size/Just-Boring-Cat/file-exporter-pcf">
</p>

Power Apps PCF control for Canvas apps that exports one or more documents from JSON input and either downloads them in the browser or exposes them as JSON output for Power Automate, SharePoint, and other external systems.

## What This Component Does

1. You pass a JSON array of documents into `documentsJson`.
2. The control validates each document, computes metadata and Base64 output, and exposes the result through `documentsOutputJson`.
3. Depending on `downloadMode`, it either:
   - does nothing and only updates outputs
   - downloads one ZIP file
   - attempts separate browser downloads

## Example `documentsJson`

Copy and paste this directly into Power Fx after wrapping it with `JSON(...)`, or use it as raw JSON when testing outside Canvas:

```json
[
  {
    "id": "doc-001",
    "fileName": "customers",
    "fileExtension": "csv",
    "mimeType": "text/csv",
    "contentText": "Id,Name,Country\n1,Ana,DE\n2,John,US\n3,Mika,JP"
  },
  {
    "id": "doc-002",
    "fileName": "notes",
    "fileExtension": "txt",
    "mimeType": "text/plain",
    "contentText": "This is a simple text file.\nSecond line.\nThird line."
  },
  {
    "id": "doc-003",
    "fileName": "config",
    "fileExtension": "json",
    "mimeType": "application/json",
    "contentText": "{\n  \"environment\": \"test\",\n  \"featureEnabled\": true,\n  \"version\": 1\n}"
  }
]
```

## JSON Schema Summary

Each document object supports:

- `id` optional string
- `fileName` required string
- `fileExtension` required string
- `mimeType` required string
- `contentText` required string

Full schema details are in [docs/json-schema.md](docs/json-schema.md).

## Common MIME Types

| File type | Extension | MIME type | Typical use |
| --- | --- | --- | --- |
| Text | `txt` | `text/plain` | plain text notes and logs |
| CSV | `csv` | `text/csv` | tabular exports |
| JSON | `json` | `application/json` | config and machine-readable data |
| XML | `xml` | `application/xml` | structured integrations |
| HTML | `html` | `text/html` | generated markup exports |
| Markdown | `md` | `text/markdown` | documentation-style exports |
| TSV | `tsv` | `text/tab-separated-values` | spreadsheet-style exports |
| PDF | `pdf` | `application/pdf` | only if your content is already valid PDF data |

More examples are listed in [docs/mime-types.md](docs/mime-types.md).

## Visual Preview

Component in Canvas:

![File Exporter Component preview](docs/images/file-exporter-component-1.jpeg)

Documents JSON property:

![Documents JSON property](docs/images/file-exporter-component-2-documentsJson.jpeg)

Download mode property:

![Download mode property](docs/images/file-exporter-component-3-download-modes.jpeg)

Documents output JSON:

![Documents output JSON property](docs/images/file-exporter-component-4-documentsOutputJson.jpeg)

Output properties:

![Output properties](docs/images/file-exporter-component-5-output-properties.jpeg)

## Key Features

- Multi-document export from one JSON input
- `downloadMode = Disabled | Zip | Separate`
- JSON outputs for Power Fx and Power Automate
- Base64 per document in `documentsOutputJson`
- `OnSelect`, `selectToken`, and `changeToken` for app-side formula patterns
- Button styling inputs for colors, fonts, borders, padding, and alignment

## Important Notes

- `Disabled` is the recommended mode when Power Automate or another system should handle storage.
- `Zip` is the most reliable built-in browser download mode.
- `Separate` is best-effort because browsers and hosts may limit multi-download behavior.
- `archiveFileName` is only used for ZIP mode.
- The control does not upload anything by itself. External storage happens through the host app or Flow.

## Quick Start

### Local PCF harness

```bash
cd pcf/FileExporterComponent
npm install
npm run refreshTypes
npm run build
npm run start
```

### Build the Dataverse solution

```bash
dotnet build solution/File_Exporter_Component_Solution/File_Exporter_Component_Solution.cdsproj -c Release
```

Managed-only:

```bash
dotnet build solution/File_Exporter_Component_Solution/File_Exporter_Component_Solution.cdsproj -c Release /p:SolutionPackageType=Managed
```

## How To Use It In Canvas

1. Import the solution into your environment.
2. Add the control to your Canvas app.
3. Bind `documentsJson` to a JSON array.
4. Choose `downloadMode`:
   - `Disabled` for output-only use
   - `Zip` for one archive download
   - `Separate` for one download per file
5. Read `documentsOutputJson` or `downloadResultJson` from the outputs.

## SharePoint / Power Automate Pattern

If you want files in SharePoint instead of local browser download:

1. Set `downloadMode = Disabled`
2. Let the control produce `documentsOutputJson`
3. Send that JSON string to a Flow
4. In Flow, parse the JSON
5. Loop over the documents
6. Convert each `base64` field to binary
7. Create the files in SharePoint

This same pattern works for other storage or integration targets too.

## Docs

- [PCF usage guide](docs/pcf-usage.md)
- [Properties and outputs](docs/properties.md)
- [JSON schema](docs/json-schema.md)
- [Common MIME types](docs/mime-types.md)
- [Solution packaging](docs/solution-package.md)

## Known Limitations

- `Separate` download mode can be constrained by browser or host download policies.
- `archiveFileName` cannot be dynamically disabled in the maker property pane; it is simply ignored outside ZIP mode.

## Contributing

Contributions are welcome.

- Fork the repo and open a pull request, or use a branch + PR if you have access.
- Prefer pull requests over direct pushes to `main`.
- Start with an issue if you want to discuss a bug or feature request.

## Creator

Harllens George de la Cruz, The Boring Cat  
https://theboringcat.com/

## License

Apache-2.0. See [LICENSE](LICENSE).
