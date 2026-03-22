# PCF Usage Guide

## Input Properties

- `documentsJson` (required): JSON array of documents to export
- `downloadMode`: `Disabled | Zip | Separate`
- `archiveFileName`: ZIP file name used only when `downloadMode = Zip`
- `displayModeInput`: `Edit | Disabled | View`
- Button style properties: text, colors, border, padding, alignment, font, and type

## Output Properties

- `documentsOutputJson`
- `downloadResultJson`
- `documentCount`
- `validDocumentCount`
- `isValid`
- `errorMessage`
- `lastAction`
- `changeToken`
- `selectToken`

## Example `documentsJson`

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
  }
]
```

See also:

- [JSON schema](json-schema.md)
- [Common MIME types](mime-types.md)

## Canvas Patterns

### Use the control for built-in downloads

- Set `downloadMode = Zip` for one archive download
- Set `downloadMode = Separate` for one browser download per valid file

### Use the control for Power Automate or external systems

- Set `downloadMode = Disabled`
- Read `documentsOutputJson`
- Send that JSON string to a Flow or custom connector

### React to button clicks in Power Fx

- Use the control `OnSelect` event
- Or observe `selectToken` when you need a click token in app state

## Parse `documentsOutputJson` In Power Apps

```powerfx
ClearCollect(
    colDocs,
    ForAll(
        Table(ParseJSON(FileExporterComponent1.documentsOutputJson)),
        {
            id: Text(Value.id),
            fileName: Text(Value.fileName),
            fileExtension: Text(Value.fileExtension),
            fullFileName: Text(Value.fullFileName),
            mimeType: Text(Value.mimeType),
            contentText: Text(Value.contentText),
            base64: Text(Value.base64),
            sizeBytes: Value(Text(Value.sizeBytes)),
            isValid: Boolean(Value.isValid),
            errorMessage: Text(Value.errorMessage)
        }
    )
)
```

## Send Files To SharePoint With Power Automate

Typical pattern:

1. Set `downloadMode = Disabled`
2. Wait for updated `documentsOutputJson`
3. Pass the JSON string to Flow
4. Use `Parse JSON` in Flow
5. Loop through each document
6. Convert `base64` to binary
7. Create the file in SharePoint

This lets the control handle file preparation while Flow handles storage and downstream integration.
