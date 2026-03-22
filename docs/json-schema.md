# JSON Schema

`documentsJson` must be a JSON array of document objects.

## Minimal shape

```json
[
  {
    "fileName": "customers",
    "fileExtension": "csv",
    "mimeType": "text/csv",
    "contentText": "Id,Name\n1,Ana"
  }
]
```

## Per-document fields

| Field | Required | Type | Notes |
| --- | --- | --- | --- |
| `id` | No | string | useful for correlation in app logic |
| `fileName` | Yes | string | file name without the extension |
| `fileExtension` | Yes | string | for example `txt`, `csv`, `json` |
| `mimeType` | Yes | string | browser or consumer content type |
| `contentText` | Yes | string | text content to export |

## Full example

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

## Validation behavior

- Missing required fields make that document invalid
- Invalid documents remain present in `documentsOutputJson`
- Each invalid document gets `isValid = false` and an `errorMessage`
- The control can still process valid documents in the same array
