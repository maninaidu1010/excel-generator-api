# Excel Generator API

Flask service that creates, edits, formats, enhances, validates and analyzes Excel workbooks.
Used by a Copilot Studio agent through a custom connector. Deploys to Render.

| Endpoint | Purpose |
|---|---|
| `GET /health` | status |
| `GET /excel-capabilities` | every operation, block, chart type and format |
| `POST /createexcel` | build a workbook from a `sheets` specification |
| `POST /editexcel` | apply `operations` to a workbook (`source.fileBase64` or `source.fileUrl`; CSV accepted) |
| `POST /analyzeexcel` | read-only inspection: structure, data quality, formula problems, health score |
| `GET /download/<file>` | download a generated workbook (kept 24 h) |

## Files
`app.py` (web app) - `excel_engine.py` (the engine) - `requirements.txt` - `render.yaml` - `openapi.json` (connector definition) - `sample_excel_payloads.json` - `smoke_test.py`

## Run locally
    pip install -r requirements.txt
    python smoke_test.py          # must end with: failures: 0
    python app.py                 # http://localhost:5000/health

## Render settings
Build `pip install -r requirements.txt` - Start `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120` - Health check `/health` - Python 3.12.

## Optional environment variables
`OUTPUT_DIR`, `FILE_TTL_SECONDS` (default 86400), `PUBLIC_BASE_URL`, `EXCEL_MAX_FILE_MB` (default 15), `EXCEL_LO_VERIFY`.

## Limits
Formulas are real but are calculated by Excel when the file is opened (the server verifies them only if LibreOffice is installed, which it is not on Render's Python runtime). Summary tables are formula-based, not PivotTables. Editing an existing file cannot keep slicers, sparklines, threaded comments or form controls (reported in `preservationWarnings`). Generated files are temporary. Max 15 MB, 80 operations per request, 50,000 rows per created table.
