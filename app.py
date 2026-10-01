"""
Excel Generator API (Flask) - create / edit / analyze Excel workbooks.

Endpoints
  GET  /, /health            service status
  GET  /excel-capabilities   every operation, block, chart type, format
  POST /createexcel          build a workbook from a sheet specification
  POST /editexcel            apply operations to an uploaded workbook / CSV
  POST /analyzeexcel         inspect + validate a workbook (read-only)
  GET  /download/<file>      download a generated workbook

Env vars (all optional)
  OUTPUT_DIR          where generated files are written (default: system temp dir)
  FILE_TTL_SECONDS    how long files are kept (default 86400)
  PUBLIC_BASE_URL     override the base of download links
  EXCEL_MAX_FILE_MB   max accepted workbook size (default 15)
  EXCEL_LO_VERIFY     set to 0 to skip LibreOffice formula checks (only used if LibreOffice is installed)
"""
import logging
import os
import re
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from werkzeug.exceptions import HTTPException
from werkzeug.middleware.proxy_fix import ProxyFix

from excel_engine import ExcelError, create_blueprint

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("excel-api")

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR") or Path(tempfile.gettempdir()) / "generated_workbooks")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")
MIME = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # base64 inflates files by ~33%
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)


def base_url() -> str:
    if PUBLIC_BASE_URL:
        return PUBLIC_BASE_URL
    render_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")  # set automatically by Render
    return render_url or request.url_root.rstrip("/")


app.register_blueprint(create_blueprint(OUTPUT_DIR, base_url, log))


@app.get("/")
@app.get("/health")
def health():
    return jsonify({"status": "ok", "service": "excel-generator", "excel": True})


@app.get("/download/<path:filename>")
def download(filename: str):
    if not re.fullmatch(r"[A-Za-z0-9._-]+\.(xlsx|xlsm)", filename):
        raise ExcelError("Invalid file name.", status=400, code="INVALID_FILENAME")
    if not (OUTPUT_DIR / filename).is_file():
        raise ExcelError("File not found. Generated files expire after 24 hours or when the service restarts; "
                         "please generate it again.", status=404, code="FILE_NOT_FOUND")
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True, download_name=filename,
                               mimetype=MIME[filename.rsplit(".", 1)[-1].lower()])


@app.errorhandler(ExcelError)
def handle_excel_error(err: ExcelError):
    return jsonify({"status": "error", "errorCode": err.code, "message": err.message}), err.status


@app.errorhandler(HTTPException)
def handle_http_error(err: HTTPException):
    code = {404: "NOT_FOUND", 405: "METHOD_NOT_ALLOWED", 413: "PAYLOAD_TOO_LARGE"}.get(err.code or 500, "HTTP_ERROR")
    return jsonify({"status": "error", "errorCode": code, "message": err.description}), err.code or 500


@app.errorhandler(Exception)
def handle_unexpected(err: Exception):
    log.exception("Unhandled error")
    return jsonify({"status": "error", "errorCode": "INTERNAL_ERROR",
                    "message": "Something went wrong while processing the workbook. Please try again."}), 500


if __name__ == "__main__":  # local development only; Render uses gunicorn
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
