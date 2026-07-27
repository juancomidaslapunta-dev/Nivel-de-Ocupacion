import sys
import traceback
from http.server import BaseHTTPRequestHandler

try:
    from api._store import download_bytes, StoreError
    IMPORT_ERR = None
except Exception as e:
    IMPORT_ERR = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"

REPORT_OBJECT = "report/reporte.html"
_NO_REPORT_HTML = "<html><body><h1>Reporte no generado aún</h1></body></html>"

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if IMPORT_ERR:
            body = IMPORT_ERR.encode("utf-8")
            self.send_response(500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        data = None
        try:
            data = download_bytes(REPORT_OBJECT)
        except Exception:
            data = None

        if data:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            body = _NO_REPORT_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
