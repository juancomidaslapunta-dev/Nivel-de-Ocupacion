"""
api/report.py — sirve el HTML de reporte desde Supabase Storage.
"""
from http.server import BaseHTTPRequestHandler

from api._store import download_bytes, StoreError

REPORT_OBJECT = "report/reporte.html"

_NO_REPORT_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sin reporte</title>
  <link href="https://fonts.googleapis.com/css2?family=Fira+Sans:wght@300;400;600&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Fira Sans', system-ui, sans-serif;
      background: #0F172A; color: #94A3B8;
      display: flex; align-items: center; justify-content: center;
      min-height: 100vh; flex-direction: column; gap: 16px;
      text-align: center; padding: 24px;
    }
    svg { opacity: .35; }
    h2 { color: #F1F5F9; font-size: 1.1rem; font-weight: 600; }
    p  { font-size: .85rem; line-height: 1.6; max-width: 340px; }
    .steps { font-size: .8rem; line-height: 2; color: #64748B; }
    .step  { display: flex; align-items: center; gap: 10px; }
    .num   { background: #1E293B; border: 1px solid #334155; border-radius: 50%;
              width: 22px; height: 22px; display: flex; align-items: center; justify-content: center;
              font-size: .7rem; font-weight: 700; flex-shrink: 0; color: #93C5FD; }
  </style>
</head>
<body>
  <svg width="52" height="52" viewBox="0 0 24 24" fill="none" stroke="#475569" stroke-width="1.5"
       stroke-linecap="round" stroke-linejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
    <line x1="16" y1="13" x2="8" y2="13"/>
    <line x1="16" y1="17" x2="8" y2="17"/>
    <polyline points="10 9 9 9 8 9"/>
  </svg>
  <h2>Reporte no generado aún</h2>
  <p>Sube el archivo Excel y pulsa <strong>Actualizar Reporte</strong> para generar el análisis.</p>
  <div class="steps">
    <div class="step"><span class="num">1</span> Sube el archivo Stock Físico (.xlsx)</div>
    <div class="step"><span class="num">2</span> Opcionalmente sube la Tabla de Capacidades</div>
    <div class="step"><span class="num">3</span> Pulsa "Actualizar Reporte"</div>
  </div>
</body>
</html>"""


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        data = None
        try:
            data = download_bytes(REPORT_OBJECT)
        except StoreError:
            data = None

        if data:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(data)
        else:
            body = _NO_REPORT_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
