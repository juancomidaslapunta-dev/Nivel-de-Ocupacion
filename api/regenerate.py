"""
api/regenerate.py — ejecuta el análisis EN EL MISMO PROCESO (sin subprocess).

Descarga los archivos desde Supabase Storage a /tmp, corre el pipeline de
analisis_ocupacion_sep.py importándolo como módulo, y sube el reporte HTML/XLSX
de vuelta al bucket. Finalmente incrementa la versión en state.json.
"""
import os
import sys
import json
import time
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from _store import (upload_bytes, download_bytes, read_state, write_state,
                    StoreError)

BASE_DIR = Path(__file__).resolve().parent.parent   # raíz del repo
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

TMP_DIR   = "/tmp"
XLSX_CT   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HTML_CT   = "text/html; charset=utf-8"


def _send_json(handler, data: dict, status: int = 200):
    body = json.dumps(data).encode()
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.end_headers()
    handler.wfile.write(body)


def _run_analysis(base_path: str, cap_path, label: str, display_name: str = None):
    """Corre el pipeline y devuelve (html_path, xlsx_path)."""
    import analisis_ocupacion_sep as an

    out_html = os.path.join(TMP_DIR, "reporte_ocupacion_sep.html")
    out_xlsx = os.path.join(TMP_DIR, "reporte_ocupacion_sep.xlsx")

    # Configurar globales del módulo (equivalente a los args del CLI).
    # INPUT_FILE se usa solo para mostrar el nombre de origen en el reporte;
    # la lectura real se hace con la ruta temporal explícita más abajo.
    an.INPUT_FILE       = display_name or base_path
    an.CAPACIDADES_FILE = cap_path
    an.LABEL            = label or ""
    an.OUTPUT_DIR       = TMP_DIR
    an.OUTPUT_XLSX      = out_xlsx
    an.OUTPUT_HTML      = out_html

    xls      = an.pd.ExcelFile(base_path)
    stock, _ = an.load_sku_master(xls)
    capacity = an.load_capacity(xls)
    result   = an.consolidate(stock, capacity)
    an.write_excel(result, out_xlsx)
    an.write_html(result, stock, out_html)
    return out_html, out_xlsx


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        try:
            data = json.loads(body) if body else {}
        except Exception:
            data = {}
        label = data.get("label", "")

        try:
            state = read_state()

            # Descargar archivos activos desde el bucket
            base_bytes = download_bytes("uploads/base.xlsx")
            if base_bytes is None:
                _send_json(self, {"ok": False, "msg": "No hay archivo de stock. Sube primero el Excel de Stock Físico."}, 400)
                return
            base_path = os.path.join(TMP_DIR, "base.xlsx")
            with open(base_path, "wb") as f:
                f.write(base_bytes)

            cap_path = None
            cap_bytes = download_bytes("uploads/cap.xlsx")
            if cap_bytes is not None:
                cap_path = os.path.join(TMP_DIR, "cap.xlsx")
                with open(cap_path, "wb") as f:
                    f.write(cap_bytes)

            # Ejecutar análisis
            out_html, out_xlsx = _run_analysis(
                base_path, cap_path, label, display_name=state.get("base"))

            # Subir resultados al bucket
            with open(out_html, "rb") as f:
                upload_bytes("report/reporte.html", f.read(), HTML_CT)
            if os.path.exists(out_xlsx):
                with open(out_xlsx, "rb") as f:
                    upload_bytes("report/reporte.xlsx", f.read(), XLSX_CT)

            # Bump de versión
            ts = time.strftime("%H:%M:%S")
            state["v"]     = int(state.get("v", 0)) + 1
            state["ts"]    = ts
            state["label"] = label
            write_state(state)

            _send_json(self, {"ok": True, "version": state["v"], "ts": ts})

        except StoreError as e:
            _send_json(self, {"ok": False, "msg": str(e)}, 500)
        except Exception as e:
            import traceback
            _send_json(self, {"ok": False, "msg": f"{e}", "trace": traceback.format_exc()[-800:]}, 500)

    def log_message(self, fmt, *args):
        pass
