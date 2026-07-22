"""
api/regenerate.py — ejecuta analisis_ocupacion_sep.py con archivos en /tmp/
"""
import os
import sys
import json
import time
import subprocess
from http.server import BaseHTTPRequestHandler
from pathlib import Path

TMP_DIR    = "/tmp/uploads"
OUTPUT_DIR = "/tmp"
BASE_DIR   = Path(__file__).parent.parent   # raíz del repo


def _read_active(ftype: str):
    meta = os.path.join(TMP_DIR, f"{ftype}_active.txt")
    if os.path.exists(meta):
        with open(meta) as f:
            path = f.read().strip()
        if os.path.exists(path):
            return path
    return None


def _read_version():
    vf = os.path.join(OUTPUT_DIR, "version.json")
    if os.path.exists(vf):
        with open(vf) as f:
            return json.load(f)
    return {"v": 0, "ts": "—", "label": ""}


def _write_version(v: int, ts: str, label: str):
    vf = os.path.join(OUTPUT_DIR, "version.json")
    with open(vf, "w") as f:
        json.dump({"v": v, "ts": ts, "label": label}, f)


def _send_json(handler, data: dict, status: int = 200):
    body = json.dumps(data).encode()
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.end_headers()
    handler.wfile.write(body)


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        os.makedirs(TMP_DIR, exist_ok=True)

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        try:
            data = json.loads(body) if body else {}
        except Exception:
            data = {}

        label = data.get("label", "")

        base_file = _read_active("base")
        cap_file  = _read_active("cap")

        if not base_file:
            _send_json(self, {"ok": False, "msg": "No hay archivo de stock. Sube primero el Excel de Stock Físico."}, 400)
            return

        # Construir comando
        analysis_py = str(BASE_DIR / "analisis_ocupacion_sep.py")
        cmd = [sys.executable, analysis_py,
               "--base",       base_file,
               "--output-dir", OUTPUT_DIR]
        if cap_file:
            cmd += ["--cap", cap_file]
        if label:
            cmd += ["--label", label]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                cwd=str(BASE_DIR),
                env={**os.environ, "PYTHONPATH": str(BASE_DIR)},
            )
            if result.returncode == 0:
                cur = _read_version()
                new_v = cur["v"] + 1
                ts = time.strftime("%H:%M:%S")
                _write_version(new_v, ts, label)
                _send_json(self, {"ok": True, "version": new_v, "ts": ts})
            else:
                err = result.stderr[-800:] if result.stderr else "error desconocido"
                _send_json(self, {"ok": False, "msg": err}, 500)
        except subprocess.TimeoutExpired:
            _send_json(self, {"ok": False, "msg": "El análisis tardó demasiado (>120s). Intenta con un archivo más pequeño."}, 500)
        except Exception as e:
            _send_json(self, {"ok": False, "msg": str(e)}, 500)

    def log_message(self, fmt, *args):
        pass
