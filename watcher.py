"""
Watcher + servidor local — un archivo de stock físico + tabla de capacidades.

Endpoints:
  GET  /              → dashboard.html
  GET  /report        → reporte_ocupacion_sep.html
  GET  /version       → {"v": N, "ts": "HH:MM:SS", "label": "..."}
  GET  /status        → estado de archivos subidos
  POST /upload        → multipart: field "type" (base|cap) + file
  POST /regenerate    → body JSON: {"label":"..."}

Uso: python3 watcher.py
"""

import os
import sys

# Forzar UTF-8 en stdout para compatibilidad con Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import json
import time
import threading
import subprocess
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler
from email import message_from_bytes
from email.policy import HTTP as EMAIL_HTTP
from urllib.parse import urlparse

PORT        = 8080
POLL_SECS   = 3
BASE_DIR    = Path(__file__).parent
WATCH_FILE  = BASE_DIR / "maestra de codigosv2.xlsx"
REPORT_HTML = BASE_DIR / "reporte_ocupacion_sep.html"
ANALYSIS_PY = BASE_DIR / "analisis_ocupacion_sep.py"
DASHBOARD   = BASE_DIR / "dashboard.html"
UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)

# Archivos activos (se actualizan al subir)
_active = {
    "base": str(WATCH_FILE),
    "cap":  None,
}
_label  = {"label": ""}
_state  = {"version": 0, "ts": "—", "running": False, "error": None}
_lock   = threading.Lock()


# ── Regenerar ─────────────────────────────────────────────────────────────────
def _build_cmd():
    cmd = [sys.executable, str(ANALYSIS_PY)]
    if _active["base"]:   cmd += ["--base",  _active["base"]]
    if _active["cap"]:    cmd += ["--cap",   _active["cap"]]
    if _label["label"]:   cmd += ["--label", _label["label"]]
    return cmd


def regenerate(reason: str = "manual"):
    with _lock:
        if _state["running"]:
            return {"ok": False, "msg": "Ya hay regeneración en curso"}
        _state["running"] = True
        _state["error"]   = None

    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] Regenerando ({reason})…", flush=True)
    try:
        result = subprocess.run(
            _build_cmd(), capture_output=True, text=True,
            timeout=90, cwd=str(BASE_DIR)
        )
        if result.returncode == 0:
            _inject_autoreload()
            with _lock:
                _state["version"] += 1
                _state["ts"] = time.strftime("%H:%M:%S")
            print(f"[{_state['ts']}] ✅ v{_state['version']}", flush=True)
            return {"ok": True, "version": _state["version"], "ts": _state["ts"]}
        else:
            err = result.stderr[-600:] if result.stderr else "error desconocido"
            with _lock:
                _state["error"] = err
            print(f"[{ts}] ❌ {err}", flush=True)
            return {"ok": False, "msg": err}
    except Exception as e:
        with _lock:
            _state["error"] = str(e)
        print(f"[{ts}] ❌ {e}", flush=True)
        return {"ok": False, "msg": str(e)}
    finally:
        with _lock:
            _state["running"] = False


def _inject_autoreload():
    """Inserta JS de auto-recarga en el HTML generado (idempotente)."""
    if not REPORT_HTML.exists():
        return
    script = (
        '\n<script id="__ar__">'
        'let _v=null;'
        'setInterval(()=>fetch("/version?_="+Date.now())'
        '.then(r=>r.json()).then(d=>{if(_v===null){_v=d.v;return;}'
        'if(d.v!==_v){_v=d.v;location.reload();}})'
        '.catch(()=>{}),3000);'
        '</script>'
    )
    txt = REPORT_HTML.read_text(encoding="utf-8")
    if "__ar__" not in txt:
        txt = txt.replace("</body>", script + "\n</body>", 1)
        REPORT_HTML.write_text(txt, encoding="utf-8")


# ── Watcher loop ──────────────────────────────────────────────────────────────
def _watch_loop():
    last = 0.0
    print(f"[{time.strftime('%H:%M:%S')}] Monitoreando: {WATCH_FILE.name}", flush=True)
    # Solo regenerar al inicio si el archivo base existe
    if Path(_active["base"]).exists():
        regenerate("inicio")
    else:
        print(f"[{time.strftime('%H:%M:%S')}] Archivo base no encontrado: {_active['base']}", flush=True)
        print(f"[{time.strftime('%H:%M:%S')}] Suba el archivo Excel desde el dashboard para generar el reporte.", flush=True)
    while True:
        time.sleep(POLL_SECS)
        watch = Path(_active["base"]) if _active["base"] else WATCH_FILE
        try:
            mt = watch.stat().st_mtime
            if mt != last:
                last = mt
                time.sleep(1.0)
                threading.Thread(target=regenerate, args=("cambio en archivo",), daemon=True).start()
        except FileNotFoundError:
            pass


# ── Parsear multipart upload ──────────────────────────────────────────────────
def _parse_multipart(content_type: str, body: bytes):
    """Retorna dict {name: value} y dict {name: (filename, bytes)}."""
    # Reconstruir mensaje compatible con email parser
    msg_bytes = f"Content-Type: {content_type}\r\n\r\n".encode() + body
    msg = message_from_bytes(msg_bytes, policy=EMAIL_HTTP)
    fields, files = {}, {}
    for part in msg.iter_parts():
        cd = part.get("Content-Disposition", "")
        name = None
        filename = None
        for item in cd.split(";"):
            item = item.strip()
            if item.startswith('name="'):
                name = item[6:-1]
            elif item.startswith('filename="'):
                filename = item[10:-1]
        if name is None:
            continue
        payload = part.get_payload(decode=True) or b""
        if filename:
            files[name] = (filename, payload)
        else:
            fields[name] = payload.decode(errors="replace").strip()
    return fields, files


# ── HTTP Handler ──────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path, mime: str = "text/html"):
        if not path.exists():
            self.send_error(404)
            return
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", len(data))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        path   = parsed.path.rstrip("/") or "/"

        if path in ("/", "/dashboard"):
            self._send_file(DASHBOARD)
        elif path == "/report":
            self._send_file(REPORT_HTML)
        elif path == "/version":
            self._send_json({
                "v":     _state["version"],
                "ts":    _state["ts"],
                "label": _label["label"],
            })
        elif path == "/status":
            self._send_json({
                "base":    _active["base"],
                "cap":     _active["cap"],
                "label":   _label["label"],
                "version": _state["version"],
                "ts":      _state["ts"],
                "running": _state["running"],
                "error":   _state["error"],
            })
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        length = int(self.headers.get("Content-Length", 0))
        body   = self.rfile.read(length)

        if path == "/upload":
            ct = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ct:
                self._send_json({"ok": False, "msg": "Se requiere multipart/form-data"}, 400)
                return

            fields, files = _parse_multipart(ct, body)
            ftype = fields.get("type", "base")   # base | cap

            if "file" not in files:
                self._send_json({"ok": False, "msg": "Campo 'file' faltante"}, 400)
                return

            orig_name, file_bytes = files["file"]
            dest = UPLOADS_DIR / f"{ftype}_{orig_name}"
            dest.write_bytes(file_bytes)

            if ftype in _active:
                _active[ftype] = str(dest)
            print(f"[{time.strftime('%H:%M:%S')}] Upload {ftype}: {orig_name} ({len(file_bytes):,} bytes)",
                  flush=True)

            self._send_json({"ok": True, "type": ftype, "file": orig_name, "path": str(dest)})

        elif path == "/regenerate":
            try:
                data = json.loads(body) if body else {}
            except json.JSONDecodeError:
                data = {}

            _label["label"] = data.get("label", _label["label"])

            threading.Thread(
                target=lambda: self._send_json(regenerate("dashboard")),
                daemon=True
            ).start()
            self._send_json({"ok": True, "msg": "Regenerando…"})

        else:
            self.send_error(404)

    def log_message(self, fmt, *args):
        if args and str(args[1]) not in ("200", "304"):
            super().log_message(fmt, *args)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    os.chdir(BASE_DIR)
    threading.Thread(target=_watch_loop, daemon=True).start()

    server = HTTPServer(("localhost", PORT), Handler)
    url    = f"http://localhost:{PORT}"

    print(f"""
  +---------------------------------------------+
  |  Watcher + servidor activo                  |
  |  URL: {url:<38}|
  |  Ctrl+C para detener                        |
  +---------------------------------------------+
""", flush=True)

    def _open():
        time.sleep(2)
        import webbrowser
        webbrowser.open(url)

    threading.Thread(target=_open, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print(f"\n[{time.strftime('%H:%M:%S')}] Detenido.")
        server.shutdown()


if __name__ == "__main__":
    main()
