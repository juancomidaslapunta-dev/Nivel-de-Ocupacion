"""
api/status.py — retorna estado de archivos activos y versión
"""
import os
import json
from http.server import BaseHTTPRequestHandler

TMP_DIR      = "/tmp/uploads"
VERSION_FILE = "/tmp/version.json"


def _read_active(ftype: str):
    meta = os.path.join(TMP_DIR, f"{ftype}_active.txt")
    if os.path.exists(meta):
        with open(meta) as f:
            path = f.read().strip()
        if os.path.exists(path):
            return path
    return None


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if os.path.exists(VERSION_FILE):
            with open(VERSION_FILE) as f:
                ver = json.load(f)
        else:
            ver = {"v": 0, "ts": "—", "label": ""}

        data = {
            "base":    _read_active("base"),
            "cap":     _read_active("cap"),
            "label":   ver.get("label", ""),
            "version": ver.get("v", 0),
            "ts":      ver.get("ts", "—"),
            "running": False,
            "error":   None,
        }

        body = json.dumps(data).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
