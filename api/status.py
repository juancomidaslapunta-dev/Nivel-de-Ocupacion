"""
api/status.py — retorna estado de archivos activos y versión (desde Supabase).
"""
import json
from http.server import BaseHTTPRequestHandler

from _store import read_state, StoreError


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        try:
            st = read_state()
            data = {
                "base":    st.get("base"),
                "cap":     st.get("cap"),
                "label":   st.get("label", ""),
                "version": st.get("v", 0),
                "ts":      st.get("ts", "—"),
                "running": False,
                "error":   None,
            }
        except StoreError as e:
            data = {"base": None, "cap": None, "label": "", "version": 0,
                    "ts": "—", "running": False, "error": str(e)}

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
