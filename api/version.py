"""
api/version.py — retorna JSON con versión actual y timestamp (desde Supabase).
"""
import json
import traceback
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        try:
            from _store import read_state
            st = read_state()
            data = {"v": st.get("v", 0), "ts": st.get("ts", "—"),
                    "label": st.get("label", "")}
        except Exception as e:
            data = {"v": 0, "ts": "—", "label": "",
                    "debug": f"{type(e).__name__}: {e}",
                    "trace": traceback.format_exc()[-1200:]}

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
