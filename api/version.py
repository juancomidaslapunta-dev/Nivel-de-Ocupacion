"""
api/version.py — retorna JSON con versión actual y timestamp
"""
import os
import json
from http.server import BaseHTTPRequestHandler

VERSION_FILE = "/tmp/version.json"


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if os.path.exists(VERSION_FILE):
            with open(VERSION_FILE) as f:
                data = json.load(f)
        else:
            data = {"v": 0, "ts": "—", "label": ""}

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
