"""
api/upload.py — recibe archivo Excel (multipart) y lo guarda en Supabase Storage.
Campos: type (base|cap), file
"""
import json
from http.server import BaseHTTPRequestHandler
from email import message_from_bytes
from email.policy import HTTP as EMAIL_HTTP

from _store import upload_bytes, read_state, write_state, StoreError

XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _parse_multipart(content_type: str, body: bytes):
    msg_bytes = f"Content-Type: {content_type}\r\n\r\n".encode() + body
    msg = message_from_bytes(msg_bytes, policy=EMAIL_HTTP)
    fields, files = {}, {}
    for part in msg.iter_parts():
        cd = part.get("Content-Disposition", "")
        name = filename = None
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
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        ct = self.headers.get("Content-Type", "")

        if "multipart/form-data" not in ct:
            _send_json(self, {"ok": False, "msg": "Se requiere multipart/form-data"}, 400)
            return

        fields, files = _parse_multipart(ct, body)
        ftype = fields.get("type", "base")
        if ftype not in ("base", "cap"):
            ftype = "base"

        if "file" not in files:
            _send_json(self, {"ok": False, "msg": "Campo 'file' faltante"}, 400)
            return

        orig_name, file_bytes = files["file"]
        try:
            upload_bytes(f"uploads/{ftype}.xlsx", file_bytes, XLSX_CT)
            state = read_state()
            state[ftype] = orig_name
            write_state(state)
        except StoreError as e:
            _send_json(self, {"ok": False, "msg": str(e)}, 500)
            return

        _send_json(self, {
            "ok": True,
            "type": ftype,
            "file": orig_name,
            "size": len(file_bytes),
        })

    def log_message(self, fmt, *args):
        pass
