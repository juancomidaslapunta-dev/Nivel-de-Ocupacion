"""
api/_store.py — capa de persistencia sobre Supabase Storage.

Reemplaza el uso de /tmp (efímero y no compartido entre lambdas de Vercel)
por un bucket de Supabase, que sí es compartido y persistente.

Requiere variables de entorno en Vercel:
  SUPABASE_URL          → https://xxxx.supabase.co
  SUPABASE_SERVICE_KEY  → service_role key (Settings ▸ API)
  SUPABASE_BUCKET       → opcional, por defecto "ocupacion"

Objetos que se guardan en el bucket:
  uploads/base.xlsx     → stock físico subido
  uploads/cap.xlsx      → tabla de capacidades subida (opcional)
  report/reporte.html   → reporte generado
  report/reporte.xlsx   → reporte generado (excel)
  state.json            → { v, ts, label, base, cap }
"""
import os
import json
import urllib.request
import urllib.error

SUPABASE_URL = (os.environ.get("SUPABASE_URL") or "").rstrip("/")
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_KEY") or ""
BUCKET       = os.environ.get("SUPABASE_BUCKET") or "ocupacion"

STATE_PATH = "state.json"
DEFAULT_STATE = {"v": 0, "ts": "—", "label": "", "base": None, "cap": None}


class StoreError(Exception):
    pass


def _check_config():
    if not SUPABASE_URL or not SERVICE_KEY:
        raise StoreError(
            "Faltan variables de entorno SUPABASE_URL / SUPABASE_SERVICE_KEY en Vercel."
        )


def _object_url(path: str) -> str:
    return f"{SUPABASE_URL}/storage/v1/object/{BUCKET}/{path}"


def _request(method: str, url: str, data: bytes = None, content_type: str = None,
             upsert: bool = False):
    headers = {
        "Authorization": f"Bearer {SERVICE_KEY}",
        "apikey": SERVICE_KEY,
    }
    if content_type:
        headers["Content-Type"] = content_type
    if upsert:
        headers["x-upsert"] = "true"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    return urllib.request.urlopen(req, timeout=30)


def upload_bytes(path: str, data: bytes, content_type: str) -> None:
    """Sube (o reemplaza) un objeto al bucket."""
    _check_config()
    try:
        _request("POST", _object_url(path), data=data,
                 content_type=content_type, upsert=True)
    except urllib.error.HTTPError as e:
        detail = e.read().decode(errors="replace")
        raise StoreError(f"Error subiendo {path}: {e.code} {detail}")


def download_bytes(path: str):
    """Descarga un objeto. Devuelve bytes o None si no existe (404)."""
    _check_config()
    try:
        resp = _request("GET", _object_url(path))
        return resp.read()
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        detail = e.read().decode(errors="replace")
        raise StoreError(f"Error descargando {path}: {e.code} {detail}")


def read_state() -> dict:
    raw = download_bytes(STATE_PATH)
    if raw is None:
        return dict(DEFAULT_STATE)
    try:
        st = json.loads(raw)
    except Exception:
        return dict(DEFAULT_STATE)
    merged = dict(DEFAULT_STATE)
    merged.update(st)
    return merged


def write_state(state: dict) -> None:
    upload_bytes(STATE_PATH, json.dumps(state).encode(), "application/json")
