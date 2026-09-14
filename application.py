from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime, date
from pathlib import Path
import hashlib
import tempfile
import zipfile
import base64
import urllib.request
import urllib.error

import streamlit as st

# ============================================================
# AGENTE TIARA 2.0
# Cerebro determinista basado en el Excel de septiembre 2026
# ============================================================

ROOT = Path(__file__).resolve().parent
APP_VERSION = "2.2.0"
APP_UPDATED = "14/09/2026"
DATA_DIR = ROOT / "data"
JSON_FILE = DATA_DIR / "base_conocimiento_tiara_septiembre_2026_v2.json"
JSON_FILE_OLD = DATA_DIR / "base_conocimiento_tiara_septiembre_2026.json"
EXCEL_CANDIDATES = [
    ROOT / "VITACORA DE BARCO  SEPTIEMBRE 2026.xlsx",
    ROOT / "bitácora barco septiembre 2026.xlsx",
    DATA_DIR / "VITACORA DE BARCO  SEPTIEMBRE 2026.xlsx",
    DATA_DIR / "bitácora barco septiembre 2026.xlsx",
]

MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


# ============================================================
# VISIÓN / FOTOGRAFÍAS DEL EXCEL
# ============================================================
# La clave de Gemini se lee exclusivamente desde Streamlit Secrets.
# Nunca se muestra ni se guarda dentro de la base de conocimiento.

def _gemini_api_key():
    try:
        cfg = st.secrets.get("gemini", {})
        key = str(cfg.get("api_key", "")).strip()
        if key:
            return key
        # Compatibilidad opcional con un secret plano.
        key = str(st.secrets.get("GEMINI_API_KEY", "")).strip()
        return key or None
    except Exception:
        return None


def _prepare_image_for_vision(raw: bytes, fmt: str):
    """Reduce una foto si es demasiado grande, sin alterar la original del Excel."""
    if not raw:
        return raw, ("jpeg" if fmt == "jpg" else fmt)
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(raw))
        img.load()
        # Gemini acepta JPEG/PNG/WEBP. JPEG reduce mucho el tamaño de fotos.
        max_side = 1800
        if max(img.size) > max_side:
            ratio = max_side / float(max(img.size))
            img = img.resize((max(1, int(img.width*ratio)), max(1, int(img.height*ratio))), Image.LANCZOS)
        out = io.BytesIO()
        if fmt.lower() in ("png",) and img.mode in ("RGBA", "LA"):
            img.save(out, format="PNG", optimize=True)
            return out.getvalue(), "png"
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        img.save(out, format="JPEG", quality=82, optimize=True)
        return out.getvalue(), "jpeg"
    except Exception:
        return raw, ("jpeg" if fmt == "jpg" else fmt)


def _gemini_vision_request(image_bytes: bytes, mime_type: str, prompt: str):
    key = _gemini_api_key()
    if not key:
        return None, "NO_API_KEY"
    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        "gemini-2.5-flash-lite:generateContent?key=" + key
    )
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inline_data": {
                        "mime_type": mime_type,
                        "data": base64.b64encode(image_bytes).decode("ascii"),
                    }
                },
            ]
        }],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 900,
        },
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        endpoint, data=body,
        headers={"Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=45) as response:
            result = json.loads(response.read().decode("utf-8"))
        parts = (((result.get("candidates") or [{}])[0].get("content") or {}).get("parts") or [])
        text = " ".join(str(x.get("text", "")).strip() for x in parts if x.get("text")).strip()
        return text or None, None
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        return None, f"HTTP {exc.code}: {detail[:300]}"
    except Exception as exc:
        return None, str(exc)


def _analyze_embedded_images(images):
    """Interpreta fotografías incrustadas y devuelve solo texto estructurado.

    Las imágenes binarias se eliminan después del análisis para no inflar el
    snapshot de Supabase. La fotografía original sigue siendo la fuente en el
    Excel del usuario.
    """
    if not images:
        return images, {"enabled": False, "analyzed": 0, "reason": "sin_fotos"}
    if not _gemini_api_key():
        # No enviamos ninguna imagen si el usuario todavía no configuró la API.
        clean = []
        for im in images:
            clean.append({k:v for k,v in im.items() if k != "data_base64"})
        return clean, {"enabled": False, "analyzed": 0, "reason": "sin_api_key"}

    analyzed = []
    max_images = 30
    for im in images[:max_images]:
        b64 = im.get("data_base64")
        if not b64:
            continue
        try:
            raw = base64.b64decode(b64)
            fmt = str(im.get("format") or "jpeg").lower()
            raw2, fmt2 = _prepare_image_for_vision(raw, fmt)
            mime = "image/png" if fmt2 == "png" else ("image/webp" if fmt2 == "webp" else "image/jpeg")
            prompt = (
                "Eres el módulo de visión del Agente Tiara. Analiza esta fotografía "
                "incrustada en un Excel de permisos/seguros de una embarcación. "
                "Extrae SOLO información realmente visible: tipo de documento, "
                "nombre de embarcación/empresa/persona si se lee, número de póliza, "
                "permiso/certificado si se lee, fechas de emisión y vencimiento, "
                "aseguradora/autoridad, y cualquier dato importante visible. "
                "Si algo no se puede leer, dilo como 'no legible'. No inventes ni "
                "completes datos faltantes. Responde en español y en formato breve."
            )
            text, err = _gemini_vision_request(raw2, mime, prompt)
            item = {k:v for k,v in im.items() if k != "data_base64"}
            item["vision_analysis"] = text or "No se pudo interpretar visualmente esta fotografía."
            item["vision_error"] = err if err and err != "NO_API_KEY" else None
            analyzed.append(item)
        except Exception as exc:
            item = {k:v for k,v in im.items() if k != "data_base64"}
            item["vision_analysis"] = "No se pudo procesar esta fotografía."
            item["vision_error"] = str(exc)
            analyzed.append(item)

    # Conserva metadata de fotos que quedaron fuera del límite, sin bytes.
    for im in images[max_images:]:
        analyzed.append({k:v for k,v in im.items() if k != "data_base64"})
    return analyzed, {"enabled": True, "analyzed": min(len(images), max_images), "total": len(images)}


# ============================================================
# PERSISTENCIA SUPABASE
# ============================================================
# La aplicación corre en el servidor de Streamlit. La clave de Supabase
# se lee exclusivamente desde st.secrets y nunca se muestra en pantalla.
SUPABASE_TABLE = "tiara_data"

def _supabase_credentials():
    try:
        cfg = st.secrets.get("supabase", {})
        url = str(cfg.get("url", "")).strip().rstrip("/")
        # La llave pública puede permanecer en `key`; la persistencia usa
        # exclusivamente la llave privada de servidor `service_key`.
        key = str(cfg.get("service_key", "")).strip()
        if url and key:
            return url, key
    except Exception:
        pass
    return None, None


def _supabase_request(method, path, payload=None, query=""):
    import urllib.request
    import urllib.error

    base_url, key = _supabase_credentials()
    if not base_url or not key:
        raise RuntimeError("No están configurados los Secrets de Supabase en Streamlit.")

    url = f"{base_url}/rest/v1/{SUPABASE_TABLE}{path}{query}"
    headers = {
        "apikey": key,
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }
    body = None if payload is None else json.dumps(payload, ensure_ascii=False, allow_nan=False).encode("utf-8")
    request = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Supabase respondió HTTP {exc.code}: {detail[:500]}") from exc
    except Exception as exc:
        # Streamlit Cloud puede tardar en recibir la respuesta de Supabase
        # aunque el servidor sí esté procesando el snapshot. Reintentamos
        # una sola vez en operaciones de escritura para tolerar ese timeout.
        if method.upper() in {"POST", "PATCH", "PUT"} and "timed out" in str(exc).lower():
            try:
                with urllib.request.urlopen(request, timeout=180) as response:
                    raw = response.read().decode("utf-8")
                    return json.loads(raw) if raw else None
            except urllib.error.HTTPError as retry_exc:
                detail = retry_exc.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"Supabase respondió HTTP {retry_exc.code}: {detail[:500]}") from retry_exc
            except Exception as retry_exc:
                raise RuntimeError(f"No se pudo conectar con Supabase tras reintentar: {retry_exc}") from retry_exc
        raise RuntimeError(f"No se pudo conectar con Supabase: {exc}") from exc


def load_supabase_data():
    """Carga el último snapshot activo guardado en Supabase."""
    try:
        rows = _supabase_request(
            "GET",
            "",
            query=(
                "?select=data,source_filename,source_hash,version_number,created_at"
                "&is_active=eq.true&order=created_at.desc&limit=1"
            ),
        )
        if not rows:
            return None
        row = rows[0]
        data = row.get("data")
        if not isinstance(data, dict):
            return None
        meta = data.setdefault("metadata", {})
        if row.get("source_filename") and not meta.get("source_file"):
            meta["source_file"] = row["source_filename"]
        if row.get("source_hash") and not meta.get("upload_sha256"):
            meta["upload_sha256"] = row["source_hash"]
        if row.get("version_number") is not None and not meta.get("version"):
            # version_number es un entero en Supabase; conservamos la versión
            # legible del snapshot cuando ya viene dentro de metadata.
            meta["version_number"] = row["version_number"]
        meta["supabase_persisted"] = True
        meta["supabase_created_at"] = row.get("created_at")
        return data
    except Exception as exc:
        st.session_state.supabase_load_error = str(exc)
        return None


def save_supabase_data(data):
    """Guarda un snapshot completo de la base normalizada del Excel."""
    meta = data.get("metadata", {})
    payload = {
        "source_filename": meta.get("source_file", "Excel Tiara"),
        "source_hash": meta.get("upload_sha256"),
        "version_number": 210,
        "is_active": True,
        "data": data,
    }
    _supabase_request("POST", "", payload=payload)


MONTH_NAMES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

STOPWORDS = {
    "el", "la", "los", "las", "de", "del", "un", "una", "unos", "unas",
    "que", "qué", "se", "ha", "han", "le", "les", "al", "en", "por",
    "para", "con", "y", "o", "a", "lo", "me", "nos", "su", "sus",
    "fue", "ser", "es", "son", "hay", "cuál", "cual", "cuáles", "cuales",
    "cuando", "cuándo", "como", "cómo", "toca", "tiene", "tienen", "dame",
    "quiero", "saber", "sobre", "del", "durante", "este", "esta", "del",
}


def norm(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s


def safe_float(value: object):
    if value is None or value == "":
        return None
    try:
        x = float(value)
        return x if x == x else None
    except Exception:
        s = str(value).replace("$", "").replace(",", "").strip()
        try:
            return float(s)
        except Exception:
            return None


def parse_date(value: object):
    if isinstance(value, (datetime, date)):
        return value.date() if isinstance(value, datetime) else value
    if not value:
        return None
    s = str(value).strip()
    for fmt in ("%d-%m-%y", "%d-%m-%Y", "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def extract_dates(value: object):
    out = []
    if value is None:
        return out
    for part in re.split(r"[|\n]+", str(value)):
        d = parse_date(part.strip())
        if d:
            out.append(d)
    return out


def extract_hours(value: object):
    out = []
    if value is None:
        return out
    for part in re.split(r"[|\n]+", str(value)):
        part = part.strip()
        if re.fullmatch(r"\d+(?:\.0)?", part):
            try:
                out.append(int(float(part)))
            except Exception:
                pass
    return out


def fmt_date(d):
    return d.strftime("%d-%m-%Y") if d else "NO DETERMINADO"


def money(v):
    return f"${v:,.2f} USD"


@st.cache_data(show_spinner=False)
def load_data():
    # 1) Preferir la base v2 si existe.
    if JSON_FILE.exists():
        with JSON_FILE.open("r", encoding="utf-8") as f:
            return normalize_loaded_data(json.load(f))

    # 2) Si está disponible el Excel original, reconstruir la base desde la fuente.
    for path in EXCEL_CANDIDATES:
        if path.exists():
            return build_from_excel(path)

    # 3) Compatibilidad con la base JSON antigua que ya está en GitHub.
    if JSON_FILE_OLD.exists():
        with JSON_FILE_OLD.open("r", encoding="utf-8") as f:
            return normalize_loaded_data(json.load(f))

    return None


def excel_serial_to_date(value):
    if value is None:
        return None
    try:
        n = float(str(value).strip())
    except Exception:
        return None
    # Fechas seriales de Excel (sistema 1900).
    if 20000 <= n <= 80000:
        from datetime import timedelta
        return date(1899, 12, 30) + timedelta(days=int(n))
    return None


def normalize_loaded_data(data):
    """Convierte la base antigua 1.0 a la estructura que usa el agente 2.0."""
    if not data:
        return data

    # Ya está normalizada.
    if "inventory_records" in data and "documents_records" in data:
        return data

    out = {
        "metadata": dict(data.get("metadata", {})),
        "sheets": dict(data.get("sheets", {})),
        "maintenance_records": [],
        "log_records": [],
        "inventory_records": [],
        "fuel_records": [],
        "budget_records": [],
        "invoice_records": [],
        "cleaning_records": [],
        "documents_records": [],
        "images": [],
    }

    out["metadata"]["version"] = APP_VERSION

    # Mantenimiento: normalizar los nombres de columnas de la base 1.0.
    for rec in data.get("maintenance_records", []):
        d = rec.get("data", {})
        def first_key(*keys):
            for k in keys:
                if k in d and d[k] not in (None, ""):
                    return d[k]
            return None

        dates_raw = first_key("fecha", "fecha cambio o mantenimiento", "FECHA")
        dates = []
        for part in re.split(r"[|\n]+", str(dates_raw or "")):
            part = part.strip()
            dt = parse_date(part) or excel_serial_to_date(part)
            if dt:
                dates.append(dt.isoformat())

        hours_raw = first_key("historial_horas", "historial mantenimientos ", "historial mantenimientos", "HISTORIAL MANTENIMIENTOS")
        interval = first_key("cada_horas", "cada (horas)", "cada horas", "CADA (HORAS)")
        normalized = {
            "mantenimiento": first_key("mantenimiento", "MANTENIMIENTO"),
            "cada_horas": interval,
            "historial_horas": hours_raw,
            "fecha": "|".join(dates),
            "observaciones": first_key("observaciones", "OBSERVACIONES"),
        }
        out["maintenance_records"].append({
            "source_sheet": rec.get("source_sheet", "CHECKLIST MANT. PREVENTIVO"),
            "source_row": rec.get("source_row"),
            "data": normalized,
        })

    # Los demás módulos conservan sus datos y además reciben un campo values
    # para que el cerebro 2.0 pueda consultarlos sin romper la base antigua.
    mappings = {
        "log_records": "log_records",
        "inventory_records": "inventory_items",
        "fuel_records": "fuel_records",
        "budget_records": "budget_records",
        "invoice_records": "invoice_records",
        "cleaning_records": "cleaning_records",
        "documents_records": "documents",
    }
    for target, old_key in mappings.items():
        for rec in data.get(old_key, []):
            vals = list((rec.get("data") or {}).values())
            if target == "log_records":
                # Reinterpretar la bitácora antigua con la nueva estructura
                # para que el estado, sistema y horas queden disponibles.
                parsed = parse_log_record(vals, rec.get("source_row"))
                parsed["source_sheet"] = rec.get("source_sheet")
                out[target].append(parsed)
            else:
                out[target].append({
                    "source_sheet": rec.get("source_sheet"),
                    "source_row": rec.get("source_row"),
                    "values": vals,
                    "data": rec.get("data", {}),
                })

    return out


def _find_sheet(wb, expected_name):
    """Encuentra una hoja tolerando tildes, espacios y pequeñas variaciones."""
    target = norm(expected_name).replace(" ", "")
    for ws in wb.worksheets:
        candidate = norm(ws.title).replace(" ", "")
        if candidate == target:
            return ws
    # Coincidencia parcial como respaldo para nombres que cambian ligeramente.
    for ws in wb.worksheets:
        candidate = norm(ws.title).replace(" ", "")
        if target in candidate or candidate in target:
            return ws
    return None


def _read_sheet_rows(ws):
    """Lee una hoja en modo read_only y devuelve solo los valores serializados."""
    if ws is None:
        return []
    rows = []
    for row in ws.iter_rows(values_only=True):
        vals = [serialize(v) for v in row]
        if any(v not in (None, "") for v in vals):
            rows.append(vals)
    return rows



# ============================================================
# BITÁCORA ESTRUCTURADA
# ============================================================

LOG_HEADERS = {
    "observation": 0,
    "report_date": 1,
    "pending": 2,
    "done": 3,
    "installation_date": 4,
    "hours": 5,
    "notes": 6,
}


def _is_checked(value):
    """Reconoce casillas/checks que puedan llegar como texto o booleano."""
    if value is True:
        return True
    if value in (1, 1.0):
        return True
    n = norm(value)
    return n in {"true", "verdadero", "si", "sí", "yes", "checked", "check", "✓", "✔", "☑"}


def _is_x(value):
    if value is None:
        return False
    return norm(value) in {"x", "✕", "✖", "×"}


def _log_hours_from_value(value):
    """Acepta 540, '540 h', etc., pero ignora números que parezcan fechas."""
    if value in (None, ""):
        return None
    try:
        x = float(value)
        if 1 <= x <= 10000:
            return int(x) if x.is_integer() else x
    except Exception:
        pass
    text = str(value).replace(",", ".")
    normalized = norm(text)
    # "kit 300 h" es un intervalo de servicio, no un horómetro actual.
    if re.search(r"\bkit\s+(?:de\s+)?300\s*(?:h|hr|hrs|hora|horas)?\b", normalized):
        return None
    m = re.search(r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hora|horas)\b", normalized)
    if m:
        x = float(m.group(1))
        return int(x) if x.is_integer() else x
    return None


def _log_text_hours(text):
    """Extrae horas explícitas escritas dentro de observaciones/notas."""
    text = norm(text)
    matches = re.finditer(r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hora|horas)\b", text)
    out = []
    for match in matches:
        # No tomar "kit 300 h" como horómetro.
        prefix = text[max(0, match.start()-12):match.start()]
        if re.search(r"\bkit\s+(?:de\s+)?$", prefix):
            continue
        x = float(match.group(1))
        if 1 <= x <= 10000:
            out.append(int(x) if x.is_integer() else x)
    # También reconoce el formato '(530) h de motor'.
    for m in re.findall(r"\((\d+(?:\.\d+)?)\)\s*h", text):
        x = float(m)
        if 1 <= x <= 10000:
            out.append(int(x) if x.is_integer() else x)
    return out


def _classify_log_system(text):
    """Clasifica la fila sin mezclar horas de motor con horas de generador."""
    p = norm(text)
    # Usamos palabras completas: "funcionando" no puede activar "onan" y
    # otras coincidencias accidentales no deben cambiar el sistema.
    has_generator = bool(re.search(r"\b(?:generador|cummins|onan)\b", p))
    has_motor = bool(re.search(r"\b(?:motor|motores|mercury|estribor|babor)\b", p))
    has_leg = bool(re.search(r"\b(?:pata|patas|outdrive|muelas)\b", p))

    # Si una misma fila habla explícitamente de motor y generador y solo hay
    # una cifra de horas, no adivinamos: esa hora NO se asigna a ninguno.
    if has_generator and has_motor:
        return "Ambos"
    if has_generator:
        return "Generador Cummins Onan"
    if has_leg:
        return "Patas Mercury"
    if has_motor:
        return "Motores Mercury"
    return "General"


def _log_status(pending, done, installation_date, text):
    """Regla de estado de la bitácora.

    C = X significa pendiente. D marcado significa resuelto/listo.
    La fecha de instalación/reparación en E confirma que el evento fue
    atendido, incluso cuando Excel entrega las casillas como objetos visuales
    sin valor de celda para openpyxl.
    """
    text_n = norm(text)
    if _is_checked(done) or installation_date is not None:
        return "RESUELTO"
    if _is_x(pending):
        return "PENDIENTE"
    if any(x in text_n for x in ("pendiente", "falta", "a la espera", "esperando")):
        return "PENDIENTE"
    return "REGISTRADO"


def parse_log_record(vals, row_number):
    observation = vals[LOG_HEADERS["observation"]] if len(vals) > 0 else None
    report_date = vals[LOG_HEADERS["report_date"]] if len(vals) > 1 else None
    pending = vals[LOG_HEADERS["pending"]] if len(vals) > 2 else None
    done = vals[LOG_HEADERS["done"]] if len(vals) > 3 else None
    installation_date = vals[LOG_HEADERS["installation_date"]] if len(vals) > 4 else None
    hours_value = vals[LOG_HEADERS["hours"]] if len(vals) > 5 else None
    notes = vals[LOG_HEADERS["notes"]] if len(vals) > 6 else None

    combined = " ".join(str(x) for x in (observation, notes) if x not in (None, ""))
    hours = _log_hours_from_value(hours_value)
    text_hours = _log_text_hours(combined)
    if hours is None and text_hours:
        hours = text_hours[-1]

    # La columna F puede contener una hora o, si algún día se escriben varias
    # en la misma celda, valores identificados por sistema. Esto permite
    # mantener separados motores y generador sin adivinar.
    hours_por_sistema = {}
    hours_text = norm(hours_value)
    labeled = re.findall(
        r"(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hora|horas)?\s*(?:de\s+)?"
        r"(motor(?:es)?|mercury|generador|cummins|onan|pata(?:s)?)\b",
        hours_text,
    )
    for raw_h, label in labeled:
        x = float(raw_h)
        if not (1 <= x <= 10000):
            continue
        label_n = norm(label)
        if label_n in {"generador", "cummins", "onan"}:
            key = "Generador Cummins Onan"
        elif label_n in {"pata", "patas"}:
            key = "Patas Mercury"
        else:
            key = "Motores Mercury"
        hours_por_sistema[key] = int(x) if x.is_integer() else x

    # Si la hora no viene etiquetada, se asigna únicamente cuando el texto de
    # la fila identifica un solo sistema. Si menciona dos sistemas, no se
    # adivina y esa cifra no contamina ningún horómetro.
    system = _classify_log_system(combined)
    if hours is not None and not hours_por_sistema and system in {
        "Motores Mercury", "Patas Mercury", "Generador Cummins Onan"
    }:
        hours_por_sistema[system] = hours

    status = _log_status(pending, done, installation_date, combined)

    return {
        "source_row": row_number,
        "values": vals,
        "data": {
            "observacion": observation,
            "fecha_reporte": report_date,
            "pendiente": pending,
            "listo": done,
            "fecha_instalacion": installation_date,
            "horas": hours,
            "horas_por_sistema": hours_por_sistema,
            "observaciones": notes,
            "sistema": system,
            "estado": status,
        },
    }


def log_records(data):
    return [r for r in data.get("log_records", []) if r.get("data")]


def _log_date(rec, key="fecha_reporte"):
    d = (rec.get("data") or {}).get(key)
    return parse_date(d) or excel_serial_to_date(d)


def _horometer_system(system):
    """Devuelve el sistema real del horómetro.

    Patas y Kit de 300 horas utilizan el mismo horómetro de motores.
    El generador mantiene su horómetro independiente.
    """
    if system in {"Motores Mercury", "Patas Mercury"}:
        return "Motores Mercury"
    return system


def latest_log_hours(data, system):
    """Obtiene el horómetro actual desde la Bitácora del Barco en Línea.

    La bitácora es la fuente de actualización del horómetro actual. El
    Checklist NO se usa para reemplazar ese valor; solo sirve como respaldo
    cuando todavía no existe una lectura en la bitácora.

    Motores y patas comparten el mismo horómetro. Generador mantiene el suyo.
    Cada fila se clasifica primero por el evento y después se asocia la hora
    de esa fila al sistema correspondiente.
    """
    target = _horometer_system(system)
    candidates = []

    for rec in log_records(data):
        d = rec.get("data") or {}
        horas_por_sistema = d.get("horas_por_sistema") or {}

        # Primero usamos una hora explícitamente etiquetada en la celda.
        # Esto permite interpretar celdas como "545 h motor / 2234 h generador"
        # sin mezclar ambos horómetros.
        if target in horas_por_sistema:
            h = horas_por_sistema[target]
        # Si la celda de horas no trae etiquetas, parse_log_record ya asignó
        # la lectura usando el texto del evento. Nunca sustituimos esta lectura
        # por el valor del Checklist cuando existe una lectura válida en Bitácora.
        # Patas se conecta deliberadamente al horómetro de motores.
        elif system == "Patas Mercury" and "Patas Mercury" in horas_por_sistema:
            h = horas_por_sistema["Patas Mercury"]
        elif d.get("sistema") == target:
            h = d.get("horas")
        elif system == "Patas Mercury" and d.get("sistema") == "Patas Mercury":
            h = d.get("horas")
        else:
            continue

        if h is None:
            continue
        rd = _log_date(rec) or date.min
        row = int(rec.get("source_row") or 0)
        candidates.append((rd, row, float(h)))

    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[-1][2]


def _format_log_record(rec):
    d = rec.get("data") or {}
    report = _log_date(rec)
    install = _log_date(rec, "fecha_instalacion")
    parts = []
    if report:
        parts.append(f"reporte {fmt_date(report)}")
    if d.get("estado"):
        parts.append(d["estado"])
    if install:
        parts.append(f"atendido/instalado {fmt_date(install)}")
    if d.get("horas") is not None:
        parts.append(f"{d['horas']:g} h")
    system = d.get("sistema")
    if system and system != "General":
        parts.append(system)
    text = str(d.get("observacion") or "").strip()
    notes = str(d.get("observaciones") or "").strip()
    if notes:
        text += f" — {notes}"
    return " — ".join([text] + parts) if text else " | ".join(parts)


def _extract_excel_images(path: Path):
    """Extrae fotografías incrustadas del Excel y las asocia a hoja/celda.

    Las imágenes no forman parte de los valores de las celdas. Se conservan
    como evidencia visual para que el cerebro pueda localizarlas y, cuando
    sea posible, aplicar OCR. No se inventa contenido visual.
    """
    import openpyxl
    images = []
    try:
        wb_img = openpyxl.load_workbook(path, data_only=False, read_only=False)
        for ws in wb_img.worksheets:
            for idx, img in enumerate(getattr(ws, "_images", []) or [], 1):
                try:
                    raw = img._data()
                    fmt = str(getattr(img, "format", None) or "png").lower()
                    if fmt == "jpg": fmt = "jpeg"
                    anchor = getattr(img, "anchor", None)
                    marker = getattr(anchor, "_from", None)
                    row = (getattr(marker, "row", 0) + 1) if marker is not None else None
                    col = (getattr(marker, "col", 0) + 1) if marker is not None else None
                    images.append({
                        "sheet": ws.title,
                        "index": idx,
                        "row": row,
                        "column": col,
                        "format": fmt,
                        "data_base64": base64.b64encode(raw).decode("ascii"),
                    })
                except Exception:
                    continue
        wb_img.close()
    except Exception:
        pass
    return images


def build_from_excel(path: Path):
    """Construye la base directamente desde Excel sin guardar una copia descomprimida.

    openpyxl abre el .xlsx como archivo ZIP y read_only evita cargar el libro
    completo en memoria de una sola vez. Esto permite trabajar mejor con
    archivos grandes y variables.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    data = {
        "metadata": {
            "agent": "Agente Tiara",
            "version": "2.2",
            "source_of_truth": "Excel cargado por el usuario",
            "source_file": path.name,
        },
        "sheets": {},
        "maintenance_records": [],
        "log_records": [],
        "inventory_records": [],
        "fuel_records": [],
        "budget_records": [],
        "invoice_records": [],
        "cleaning_records": [],
        "documents_records": [],
    }

    # Conservamos la estructura de hojas para que el resto de los agentes
    # siga funcionando, pero usamos lectura secuencial.
    for ws in wb.worksheets:
        rows = _read_sheet_rows(ws)
        data["sheets"][ws.title] = {
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "rows": rows,
        }

    # Mantenimiento
    ws = _find_sheet(wb, "CHECKLIST MANT. PREVENTIVO")
    if ws:
        for r, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = [serialize(v) for v in row]
            if any(v not in (None, "") for v in vals):
                data["maintenance_records"].append({
                    "source_row": r,
                    "data": {
                        "mantenimiento": vals[0] if len(vals) > 0 else None,
                        "cada_horas": vals[1] if len(vals) > 1 else None,
                        "historial_horas": vals[2] if len(vals) > 2 else None,
                        "fecha": vals[8] if len(vals) > 8 else None,
                        "observaciones": vals[9] if len(vals) > 9 else None,
                    },
                })

    mapping = {
        "log_records": "VITACORA BARCO EN LINEA",
        "inventory_records": "INVENTARIO",
        "fuel_records": "CONSUMO COMBUSTIBLE ",
        "budget_records": "PRESUPUESTO TIARA 2026",
        "invoice_records": "FACTURAS 2026...",
        "cleaning_records": "S.LIMPIEZA & OTROS",
        "documents_records": "PERMISOS & SEGUROS",
    }
    for key, sheet_name in mapping.items():
        ws = _find_sheet(wb, sheet_name)
        if not ws:
            continue
        for r, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = [serialize(v) for v in row]
            if not any(v not in (None, "") for v in vals):
                continue

            if key == "log_records":
                # La bitácora tiene una estructura propia que el agente debe
                # entender, no tratar como una fila de texto genérica.
                data[key].append(parse_log_record(vals, r))
            else:
                data[key].append({"source_row": r, "values": vals})

    wb.close()
    # Segunda pasada únicamente para fotografías incrustadas; read_only no las expone.
    # Se analizan con visión y luego se eliminan los bytes para que Supabase no
    # reciba el contenido binario de todas las fotos.
    raw_images = _extract_excel_images(path)
    data["images"], vision_meta = _analyze_embedded_images(raw_images)
    data["metadata"]["vision"] = vision_meta
    return data


def build_from_uploaded_excel(uploaded_file):
    """Procesa un Excel subido desde la aplicación sin descomprimirlo manualmente.

    El formato .xlsx/.xlsm ya es un contenedor ZIP. openpyxl lo abre
    directamente y read_only permite leerlo por partes. Primero copiamos el
    archivo a un temporal por bloques y comprobamos que el contenedor sea
    válido; después construimos la misma base que usa el agente.
    """
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix not in (".xlsx", ".xlsm"):
        raise ValueError("El archivo debe ser Excel .xlsx o .xlsm")

    # Límite de seguridad razonable para evitar cargar accidentalmente un
    # archivo enorme. El límite de Streamlit también puede existir en el
    # servidor, pero este control ocurre dentro de la aplicación.
    max_bytes = 500 * 1024 * 1024
    declared_size = getattr(uploaded_file, "size", None)
    if declared_size and declared_size > max_bytes:
        raise ValueError("El Excel supera el límite de 500 MB de esta aplicación.")

    h = hashlib.sha256()

    with tempfile.NamedTemporaryFile(
        prefix="tiara_excel_",
        suffix=suffix,
        delete=False,
    ) as tmp:
        temp_path = Path(tmp.name)

        while True:
            chunk = uploaded_file.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
            tmp.write(chunk)

    uploaded_file.seek(0)

    try:
        # Un XLSX/XLSM válido es un ZIP. No extraemos el contenido a una
        # carpeta: openpyxl lo lee directamente desde el contenedor.
        if not zipfile.is_zipfile(temp_path):
            raise ValueError(
                "El archivo no parece ser un Excel .xlsx/.xlsm válido. "
                "Vuelve a cargar el archivo original."
            )

        data = build_from_excel(temp_path)

        # Metadatos de la carga actual.
        data["metadata"]["upload_sha256"] = h.hexdigest()
        data["metadata"]["version"] = APP_VERSION
        data["metadata"]["source_file"] = uploaded_file.name
        data["metadata"]["source_size_mb"] = round(
            temp_path.stat().st_size / (1024 * 1024), 2
        )
        data["metadata"]["source_updated_in_app"] = datetime.now().isoformat(
            timespec="seconds"
        )

        # Verificación mínima: el libro debe tener hojas.
        if not data.get("sheets"):
            raise ValueError("El Excel no contiene hojas de datos legibles.")

        return data

    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass

def serialize(v):
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return v


def sheet_rows(data, name):
    return data.get("sheets", {}).get(name, {}).get("rows", [])


def maintenance_records(data):
    return [r for r in data.get("maintenance_records", []) if r.get("data", {}).get("mantenimiento")]


# ============================================================
# CEREBRO / PLANIFICADOR
# ============================================================

def detect_system(p):
    p = norm(p)
    if any(x in p for x in ("generador", "cummins", "onan")): return "Generador Cummins Onan"
    if any(x in p for x in ("pata", "patas", "outdrive", "muelas")): return "Patas Mercury"
    if any(x in p for x in ("motor", "motores", "mercury", "estribor", "babor")): return "Motores Mercury"
    if any(x in p for x in ("combustible", "gasolina", "diesel", "diésel")): return "Combustible"
    if any(x in p for x in ("inventario", "existencias", "stock", "producto", "vasos", "fender", "chaleco")): return "Inventario"
    if any(x in p for x in ("factura", "facturas", "invoice")): return "Facturas"
    if any(x in p for x in ("permiso", "permisos", "seguro", "seguros", "garantia", "garantía", "documento")): return "Permisos y seguros"
    if any(x in p for x in ("bitacora", "bitácora", "evento", "problema", "falla", "alarma", "reparacion", "reparación")): return "Bitácora"
    if any(x in p for x in ("limpieza", "limpiar", "aseo")): return "Limpieza"
    if any(x in p for x in ("presupuesto", "gasto", "gastamos", "costo", "costos", "gastado", "gastó")): return "Finanzas"
    return "General"


def detect_intent(p):
    p = norm(p)
    if any(x in p for x in ("estado general", "estado del barco", "estado completo", "revision general", "revisión general")): return "estado_general"

    # Finanzas: primero identificar un rubro, pero separar INVENTARIO operativo
    # de INVENTARIO.COCINA del presupuesto.
    finance_words = ("cuanto", "gasto", "gaste", "gastamos", "costo", "presupuesto", "presupuestado", "gastado", "gastó", "cuanto se gasto")
    rubric = identify_rubric(p)
    operational_inventory = any(x in p for x in ("que hay en inventario", "qué hay en inventario", "que tenemos en inventario", "existencias", "stock"))
    if rubric and any(x in p for x in finance_words) and not operational_inventory:
        return "gasto_mensual" if (find_month(p) or any(x in p for x in ("ese mes", "mismo mes"))) else "gasto_anual"
    if any(x in p for x in ("cuanto gastamos", "cuánto gastamos", "cuanto gasto", "cuánto gasto", "gasto en", "gastamos en", "presupuesto", "gastado")):
        return "gasto_mensual" if find_month(p) else "gasto_anual"

    if any(x in p for x in ("cuantas veces", "cuántas veces", "cuantos cambios", "cuántos cambios", "cantidad de cambios", "cantidad de mantenimientos")): return "cantidad_mantenimiento"
    if any(x in p for x in ("historial completo", "todo el historial", "todas las veces", "todos los cambios", "todos los mantenimientos", "muestrame el historial", "muéstrame el historial")): return "historial_completo"
    if any(x in p for x in ("ultimo cambio", "último cambio", "ultimo mantenimiento", "último mantenimiento", "cuando fue", "cuándo fue", "cuando se hizo", "cuándo se hizo", "ultima vez", "última vez")): return "ultimo_mantenimiento"
    if any(x in p for x in ("kit 300", "kit de 300", "300 horas", "servicio 300")): return "kit_300"
    if any(x in p for x in ("cuando toca", "cuándo toca", "proximo", "próximo", "faltan horas", "que mantenimiento", "qué mantenimiento", "mantenimiento")): return "mantenimiento"
    if operational_inventory or any(x in p for x in ("inventario", "existencias", "stock")): return "inventario"
    if any(x in p for x in ("combustible", "gasolina", "diesel", "diésel")): return "combustible"
    if any(x in p for x in ("factura", "facturas", "invoice")): return "facturas"
    if any(x in p for x in ("foto", "fotografia", "fotografía", "imagen", "documento fotografiado", "que dice la foto", "qué dice la foto", "lee la foto", "leer la foto", "mira la foto", "que muestra la foto", "qué muestra la foto")): return "fotografias"
    if any(x in p for x in ("permiso", "permisos", "seguro", "seguros", "vence", "vencimiento", "documentos")): return "documentos"
    if any(x in p for x in ("limpieza", "limpiar", "aseo")): return "limpieza"
    if any(x in p for x in ("bitacora", "bitácora", "evento", "falla", "alarma", "problema")): return "bitacora"
    if any(x in p for x in ("horas", "horometro", "horómetro")): return "horometros"
    return "general"


def find_month(p):
    for name, number in MONTHS.items():
        if name in p:
            return number
    return None


def find_year(p):
    years = [int(x) for x in re.findall(r"\b(20\d{2})\b", p)]
    return years[0] if years else None


def think(question):
    p = norm(question)
    return {
        "intent": detect_intent(p),
        "system": detect_system(p),
        "month": find_month(p),
        "year": find_year(p),
        "normalized": p,
    }


# ============================================================
# MÓDULO DE MANTENIMIENTO
# ============================================================

SYSTEM_ALIASES = {
    "Motores Mercury": ["motor", "motores", "mercury"],
    "Patas Mercury": ["pata", "patas"],
    "Generador Cummins Onan": ["generador", "cummins", "onan"],
}


def system_matches(name, system):
    n = norm(name)
    return any(a in n for a in SYSTEM_ALIASES.get(system, []))


def maintenance_candidates(data, system, p):
    records = maintenance_records(data)
    exact = []
    broad = []
    for r in records:
        d = r["data"]
        name = norm(d.get("mantenimiento"))
        row = int(r.get("source_row", 0) or 0)

        if system == "Motores Mercury":
            # Bloque de motores: filas 4-9 del checklist.
            if 4 <= row <= 9:
                broad.append(r)
                if "aceite motores" in name or "filtros de aceite" in name:
                    exact.append(r)

        elif system == "Patas Mercury":
            # Las patas aparecen dentro del bloque de motores, pero se
            # identifican por su nombre específico.
            if "pata" in name:
                broad.append(r)
                if "aceite de pata" in name:
                    exact.append(r)

        elif system == "Generador Cummins Onan":
            # Bloque del generador: filas 11-31 del checklist.
            if 11 <= row <= 31:
                broad.append(r)
                if "cambio de aceite" in name or "cambio filtro de aceite" in name:
                    exact.append(r)

    return exact, broad


def maintenance_row_summary(r):
    d = r["data"]
    hours = extract_hours(d.get("historial_horas"))
    dates = extract_dates(d.get("fecha"))
    interval = safe_float(d.get("cada_horas"))
    return {
        "name": str(d.get("mantenimiento") or "").strip(),
        "interval": int(interval) if interval is not None and interval.is_integer() else interval,
        "hours": hours,
        "dates": dates,
        "last_hour": hours[-1] if hours else None,
        "last_date": dates[-1] if dates else None,
        "next_hour": (hours[-1] + int(interval)) if hours and interval is not None else None,
        "observations": str(d.get("observaciones") or "").strip(),
    }


def answer_ultimo_mantenimiento(data, thought):
    system = thought["system"]
    p = thought["normalized"]
    if system not in SYSTEM_ALIASES:
        return "No pude determinar si preguntas por motores, patas o generador."

    exact, broad = maintenance_candidates(data, system, p)
    wants_oil = "aceite" in p or "filtro" in p
    candidates = exact if wants_oil and exact else broad

    # Si es una pregunta de motores sin detalle, devolver SOLO el mantenimiento principal,
    # no todas las filas relacionadas.
    if not wants_oil and system == "Motores Mercury":
        candidates = [r for r in broad if "aceite motores" in norm(r["data"].get("mantenimiento"))]
    elif not wants_oil and system == "Patas Mercury":
        # Para una pregunta genérica por las patas, usar el mantenimiento
        # principal (aceite), no escoger arbitrariamente un ánodo u otro ítem.
        candidates = [r for r in broad if "aceite de pata" in norm(r["data"].get("mantenimiento"))]

    if not candidates:
        return "No encontré un registro de mantenimiento confirmado para ese sistema."

    # Escoger la fila más reciente por fecha; en empate, por horómetro.
    summaries = [maintenance_row_summary(r) for r in candidates]
    summaries.sort(key=lambda x: (x["last_date"] or date.min, x["last_hour"] or -1))
    s = summaries[-1]

    lines = [
        f"### Último mantenimiento — {system}",
        f"**Mantenimiento:** {s['name']}",
        f"**Última fecha registrada:** {fmt_date(s['last_date'])}",
        f"**Último horómetro registrado:** {s['last_hour']} h" if s["last_hour"] is not None else "**Último horómetro registrado:** NO DETERMINADO",
    ]
    if s["interval"] is not None and s["last_hour"] is not None:
        lines.append(f"**Próximo por horas:** {s['next_hour']} h")
    return "\n".join(lines)


def answer_historial(data, thought, complete=False):
    system = thought["system"]
    if system not in SYSTEM_ALIASES:
        return "No pude determinar el sistema para consultar el historial."
    exact, broad = maintenance_candidates(data, system, thought["normalized"])
    rows = broad if complete else exact
    if not rows:
        rows = broad
    if not rows:
        return "No encontré historial de mantenimiento para ese sistema."

    # Historial completo de un sistema: máximo los registros que realmente pertenecen al sistema.
    items = []
    for r in rows:
        s = maintenance_row_summary(r)
        if s["last_date"] or s["last_hour"]:
            items.append(s)
    items.sort(key=lambda x: (x["last_date"] or date.min, x["last_hour"] or -1), reverse=True)

    lines = [f"### Historial de mantenimiento — {system}"]
    for s in items:
        if s["last_date"]:
            lines.append(f"- **{s['name']}** — última fecha: {fmt_date(s['last_date'])}" + (f", {s['last_hour']} h" if s["last_hour"] is not None else ""))
        else:
            lines.append(f"- **{s['name']}** — última hora: {s['last_hour']} h")
    return "\n".join(lines)


def answer_count(data, thought):
    system = thought["system"]
    exact, broad = maintenance_candidates(data, system, thought["normalized"])
    rows = exact if exact else broad
    if not rows:
        return "No encontré registros suficientes para contar ese mantenimiento."
    total = 0
    details = []
    for r in rows:
        s = maintenance_row_summary(r)
        n = len(s["hours"]) if s["hours"] else len(s["dates"])
        if n:
            total += n
            details.append(f"- {s['name']}: {n} registros")
    if not details:
        return "No encontré registros fechados o por horas para ese mantenimiento."
    return f"### Cantidad de mantenimientos — {system}\n**Total:** {total}\n" + "\n".join(details)


def answer_mantenimiento(data, thought):
    system = thought["system"]
    p = thought["normalized"]
    if system not in SYSTEM_ALIASES:
        return "No pude determinar el sistema del barco."

    if system == "Motores Mercury":
        rows = [r for r in maintenance_records(data) if "aceite motores" in norm(r["data"].get("mantenimiento"))]
    elif system == "Patas Mercury":
        rows = [r for r in maintenance_records(data) if "aceite de pata" in norm(r["data"].get("mantenimiento"))]
    else:
        rows = [r for r in maintenance_records(data) if "cambio de aceite" in norm(r["data"].get("mantenimiento")) and 11 <= int(r.get("source_row", 0)) <= 31]

    if not rows:
        return "No encontré el mantenimiento programado para ese sistema."

    # Si existen varias filas del mismo sistema, tomar la que tenga el
    # registro más reciente, no simplemente la primera fila del Excel.
    summaries = [maintenance_row_summary(r) for r in rows]
    s = max(summaries, key=lambda x: (x["last_date"] or date.min, x["last_hour"] or -1))
    current = current_hours(data).get(system)
    lines = [
        f"### 🔧 Próximo mantenimiento — {system}",
        f"**Mantenimiento:** {s['name']}",
        f"**Último registrado:** {fmt_date(s['last_date'])}",
        f"**Horómetro del último registro:** {s['last_hour']} h" if s["last_hour"] is not None else "**Horómetro del último registro:** NO DETERMINADO",
        f"**Intervalo:** {s['interval']} h" if s["interval"] is not None else "**Intervalo:** NO DETERMINADO",
        f"**Próximo:** {s['next_hour']} h" if s["next_hour"] is not None else "**Próximo:** NO DETERMINADO",
    ]
    if current is not None and s["next_hour"] is not None:
        faltan = s["next_hour"] - current
        lines.append(f"**Horómetro actual:** {current} h")
        lines.append(f"**Faltan:** {faltan} h" if faltan >= 0 else f"**Excedido por:** {abs(faltan)} h")
    return "\n".join(lines)


def answer_kit(data):
    rows = [r for r in maintenance_records(data) if "kit 300 horas" in norm(r["data"].get("mantenimiento"))]
    if not rows:
        return "No encontré el Kit de 300 horas en el Excel."

    summaries = [maintenance_row_summary(r) for r in rows]
    s = max(summaries, key=lambda x: (x["last_date"] or date.min, x["last_hour"] or -1))
    current = current_hours(data).get("Motores Mercury")
    lines = [
        "### 🔩 Kit 300 horas — Motores Mercury",
        f"**Último kit registrado:** {s['last_hour']} h" if s["last_hour"] is not None else "**Último kit registrado:** NO DETERMINADO",
        f"**Fecha del último kit:** {fmt_date(s['last_date'])}",
        f"**Intervalo:** {s['interval']} h" if s["interval"] is not None else "**Intervalo:** NO DETERMINADO",
        f"**Próximo kit:** {s['next_hour']} h" if s["next_hour"] is not None else "**Próximo kit:** NO DETERMINADO",
    ]
    if current is not None and s["next_hour"] is not None:
        faltan = s["next_hour"] - current
        lines.append(f"**Horómetro actual de motores:** {current} h")
        lines.append(f"**Faltan:** {faltan} h" if faltan >= 0 else f"**Excedido por:** {abs(faltan)} h")
    return "\n".join(lines)


def current_hours(data):
    result = {}
    # La bitácora es la fuente más reciente cuando registra un evento con
    # horas. El checklist queda como respaldo histórico.
    fallback_rows = maintenance_records(data)
    targets = {
        "Motores Mercury": 4,
        "Generador Cummins Onan": 11,
    }
    for system, source_row in targets.items():
        log_h = latest_log_hours(data, system)
        if log_h is not None:
            result[system] = int(log_h) if float(log_h).is_integer() else log_h
            continue

        match = next((r for r in fallback_rows if int(r.get("source_row", -1)) == source_row), None)
        if match:
            hours = extract_hours(match["data"].get("historial_horas"))
            result[system] = hours[-1] if hours else None
        else:
            result[system] = None

    # Patas utilizan deliberadamente el mismo horómetro de motores.
    result["Patas Mercury"] = result.get("Motores Mercury")
    return result


def answer_state(data):
    hours = current_hours(data)
    lines = ["### Estado general del barco"]

    target_names = {
        "Motores Mercury": {"cambio aceite motores"},
        "Patas Mercury": {"cambio de aceite de pata motores"},
        "Generador Cummins Onan": {"cambio de aceite"},
    }

    for system in ("Motores Mercury", "Patas Mercury", "Generador Cummins Onan"):
        h = hours.get(system)
        lines.append(
            f"**{system}:** {h} h" if h is not None
            else f"**{system}:** NO DETERMINADO"
        )

        rows = [
            r for r in maintenance_records(data)
            if norm(r["data"].get("mantenimiento")) in target_names[system]
        ]
        if not rows:
            continue

        summaries = [maintenance_row_summary(r) for r in rows]
        s = max(summaries, key=lambda x: (x["last_date"] or date.min, x["last_hour"] or -1))

        if s["last_date"]:
            lines.append(f"  - Último mantenimiento: {fmt_date(s['last_date'])}")
        if s["last_hour"] is not None:
            lines.append(f"  - Horómetro del último mantenimiento: {s['last_hour']} h")
        if s["interval"] is not None:
            lines.append(f"  - Intervalo: {s['interval']} h")
        if s["next_hour"] is not None:
            lines.append(f"  - Próximo: {s['next_hour']} h")
            if h is not None:
                faltan = s["next_hour"] - h
                lines.append(
                    f"  - Faltan: {faltan} h" if faltan >= 0
                    else f"  - Excedido por: {abs(faltan)} h"
                )

    return "\n".join(lines)


# ============================================================
# FINANZAS / COMBUSTIBLE
# ============================================================

# Alias de todos los rubros financieros presentes en la base.
# La consulta puede usar el nombre exacto del Excel o una forma coloquial.
# La lógica devuelve siempre el nombre canónico que usa el presupuesto.
RUBRIC_RULES = [
    (("mantenimiento", "anual"), "MANTENIMIENTO ANUAL TIARA"),
    (("mantenimiento", "motor"), "MANT.MOTORES"),
    (("mantenimiento", "motores"), "MANT.MOTORES"),
    (("motores",), "MANT.MOTORES"),
    (("mantenimiento", "generador"), "MANT.GENERADOR"),
    (("generador",), "MANT.GENERADOR"),
    (("combustible",), "COMSUMO COMBUSTIBLE"),
    (("gasolina",), "COMSUMO COMBUSTIBLE"),
    (("diesel",), "COMSUMO COMBUSTIBLE"),
    (("diésel",), "COMSUMO COMBUSTIBLE"),
    (("repuestos",), "REPUESTOS & COTIZACIONES"),
    (("repuesto",), "REPUESTOS & COTIZACIONES"),
    (("cotizaciones",), "REPUESTOS & COTIZACIONES"),
    (("cotizacion",), "REPUESTOS & COTIZACIONES"),
    (("seguros",), "SEGUROS & MEMBRESIAS"),
    (("seguro",), "SEGUROS & MEMBRESIAS"),
    (("membresias",), "SEGUROS & MEMBRESIAS"),
    (("membresía",), "SEGUROS & MEMBRESIAS"),
    (("salario",), "SALARIO ANDRES"),
    (("sueldo",), "SALARIO ANDRES"),
    (("limpieza",), "S.LIMPIEZA"),
    (("inventario",), "INVENTARIO.COCINA"),
    (("cocina",), "INVENTARIO.COCINA"),
    (("s.cocina",), "INVENTARIO.COCINA"),
    (("agua",), "PAGO POR AGUA DULCE HIELO ETC."),
    (("hielo",), "PAGO POR AGUA DULCE HIELO ETC."),
    (("agua dulce",), "PAGO POR AGUA DULCE HIELO ETC."),
    (("visitas",), "VISITAS JUAN MANUEL"),
    (("juan manuel",), "VISITAS JUAN MANUEL"),
    (("marina",), "TRAVEL/SLIP MARINA"),
    (("slip",), "TRAVEL/SLIP MARINA"),
    (("travel",), "TRAVEL/SLIP MARINA"),
    (("comida",), "COMIDA ANDRES"),
    (("ccss",), "CCSS ANDRES"),
    (("otros",), "OTROS/"),
]


def identify_rubric(p):
    """Reconoce rubros financieros por lenguaje natural y por aliases del Excel."""
    p = norm(p)
    if any(x in p for x in ("que hay en inventario", "qué hay en inventario", "existencias", "stock")) and not any(x in p for x in ("gasto", "gastamos", "costo", "presupuesto", "gastado")):
        return None
    aliases = [
        (("mantenimiento anual", "mantenimiento general anual", "mant anual"), "MANTENIMIENTO ANUAL TIARA"),
        (("mantenimiento de motores", "mantenimiento motores", "mant motores", "motor mercury"), "MANT.MOTORES"),
        (("mantenimiento del generador", "mantenimiento generador", "mant generador"), "MANT.GENERADOR"),
        (("combustible", "gasolina", "diesel", "diésel", "consumo de combustible"), "COMSUMO COMBUSTIBLE"),
        (("repuestos", "repuesto", "cotizaciones", "cotizacion"), "REPUESTOS & COTIZACIONES"),
        (("seguros", "seguro", "membresias", "membresía", "membresias y seguros"), "SEGUROS & MEMBRESIAS"),
        (("salario", "sueldo"), "SALARIO ANDRES"),
        (("limpieza", "aseo"), "S.LIMPIEZA"),
        (("inventario cocina", "inventario de cocina", "s cocina", "s.cocina"), "INVENTARIO.COCINA"),
        (("agua dulce", "hielo", "agua e hielo", "agua"), "PAGO POR AGUA DULCE HIELO ETC."),
        (("visitas", "juan manuel"), "VISITAS JUAN MANUEL"),
        (("marina", "slip", "travel slip", "travel"), "TRAVEL/SLIP MARINA"),
        (("comida", "comidas", "alimentacion", "alimentación"), "COMIDA ANDRES"),
        (("ccss",), "CCSS ANDRES"),
        (("otros", "otros gastos"), "OTROS/"),
    ]
    for variants, rubric in aliases:
        if any(v in p for v in variants):
            return rubric
    ordered = sorted(RUBRIC_RULES, key=lambda item: len(item[0]), reverse=True)
    for words, rubric in ordered:
        if all(norm(w) in p for w in words): return rubric
    return None


def _budget_name(rec):
    data = rec.get("data") or {}
    if data:
        for key in ("RUBROS", "rubros", "RUBRO", "rubro"):
            if data.get(key) not in (None, ""):
                return str(data[key]).strip()
    vals = rec.get("values") or []
    if len(vals) > 2 and vals[2] not in (None, ""):
        return str(vals[2]).strip()
    return ""


def _budget_number(value):
    if value in (None, "", "-", "?"):
        return None
    return safe_float(value)


def _budget_month_key(year, month):
    return f"{MONTH_NAMES[month][:3]}-{str(year)[2:]}"


def _budget_dict_value(rec, year, month):
    """Lee un mes desde la estructura JSON real del presupuesto."""
    data = rec.get("data") or {}
    if not data:
        return None

    target = _budget_month_key(year, month)
    for key, value in data.items():
        nk = norm(key).replace(" ", "")
        # Acepta ene-26, ene26, feb-26 y también ' feb 26'.
        if nk.replace("-", "") == target.replace("-", ""):
            return _budget_number(value)
    return None


def _budget_rows_from_excel_values(data):
    """Convierte filas antiguas del Excel en registros con encabezados reales."""
    rows = budget_rows(data)
    headers = None
    raw = sheet_rows(data, "PRESUPUESTO TIARA 2026")
    if raw:
        # En el Excel del proyecto, la fila 3 (índice 2) contiene los encabezados.
        if len(raw) > 2:
            headers = [norm(x).replace(" ", "") if x is not None else "" for x in raw[2]]

    out = []
    for rec in rows:
        if rec.get("data"):
            out.append(rec)
            continue
        vals = rec.get("values") or []
        mapped = {}
        if headers:
            for i, value in enumerate(vals):
                if i < len(headers) and headers[i]:
                    mapped[headers[i]] = value
        clone = dict(rec)
        clone["data"] = mapped
        out.append(clone)
    return out


def budget_rows(data):
    rows = data.get("budget_records", [])
    if rows:
        return rows
    raw = sheet_rows(data, "PRESUPUESTO TIARA 2026")
    return [{"source_row": i + 1, "values": row} for i, row in enumerate(raw)]


def _is_budget_total_row(name, rec=None):
    n = norm(name)
    if n in ("total", "totales") or n.startswith("total "):
        return True
    # En la estructura real del Excel, la fila TOTAL usa "Total" en
    # OBSERVACIONES y deja RUBROS vacío.
    d = (rec or {}).get("data") or {}
    for key in ("OSERVACIONES", "OBSERVACIONES", "observaciones"):
        if norm(d.get(key)) in ("total", "totales"):
            return True
    return False


def _budget_name_matches(name, rubric):
    """Compara nombres del Excel tolerando tildes, espacios y variantes."""
    a = norm(name)
    b = norm(rubric)
    variants = {
        "comsumo combustible": {"comsumo combustible", "consumo combustible"},
        "repuestos & cotizaciones": {"repuestos & cotizaciones", "repuestos &cotizaciones"},
        "s.limpieza": {"s.limpieza", "s limpieza"},
        "seguros & membresias": {"seguros & membresias", "seguros/ membresias", "seguros/membresias"},
        "pago por agua dulce hielo etc.": {
            "pago por agua dulce hielo etc.",
            "pago por agua dulce,hielo,..",
        },
    }
    if a == b:
        return True
    return a in variants.get(b, set()) or b in variants.get(a, set())


def _find_budget_record(data, rubric):
    target = norm(rubric)
    for rec in _budget_rows_from_excel_values(data):
        if _budget_name_matches(_budget_name(rec), rubric):
            return rec
    return None


def _find_budget_record_for_month(data, rubric, year, month):
    """Busca el registro del rubro que realmente contiene el mes solicitado.

    El Excel puede contener más de una fila con el mismo rubro (por ejemplo,
    una fila de referencia/presupuesto y otra con los importes mensuales).
    Para una pregunta mensual no debemos quedarnos con la primera fila si esa
    fila no tiene el mes solicitado.
    """
    target = norm(rubric)
    candidates = []
    for rec in _budget_rows_from_excel_values(data):
        if not _budget_name_matches(_budget_name(rec), rubric):
            continue
        value = _budget_dict_value(rec, year, month)
        if value is not None:
            candidates.append((rec, value))
    return candidates[0] if candidates else (None, None)


def budget_value_for_month(data, rubric, year, month):
    rec = _find_budget_record(data, rubric)
    if not rec:
        return None

    value = _budget_dict_value(rec, year, month)
    if value is not None:
        return value

    # Compatibilidad con el formato antiguo de filas del Excel.
    vals = rec.get("values", [])
    if year == 2026:
        col = 3 + (month - 1) * 2
    elif year == 2025:
        col = 4 + (month - 1) * 2
    elif year == 2024:
        col = 3 + (month - 1) * 2
    else:
        return None
    return _budget_number(vals[col]) if col < len(vals) else None


def budget_total(data, rubric, year):
    rec = _find_budget_record(data, rubric)
    if not rec:
        return None
    d = rec.get("data") or {}
    for key in (f"TOTAL  {year}", f"TOTAL {year}"):
        if key in d:
            return _budget_number(d[key])
    # Buscar de forma tolerante.
    for key, value in d.items():
        if norm(key).replace(" ", "") == f"total{year}":
            return _budget_number(value)

    vals = rec.get("values", [])
    if year == 2026 and len(vals) > 27:
        return _budget_number(vals[27])
    if year == 2025 and len(vals) > 28:
        return _budget_number(vals[28])
    return None


def budget_total_all(data, year, month=None):
    """Total general del presupuesto. Usa la fila TOTAL del Excel cuando existe."""
    rows = _budget_rows_from_excel_values(data)

    # 1) La fuente tiene una fila explícita TOTAL. Es la autoridad para una
    #    pregunta como 'cuánto se gastó en agosto de 2026'.
    for rec in rows:
        if _is_budget_total_row(_budget_name(rec), rec):
            if month is not None:
                value = _budget_dict_value(rec, year, month)
                if value is not None:
                    return value
            else:
                d = rec.get("data") or {}
                for key, value in d.items():
                    if norm(key).replace(" ", "") == f"total{year}":
                        return _budget_number(value)

    # 2) Compatibilidad: sumar los rubros si la fila TOTAL no está disponible.
    total = 0.0
    found = False
    for rec in rows:
        name = _budget_name(rec)
        if not name or _is_budget_total_row(name, rec):
            continue
        value = _budget_dict_value(rec, year, month) if month else budget_total(data, name, year)
        if value is not None:
            total += value
            found = True
    return total if found else None


def answer_expenses(data, thought):
    p = thought["normalized"]
    rubric = identify_rubric(p)
    year = thought["year"] or 2026
    month = thought["month"]

    # Pregunta mensual sin rubro: responder con el TOTAL GENERAL del mes.
    # Ejemplo: '¿Cuánto se gastó en agosto del 2026?'
    if month and not rubric:
        value = budget_total_all(data, year, month)
        if value is None:
            return (f"### Gasto mensual\n**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                    "**Gasto registrado:** NO DETERMINADO")
        return (f"### Gasto mensual — Presupuesto Tiara\n"
                f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                f"**Gasto total registrado:** {money(value)}\n"
                "**Fuente:** PRESUPUESTO TIARA 2026\n"
                "**Estado:** CONFIRMADO")

    # Pregunta anual sin rubro: responder con el total general anual.
    if not month and not rubric:
        value = budget_total_all(data, year)
        if value is None:
            return f"### Gasto anual\n**Año:** {year}\n**Gasto total:** NO DETERMINADO"
        return (f"### Gasto anual — Presupuesto Tiara\n**Año:** {year}\n"
                f"**Gasto total registrado:** {money(value)}\n"
                "**Fuente:** PRESUPUESTO TIARA 2026\n**Estado:** CONFIRMADO")

    if not rubric:
        return "No pude determinar el rubro del gasto. Puedes indicar el concepto (combustible, motores, generador, limpieza, etc.)."

    if month:
        value = budget_value_for_month(data, rubric, year, month)
        if value is None:
            return (f"### Gasto mensual\n**Rubro:** {rubric}\n"
                    f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                    "**Gasto registrado:** NO DETERMINADO")
        return (f"### Gasto mensual\n**Rubro:** {rubric}\n"
                f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                f"**Gasto registrado:** {money(value)}\n"
                "**Fuente:** PRESUPUESTO TIARA 2026\n**Estado:** CONFIRMADO")

    value = budget_total(data, rubric, year)
    if value is None:
        return f"### Gasto anual\n**Rubro:** {rubric}\n**Año:** {year}\n**Gasto:** NO DETERMINADO"
    return (f"### Gasto anual\n**Rubro:** {rubric}\n**Año:** {year}\n"
            f"**Gasto registrado:** {money(value)}\n"
            "**Fuente:** PRESUPUESTO TIARA 2026\n**Estado:** CONFIRMADO")

def answer_fuel(data, thought):
    year = thought["year"] or 2026
    month = thought.get("month")

    # Cuando la consulta es mensual y viene dentro de una conversación de
    # gastos (por ejemplo: "En ese mismo mes cuanto en combustible"),
    # consultar primero el rubro financiero de combustible del PRESUPUESTO.
    # Esto evita devolver los 12 meses cuando el usuario pidió un solo mes.
    if month:
        # El Excel contiene el rubro escrito de dos formas:
        # "COMSUMO COMBUSTIBLE" y "CONSUMO COMBUSTIBLE". Buscar ambas
        # variantes para no perder el mes solicitado.
        rec = None
        value = None
        for rubric in ("COMSUMO COMBUSTIBLE", "CONSUMO COMBUSTIBLE"):
            rec, value = _find_budget_record_for_month(data, rubric, year, month)
            if rec is not None and value is not None:
                break
        if rec is not None and value is not None:
            return (
                f"### Gasto de combustible\n"
                f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                f"**Combustible registrado:** {money(value)}\n"
                f"**Fuente:** PRESUPUESTO TIARA 2026\n"
                f"**Estado:** CONFIRMADO"
            )

    # Si no hay mes, o el presupuesto no contiene el mes solicitado, usar la
    # bitácora de consumo de combustible como respaldo.
    rows = data.get("fuel_records", [])
    totals = {m: 0.0 for m in range(1, 13)}
    found = {m: False for m in range(1, 13)}

    for rec in rows:
        vals = rec.get("values", [])
        if not vals:
            continue
        # En la base normalizada de combustible, MES es la primera columna y
        # contiene el mes como fecha serial de Excel.
        d = parse_date(vals[0]) or excel_serial_to_date(vals[0])
        if not d or d.year != year:
            continue
        total = safe_float(vals[5]) if len(vals) > 5 else None
        if total is not None:
            totals[d.month] += total
            found[d.month] = True

    if month:
        if found[month]:
            return (
                f"### Consumo de combustible\n"
                f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
                f"**Combustible registrado:** {money(totals[month])}\n"
                f"**Fuente:** CONSUMO COMBUSTIBLE\n"
                f"**Estado:** CONFIRMADO"
            )
        return (
            f"### Consumo de combustible\n"
            f"**Mes:** {MONTH_NAMES[month].capitalize()} {year}\n"
            f"**Combustible registrado:** NO DETERMINADO\n"
            f"**Fuente:** CONSUMO COMBUSTIBLE / PRESUPUESTO TIARA 2026"
        )

    annual = sum(totals.values())
    lines = [f"### Consumo de combustible — {year}"]
    for m in range(1, 13):
        lines.append(
            f"- **{MONTH_NAMES[m].capitalize()}:** "
            f"{money(totals[m]) if found[m] else 'NO DETERMINADO'}"
        )
    lines.append(f"\n**Total del año:** {money(annual)}")
    return "\n".join(lines)


# ============================================================
# INVENTARIO / BITÁCORA / DOCUMENTOS
# ============================================================

def answer_inventory(data, thought):
    p = thought["normalized"]
    rows = data.get("inventory_records", [])
    if not rows:
        return "No hay datos de inventario disponibles."
    tokens = [t for t in re.findall(r"[a-z0-9]+", p) if len(t) >= 4 and t not in STOPWORDS]
    hits = []
    for rec in rows:
        vals = rec.get("values", [])
        if len(vals) < 3:
            continue
        product = norm(vals[1])
        score = sum(1 for t in tokens if t in product)
        if score:
            hits.append((score, product, vals))
    hits.sort(reverse=True)
    if not hits:
        return "No encontré ese producto en el inventario."
    lines = ["### Inventario"]
    for _, _, vals in hits[:8]:
        lines.append(f"- **{vals[1]}** — existencias: **{vals[2]}**")
    return "\n".join(lines)


def answer_log(data, thought):
    p = thought["normalized"]
    rows = log_records(data)
    if not rows:
        return "No hay registros de bitácora disponibles."

    # Filtros de estado: primero usamos la estructura de la bitácora y solo
    # después el texto libre.
    wants_pending = any(x in p for x in ("pendiente", "pendientes", "sin resolver", "no resuelto"))
    wants_done = any(x in p for x in ("resuelto", "resueltos", "listo", "reparado", "reparados", "atendido"))
    system = thought.get("system")
    if system not in {"Motores Mercury", "Patas Mercury", "Generador Cummins Onan"}:
        system = None

    candidates = []
    for rec in rows:
        d = rec.get("data") or {}
        if system and d.get("sistema") != system:
            continue
        if wants_pending and d.get("estado") != "PENDIENTE":
            continue
        if wants_done and d.get("estado") != "RESUELTO":
            continue
        candidates.append(rec)

    # Si preguntó explícitamente por pendientes/resueltos, ordenar por fecha
    # y mostrar primero los eventos más recientes.
    if wants_pending or wants_done:
        candidates.sort(key=lambda r: _log_date(r) or date.min, reverse=True)
        if not candidates:
            estado = "pendientes" if wants_pending else "resueltos"
            return f"No encontré eventos {estado} en la bitácora para ese sistema."
        title = "Pendientes" if wants_pending else "Resueltos / atendidos"
        lines = [f"### Bitácora — {title}"]
        for rec in candidates[:15]:
            lines.append(f"- {_format_log_record(rec)}")
        return "\n".join(lines)

    # Consulta normal: búsqueda relevante, pero sin devolver filas ajenas solo
    # porque compartan palabras genéricas como 'motor' o 'generador'.
    tokens = [t for t in re.findall(r"[a-z0-9]+", p) if len(t) >= 4 and t not in STOPWORDS]
    hits = []
    for rec in candidates:
        d = rec.get("data") or {}
        text = norm(" ".join(str(d.get(k) or "") for k in ("observacion", "observaciones", "sistema", "estado")))
        score = 0
        for t in tokens:
            if t in text:
                score += 1
        if score:
            hits.append((score, _log_date(rec) or date.min, rec))
    hits.sort(key=lambda x: (x[0], x[1]), reverse=True)
    if not hits:
        return "No encontré una entrada de bitácora relacionada con la consulta."

    lines = ["### Bitácora — resultados relevantes"]
    for _, _, rec in hits[:8]:
        lines.append(f"- {_format_log_record(rec)}")
    return "\n".join(lines)


def answer_documents(data):
    rows = data.get("documents_records", [])
    if not rows:
        return "No hay información de permisos y seguros disponible."
    lines = ["### Permisos y seguros"]
    for rec in rows:
        vals = rec.get("values", [])
        text = " | ".join(str(v) for v in vals if v not in (None, ""))
        if text:
            lines.append("- " + text)
    return "\n".join(lines[:15])


def answer_sheet(data, sheet_name, title):
    rows = sheet_rows(data, sheet_name)
    if not rows:
        return f"No hay información disponible en la hoja **{sheet_name}**."
    lines = [f"### {title}"]
    for row in rows[:20]:
        vals = row.get("values", []) if isinstance(row, dict) else row
        text = " | ".join(str(v) for v in vals if v not in (None, ""))
        if text: lines.append("- " + text)
    return "\n".join(lines)


def answer_photos(data, question):
    images = data.get("images", [])
    if not images:
        return "No encontré fotografías incrustadas en el Excel."
    target = norm(question)
    preferred = [im for im in images if norm(im.get("sheet", "")) in target or ("permiso" in target and "permiso" in norm(im.get("sheet", "")))]
    chosen = preferred or images
    lines = [f"### Fotografías — {len(images)} encontradas"]
    for im in chosen[:12]:
        loc = f"hoja **{im.get('sheet','desconocida')}**"
        if im.get("row"): loc += f", fila {im['row']}"
        lines.append(f"\n**Fotografía {im.get('index','?')}** — {loc}")
        analysis = str(im.get("vision_analysis") or "").strip()
        if analysis:
            lines.append(analysis)
        else:
            lines.append("No hay interpretación visual guardada para esta fotografía.")
    vision = data.get("metadata", {}).get("vision", {})
    if not vision.get("enabled"):
        lines.append("\n⚠️ La lectura visual todavía no está conectada. Configura el secret `gemini.api_key` en Streamlit para activar la interpretación de las fotos.")
    return "\n".join(lines)


def answer_general(data, thought):
    if thought["system"] == "Combustible":
        return answer_fuel(data, thought)
    if thought["system"] == "Inventario": return answer_inventory(data, thought)
    if thought["system"] == "Facturas": return answer_sheet(data, "FACTURAS 2026...", "Facturas")
    if thought["system"] == "Limpieza": return answer_sheet(data, "S.LIMPIEZA & OTROS", "Limpieza y otros")
    if thought["system"] == "Permisos y seguros": return answer_documents(data)
    return answer_sheet(data, thought.get("system") or "", "Información") if thought.get("system") not in (None, "General") else ("Puedo consultar las 8 hojas del Excel: mantenimiento, bitácora, inventario, combustible, presupuesto, facturas, limpieza y permisos/seguros.\n\nPrueba: **¿Cuánto gastamos en limpieza en septiembre?**")


def run_agent(data, question, context=None):
    thought = think(question)

    # Contexto conversacional INVISIBLE: solo conserva datos necesarios para
    # interpretar referencias como "ese mismo mes". No guarda ni muestra
    # preguntas ni respuestas anteriores.
    context = context or {}
    p = thought["normalized"]
    refers_to_previous_month = bool(re.search(
        r"\b(ese mismo mes|ese mes|en ese mismo mes|en ese mes|mismo mes)\b", p
    ))
    if refers_to_previous_month and thought.get("month") is None:
        thought["month"] = context.get("month")
        if thought.get("year") is None:
            thought["year"] = context.get("year")

    intent = thought["intent"]

    if intent == "estado_general":
        return answer_state(data)
    if intent == "ultimo_mantenimiento":
        return answer_ultimo_mantenimiento(data, thought)
    if intent == "historial_completo":
        return answer_historial(data, thought, complete=True)
    if intent == "cantidad_mantenimiento":
        return answer_count(data, thought)
    if intent == "kit_300":
        return answer_kit(data)
    if intent == "mantenimiento":
        return answer_mantenimiento(data, thought)
    if intent in ("gasto_mensual", "gasto_anual"):
        return answer_expenses(data, thought)
    if intent == "combustible":
        return answer_fuel(data, thought)
    if intent == "inventario":
        return answer_inventory(data, thought)
    if intent == "bitacora":
        return answer_log(data, thought)
    if intent == "documentos":
        # Si preguntan por el contenido visual de permisos/seguros, usar visión.
        if any(x in p for x in ("foto", "fotografia", "fotografía", "imagen", "que dice", "qué dice", "lee", "leer", "muestra")):
            return answer_photos(data, question)
        return answer_documents(data)
    if intent == "fotografias":
        return answer_photos(data, question)
    if intent == "facturas":
        return answer_sheet(data, "FACTURAS 2026...", "Facturas")
    if intent == "limpieza":
        return answer_sheet(data, "S.LIMPIEZA & OTROS", "Limpieza y otros")
    if intent == "fotografias":
        return answer_photos(data, question)
    if intent == "horometros":
        h = current_hours(data)
        return "### Horómetros actuales\n" + "\n".join(f"- **{k}:** {v} h" if v is not None else f"- **{k}:** NO DETERMINADO" for k, v in h.items())
    return answer_general(data, thought)


# ============================================================
# INTERFAZ STREAMLIT — OSCURA Y SIN HISTORIAL
# ============================================================

st.set_page_config(page_title="Agente Tiara", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    background: #0e1117 !important;
}
[data-testid="stHeader"] { background: #0e1117 !important; }
.block-container { max-width: 1100px; padding-top: 1.2rem; padding-bottom: 2rem; }
.hero {
    background: linear-gradient(135deg, #10283d 0%, #174f6d 100%);
    border-radius: 22px;
    padding: 28px 30px;
    margin-bottom: 18px;
}
.hero h1 { color: #ffffff !important; margin: 0; font-size: 2.4rem; }
.hero p { color: #dbe8ef !important; font-size: 1.05rem; margin: 8px 0 0; }
.card {
    background: #171b23;
    border: 1px solid #2a3040;
    border-radius: 18px;
    padding: 18px;
    margin-bottom: 12px;
}
.section-title { color: #dbe8ef; font-size: 1.05rem; font-weight: 700; margin: 1rem 0 .6rem; }
div.stButton > button {
    background: #171b23 !important;
    color: #e7edf2 !important;
    border: 1px solid #394454 !important;
    border-radius: 12px !important;
    min-height: 48px !important;
    font-weight: 650 !important;
}
div.stButton > button:hover { border-color: #6a8da2 !important; }
textarea, input { background: #171b23 !important; color: #f2f5f7 !important; }
[data-testid="stChatInput"] { background: #171b23 !important; }
[data-testid="stChatMessage"] { background: #171b23 !important; border-radius: 16px; }
[data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li { color: #e7edf2; }
small, .stCaption { color: #9da9b5 !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CARGA DEL EXCEL DESDE LA PROPIA APLICACIÓN
# ============================================================
# El usuario puede reemplazar el Excel sin tocar GitHub.
# El .xlsx/.xlsm se procesa internamente: no hay que descomprimirlo.
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data = None
if "uploaded_signature" not in st.session_state:
    st.session_state.uploaded_signature = None
if "upload_message" not in st.session_state:
    st.session_state.upload_message = ""

with st.sidebar:
    st.markdown("### 📥 Actualizar Excel")
    st.caption(
        "Carga el Excel nuevo del barco. La aplicación lo procesa internamente "
        "y guarda la base activa para futuras sesiones."
    )
    st.caption("Límite de carga configurado: 500 MB")
    if _gemini_api_key():
        st.caption("👁️ Visión de fotografías: conectada")
    else:
        st.caption("👁️ Visión de fotografías: requiere configurar Gemini en Secrets")

    uploaded_file = st.file_uploader(
        "Seleccionar Excel",
        type=["xlsx", "xlsm"],
        key="tiara_excel_uploader",
        max_upload_size=500,
        accept_multiple_files=False,
        help=(
            "Carga el Excel original directamente desde Descargas. La aplicación "
            "lo lee internamente, sin que tengas que descomprimirlo. "
            "Se admiten archivos de hasta 500 MB en este cargador."
        ),
    )

    if uploaded_file is not None:
        # La firma se basa en el CONTENIDO real del Excel, no solo en nombre/tamaño.
        # Así, si reemplazas el Excel en Descargas y conserva el mismo nombre
        # y tamaño, la aplicación igualmente detecta la actualización y vuelve
        # a interpretar la Bitácora y los horómetros.
        file_bytes = uploaded_file.getvalue()
        signature = (
            uploaded_file.name,
            uploaded_file.size,
            hashlib.sha256(file_bytes).hexdigest(),
        )

        if st.session_state.uploaded_signature != signature:
            previous_data = st.session_state.uploaded_data

            try:
                with st.spinner(
                    "Procesando Excel… no necesitas descomprimirlo manualmente."
                ):
                    # El archivo se vuelve a procesar completo cada vez que cambia
                    # su contenido. Esto incluye la hoja VITACORA BARCO EN LINEA.
                    uploaded_file.seek(0)
                    new_data = build_from_uploaded_excel(uploaded_file)

                # Guardamos el snapshot completo en Supabase antes de cambiar
                # la fuente activa de la sesión. Si Supabase falla, mantenemos
                # la base anterior y mostramos el error sin romper la aplicación.
                save_supabase_data(new_data)

                # Solo sustituimos la fuente activa cuando el procesamiento y
                # la persistencia terminaron correctamente.
                st.session_state.uploaded_data = new_data
                st.session_state.uploaded_signature = signature
                st.session_state.last_question = ""
                st.session_state.last_answer = ""
                st.session_state.query_context = {"month": None, "year": None}

                meta = new_data.get("metadata", {})
                st.session_state.upload_message = (
                    f"✓ Excel activo: {meta.get('source_file', uploaded_file.name)} "
                    f"· {meta.get('source_size_mb', 0):.2f} MB"
                )

            except Exception as exc:
                # No destruir una base que ya funcionaba por un archivo inválido.
                st.session_state.uploaded_data = previous_data
                st.session_state.upload_message = (
                    f"✕ No se pudo actualizar el Excel: {exc}"
                )

    if st.session_state.upload_message:
        st.caption(st.session_state.upload_message)

# Prioridad de fuentes:
# 1) Excel recién cargado en esta sesión.
# 2) Último snapshot persistente guardado en Supabase.
# 3) Base incluida en GitHub como respaldo.
data = st.session_state.uploaded_data
if not data:
    data = load_supabase_data()
if not data:
    data = load_data()

if not data:
    st.write("No se encontró la base de datos del Agente Tiara.")
    st.stop()

with st.sidebar:
    meta = data.get("metadata", {})
    st.markdown("### 📚 Base de datos")
    st.caption(f"Fuente activa: {meta.get('source_file', 'Excel de septiembre 2026')}")
    st.caption(f"Versión de la aplicación: {APP_VERSION}")
    st.caption(f"Actualización: {APP_UPDATED}")
    st.caption(f"Versión de la base: {meta.get('version', APP_VERSION)}")
    if meta.get("source_size_mb") is not None:
        st.caption(f"Tamaño procesado: {meta.get('source_size_mb')} MB")

# Estado de interfaz: solo mostramos la pregunta y respuesta ACTUALES.
# Se conserva internamente un contexto mínimo para entender frases como
# "ese mismo mes", sin conservar ni mostrar el historial de chat.
if "last_question" not in st.session_state:
    st.session_state.last_question = ""
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""
if "query_context" not in st.session_state:
    st.session_state.query_context = {"month": None, "year": None}

st.markdown("""
<div class="hero">
    <h1>⚓ AGENTE TIARA</h1>
    <p>Centro inteligente de mantenimiento, operación, inventario, combustible y documentación del barco.</p>
</div>
""", unsafe_allow_html=True)

# Los datos administrativos quedan fuera de la pantalla principal.
# Se mantienen disponibles en la barra lateral.
hours = current_hours(data)
maintenance_n = len([
    r for r in maintenance_records(data)
    if norm(r["data"].get("mantenimiento")) not in ("motores", "generador")
])

with st.sidebar:
    st.markdown("### 📊 Datos de la base")
    st.write(f"**Hojas de datos:** {len(data.get('sheets', {}))}")
    st.write(f"**Registros:** {sum(len(data.get(k, [])) for k in ('maintenance_records', 'log_records', 'inventory_records', 'fuel_records', 'budget_records', 'invoice_records', 'cleaning_records', 'documents_records')):,}")
    st.write(f"**Mantenimientos:** {maintenance_n}")
    st.divider()
    st.markdown("### ⏱️ Horómetros")
    for system, value in hours.items():
        st.write(f"**{system}:** {value} h" if value is not None else f"**{system}:** NO DETERMINADO")

st.markdown('<div class="section-title">⚡ Consultas rápidas</div>', unsafe_allow_html=True)

# Exactamente los cinco accesos rápidos definidos en Colab.
q1, q2, q3, q4, q5 = st.columns(5)
quick = None
with q1:
    if st.button("🛥️ Estado general", use_container_width=True):
        quick = "Dame el estado general del barco."
with q2:
    if st.button("🔧 Motores", use_container_width=True):
        quick = "¿Qué mantenimiento les toca próximamente a los motores?"
with q3:
    if st.button("🔩 Kit 300 h", use_container_width=True):
        quick = "¿Cuándo toca el próximo kit de 300 horas de los motores?"
with q4:
    if st.button("⚙️ Patas", use_container_width=True):
        quick = "¿Cuándo toca el próximo cambio de aceite de las patas?"
with q5:
    if st.button("🔌 Generador", use_container_width=True):
        quick = "¿Cuándo toca el próximo cambio de aceite al generador?"

st.markdown('<div class="section-title">🤖 Asistente</div>', unsafe_allow_html=True)

# Solo existe UNA pregunta y UNA respuesta visibles.
question = st.chat_input("Escribe aquí tu consulta sobre el Tiara…")
if question:
    quick = question

if quick:
    # Ejecutar con el contexto oculto de la consulta inmediatamente anterior.
    st.session_state.last_question = quick
    st.session_state.last_answer = run_agent(data, quick, st.session_state.query_context)

    # Actualizar SOLO el contexto útil (mes/año). No se guarda la pregunta ni
    # la respuesta anterior como historial.
    new_thought = think(quick)
    if new_thought.get("month") is not None:
        st.session_state.query_context["month"] = new_thought["month"]
    if new_thought.get("year") is not None:
        st.session_state.query_context["year"] = new_thought["year"]

if st.session_state.last_question:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f"**Pregunta:** {st.session_state.last_question}")
    st.markdown(st.session_state.last_answer)
    st.markdown('</div>', unsafe_allow_html=True)

# La fuente de datos queda disponible en la barra lateral.
