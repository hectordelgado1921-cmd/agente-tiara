import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
import streamlit as st

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
DATA.mkdir(exist_ok=True)
KB_FILE = DATA / "base_conocimiento_tiara.json"
DEFAULT_EXCEL = next(DATA.glob("*.xlsx"), None)
DEFAULT_JSON = DATA / "base_conocimiento_tiara_septiembre_2026.json"


def clean(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return re.sub(r"\s+", " ", v.strip()) if isinstance(v, str) else v


def norm(v):
    s = "" if v is None else str(v).lower()
    return s.translate(str.maketrans("áéíóúüñ", "aeiouun"))


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    m = re.search(r"-?\d+(?:[.,]\d+)?", str(v).replace(",", ""))
    return float(m.group()) if m else None


def header_row(ws):
    best = (1, -1)
    for r in range(1, min(20, ws.max_row) + 1):
        score = sum(
            clean(ws.cell(r, c).value) not in (None, "")
            for c in range(1, ws.max_column + 1)
        )
        if score > best[1]:
            best = (r, score)
    return best[0]


def read_sheet(ws):
    hr = header_row(ws)
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = clean(ws.cell(hr, c).value)
        if v not in (None, ""):
            headers[c] = str(v)

    rows = []
    for r in range(hr + 1, ws.max_row + 1):
        d = {}
        for c, h in headers.items():
            v = clean(ws.cell(r, c).value)
            if v not in (None, ""):
                d[h] = v
        if d:
            rows.append({"source_row": r, "data": d})
    return {"header_row": hr, "records": rows}


def classify(text):
    t = norm(text)
    if "garantia" in t:
        return "GARANTIA"
    if "instal" in t:
        return "INSTALACION"
    if "mantenimiento" in t or "aceite" in t or "filtro" in t:
        return "MANTENIMIENTO"
    if "repar" in t or "arreglo" in t:
        return "REPARACION"
    if any(x in t for x in ("falla", "calent", "problema", "no funciona", "no trabaja")):
        return "FALLA"
    return "OBSERVACION"


def import_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    kb = {
        "metadata": {
            "agent": "Agente Tiara",
            "version": "1.0",
            "source_of_truth": "Excel",
            "source_file": Path(path).name,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        },
        "sheets": {},
        "maintenance_records": [],
        "log_records": [],
        "inventory_items": [],
        "fuel_records": [],
        "budget_records": [],
        "invoice_records": [],
        "cleaning_records": [],
        "documents": [],
    }

    for name in wb.sheetnames:
        block = read_sheet(wb[name])
        kb["sheets"][name] = block
        n = norm(name)
        group = None

        if "checklist" in n:
            group = "maintenance_records"
        elif "vitacora" in n or "bitacora" in n:
            group = "log_records"
            for rec in block["records"]:
                rec["event_type"] = classify(
                    " | ".join(map(str, rec["data"].values()))
                )
        elif n.strip() == "inventario":
            group = "inventory_items"
        elif "combustible" in n:
            group = "fuel_records"
        elif "presupuesto" in n:
            group = "budget_records"
        elif "facturas" in n:
            group = "invoice_records"
        elif "limpieza" in n:
            group = "cleaning_records"
        elif "permisos" in n or "seguros" in n:
            group = "documents"

        if group:
            kb[group].extend([{"source_sheet": name, **r} for r in block["records"]])

    save(kb)
    return kb


def save(kb):
    KB_FILE.write_text(
        json.dumps(kb, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def load():
    return json.loads(KB_FILE.read_text(encoding="utf-8")) if KB_FILE.exists() else None


def all_records(kb):
    for g in (
        "maintenance_records",
        "log_records",
        "inventory_items",
        "fuel_records",
        "budget_records",
        "invoice_records",
        "cleaning_records",
        "documents",
    ):
        for r in kb[g]:
            yield g, r


def search(kb, q, groups=None, limit=20):
    terms = [x for x in re.findall(r"[a-z0-9]+", norm(q)) if len(x) > 2]
    out = []
    for g, r in all_records(kb):
        if groups and g not in groups:
            continue
        blob = norm(json.dumps(r, ensure_ascii=False))
        score = sum(blob.count(t) for t in terms)
        if score:
            out.append((score, g, r))
    out.sort(key=lambda x: x[0], reverse=True)
    return out[:limit]


def costs(kb, q):
    terms = [x for x in re.findall(r"[a-z0-9]+", norm(q)) if len(x) > 2]
    total = 0
    rows = []
    for g, r in all_records(kb):
        d = r["data"]
        blob = norm(json.dumps(d, ensure_ascii=False))
        if terms and not any(t in blob for t in terms):
            continue
        for k, v in d.items():
            if "total" in norm(k):
                n = num(v)
                if n is not None:
                    total += n
                    rows.append((g, r["source_row"], k, n))
    return total, rows


def choose(q):
    t = norm(q)
    if any(x in t for x in (
        "mantenimiento", "aceite", "filtro", "anodo", "bujia",
        "impeler", "impeller", "servicio"
    )):
        return "mantenimiento"
    if any(x in t for x in (
        "falla", "problema", "calent", "temperatura", "bitacora", "repar"
    )):
        return "eventos"
    if any(x in t for x in ("inventario", "stock", "existencia", "repuesto")):
        return "inventario"
    if any(x in t for x in ("factura", "proveedor")):
        return "facturas"
    if any(x in t for x in ("combustible", "gasolina", "diesel", "litros")):
        return "combustible"
    if any(x in t for x in (
        "permiso", "seguro", "vence", "vencimiento", "documento"
    )):
        return "documentos"
    if any(x in t for x in ("cuanto gaste", "cuánto gasté", "costo", "costos", "presupuesto")):
        return "costos"
    return "historial"


def render(results):
    if not results:
        return "No encontré registros relacionados en la base."
    lines = []
    for score, g, r in results[:10]:
        text = " | ".join(
            f"{k}: {v}" for k, v in r["data"].items() if v not in (None, "")
        )
        lines.append(f"**[{g}]** · fila {r['source_row']}\n\n{text[:1000]}")
    return "\n\n---\n\n".join(lines)


def answer(kb, q):
    c = choose(q)
    groups = {
        "mantenimiento": ["maintenance_records"],
        "eventos": ["log_records"],
        "inventario": ["inventory_items"],
        "facturas": ["invoice_records"],
        "combustible": ["fuel_records"],
        "documentos": ["documents"],
        "historial": None,
    }.get(c)

    if c == "costos":
        total, rows = costs(kb, q)
        return f"### Resultado de costos\n**Total: ${total:,.2f} USD**\n\nRegistros encontrados: **{len(rows)}**"

    return f"### Resultados: {c.title()}\n\n" + render(search(kb, q, groups))


def load_default():
    if KB_FILE.exists():
        return load()
    if DEFAULT_JSON.exists():
        return json.loads(DEFAULT_JSON.read_text(encoding="utf-8"))
    if DEFAULT_EXCEL and DEFAULT_EXCEL.exists():
        return import_excel(DEFAULT_EXCEL)
    return None


st.set_page_config(
    page_title="Agente Tiara",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp {
        background: #f4f7fb;
    }
    [data-testid="stHeader"] {
        background: rgba(255,255,255,0);
    }
    .hero {
        padding: 1.6rem 1.8rem;
        border-radius: 22px;
        background: linear-gradient(135deg, #0b1f33 0%, #164a70 55%, #1c718c 100%);
        color: white;
        margin-bottom: 1.2rem;
        box-shadow: 0 12px 35px rgba(11,31,51,.18);
    }
    .hero h1 {
        margin: 0;
        font-size: 2.1rem;
        letter-spacing: -.04em;
    }
    .hero p {
        margin: .35rem 0 0;
        opacity: .86;
        font-size: 1rem;
    }
    .metric {
        background: white;
        border-radius: 17px;
        padding: 1rem 1.1rem;
        border: 1px solid #e5ebf2;
        box-shadow: 0 5px 18px rgba(20,40,60,.06);
    }
    .metric-label {
        color: #6b7785;
        font-size: .82rem;
        margin-bottom: .25rem;
    }
    .metric-value {
        color: #10283d;
        font-size: 1.55rem;
        font-weight: 750;
    }
    .section-title {
        color: #10283d;
        font-size: 1.15rem;
        font-weight: 700;
        margin: 1.1rem 0 .6rem;
    }
    .side-note {
        color: #6b7785;
        font-size: .78rem;
        line-height: 1.45;
    }
    div[data-testid="stChatMessage"] {
        border-radius: 16px;
    }
    .stButton > button {
        border-radius: 11px;
    }
</style>
""", unsafe_allow_html=True)

if "kb" not in st.session_state:
    st.session_state.kb = load_default()
if "messages" not in st.session_state:
    st.session_state.messages = []

kb = st.session_state.kb

with st.sidebar:
    st.markdown("## ⚓ Agente Tiara")
    st.caption("Centro de control del barco")
    st.divider()

    if kb:
        st.success("● Base activa")
        st.caption(f"Fuente: {kb['metadata']['source_file']}")
    else:
        st.warning("○ Sin base cargada")

    uploaded = st.file_uploader(
        "Actualizar fuente de datos",
        type=["xlsx", "xlsm"],
        help="Carga una nueva versión del Excel para reconstruir la base de conocimiento.",
    )
    if uploaded and st.button("Importar Excel", use_container_width=True):
        with st.spinner("Procesando Excel..."):
            st.session_state.kb = import_excel(uploaded)
            st.session_state.messages = []
        st.rerun()

    st.divider()
    st.markdown("**Módulos**")
    st.caption("🔧 Mantenimiento")
    st.caption("📋 Bitácora y eventos")
    st.caption("📦 Inventario")
    st.caption("⛽ Combustible")
    st.caption("🧾 Facturas y costos")
    st.caption("📄 Permisos y seguros")

    st.divider()
    st.markdown(
        '<div class="side-note">La información de operación proviene del Excel cargado como fuente de verdad.</div>',
        unsafe_allow_html=True,
    )

st.markdown("""
<div class="hero">
    <h1>⚓ AGENTE TIARA</h1>
    <p>Tu centro de información para mantenimiento, operación, inventario y documentación del barco.</p>
</div>
""", unsafe_allow_html=True)

if not kb:
    st.info("Carga el Excel de septiembre desde la barra lateral para activar el centro de control.")
    st.stop()

# Dashboard
sheets = len(kb["sheets"])
records = sum(len(kb[g]) for g in (
    "maintenance_records", "log_records", "inventory_items", "fuel_records",
    "budget_records", "invoice_records", "cleaning_records", "documents"
))
maintenance_count = len(kb["maintenance_records"])

c1, c2, c3 = st.columns(3)
for col, label, value in (
    (c1, "Hojas de datos", sheets),
    (c2, "Registros", records),
    (c3, "Mantenimientos", maintenance_count),
):
    col.markdown(
        f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value:,}</div></div>',
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-title">Asistente</div>', unsafe_allow_html=True)

if not st.session_state.messages:
    st.markdown(
        "Escribe una consulta sobre el barco. Puedes preguntar por mantenimiento, bitácora, inventario, combustible, facturas, costos o documentos."
    )

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

prompt = st.chat_input("Escribe aquí tu consulta sobre el Tiara…")
if prompt:
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    response = answer(kb, prompt)
    with st.chat_message("assistant"):
        st.markdown(response)
    st.session_state.messages.append({"role": "assistant", "content": response})

st.markdown('<div class="section-title">Estado del barco</div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">Fuente de datos</div>', unsafe_allow_html=True)

with st.container():
    st.markdown("#### 📚 Fuente de datos")
    st.write(f"**Archivo:** {kb['metadata']['source_file']}")
    st.write(f"**Actualizado:** {kb['metadata'].get('updated_at', 'No disponible')}")
    st.write(f"**Hojas:** {len(kb['sheets'])}")
    st.write(f"**Registros:** {records}")

    with st.expander("Ver hojas"):
        for sheet_name in kb["sheets"]:
            st.write(f"• {sheet_name}")
            st.markdown("""
<style>
.stApp {
    background-color: #0e1117;
}
</style>
""", unsafe_allow_html=True)
