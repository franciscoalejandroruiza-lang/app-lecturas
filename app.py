import io
import json
import shutil
import zipfile
import tempfile
from pathlib import Path

import streamlit as st

from excel_reader import load_ok_sheet, detect_month_blocks, read_block_values
from fill_formats import fill_template, docx_to_pdf

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
CATALOG_PATH = BASE_DIR / "data" / "catalog_final.json"

st.set_page_config(page_title="Lecturas ISSSTE", layout="wide")
st.title("📋 Generador de Formatos de Lectura — ISSSTE Delegación Estatal Chihuahua")

# ---------------- Catálogo maestro ----------------
@st.cache_data
def load_catalog():
    recs = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    by_archivo = {}
    for r in recs:
        by_archivo.setdefault(r["ARCHIVO_ORIGEN"], []).append(r)
    return recs, by_archivo

catalog, catalog_by_archivo = load_catalog()
ubicaciones = sorted(catalog_by_archivo.keys())

st.sidebar.header("1) Excel de consumo mensual")
excel_file = st.sidebar.file_uploader("Sube ISSSTE_DETALLE_CONSUMO (.xlsx)", type=["xlsx"])

if not excel_file:
    st.info("⬅️ Sube el Excel de consumo mensual actualizado para empezar.")
    st.caption(f"Catálogo maestro cargado: {len(catalog)} equipos en {len(ubicaciones)} ubicaciones.")
    st.stop()

with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp.write(excel_file.getvalue())
    excel_path = tmp.name

try:
    ws = load_ok_sheet(excel_path)
    blocks = detect_month_blocks(ws)
except Exception as e:
    st.error(f"No se pudo leer el Excel: {e}")
    st.stop()

if len(blocks) < 1:
    st.error("No se detectaron bloques de lectura mensual en la hoja 'OK'.")
    st.stop()

st.sidebar.header("2) Modo de llenado")
modo_label = st.sidebar.radio(
    "¿Qué vas a generar?",
    ["Lectura inicial (formato en blanco para enviar a sitio)",
     "Lectura final (formato completo, listo para firma)"],
)
modo = "inicial" if modo_label.startswith("Lectura inicial") else "final"

st.sidebar.header("3) Bloque de lectura")
block_options = [f"{b['label']}  ({b['col_range']})" for b in blocks]
default_idx = len(blocks) - 1
actual_idx = st.sidebar.selectbox(
    "Bloque ACTUAL (el mes que estás procesando)",
    range(len(blocks)), format_func=lambda i: block_options[i], index=default_idx,
)

if actual_idx == 0:
    st.sidebar.error("El bloque elegido es el primero de la hoja: no hay bloque anterior con el cual generar la lectura inicial / calcular el volumen.")
    st.stop()
anterior_idx = actual_idx - 1

if modo == "inicial":
    st.sidebar.caption(
        f"En modo 'Lectura inicial' se usa el mes ANTERIOR ya capturado como punto de "
        f"partida: {block_options[anterior_idx]}. Su valor se copia al campo 'Lectura "
        f"inicial' del formato en blanco, listo para imprimir y mandar a sitio a que "
        f"capturen la lectura final de {block_options[actual_idx]}."
    )
else:
    st.sidebar.caption(
        f"Bloque ANTERIOR (Lectura inicial) detectado automáticamente: "
        f"{block_options[anterior_idx]}"
    )

anterior_vals = read_block_values(ws, blocks[anterior_idx])
actual_vals = read_block_values(ws, blocks[actual_idx]) if modo == "final" else {}

st.sidebar.header("4) Formato de salida")
formato_salida = st.sidebar.radio(
    "¿Qué archivos quieres generar?",
    ["Word (.docx) — no requiere nada extra instalado",
     "PDF — requiere LibreOffice instalado en esta máquina"],
)
generar_pdf = formato_salida.startswith("PDF")

st.sidebar.header("5) Ubicaciones a generar")
select_all = st.sidebar.checkbox("Seleccionar todas", value=True)
chosen = st.sidebar.multiselect(
    "Ubicaciones", ubicaciones, default=ubicaciones if select_all else [],
)

st.subheader("Vista previa")
st.write(f"**Modo:** {modo_label}")
st.write(f"**Bloque anterior (Lectura inicial):** {block_options[anterior_idx]}")
if modo == "final":
    st.write(f"**Bloque actual (Lectura final):** {block_options[actual_idx]}")
st.write(f"**Ubicaciones seleccionadas:** {len(chosen)} de {len(ubicaciones)}")

preview_rows = []
for archivo in chosen[:3]:
    for rec in catalog_by_archivo[archivo][:2]:
        serie = str(rec["SERIE"]).strip()
        preview_rows.append({
            "Archivo": archivo,
            "Unidad": rec["UNIDAD_MEDICA_O_ADMINISTRATIVA"],
            "Serie": serie,
            "Inicial Carta B&N": anterior_vals.get(serie, {}).get("carta_bn"),
            "Final Carta B&N": actual_vals.get(serie, {}).get("carta_bn") if modo == "final" else "",
            "Inicial Oficio B&N": anterior_vals.get(serie, {}).get("oficio_bn"),
            "Final Oficio B&N": actual_vals.get(serie, {}).get("oficio_bn") if modo == "final" else "",
            "Fecha y Hora": actual_vals.get(serie, {}).get("fecha_hora") if modo == "final" else "",
        })
if preview_rows:
    st.dataframe(preview_rows, use_container_width=True)

st.divider()

if st.button("🚀 Generar formatos", type="primary", disabled=not chosen):
    out_dir = Path(tempfile.mkdtemp())
    output_paths = []
    warnings = []
    progress = st.progress(0.0, text="Generando...")

    for n, archivo in enumerate(chosen, start=1):
        template_path = TEMPLATES_DIR / archivo
        if not template_path.exists():
            warnings.append(f"⚠️ No se encontró la plantilla original para: {archivo}")
            continue

        recs = catalog_by_archivo[archivo]
        readings_by_serie = {}
        for rec in recs:
            serie = str(rec["SERIE"]).strip()
            ini = anterior_vals.get(serie, {})
            fin = actual_vals.get(serie, {}) if modo == "final" else {}
            readings_by_serie[serie] = {"inicial": ini, "final": fin}
            if serie not in anterior_vals:
                warnings.append(f"⚠️ Serie {serie} ({archivo}) no encontrada en el Excel — quedará en blanco.")

        docx_out = out_dir / archivo
        try:
            fill_template(template_path, readings_by_serie, modo, docx_out)
            if generar_pdf:
                final_path = docx_to_pdf(docx_out, out_dir)
            else:
                final_path = docx_out
            output_paths.append(final_path)
        except Exception as e:
            warnings.append(f"❌ Error generando {archivo}: {e}")

        progress.progress(n / len(chosen), text=f"Generando... {n}/{len(chosen)}")

    progress.empty()

    if warnings:
        with st.expander(f"⚠️ {len(warnings)} avisos", expanded=True):
            for w in warnings:
                st.write(w)

    if output_paths:
        ext = "pdf" if generar_pdf else "docx"
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in output_paths:
                zf.write(p, arcname=p.name)
        zip_buf.seek(0)
        st.success(f"✅ Se generaron {len(output_paths)} archivos ({ext.upper()}).")
        st.download_button(
            f"⬇️ Descargar todos los {ext.upper()} (.zip)",
            data=zip_buf,
            file_name=f"formatos_lectura_{modo}.zip",
            mime="application/zip",
        )
    else:
        st.error("No se generó ningún archivo.")
