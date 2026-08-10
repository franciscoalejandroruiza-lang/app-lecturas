import io
import json
import shutil
import zipfile
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import openpyxl

from excel_reader import load_ok_sheet, detect_month_blocks, read_block_values
from excel_writer import write_values_to_block, create_new_month_block
from fill_formats import fill_template, docx_to_pdf
from capture_formulas import get_profile_for_model, compute_final_values, MODEL_PROFILES

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
    by_serie = {}
    for r in recs:
        by_archivo.setdefault(r["ARCHIVO_ORIGEN"], []).append(r)
        by_serie[str(r["SERIE"]).strip()] = r
    return recs, by_archivo, by_serie

catalog, catalog_by_archivo, catalog_by_serie = load_catalog()
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

tab_generar, tab_capturar = st.tabs(["📄 Generar formatos", "✍️ Capturar lectura por serie"])

# =========================================================================
# TAB 1: Generar formatos (funcionalidad existente)
# =========================================================================
with tab_generar:
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

# =========================================================================
# TAB 2: Capturar lectura por serie (calcula según modelo y guarda en Excel)
# =========================================================================
with tab_capturar:
    st.subheader("Buscar equipo por número de serie")
    serie_input = st.text_input("Número de serie", key="captura_serie").strip()

    rec = catalog_by_serie.get(serie_input) if serie_input else None

    if serie_input and not rec:
        st.warning("Esa serie no está en el catálogo maestro.")

    if rec:
        st.write(f"**Unidad:** {rec['UNIDAD_MEDICA_O_ADMINISTRATIVA']}  |  "
                 f"**Marca/Modelo:** {rec['MARCA']} {rec['MODELO']}  |  "
                 f"**Archivo:** {rec['ARCHIVO_ORIGEN']}")

        profile = get_profile_for_model(rec["MODELO"])

        if not profile:
            st.error(
                f"El modelo '{rec['MODELO']}' no está documentado en la guía de fórmulas. "
                "Modelos disponibles: " + ", ".join(sorted(set(p['nombre'] for p in MODEL_PROFILES.values())))
            )
        else:
            st.info(f"📐 {profile['nombre']}")
            if not profile.get("verificado", True):
                st.warning(
                    "⚠️ La fórmula de este modelo NO se ha confirmado con una guía real (es un mejor-intento). "
                    "Revisa el resultado con cuidado antes de usarlo, o mándame la guía de este perfil para ajustarla."
                )
            st.caption("Captura los totales tal cual aparecen en el reporte de tu impresora.")

            raw_values = {}
            for field_key, field_label, field_help in profile["fields"]:
                raw_values[field_key] = st.number_input(
                    field_label, min_value=0, step=1, key=f"raw_{field_key}", help=field_help,
                )

            computed = compute_final_values(rec["MODELO"], raw_values)

            fecha_hora_dt = st.text_input(
                "Fecha y hora de lectura (como aparece en el reporte, ej. 20/06/2026 13:08)",
                key="captura_fecha",
            )

            st.markdown("**Vista previa del resultado calculado:**")
            st.table({
                "Campo": ["Carta B&N", "Oficio B&N", "Carta Color", "Oficio Color", "Digitalización"],
                "Valor calculado": [
                    computed["carta_bn"], computed["oficio_bn"],
                    computed["carta_color"], computed["oficio_color"], computed["digitalizacion"],
                ],
            })

            st.divider()
            st.markdown("**¿En qué bloque de mes se guarda?**")
            block_options2 = [f"{b['label']}  ({b['col_range']})" for b in blocks]
            target_mode = st.radio(
                "Destino", ["Actualizar un bloque existente", "Crear un bloque de mes nuevo"],
                key="captura_destino",
            )

            if target_mode == "Actualizar un bloque existente":
                target_idx = st.selectbox(
                    "Bloque a actualizar", range(len(blocks)),
                    format_func=lambda i: block_options2[i], index=len(blocks) - 1,
                    key="captura_bloque_existente",
                )
                target_block = blocks[target_idx]
            else:
                nuevo_mes = st.text_input("Nombre del mes nuevo (ej. AGOSTO)", key="captura_mes_nuevo").strip().upper()
                target_block = None

            if st.button("💾 Guardar en el Excel", type="primary"):
                wb = openpyxl.load_workbook(excel_path)
                ws_write = wb["OK"]

                if target_mode == "Crear un bloque de mes nuevo":
                    if not nuevo_mes:
                        st.error("Escribe el nombre del mes nuevo antes de guardar.")
                        st.stop()
                    target_block = create_new_month_block(ws_write, nuevo_mes)
                else:
                    # Re-detectar bloques sobre el workbook de escritura (mismas columnas)
                    target_block = detect_month_blocks(ws_write)[target_idx]

                ok, msg = write_values_to_block(ws_write, serie_input, target_block, computed, fecha_hora_dt or None)
                if not ok:
                    st.error(msg)
                else:
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    st.success(f"✅ Guardado en el bloque: {target_block['label']}")
                    st.download_button(
                        "⬇️ Descargar Excel actualizado",
                        data=buf,
                        file_name="ISSSTE_DETALLE_CONSUMO_ok.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.caption("Reemplaza tu archivo local con este para que quede guardado de forma permanente.")
