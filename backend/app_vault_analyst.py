"""
VaultAnalyst AI - Interfaz Web de Privacidad Zero-Trust (Localhost)
Procesa archivos Excel y CSV 100% de manera local y desconectada sin enviar datos a internet.
Reutiliza el motor analítico universal de Agente 007 con fórmulas vivas, formatos y gráficos.
"""

import os
import sys
import io
import streamlit as st
import pandas as pd
import openpyxl

# Añadir el backend al path para importar el agente universal
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, CURRENT_DIR)

# ── Cargar variables de entorno desde backend/.env ──────────────────────────
# Se hace ANTES de importar módulos que lean GROQ_API_KEY al cargarse.
from dotenv import load_dotenv
_env_path = os.path.join(CURRENT_DIR, ".env")
load_dotenv(dotenv_path=_env_path, override=False)
# ─────────────────────────────────────────────────────────────────────────────

import asyncio
import re
from universal_excel_agent import process_universal_excel, inspect_workbook
from universal_code_interpreter import process_any_excel_instruction, fetch_google_sheet_as_excel_bytes
from agent_graph import run_agent_graph

def get_smart_preview_df(file_bytes: bytes, max_rows: int = 15) -> pd.DataFrame:
    """Construye un DataFrame limpio para la vista previa visual, calculando fórmulas y eliminando 'None'."""
    try:
        wb_f = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        wb_t = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sname = "Base_Datos" if "Base_Datos" in wb_f.sheetnames else wb_f.active.title
        ws_f = wb_f[sname]
        ws_t = wb_t[sname]

        rows = []
        for r in range(1, min(ws_f.max_row + 1, max_rows + 1)):
            r_vals = [ws_t.cell(r, c).value for c in range(1, ws_f.max_column + 1)]
            if any(v is not None for v in r_vals):
                rows.append(r_vals)

        if not rows or len(rows) < 2:
            return pd.DataFrame()

        headers = [str(h).strip() if h is not None and str(h).strip() != "" else f"Col_{i+1}" for i, h in enumerate(rows[0])]
        seen = {}
        clean_h = []
        for h in headers:
            cnt = seen.get(h, 0)
            seen[h] = cnt + 1
            clean_h.append(h if cnt == 0 else f"{h}_{cnt}")

        df = pd.DataFrame(rows[1:], columns=clean_h)

        # 1. Resolver Ganancia ($) si aparece None calculando Ventas - Costo
        for col in df.columns:
            if "ganancia" in col.lower() or "utilidad" in col.lower():
                col_v = next((c for c in df.columns if ("monto" in c.lower() or "venta" in c.lower()) and "id" not in c.lower()), None)
                col_c = next((c for c in df.columns if "costo" in c.lower() or "gasto" in c.lower()), None)
                if col_v and col_c:
                    for idx in df.index:
                        val_curr = df.at[idx, col]
                        if pd.isna(val_curr) or val_curr is None or str(val_curr).strip().lower() in ["none", ""]:
                            try:
                                v_num = float(str(df.at[idx, col_v]).replace("$", "").replace(",", "").strip())
                                c_num = float(str(df.at[idx, col_c]).replace("$", "").replace(",", "").strip())
                                df.at[idx, col] = v_num - c_num
                            except (ValueError, TypeError):
                                pass

        # 2. Resolver % Margen si aparece None calculando Ganancia / Ventas
        for col in df.columns:
            if "margen" in col.lower():
                col_v = next((c for c in df.columns if ("monto" in c.lower() or "venta" in c.lower()) and "id" not in c.lower()), None)
                col_g = next((c for c in df.columns if "ganancia" in c.lower() or "utilidad" in c.lower()), None)
                if col_v and col_g:
                    for idx in df.index:
                        val_curr = df.at[idx, col]
                        if pd.isna(val_curr) or val_curr is None or str(val_curr).strip().lower() in ["none", ""]:
                            try:
                                v_num = float(str(df.at[idx, col_v]).replace("$", "").replace(",", "").strip())
                                g_num = float(str(df.at[idx, col_g]).replace("$", "").replace(",", "").strip())
                                if v_num > 0:
                                    df.at[idx, col] = f"{(g_num / v_num) * 100:.1f}%"
                            except (ValueError, TypeError):
                                pass

        # 3. Limpiar cualquier 'None' residual para que no se vea en la web
        df = df.astype(object)
        for col in df.columns:
            df[col] = df[col].apply(lambda x: "" if x is None or str(x).strip().lower() == "none" else x)

        return df
    except Exception:
        return pd.DataFrame()

st.set_page_config(
    page_title="VaultAnalyst AI - Zero-Trust Local",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS corporativos
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E78;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #555;
        margin-bottom: 20px;
    }
    .security-badge {
        background-color: #E2EFDA;
        border: 1px solid #375623;
        color: #375623;
        padding: 6px 14px;
        border-radius: 6px;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Barra lateral informativa
with st.sidebar:
    st.image("https://img.icons8.com/color/96/shield.png", width=64)
    st.markdown("### 🛡️ **VaultAnalyst AI**")
    st.markdown("**Arquitectura Zero-Trust On-Premise**")
    st.markdown("---")
    st.markdown("""
    **🔒 Garantías de Privacidad:**
    - Cero transmisión de datos a servidores externos.
    - Cómputo ejecutado al 100% en la RAM de este equipo.
    - Fórmulas dinámicas de Excel nativas (`VLOOKUP`, `IFERROR`, `COUNTIF`).
    - Compatible con políticas de Compliance bancario (GDPR / ISO 27001).
    """)
    st.markdown("---")
    st.caption("Agente 007 Suite • Versión 3.6 Localhost")

# Cabecera principal
st.markdown('<div class="main-header">🛡️ VaultAnalyst AI — Analista Bóveda Local</div>', unsafe_allow_html=True)
st.markdown('<div class="security-badge">🟢 MODO AIR-GAPPED ACTIVO: Procesamiento 100% Desconectado y Privado</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Arrastra tu archivo Excel o CSV para analizar márgenes, crear tablas dinámicas, incrustar gráficos y generar fórmulas vivas sin exponer información confidencial.</div>', unsafe_allow_html=True)

# Sección de Carga de Archivos Drag & Drop
tab_local, tab_sheet = st.tabs(["📁 Archivo Local (.xlsx, .csv)", "🌐 Enlace Google Sheets en Vivo"])

uploaded_file = None
file_bytes = None
file_name = "archivo.xlsx"

with tab_local:
    uploaded_local = st.file_uploader(
        "Arrastra o selecciona tu archivo (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        help="El archivo se procesa estrictamente en la memoria RAM de tu computadora. Ningún byte viaja a la nube."
    )
    if uploaded_local is not None:
        file_bytes = uploaded_local.read()
        file_name = uploaded_local.name

with tab_sheet:
    sheet_url_input = st.text_input(
        "Pega el enlace público o compartido de Google Sheets:",
        placeholder="https://docs.google.com/spreadsheets/d/.../edit#gid=0"
    )
    if st.button("📥 Importar Google Sheet"):
        if sheet_url_input:
            with st.spinner("Descargando Google Sheet en memoria segura..."):
                try:
                    file_bytes, file_name = asyncio.run(fetch_google_sheet_as_excel_bytes(sheet_url_input))
                    st.session_state["sheet_bytes"] = file_bytes
                    st.session_state["sheet_name"] = file_name
                    st.success("✅ ¡Google Sheet importado con éxito a memoria local!")
                except Exception as e_sh:
                    st.error(f"Error importando Google Sheet: {e_sh}")

if file_bytes is None and "sheet_bytes" in st.session_state:
    file_bytes = st.session_state["sheet_bytes"]
    file_name = st.session_state.get("sheet_name", "Google_Sheet_En_Vivo.xlsx")

if file_bytes is not None:

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric(label="Nombre del Archivo", value=file_name)
    with col_info2:
        st.metric(label="Tamaño", value=f"{len(file_bytes) / 1024:.1f} KB")
    with col_info3:
        st.metric(label="Estado de Privacidad", value="Aislado Local")

    # Vista previa inmediata del archivo cargado para que el usuario siempre vea sus datos
    with st.expander("👁️ Vista Previa del Archivo Cargado (Primeras 10 filas)", expanded=True):
        df_init_preview = get_smart_preview_df(file_bytes)
        if not df_init_preview.empty:
            st.dataframe(df_init_preview.head(10), use_container_width=True)

    st.markdown("---")
    st.markdown("### ⚙️ Selecciona las Transformaciones a Aplicar:")

    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
    with col_btn1:
        opt1 = st.button("1️⃣ Margen + Verde > 35%", use_container_width=True)
    with col_btn2:
        opt2 = st.button("2️⃣ Tabla + Gráfico Circular", use_container_width=True)
    with col_btn3:
        opt3 = st.button("3️⃣ Buscador BUSCARV", use_container_width=True)
    with col_btn4:
        opt_all = st.button("⭐ APLICAR TODAS", type="primary", use_container_width=True)

    custom_instruction = st.text_input(
        "💬 O escribe una instrucción personalizada en lenguaje natural:",
        placeholder="Ej: Calcula el margen de ganancia y resalta los mejores en verde...",
        key="custom_inst"
    )
    btn_custom = st.button("⚡ Ejecutar Instrucción Personalizada")

    # Determinar qué acción ejecutar
    instruction = None
    if opt1:
        instruction = "Calcula el margen y resalta en verde > 35%"
    elif opt2:
        instruction = "Genera tabla dinámica y gráfico circular por Estado"
    elif opt3:
        instruction = "Agrega módulo de búsqueda con BUSCARV / XLOOKUP"
    elif opt_all:
        instruction = "Calcula el margen y resalta en verde > 35%, genera tabla dinámica con gráfico circular por Estado y agrega módulo de búsqueda con BUSCARV"
    elif btn_custom and custom_instruction:
        instruction = custom_instruction

    if instruction:
        with st.spinner("⚡ Procesando con Inteligencia Artificial..."):
            try:
                # Orquestación formal mediante LangGraph StateGraph
                graph_history = []
                if opt1 or opt2 or opt3 or opt_all:
                    res_bytes, out_name, out_msg = process_universal_excel(file_bytes, instruction, file_name)
                    graph_history = ["router_node -> Preset rápido", "deterministic_node: OpenPyXL directo"]
                else:
                    groq_key = os.environ.get("GROQ_API_KEY", "").strip()
                    if groq_key:
                        res_bytes, out_name, out_msg, graph_history = asyncio.run(run_agent_graph(file_bytes, instruction, file_name))
                    else:
                        res_bytes, out_name, out_msg = process_universal_excel(file_bytes, instruction, file_name)
                        graph_history = ["router_node: Sin clave Groq -> Fallback a motor determinista"]

                st.success("✅ **¡Análisis completado con éxito!**")
                st.info(out_msg)

                # Mostrar traza de ejecución de LangGraph
                if graph_history:
                    with st.expander("🧩 Nodos ejecutados en LangGraph StateGraph", expanded=False):
                        for step in graph_history:
                            st.write(f"• `{step}`")

                # Botón de descarga directa (garantizado siempre visible)
                st.download_button(
                    label=f"📥 Descargar {out_name} (Con Fórmulas Vivas)",
                    data=res_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Previsualización rápida inteligente de los datos resultantes
                df_res_preview = get_smart_preview_df(res_bytes)
                if not df_res_preview.empty:
                    st.markdown("#### 📊 Vista Previa de los Datos con Fórmulas Calculadas:")
                    st.dataframe(df_res_preview.head(10), use_container_width=True)

            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("👆 Arrastra un archivo Excel arriba para comenzar la prueba local.")
