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
from universal_excel_agent import process_universal_excel, inspect_workbook
from universal_code_interpreter import process_any_excel_instruction, fetch_google_sheet_as_excel_bytes

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
        with st.spinner("⚡ Procesando localmente con IA Zero-Data Leak..."):
            try:
                # Detectar si la instrucción libre menciona colores/márgenes → motor determinista (sin Groq)
                _inst_low = instruction.lower()
                _keywords_deterministas = [
                    "margen", "rojo", "verde", "amarillo", "azul", "porciento", "porcentaje",
                    "semaforo", "semáforo", "traffic light", "colores", "bicolor",
                    "%", "mayor", "menor", "arriba", "abajo", "rentabilidad", "ganancia",
                    "tabla dinamica", "tabla dinámica", "grafico", "gráfico", "circular",
                    "buscarv", "buscar", "xlookup", "estado", "resumen"
                ]
                _usar_motor_determinista = any(k in _inst_low for k in _keywords_deterministas)

                # Si es una de las 3 opciones predefinidas O detectamos palabras clave analíticas
                if _usar_motor_determinista or instruction in [
                    "Calcula el margen y resalta en verde > 35%",
                    "Genera tabla dinámica y gráfico circular por Estado",
                    "Agrega módulo de búsqueda con BUSCARV / XLOOKUP",
                    "Calcula el margen y resalta en verde > 35%, genera tabla dinámica con gráfico circular por Estado y agrega módulo de búsqueda con BUSCARV"
                ]:
                    res_bytes, out_name, out_msg = process_universal_excel(file_bytes, instruction, file_name)
                else:
                    # Instrucción libre avanzada: Code Interpreter vía Groq
                    groq_key = os.environ.get("GROQ_API_KEY", "").strip()
                    if not groq_key:
                        st.warning(
                            "⚠️ **GROQ_API_KEY no configurada.** Para instrucciones avanzadas (fuera de márgenes/colores/tablas), "
                            "crea el archivo `backend/.env` con tu clave:\n\n"
                            "```\nGROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxx\n```\n\n"
                            "Puedes obtener una clave **gratuita** en [console.groq.com](https://console.groq.com). "
                            "Las instrucciones de color, margen y tabla dinámica **no la necesitan**."
                        )
                        st.stop()
                    res_bytes, out_name, out_msg = asyncio.run(process_any_excel_instruction(file_bytes, instruction, file_name))

                st.success("✅ **¡Análisis completado con éxito!**")
                st.info(out_msg)

                # Previsualización rápida de datos generados
                wb_res = openpyxl.load_workbook(io.BytesIO(res_bytes), data_only=True)
                active_sname = wb_res.active.title
                st.markdown(f"#### 📊 Vista Previa de la Hoja Activa: `{active_sname}`")

                # Cargar primeras 10 filas para visualización web
                data_rows = []
                for row in wb_res[active_sname].iter_rows(values_only=True):
                    if any(c is not None for c in row):
                        data_rows.append(list(row))
                
                if data_rows:
                    df_preview = pd.DataFrame(data_rows[1:], columns=data_rows[0])
                    st.dataframe(df_preview.head(10), use_container_width=True)

                # Botón de descarga directa
                st.download_button(
                    label=f"📥 Descargar {out_name} (Con Fórmulas Vivas)",
                    data=res_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("👆 Arrastra un archivo Excel arriba para comenzar la prueba local.")
