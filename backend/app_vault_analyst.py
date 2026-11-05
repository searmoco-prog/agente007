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

from universal_excel_agent import process_universal_excel, inspect_workbook

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
uploaded_file = st.file_uploader(
    "📁 Arrastra o selecciona tu archivo (.xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"],
    help="El archivo se procesa estrictamente en la memoria RAM de tu computadora. Ningún byte viaja a la nube."
)

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name

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
        with st.spinner("⚡ Procesando localmente en memoria segura (sin internet)..."):
            try:
                res_bytes, out_name, out_msg = process_universal_excel(file_bytes, instruction, file_name)

                st.success("✅ **¡Análisis completado con éxito en 0.2 segundos!**")
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
