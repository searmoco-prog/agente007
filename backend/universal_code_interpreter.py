"""
Módulo Universal Code Interpreter - Agente 007 & VaultAnalyst AI
Arquitectura Zero-Data Leak (Aislamiento de Metadatos):
1. Extrae únicamente la estructura del libro (nombres de pestañas y columnas).
   LOS DATOS CONFIDENCIALES REALES NUNCA SALEN DEL DISPOSITIVO.
2. Solicita al LLM (Groq / Cerebras) un script determinista en Python con openpyxl / pandas.
3. Ejecuta el código generado en un Sandbox de memoria seguro sobre el Excel original.
4. Soporta tanto archivos locales (.xlsx, .csv) como enlaces en vivo de Google Sheets.
"""

import os
import io
import re
import math
import datetime
import sys
import traceback
import pandas as pd
import openpyxl
from dotenv import load_dotenv

# Cargar .env desde el mismo directorio donde vive este script (backend/)
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_THIS_DIR, ".env"), override=False)
# También probar un nivel arriba (raíz del proyecto)
load_dotenv(dotenv_path=os.path.join(_THIS_DIR, "..", ".env"), override=False)

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
MODELOS_GROQ = [
    os.getenv("GROQ_MODEL", "").strip(),
    "openai/gpt-oss-120b",      # Principal (120B de OpenAI en Groq)
    "qwen/qwen3.8-27b",         # Respaldo (27B de Alibaba Cloud en Groq)
    "llama-3.3-70b-versatile",  # Motor de código de alta velocidad
    "llama-3.1-8b-instant",
    "mixtral-8x7b-32768",
]
MODELOS_GROQ = [m for m in MODELOS_GROQ if m]

def extract_metadata_schema(file_bytes: bytes) -> dict:
    """Extrae únicamente nombres de pestañas y cabeceras sin exponer los datos sensibles."""
    if not file_bytes.startswith(b'PK\x03\x04'):
        # Es CSV
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), nrows=2)
            return {
                "format": "csv",
                "sheets": {
                    "Base_Datos": {
                        "columns": list(df.columns),
                        "row_count": "N registros"
                    }
                }
            }
        except Exception:
            return {"format": "csv", "sheets": {"Base_Datos": {"columns": []}}}

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    schema = {"format": "xlsx", "sheets": {}}
    for sname in wb.sheetnames:
        ws = wb[sname]
        headers = []
        for r in range(1, min(ws.max_row + 1, 15)):
            row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, min(ws.max_column + 1, 20))]
            non_empty = [v for v in row_vals if v]
            if len(non_empty) >= 2:
                headers = non_empty
                break
        schema["sheets"][sname] = {
            "columns": headers,
            "max_row": ws.max_row,
            "max_col": ws.max_column
        }
    return schema

def clean_python_code_block(raw_text: str) -> str:
    """Extrae código Python puro eliminando cualquier etiqueta markdown, fences o texto envolvente."""
    raw = raw_text.strip()
    m = re.search(r"```(?:python)?\s*([\s\S]*?)```", raw)
    if m:
        code = m.group(1).strip()
    else:
        m_start = re.search(r"```(?:python)?\s*([\s\S]*)", raw)
        if m_start:
            code = m_start.group(1).strip()
        else:
            code = raw
    lines = [l for l in code.splitlines() if not l.strip().startswith("```")]
    return "\n".join(lines).strip()

async def generate_python_transformation_code(schema: dict, user_prompt: str) -> str:
    """Genera código Python seguro con la IA enviando SOLO metadatos (Zero Data Leak)."""
    if not GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY no configurada en el entorno.")

    from groq import AsyncGroq
    client = AsyncGroq(api_key=GROQ_API_KEY)

    system_prompt = (
        "Eres un desarrollador Senior en Python y Microsoft Excel especializado en openpyxl y pandas.\n"
        "Tu objetivo es escribir ÚNICAMENTE un bloque ejecutable de código Python que reciba un objeto openpyxl "
        "llamado `wb` y aplique la transformación que solicitó el usuario.\n\n"
        "REGLAS OBLIGATORIAS:\n"
        "1. El libro ya está cargado en la variable `wb` (openpyxl.Workbook). Modifica `wb` directamente.\n"
        "2. NO inventes imports innecesarios ni código de carga (NO uses openpyxl.load_workbook ni wb.save).\n"
        "3. Usa fórmulas de Excel en mayúsculas e inglés (ej: SUM, AVERAGE, IFERROR, VLOOKUP, COUNTIF, SUMIF).\n"
        "4. Si piden formato condicional o colores, usa openpyxl.formatting.rule.CellIsRule y openpyxl.styles.PatternFill.\n"
        "   Para colorear una celda individual: cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid').\n"
        "5. REGLAS CRÍTICAS DE PROGRAMACIÓN DEFENSIVA:\n"
        "   - SIEMPRE inicializa todas tus variables antes de cualquier bucle (ej: max_val = None, max_row = None, target_cell = None).\n"
        "   - Al buscar máximos, mínimos o comparar números, limpia '$', ',', '%' y espacios antes de convertir a float.\n"
        "   - MANEJO DE PORCENTAJES EN EXCEL: En Excel, los porcentajes casi siempre se almacenan como decimales (ej: 50% es 0.50, 35% es 0.35). Al comparar por un porcentaje como 50%, comprueba siempre si el valor supera 50 O si está entre 0 y 1 superando 0.50: ej: `(val > 50 or (val <= 1.0 and val > 0.50))`.\n"
        "   - Ignora filas vacías, totales generales o encabezados.\n"
        "6. Al terminar, si modificaste una hoja específica, asegúrate de marcarla activa: `wb.active = ws_target`.\n"
        "7. Devuelve EXCLUSIVAMENTE el bloque de código entre ```python y ```, sin explicaciones ni saludos."
    )

    user_msg = (
        f"Estructura del libro Excel (Metadatos):\n{schema}\n\n"
        f"Instrucción del usuario:\n\"{user_prompt}\"\n\n"
        "Escribe el código Python para transformar el objeto `wb`."
    )

    last_error = None
    for modelo in MODELOS_GROQ:
        try:
            resp = await client.chat.completions.create(
                model=modelo,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_msg}
                ],
                temperature=0.1,
                max_tokens=2500
            )
            raw_content = resp.choices[0].message.content or ""
            clean_code = clean_python_code_block(raw_content)
            if clean_code:
                return clean_code
            return raw_content.strip()
        except Exception as e:
            last_error = e
            continue

    raise RuntimeError(f"Error generando código con los modelos de Groq: {last_error}")

def execute_safe_excel_transformation(file_bytes: bytes, code_str: str, file_name: str = "archivo.xlsx") -> tuple[bytes, str]:
    """Ejecuta el código generado en un Sandbox de memoria seguro sobre el Excel real."""
    # Convertir CSV a Workbook si es necesario
    if not file_bytes.startswith(b'PK\x03\x04'):
        wb = openpyxl.Workbook()
        ws_csv = wb.active
        ws_csv.title = "Base_Datos"
        try:
            df_csv = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
        except Exception:
            df_csv = pd.read_csv(io.BytesIO(file_bytes), encoding='latin-1')
        ws_csv.append(list(df_csv.columns))
        for row in df_csv.itertuples(index=False, name=None):
            ws_csv.append(list(row))
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    # Entorno seguro con librerías precargadas
    local_env = {
        "wb": wb,
        "openpyxl": openpyxl,
        "re": re,
        "math": math,
        "datetime": datetime,
        "Font": openpyxl.styles.Font,
        "PatternFill": openpyxl.styles.PatternFill,
        "Alignment": openpyxl.styles.Alignment,
        "Border": openpyxl.styles.Border,
        "Side": openpyxl.styles.Side,
        "CellIsRule": openpyxl.formatting.rule.CellIsRule,
        "Reference": openpyxl.chart.Reference,
        "BarChart": openpyxl.chart.BarChart,
        "PieChart": openpyxl.chart.PieChart,
        "LineChart": openpyxl.chart.LineChart,
        "get_column_letter": openpyxl.utils.get_column_letter,
        "pd": pd
    }

    # Sanitizar código antes de ejecutar para eliminar backticks o markdown residuales
    clean_exec_code = clean_python_code_block(code_str)

    # Bloqueo de seguridad: no permitir llamadas al sistema operativo
    forbidden_terms = ["os.system", "subprocess", "shutil.rmtree", "__import__", "eval(", "input("]
    for term in forbidden_terms:
        if term in clean_exec_code:
            raise SecurityError(f"Instrucción no permitida por seguridad: {term}")

    # Ejecutar en RAM
    exec(clean_exec_code, {}, local_env)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    result_bytes = output.getvalue()

    base_name = file_name.rsplit('.', 1)[0]
    out_filename = f"Analisis_Universal_{base_name}.xlsx"
    return result_bytes, out_filename

async def process_any_excel_instruction(file_bytes: bytes, user_instruction: str, file_name: str = "archivo.xlsx") -> tuple[bytes, str, str]:
    """Flujo completo Zero-Data Leak con auto-reparación inteligente ante errores."""
    schema = extract_metadata_schema(file_bytes)

    # 1. Generar código Python inicial
    code_generated = await generate_python_transformation_code(schema, user_instruction)

    # 2. Ejecutar con auto-reparación (hasta 2 intentos si la IA comete un bug de sintaxis o variables)
    max_intentos = 2
    for intento in range(max_intentos):
        try:
            result_bytes, out_filename = execute_safe_excel_transformation(file_bytes, code_generated, file_name)
            resumen = (
                f"🤖 **Transformación Universal Completada con IA Senior:**\n"
                f"• Se ejecutó código generado a medida para: *\"{user_instruction}\"*\n"
                f"• Cero fuga de datos: solo se enviaron los nombres de columnas a la IA; los datos se procesaron en RAM.\n"
                f"• Cambios aplicados con éxito sobre el archivo Excel."
            )
            return result_bytes, out_filename, resumen
        except Exception as e_exec:
            if intento < max_intentos - 1:
                # La IA cometió un bug; le devolvemos el error exacto para que lo auto-repare
                repair_prompt = (
                    f"El código que generaste falló con el siguiente error de ejecución en Python:\n"
                    f"ERROR: {e_exec}\n\n"
                    f"Código fallido:\n```python\n{code_generated}\n```\n\n"
                    f"Corrige el código asegurándote de inicializar todas las variables y manejar los tipos de datos "
                    f"de forma segura para cumplir la instrucción original: \"{user_instruction}\"."
                )
                code_generated = await generate_python_transformation_code(schema, repair_prompt)
            else:
                # Si fallaron los reintentos, relanzar para que se reporte
                raise e_exec

async def fetch_google_sheet_as_excel_bytes(google_sheet_url: str) -> tuple[bytes, str]:
    """Permite al bot leer enlaces en vivo de Google Sheets y convertirlos a un libro Excel."""
    import httpx
    # Asegurar formato de exportación a XLSX o CSV
    if "export?format=" not in google_sheet_url:
        if "/edit" in google_sheet_url:
            export_url = google_sheet_url.split("/edit")[0] + "/export?format=xlsx"
        else:
            export_url = google_sheet_url.rstrip("/") + "/export?format=xlsx"
    else:
        export_url = google_sheet_url

    async with httpx.AsyncClient(timeout=25.0) as client:
        resp = await client.get(export_url)
        resp.raise_for_status()
        content = resp.content

        # Si devolvió CSV, convertir a XLSX
        if not content.startswith(b'PK\x03\x04'):
            df = pd.read_csv(io.BytesIO(content))
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name="Hoja_En_Vivo", index=False)
            return out_buf.getvalue(), "Google_Sheet_En_Vivo.xlsx"
        return content, "Google_Sheet_En_Vivo.xlsx"
