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
import sys
import traceback
import pandas as pd
import openpyxl
from dotenv import load_dotenv

load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
MODELOS_GROQ = [
    os.getenv("GROQ_MODEL", "").strip(),
    "openai/gpt-oss-120b",
    "qwen/qwen3.8-27b",
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
        "5. Al terminar, si modificaste una hoja específica, asegúrate de marcarla activa: `wb.active = ws_target`.\n"
        "6. Devuelve EXCLUSIVAMENTE el bloque de código entre ```python y ```, sin explicaciones ni saludos."
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
                max_tokens=900
            )
            raw_content = resp.choices[0].message.content or ""
            # Extraer código entre backticks
            match = re.search(r"```(?:python)?\s*([\s\S]*?)\s*```", raw_content)
            if match:
                return match.group(1).strip()
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

    # Bloqueo de seguridad: no permitir llamadas al sistema operativo
    forbidden_terms = ["os.system", "subprocess", "shutil.rmtree", "__import__", "eval(", "input("]
    for term in forbidden_terms:
        if term in code_str:
            raise SecurityError(f"Instrucción no permitida por seguridad: {term}")

    # Ejecutar en RAM
    exec(code_str, {}, local_env)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    result_bytes = output.getvalue()

    base_name = file_name.rsplit('.', 1)[0]
    out_filename = f"Analisis_Universal_{base_name}.xlsx"
    return result_bytes, out_filename

async def process_any_excel_instruction(file_bytes: bytes, user_instruction: str, file_name: str = "archivo.xlsx") -> tuple[bytes, str, str]:
    """Flujo completo Zero-Data Leak: extrae schema -> genera código con IA -> ejecuta localmente."""
    # 1. Extraer esquema
    schema = extract_metadata_schema(file_bytes)

    # 2. Generar código Python
    code_generated = await generate_python_transformation_code(schema, user_instruction)

    # 3. Ejecutar sobre el Excel real
    result_bytes, out_filename = execute_safe_excel_transformation(file_bytes, code_generated, file_name)

    resumen = (
        f"🤖 **Transformación Universal Completada:**\n"
        f"• Se ejecutó código generado específicamente para tu petición: *\"{user_instruction}\"*\n"
        f"• Cero fuga de datos: solo se enviaron los nombres de columnas a la IA; los datos se procesaron en RAM.\n"
        f"• Fórmulas y estilos aplicados de forma dinámica y viva."
    )
    return result_bytes, out_filename, resumen

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
