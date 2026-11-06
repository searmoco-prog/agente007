"""
Módulo de Higiene y Limpieza Automática de Datos para Agente 007
Propósito:
1. Inspeccionar archivos Excel/CSV crudos antes de que los procese cualquier agente o LLM.
2. Resolver fórmulas vacías (None) calculando valores numéricos planos (ej. Ganancia = Ventas - Costo).
3. Convertir strings numéricos ("$15,000.00", "35.5%") a números float puros en las celdas.
4. Identificar y etiquetar filas de Totales/Subtotales para que no confundan a la IA.
5. Garantizar Zero Data Leak: todo el proceso ocurre 100% en RAM local.
"""

import io
import re
import openpyxl
import pandas as pd
from typing import Tuple, Dict, Any, List

def to_clean_float(val: Any) -> float | None:
    """Convierte un valor de celda sucio ($1,234.50, 35%, etc.) a float puro."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null", "-", "n/a"):
        return None
    # Eliminar símbolos monetarios y separadores de miles
    s_clean = s.replace("$", "").replace("€", "").replace("£", "").replace(",", "").strip()
    is_percent = "%" in s_clean
    s_clean = s_clean.replace("%", "").strip()
    try:
        f = float(s_clean)
        if is_percent and f > 1.0:
            return f / 100.0
        return f
    except ValueError:
        return None

def is_total_row_text(val: Any) -> bool:
    """Verifica si un texto corresponde a una fila de resumen/total."""
    if val is None:
        return False
    s = str(val).lower().strip()
    keywords = ["total general", "total", "subtotal", "promedio", "average", "totales"]
    return any(kw == s or kw in s for kw in keywords)

def clean_and_evaluate_workbook(file_bytes: bytes) -> Tuple[bytes, Dict[str, Any]]:
    """
    Higieniza un libro Excel en memoria.
    Evalúa fórmulas vacías, limpia formatos y genera números planos en RAM.
    Retorna: (cleaned_file_bytes, hygiene_report)
    """
    report = {
        "sheets_processed": [],
        "formulas_evaluated": 0,
        "strings_numeric_converted": 0,
        "total_rows_detected": {},
        "calculated_columns_added": []
    }

    # Si es CSV, convertir a XLSX limpio
    if not file_bytes.startswith(b'PK\x03\x04'):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="latin-1")
        
        # Limpiar columnas de texto
        for col in df.columns:
            if df[col].dtype == object:
                # Intentar convertir a numérico si tiene signos $ o %
                sample = df[col].dropna().astype(str).head(10)
                if any("$" in s or "%" in s for s in sample):
                    df[col] = df[col].apply(lambda x: to_clean_float(x) if x is not None else x)

        out_buf = io.BytesIO()
        df.to_excel(out_buf, sheet_name="Base_Datos", index=False)
        out_buf.seek(0)
        report["sheets_processed"].append("Base_Datos (desde CSV)")
        return out_buf.getvalue(), report

    # Abrir el workbook
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)

    for sname in wb.sheetnames:
        ws = wb[sname]
        if ws.max_row is None or ws.max_row < 2:
            continue

        report["sheets_processed"].append(sname)
        total_rows_sheet = []

        # 1. Encontrar la fila de cabeceras (primera fila con >= 2 celdas no vacías)
        header_row_idx = 1
        headers: List[str] = []
        col_map: Dict[str, int] = {}  # nombre_normalizado -> col_idx (1-based)

        for r in range(1, min(ws.max_row + 1, 15)):
            row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, min(ws.max_column + 1, 30))]
            non_empty = [v for v in row_vals if v]
            if len(non_empty) >= 2:
                header_row_idx = r
                headers = row_vals
                for c_idx, h_text in enumerate(row_vals, start=1):
                    if h_text:
                        col_map[h_text.lower()] = c_idx
                break

        if not headers:
            continue

        # 2. Identificar columnas clave por similitud semántica
        def find_col_idx(keywords: List[str]) -> int | None:
            for kw in keywords:
                for col_name, idx in col_map.items():
                    if kw in col_name:
                        return idx
            return None

        idx_ventas = find_col_idx(["monto venta", "venta", "ventas", "revenue", "sales", "ingreso"])
        idx_costo = find_col_idx(["costo", "cost", "gasto"])
        idx_ganancia = find_col_idx(["ganancia", "utilidad", "profit", "beneficio"])
        idx_margen = find_col_idx(["margen", "% margen", "margin"])

        # Si existen Ventas y Costo pero NO existe columna de Margen, crearla automáticamente
        if idx_ventas and idx_costo and not idx_margen:
            idx_margen = ws.max_column + 1
            header_cell = ws.cell(header_row_idx, idx_margen)
            header_cell.value = "% Margen de Ganancia"
            # Copiar estilo del encabezado anterior si existe
            prev_h = ws.cell(header_row_idx, idx_margen - 1)
            if prev_h.fill and prev_h.fill.fill_type:
                header_cell.fill = openpyxl.styles.PatternFill(
                    fill_type=prev_h.fill.fill_type,
                    start_color=getattr(prev_h.fill.start_color, 'rgb', '001F497D') or '001F497D',
                    end_color=getattr(prev_h.fill.end_color, 'rgb', '001F497D') or '001F497D'
                )
            if prev_h.font:
                header_cell.font = openpyxl.styles.Font(
                    name=prev_h.font.name,
                    size=prev_h.font.size,
                    bold=True,
                    color=getattr(prev_h.font.color, 'rgb', 'FFFFFFFF') or 'FFFFFFFF'
                )
            report["calculated_columns_added"].append(f"{sname}: % Margen de Ganancia")

        # 3. Recorrer filas de datos (después de header_row_idx)
        sum_ventas_sheet = 0.0
        sum_ganancia_sheet = 0.0

        for r in range(header_row_idx + 1, ws.max_row + 1):
            # Verificar si la fila es de Totales/Resumen
            row_is_total = False
            for c in range(1, ws.max_column + 1):
                cell_val = ws.cell(r, c).value
                if is_total_row_text(cell_val):
                    row_is_total = True
                    break

            if row_is_total:
                total_rows_sheet.append(r)
                continue  # No tocar filas de totales generales en el bucle principal

            # Limpiar valores monetarios y numéricos en texto a floats reales
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                c_val = cell.value
                if isinstance(c_val, str) and not str(c_val).startswith("="):
                    cleaned_f = to_clean_float(c_val)
                    if cleaned_f is not None and any(symbol in str(c_val) for symbol in ["$", "€", ","]):
                        cell.value = cleaned_f
                        report["strings_numeric_converted"] += 1

            # 4. Resolver Fórmulas de Ganancia y Margen si vienen vacías o con fórmula sin evaluar
            v_val = to_clean_float(ws.cell(r, idx_ventas).value) if idx_ventas else None
            c_val = to_clean_float(ws.cell(r, idx_costo).value) if idx_costo else None

            if v_val is not None and c_val is not None:
                calc_ganancia = round(v_val - c_val, 2)
                calc_margen = round(calc_ganancia / v_val, 4) if v_val > 0 else 0.0
                sum_ventas_sheet += v_val
                sum_ganancia_sheet += calc_ganancia

                if idx_ganancia:
                    cell_g = ws.cell(r, idx_ganancia)
                    # Si tiene fórmula (=E-F) o está vacío (None), estampar el número real
                    if cell_g.value is None or str(cell_g.value).startswith("="):
                        cell_g.value = calc_ganancia
                        report["formulas_evaluated"] += 1

                if idx_margen:
                    cell_m = ws.cell(r, idx_margen)
                    if cell_m.value is None or str(cell_m.value).startswith("="):
                        cell_m.value = calc_margen
                        cell_m.number_format = "0.0%"
                        report["formulas_evaluated"] += 1

        # Rellenar fila de totales con el margen global si aplica
        if idx_margen and total_rows_sheet and sum_ventas_sheet > 0:
            for t_row in total_rows_sheet:
                cell_tot_m = ws.cell(t_row, idx_margen)
                cell_tot_m.value = round(sum_ganancia_sheet / sum_ventas_sheet, 4)
                cell_tot_m.number_format = "0.0%"

        if total_rows_sheet:
            report["total_rows_detected"][sname] = total_rows_sheet

    # Guardar en RAM
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue(), report
