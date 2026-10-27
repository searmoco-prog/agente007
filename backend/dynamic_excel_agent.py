import io
import re
import json
import unicodedata
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

def strip_accents(text: str) -> str:
    """Elimina acentos y normaliza a minúsculas para matching perfecto."""
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    return ''.join(
        c for c in unicodedata.normalize('NFD', text.lower())
        if unicodedata.category(c) != 'Mn'
    ).strip()

def _safe_row_text(row: pd.Series) -> str:
    """Convierte todas las celdas de una fila a texto seguro sin fallar por float/NaN."""
    parts = []
    for val in row:
        if pd.notna(val) and val is not None:
            parts.append(str(val))
    return " ".join(parts)

# Paleta corporativa de colores suaves
COLOR_MAP = {
    'azul': {'fill': 'D9E1F2', 'font': '1F4E78', 'nombre': 'Azul suave'},
    'blue': {'fill': 'D9E1F2', 'font': '1F4E78', 'nombre': 'Azul suave'},
    'verde': {'fill': 'E2EFDA', 'font': '375623', 'nombre': 'Verde suave'},
    'green': {'fill': 'E2EFDA', 'font': '375623', 'nombre': 'Verde suave'},
    'amarillo': {'fill': 'FFF2CC', 'font': '7F6000', 'nombre': 'Amarillo suave'},
    'yellow': {'fill': 'FFF2CC', 'font': '7F6000', 'nombre': 'Amarillo suave'},
    'rojo': {'fill': 'FCE4D6', 'font': 'C00000', 'nombre': 'Coral / Rojo suave'},
    'red': {'fill': 'FCE4D6', 'font': 'C00000', 'nombre': 'Coral / Rojo suave'},
    'naranja': {'fill': 'FCE4D6', 'font': 'C65911', 'nombre': 'Naranja suave'},
    'orange': {'fill': 'FCE4D6', 'font': 'C65911', 'nombre': 'Naranja suave'},
    'morado': {'fill': 'E1D5E7', 'font': '674EA7', 'nombre': 'Morado suave'},
    'purple': {'fill': 'E1D5E7', 'font': '674EA7', 'nombre': 'Morado suave'},
    'gris': {'fill': 'F2F2F2', 'font': '595959', 'nombre': 'Gris suave'},
    'gray': {'fill': 'F2F2F2', 'font': '595959', 'nombre': 'Gris suave'}
}

def parse_excel_intent(pregunta: str, df: pd.DataFrame) -> dict:
    """
    Analiza la petición en lenguaje natural para extraer:
    - Entidades a filtrar o resaltar (ej: Ariel, Jesús, APP, etc.)
    - Si se desea filtrar solo esas filas o mantener todas
    - Colores solicitados
    - Fórmulas o concatenaciones deseadas
    """
    p_norm = strip_accents(pregunta)
    
    # 1. Detectar color deseado
    color_detectado = None
    color_info = COLOR_MAP['azul']  # Default elegante
    for c_key, c_val in COLOR_MAP.items():
        if f'en {c_key}' in p_norm or f'de {c_key}' in p_norm or f'color {c_key}' in p_norm or f'pintalo {c_key}' in p_norm or f'pinta {c_key}' in p_norm:
            color_detectado = c_key
            color_info = c_val
            break

    # 2. Detectar si quiere SOLO las filas filtradas
    es_filtro_exclusivo = any(f in p_norm for f in [
        'solo ', 'solamente ', 'unicamente ', 'extrae solo', 'extraer solo', 
        'filtra solo', 'filtrar solo', 'solo los de', 'unicos'
    ])

    # 3. Detectar palabras clave / nombres objetivo
    # Revisamos valores frecuentes en 'Acción' y 'Escenario'
    candidatos_entidad = []
    
    # Nombres de personas comunes en la base
    nombres_clave = ['ariel', 'jesus', 'regularizacion', 'baja', 'soporte']
    for n in nombres_clave:
        if n in p_norm:
            candidatos_entidad.append(n)

    # Escenarios comunes
    escenarios_clave = {'app': 'APP', 'portabilidad': 'PORTABILIDAD', 'activacion': 'ACTIVACIÓN', 'cambio nir': 'CAMBIO NIR'}
    for esc_k, esc_v in escenarios_clave.items():
        if esc_k in p_norm:
            candidatos_entidad.append(esc_k)

    # Si no detectó nombres conocidos, buscar palabras que sigan a "dicen", "de", "con"
    if not candidatos_entidad:
        m = re.search(r'(?:dicen|tienen|de|para|nombre|persona|accion)\s+([a-zA-Z0-9_-]+)', p_norm)
        if m:
            candidatos_entidad.append(m.group(1))

    # 4. Concatenación
    quiere_concatenar = any(w in p_norm for w in ['concatena', 'concatenar', 'une', 'unir', 'junta', 'juntar', 'formula'])
    
    return {
        'entidades': candidatos_entidad,
        'filtro_exclusivo': es_filtro_exclusivo,
        'color_info': color_info,
        'quiere_concatenar': True if (quiere_concatenar or True) else True, # Siempre por default agregar la fórmula
        'pregunta_original': pregunta
    }

def generate_dynamic_custom_excel(pregunta: str, df: pd.DataFrame) -> tuple:
    """
    Construye un libro Excel personalizado según lenguaje natural:
    - Filtra filas si se pidió exclusividad (ej. solo Ariel).
    - Aplica color específico a celdas o filas correspondientes.
    - Inserta fórmula viva de concatenación: =A{row} & " - " & B{row}
    - Inserta Dashboard Ejecutivo con KPIs dinámicos del subconjunto.
    - Protege el Google Sheet original manteniéndolo intacto.
    
    Retorna: (excel_bytes, filename, mensaje_explicativo)
    """
    intent = parse_excel_intent(pregunta, df)
    entidades = intent['entidades']
    filtro_exclusivo = intent['filtro_exclusivo']
    color_info = intent['color_info']
    
    # Filtrar o marcar en el DataFrame
    df_work = df.copy()
    
    # Crear máscara de coincidencia
    if entidades:
        mask = pd.Series(False, index=df_work.index)
        for ent in entidades:
            ent_norm = strip_accents(ent)
            mask_ent = df_work.apply(lambda r: ent_norm in strip_accents(_safe_row_text(r)), axis=1)
            mask = mask | mask_ent
    else:
        # Si no especificó nadie en particular, se toman todos
        mask = pd.Series(True, index=df_work.index)

    total_coincidencias = int(mask.sum())
    
    if filtro_exclusivo and total_coincidencias > 0:
        df_target = df_work[mask].copy()
        modo_reporte = f"Filtro Exclusivo: {', '.join(entidades).upper()}"
        filename = f"Reporte_{'_'.join(entidades).capitalize()}_007.xlsx"
    else:
        df_target = df_work.copy()
        modo_reporte = f"Base Completa con Resaltado: {', '.join(entidades).upper()}" if entidades else "Base Completa Personalizada"
        filename = "Reporte_Personalizado_007.xlsx"

    wb = openpyxl.Workbook()
    
    # Estilos
    font_titulo = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Calibri", size=10, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sub_navy = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    
    custom_fill = PatternFill(start_color=color_info['fill'], end_color=color_info['fill'], fill_type="solid")
    custom_font = Font(name="Calibri", size=11, bold=True, color=color_info['font'])
    
    thin_border_side = Side(style='thin', color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    total_rows = len(df_target)

    # -------------------------------------------------------------
    # HOJA 2: Datos_Operativos (Filtrados / Procesados)
    # -------------------------------------------------------------
    ws_data = wb.create_sheet(title="Datos_Operativos")
    ws_data.views.sheetView[0].showGridLines = True

    headers = [
        "Ticket", 
        "REQ", 
        "Folio Concatenado (Fórmula)", 
        "Escenario Normalizado", 
        "Acción Requerida", 
        "Estatus Operativo"
    ]
    ws_data.append(headers)

    for col_idx in range(1, len(headers) + 1):
        c = ws_data.cell(row=1, column=col_idx)
        c.font = font_header
        c.fill = fill_navy
        c.alignment = align_center
    ws_data.row_dimensions[1].height = 26

    # Insertar filas
    for row_idx, row in enumerate(df_target.itertuples(), start=2):
        ticket_val = str(row.Ticket) if pd.notna(row.Ticket) else ""
        req_val = str(row.REQ) if pd.notna(row.REQ) else ""
        formula_concat = f'=A{row_idx} & " - " & B{row_idx}'
        esc_val = str(row.Escenario_Limpio)
        acc_val = str(row.Accion_Limpia)
        stat_val = str(getattr(row, 'status_criterio', getattr(row, 'status_normalizado', 'RESUELTO')))

        ws_data.append([ticket_val, req_val, formula_concat, esc_val, acc_val, stat_val])

        ws_data.cell(row=row_idx, column=1).alignment = align_center
        ws_data.cell(row=row_idx, column=2).alignment = align_center
        ws_data.cell(row=row_idx, column=3).alignment = align_center
        ws_data.cell(row=row_idx, column=4).alignment = align_left
        ws_data.cell(row=row_idx, column=5).alignment = align_left
        
        stat_cell = ws_data.cell(row=row_idx, column=6)
        stat_cell.alignment = align_center
        stat_cell.font = font_bold

        # Verificar si esta fila coincide con la entidad buscada para pintarla del color pedido
        row_str = strip_accents(f"{ticket_val} {req_val} {esc_val} {acc_val} {stat_val}")
        es_coincidencia = any(strip_accents(ent) in row_str for ent in entidades) if entidades else True

        for c_idx in range(1, 7):
            cell_item = ws_data.cell(row=row_idx, column=c_idx)
            cell_item.border = thin_border
            if es_coincidencia:
                cell_item.fill = custom_fill
                if c_idx in [3, 5, 6]:
                    cell_item.font = custom_font

    # Auto-ajustar anchos
    for col in ws_data.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws_data.auto_filter.ref = f"A1:F{total_rows + 1}"

    # -------------------------------------------------------------
    # HOJA 1: Dashboard_Personalizado
    # -------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "Dashboard_Personalizado"
    ws_dash.views.sheetView[0].showGridLines = True

    # Banner
    ws_dash.merge_cells("A1:G2")
    banner = ws_dash["A1"]
    banner.value = f"  AGENTE 007 — {modo_reporte.upper()}"
    banner.font = font_titulo
    banner.fill = fill_navy
    banner.alignment = Alignment(horizontal="left", vertical="center")

    ws_dash["A3"] = f"Reporte Generado a Medida por IA | {color_info['nombre']} | Fórmulas Vivas | Fuente Original Protegida"
    ws_dash["A3"].font = font_subtitulo

    # Tabla KPI
    ws_dash["A5"] = "Métrica del Reporte"
    ws_dash["B5"] = "Valor"
    ws_dash["C5"] = "% del Total Global"
    for col in ["A", "B", "C"]:
        ws_dash[f"{col}5"].font = font_header
        ws_dash[f"{col}5"].fill = fill_sub_navy
        ws_dash[f"{col}5"].alignment = align_center

    ws_dash["A6"] = "Tickets en este Reporte"
    ws_dash["B6"] = f"=COUNTA(Datos_Operativos!A2:A{total_rows + 1})"
    ws_dash["C6"] = f"={total_rows}/{len(df)}"

    ws_dash["A7"] = "Coincidencias Clave"
    ws_dash["B7"] = total_coincidencias
    ws_dash["C7"] = f"={total_coincidencias}/{len(df)}"

    ws_dash["A8"] = "Base Global de Tickets"
    ws_dash["B8"] = len(df)
    ws_dash["C8"] = "100.0%"

    for r in range(6, 9):
        ws_dash[f"A{r}"].font = font_bold
        ws_dash[f"B{r}"].font = font_bold
        ws_dash[f"B{r}"].alignment = align_right
        ws_dash[f"C{r}"].alignment = align_right
        ws_dash[f"C{r}"].number_format = "0.0%"
        for c in ["A", "B", "C"]:
            ws_dash[f"{c}{r}"].border = thin_border

    # Fila de color personalizado en KPI
    ws_dash["A6"].fill = custom_fill
    ws_dash["B6"].fill = custom_fill
    ws_dash["C6"].fill = custom_fill

    # Tabla Desglose por Escenario en este reporte
    ws_dash["E5"] = "Escenario"
    ws_dash["F5"] = "Tickets"
    ws_dash["E5"].font = font_header
    ws_dash["E5"].fill = fill_sub_navy
    ws_dash["F5"].font = font_header
    ws_dash["F5"].fill = fill_sub_navy
    ws_dash["F5"].alignment = align_center

    top_esc = df_target['Escenario_Limpio'].value_counts()
    r_esc = 6
    for esc_name, count_val in top_esc.items():
        ws_dash[f"E{r_esc}"] = esc_name
        ws_dash[f"F{r_esc}"] = int(count_val)
        ws_dash[f"E{r_esc}"].font = font_regular
        ws_dash[f"F{r_esc}"].font = font_bold
        ws_dash[f"F{r_esc}"].alignment = align_right
        ws_dash[f"E{r_esc}"].border = thin_border
        ws_dash[f"F{r_esc}"].border = thin_border
        r_esc += 1

    # Ajustar anchos Dashboard
    ws_dash.column_dimensions["A"].width = 28
    ws_dash.column_dimensions["B"].width = 16
    ws_dash.column_dimensions["C"].width = 18
    ws_dash.column_dimensions["D"].width = 5
    ws_dash.column_dimensions["E"].width = 24
    ws_dash.column_dimensions["F"].width = 14

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    excel_bytes = output.getvalue()

    # Mensaje humano para Telegram
    entidades_texto = ", ".join(entidades).upper() if entidades else "solicitados"
    resumen_humano = (
        f"¡Listo jefe! 🫡 Le preparé su reporte a medida:\n\n"
        f"• 📋 **Filtro:** {modo_reporte}\n"
        f"• 🎯 **Registros incluidos:** {total_rows} tickets (de {len(df)} globales)\n"
        f"• 🎨 **Formato:** Celdas resaltadas en **{color_info['nombre']}**\n"
        f"• ⚡ **Fórmula viva aplicada:** Columna C con `=A# & \" - \" & B#`\n"
        f"• 🔒 **Seguridad:** Su Google Sheet original quedó 100% intacto y protegido.\n\n"
        f"Le adjunto el archivo listo para descargar y presentar. 👇"
    )

    return excel_bytes, filename, resumen_humano
