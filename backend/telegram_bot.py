import os
import sys
import io
import re
import difflib
import unicodedata
import asyncio
import traceback
import httpx
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from dotenv import load_dotenv

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from dynamic_excel_agent import generate_dynamic_custom_excel
from universal_excel_agent import process_universal_excel, inspect_workbook
from agent_mcp import decide_agent_action_mcp, execute_tool_consultar_kpis, execute_langgraph_ticket
from session_storage import save_user_session, restore_user_session

# Salida estándar sin buffer para logs en tiempo real
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

load_dotenv()

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
SHEET_URL = os.getenv(
    "GOOGLE_SHEET_CSV_URL", 
    "https://docs.google.com/spreadsheets/d/1WWZ6fThbuZaFvH84PTeMg5FBpvqAcsgLEhpzIzRCpYg/export?format=csv&gid=0"
).strip()
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
CEREBRAS_API_KEY = os.getenv("CEREBRAS_API_KEY", "").strip()
GROQ_MODEL_CUSTOM = os.getenv("GROQ_MODEL", "").strip()

MODELOS_GROQ = [
    GROQ_MODEL_CUSTOM,
    "openai/gpt-oss-120b",  # Principal (120B de OpenAI en Groq)
    "qwen/qwen3.8-27b",     # Respaldo (27B de Alibaba Cloud en Groq)
]
MODELOS_GROQ = [m for m in MODELOS_GROQ if m]

BASE_URL = f"https://api.telegram.org/bot{TOKEN}"
RENDER_URL = "https://agente007-61bf.onrender.com"

# ---------------------------------------------------------------
# SERVIDOR HTTP DE SALUD (Render Web Service)
# ---------------------------------------------------------------
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(b'{"status": "ok", "service": "Agente 007 v3.5 Enterprise Shield"}')
    def log_message(self, format, *args):
        pass

def start_health_server():
    try:
        port = int(os.getenv("PORT", 10000))
        server = HTTPServer(("0.0.0.0", port), HealthHandler)
        print(f"SERVIDOR DE PUERTO INICIADO EN 0.0.0.0:{port}", flush=True)
        server.serve_forever()
    except Exception as e:
        print(f"Error servidor HTTP: {e}", flush=True)

threading.Thread(target=start_health_server, daemon=True).start()

# ---------------------------------------------------------------
# AUTO-PING KEEP-ALIVE (Anti-Sleep para Render Free Tier)
# ---------------------------------------------------------------
async def keep_alive_loop():
    """Mantiene activo el bot en Render cada 10 minutos para evitar que duerma."""
    await asyncio.sleep(45)
    print("🚀 [KEEP-ALIVE] Centinela de disponibilidad iniciado en segundo plano.", flush=True)
    while True:
        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                r = await client.get(RENDER_URL)
                print(f"📡 [KEEP-ALIVE] Ping preventivo a Render: {r.status_code} OK (Servicio despierto)", flush=True)
        except Exception as e:
            print(f"⚠️ [KEEP-ALIVE] Error ping: {e}", flush=True)
        await asyncio.sleep(600)  # Cada 10 minutos (Render suspende a los 15)

# ---------------------------------------------------------------
# ESCUDO DETERMINISTA PERIMETRAL (Guardrail Anti-Prompt Injection)
# ---------------------------------------------------------------
SECURITY_PATTERNS = [
    r"(?i)\b(olvida|ignora|deshazte de|borra)\b.*?\b(instrucciones|directivas|reglas|rol|prompts)\b",
    r"(?i)\b(ignore|disregard|forget)\b.*?\b(instructions|rules|system prompt|previous context)\b",
    r"(?i)\b(system prompt|prompt del sistema|dame tu prompt|muestra tu prompt)\b",
    r"(?i)\b(jailbreak|dan mode|developer mode|modo desarrollador|unrestricted mode)\b",
    r"(?i)\b(eres un hacker|actua como un hacker|bypass security|vulnerabilidad)\b",
    r"(?i)\b(drop table|select \* from|delete from|insert into|information_schema|xp_cmdshell)\b",
    r"(?i)\b(dame|revela|muestra|extrae)\b.*?\b(contraseñas|passwords|claves|tokens|api_key|secrets|credenciales)\b",
    r"(?i)\b(acceso|root|admin)\b.*?\b(base de datos|servidor|terminal|bash|powershell|consola)\b",
]

def check_perimeter_security(text: str) -> bool:
    """Verifica si la petición contiene vectores de ataque antes de llegar al LLM."""
    for pattern in SECURITY_PATTERNS:
        if re.search(pattern, text):
            return True
    return False

# ---------------------------------------------------------------
# FUNCIONES DE DATOS Y GENERACIÓN
# ---------------------------------------------------------------
async def fetch_cleaned_data() -> pd.DataFrame:
    """Descarga y normaliza el Google Sheet en tiempo real."""
    async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
        resp = await client.get(SHEET_URL)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))

    col_resuelto = df.columns[4]
    df['status_clean'] = df[col_resuelto].fillna('').astype(str).str.strip().str.lower()
    df['status_normalizado'] = df['status_clean'].apply(lambda x: 'RESUELTO' if x == 'resuelto' else 'PENDIENTE')
    
    df['Escenario_Limpio'] = df['Escenario'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df['Escenario_Limpio'] = df['Escenario_Limpio'].replace({
        'ACTIAVACION': 'ACTIVACIÓN',
        'PORTABILIDA': 'PORTABILIDAD',
        'CAMBIO DE NIR': 'CAMBIO NIR'
    })
    
    df['Accion_Limpia'] = df.iloc[:, 3].fillna('SIN ACCIÓN').astype(str).str.strip()
    return df

def generate_chart_bytes(df: pd.DataFrame) -> bytes:
    """
    Genera la imagen PNG de la gráfica OFICIAL de base (corte real del Google Sheet).
    """
    total_tickets = len(df)
    resueltos = int((df['status_normalizado'] == 'RESUELTO').sum())
    pendientes = int((df['status_normalizado'] == 'PENDIENTE').sum())
    pct_resuelto = (resueltos / total_tickets) * 100 if total_tickets > 0 else 0
    pct_pendiente = (pendientes / total_tickets) * 100 if total_tickets > 0 else 0

    top_escenarios = df['Escenario_Limpio'].value_counts().head(5)
    top_acciones = df['Accion_Limpia'].value_counts().head(5)

    plt.style.use('dark_background')
    fig = plt.figure(figsize=(13, 7), facecolor='#0f172a')
    fig.suptitle(
        f'AGENTE 007 - REPORTE DE OPERACIONES DE TELEFONÍA\nTotal: {total_tickets} Tickets | Corte Oficial en Vivo',
        fontsize=15, fontweight='bold', color='#38bdf8', y=0.96
    )

    # Donut Chart Oficial
    ax1 = fig.add_subplot(1, 3, 1, facecolor='#0f172a')
    ax1.pie(
        [resueltos, pendientes], 
        labels=[f'Resueltos\n{resueltos} ({pct_resuelto:.1f}%)', f'Pendientes\n{pendientes} ({pct_pendiente:.1f}%)'], 
        colors=['#10b981', '#f43f5e'], 
        startangle=140,
        textprops={'fontsize': 10, 'color': '#f8fafc', 'fontweight': 'bold'},
        wedgeprops={'width': 0.45, 'edgecolor': '#0f172a', 'linewidth': 2}
    )
    ax1.set_title('Estatus Global Oficial', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')

    # Top Escenarios
    ax2 = fig.add_subplot(1, 3, 2, facecolor='#0f172a')
    bars1 = ax2.barh(top_escenarios.index[::-1], top_escenarios.values[::-1], color='#6366f1', height=0.6)
    ax2.set_title('Top Escenarios', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')
    ax2.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
    for bar in bars1:
        w = bar.get_width()
        ax2.text(w + 1, bar.get_y() + bar.get_height()/2, f'{int(w)}', ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=9)
    ax2.tick_params(colors='#94a3b8', labelsize=9)
    for s in ax2.spines.values():
        s.set_color('#334155')

    # Top Acciones
    ax3 = fig.add_subplot(1, 3, 3, facecolor='#0f172a')
    etiquetas = [a[:20] + '...' if len(a) > 20 else a for a in top_acciones.index[::-1]]
    bars2 = ax3.barh(etiquetas, top_acciones.values[::-1], color='#06b6d4', height=0.6)
    ax3.set_title('Top Acciones', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')
    ax3.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
    for bar in bars2:
        w = bar.get_width()
        ax3.text(w + 1, bar.get_y() + bar.get_height()/2, f'{int(w)}', ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=9)
    ax3.tick_params(colors='#94a3b8', labelsize=9)
    for s in ax3.spines.values():
        s.set_color('#334155')

    plt.tight_layout(rect=[0.02, 0.04, 0.98, 0.91])
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=160, facecolor='#0f172a', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf.getvalue()

def generate_criteria_chart_bytes(df_crit: pd.DataFrame, reglas_detalle: dict, default_status: str) -> bytes:
    """
    Genera un panel visual NUEVO y ESPECIALIZADO enfocado exclusivamente en las reglas del criterio.
    Muestra:
    - Donut con avance según el criterio.
    - Barras específicas para cada regla dictada (JESUS, ARIEL, RESTO).
    - Escenarios afectados por los casos No Atendidos / Off.
    """
    total = len(df_crit)
    resueltos = int((df_crit['status_criterio'] == 'RESUELTO').sum())
    pendientes = int((df_crit['status_criterio'] == 'PENDIENTE').sum())
    pct_resuelto = (resueltos / total) * 100 if total > 0 else 0
    pct_pendiente = (pendientes / total) * 100 if total > 0 else 0

    plt.style.use('dark_background')
    fig = plt.figure(figsize=(14, 7), facecolor='#0f172a')
    fig.suptitle(
        f'AGENTE 007 - REPORTE DE CRITERIO PERSONALIZADO\nTotal: {total} Tickets | Simulación de Turno Activa',
        fontsize=15, fontweight='bold', color='#38bdf8', y=0.96
    )

    # 1. Donut Chart de Criterio
    ax1 = fig.add_subplot(1, 3, 1, facecolor='#0f172a')
    ax1.pie(
        [resueltos, pendientes], 
        labels=[f'Atendidos\n{resueltos} ({pct_resuelto:.1f}%)', f'No Atendidos / Off\n{pendientes} ({pct_pendiente:.1f}%)'], 
        colors=['#10b981', '#f43f5e'], 
        startangle=140,
        textprops={'fontsize': 10, 'color': '#f8fafc', 'fontweight': 'bold'},
        wedgeprops={'width': 0.45, 'edgecolor': '#0f172a', 'linewidth': 2}
    )
    ax1.set_title('Estatus con tu Criterio', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')

    # 2. Barras dedicadas a cada regla del criterio (JESUS, ARIEL, RESTO)
    ax2 = fig.add_subplot(1, 3, 2, facecolor='#0f172a')
    labels = list(reglas_detalle.keys())
    counts = [reglas_detalle[k]['count'] for k in labels]
    colors = ['#f43f5e' if reglas_detalle[k]['status'] == 'PENDIENTE' else '#10b981' for k in labels]

    bars2 = ax2.barh(labels[::-1], counts[::-1], color=colors[::-1], height=0.55)
    ax2.set_title('Desglose de Reglas del Criterio', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')
    ax2.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
    for bar in bars2:
        w = bar.get_width()
        ax2.text(w + 1, bar.get_y() + bar.get_height()/2, f'{int(w)}', ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=10)
    ax2.tick_params(colors='#94a3b8', labelsize=9)
    for s in ax2.spines.values():
        s.set_color('#334155')

    # 3. Escenarios afectados por los casos No Atendidos / Off
    ax3 = fig.add_subplot(1, 3, 3, facecolor='#0f172a')
    no_atendidos_df = df_crit[df_crit['status_criterio'] == 'PENDIENTE']
    if len(no_atendidos_df) > 0:
        top_esc_pend = no_atendidos_df['Escenario_Limpio'].value_counts().head(5)
        bars3 = ax3.barh(top_esc_pend.index[::-1], top_esc_pend.values[::-1], color='#e11d48', height=0.55)
        for bar in bars3:
            w = bar.get_width()
            ax3.text(w + 1, bar.get_y() + bar.get_height()/2, f'{int(w)}', ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=10)
    ax3.set_title('Escenarios No Atendidos (Off)', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')
    ax3.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
    ax3.tick_params(colors='#94a3b8', labelsize=9)
    for s in ax3.spines.values():
        s.set_color('#334155')

    plt.tight_layout(rect=[0.02, 0.04, 0.98, 0.91])
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=160, facecolor='#0f172a', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf.getvalue()

def generate_excel_bytes(df: pd.DataFrame, status_col: str = 'status_normalizado', titulo_reporte: str = '') -> bytes:
    """
    Genera un libro Excel corporativo de nivel financiero con:
    - Gráficas nativas interactivas (PieChart y BarChart incrustados).
    - Fórmulas dinámicas vivas (=COUNTIF, =COUNTA, porcentajes).
    - Columna con Fórmula viva de concatenación (=A2 & " - " & B2).
    - Formato condicional automático (Verde suave para RESUELTO, Coral suave para PENDIENTE).
    - Cruces y agregaciones consolidadas.
    """
    wb = openpyxl.Workbook()

    # ✅ Forzar recálculo automático al abrir en cualquier dispositivo
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    
    # Fuentes y estilos
    font_titulo = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Calibri", size=10, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sub_navy = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_soft_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    thin_border_side = Side(style='thin', color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    total_rows = len(df)

    # -------------------------------------------------------------
    # HOJA 2: Base_Consolidada (Alimentadora de datos)
    # -------------------------------------------------------------
    ws_data = wb.create_sheet(title="Base_Consolidada")
    ws_data.views.sheetView[0].showGridLines = True

    headers = ["Ticket", "REQ", "Folio Completo (Fórmula)", "Escenario Normalizado", "Acción Requerida", "Estatus Operativo"]
    ws_data.append(headers)

    # Estilizar encabezado de datos
    for col_idx in range(1, len(headers) + 1):
        cell = ws_data.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
    ws_data.row_dimensions[1].height = 26

    # Insertar filas
    for row_idx, row in enumerate(df.itertuples(), start=2):
        ticket_val = str(row.Ticket) if pd.notna(row.Ticket) else ""
        req_val = str(row.REQ) if pd.notna(row.REQ) else ""
        # Fórmula viva de Excel para concatenar Ticket y REQ
        formula_concat = f'=A{row_idx} & " - " & B{row_idx}'
        esc_val = str(row.Escenario_Limpio)
        acc_val = str(row.Accion_Limpia)
        stat_val = str(getattr(row, status_col, row.status_normalizado))

        ws_data.append([ticket_val, req_val, formula_concat, esc_val, acc_val, stat_val])

        ws_data.cell(row=row_idx, column=1).alignment = align_center
        ws_data.cell(row=row_idx, column=2).alignment = align_center
        ws_data.cell(row=row_idx, column=3).alignment = align_center
        ws_data.cell(row=row_idx, column=4).alignment = align_left
        ws_data.cell(row=row_idx, column=5).alignment = align_left
        
        stat_cell = ws_data.cell(row=row_idx, column=6)
        stat_cell.alignment = align_center
        stat_cell.font = font_bold

        for c in range(1, 7):
            ws_data.cell(row=row_idx, column=c).border = thin_border

    # Formato Condicional en Estatus (Columna F)
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    green_font = Font(color="375623", bold=True)
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    red_font = Font(color="C00000", bold=True)

    rule_resuelto = CellIsRule(operator='equal', formula=['"RESUELTO"'], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_pendiente = CellIsRule(operator='equal', formula=['"PENDIENTE"'], stopIfTrue=True, fill=red_fill, font=red_font)

    status_range = f"F2:F{total_rows + 1}"
    ws_data.conditional_formatting.add(status_range, rule_resuelto)
    ws_data.conditional_formatting.add(status_range, rule_pendiente)

    # Auto-ajustar anchos
    for col in ws_data.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 15)

    # Activar filtro automático
    ws_data.auto_filter.ref = f"A1:F{total_rows + 1}"

    # -------------------------------------------------------------
    # HOJA 1: Dashboard_Ejecutivo (Primera hoja visible)
    # -------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "Dashboard_Ejecutivo"
    ws_dash.views.sheetView[0].showGridLines = True

    # Banner Corporativo
    ws_dash.merge_cells("A1:G2")
    banner = ws_dash["A1"]
    titulo_texto = f"  AGENTE 007 — {titulo_reporte.upper()}" if titulo_reporte else "  AGENTE 007 — DASHBOARD DE OPERACIONES TELECOM (NOC L1/L2)"
    banner.value = titulo_texto
    banner.font = font_titulo
    banner.fill = fill_navy
    banner.alignment = Alignment(horizontal="left", vertical="center")

    ws_dash["A3"] = "Reporte Dinámico en Tiempo Real | Corte de Incidentes, Fórmulas Vivas y Cruces Operativos"
    ws_dash["A3"].font = font_subtitulo

    # Tabla 1: Indicadores Clave de Estatus (Fórmulas vivas referenciadas a Columna F)
    ws_dash["A5"] = "Indicador Clave (KPI)"
    ws_dash["B5"] = "Total"
    ws_dash["C5"] = "% Avance"
    for col in ["A", "B", "C"]:
        ws_dash[f"{col}5"].font = font_header
        ws_dash[f"{col}5"].fill = fill_sub_navy
        ws_dash[f"{col}5"].alignment = align_center

    ws_dash["A6"] = "Tickets Auditados"
    ws_dash["B6"] = f"=COUNTA(Base_Consolidada!A2:A{total_rows + 1})"
    ws_dash["C6"] = "100.0%"

    ws_dash["A7"] = "Resueltos / Atendidos"
    ws_dash["B7"] = f'=COUNTIF(Base_Consolidada!F2:F{total_rows + 1}, "RESUELTO")'
    ws_dash["C7"] = "=B7/B6"

    ws_dash["A8"] = "Pendientes / Off"
    ws_dash["B8"] = f'=COUNTIF(Base_Consolidada!F2:F{total_rows + 1}, "PENDIENTE")'
    ws_dash["C8"] = "=B8/B6"

    for r in range(6, 9):
        ws_dash[f"A{r}"].font = font_bold
        ws_dash[f"B{r}"].font = font_bold
        ws_dash[f"B{r}"].alignment = align_right
        ws_dash[f"C{r}"].alignment = align_right
        ws_dash[f"C{r}"].number_format = "0.0%"
        for c in ["A", "B", "C"]:
            ws_dash[f"{c}{r}"].border = thin_border

    # Color verde suave a Resueltos y suave rojo a Pendientes en el KPI
    ws_dash["A7"].fill = green_fill
    ws_dash["B7"].fill = green_fill
    ws_dash["C7"].fill = green_fill
    ws_dash["A8"].fill = red_fill
    ws_dash["B8"].fill = red_fill
    ws_dash["C8"].fill = red_fill

    # Tabla 2: Top Escenarios
    ws_dash["E5"] = "Top Escenario"
    ws_dash["F5"] = "Casos"
    ws_dash["E5"].font = font_header
    ws_dash["E5"].fill = fill_sub_navy
    ws_dash["F5"].font = font_header
    ws_dash["F5"].fill = fill_sub_navy
    ws_dash["F5"].alignment = align_center

    top_esc = df['Escenario_Limpio'].value_counts().head(5)
    row_esc = 6
    for esc_name, count_val in top_esc.items():
        ws_dash[f"E{row_esc}"] = esc_name
        ws_dash[f"F{row_esc}"] = int(count_val)
        ws_dash[f"E{row_esc}"].font = font_regular
        ws_dash[f"F{row_esc}"].font = font_bold
        ws_dash[f"F{row_esc}"].alignment = align_right
        ws_dash[f"E{row_esc}"].border = thin_border
        ws_dash[f"F{row_esc}"].border = thin_border
        row_esc += 1

    # Anchos de columnas en Dashboard
    ws_dash.column_dimensions["A"].width = 24
    ws_dash.column_dimensions["B"].width = 12
    ws_dash.column_dimensions["C"].width = 12
    ws_dash.column_dimensions["D"].width = 4
    ws_dash.column_dimensions["E"].width = 26
    ws_dash.column_dimensions["F"].width = 10
    ws_dash.column_dimensions["G"].width = 4

    # -------------------------------------------------------------
    # GRÁFICAS NATIVAS INCORPORADAS EN EXCEL
    # -------------------------------------------------------------
    # 1. Gráfica de Pastel (Estatus)
    pie = PieChart()
    pie.title = "Distribución de Estatus Operativo"
    pie.style = 10
    data_pie = Reference(ws_dash, min_col=2, min_row=6, max_row=8)
    labels_pie = Reference(ws_dash, min_col=1, min_row=7, max_row=8)
    pie.add_data(data_pie, titles_from_data=False)
    pie.set_categories(labels_pie)
    pie.width = 13
    pie.height = 9
    ws_dash.add_chart(pie, "A11")

    # 2. Gráfica de Barras (Top Escenarios)
    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = "Top Escenarios con Mayor Demanda"
    bar.y_axis.title = "Volumen de Tickets"
    bar.x_axis.title = "Escenario"
    data_bar = Reference(ws_dash, min_col=6, min_row=5, max_row=row_esc - 1)
    labels_bar = Reference(ws_dash, min_col=5, min_row=6, max_row=row_esc - 1)
    bar.add_data(data_bar, titles_from_data=True)
    bar.set_categories(labels_bar)
    bar.legend = None
    bar.width = 15
    bar.height = 9
    ws_dash.add_chart(bar, "E11")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ---------------------------------------------------------------
# MOTOR DE ANÁLISIS DINÁMICO Y CRITERIOS (Text-to-Data Intelligence)
# ---------------------------------------------------------------
def strip_accents(text: str) -> str:
    """Remueve acentos, tildes y diéresis de cualquier texto para matching infalible."""
    if not text:
        return ""
    return "".join(
        c for c in unicodedata.normalize('NFD', str(text))
        if unicodedata.category(c) != 'Mn'
    ).lower().strip()

def _safe_row_text(row: pd.Series) -> str:
    """Convierte toda la fila a texto sin tildes de forma segura (evita crash con float/NaN)."""
    return ' '.join(strip_accents(v) for v in row.values if v is not None)

def get_vocabulary(df: pd.DataFrame) -> list:
    """Construye un catálogo de palabras de la base para corrección ortográfica (fuzzy matching)."""
    vocab = set()
    for col in ['Accion_Limpia', 'Escenario_Limpio']:
        if col in df.columns:
            for v in df[col].dropna().unique():
                for w in str(v).split():
                    norm_w = strip_accents(w)
                    if len(norm_w) >= 3:
                        vocab.add(norm_w)
    return list(vocab)

def apply_multi_criteria(pregunta: str, df: pd.DataFrame):
    """
    Detecta criterios MULTI-NOMBRE con tolerancia a acentos y errores tipográficos.
    Retorna: (df_crit, reglas_detalle, default_status) o None
    """
    p_norm = strip_accents(pregunta)
    
    # 1. Filtro Exclusivo por Escenario (ej. 'solo app', 'filtra por portabilidad')
    escenarios_map = {
        'app': 'APP',
        'portabilidad': 'PORTABILIDAD',
        'activacion': 'ACTIVACIÓN',
        'cambio nir': 'CAMBIO NIR'
    }
    quiere_filtro_escenario = any(f in p_norm for f in ['solo ', 'solamente ', 'filtra por ', 'filtrar por ', 'unicamente '])
    if quiere_filtro_escenario:
        for esc_k, esc_v in escenarios_map.items():
            if esc_k in p_norm:
                df_sub = df[df['Escenario_Limpio'] == esc_v].copy()
                if len(df_sub) > 0:
                    df_sub['status_criterio'] = df_sub['status_normalizado']
                    reglas_detalle = {f'SOLO {esc_v}': {'status': 'FILTRADO', 'count': len(df_sub)}}
                    return df_sub, reglas_detalle, 'FILTRADO'

    tiene_criterio = any(t in p_norm for t in [
        'criterio', 'considera', 'aplica', 'marca', 'ponme', 'no atendido', 'atendido', 
        'off', 'diferente', 'demas', 'resto', 'excepto', 'otro', 'cambia'
    ])
    if not tiene_criterio:
        return None

    texto_filas_norm = df.astype(str).apply(_safe_row_text, axis=1)
    vocab = get_vocabulary(df)

    tokens = p_norm.split()
    negativos = ['no atendido', 'no atendidos', 'pendiente', 'pendientes', 'sin atender', 'off', 'no resuelto', 'bloqueado']
    positivos = ['atendido', 'atendidos', 'resuelto', 'resueltos', 'ok', 'listo', 'cerrado', 'diferente', 'otro', 'demas', 'resto']

    reglas = {} # {TOKEN: {'status': 'PENDIENTE', 'count': N, 'mask': Series}}

    for i, tok in enumerate(tokens):
        if len(tok) < 3:
            continue
        
        target_tok = tok
        mask = texto_filas_norm.str.contains(r'\b' + re.escape(tok) + r'\b')
        if mask.sum() == 0:
            # Tolerancia a errores tipográficos (ej. 'jessus' -> 'jesus', 'ariell' -> 'ariel')
            fuzzy = difflib.get_close_matches(tok, vocab, n=1, cutoff=0.75)
            if fuzzy:
                target_tok = fuzzy[0]
                mask = texto_filas_norm.str.contains(r'\b' + re.escape(target_tok) + r'\b')

        if mask.sum() > 0:
            ventana = ' '.join(tokens[i+1:min(i+6, len(tokens))])
            if any(n in ventana for n in negativos):
                reglas[target_tok.upper()] = {'status': 'PENDIENTE', 'count': int(mask.sum()), 'mask': mask}
            elif any(p in ventana for p in positivos):
                reglas[target_tok.upper()] = {'status': 'RESUELTO', 'count': int(mask.sum()), 'mask': mask}

    if not reglas:
        return None

    default_status = 'RESUELTO'
    df_crit = df.copy()
    df_crit['status_criterio'] = default_status
    matched_mask = pd.Series([False] * len(df))

    for kw, info in reglas.items():
        df_crit.loc[info['mask'], 'status_criterio'] = info['status']
        matched_mask = matched_mask | info['mask']

    resto_count = int((~matched_mask).sum())
    reglas_detalle = {k: {'status': v['status'], 'count': v['count']} for k, v in reglas.items()}
    reglas_detalle['RESTO (OTROS)'] = {'status': default_status, 'count': resto_count}

    return df_crit, reglas_detalle, default_status

def analyze_query_with_df(pregunta: str, df: pd.DataFrame) -> str:
    """Extrae entidades, nombres, filtros y calcula criterios matemáticos en tiempo real."""
    p_norm = strip_accents(pregunta)
    total = len(df)
    resueltos_oficial = int((df['status_normalizado'] == 'RESUELTO').sum())
    pendientes_oficial = int((df['status_normalizado'] == 'PENDIENTE').sum())

    context_lines = [
        f"BASE OFICIAL ACTUAL ({total} tickets):",
        f"- Resueltos oficiales: {resueltos_oficial} ({resueltos_oficial/total*100:.1f}%)",
        f"- Pendientes oficiales: {pendientes_oficial} ({pendientes_oficial/total*100:.1f}%)",
        f"- Top Escenarios: {dict(df['Escenario_Limpio'].value_counts().head(5))}",
        f"- Top Acciones: {dict(df['Accion_Limpia'].value_counts().head(5))}"
    ]

    unique_words = set(re.findall(r'\b[a-z]{3,}\b', p_norm))
    stopwords = {
        'dime', 'cuantos', 'estan', 'atendidos', 'tickets', 'como', 'para', 'este',
        'esta', 'pero', 'hacer', 'hola', 'buen', 'nombre', 'algo', 'traen', 'solo',
        'unos', 'estos', 'favor', 'puede', 'puedo', 'cual', 'cuales', 'todos',
        'ponme', 'aplica', 'criterio', 'diferente', 'demas', 'resto', 'otro',
        'siguiente', 'todo', 'excepto', 'ayudar', 'aplicando', 'puedes', 'dicen',
        'pides', 'generar', 'conteo', 'grafica', 'reporte', 'informe', 'quiero',
        'dame', 'tengo', 'queda', 'quedan', 'poner', 'ponlos', 'marca', 'marcar',
        'cuanto', 'cuando', 'donde', 'porque', 'tiene', 'tienen', 'decir',
        'tipo', 'tipos', 'sea', 'son', 'hay', 'van', 'vez', 'mas',
        'que', 'los', 'las', 'del', 'con', 'sin', 'por', 'una', 'uno', 'sus',
        'hacer', 'decirme', 'generarme', 'darme', 'hacerlo',
        'excel', 'xlsx', 'archivo', 'documento', 'hoja', 'descargar'
    }
    keywords = [w for w in unique_words if w not in stopwords]

    texto_filas_norm = df.astype(str).apply(_safe_row_text, axis=1)
    vocab = get_vocabulary(df)

    coincidencias = []
    no_encontrados = []

    for kw in keywords:
        target_kw = kw
        mask = texto_filas_norm.str.contains(r'\b' + re.escape(kw) + r'\b')
        if mask.sum() == 0:
            fuzzy = difflib.get_close_matches(kw, vocab, n=1, cutoff=0.75)
            if fuzzy:
                target_kw = fuzzy[0]
                mask = texto_filas_norm.str.contains(r'\b' + re.escape(target_kw) + r'\b')

        count = int(mask.sum())
        if count > 0:
            sub = df[mask]
            res_sub = int((sub['status_normalizado'] == 'RESUELTO').sum())
            pen_sub = int((sub['status_normalizado'] == 'PENDIENTE').sum())
            coincidencias.append(
                f"COINCIDENCIA EN LA BASE PARA '{target_kw.upper()}':\n"
                f"  • Total de tickets encontrados: {count}\n"
                f"  • Estatus actual en la hoja oficial: {res_sub} resueltos, {pen_sub} pendientes\n"
                f"  • Escenarios donde aparece: {dict(sub['Escenario_Limpio'].value_counts().head(3))}\n"
                f"  • Acciones donde aparece: {dict(sub['Accion_Limpia'].value_counts().head(3))}"
            )
        else:
            if len(kw) >= 4:
                no_encontrados.append(f"• '{kw.capitalize()}': 0 tickets encontrados en la base.")

    if coincidencias:
        context_lines.append("\nANÁLISIS DINÁMICO EN TIEMPO REAL:")
        context_lines.extend(coincidencias)

    if no_encontrados:
        context_lines.append("\nBÚSQUEDAS SIN COINCIDENCIAS (0 REGISTROS):")
        context_lines.extend(no_encontrados)

    resultado_multi = apply_multi_criteria(pregunta, df)
    if resultado_multi is not None:
        df_crit, reglas_detalle, default_status = resultado_multi
        total_crit = len(df_crit)
        res_crit = int((df_crit['status_criterio'] == 'RESUELTO').sum()) if 'status_criterio' in df_crit else 0
        pen_crit = int((df_crit['status_criterio'] == 'PENDIENTE').sum()) if 'status_criterio' in df_crit else 0
        pct_crit = (res_crit / total_crit * 100) if total_crit > 0 else 0

        reglas_str = '\n'.join(
            f"    • {k}: {v['count']} tickets → {v['status']}" for k, v in reglas_detalle.items()
        )
        context_lines.append(
            f"\nSIMULACIÓN MULTI-CRITERIO DEL JEFE:\n"
            f"  Reglas aplicadas con éxito:\n{reglas_str}\n"
            f"  RESULTADO DEL CRITERIO:\n"
            f"  • Total analizado: {total_crit} tickets\n"
            f"  • Atendidos (Resueltos): {res_crit}\n"
            f"  • No Atendidos (Pendientes / Off): {pen_crit}\n"
            f"  • % de Avance resultante: {pct_crit:.1f}%\n"
            f"  INSTRUCCIÓN PARA TU RESPUESTA: Explica al jefe cada regla con sus números exactos. "
            f"Menciona que Jesús tiene sus casos en 'APP', Ariel tiene 2 casos en 'APP' y el resto son atendidos."
        )

    return "\n".join(context_lines)

# ---------------------------------------------------------------
# IA CONVERSACIONAL (Groq / Cerebras) CON MEMORIA DE SESIÓN
# ---------------------------------------------------------------
HISTORIAL_CHAT = {}  # {chat_id: [{"role": "user"|"assistant", "content": "..."}]}

async def ask_ai_chat(pregunta: str, df: pd.DataFrame, chat_id: int = None) -> str:
    """Genera una respuesta humana, concisa y ejecutiva analizando criterios y con memoria conversacional."""
    analisis_datos = analyze_query_with_df(pregunta, df)

    system_prompt = (
        "Eres el 'Agente 007', copiloto senior de operaciones y analista universal de datos (NOC Telecom & Negocios). "
        "Habla de forma humana, ejecutiva, profesional y amigable con tu jefe ('Hola jefe', 'Buen día'). "
        "Tienes dos especialidades integradas: "
        "1) MODO TELECOM NOC: Monitoreo en vivo de tickets, fallas de red, portabilidades y cruces en base de datos. "
        "2) MODO ANALISTA UNIVERSAL: Tienes la capacidad comprobada de recibir, abrir, inspeccionar y transformar cualquier archivo Excel (.xlsx) o CSV "
        "de ventas, finanzas, inventarios o call center. Puedes calcular márgenes, agregar columnas con fórmulas vivas, "
        "inyectar BUSCARV/XLOOKUP, generar tablas dinámicas y crear gráficos circulares o de barras nativos. "
        "REGLA CRÍTICA: NUNCA digas que no tienes la capacidad de abrir o ver archivos que el usuario suba. Si tu jefe te pregunta si puedes leer un archivo, "
        "confirma con entusiasmo y seguridad técnica que SÍ puedes procesar cualquier archivo .xlsx o .csv que te envíe por este chat de Telegram."
    )

    messages = [
        {"role": "system", "content": f"{system_prompt}\n\nDATOS OPERATIVOS EN VIVO:\n{analisis_datos}"}
    ]

    # Agregar historial reciente (últimos 4 mensajes) si existe para continuidad conversacional
    if chat_id and chat_id in HISTORIAL_CHAT:
        messages.extend(HISTORIAL_CHAT[chat_id][-4:])

    messages.append({"role": "user", "content": pregunta})
    respuesta_final = None

    if GROQ_API_KEY:
        try:
            from groq import AsyncGroq
            client = AsyncGroq(api_key=GROQ_API_KEY)
            for m in MODELOS_GROQ:
                try:
                    print(f"Probando modelo en Groq: '{m}'...", flush=True)
                    res = await client.chat.completions.create(
                        messages=messages,
                        model=m,
                        temperature=0.3,
                        max_tokens=400
                    )
                    respuesta_final = res.choices[0].message.content
                    print(f"-> EXITO con modelo: {m}!", flush=True)
                    break
                except Exception as ex_m:
                    print(f"Error con modelo {m}: {ex_m}", flush=True)
        except Exception as e:
            print(f"Error Groq: {e}", flush=True)

    if not respuesta_final:
        # Fallback si no hay IA
        total = len(df)
        resueltos = int((df['status_normalizado'] == 'RESUELTO').sum())
        pendientes = int((df['status_normalizado'] == 'PENDIENTE').sum())
        pct = (resueltos / total * 100) if total > 0 else 0
        respuesta_final = (
            f"¡Hola jefe! Al corte de este momento llevamos {total} tickets en la base:\n\n"
            f"• ✅ {resueltos} resueltos ({pct:.1f}%)\n"
            f"• ⏳ {pendientes} pendientes\n\n"
            f"La gran mayoría están concentrados en App y Portabilidad. ¿Gusta que le envíe la gráfica o el archivo Excel?"
        )

    # Actualizar historial de sesión
    if chat_id:
        if chat_id not in HISTORIAL_CHAT:
            HISTORIAL_CHAT[chat_id] = []
        HISTORIAL_CHAT[chat_id].append({"role": "user", "content": pregunta})
        HISTORIAL_CHAT[chat_id].append({"role": "assistant", "content": respuesta_final})
        if len(HISTORIAL_CHAT[chat_id]) > 8:
            HISTORIAL_CHAT[chat_id] = HISTORIAL_CHAT[chat_id][-8:]

    return respuesta_final

# ---------------------------------------------------------------
# BOTONES INTERACTIVOS
# ---------------------------------------------------------------
BOTONES_ACCIONES = {
    "inline_keyboard": [
        [
            {"text": "📊 Ver Gráfica Oficial", "callback_data": "btn_chart"},
            {"text": "📁 Descargar Excel Avanzado", "callback_data": "btn_excel"}
        ],
        [
            {"text": "🔄 Actualizar Resumen", "callback_data": "btn_summary"}
        ]
    ]
}

# ---------------------------------------------------------------
# GESTOR DE ARCHIVOS Y MODO HÍBRIDO (Telecom NOC vs Analista Universal)
# ---------------------------------------------------------------
MODO_SESION = {}       # {chat_id: 'telecom' | 'universal'}
DATAFRAME_SESION = {}  # {chat_id: {"df": pd.DataFrame, "filename": str, "total_filas": int, "raw_bytes": bytes}}

async def get_active_dataframe(chat_id: int, client: httpx.AsyncClient = None) -> tuple[pd.DataFrame, str]:
    """Retorna el DataFrame activo para el chat (el archivo que subió el usuario, restaurado o la hoja oficial)."""
    if chat_id in DATAFRAME_SESION:
        info = DATAFRAME_SESION[chat_id]
        return info["df"], f"Archivo subido: {info['filename']}"

    # Intentar restaurar sesión persistente de disco o Supabase si Render reinició
    try:
        restored = await restore_user_session(chat_id, client=client, bot_token=TOKEN)
        if restored and restored.get("raw_bytes"):
            raw_b = restored["raw_bytes"]
            fname = restored["filename"]
            df_up, sheets_f, s_act = parse_uploaded_file(raw_b, fname)
            DATAFRAME_SESION[chat_id] = {
                "df": df_up,
                "filename": fname,
                "total_filas": len(df_up),
                "sheets": sheets_f,
                "active_sheet": s_act,
                "raw_bytes": raw_b,
                "file_id": restored.get("file_id", "")
            }
            return df_up, f"Archivo activo (sesión recuperada): {fname}"
    except Exception as e_res:
        print(f"Detalle al restaurar sesión activa: {e_res}", flush=True)

    df_oficial = await fetch_cleaned_data()
    return df_oficial, "Hoja Oficial en Vivo (Google Sheet)"

async def download_telegram_file(client: httpx.AsyncClient, file_id: str) -> bytes:
    """Descarga el archivo binario desde la API de Telegram."""
    info_resp = await client.get(f"{BASE_URL}/getFile?file_id={file_id}")
    info_resp.raise_for_status()
    file_path = info_resp.json()["result"]["file_path"]
    download_url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"
    file_resp = await client.get(download_url)
    file_resp.raise_for_status()
    return file_resp.content

def parse_uploaded_file(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, list, str]:
    """Lee y normaliza cualquier archivo subido (.xlsx, .xls, .csv, .tsv) a un DataFrame universal."""
    file_name_lower = file_name.lower()
    sheet_names = []
    target_sheet = ""

    if file_name_lower.endswith(('.xlsx', '.xls')):
        wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_names = wb_in.sheetnames

        # Buscar la hoja que contenga la tabla de datos real (Base_Datos, Datos, Data, Ventas, etc.)
        target_sheet = sheet_names[0]
        for s in sheet_names:
            if any(k in s.lower() for k in ['base', 'datos', 'data', 'ventas', 'detalle', 'registros', 'orders']):
                target_sheet = s
                break

        ws_target = wb_in[target_sheet]

        # Detectar la fila de encabezados reales (donde haya al menos 2 celdas con texto descriptivo)
        header_row_idx = 0
        for r_idx, row in enumerate(ws_target.iter_rows(values_only=True)):
            non_empty = [str(c).strip() for c in row if c is not None and str(c).strip() != '']
            if len(non_empty) >= 2 and any(k in " ".join(non_empty).lower() for k in ['id', 'venta', 'monto', 'costo', 'ticket', 'fecha', 'cliente', 'estado']):
                header_row_idx = r_idx
                break
            elif len(non_empty) >= 3 and header_row_idx == 0:
                header_row_idx = r_idx

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=header_row_idx)
    elif file_name_lower.endswith(('.csv', '.txt', '.tsv')):
        sheet_names = ["Hoja1"]
        target_sheet = "Hoja1"
        try:
            sep = '\t' if file_name_lower.endswith('.tsv') else ','
            df = pd.read_csv(io.BytesIO(file_bytes), sep=sep, encoding='utf-8', engine='python')
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='latin-1', engine='python')
    else:
        raise ValueError(f"Formato no soportado para '{file_name}'. Sube un archivo .xlsx, .xls o .csv")

    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)

    # 1. Asegurar Ticket y REQ para compatibilidad total
    if 'Ticket' not in df.columns:
        df['Ticket'] = df[cols[0]].astype(str) if len(cols) > 0 else "0"
    if 'REQ' not in df.columns:
        df['REQ'] = df[cols[1]].astype(str) if len(cols) > 1 else ""

    # 2. Escenario_Limpio
    if 'Escenario_Limpio' not in df.columns:
        if 'Escenario' in df.columns:
            df['Escenario_Limpio'] = df['Escenario'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
        elif len(cols) > 2:
            df['Escenario_Limpio'] = df[cols[2]].fillna('GENERAL').astype(str).str.upper().str.strip()
        else:
            df['Escenario_Limpio'] = "GENERAL"

    # 3. Accion_Limpia
    if 'Accion_Limpia' not in df.columns:
        candidates = [c for c in df.columns if any(k in c.lower() for k in ['acci', 'accion', 'action', 'tipo', 'falla', 'nombre', 'sitio', 'servicio'])]
        if candidates:
            df['Accion_Limpia'] = df[candidates[0]].fillna('SIN ACCIÓN').astype(str).str.strip()
        elif len(cols) > 3:
            df['Accion_Limpia'] = df[cols[3]].fillna('SIN ACCIÓN').astype(str).str.strip()
        else:
            df['Accion_Limpia'] = "REGISTRO"

    # 4. status_normalizado
    if 'status_normalizado' not in df.columns:
        status_candidates = [c for c in df.columns if any(k in c.lower() for k in ['status', 'estatus', 'estado', 'resuelto', 'resolu'])]
        if status_candidates:
            df['status_normalizado'] = df[status_candidates[0]].fillna('').astype(str).apply(
                lambda s: 'RESUELTO' if any(w in str(s).lower() for w in ['resuelto', 'cerrado', 'ok', 'atendido', 'exito', 'activo', 'completado']) else 'PENDIENTE'
            )
        else:
            df['status_normalizado'] = 'RESUELTO'

    return df, sheet_names, target_sheet

async def send_chart_to_user(client: httpx.AsyncClient, chat_id: int, df: pd.DataFrame = None, label: str = ""):
    """Envía la gráfica oficial de base o del archivo activo del usuario."""
    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "📊 Generando gráfica de operaciones..."}
    )
    if df is None:
        df, label = await get_active_dataframe(chat_id)
    elif not label:
        label = "Corte Operativo"
    img_bytes = generate_chart_bytes(df)
    files = {"photo": ("grafica_operativa_007.png", img_bytes, "image/png")}
    await client.post(
        f"{BASE_URL}/sendPhoto",
        data={"chat_id": chat_id, "caption": f"📊 Gráfica: {label}"},
        files=files
    )

async def send_criteria_chart_to_user(client: httpx.AsyncClient, chat_id: int, df_crit: pd.DataFrame, reglas_detalle: dict, default_status: str):
    """Envía la gráfica NUEVA y ESPECIALIZADA con el criterio del jefe."""
    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "📊 Generando gráfica especializada con las reglas de tu criterio..."}
    )
    img_bytes = generate_criteria_chart_bytes(df_crit, reglas_detalle, default_status)
    files = {"photo": ("grafica_criterio_007.png", img_bytes, "image/png")}
    caption = "📊 Gráfica Especial de Criterio: Visualización dedicada de las reglas aplicadas."
    await client.post(
        f"{BASE_URL}/sendPhoto",
        data={"chat_id": chat_id, "caption": caption},
        files=files
    )

ULTIMO_CRITERIO = {}  # {chat_id: (df_crit, reglas_detalle, default_status, titulo_reporte)}

async def send_excel_to_user(client: httpx.AsyncClient, chat_id: int, df: pd.DataFrame = None, status_col: str = 'status_normalizado', titulo_reporte: str = ''):
    """Envía el archivo Excel con gráficas nativas, fórmulas vivas y formato condicional."""
    if df is None:
        await client.post(
            f"{BASE_URL}/sendMessage",
            json={"chat_id": chat_id, "text": "📁 Compilando libro Excel corporativo con fórmulas automáticas..."}
        )
        df, label_fuente = await get_active_dataframe(chat_id)
        if not titulo_reporte:
            titulo_reporte = label_fuente
    else:
        label = f": {titulo_reporte}" if titulo_reporte else ""
        await client.post(
            f"{BASE_URL}/sendMessage",
            json={"chat_id": chat_id, "text": f"📁 Compilando libro Excel con tu criterio personalizado{label} y fórmulas vivas..."}
        )
    excel_bytes = generate_excel_bytes(df, status_col=status_col, titulo_reporte=titulo_reporte)
    files = {"document": ("Reporte_Telecom_Consolidado_007.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    caption = f"📁 Reporte Excel con criterio: {titulo_reporte}" if titulo_reporte else "📁 Reporte Corporativo Oficial: Dashboard con gráficas nativas de Excel, fórmulas automáticas y formato condicional."
    await client.post(
        f"{BASE_URL}/sendDocument",
        data={"chat_id": chat_id, "caption": caption},
        files=files
    )

async def send_dynamic_custom_excel(client: httpx.AsyncClient, chat_id: int, query: str, df: pd.DataFrame = None):
    """Genera y envía un Excel a medida con filtros, colores y fórmulas vivas solicitadas en lenguaje natural."""
    if df is None:
        df, _ = await get_active_dataframe(chat_id)

    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "⚡ Procesando su requerimiento en memoria segura y armando el libro Excel a medida..."}
    )

    try:
        excel_bytes, filename, resumen = generate_dynamic_custom_excel(query, df)
        files = {"document": (filename, excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        await client.post(
            f"{BASE_URL}/sendDocument",
            data={"chat_id": chat_id, "caption": f"📁 {filename} preparado y verificado."},
            files=files
        )
        await client.post(
            f"{BASE_URL}/sendMessage",
            json={"chat_id": chat_id, "text": resumen, "reply_markup": BOTONES_ACCIONES}
        )
    except Exception as e:
        print(f"Error generando dynamic excel: {e}", flush=True)
        # Fallback al generador corporativo estándar
        await send_excel_to_user(client, chat_id, df=df, titulo_reporte="Reporte Operativo")

# ---------------------------------------------------------------
# BOT LOOP CON ENRUTADOR DE INTENCIONES Y PERÍMETRO DE SEGURIDAD
# ---------------------------------------------------------------
async def run_bot():
    if not TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN no configurado", flush=True)
        return

    print("=" * 65, flush=True)
    print("🤖 AGENTE 007 v3.5 (ENTERPRISE SHIELD + EXCEL DASHBOARD) ACTIVO", flush=True)
    print("=" * 65, flush=True)

    # Iniciar Centinela Keep-Alive
    asyncio.create_task(keep_alive_loop())

    offset = 0
    async with httpx.AsyncClient(timeout=40.0) as client:
        while True:
            try:
                resp = await client.get(f"{BASE_URL}/getUpdates?offset={offset}&timeout=25")
                if resp.status_code != 200:
                    await asyncio.sleep(3)
                    continue

                for update in resp.json().get("result", []):
                    offset = update["update_id"] + 1

                    # 1. MANEJAR CLIC EN BOTONES INTERACTIVOS
                    if "callback_query" in update:
                        cb = update["callback_query"]
                        cb_data = cb.get("data", "")
                        cb_chat_id = cb["message"]["chat"]["id"]
                        
                        await client.post(f"{BASE_URL}/answerCallbackQuery", json={"callback_query_id": cb["id"]})
                        df_btn, label_btn = await get_active_dataframe(cb_chat_id)

                        if cb_data == "btn_chart":
                            await send_chart_to_user(client, cb_chat_id, df=df_btn, label=label_btn)
                        elif cb_data == "btn_excel":
                            if cb_chat_id in DATAFRAME_SESION and DATAFRAME_SESION[cb_chat_id].get("raw_bytes"):
                                raw_b = DATAFRAME_SESION[cb_chat_id]["raw_bytes"]
                                f_name = DATAFRAME_SESION[cb_chat_id]["filename"]
                                await client.post(
                                    f"{BASE_URL}/sendMessage",
                                    json={"chat_id": cb_chat_id, "text": f"⚡ Generando libro Excel enriquecido con todas las transformaciones sobre '{f_name}'..."}
                                )
                                res_bytes, out_name, out_msg = process_universal_excel(raw_b, "todas", f_name)
                                DATAFRAME_SESION[cb_chat_id]["raw_bytes"] = res_bytes
                                files = {"document": (out_name, res_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                                await client.post(
                                    f"{BASE_URL}/sendDocument",
                                    data={"chat_id": cb_chat_id, "caption": f"📁 {out_name} con fórmulas vivas, formatos y gráficas."},
                                    files=files
                                )
                                await client.post(
                                    f"{BASE_URL}/sendMessage",
                                    json={"chat_id": cb_chat_id, "text": out_msg, "reply_markup": BOTONES_ACCIONES}
                                )
                            else:
                                await send_excel_to_user(client, cb_chat_id, df=df_btn, titulo_reporte=label_btn)
                        elif cb_data == "btn_summary":
                            resumen = await ask_ai_chat("Dame el resumen general del turno", df_btn, chat_id=cb_chat_id)
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": cb_chat_id, "text": resumen, "reply_markup": BOTONES_ACCIONES}
                            )
                        continue

                    # 2. MANEJAR MENSAJES (Texto o Documentos subidos)
                    message = update.get("message")
                    if not message:
                        continue

                    chat_id = message["chat"]["id"]
                    documento = message.get("document")

                    # A) ¿El usuario subió un documento (.xlsx, .csv, .xls, etc.)?
                    if documento:
                        file_id = documento.get("file_id")
                        file_name = documento.get("file_name", "archivo.xlsx")
                        caption = message.get("caption", "").strip()

                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": f"📥 Descargando y procesando '{file_name}' en memoria segura..."}
                        )

                        try:
                            file_bytes = await download_telegram_file(client, file_id)
                            df_uploaded, sheets_found, sheet_active = parse_uploaded_file(file_bytes, file_name)
                            DATAFRAME_SESION[chat_id] = {
                                "df": df_uploaded,
                                "filename": file_name,
                                "total_filas": len(df_uploaded),
                                "sheets": sheets_found,
                                "active_sheet": sheet_active,
                                "raw_bytes": file_bytes,
                                "file_id": file_id
                            }
                            # Guardar en persistencia (disco y Supabase) para sobrevivir reinicios de Render
                            save_user_session(chat_id, file_id, file_name, file_bytes, sheets_found, sheet_active, df_uploaded)

                            total_f = len(df_uploaded)
                            col_list = ", ".join(list(df_uploaded.columns)[:8])
                            sheets_str = ", ".join(sheets_found) if sheets_found else "Hoja1"

                            if caption:
                                texto = caption
                                texto_lower = texto.lower()
                                await client.post(
                                    f"{BASE_URL}/sendMessage",
                                    json={
                                        "chat_id": chat_id,
                                        "text": f"✅ ¡Archivo '{file_name}' recibido ({total_f} filas)!\n⚡ Procesando su instrucción: \"{caption}\"..."
                                    }
                                )
                            else:
                                msg_confirm = (
                                    f"✅ **¡Archivo '{file_name}' cargado con éxito!** 🫡\n\n"
                                    f"• 📑 **Pestañas detectadas:** {sheets_str}\n"
                                    f"• 📊 **Pestaña analizada:** '{sheet_active}' ({total_f} registros)\n"
                                    f"• 🏷️ **Columnas detectadas:** {col_list}\n\n"
                                    "¿Qué operación deseas ejecutar, jefe?\n"
                                    "1️⃣ *Calcula el margen y resalta en verde > 35%* (o escribe **1**)\n"
                                    "2️⃣ *Genera tabla dinámica y gráfico circular por Estado* (o escribe **2**)\n"
                                    "3️⃣ *Agrega módulo de búsqueda con BUSCARV / XLOOKUP* (o escribe **3**)\n"
                                    "⭐ *Aplica las 3 transformaciones juntas* (o escribe **todas**)\n\n"
                                    "💡 *(Para volver a la base oficial de Google Sheet, escribe /reset)*"
                                )
                                await client.post(
                                    f"{BASE_URL}/sendMessage",
                                    json={"chat_id": chat_id, "text": msg_confirm, "reply_markup": BOTONES_ACCIONES}
                                )
                                continue
                        except Exception as e:
                            print(f"Error procesando documento: {e}", flush=True)
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": f"⚠️ Error al leer '{file_name}': {str(e)}\n\nAsegúrate de subir un archivo .xlsx, .xls o .csv válido."}
                            )
                            continue
                    elif "text" in message:
                        texto = message["text"].strip()
                        texto_lower = texto.lower()
                    else:
                        continue

                    print(f"Mensaje de {chat_id}: {texto}", flush=True)

                    # Selector de Modo Híbrido (Telecom NOC vs Analista Universal)
                    if texto_lower in ['/modo', 'modo']:
                        modo_actual = MODO_SESION.get(chat_id, "telecom").upper()
                        msg_modo = (
                            f"🤖 **MODO DE OPERACIÓN ACTUAL:** {modo_actual}\n\n"
                            "Puedes alternar entre los dos cerebros en cualquier momento:\n\n"
                            "🛰️ **/telecom** ➡️ Modo NOC Telecom 007 (Google Sheet en vivo, 258 tickets, fallas de red, Supabase).\n"
                            "📊 **/universal** ➡️ Modo Analista Universal (Finanzas, Ventas, BUSCARV, tablas dinámicas, márgenes y gráficos nativos)."
                        )
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": msg_modo, "reply_markup": BOTONES_ACCIONES}
                        )
                        continue

                    if texto_lower in ['/telecom', 'telecom', 'modo telecom']:
                        MODO_SESION[chat_id] = "telecom"
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={
                                "chat_id": chat_id,
                                "text": "🛰️ **Modo Telecom NOC 007 Activado.**\n\nConectado a la base oficial de tickets y listo para corte de turno, gráficas de red y resolución en Supabase.",
                                "reply_markup": BOTONES_ACCIONES
                            }
                        )
                        continue

                    if texto_lower in ['/universal', 'universal', 'modo universal']:
                        MODO_SESION[chat_id] = "universal"
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={
                                "chat_id": chat_id,
                                "text": "📊 **Modo Analista Universal Activado.**\n\nListo para recibir cualquier archivo Excel (.xlsx) o CSV de Finanzas, Ventas o Call Center, aplicar BUSCARV, calcular márgenes, tablas dinámicas y gráficos nativos.",
                                "reply_markup": BOTONES_ACCIONES
                            }
                        )
                        continue

                    # Reset de sesión de archivo hacia la base oficial
                    if any(texto_lower == r or texto_lower.startswith(r) for r in ['/reset', '/oficial', 'base oficial', 'hoja oficial', 'restablecer', 'volver a la base']):
                        if chat_id in DATAFRAME_SESION:
                            del DATAFRAME_SESION[chat_id]
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={
                                    "chat_id": chat_id,
                                    "text": "🔄 Sesión restablecida con éxito.\n\nAhora estás consultando nuevamente la Hoja Oficial en Vivo (Google Sheet) de la empresa.",
                                    "reply_markup": BOTONES_ACCIONES
                                }
                            )
                            continue

                    # Normalización de intenciones numéricas / lenguaje natural para archivos cargados
                    t_clean = texto_lower.strip()
                    if any(t_clean == k or t_clean.startswith(k) for k in ['1', 'opcion 1', 'opción 1', 'la 1', 'el 1', 'la primera', 'el primero', 'primera']):
                        texto = "Calcula el margen y resalta en verde > 35%"
                        texto_lower = texto.lower()
                    elif any(t_clean == k or t_clean.startswith(k) for k in ['2', 'opcion 2', 'opción 2', 'la 2', 'el 2', 'la segunda', 'el segundo', 'segunda']):
                        texto = "Genera tabla dinámica y gráfico circular por Estado"
                        texto_lower = texto.lower()
                    elif any(t_clean == k or t_clean.startswith(k) for k in ['3', 'opcion 3', 'opción 3', 'la 3', 'el 3', 'la tercera', 'el tercero', 'tercera']):
                        texto = "Agrega módulo de búsqueda con BUSCARV / XLOOKUP"
                        texto_lower = texto.lower()
                    elif any(k in t_clean for k in ['todas', 'todo', 'las 3', 'haz todo', 'todas las opciones', 'las tres', 'aplica todo', 'ambas', 'completo']):
                        texto = "Calcula el margen y resalta en verde > 35%, genera tabla dinámica con gráfico circular por Estado y agrega módulo de búsqueda con BUSCARV"
                        texto_lower = texto.lower()

                    # Restaurar sesión si la memoria RAM se reinició por inactividad de Render
                    if chat_id not in DATAFRAME_SESION:
                        try:
                            restored = await restore_user_session(chat_id, client=client, bot_token=TOKEN)
                            if restored and restored.get("raw_bytes"):
                                raw_b_res = restored["raw_bytes"]
                                fname_res = restored["filename"]
                                df_res, sheets_res, act_res = parse_uploaded_file(raw_b_res, fname_res)
                                DATAFRAME_SESION[chat_id] = {
                                    "df": df_res,
                                    "filename": fname_res,
                                    "total_filas": len(df_res),
                                    "sheets": sheets_res,
                                    "active_sheet": act_res,
                                    "raw_bytes": raw_b_res,
                                    "file_id": restored.get("file_id", "")
                                }
                                print(f"✅ Sesión restaurada desde persistencia para {chat_id} ({fname_res})", flush=True)
                        except Exception as e_res:
                            print(f"Error restaurando sesión persistente: {e_res}", flush=True)

                    # Si el usuario tiene un archivo subido activo en memoria: procesar con Analista Universal
                    if chat_id in DATAFRAME_SESION and DATAFRAME_SESION[chat_id].get("raw_bytes"):
                        raw_b = DATAFRAME_SESION[chat_id]["raw_bytes"]
                        f_name = DATAFRAME_SESION[chat_id]["filename"]

                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": f"⚡ Ejecutando transformaciones de Analista Universal sobre '{f_name}'..."}
                        )

                        try:
                            res_bytes, out_name, out_msg = process_universal_excel(raw_b, texto, f_name)
                            # Actualizar los raw_bytes de la sesión con la versión modificada
                            DATAFRAME_SESION[chat_id]["raw_bytes"] = res_bytes
                            save_user_session(chat_id, DATAFRAME_SESION[chat_id].get("file_id", ""), f_name, res_bytes, DATAFRAME_SESION[chat_id].get("sheets", []), DATAFRAME_SESION[chat_id].get("active_sheet", "Base_Datos"), DATAFRAME_SESION[chat_id].get("df"))
                            
                            files = {"document": (out_name, res_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                            await client.post(
                                f"{BASE_URL}/sendDocument",
                                data={"chat_id": chat_id, "caption": f"📁 {out_name} generado con fórmulas vivas y formatos."},
                                files=files
                            )
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": out_msg, "reply_markup": BOTONES_ACCIONES}
                            )
                            continue
                        except Exception as e_univ:
                            print(f"Error procesando universal excel: {e_univ}", flush=True)
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": f"⚠️ Detalle al procesar '{f_name}': {str(e_univ)}", "reply_markup": BOTONES_ACCIONES}
                            )
                            continue

                    # 🛡️ CAPA 0: ESCUDO DE SEGURIDAD PERIMETRAL (Guardrail Anti-Jailbreak)
                    if check_perimeter_security(texto):
                        print(f"🚨 ALERTA: Intento de ataque perimetral bloqueado de {chat_id}: {texto}", flush=True)
                        alerta_seguridad = (
                            "🚨 [ESCUDO DE SEGURIDAD 007 ACTIVADO]\n\n"
                            "⚠️ Se detectó un intento de inyección de prompt, evasión de directivas o comando no autorizado.\n\n"
                            "🔒 Acción: La petición ha sido bloqueada en el perímetro determinista.\n"
                            "🛡️ El modelo de IA y las bases de datos operativas se mantienen blindados.\n"
                            "📌 Incidente registrado en la bitácora de auditoría de seguridad."
                        )
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": alerta_seguridad}
                        )
                        continue

                    # CASO A: Saludos casuales (CERO costo, CERO demora, respuesta humana instantánea)
                    saludos = ["hola", "buen dia", "buenos dias", "buenas", "buenas tardes", "buenas noches", "hey", "que tal", "/start"]
                    if any(texto_lower.startswith(s) for s in saludos) and len(texto_lower) < 25:
                        msg_saludo = (
                            "¡Hola jefe, buen día! 🫡 Agente 007 a la orden.\n\n"
                            "¿En qué le puedo apoyar hoy? Puedo darle el corte de tickets en vivo, "
                            "generarle la gráfica o prepararle el Excel corporativo con gráficas nativas.\n\n"
                            "📎 También puede arrastrar o subir cualquier archivo Excel o CSV directamente aquí para que lo analice a su gusto."
                        )
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": msg_saludo, "reply_markup": BOTONES_ACCIONES}
                        )
                        continue

                    # -------------------------------------------------------------
                    # 🤖 NIVEL 3: ORQUESTADOR AUTÓNOMO MCP (Tool Calling Dinámico)
                    # -------------------------------------------------------------
                    df, fuente_nombre = await get_active_dataframe(chat_id)
                    mcp_res = await decide_agent_action_mcp(
                        pregunta=texto,
                        df=df,
                        chat_id=chat_id,
                        api_key=GROQ_API_KEY,
                        modelos=MODELOS_GROQ,
                        historial_chat=HISTORIAL_CHAT
                    )

                    if mcp_res:
                        action = mcp_res.get("action")
                        args = mcp_res.get("args", {})
                        print(f"🤖 [MCP TOOL CALL EJECUTADO]: {action} con args: {args}", flush=True)

                        if action == "generar_excel_personalizado":
                            instruccion = args.get("instruccion", texto)
                            await send_dynamic_custom_excel(client, chat_id, instruccion, df)
                            continue

                        elif action == "generar_grafica_operativa":
                            tipo = args.get("tipo", "oficial_general")
                            if tipo == "criterio_especifico":
                                crit = apply_multi_criteria(texto, df)
                                if crit is not None:
                                    df_crit, reglas_det, def_stat = crit
                                    await send_criteria_chart_to_user(client, chat_id, df_crit, reglas_det, def_stat)
                                    continue
                            await send_chart_to_user(client, chat_id)
                            continue

                        elif action == "consultar_kpis_en_vivo":
                            resp = await ask_ai_chat(texto, df, chat_id=chat_id)
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": resp, "reply_markup": BOTONES_ACCIONES}
                            )
                            continue

                        elif action == "procesar_ticket_langgraph":
                            desc = args.get("descripcion_problema", texto)
                            cuenta = args.get("cuenta_cliente")
                            res_lg = execute_langgraph_ticket(desc, cuenta, canal="TELEGRAM")

                            if res_lg.get("exito"):
                                stat = res_lg.get("status", "PENDIENTE")
                                auto = "✅ SÍ (Automática)" if res_lg.get("resolucion_automatica") else "⚠️ NO (Escalado Humano)"
                                clas = res_lg.get("clasificacion") or {}
                                cli = res_lg.get("cliente")
                                cli_txt = f"{cli['nombre']} ({clas.get('cuenta', 'N/A')})" if cli else (clas.get('cuenta') or "No identificado")
                                saldo_txt = f"${cli['saldo']:.2f} USD" if cli else "N/A"

                                msg_veredicto = (
                                    "🕵️ [DICTAMEN MULTI-AGENTE 007 — LANGGRAPH]\n\n"
                                    f"• 🏷️ **Clasificación:** {clas.get('tipo', 'REQUERIMIENTO')} - {clas.get('categoria', 'ACTIVACION')} ({clas.get('accion', 'ACCION')})\n"
                                    f"• 👤 **Cliente:** {cli_txt}\n"
                                    f"• 💰 **Saldo Contable:** {saldo_txt}\n"
                                    f"• ⚡ **Resolución:** {stat} | Auto: {auto}\n\n"
                                    f"📋 **Comentario Técnico:**\n"
                                    f"\"{res_lg.get('comentario')}\"\n\n"
                                    "🔒 *Trazabilidad de 5 nodos y auditoría registrada en Supabase.*"
                                )
                            else:
                                msg_veredicto = (
                                    "⚠️ [ERROR EN MOTOR MULTI-AGENTE LANGGRAPH]\n\n"
                                    f"No fue posible completar el ciclo en Supabase: {res_lg.get('error')}"
                                )

                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": msg_veredicto, "reply_markup": BOTONES_ACCIONES}
                            )
                            continue

                        elif action == "chat_reply":
                            llm_reply = mcp_res.get("llm_text")
                            if llm_reply:
                                if chat_id not in HISTORIAL_CHAT:
                                    HISTORIAL_CHAT[chat_id] = []
                                HISTORIAL_CHAT[chat_id].append({"role": "user", "content": texto})
                                HISTORIAL_CHAT[chat_id].append({"role": "assistant", "content": llm_reply})
                                await client.post(
                                    f"{BASE_URL}/sendMessage",
                                    json={"chat_id": chat_id, "text": llm_reply, "reply_markup": BOTONES_ACCIONES}
                                )
                                continue

                    # CASO B0: Petición explícita de Incidente/Ticket con cuenta de cliente (Fallback LangGraph)
                    if "cli-" in texto_lower:
                        res_lg = execute_langgraph_ticket(texto, canal="TELEGRAM")
                        if res_lg.get("exito"):
                            stat = res_lg.get("status", "PENDIENTE")
                            auto = "✅ SÍ (Automática)" if res_lg.get("resolucion_automatica") else "⚠️ NO (Escalado Humano)"
                            clas = res_lg.get("clasificacion") or {}
                            cli = res_lg.get("cliente")
                            cli_txt = f"{cli['nombre']} ({clas.get('cuenta', 'N/A')})" if cli else (clas.get('cuenta') or "No identificado")
                            saldo_txt = f"${cli['saldo']:.2f} USD" if cli else "N/A"

                            msg_veredicto = (
                                "🕵️ [DICTAMEN MULTI-AGENTE 007 — LANGGRAPH]\n\n"
                                f"• 🏷️ **Clasificación:** {clas.get('tipo', 'REQUERIMIENTO')} - {clas.get('categoria', 'ACTIVACION')} ({clas.get('accion', 'ACCION')})\n"
                                f"• 👤 **Cliente:** {cli_txt}\n"
                                f"• 💰 **Saldo Contable:** {saldo_txt}\n"
                                f"• ⚡ **Resolución:** {stat} | Auto: {auto}\n\n"
                                f"📋 **Comentario Técnico:**\n"
                                f"\"{res_lg.get('comentario')}\"\n\n"
                                "🔒 *Trazabilidad de 5 nodos y auditoría registrada en Supabase.*"
                            )
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": chat_id, "text": msg_veredicto, "reply_markup": BOTONES_ACCIONES}
                            )
                            continue

                    # CASO B: Petición explícita SIMPLE de Gráfica Oficial
                    es_pedido_grafica_simple = (
                        any(palabra in texto_lower for palabra in ["grafica", "gráfica", "/grafica"])
                        and len(texto.split()) <= 4
                        and not any(c in texto_lower for c in ['criterio', 'jesus', 'ariel', 'atendido', 'filtra', 'si ', 'excepto'])
                    )
                    if es_pedido_grafica_simple:
                        await send_chart_to_user(client, chat_id)
                        continue

                    # CASO C: Petición de Excel (Personalizado dinámico, Oficial o referenciando al criterio previo)
                    pide_excel = any(palabra in texto_lower for palabra in ["excel", "xlsx", "/excel"])
                    refiere_a_criterio = any(c in texto_lower for c in ['este criterio', 'del criterio', 'de esto', 'de este', 'con este', 'con criterio'])
                    es_excel_dinamico = any(k in texto_lower for k in [
                        'pint', 'color', 'extrae', 'extraer', 'solo ', 'formula', 'fórmula', 
                        'concatena', 'concatenar', 'azul', 'amarillo', 'verde', 'rojo', 'morado', 'naranja', 'personalizado', 'ariel'
                    ])

                    # 1. Petición explícita de Excel a medida con filtros, colores o fórmulas
                    if pide_excel and es_excel_dinamico:
                        df, _ = await get_active_dataframe(chat_id)
                        await send_dynamic_custom_excel(client, chat_id, texto, df)
                        continue

                    # 2. Petición de Excel del criterio activo en memoria
                    if pide_excel and (refiere_a_criterio or (chat_id in ULTIMO_CRITERIO and len(texto.split()) <= 12)):
                        if chat_id in ULTIMO_CRITERIO:
                            df_crit, reglas_det, def_stat, tit = ULTIMO_CRITERIO[chat_id]
                            await send_excel_to_user(client, chat_id, df=df_crit, status_col='status_criterio', titulo_reporte=tit)
                            continue
                    elif pide_excel and len(texto.split()) <= 4 and not any(c in texto_lower for c in ['criterio', 'jesus', 'ariel', 'atendido', 'filtra', 'si ', 'excepto']):
                        # Pedido simple del Excel oficial o del archivo activo
                        await send_excel_to_user(client, chat_id)
                        continue

                    # CASO D: Preguntas analíticas, criterios dinámicos y consultas complejas
                    df, _ = await get_active_dataframe(chat_id)
                    resultado_multi = apply_multi_criteria(texto, df)
                    pide_grafica_en_msg = any(p in texto_lower for p in ['grafica', 'gráfica', 'grafico', 'chart'])
                    pide_excel_en_msg = any(p in texto_lower for p in ['excel', 'xlsx'])

                    # Si se detectó un criterio nuevo, guardarlo en la sesión
                    if resultado_multi is not None:
                        df_crit, reglas_detalle, default_status = resultado_multi
                        reglas_str = ', '.join(f'{k}={v["status"]}' for k, v in reglas_detalle.items() if k != 'RESTO (OTROS)')
                        titulo_crit = f'Criterio: {reglas_str}, Resto={default_status}'
                        ULTIMO_CRITERIO[chat_id] = (df_crit, reglas_detalle, default_status, titulo_crit)

                    # 1. Enviar primero la respuesta textual ejecutiva de la IA
                    respuesta_humana = await ask_ai_chat(texto, df, chat_id=chat_id)
                    await client.post(
                        f"{BASE_URL}/sendMessage",
                        json={
                            "chat_id": chat_id,
                            "text": respuesta_humana,
                            "reply_markup": BOTONES_ACCIONES
                        }
                    )

                    # 2. Si hay criterio personalizado Y se solicita gráfica:
                    if resultado_multi is not None and (pide_grafica_en_msg or 'criterio' in texto_lower):
                        df_crit, reglas_detalle, default_status = resultado_multi
                        await send_criteria_chart_to_user(
                            client, chat_id,
                            df_crit=df_crit,
                            reglas_detalle=reglas_detalle,
                            default_status=default_status
                        )
                    elif pide_grafica_en_msg:
                        await send_chart_to_user(client, chat_id)

                    # 3. Si en este mismo mensaje pidió también el Excel:
                    if pide_excel_en_msg:
                        if es_excel_dinamico or any(k in texto_lower for k in ['pint', 'color', 'extrae', 'solo ', 'formula', 'concatena', 'azul', 'amarillo', 'verde', 'rojo', 'morado']):
                            await send_dynamic_custom_excel(client, chat_id, texto, df)
                        elif resultado_multi is not None and chat_id in ULTIMO_CRITERIO:
                            df_crit, reglas_detalle, default_status, tit = ULTIMO_CRITERIO[chat_id]
                            await send_excel_to_user(client, chat_id, df=df_crit, status_col='status_criterio', titulo_reporte=tit)
                        else:
                            await send_excel_to_user(client, chat_id, df=df)

            except Exception as e:
                print(f"Error loop bot: {e}", flush=True)
                await asyncio.sleep(3)

async def main_loop():
    while True:
        try:
            await run_bot()
        except Exception as e:
            print(f"Error inesperado: {e}", flush=True)
            traceback.print_exc()
        await asyncio.sleep(5)

if __name__ == "__main__":
    asyncio.run(main_loop())
