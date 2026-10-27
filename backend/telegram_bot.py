import os
import sys
import io
import re
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
    "qwen/qwen3.8-27b",
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant"
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
    """Genera la imagen PNG de la gráfica para visualización rápida."""
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
        f'AGENTE 007 - REPORTE DE OPERACIONES DE TELEFONÍA\nTotal: {total_tickets} Tickets | Corte en Vivo',
        fontsize=15, fontweight='bold', color='#38bdf8', y=0.96
    )

    # Donut Chart
    ax1 = fig.add_subplot(1, 3, 1, facecolor='#0f172a')
    ax1.pie(
        [resueltos, pendientes], 
        labels=[f'Resueltos\n{resueltos} ({pct_resuelto:.1f}%)', f'Pendientes\n{pendientes} ({pct_pendiente:.1f}%)'], 
        colors=['#10b981', '#f43f5e'], 
        startangle=140,
        textprops={'fontsize': 10, 'color': '#f8fafc', 'fontweight': 'bold'},
        wedgeprops={'width': 0.45, 'edgecolor': '#0f172a', 'linewidth': 2}
    )
    ax1.set_title('Estatus Global', fontsize=12, color='#e2e8f0', pad=12, fontweight='bold')

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

def generate_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Genera un libro Excel corporativo de nivel financiero con:
    - Gráficas nativas interactivas (PieChart y BarChart incrustados).
    - Fórmulas dinámicas vivas (=COUNTIF, =COUNTA, porcentajes).
    - Formato condicional automático (Verde suave para RESUELTO, Coral suave para PENDIENTE).
    - Cruces y agregaciones consolidadas.
    """
    wb = openpyxl.Workbook()
    
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

    headers = ["Ticket", "REQ", "Escenario Normalizado", "Acción Requerida", "Estatus Operativo"]
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
        esc_val = str(row.Escenario_Limpio)
        acc_val = str(row.Accion_Limpia)
        stat_val = str(row.status_normalizado)

        ws_data.append([ticket_val, req_val, esc_val, acc_val, stat_val])

        ws_data.cell(row=row_idx, column=1).alignment = align_center
        ws_data.cell(row=row_idx, column=2).alignment = align_center
        ws_data.cell(row=row_idx, column=3).alignment = align_left
        ws_data.cell(row=row_idx, column=4).alignment = align_left
        
        stat_cell = ws_data.cell(row=row_idx, column=5)
        stat_cell.alignment = align_center
        stat_cell.font = font_bold

        for c in range(1, 6):
            ws_data.cell(row=row_idx, column=c).border = thin_border

    # Formato Condicional en Estatus (Columna E)
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    green_font = Font(color="375623", bold=True)
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    red_font = Font(color="C00000", bold=True)

    rule_resuelto = CellIsRule(operator='equal', formula=['"RESUELTO"'], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_pendiente = CellIsRule(operator='equal', formula=['"PENDIENTE"'], stopIfTrue=True, fill=red_fill, font=red_font)

    status_range = f"E2:E{total_rows + 1}"
    ws_data.conditional_formatting.add(status_range, rule_resuelto)
    ws_data.conditional_formatting.add(status_range, rule_pendiente)

    # Auto-ajustar anchos
    for col in ws_data.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Activar filtro automático
    ws_data.auto_filter.ref = f"A1:E{total_rows + 1}"

    # -------------------------------------------------------------
    # HOJA 1: Dashboard_Ejecutivo (Primera hoja visible)
    # -------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "Dashboard_Ejecutivo"
    ws_dash.views.sheetView[0].showGridLines = True

    # Banner Corporativo
    ws_dash.merge_cells("A1:G2")
    banner = ws_dash["A1"]
    banner.value = "  AGENTE 007 — DASHBOARD DE OPERACIONES TELECOM (NOC L1/L2)"
    banner.font = font_titulo
    banner.fill = fill_navy
    banner.alignment = Alignment(horizontal="left", vertical="center")

    ws_dash["A3"] = "Reporte Dinámico en Tiempo Real | Corte de Incidentes y Cruce de Acciones Operativas"
    ws_dash["A3"].font = font_subtitulo

    # Tabla 1: Indicadores Clave de Estatus (Fórmulas vivas)
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

    ws_dash["A7"] = "Resueltos"
    ws_dash["B7"] = f'=COUNTIF(Base_Consolidada!E2:E{total_rows + 1}, "RESUELTO")'
    ws_dash["C7"] = "=B7/B6"

    ws_dash["A8"] = "Pendientes"
    ws_dash["B8"] = f'=COUNTIF(Base_Consolidada!E2:E{total_rows + 1}, "PENDIENTE")'
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
# IA CONVERSACIONAL (Groq / Cerebras)
# ---------------------------------------------------------------
async def ask_ai_chat(pregunta: str, df: pd.DataFrame) -> str:
    """Genera una respuesta humana, concisa y ejecutiva sin textos kilométricos."""
    total = len(df)
    resueltos = int((df['status_normalizado'] == 'RESUELTO').sum())
    pendientes = int((df['status_normalizado'] == 'PENDIENTE').sum())
    pct = (resueltos / total * 100) if total > 0 else 0
    top_esc = df['Escenario_Limpio'].value_counts().head(3).to_dict()

    system_prompt = (
        "Eres el 'Agente 007', el copiloto inteligente de operaciones de telecomunicaciones. "
        "Habla de forma humana, profesional, amigable y respetuosa con tu jefe ('Hola jefe', 'Buen día'). "
        "Sé breve y conciso: máximo 2 o 3 párrafos cortos con viñetas. "
        "Nunca uses markdown complejo ni código. Responde con calidez y precisión técnica."
    )

    contexto = (
        f"ESTADO DE TICKETS EN GOOGLE SHEETS HOY:\n"
        f"• Total: {total} tickets\n"
        f"• Resueltos: {resueltos} ({pct:.1f}%)\n"
        f"• Pendientes: {pendientes} ({100 - pct:.1f}%)\n"
        f"• Principales escenarios: {top_esc}\n\n"
        f"Mensaje de tu jefe: \"{pregunta}\"\n\n"
        f"Responde directamente a lo que te pregunta de forma natural y breve."
    )

    if GROQ_API_KEY:
        try:
            from groq import AsyncGroq
            client = AsyncGroq(api_key=GROQ_API_KEY)
            for m in MODELOS_GROQ:
                try:
                    res = await client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": contexto}
                        ],
                        model=m,
                        temperature=0.4,
                        max_tokens=350
                    )
                    return res.choices[0].message.content
                except Exception:
                    pass
        except Exception as e:
            print(f"Error Groq: {e}", flush=True)

    # Fallback si no hay IA
    return (
        f"¡Hola jefe! Al corte de este momento llevamos {total} tickets en la base:\n\n"
        f"• ✅ {resueltos} resueltos ({pct:.1f}%)\n"
        f"• ⏳ {pendientes} pendientes\n\n"
        f"La gran mayoría están concentrados en App y Portabilidad. ¿Gusta que le envíe la gráfica o el archivo Excel?"
    )

# ---------------------------------------------------------------
# BOTONES INTERACTIVOS
# ---------------------------------------------------------------
BOTONES_ACCIONES = {
    "inline_keyboard": [
        [
            {"text": "📊 Ver Gráfica", "callback_data": "btn_chart"},
            {"text": "📁 Descargar Excel Avanzado", "callback_data": "btn_excel"}
        ],
        [
            {"text": "🔄 Actualizar Resumen", "callback_data": "btn_summary"}
        ]
    ]
}

async def send_chart_to_user(client: httpx.AsyncClient, chat_id: int):
    """Envía únicamente la gráfica."""
    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "📊 Generando gráfica actualizada de operaciones..."}
    )
    df = await fetch_cleaned_data()
    img_bytes = generate_chart_bytes(df)
    files = {"photo": ("grafica_007.png", img_bytes, "image/png")}
    await client.post(
        f"{BASE_URL}/sendPhoto",
        data={"chat_id": chat_id, "caption": "📊 Gráfica oficial al corte de operaciones de hoy."},
        files=files
    )

async def send_excel_to_user(client: httpx.AsyncClient, chat_id: int):
    """Envía el archivo Excel con gráficas nativas, fórmulas y formato condicional."""
    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "📁 Compilando libro Excel corporativo con gráficas nativas y fórmulas dinámicas..."}
    )
    df = await fetch_cleaned_data()
    excel_bytes = generate_excel_bytes(df)
    files = {"document": ("Reporte_Telecom_Consolidado_007.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    await client.post(
        f"{BASE_URL}/sendDocument",
        data={"chat_id": chat_id, "caption": "📁 Reporte Corporativo Oficial: Dashboard con gráficas nativas de Excel, fórmulas automáticas y formato condicional."},
        files=files
    )

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

                        if cb_data == "btn_chart":
                            await send_chart_to_user(client, cb_chat_id)
                        elif cb_data == "btn_excel":
                            await send_excel_to_user(client, cb_chat_id)
                        elif cb_data == "btn_summary":
                            df = await fetch_cleaned_data()
                            resumen = await ask_ai_chat("Dame el resumen general del turno", df)
                            await client.post(
                                f"{BASE_URL}/sendMessage",
                                json={"chat_id": cb_chat_id, "text": resumen, "reply_markup": BOTONES_ACCIONES}
                            )
                        continue

                    # 2. MANEJAR MENSAJES DE TEXTO
                    message = update.get("message")
                    if not message or "text" not in message:
                        continue

                    chat_id = message["chat"]["id"]
                    texto = message["text"].strip()
                    texto_lower = texto.lower()

                    print(f"Mensaje de {chat_id}: {texto}", flush=True)

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
                            "generarle la gráfica o prepararle el Excel corporativo con gráficas nativas."
                        )
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": msg_saludo, "reply_markup": BOTONES_ACCIONES}
                        )
                        continue

                    # CASO B: Petición explícita de Gráfica
                    if any(palabra in texto_lower for palabra in ["grafica", "gráfica", "graficas", "gráficas", "/grafica"]):
                        await send_chart_to_user(client, chat_id)
                        continue

                    # CASO C: Petición explícita de Excel
                    if any(palabra in texto_lower for palabra in ["excel", "xlsx", "/excel", "archivo", "descargar excel"]):
                        await send_excel_to_user(client, chat_id)
                        continue

                    # CASO D: Preguntas sobre el trabajo / estado ("cómo vamos", "cuántos van", preguntas operativas)
                    df = await fetch_cleaned_data()
                    respuesta_humana = await ask_ai_chat(texto, df)

                    await client.post(
                        f"{BASE_URL}/sendMessage",
                        json={
                            "chat_id": chat_id,
                            "text": respuesta_humana,
                            "reply_markup": BOTONES_ACCIONES
                        }
                    )

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
