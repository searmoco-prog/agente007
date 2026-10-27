import os
import sys
import io
import asyncio
import traceback
import httpx
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from dotenv import load_dotenv

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
        self.wfile.write(b'{"status": "ok", "service": "Agente 007 v3.0"}')
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
    """Genera la imagen PNG de la gráfica."""
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
    """Genera el archivo Excel profesional."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        col_export = ['Ticket', 'REQ', 'Escenario_Limpio', 'Accion_Limpia', 'status_normalizado']
        df_export = df[col_export].rename(columns={
            'Escenario_Limpio': 'Escenario Normalizado',
            'Accion_Limpia': 'Acción Requerida',
            'status_normalizado': 'Estatus'
        })
        df_export.to_excel(writer, sheet_name='Tickets_Consolidados', index=False)

        resumen_data = {
            'Métrica': ['Total Tickets', 'Resueltos', 'Pendientes', '% Avance'],
            'Valor': [
                len(df),
                int((df['status_normalizado'] == 'RESUELTO').sum()),
                int((df['status_normalizado'] == 'PENDIENTE').sum()),
                f"{(df['status_normalizado'] == 'RESUELTO').sum() / len(df) * 100:.1f}%"
            ]
        }
        pd.DataFrame(resumen_data).to_excel(writer, sheet_name='Resumen_Operativo', index=False)

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
        "Eres el 'Agente 007', el asistente y copiloto inteligente de operaciones de telecomunicaciones. "
        "Habla de forma humana, profesional, amigable y respetuosa con tu jefe ('Hola jefe', 'Buen día'). "
        "Sé breve y conciso: máximo 2 o 3 párrafos cortos con viñetas. "
        "Nunca uses markdown complejo ni código. Responde con calidez y precisión."
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
# BOT LOOP CON ENRUTADOR DE INTENCIONES Y BOTONES
# ---------------------------------------------------------------
BOTONES_ACCIONES = {
    "inline_keyboard": [
        [
            {"text": "📊 Ver Gráfica", "callback_data": "btn_chart"},
            {"text": "📁 Descargar Excel", "callback_data": "btn_excel"}
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
        json={"chat_id": chat_id, "text": "📊 Generando gráfica actualizada..."}
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
    """Envía únicamente el archivo Excel."""
    await client.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": "📁 Compilando archivo Excel consolidado..."}
    )
    df = await fetch_cleaned_data()
    excel_bytes = generate_excel_bytes(df)
    files = {"document": ("Reporte_Telecom_Consolidado_007.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    await client.post(
        f"{BASE_URL}/sendDocument",
        data={"chat_id": chat_id, "caption": "📁 Base consolidada con los 258 tickets normalizados y resumen ejecutivo."},
        files=files
    )

async def run_bot():
    if not TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN no configurado", flush=True)
        return

    print("=" * 60, flush=True)
    print("🤖 AGENTE 007 v3.0 (MODO CONVERSACIONAL HUMANO) ACTIVO", flush=True)
    print("=" * 60, flush=True)

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
                        
                        # Confirmar recepción del botón a Telegram
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

                    # CASO A: Saludos casuales (CERO costo, CERO demora, respuesta humana instantánea)
                    saludos = ["hola", "buen dia", "buenos dias", "buenas", "buenas tardes", "buenas noches", "hey", "que tal", "/start"]
                    if any(texto_lower.startswith(s) for s in saludos) and len(texto_lower) < 25:
                        msg_saludo = (
                            "¡Hola jefe, buen día! 🫡 Agente 007 a la orden.\n\n"
                            "¿En qué le puedo apoyar hoy? Puedo darle el corte de tickets en vivo, "
                            "generarle la gráfica o prepararle el Excel consolidado."
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
                    # Responde con texto amigable y ofrece los botones sin spamear archivos pesados
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
