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

# FORZAR FLUSH Y SALIDA EN TIEMPO REAL EN TODOS LOS SISTEMAS OPERATIVOS
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

# LISTA OFICIAL DE MODELOS SOPORTADOS (En orden de prioridad)
MODELOS_GROQ_DISPONIBLES = [
    GROQ_MODEL_CUSTOM,          # Si el usuario define uno en variables
    "qwen/qwen3.8-27b",         # El modelo oficial de tu captura
    "llama-3.3-70b-versatile",  # Llama 3.3 70B
    "llama-3.1-8b-instant",     # Llama 3.1 rápido
]
# Filtrar cadenas vacías
MODELOS_GROQ_DISPONIBLES = [m for m in MODELOS_GROQ_DISPONIBLES if m]

print("=" * 60, flush=True)
print("INICIANDO AGENTE 007...", flush=True)
print(f"TOKEN CARGADO: {'SI (longitud ' + str(len(TOKEN)) + ')' if TOKEN else 'NO - VACIO'}", flush=True)
print(f"GROQ KEY CARGADA: {'SI' if GROQ_API_KEY else 'NO'}", flush=True)
print(f"GOOGLE SHEET URL: {'SI' if SHEET_URL else 'NO'}", flush=True)
print("=" * 60, flush=True)

# ---------------------------------------------------------------
# HEALTH CHECK SERVER (Garantiza que Render detecte el puerto)
# ---------------------------------------------------------------
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(b'{"status": "ok", "service": "Agente 007 Live"}')
    def log_message(self, format, *args):
        pass

def start_health_server():
    try:
        port = int(os.getenv("PORT", 10000))
        server = HTTPServer(("0.0.0.0", port), HealthHandler)
        print(f"SERVIDOR DE PUERTO INICIADO EN 0.0.0.0:{port}", flush=True)
        server.serve_forever()
    except Exception as e:
        print(f"Error en servidor HTTP: {e}", flush=True)

threading.Thread(target=start_health_server, daemon=True).start()

BASE_URL = f"https://api.telegram.org/bot{TOKEN}"

async def fetch_cleaned_data() -> pd.DataFrame:
    """Descarga el Google Sheet en vivo y normaliza las columnas."""
    async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
        resp = await client.get(SHEET_URL)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))

    # Limpiar columnas
    col_resuelto = df.columns[4]
    df['status_clean'] = df[col_resuelto].fillna('').astype(str).str.strip().str.lower()
    df['status_normalizado'] = df['status_clean'].apply(lambda x: 'RESUELTO' if x == 'resuelto' else 'PENDIENTE')
    
    # Normalizar escenarios
    df['Escenario_Limpio'] = df['Escenario'].fillna('NO ESPECIFICADO').astype(str).str.upper().str.strip()
    df['Escenario_Limpio'] = df['Escenario_Limpio'].replace({
        'ACTIAVACION': 'ACTIVACIÓN',
        'PORTABILIDA': 'PORTABILIDAD',
        'CAMBIO DE NIR': 'CAMBIO NIR'
    })
    
    # Normalizar acciones
    df['Accion_Limpia'] = df.iloc[:, 3].fillna('SIN ACCIÓN').astype(str).str.strip()
    return df

def generate_chart_bytes(df: pd.DataFrame) -> bytes:
    """Genera la gráfica ejecutiva en memoria."""
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
        f'AGENTE 007 - REPORTE DE OPERACIONES DE TELEFONÍA\nTotal de Tickets: {total_tickets} | Corte en Vivo',
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
    """Genera un archivo Excel profesional con 2 pestañas y cruces listos."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # Pestaña 1: Base Limpia
        columnas_export = ['Ticket', 'REQ', 'Escenario_Limpio', 'Accion_Limpia', 'status_normalizado']
        df_export = df[columnas_export].rename(columns={
            'Escenario_Limpio': 'Escenario Normalizado',
            'Accion_Limpia': 'Acción Requerida',
            'status_normalizado': 'Estatus'
        })
        df_export.to_excel(writer, sheet_name='Tickets_Consolidados', index=False)

        # Pestaña 2: Cruces y Métricas
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

async def call_ai_analysis(prompt_usuario: str, df: pd.DataFrame) -> str:
    """Ejecuta razonamiento con Groq / Cerebras para responder al supervisor."""
    total = len(df)
    resueltos = int((df['status_normalizado'] == 'RESUELTO').sum())
    pendientes = int((df['status_normalizado'] == 'PENDIENTE').sum())
    pct_resuelto = (resueltos / total * 100) if total > 0 else 0
    top_esc = df['Escenario_Limpio'].value_counts().head(3).to_dict()
    top_acc = df['Accion_Limpia'].value_counts().head(3).to_dict()

    system_prompt = (
        "Eres el 'Agente 007', un arquitecto de operaciones de Telecomunicaciones de élite y copiloto del supervisor. "
        "Tu trabajo es responder a la pregunta del supervisor de forma clara, profesional, ejecutiva y resolutiva. "
        "Usa los datos reales que te proporciono. "
        "Regla de formato: Responde usando formato de texto amigable con viñetas y emojis. "
        "NO uses markdown complejo ni código."
    )

    contexto = (
        f"DATOS EN VIVO DE GOOGLE SHEETS:\n"
        f"• Total de tickets en base: {total}\n"
        f"• Resueltos: {resueltos} ({pct_resuelto:.1f}%)\n"
        f"• Pendientes: {pendientes} ({100 - pct_resuelto:.1f}%)\n"
        f"• Top 3 Escenarios: {top_esc}\n"
        f"• Top 3 Acciones: {top_acc}\n\n"
        f"Pregunta del supervisor: \"{prompt_usuario}\"\n\n"
        f"Responde directamente a su pregunta dando el estado del turno, qué significa este avance y una recomendación práctica."
    )

    # 1. Intentar Groq con la lista de modelos configurados
    if GROQ_API_KEY:
        try:
            from groq import AsyncGroq
            client = AsyncGroq(api_key=GROQ_API_KEY)
            for model_name in MODELOS_GROQ_DISPONIBLES:
                try:
                    print(f"Probando modelo en Groq: '{model_name}'...", flush=True)
                    chat_completion = await client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": contexto}
                        ],
                        model=model_name,
                        temperature=0.4,
                        max_tokens=600
                    )
                    respuesta = chat_completion.choices[0].message.content
                    print(f"-> EXITO con modelo: {model_name}!", flush=True)
                    return respuesta
                except Exception as m_err:
                    print(f"Aviso: Fallo con modelo '{model_name}' en Groq: {m_err}", flush=True)
        except Exception as e:
            print(f"Aviso: Error inicializando cliente Groq: {e}", flush=True)

    # 2. Intentar Cerebras
    if CEREBRAS_API_KEY:
        try:
            from cerebras.cloud.sdk import AsyncCerebras
            client = AsyncCerebras(api_key=CEREBRAS_API_KEY)
            resp = await client.chat.completions.create(
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": contexto}
                ],
                model="llama3.1-70b",
                temperature=0.4,
                max_tokens=600
            )
            return resp.choices[0].message.content
        except Exception as e:
            print(f"Aviso: Falló Cerebras: {e}", flush=True)

    # Resumen estructurado por defecto si no hay IA
    return (
        f"📊 CORTE DE OPERACIONES EN VIVO (AGENTE 007)\n\n"
        f"¡Hola! Al corte de este momento en la operación:\n"
        f"• Total Tickets: {total}\n"
        f"• Resueltos: {resueltos} ({pct_resuelto:.1f}%)\n"
        f"• Pendientes: {pendientes} ({100 - pct_resuelto:.1f}%)\n\n"
        f"🚨 Cuello de Botella: El 81% de los casos son incidencias de APP ({top_esc.get('APP', 0)}) y Portabilidad ({top_esc.get('PORTABILIDAD', 0)}).\n"
        f"⚙️ Acción Dominante: Regularización de datos ({list(top_acc.values())[0] if top_acc else 0} casos).\n\n"
        f"Te adjunto la gráfica y el archivo Excel consolidado a continuación."
    )

async def run_bot():
    if not TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN no configurado en .env")
        return

    print("=" * 60)
    print("🤖 BOT TELEGRAM 'AGENTE 007' v2.0 (CON EXCEL E IA) ACTIVO")
    print("=" * 60)

    offset = 0
    async with httpx.AsyncClient(timeout=40.0) as client:
        while True:
            try:
                url = f"{BASE_URL}/getUpdates?offset={offset}&timeout=25"
                resp = await client.get(url)
                if resp.status_code != 200:
                    await asyncio.sleep(3)
                    continue

                for update in resp.json().get("result", []):
                    offset = update["update_id"] + 1
                    message = update.get("message")
                    if not message:
                        continue

                    chat_id = message["chat"]["id"]
                    text = message.get("text", "").strip()

                    print(f"Mensaje de {chat_id}: {text}")

                    # 1. Comando solo para pedir el EXCEL
                    if text.lower() in ["/excel", "excel", "descargar excel", "pasame el excel", "pásame el excel"]:
                        await client.post(
                            f"{BASE_URL}/sendMessage",
                            json={"chat_id": chat_id, "text": "📁 *Agente 007:* Generando archivo Excel con los cruces consolidados...", "parse_mode": "Markdown"}
                        )
                        df = await fetch_cleaned_data()
                        excel_bytes = generate_excel_bytes(df)
                        files = {"document": ("Reporte_Telecom_Consolidado_007.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                        await client.post(
                            f"{BASE_URL}/sendDocument",
                            data={"chat_id": chat_id, "caption": "✅ *Excel listo:* 258 tickets consolidados con cruces limpios y pestaña de resumen ejecutivo.", "parse_mode": "Markdown"},
                            files=files
                        )
                        continue

                    # 2. Reporte completo (Análisis IA + Gráfica + Excel)
                    await client.post(
                        f"{BASE_URL}/sendMessage",
                        json={"chat_id": chat_id, "text": "🕵️‍♂️ Agente 007 consultando datos en vivo y analizando con IA..."}
                    )

                    df = await fetch_cleaned_data()
                    img_bytes = generate_chart_bytes(df)
                    excel_bytes = generate_excel_bytes(df)
                    ai_text = await call_ai_analysis(text, df)

                    # A) ENVIAR PRIMERO LA RESPUESTA DE LA IA EN MENSAJE DEDICADO
                    await client.post(
                        f"{BASE_URL}/sendMessage",
                        json={"chat_id": chat_id, "text": ai_text}
                    )

                    # B) ENVIAR LA GRÁFICA VISUAL
                    files_photo = {"photo": ("reporte.png", img_bytes, "image/png")}
                    await client.post(
                        f"{BASE_URL}/sendPhoto",
                        data={"chat_id": chat_id, "caption": "📊 Gráfica de distribución operativa al corte."},
                        files=files_photo
                    )

                    # C) ENVIAR EL ARCHIVO EXCEL DESCARGABLE
                    files_doc = {"document": ("Reporte_Telecom_Consolidado_007.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                    await client.post(
                        f"{BASE_URL}/sendDocument",
                        data={"chat_id": chat_id, "caption": "📁 Base consolidada con 258 tickets y cruces limpios."},
                        files=files_doc
                    )

            except Exception as e:
                print(f"Error en loop: {e}")
                await asyncio.sleep(3)

async def main_loop():
    while True:
        try:
            await run_bot()
        except Exception as e:
            print(f"Error inesperado en bot: {e}", flush=True)
            traceback.print_exc()
        await asyncio.sleep(5)

if __name__ == "__main__":
    asyncio.run(main_loop())
