from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from app.schemas.ticket import TicketIngress
from app.graph.workflow import agente007_app

app = FastAPI(
    title="Agente 007 - Autonomous Trouble-Ticket & Cross-Resolution Engine",
    description="API empresarial de resolución y cruce automático de tickets con LangGraph y Supabase",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def health_check():
    return {
        "status": "online",
        "system": "Agente 007 Core API",
        "version": "1.0.0"
    }

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/v1/tickets/process", status_code=status.HTTP_200_OK)
async def process_ticket(ticket_input: TicketIngress):
    """Recibe un ticket y ejecuta el flujo completo de Agente 007."""
    try:
        initial_state = {
            "ticket_id": None,
            "descripcion_original": ticket_input.descripcion_original,
            "cuenta_cliente_input": ticket_input.cuenta_cliente,
            "canal_origen": ticket_input.canal_origen,
            "es_seguro": True,
            "motivo_bloqueo_seguridad": None,
            "clasificacion": None,
            "datos_cliente": None,
            "catalogo_regla": None,
            "resultado_resolucion": None,
            "trazas_auditoria": []
        }
        final_state = agente007_app.invoke(initial_state)
        return {
            "success": True,
            "ticket_id": final_state.get("ticket_id"),
            "es_seguro": final_state.get("es_seguro"),
            "clasificacion": final_state.get("clasificacion"),
            "datos_cliente": final_state.get("datos_cliente"),
            "resolucion": final_state.get("resultado_resolucion")
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error en la ejecución del Agente 007: {str(e)}"
        )
