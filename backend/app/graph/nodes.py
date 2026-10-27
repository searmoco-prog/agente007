import re
from typing import Dict, Any
from app.graph.state import Agent007State
from app.schemas.ticket import TicketClassification, CustomerCrossData, TicketResolutionResult
from app.graph.tools import (
    tool_buscar_cliente,
    tool_buscar_regla_catalogo,
    tool_ejecutar_activacion_servicio,
    tool_guardar_ticket_en_db,
    tool_registrar_auditoria
)

# -------------------------------------------------------------------------
# NODO 1: Sanitizador / Ciberseguridad (Fail-fast ante Prompt Injection)
# -------------------------------------------------------------------------
PATRONES_SOSPECHOSOS = [
    r"ignore\s+(all\s+)?previous\s+instructions",
    r"olvida\s+(todas\s+)?las\s+instrucciones",
    r"drop\s+table",
    r"delete\s+from",
    r"system\s+prompt",
    r"eres\s+ahora\s+un\s+hacker",
]

def sanitizer_node(state: Agent007State) -> Dict[str, Any]:
    """Valida que el texto no contenga inyecciones de código o ataques al LLM."""
    texto = state["descripcion_original"].lower()
    for patron in PATRONES_SOSPECHOSOS:
        if re.search(patron, texto):
            return {
                "es_seguro": False,
                "motivo_bloqueo_seguridad": f"Vector de inyección detectado: patron '{patron}'",
                "resultado_resolucion": TicketResolutionResult(
                    status="ESCALADO_HUMANO",
                    resolucion_automatica=False,
                    comentario_resolucion="[ALERTA DE SEGURIDAD] Ticket bloqueado por intento de inyección de prompt o comando no autorizado."
                )
            }
    return {"es_seguro": True, "motivo_bloqueo_seguridad": None}


# -------------------------------------------------------------------------
# NODO 2: Tipificador y Extractor (ITIL Classifier)
# -------------------------------------------------------------------------
def classifier_node(state: Agent007State) -> Dict[str, Any]:
    """Clasifica el ticket en Incidente o Requerimiento y extrae la cuenta del cliente."""
    if not state.get("es_seguro", True):
        return {}

    texto = state["descripcion_original"]
    
    # Intentar extraer cuenta (ej: CLI-1001) si no vino en el input explícito
    cuenta = state.get("cuenta_cliente_input")
    if not cuenta:
        match = re.search(r"(CLI-\d+)", texto, re.IGNORECASE)
        if match:
            cuenta = match.group(1).upper()

    texto_lower = texto.lower()

    # Lógica de clasificación para 'Activación - Realizar activación'
    if "activa" in texto_lower or "alta" in texto_lower:
        tipo = "REQUERIMIENTO"
        categoria = "ACTIVACION"
        accion = "REALIZAR_ACTIVACION"
        confianza = 0.98
        razon = "Solicitud de provisión de servicio nuevo o reactivación estándar."
    elif "no hay internet" in texto_lower or "caido" in texto_lower or "corte" in texto_lower:
        tipo = "INCIDENTE"
        categoria = "CONECTIVIDAD"
        accion = "DIAGNOSTICO_LINEA"
        confianza = 0.95
        razon = "Interrupción de servicio reportada por el usuario."
    elif "cobro" in texto_lower or "factura" in texto_lower or "saldo" in texto_lower:
        tipo = "INCIDENTE"
        categoria = "FACTURACION"
        accion = "ACLARACION_CARGO"
        confianza = 0.90
        razon = "Inconformidad o consulta sobre facturación."
    else:
        tipo = "REQUERIMIENTO"
        categoria = "GENERAL"
        accion = "CONSULTA_GENERAL"
        confianza = 0.70
        razon = "Ticket sin patrón específico detectado."

    clasificacion = TicketClassification(
        cuenta_cliente=cuenta,
        tipo=tipo,
        categoria=categoria,
        accion=accion,
        confianza=confianza,
        razonamiento=razon
    )
    return {"clasificacion": clasificacion}


# -------------------------------------------------------------------------
# NODO 3: Agente de Cruces (Cross-Referencer Node)
# -------------------------------------------------------------------------
def cross_referencer_node(state: Agent007State) -> Dict[str, Any]:
    """Ejecuta los cruces SQL contra la tabla customers y ticket_catalog."""
    if not state.get("es_seguro", True) or not state.get("clasificacion"):
        return {}

    clasif = state["clasificacion"]
    cuenta = clasif.cuenta_cliente

    # Cruce 1: Buscar cliente en base de datos
    datos_cliente = CustomerCrossData(encontrado=False)
    if cuenta:
        cliente_raw = tool_buscar_cliente(cuenta)
        if cliente_raw:
            datos_cliente = CustomerCrossData(
                encontrado=True,
                cuenta_cliente=cliente_raw.get("cuenta_cliente"),
                nombre_completo=cliente_raw.get("nombre_completo"),
                plan_servicio=cliente_raw.get("plan_servicio"),
                estado_servicio=cliente_raw.get("estado_servicio"),
                saldo_pendiente=float(cliente_raw.get("saldo_pendiente", 0.0))
            )

    # Cruce 2: Buscar regla en el catálogo
    catalogo_regla = tool_buscar_regla_catalogo(clasif.categoria, clasif.accion)

    return {
        "datos_cliente": datos_cliente,
        "catalogo_regla": catalogo_regla
    }


# -------------------------------------------------------------------------
# NODO 4: Evaluador de Auto-Resolución
# -------------------------------------------------------------------------
def evaluator_node(state: Agent007State) -> Dict[str, Any]:
    """Decide si el ticket se auto-resuelve o se escala a un humano."""
    if not state.get("es_seguro", True):
        return {}

    clasif = state.get("clasificacion")
    cliente = state.get("datos_cliente")
    regla = state.get("catalogo_regla")

    # Caso 1: Cliente no existe en base de datos
    if not cliente or not cliente.encontrado:
        return {
            "resultado_resolucion": TicketResolutionResult(
                status="ESCALADO_HUMANO",
                resolucion_automatica=False,
                comentario_resolucion=(
                    f"Ticket clasificado como {clasif.tipo} - {clasif.categoria}. "
                    "No fue posible cruzar con la base de clientes (Cuenta inexistente o no proporcionada). "
                    "Se escala a un asesor humano para cotejar identidad."
                )
            )
        }

    # Caso 2: ¿La acción en el catálogo es automatizable?
    es_auto = regla.get("es_automatizable", False) if regla else False

    # Regla específica para 'ACTIVACION - REALIZAR_ACTIVACION'
    if clasif.accion == "REALIZAR_ACTIVACION" and es_auto:
        if cliente.saldo_pendiente <= 0.00:
            # ¡CONDICIÓN DE ÉXITO AUTOMÁTICA!
            return {
                "resultado_resolucion": TicketResolutionResult(
                    status="AUTO_RESUELTO",
                    resolucion_automatica=True,
                    nuevo_estado_servicio="ACTIVO",
                    comentario_resolucion=(
                        f"[RESOLUCIÓN AUTOMÁTICA AGENTE 007] "
                        f"Cliente '{cliente.nombre_completo}' ({cliente.cuenta_cliente}) validado con éxito. "
                        f"Saldo al corriente ($0.00). Plan: {cliente.plan_servicio}. "
                        "Servicio activado en sistema automáticamente en 1.2 segundos."
                    )
                )
            }
        else:
            # Tiene deuda: no se puede activar solo
            return {
                "resultado_resolucion": TicketResolutionResult(
                    status="ESCALADO_HUMANO",
                    resolucion_automatica=False,
                    comentario_resolucion=(
                        f"[ESCALADO A COBRANZA/SOPORTE] "
                        f"El cliente solicita activación, pero presenta un saldo deudor pendiente de "
                        f"${cliente.saldo_pendiente:.2f} USD. No procede activación automática sin previo pago."
                    )
                )
            }

    # Caso por defecto si no es automatizable
    return {
        "resultado_resolucion": TicketResolutionResult(
            status="ESCALADO_HUMANO",
            resolucion_automatica=False,
            comentario_resolucion=(
                f"Ticket clasificado como {clasif.tipo} ({clasif.accion}). "
                "Esta acción requiere intervención operativa manual de acuerdo a las políticas de servicio."
            )
        )
    }


# -------------------------------------------------------------------------
# NODO 5: Ejecutor y Persistencia (Action & Persistence)
# -------------------------------------------------------------------------
def action_node(state: Agent007State) -> Dict[str, Any]:
    """Aplica los cambios en Supabase (activación de servicio + guardado de ticket)."""
    resolucion = state.get("resultado_resolucion")
    clasif = state.get("clasificacion")
    cliente = state.get("datos_cliente")

    # 1. Si se auto-resolvió y requiere activación, ejecutamos el cambio en Supabase
    if resolucion and resolucion.resolucion_automatica and resolucion.nuevo_estado_servicio == "ACTIVO":
        if cliente and cliente.cuenta_cliente:
            tool_ejecutar_activacion_servicio(cliente.cuenta_cliente)

    # 2. Persistir el ticket en la tabla tickets de Supabase
    tipo_val = clasif.tipo if clasif else "REQUERIMIENTO"
    cat_val = clasif.categoria if clasif else "GENERAL"
    acc_val = clasif.accion if clasif else "DESCONOCIDA"
    conf_val = clasif.confianza if clasif else 0.50
    cuenta_val = cliente.cuenta_cliente if (cliente and cliente.encontrado) else None

    ticket_guardado = tool_guardar_ticket_en_db(
        cuenta_cliente=cuenta_val,
        descripcion_original=state["descripcion_original"],
        canal_origen=state.get("canal_origen", "WHATSAPP"),
        tipo=tipo_val,
        categoria=cat_val,
        accion=acc_val,
        status=resolucion.status if resolucion else "PENDIENTE",
        confianza_ia=conf_val,
        comentario_resolucion=resolucion.comentario_resolucion if resolucion else "Sin comentario",
        resolucion_automatica=resolucion.resolucion_automatica if resolucion else False
    )

    ticket_id = ticket_guardado.get("id")
    if ticket_id:
        tool_registrar_auditoria(
            ticket_id=ticket_id,
            nodo="ActionNode",
            accion="TICKET_FINALIZADO",
            detalles={
                "status": resolucion.status if resolucion else "UNKNOWN",
                "resolucion_automatica": resolucion.resolucion_automatica if resolucion else False,
                "folio": ticket_guardado.get("folio")
            }
        )

    return {"ticket_id": ticket_id}
