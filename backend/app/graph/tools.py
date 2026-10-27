from typing import Optional, Dict, Any
from app.db.supabase import get_supabase_client

def tool_buscar_cliente(cuenta_cliente: str) -> Optional[Dict[str, Any]]:
    """
    Herramienta de cruce: Busca la información de contrato, saldo y estado del cliente en Supabase.
    """
    supabase = get_supabase_client()
    response = supabase.table("customers").select("*").eq("cuenta_cliente", cuenta_cliente).execute()
    if response.data and len(response.data) > 0:
        return response.data[0]
    return None

def tool_buscar_regla_catalogo(categoria: str, accion: str) -> Optional[Dict[str, Any]]:
    """
    Herramienta de cruce: Verifica en ticket_catalog si la acción está permitida y si es automatizable.
    """
    supabase = get_supabase_client()
    response = (
        supabase.table("ticket_catalog")
        .select("*")
        .ilike("categoria", f"%{categoria}%")
        .ilike("accion_sugerida", f"%{accion}%")
        .execute()
    )
    if response.data and len(response.data) > 0:
        return response.data[0]
    return None

def tool_ejecutar_activacion_servicio(cuenta_cliente: str) -> bool:
    """
    Herramienta de acción: Cambia el estado del servicio a 'ACTIVO' en la base de datos de Supabase.
    """
    supabase = get_supabase_client()
    response = (
        supabase.table("customers")
        .update({"estado_servicio": "ACTIVO"})
        .eq("cuenta_cliente", cuenta_cliente)
        .execute()
    )
    return bool(response.data)

def tool_guardar_ticket_en_db(
    cuenta_cliente: Optional[str],
    descripcion_original: str,
    canal_origen: str,
    tipo: str,
    categoria: str,
    accion: str,
    status: str,
    confianza_ia: float,
    comentario_resolucion: str,
    resolucion_automatica: bool
) -> Dict[str, Any]:
    """
    Herramienta de persistencia: Inserta o actualiza el registro en la tabla tickets de Supabase.
    """
    supabase = get_supabase_client()
    payload = {
        "cuenta_cliente": cuenta_cliente,
        "descripcion_original": descripcion_original,
        "canal_origen": canal_origen,
        "tipo": tipo,
        "categoria": categoria,
        "accion": accion,
        "status": status,
        "confianza_ia": confianza_ia,
        "comentario_resolucion": comentario_resolucion,
        "resolucion_automatica": resolucion_automatica
    }
    response = supabase.table("tickets").insert(payload).execute()
    if response.data and len(response.data) > 0:
        return response.data[0]
    return {}

def tool_registrar_auditoria(ticket_id: str, nodo: str, accion: str, detalles: Dict[str, Any]):
    """
    Herramienta de trazabilidad: Inserta en ticket_audit_logs.
    """
    try:
        supabase = get_supabase_client()
        supabase.table("ticket_audit_logs").insert({
            "ticket_id": ticket_id,
            "nodo_ejecutor": nodo,
            "accion_ejecutada": accion,
            "detalles": detalles
        }).execute()
    except Exception as e:
        print(f"Error registrando auditoría: {e}")
