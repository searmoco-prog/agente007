from typing import TypedDict, Optional, List, Dict, Any
from app.schemas.ticket import TicketClassification, CustomerCrossData, TicketResolutionResult

class Agent007State(TypedDict):
    """
    Estado inmutable del Grafo de LangGraph para Agente 007.
    Viaja y se enriquece a través de cada nodo.
    """
    # 1. Ingesta
    ticket_id: Optional[str]
    descripcion_original: str
    cuenta_cliente_input: Optional[str]
    canal_origen: str
    
    # 2. Seguridad y Sanitización
    es_seguro: bool
    motivo_bloqueo_seguridad: Optional[str]
    
    # 3. Tipificación e Interpretación
    clasificacion: Optional[TicketClassification]
    
    # 4. Cruces de Datos en Supabase
    datos_cliente: Optional[CustomerCrossData]
    catalogo_regla: Optional[Dict[str, Any]]
    
    # 5. Resolución y Acción
    resultado_resolucion: Optional[TicketResolutionResult]
    
    # 6. Trazas de auditoría
    trazas_auditoria: List[Dict[str, Any]]
