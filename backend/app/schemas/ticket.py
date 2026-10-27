from typing import Optional, Literal
from pydantic import BaseModel, Field

class TicketIngress(BaseModel):
    """Esquema de entrada: Lo que envía WhatsApp, el técnico o la fila de Excel."""
    descripcion_original: str = Field(..., description="Texto libre del reporte o solicitud")
    cuenta_cliente: Optional[str] = Field(None, description="Cuenta o contrato si ya se conoce (ej. CLI-1001)")
    canal_origen: str = Field(default="WHATSAPP", description="Origen: WHATSAPP, EXCEL_BATCH, WEB")

class TicketClassification(BaseModel):
    """Esquema de tipificación ITIL extraído por el Agente 007."""
    cuenta_cliente: Optional[str] = Field(None, description="Número de cuenta detectado en el texto")
    tipo: Literal["INCIDENTE", "REQUERIMIENTO"] = Field(..., description="Clasificación ITIL")
    categoria: str = Field(..., description="Categoría detectada (ej. ACTIVACION, FACTURACION, CONECTIVIDAD)")
    accion: str = Field(..., description="Acción normalizada (ej. REALIZAR_ACTIVACION)")
    confianza: float = Field(..., ge=0.0, le=1.0, description="Nivel de certeza de la IA")
    razonamiento: str = Field(..., description="Explicación técnica de por qué se clasificó así")

class CustomerCrossData(BaseModel):
    """Datos obtenidos del cruce automático contra la tabla de clientes en Supabase."""
    encontrado: bool = False
    cuenta_cliente: Optional[str] = None
    nombre_completo: Optional[str] = None
    plan_servicio: Optional[str] = None
    estado_servicio: Optional[str] = None
    saldo_pendiente: float = 0.00

class TicketResolutionResult(BaseModel):
    """Resultado final de resolución y comentario formal generado por Agente 007."""
    status: Literal["AUTO_RESUELTO", "ESCALADO_HUMANO"] = Field(...)
    resolucion_automatica: bool = Field(...)
    comentario_resolucion: str = Field(..., description="Comentario técnico formal para el ticket")
    nuevo_estado_servicio: Optional[str] = Field(None, description="Si aplica cambio (ej. 'ACTIVO')")
