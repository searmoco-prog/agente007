"""
Agente 007 - Motor de Orquestación Autónomo (Nivel 3: Tool Calling / Protocolo MCP)
Permite al LLM (Groq / Llama-3.3-70b) decidir autónomamente qué herramienta ejecutar
según el lenguaje natural del usuario, sin if/else rígidos.
"""

import json
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd

# Definición formal de herramientas bajo el estándar de Tool Calling / MCP
TOOLS_DEFINITIONS = [
    {
        "type": "function",
        "function": {
            "name": "generar_excel_personalizado",
            "description": "Genera y prepara para descarga un archivo Excel (.xlsx) personalizado. Úsalo SIEMPRE que el usuario pida un archivo Excel, documento, xlsx, descargar archivo, o requiera filtrar por una persona (ej: Ariel, Jesús), aplicar un color específico (azul, amarillo, verde, rojo, morado) o aplicar fórmulas vivas de concatenación.",
            "parameters": {
                "type": "object",
                "properties": {
                    "instruccion": {
                        "type": "string",
                        "description": "La instrucción detallada para generar el Excel (ej: 'extrae solo los de Ariel en azul con fórmula viva de concatenación')"
                    },
                    "entidad_filtro": {
                        "type": "string",
                        "description": "Nombre de la persona, acción o escenario a filtrar (ej: 'ariel', 'jesus', 'app', 'portabilidad')."
                    },
                    "color": {
                        "type": "string",
                        "enum": ["azul", "verde", "amarillo", "rojo", "morado", "naranja", "gris"],
                        "description": "Color solicitado para resaltar celdas. Default 'azul'."
                    },
                    "solo_coincidencias": {
                        "type": "boolean",
                        "description": "True si el usuario pidió exclusivamente esos registros ('solo ariel', 'únicamente jesus'). False si quiere toda la base resaltada."
                    }
                },
                "required": ["instruccion"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generar_grafica_operativa",
            "description": "Genera una visualización gráfica (gráfico de barras o desglose de regla). Úsalo cuando el usuario pida ver la gráfica, diagrama, visualización o gráfica de un criterio (ej: gráfica de Jesús vs Ariel).",
            "parameters": {
                "type": "object",
                "properties": {
                    "tipo": {
                        "type": "string",
                        "enum": ["oficial_general", "criterio_especifico"],
                        "description": "Tipo de gráfica. 'oficial_general' para el corte estándar de tickets. 'criterio_especifico' si pidió un desglose por personas/criterios."
                    },
                    "criterio_texto": {
                        "type": "string",
                        "description": "Detalle del criterio si aplica (ej: 'Jesús no atendidos excepto Ariel off')."
                    }
                },
                "required": ["tipo"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "consultar_kpis_en_vivo",
            "description": "Consulta las métricas operativas y estadísticas matemáticas exactas en tiempo real de la base de tickets (total de tickets, porcentaje resueltos, pendientes, top escenarios, conteo por personas). Úsalo cuando el usuario pregunte por cifras, números, estatus o resumen general.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filtro_especifico": {
                        "type": "string",
                        "description": "Filtro opcional sobre el cual consultar estadísticas (ej: 'jesus', 'ariel', 'app', 'portabilidad')."
                    }
                }
            }
        }
    }
]

def execute_tool_consultar_kpis(df: pd.DataFrame, filtro_especifico: Optional[str] = None) -> Dict[str, Any]:
    """Herramienta 1: Calcula KPIs y estadísticas exactas de la base en memoria."""
    total_general = len(df)
    resueltos_gen = int((df['status_normalizado'] == 'RESUELTO').sum())
    pendientes_gen = int((df['status_normalizado'] == 'PENDIENTE').sum())
    pct_gen = round((resueltos_gen / total_general * 100), 1) if total_general > 0 else 0
    
    top_escenarios = df['Escenario_Limpio'].value_counts().head(4).to_dict()

    resultado = {
        "total_tickets_base": total_general,
        "resueltos": resueltos_gen,
        "pendientes": pendientes_gen,
        "porcentaje_cumplimiento": f"{pct_gen}%",
        "top_escenarios": top_escenarios
    }

    if filtro_especifico:
        f_norm = filtro_especifico.lower().strip()
        mask = df.apply(lambda r: f_norm in str(r.to_dict()).lower(), axis=1)
        sub_df = df[mask]
        sub_total = len(sub_df)
        sub_res = int((sub_df['status_normalizado'] == 'RESUELTO').sum())
        sub_pen = int((sub_df['status_normalizado'] == 'PENDIENTE').sum())
        resultado["detalle_filtro"] = {
            "termino": filtro_especifico,
            "coincidencias": sub_total,
            "resueltos": sub_res,
            "pendientes": sub_pen
        }

    return resultado

async def decide_agent_action_mcp(
    pregunta: str, 
    df: pd.DataFrame, 
    chat_id: int, 
    api_key: str, 
    modelos: List[str],
    historial_chat: Dict[int, list]
) -> Optional[Dict[str, Any]]:
    """
    Orquestador MCP: El LLM evalúa la petición en lenguaje natural junto con
    el registro formal de herramientas disponibles y decide autónomamente la acción.
    
    Retorna:
    - dict con {"action": "generar_excel_personalizado"|"generar_grafica_operativa"|"consultar_kpis"|"chat_reply", ...}
    - None si no hay API key o hay error para delegar al fallback determinista.
    """
    if not api_key:
        return None

    try:
        from groq import AsyncGroq
        client = AsyncGroq(api_key=api_key)

        system_prompt = (
            "Eres el 'Agente 007', copiloto senior de operaciones de telecomunicaciones. "
            "Tienes a tu disposición un conjunto de herramientas operativas (MCP Tools). "
            "Analiza el mensaje de tu jefe y decide si requieres ejecutar una herramienta "
            "(como generar un Excel a medida, generar una gráfica, o calcular KPIs) "
            "o si solo debes responder cordialmente con información general. "
            "Si el usuario pide un Excel, documento, descargar archivo o filtrar personas (ej. Ariel, Jesús), "
            "DEBES invocar la herramienta generar_excel_personalizado con los parámetros adecuados."
        )

        messages = [
            {"role": "system", "content": system_prompt}
        ]

        # Añadir contexto reciente
        if chat_id and chat_id in historial_chat:
            messages.extend(historial_chat[chat_id][-4:])

        messages.append({"role": "user", "content": pregunta})

        for m in modelos:
            try:
                res = await client.chat.completions.create(
                    model=m,
                    messages=messages,
                    tools=TOOLS_DEFINITIONS,
                    tool_choice="auto",
                    temperature=0.1,
                    max_tokens=300
                )

                msg = res.choices[0].message
                
                # Caso 1: El LLM decidió invocar una o más herramientas (Tool Calling activo)
                if msg.tool_calls and len(msg.tool_calls) > 0:
                    tc = msg.tool_calls[0]
                    tool_name = tc.function.name
                    try:
                        tool_args = json.loads(tc.function.arguments)
                    except Exception:
                        tool_args = {}

                    return {
                        "action": tool_name,
                        "args": tool_args,
                        "llm_text": msg.content or ""
                    }

                # Caso 2: El LLM respondió directamente texto conversacional
                if msg.content:
                    return {
                        "action": "chat_reply",
                        "args": {},
                        "llm_text": msg.content
                    }
            except Exception:
                continue

    except Exception as e:
        print(f"Error en decide_agent_action_mcp: {e}", flush=True)

    return None
