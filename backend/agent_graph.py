"""
Módulo Agente 007 - Orquestador LangGraph + LangChain
Estructura formal de grafo de estados con nodos de enrutamiento, herramientas deterministas,
generador de código IA (Groq) y bucle cíclico de auto-reparación (Self-Healing).
"""

import os
import io
import asyncio
from typing import TypedDict, Optional, Dict, Any, List
from dotenv import load_dotenv

# Cargar .env
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_THIS_DIR, ".env"), override=False)
load_dotenv(dotenv_path=os.path.join(_THIS_DIR, "..", ".env"), override=False)

from langgraph.graph import StateGraph, END
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_groq import ChatGroq

from universal_excel_agent import process_universal_excel
from universal_code_interpreter import (
    extract_metadata_schema,
    execute_safe_excel_transformation,
    clean_python_code_block,
    MODELOS_GROQ,
    GROQ_API_KEY
)

# 1. Definición del Estado del Grafo
class AgentState(TypedDict):
    file_bytes: bytes
    file_name: str
    instruction: str
    schema: Optional[Dict[str, Any]]
    route: Optional[str]              # "deterministic" o "code_interpreter"
    generated_code: Optional[str]
    error_message: Optional[str]
    retry_count: int
    result_bytes: Optional[bytes]
    output_filename: Optional[str]
    summary_message: Optional[str]
    status: str                       # "routing", "executing", "evaluating", "done", "failed"
    graph_history: List[str]          # Para registrar el camino tomado por el grafo

# 2. Nodos del Grafo

def router_node(state: AgentState) -> Dict[str, Any]:
    """Nodo 1: Analiza la petición y clasifica hacia qué rama del grafo debe ir."""
    inst_lower = state["instruction"].lower().strip()
    history = list(state.get("graph_history", []))
    history.append("router_node: Analizando intención del usuario")

    # Si es una de las opciones predefinidas clásicas (sin lógica personalizada compleja)
    es_preset_fijo = any(k in inst_lower for k in [
        "1", "2", "3", "todas", "aplicar todas",
        "tabla + gráfico circular", "buscador buscarv"
    ])

    # Si pide explícitamente solo semáforo estándar sin condiciones raras
    es_semaforo_estandar = ("semaforo" in inst_lower or "semáforo" in inst_lower) and not any(
        k in inst_lower for k in ["vendedor", "cliente", "fila", "columna especifica", "eliminar", "filtrar"]
    )

    if es_preset_fijo or es_semaforo_estandar:
        route = "deterministic"
        history.append("router_node -> Enrutado a [deterministic_node] (OpenPyXL nativo)")
    else:
        route = "code_interpreter"
        history.append("router_node -> Enrutado a [code_generator_node] (LangChain + Groq)")

    return {
        "route": route,
        "status": "executing",
        "graph_history": history
    }

def deterministic_node(state: AgentState) -> Dict[str, Any]:
    """Nodo 2: Ejecuta transformaciones fijas y deterministas de alta velocidad con OpenPyXL."""
    history = list(state.get("graph_history", []))
    history.append("deterministic_node: Aplicando fórmulas de Excel nativas")
    
    try:
        res_bytes, out_name, summary = process_universal_excel(
            state["file_bytes"],
            state["instruction"],
            state["file_name"]
        )
        history.append("deterministic_node: Transformación exitosa")
        return {
            "result_bytes": res_bytes,
            "output_filename": out_name,
            "summary_message": summary,
            "status": "done",
            "graph_history": history
        }
    except Exception as e:
        history.append(f"deterministic_node falló ({e}). Derivando a intérprete IA...")
        return {
            "route": "code_interpreter",
            "error_message": str(e),
            "status": "executing",
            "graph_history": history
        }

async def code_generator_node(state: AgentState) -> Dict[str, Any]:
    """Nodo 3: Genera código seguro en Python mediante LangChain + Groq con arquitectura Zero-Data Leak."""
    history = list(state.get("graph_history", []))
    history.append(f"code_generator_node: Generando script Python con IA (Intento {state.get('retry_count', 0) + 1})")
    
    schema = state.get("schema") or extract_metadata_schema(state["file_bytes"])
    instruction = state["instruction"]
    error_msg = state.get("error_message")
    retry_count = state.get("retry_count", 0)

    system_prompt = (
        "Eres el Motor de Código Autónomo del Agente 007 en LangGraph.\n"
        "Tu única tarea es generar un script en Python usando openpyxl y pandas para transformar un archivo Excel en memoria.\n"
        "Reglas estrictas:\n"
        "1. El archivo está cargado como objeto openpyxl llamado 'wb' (wb = openpyxl.load_workbook(io.BytesIO(file_bytes)) ya existe).\n"
        "2. NUNCA inventes columnas ni uses openpyxl.load_workbook ni wb.save() al final. El entorno ya las maneja.\n"
        "3. SIEMPRE inicializa todas las variables antes de usarlas en bucles o condicionales.\n"
        "4. Limpia strings numéricos (quita '$', ',' y espacios) antes de convertirlos a float.\n"
        "5. Devuelve ÚNICAMENTE el código Python puro, sin explicaciones ni markdown."
    )

    if error_msg and retry_count > 0:
        user_prompt = (
            f"El código previo falló con el siguiente error: {error_msg}\n"
            f"Estructura del Excel: {schema}\n"
            f"Corrige el código asegurando que no queden variables sin definir para cumplir: \"{instruction}\""
        )
    else:
        user_prompt = (
            f"Estructura de pestañas y columnas del archivo: {schema}\n"
            f"Instrucción del usuario: \"{instruction}\"\n"
            f"Genera el código Python para aplicarlo sobre 'wb'."
        )

    # Invocar Groq a través de LangChain
    code_generated = ""
    api_key = GROQ_API_KEY or os.getenv("GROQ_API_KEY", "")
    for modelo in MODELOS_GROQ:
        try:
            llm = ChatGroq(
                groq_api_key=api_key,
                model_name=modelo,
                temperature=0.1,
                max_tokens=2500
            )
            response = await llm.ainvoke([
                SystemMessage(content=system_prompt),
                HumanMessage(content=user_prompt)
            ])
            code_generated = clean_python_code_block(response.content)
            if code_generated:
                break
        except Exception as e_model:
            history.append(f"Modelo {modelo} no disponible: {e_model}")
            continue

    if not code_generated:
        raise RuntimeError("No fue posible generar código con ninguno de los modelos de LangChain/Groq configurados.")

    return {
        "schema": schema,
        "generated_code": code_generated,
        "retry_count": retry_count + 1,
        "status": "evaluating",
        "graph_history": history
    }

def executor_evaluator_node(state: AgentState) -> Dict[str, Any]:
    """Nodo 4: Ejecuta el código generado en RAM. Si hay error, activa el ciclo de auto-reparación."""
    history = list(state.get("graph_history", []))
    history.append("executor_evaluator_node: Ejecutando y evaluando código en Sandbox seguro")
    
    code_str = state.get("generated_code", "")
    file_bytes = state["file_bytes"]
    file_name = state["file_name"]
    instruccion = state["instruction"]

    try:
        res_bytes, out_filename = execute_safe_excel_transformation(file_bytes, code_str, file_name)
        history.append("executor_evaluator_node: Código ejecutado con éxito total ✅")
        
        summary = (
            f"⚡ **Procesado con éxito por Agente 007 (LangGraph + Groq):**\n"
            f"• Ruta ejecutada: `StateGraph -> CodeInterpreter -> Evaluator -> Done`\n"
            f"• Petición resuelta: *\"{instruccion}\"*\n"
            f"• Privacidad Zero-Data Leak garantizada (ejecución 100% en RAM local)."
        )
        return {
            "result_bytes": res_bytes,
            "output_filename": out_filename,
            "summary_message": summary,
            "error_message": None,
            "status": "done",
            "graph_history": history
        }
    except Exception as e:
        history.append(f"executor_evaluator_node detectó fallo: {e}")
        return {
            "error_message": str(e),
            "status": "evaluating",
            "graph_history": history
        }

# 3. Funciones de Enrutamiento Condicional (Edges)

def decide_next_route(state: AgentState) -> str:
    """Decide si ir al nodo determinista o al generador de código."""
    if state.get("route") == "deterministic":
        return "deterministic_node"
    return "code_generator_node"

def evaluate_execution_status(state: AgentState) -> str:
    """Decide si el grafo termina con éxito o regresa al generador para auto-corregirse (Self-Healing)."""
    if state.get("status") == "done":
        return "end"
    
    # Si falló y aún tenemos intentos de reintento (< 3)
    if state.get("retry_count", 0) < 3:
        return "retry"
    
    return "end"

# 4. Construcción y Compilación del StateGraph de LangGraph

def build_agent_graph():
    """Construye y compila el flujo formal de LangGraph."""
    workflow = StateGraph(AgentState)

    # Agregar nodos
    workflow.add_node("router_node", router_node)
    workflow.add_node("deterministic_node", deterministic_node)
    workflow.add_node("code_generator_node", code_generator_node)
    workflow.add_node("executor_evaluator_node", executor_evaluator_node)

    # Definir punto de entrada
    workflow.set_entry_point("router_node")

    # Conexiones condicionales desde el Router
    workflow.add_conditional_edges(
        "router_node",
        decide_next_route,
        {
            "deterministic_node": "deterministic_node",
            "code_generator_node": "code_generator_node"
        }
    )

    # El nodo determinista termina directo
    workflow.add_edge("deterministic_node", END)

    # Del generador de código pasa siempre al ejecutor/evaluador
    workflow.add_edge("code_generator_node", "executor_evaluator_node")

    # Conexión cíclica de Auto-Reparación desde el Evaluador
    workflow.add_conditional_edges(
        "executor_evaluator_node",
        evaluate_execution_status,
        {
            "end": END,
            "retry": "code_generator_node"   # Bucle de retorno auto-reparador
        }
    )

    return workflow.compile()

# Instancia global compilada del grafo
app_agent_graph = build_agent_graph()

# 5. Función de Ejecución Externa (Interfaz Unificada)

async def run_agent_graph(file_bytes: bytes, instruction: str, file_name: str = "archivo.xlsx") -> tuple[bytes, str, str, List[str]]:
    """
    Ejecuta el StateGraph completo de LangGraph de inicio a fin.
    Retorna: (result_bytes, output_filename, summary_message, graph_history)
    """
    initial_state: AgentState = {
        "file_bytes": file_bytes,
        "file_name": file_name,
        "instruction": instruction,
        "schema": None,
        "route": None,
        "generated_code": None,
        "error_message": None,
        "retry_count": 0,
        "result_bytes": None,
        "output_filename": None,
        "summary_message": None,
        "status": "routing",
        "graph_history": []
    }

    final_state = await app_agent_graph.ainvoke(initial_state)

    if final_state.get("result_bytes"):
        return (
            final_state["result_bytes"],
            final_state["output_filename"],
            final_state["summary_message"],
            final_state.get("graph_history", [])
        )
    else:
        err = final_state.get("error_message", "Error desconocido durante la ejecución del grafo.")
        raise RuntimeError(f"El Agente LangGraph no pudo completar la instrucción: {err}")
