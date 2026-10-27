from langgraph.graph import StateGraph, END
from app.graph.state import Agent007State
from app.graph.nodes import (
    sanitizer_node,
    classifier_node,
    cross_referencer_node,
    evaluator_node,
    action_node
)

def decidir_siguiente_paso(state: Agent007State) -> str:
    """Borde condicional: Si fue bloqueado por seguridad, salta directo a persistir la alerta."""
    if not state.get("es_seguro", True):
        return "action"
    return "classifier"

def build_agente007_graph():
    """Construye y compila el flujo del Agente 007 con LangGraph."""
    workflow = StateGraph(Agent007State)

    # 1. Registrar Nodos
    workflow.add_node("sanitizer", sanitizer_node)
    workflow.add_node("classifier", classifier_node)
    workflow.add_node("cross_referencer", cross_referencer_node)
    workflow.add_node("evaluator", evaluator_node)
    workflow.add_node("action", action_node)

    # 2. Definir Punto de Entrada
    workflow.set_entry_point("sanitizer")

    # 3. Bordes y Transiciones
    workflow.add_conditional_edges(
        "sanitizer",
        decidir_siguiente_paso,
        {
            "action": "action",
            "classifier": "classifier"
        }
    )
    workflow.add_edge("classifier", "cross_referencer")
    workflow.add_edge("cross_referencer", "evaluator")
    workflow.add_edge("evaluator", "action")
    workflow.add_edge("action", END)

    # 4. Compilar Grafo
    return workflow.compile()

# Instancia singleton lista para usar
agente007_app = build_agente007_graph()
