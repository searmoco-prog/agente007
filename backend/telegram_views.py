"""
Vistas, Tarjetas Ejecutivas y Teclados Interactivos para Telegram Bot
Agente 007 v4.0 - Enterprise Executive UI
"""

from typing import List, Optional, Dict, Any
from shortcuts_config import CATEGORIAS, ATAJOS, get_atajos_por_categoria

def build_main_menu_keyboard(has_file: bool = False) -> dict:
    """Genera el teclado principal organizado por niveles ejecutivos."""
    if has_file:
        buttons = [
            [
                {"text": "⚡ Atajos Rápidos por Categoría", "callback_data": "menu_cats"},
            ],
            [
                {"text": "⭐ Aplicar Todo (Pack 3 en 1)", "callback_data": "shortcut:todas"},
                {"text": "📁 Descargar Excel", "callback_data": "btn_excel"}
            ],
            [
                {"text": "📊 Ver Gráfica", "callback_data": "btn_chart"},
                {"text": "🔄 Resumen Ejecutivo", "callback_data": "btn_summary"}
            ],
            [
                {"text": "❓ Ayuda y Comandos", "callback_data": "btn_help"},
                {"text": "🔄 Resetear Archivo", "callback_data": "btn_reset"}
            ]
        ]
    else:
        buttons = [
            [
                {"text": "📊 Ver Gráfica Oficial", "callback_data": "btn_chart"},
                {"text": "📁 Descargar Excel Oficial", "callback_data": "btn_excel"}
            ],
            [
                {"text": "🔄 Resumen del Turno", "callback_data": "btn_summary"},
                {"text": "❓ Ayuda y Guía", "callback_data": "btn_help"}
            ]
        ]
    return {"inline_keyboard": buttons}

def build_categories_keyboard() -> dict:
    """Genera el menú interactivo de categorías de atajos."""
    buttons = [
        [
            {"text": "🧹 Limpieza y Formato", "callback_data": "cat:limpieza"},
            {"text": "🎨 Semáforos y Colores", "callback_data": "cat:color"}
        ],
        [
            {"text": "📊 Tablas y Gráficas", "callback_data": "cat:tablas"},
            {"text": "🔢 Fórmulas y Cálculos", "callback_data": "cat:formulas"}
        ],
        [
            {"text": "⭐ Aplicar Pack 3 en 1", "callback_data": "shortcut:todas"}
        ],
        [
            {"text": "⬅️ Volver al Menú Principal", "callback_data": "menu_main"}
        ]
    ]
    return {"inline_keyboard": buttons}

def build_category_shortcuts_keyboard(cat_key: str) -> dict:
    """Genera el teclado con las acciones específicas de una categoría seleccionada."""
    cat_info = CATEGORIAS.get(cat_key, {})
    cat_atajos = get_atajos_por_categoria(cat_key)
    
    buttons = []
    # Añadir botones de atajos (de 1 en 1 o de 2 en 2 para que sean legibles en móvil)
    for key, atajo in cat_atajos.items():
        buttons.append([{"text": atajo["titulo"], "callback_data": f"shortcut:{key}"}])
    
    # Botones de navegación
    buttons.append([
        {"text": "📂 Ver Otras Categorías", "callback_data": "menu_cats"},
        {"text": "⬅️ Menú Principal", "callback_data": "menu_main"}
    ])
    
    return {"inline_keyboard": buttons}

def format_file_uploaded_card(file_name: str, total_filas: int, sheets: List[str], active_sheet: str, columns: List[str]) -> str:
    """Diseña la tarjeta ejecutiva cuando un usuario sube un archivo nuevo."""
    cols_preview = ", ".join(columns[:6])
    if len(columns) > 6:
        cols_preview += f" (+{len(columns) - 6} más)"
    
    sheets_str = ", ".join(sheets) if sheets else "Hoja1"
    
    card = (
        "╔═══════════════════════════════════╗\n"
        "   📥 ARCHIVO RECIBIDO Y BLINDADO   \n"
        "╚═══════════════════════════════════╝\n\n"
        f"📄 **Archivo:** `{file_name}`\n"
        f"📊 **Pestaña Activa:** `{active_sheet}`\n"
        f"📑 **Pestañas Detectadas:** {sheets_str}\n"
        f"📈 **Registros:** `{total_filas} filas`\n"
        f"🏷️ **Columnas:** _{cols_preview}_\n\n"
        "🛡️ **Higiene Zero-Leak:** Fórmulas evaluadas y datos preparados en RAM.\n\n"
        "👇 **Selecciona una acción rápida en los botones o escribe tu propia instrucción:**"
    )
    return card

def format_executive_report_card(
    titulo_accion: str,
    file_name: str,
    total_filas: int,
    sheet_name: str,
    summary_detail: str,
    graph_history: Optional[List[str]] = None
) -> str:
    """Diseña la tarjeta de confirmación de ejecución de alta calidad ejecutiva."""
    # Extraer los pasos clave del historial si existen
    pasos_clave = []
    if graph_history:
        for h in graph_history:
            if "hygiene_node" in h and "precalculadas" in h:
                pasos_clave.append("🧹 Higiene de datos y pre-cálculo")
            elif "router_node" in h and "Enrutado" in h:
                pasos_clave.append("🧭 Enrutamiento inteligente")
            elif "executor_evaluator_node" in h and "éxito" in h:
                pasos_clave.append("⚡ Ejecución en sandbox local")

    resumen_limpio = summary_detail.replace("🤖 **Transformación Universal Completada con IA Senior:**", "").strip()
    resumen_limpio = resumen_limpio.replace("⚡ **Procesado con éxito por Agente 007 (LangGraph + Groq):**", "").strip()

    card = (
        "╔═══════════════════════════════════╗\n"
        "   ⚡ TRANSFORMACIÓN EJECUTADA      \n"
        "╚═══════════════════════════════════╝\n\n"
        f"🎯 **Acción:** {titulo_accion}\n"
        f"📄 **Archivo:** `{file_name}`\n"
        f"📊 **Hoja Modificada:** `{sheet_name}`\n"
        f"📈 **Registros Evaluados:** `{total_filas}`\n\n"
        f"💡 **Resultado:**\n{resumen_limpio}\n\n"
        "📁 Tu archivo enriquecido con fórmulas vivas está listo abajo 👇"
    )
    return card

def format_categories_menu_card() -> str:
    """Texto descriptivo para el catálogo de categorías."""
    card = (
        "╔═══════════════════════════════════╗\n"
        "   ⚡ CATÁLOGO DE ATAJOS RÁPIDOS    \n"
        "╚═══════════════════════════════════╝\n\n"
        "Selecciona una categoría de operaciones para ver las acciones disponibles:\n\n"
        "• 🧹 **Limpieza y Formato:** Separar nombres, duplicados, espacios.\n"
        "• 🎨 **Semáforos y Colores:** Resaltar márgenes, top ventas, alertas.\n"
        "• 📊 **Tablas y Gráficas:** Dinámicas, resúmenes agrupados, donas.\n"
        "• 🔢 **Fórmulas y Cálculos:** Fila de totales, promedios, BUSCARV.\n"
        "• ⭐ **Pack 3 en 1:** Aplica las transformaciones más completas juntas."
    )
    return card

def format_category_detail_card(cat_key: str) -> str:
    """Texto descriptivo para una categoría específica."""
    cat_info = CATEGORIAS.get(cat_key, {})
    nombre = cat_info.get("nombre", "Operaciones")
    desc = cat_info.get("descripcion", "")
    
    card = (
        f"╔═══════════════════════════════════╗\n"
        f"   {cat_info.get('icono', '⚡')} {nombre.upper()}\n"
        f"╚═══════════════════════════════════╝\n\n"
        f"_{desc}_\n\n"
        "Toca una opción para aplicarla directamente a tu archivo:"
    )
    return card

def format_help_card() -> str:
    """Tarjeta de ayuda y comandos para el usuario."""
    card = (
        "╔═══════════════════════════════════╗\n"
        "   🤖 GUÍA DEL AGENTE 007 v4.0       \n"
        "╚═══════════════════════════════════╝\n\n"
        "Puedes interactuar conmigo de dos formas:\n\n"
        "1️⃣ **Con Botones y Atajos:**\n"
        "• Toca **⚡ Atajos Rápidos** para navegar por categorías sin escribir.\n"
        "• Usa **⭐ Aplicar Todo** para agregar semáforos, tabla dinámica y buscador en 1 clic.\n\n"
        "2️⃣ **Con Lenguaje Natural:**\n"
        "Escríbeme lo que necesitas en tus propias palabras:\n"
        "• _\"pon en rojo el vendedor con mayor margen\"_\n"
        "• _\"separa los nombres y apellidos\"_\n"
        "• _\"colorea en amarillo las ventas mayores a 50%\"_\n\n"
        "📌 **Comandos Disponibles:**\n"
        "• `/atajos` o `/menu` ➡️ Abre el catálogo de acciones rápidas.\n"
        "• `/reset` ➡️ Restablece hacia la base oficial en vivo.\n"
        "• `/modo` ➡️ Alterna entre Analista Universal y NOC Telecom."
    )
    return card
