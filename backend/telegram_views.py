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
    cols_preview = ", ".join(columns[:5])
    if len(columns) > 5:
        cols_preview += f" (+{len(columns) - 5} más)"
    
    sheets_str = ", ".join(sheets) if sheets else "Hoja1"
    
    card = (
        "📥 **Archivo listo para trabajar**\n\n"
        f"📄 **Documento:** `{file_name}`\n"
        f"📈 **Registros:** `{total_filas} filas` • `{len(columns)} columnas`\n"
        f"📑 **Pestaña activa:** `{active_sheet}`\n"
        f"🏷️ **Columnas:** _{cols_preview}_\n\n"
        "👇 *Toca un atajo rápido abajo o escribe tu instrucción en el chat:*"
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
    """Diseña la tarjeta de confirmación de ejecución ejecutiva minimalista."""
    # Limpieza profunda de textos técnicos
    resumen_limpio = summary_detail
    for remove_str in [
        "🤖 **Transformación Universal Completada con IA Senior:**",
        "⚡ **Procesado con éxito por Agente 007 (LangGraph + Groq):**",
        "• Ruta ejecutada: `StateGraph -> CodeInterpreter -> Evaluator -> Done`",
        "• Privacidad Zero-Data Leak garantizada (ejecución 100% en RAM local).",
        "Cero fuga de datos: solo se enviaron los nombres de columnas a la IA; los datos se procesaron en RAM.",
        "• Se ejecutó código generado a medida para:"
    ]:
        resumen_limpio = resumen_limpio.replace(remove_str, "").strip()

    # Eliminar viñetas vacías o líneas repetitivas
    lineas = [l.strip() for l in resumen_limpio.splitlines() if l.strip() and not l.strip().startswith("• Petición resuelta:")]
    resumen_limpio = " ".join(lineas).strip()
    if not resumen_limpio:
        resumen_limpio = "Cambios aplicados exitosamente con fórmulas vivas y formatos."

    card = (
        "✨ **Transformación completada con éxito**\n\n"
        f"🎯 **Acción:** {titulo_accion}\n"
        f"📄 **Archivo:** `{file_name}` ({total_filas} registros)\n"
        f"📊 **Hoja:** `{sheet_name}`\n"
        f"✅ **Resultado:** {resumen_limpio}\n\n"
        "📁 *Tu archivo enriquecido está listo arriba 👆*"
    )
    return card

def format_categories_menu_card() -> str:
    """Texto descriptivo para el catálogo de categorías."""
    card = (
        "⚡ **Catálogo de Atajos Rápidos**\n\n"
        "Elige una categoría para ver sus opciones:\n\n"
        "🧹 **Limpieza:** Separar nombres, duplicados, espacios.\n"
        "🎨 **Semáforos:** Márgenes, top ventas, alertas de color.\n"
        "📊 **Tablas y Gráficas:** Dinámicas, donas, barras.\n"
        "🔢 **Cálculos:** Totales, promedios, BUSCARV / comisiones.\n"
        "⭐ **Pack 3 en 1:** Todas las mejoras juntas en 1 clic."
    )
    return card

def format_category_detail_card(cat_key: str) -> str:
    """Texto descriptivo para una categoría específica."""
    cat_info = CATEGORIAS.get(cat_key, {})
    nombre = cat_info.get("nombre", "Operaciones")
    desc = cat_info.get("descripcion", "")
    icono = cat_info.get("icono", "⚡")
    
    card = (
        f"{icono} **{nombre}**\n\n"
        f"_{desc}_\n\n"
        "👇 *Toca una acción para aplicarla a tu Excel:*"
    )
    return card

def format_help_card() -> str:
    """Tarjeta de ayuda y comandos para el usuario."""
    card = (
        "🤖 **Agente 007 • Guía de Uso**\n\n"
        "Puedes usarme de dos formas muy sencillas:\n\n"
        "1️⃣ **Con Atajos Rápidos:**\n"
        "• Toca **⚡ Atajos Rápidos** para elegir acciones por categoría.\n"
        "• Usa **⭐ Pack 3 en 1** para enriquecer todo en 1 clic.\n\n"
        "2️⃣ **Con Lenguaje Natural (Chat):**\n"
        "Escríbeme lo que quieres tal como lo piensas:\n"
        "• _\"pon en morado los vendedores de la ciudad de mexico\"_\n"
        "• _\"resalta en rojo los gastos mayores a 20,000\"_\n"
        "• _\"agrega una columna de comisión del 5%\"_\n\n"
        "📌 **Comandos útiles:**\n"
        "• `/atajos` ➡️ Abre el menú de botones rápidos.\n"
        "• `/reset` ➡️ Restablece el archivo al original."
    )
    return card
