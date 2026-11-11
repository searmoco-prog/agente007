"""
Catálogo Unificado de Atajos Rápidos y Operaciones de Analista
Compatible con Telegram Bot y VaultAnalyst (Streamlit)
"""

CATEGORIAS = {
    "limpieza": {
        "nombre": "Limpieza y Formato",
        "icono": "🧹",
        "descripcion": "Separar columnas, eliminar duplicados y normalizar textos"
    },
    "color": {
        "nombre": "Semáforos y Colores",
        "icono": "🎨",
        "descripcion": "Resaltar mejores vendedores, márgenes altos o alertas de ventas"
    },
    "tablas": {
        "nombre": "Tablas y Gráficas",
        "icono": "📊",
        "descripcion": "Tablas dinámicas, resúmenes agrupados y gráficos nativos"
    },
    "formulas": {
        "nombre": "Fórmulas y Cálculos",
        "icono": "🔢",
        "descripcion": "Totales automáticos, promedios, comisiones y BUSCARV"
    }
}

ATAJOS = {
    # 🧹 Limpieza y Formato
    "separar_nombres": {
        "categoria": "limpieza",
        "titulo": "✂️ Separar Nombre y Apellido",
        "instruccion": "Si existe una columna de nombre completo o cliente con dos o más palabras, sepárala en dos columnas: Nombre y Apellido",
        "icono": "✂️"
    },
    "quitar_duplicados": {
        "categoria": "limpieza",
        "titulo": "🗑️ Eliminar Filas Duplicadas",
        "instruccion": "Elimina todas las filas duplicadas de la base de datos manteniendo únicamente la primera aparición",
        "icono": "🗑️"
    },
    "formato_titulos": {
        "categoria": "limpieza",
        "titulo": "🔤 Formato Tipo Título",
        "instruccion": "Convierte todas las columnas de texto a formato Tipo Título (primera letra de cada palabra mayúscula) y elimina espacios innecesarios",
        "icono": "🔤"
    },

    # 🎨 Semáforos y Colores
    "top_margen_rojo": {
        "categoria": "color",
        "titulo": "🔴 Top Vendedor Mayor Margen",
        "instruccion": "pueds poner en rojo el vendedor con el mayor margen de ganancia",
        "icono": "🔴"
    },
    "verde_margen_35": {
        "categoria": "color",
        "titulo": "🟢 Resaltar Margen > 35%",
        "instruccion": "Calcula el margen y resalta en verde > 35%",
        "icono": "🟢"
    },
    "amarillo_margen_50": {
        "categoria": "color",
        "titulo": "🟡 Resaltar Margen > 50%",
        "instruccion": "puedes colorear de amarillo al vendedor que vendio arriba del 50 porciento",
        "icono": "🟡"
    },
    "alerta_bajas_ventas": {
        "categoria": "color",
        "titulo": "🟠 Alerta de Ventas Bajas",
        "instruccion": "Resalta en color naranja suave las ventas que estén por debajo del promedio general",
        "icono": "🟠"
    },

    # 📊 Tablas y Gráficas
    "tabla_grafico_circular": {
        "categoria": "tablas",
        "titulo": "🍩 Tabla Dinámica + Gráfico Circular",
        "instruccion": "Genera tabla dinámica y gráfico circular por Estado",
        "icono": "🍩"
    },
    "resumen_categoria": {
        "categoria": "tablas",
        "titulo": "📊 Resumen por Categoría con Barras",
        "instruccion": "Genera una tabla de resumen de ventas y ganancia agrupada por Categoría con un gráfico de barras nativo",
        "icono": "📊"
    },

    # 🔢 Fórmulas y Cálculos
    "buscador_buscarv": {
        "categoria": "formulas",
        "titulo": "🔍 Buscador BUSCARV / XLOOKUP",
        "instruccion": "Agrega módulo de búsqueda con BUSCARV / XLOOKUP",
        "icono": "🔍"
    },
    "fila_totales": {
        "categoria": "formulas",
        "titulo": "∑ Fila de Totales y Promedios",
        "instruccion": "Agrega al final de la tabla una fila con fórmulas SUM para ventas y costos, y AVERAGE para márgenes",
        "icono": "∑"
    },
    "columna_comision": {
        "categoria": "formulas",
        "titulo": "💵 Calcular Comisión (5%)",
        "instruccion": "Inserta una nueva columna llamada 'Comisión (5%)' calculada con fórmula de Excel como el 5% del Monto de Ventas",
        "icono": "💵"
    }
}

def get_atajos_por_categoria(cat_key: str) -> dict:
    """Retorna los atajos correspondientes a una categoría específica."""
    return {k: v for k, v in ATAJOS.items() if v["categoria"] == cat_key}
