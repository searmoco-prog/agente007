# 📋 BITÁCORA TÉCNICA DE PRUEBAS, CONTROL DE CALIDAD Y RESOLUCIÓN DE ERRORES
## Proyecto: Agente 007 - Telecom NOC & Universal Excel Analyst
**Versión:** 3.6 Enterprise Shield  
**Entorno de Producción:** Render Web Service (`https://agente007-61bf.onrender.com`)  
**Canal Oficial de Usuario:** Bot de Telegram (`@agente007_telecom_bot`)  
**Base de Datos & Auditoría:** Supabase PostgreSQL + Google Sheets en Vivo  
**Fecha de Certificación:** Noviembre 2026  

---

## 1. Resumen Ejecutivo de la Arquitectura
El **Agente 007** opera bajo una arquitectura híbrida con dos cerebros operativos independientes:
1. **Modo Telecom NOC 007 (`/telecom`):** Conectado en tiempo real a la base de 258 tickets de telecomunicaciones (Google Sheet Oficial), con generación de gráficas operativas, filtrado por criticidad y resolución multi-agente en LangGraph + Supabase.
2. **Modo Analista Universal de Datos (`/universal`):** Capaz de ingerir cualquier archivo corporativo (`.xlsx`, `.csv`), detectar automáticamente estructuras complejas con múltiples pestañas (`Dashboard`, `Base_Datos`, `Parametros`), calcular indicadores de negocio con fórmulas dinámicas vivas, aplicar formato condicional y generar gráficos incrustados nativos de Microsoft Excel sin depender de APIs de pago.

---

## 2. Matriz de Casos de Prueba Validados (Criterios de Aceptación)

| ID Prueba | Módulo / Criterio | Descripción del Requerimiento | Resultado Técnico Validado | Estado |
| :--- | :--- | :--- | :--- | :---: |
| **TC-01** | **Cita 1:** Margen de Ganancia | Calcular `% Margen de Ganancia` con fórmula viva `=IFERROR(Ganancia/Ventas, 0)` y aplicar formato condicional en **verde suave** (`#E2EFDA`) a celdas con margen **> 35%**. | Columna generada con fórmula viva. Formato condicional `CellIsRule` aplicado a rango `I2:I12`. Margen reflejado también en portada `Dashboard`. | ✅ **APROBADO** |
| **TC-02** | **Cita 2:** Tabla Dinámica + Gráfico | Resumir órdenes por Estado (`Completado`, `Pendiente`, `Cancelado`) con conteos (`COUNTIF`), totales (`SUMIF`) e incrustar gráfico circular nativo de Excel. | Creada pestaña dedicada `Resumen_Estado_Ordenes`. `PieChart` nativo insertado en celda `F4` con vinculación dinámica a los totales. | ✅ **APROBADO** |
| **TC-03** | **Cita 3:** Búsqueda Cruzada BUSCARV | Implementar módulo interactivo en portada para buscar por `ID Ventas` y traer en vivo Cliente, Monto y Estado mediante `=VLOOKUP(...)`. | Panel de búsqueda insertado en `Dashboard`. Fórmulas `=IFERROR(VLOOKUP(B18, 'Base_Datos'!A:H, Col, FALSE), ...)` 100% dinámicas. | ✅ **APROBADO** |
| **TC-04** | Modo Combo ("Todas") | Ejecutar Cita 1, Cita 2 y Cita 3 simultáneamente en una sola instrucción ("todas" o "aplica todo"). | Generación del libro integral con todas las pestañas, fórmulas, estilos corporativos azul marino y gráfico circular. | ✅ **APROBADO** |
| **TC-05** | Ingesta Multi-Pestaña | Leer un libro con portada vacía (`Dashboard`) y base de datos en hoja secundaria (`Base_Datos`). | Detección automática de la pestaña con datos reales y escaneo del offset de fila para encabezados. | ✅ **APROBADO** |
| **TC-06** | Resiliencia de Sesión | Evitar pérdida del archivo cargado cuando el servidor gratuito de Render se duerme por inactividad. | Implementado caché en disco local + re-descarga transparente vía `file_id` de Telegram / Supabase. | ✅ **APROBADO** |

---

## 3. Bitácora de Errores Encontrados, Causa Raíz y Solución Aplicada

### 🔴 Incidencia 1: Lectura de pestaña vacía (`Unnamed: 0..5`, 14 filas)
- **Síntoma reportado:** Al subir el archivo `Analisis_Ventas_Recargable(1).xlsx`, el bot reportaba: `Registros detectados: 14 filas. Columnas clave: Unnamed: 0, Unnamed: 1...` y no reconocía las ventas ni los clientes.
- **Causa Raíz:** Por defecto, `pandas.read_excel()` lee la primera hoja (índice 0, que correspondía a `Dashboard`). Esa hoja contenía un banner publicitario, tarjetas vacías y celdas combinadas en las primeras 15 filas, por lo que pandas creaba columnas sin nombre y leía filas en blanco.
- **Solución Técnica:** Se implementó en `backend/telegram_bot.py` un escáner inteligente con `openpyxl`:
  1. Inspecciona `wb.sheetnames`.
  2. Prioriza pestañas nombradas `Base_Datos`, `Datos`, `Data`, `Ventas` o `Registros`.
  3. Realiza un escaneo vertical de filas buscando la primera fila que contenga al menos 2 palabras clave de negocio (`id`, `venta`, `monto`, `costo`, `cliente`, `estado`).
  4. Pasa `header=header_row_idx` a pandas.

---

### 🔴 Incidencia 2: Incompatibilidad al subir archivo en formato `.csv`
- **Síntoma reportado:** El usuario subió inicialmente `Analisis_Ventas_Recargable(1).csv` y el cálculo de margen arrojó error interno.
- **Causa Raíz:** El formato `.csv` es texto plano separado por comas: no soporta múltiples pestañas (solo exporta la pestaña activa en Excel, que era la portada vacía), no almacena estilos, ni fórmulas vivas, ni gráficos. Además, `openpyxl.load_workbook()` solo admite binarios OpenXML (`.xlsx`), lanzando `InvalidFileException`.
- **Solución Técnica:** Se implementó un adaptador universal en `backend/universal_excel_agent.py`:
  - Detecta la firma binaria del archivo (`PK\x03\x04`).
  - Si es texto plano / CSV, crea dinámicamente un `Workbook` de openpyxl, transfiere los registros a una pestaña `Base_Datos`, y genera un archivo final `.xlsx` con todas las fórmulas habilitadas.

---

### 🔴 Incidencia 3: Peticiones cortas ("1", "la primera opción") caían a respuesta genérica de ChatGPT
- **Síntoma reportado:** El bot ofrecía las opciones 1, 2 y 3. Al responder el usuario *"La primera opción"*, el bot respondía: *"¡Claro! Para calcular el margen necesito que subas el archivo..."*.
- **Causa Raíz:** El filtro de intención buscaba palabras como `margen` o `rentabilidad`. Al recibir simplemente `"1"` o `"la primera opción"`, la condición no hacía match y el flujo caía en el fallback conversacional de IA general (`ask_ai_chat`), el cual no tenía contexto del archivo en memoria.
- **Solución Técnica:** Se construyó una capa previa de **Normalización Determinista de Intenciones**:
  - Mapea variaciones: `["1", "opcion 1", "la 1", "el 1", "la primera", "primera"]` ➡️ Ejecuta Cita 1.
  - Mapea variaciones: `["2", "opcion 2", "la 2", "la segunda", "segunda"]` ➡️ Ejecuta Cita 2.
  - Mapea variaciones: `["3", "opcion 3", "la 3", "la tercera", "tercera"]` ➡️ Ejecuta Cita 3.
  - Mapea variaciones: `["todas", "todo", "las 3", "aplica todo", "haz todo"]` ➡️ Ejecuta las 3 citas.
  - En `universal_excel_agent.py`, si ninguna coincide específicamente, el motor por defecto aplica las 3 transformaciones completas, asegurando que el usuario jamás reciba un archivo sin cambios.

---

### 🔴 Incidencia 4: Fallo de tipo al normalizar estatus (`AttributeError: 'float' object has no attribute 'lower'`)
- **Síntoma reportado:** Error de servidor `500` al procesar archivos con celdas de estado vacías.
- **Causa Raíz:** En `parse_uploaded_file()`, la columna de estado contenía valores `NaN` que pandas trata como punto flotante (`float`). Al ejecutar `s.lower()`, Python lanzaba `AttributeError`.
- **Solución Técnica:** Se blindó la función con `.fillna('')` y conversión explícita `str(s).lower()`, añadiendo compatibilidad con el estado corporativo `'completado'` para su mapeo a `'RESUELTO'`.

---

### 🔴 Incidencia 5: El archivo devuelto parecía idéntico / Formato no visible
- **Síntoma reportado:** *"El excel me lo entregó igual, no me entregó lo que le pedí"*.
- **Causa Raíz:** Doble factor técnico:
  1. **Foco de visualización al abrir:** Excel guarda en el XML qué pestaña está seleccionada (`tabSelected`). El libro original tenía seleccionada la pestaña `Dashboard`. Al abrirlo, el usuario veía la portada sin cambios, ya que la columna `% Margen de Ganancia` y el color verde estaban en la pestaña inferior `Base_Datos`.
  2. **Compatibilidad en visores móviles:** Aplicaciones de terceros en Android/iOS no recalculan reglas `CellIsRule` de OpenXML hasta que se fuerza una edición.
- **Solución Técnica:**
  - Se configuró `wb.active = ws_data` y `ws.views.sheetView[0].tabSelected = True` para forzar a Excel a abrir directamente en la hoja con los datos calculados.
  - Se sincronizó también la portada `Dashboard`: ahora añade la tarjeta KPI **`Margen Global`** (`=IFERROR(E6/C6, 0)`) en fila 5 y la columna `% Margen` en la tabla de categorías por producto.
  - Se validó el funcionamiento perfecto con la app oficial de **Microsoft Excel**.

---

### 🔴 Incidencia 6: Pérdida del archivo cargado por "spin-down" (apagado) de Render
- **Síntoma reportado:** Si pasaban más de 15 minutos entre la subida del archivo y la petición, el bot pedía subir el archivo de nuevo.
- **Causa Raíz:** En la capa gratuita de Render, los servidores entran en modo de suspensión tras inactividad. La memoria RAM (`DATAFRAME_SESION`) se reinicia a cero.
- **Solución Técnica:** Se creó el módulo `backend/session_storage.py` con persistencia en 3 capas:
  1. **Caché en disco local (`sessions_cache/`):** Almacena el binario `.xlsx` y los metadatos JSON por `chat_id`.
  2. **Persistencia en Supabase:** Registra el `chat_id`, `file_id` de Telegram y fecha de actualización.
  3. **Auto-restauración transparente:** Cuando entra un mensaje de texto, el bot verifica si la RAM está vacía; si lo está, restaura la sesión desde el disco o re-descarga el archivo desde el CDN de Telegram sin molestar al usuario.

---

## 4. Guía de Ejecución y Verificación Rápida para Usuarios

```text
Paso 1: Abre Telegram y entra al bot @agente007_telecom_bot
Paso 2: Arrastra o sube tu archivo 'Analisis_Ventas_Recargable(1).xlsx'
Paso 3: El bot confirmará:
        • 📑 Pestañas detectadas: Dashboard, Base_Datos, Parametros
        • 📊 Pestaña analizada: 'Base_Datos' (11 registros)
        • 🏷️ Columnas detectadas: ID Ventas, Fecha, Cliente, Monto Ventas ($), Costo ($), Estado...
Paso 4: Envía tu comando preferido:
        • Escribe '1' ➡️ Margen de Ganancia vivo + Verde > 35%
        • Escribe '2' ➡️ Pestaña de Resumen + Gráfico Circular Nativo
        • Escribe '3' ➡️ Buscador interactivo BUSCARV / XLOOKUP
        • Escribe 'todas' ➡️ Aplica las 3 transformaciones juntas
Paso 5: Descarga el archivo generado y ábrelo con Microsoft Excel oficial en PC o Celular.
```

---

## 5. Estado Final del Sistema
- **Compilación Python:** 100% libre de advertencias y errores sintácticos (`py_compile`).
- **Pipeline de Pruebas Unitarias:** TC-01 a TC-06 ejecutados y superados con éxito.
- **Repositorio:** Sincronizado en rama `main` en GitHub (`https://github.com/searmoco-prog/agente007.git`).
- **Despliegue:** Automático y operativo en Render Web Service.
