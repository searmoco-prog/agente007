import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib

# Configurar backend sin interfaz gráfica
matplotlib.use('Agg')

# Rutas
csv_path = 'C:/Users/LENOVO/.gemini/antigravity/scratch/agente007/backend/test_sheet.csv'
output_img = 'C:/Users/LENOVO/.gemini/antigravity/brain/723fa582-5af7-4484-af8b-b1e3cd08a89f/reporte_telecom_007.png'

# Cargar datos
df = pd.read_csv(csv_path)

# Normalizar columnas
col_resuelto = df.columns[4]
df['status_clean'] = df[col_resuelto].fillna('').astype(str).str.strip().str.lower()
df['status_normalizado'] = df['status_clean'].apply(lambda x: 'RESUELTO' if x == 'resuelto' else 'PENDIENTE')

# Métricas
total_tickets = len(df)
resueltos = (df['status_normalizado'] == 'RESUELTO').sum()
pendientes = (df['status_normalizado'] == 'PENDIENTE').sum()
pct_resuelto = (resueltos / total_tickets) * 100
pct_pendiente = (pendientes / total_tickets) * 100

# Top Escenarios (limpiando nombres de personas)
escenarios_validos = df['Escenario'].dropna().astype(str).str.upper().str.strip()
# Reemplazar typos
escenarios_validos = escenarios_validos.replace({
    'ACTIAVACION': 'ACTIVACIÓN',
    'PORTABILIDA': 'PORTABILIDAD',
    'CAMBIO DE NIR': 'CAMBIO NIR'
})
top_escenarios = escenarios_validos.value_counts().head(5)

# Top Acciones
acciones = df.iloc[:, 3].dropna().astype(str).str.strip()
top_acciones = acciones.value_counts().head(5)

# -------------------------------------------------------------
# DIBUJAR GRÁFICA ESTILO EXECUTIVE DASHBOARD (Estilo Moderno Dark)
# -------------------------------------------------------------
plt.style.use('dark_background')
fig = plt.figure(figsize=(14, 8), facecolor='#0f172a')

# Título Principal
fig.suptitle(
    f'AGENTE 007 - REPORTE DE OPERACIONES DE TELEFONÍA EN VIVO\nTotal de Tickets: {total_tickets} | Corte en Tiempo Real',
    fontsize=16,
    fontweight='bold',
    color='#38bdf8',
    y=0.96
)

# 1. Subplot Donut: Resueltos vs Pendientes
ax1 = fig.add_subplot(1, 3, 1, facecolor='#0f172a')
colores_donut = ['#10b981', '#f43f5e']
sizes = [resueltos, pendientes]
labels = [f'Resueltos\n{resueltos} ({pct_resuelto:.1f}%)', f'Pendientes\n{pendientes} ({pct_pendiente:.1f}%)']
wedges, texts = ax1.pie(
    sizes, 
    labels=labels, 
    colors=colores_donut, 
    startangle=140,
    textprops={'fontsize': 11, 'color': '#f8fafc', 'fontweight': 'bold'},
    wedgeprops={'width': 0.45, 'edgecolor': '#0f172a', 'linewidth': 2}
)
ax1.set_title('Estatus Global de Resolución', fontsize=13, color='#e2e8f0', pad=15, fontweight='bold')

# 2. Subplot Barras: Top Escenarios
ax2 = fig.add_subplot(1, 3, 2, facecolor='#0f172a')
bars1 = ax2.barh(top_escenarios.index[::-1], top_escenarios.values[::-1], color='#6366f1', height=0.6)
ax2.set_title('Top Escenarios de Falla / Solicitud', fontsize=13, color='#e2e8f0', pad=15, fontweight='bold')
ax2.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
for bar in bars1:
    width = bar.get_width()
    ax2.text(width + 2, bar.get_y() + bar.get_height()/2, f'{int(width)}',
             ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=10)
ax2.set_xlim(0, max(top_escenarios.values) * 1.2)
ax2.tick_params(colors='#94a3b8', labelsize=10)
for spine in ax2.spines.values():
    spine.set_color('#334155')

# 3. Subplot Barras: Top Acciones
ax3 = fig.add_subplot(1, 3, 3, facecolor='#0f172a')
# Truncar nombres muy largos para que quepan
etiquetas_acciones = [a[:22] + '...' if len(a) > 22 else a for a in top_acciones.index[::-1]]
bars2 = ax3.barh(etiquetas_acciones, top_acciones.values[::-1], color='#06b6d4', height=0.6)
ax3.set_title('Top Acciones Operativas', fontsize=13, color='#e2e8f0', pad=15, fontweight='bold')
ax3.grid(axis='x', linestyle='--', alpha=0.2, color='#94a3b8')
for bar in bars2:
    width = bar.get_width()
    ax3.text(width + 2, bar.get_y() + bar.get_height()/2, f'{int(width)}',
             ha='left', va='center', color='#f8fafc', fontweight='bold', fontsize=10)
ax3.set_xlim(0, max(top_acciones.values) * 1.2)
ax3.tick_params(colors='#94a3b8', labelsize=10)
for spine in ax3.spines.values():
    spine.set_color('#334155')

plt.tight_layout(rect=[0.03, 0.05, 0.97, 0.90])

# Guardar
os.makedirs(os.path.dirname(output_img), exist_ok=True)
plt.savefig(output_img, dpi=180, facecolor='#0f172a', bbox_inches='tight')
plt.close()
print("GRÁFICA GENERADA EXITOSAMENTE EN:", output_img)
