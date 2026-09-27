import streamlit as st
import pandas as pd
import sqlite3
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import plotly.express as px

# Configuración de página
st.set_page_config(page_title="Gestor de Tarjetas y Cuotas", layout="wide")

# --- BASE DE DATOS LOCAL (SQLite) ---
conn = sqlite3.connect("finanzas_tarjetas.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS consumos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tarjeta TEXT,
    concepto TEXT,
    monto REAL,
    cuota_actual INTEGER,
    cuota_total INTEGER,
    primer_vencimiento TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS meses_pagados (
    tarjeta TEXT,
    mes_vencimiento TEXT,
    PRIMARY KEY (tarjeta, mes_vencimiento)
)
""")
conn.commit()

# --- FUNCIONES AUXILIARES ---
def limpiar_monto_es(val):
    """
    Convierte montos en formato de moneda de Argentina/España/Excel ($ 14.308,34 o $ 5.903)
    a valores numéricos float exactos en Python (14308.34 o 5903.0).
    """
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    s = re.sub(r'[^\d\,\.\-]', '', s)
    if not s:
        return 0.0
    
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace('.', '')
        
    try:
        return float(s)
    except ValueError:
        return 0.0

def parsear_cuota(val):
    if pd.isna(val) or str(val).strip() in ('', '-', 'nan', 'None'):
        return 1, 1
    val_str = str(val).strip()
    match = re.match(r"^(\d+)\s+de\s+(\d+)$", val_str, re.IGNORECASE)
    if match:
        return int(match.group(1)), int(match.group(2))
    try:
        num = int(float(val_str))
        return num, num
    except ValueError:
        return 1, 1

def fmt_moneda(val):
    return f"${val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

meses_nombres = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# --- INTERFAZ PRINCIPAL ---
st.title("💳 Control de Tarjetas, Cuotas y Vencimientos")

col_left, col_right = st.columns(2)

# --- SECCIÓN 1: CARGA DE RESÚMENES ---
with col_left:
    with st.expander("📥 Cargar nuevo resumen (Excel)", expanded=True):
        archivo = st.file_uploader("Sube el archivo Excel del resumen", type=["xlsx"])
        
        if archivo:
            df_upload = pd.read_excel(archivo)
            
            # Identificar columna Metodo / Tarjeta de forma dinámica
            col_metodo = next((c for c in df_upload.columns if 'metodo' in c.lower()), None)
            tarjeta_detectada = df_upload[col_metodo].dropna().iloc[0] if col_metodo and not df_upload[col_metodo].dropna().empty else "Visa Sin Nombre"
            
            col1, col2, col3 = st.columns(3)
            with col1:
                nombre_tarjeta = st.text_input("Identificador de tarjeta", value=str(tarjeta_detectada))
            with col2:
                mes_vto = st.selectbox("Mes de 1er Vencimiento", meses_nombres, index=datetime.now().month % 12)
            with col3:
                anio_vto = st.number_input("Año de 1er Vencimiento", min_value=2024, max_value=2035, value=datetime.now().year)
                
            fecha_primer_vto = f"{anio_vto}-{meses_nombres.index(mes_vto)+1:02d}-01"
            
            if st.button("Guardar / Actualizar Tarjeta"):
                cursor.execute("DELETE FROM consumos WHERE tarjeta = ?", (nombre_tarjeta,))
                
                col_cuota = next((c for c in df_upload.columns if 'cuota' in c.lower() or 'couta' in c.lower()), None)
                col_monto = next((c for c in df_upload.columns if 'monto' in c.lower()), None)
                col_concepto = next((c for c in df_upload.columns if 'concepto' in c.lower()), None)
                
                for _, row in df_upload.iterrows():
                    val_cuota = row[col_cuota] if col_cuota else ''
                    ca, ct = parsear_cuota(val_cuota)
                    
                    val_monto = row[col_monto] if col_monto else 0
                    monto_float = limpiar_monto_es(val_monto)
                    
                    val_concepto = str(row[col_concepto]) if col_concepto else ''
                    
                    cursor.execute("""
                        INSERT INTO consumos (tarjeta, concepto, monto, cuota_actual, cuota_total, primer_vencimiento)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (nombre_tarjeta, val_concepto, monto_float, ca, ct, fecha_primer_vto))
                    
                conn.commit()
                st.success(f"¡Resumen de {nombre_tarjeta} cargado con éxito!")
                st.rerun()

# --- SECCIÓN 2: OPCIONES BORRADO Y GESTIÓN ---
with col_right:
    with st.expander("🗑️ Eliminar / Gestionar Datos", expanded=False):
        df_existente = pd.read_sql("SELECT DISTINCT tarjeta FROM consumos", conn)
        tarjetas_cargadas = df_existente['tarjeta'].tolist() if not df_existente.empty else []
        
        if tarjetas_cargadas:
            st.markdown("### 🗑️ Borrar una tarjeta específica")
            tarjeta_a_eliminar = st.selectbox("Selecciona la tarjeta que deseas eliminar", tarjetas_cargadas, key="del_card")
            
            if st.button("❌ Eliminar esta tarjeta"):
                cursor.execute("DELETE FROM consumos WHERE tarjeta = ?", (tarjeta_a_eliminar,))
                cursor.execute("DELETE FROM meses_pagados WHERE tarjeta = ?", (tarjeta_a_eliminar,))
                conn.commit()
                st.success(f"Se eliminaron todos los registros de {tarjeta_a_eliminar}.")
                st.rerun()
                
            st.markdown("---")
            st.markdown("### ⚠️ Borrado Completo")
            confirmar = st.checkbox("Confirmar que deseo borrar TODOS los datos")
            if st.button("🔥 Borrar TODO y Reiniciar"):
                if confirmar:
                    cursor.execute("DELETE FROM consumos")
                    cursor.execute("DELETE FROM meses_pagados")
                    conn.commit()
                    st.success("Se eliminaron todos los datos de la base de datos.")
                    st.rerun()
                else:
                    st.warning("Debes marcar la casilla de confirmación para borrar todo.")
        else:
            st.info("No hay tarjetas cargadas actualmente para borrar.")

# --- CONSULTA DE DATOS ---
df_consumos = pd.read_sql("SELECT * FROM consumos", conn)
df_pagados = pd.read_sql("SELECT * FROM meses_pagados", conn)

if df_consumos.empty:
    st.info("Aún no tienes consumos cargados. Despliega la sección superior para subir tu primer Excel.")
    st.stop()

# --- SECCIÓN: MARCAR MESES PAGADOS ---
st.subheader("✅ Estado de Pagos")
tarjetas_activas = df_consumos['tarjeta'].unique()

col_t1, col_t2 = st.columns([2, 3])
with col_t1:
    tarjeta_sel = st.selectbox("Seleccionar Tarjeta", tarjetas_activas, key="sel_pay_card")
with col_t2:
    vto_base_str = df_consumos[df_consumos['tarjeta'] == tarjeta_sel]['primer_vencimiento'].iloc[0]
    vto_base = datetime.strptime(vto_base_str, "%Y-%m-%d")
    
    opciones_meses = [(vto_base + relativedelta(months=i)).strftime("%Y-%m") for i in range(12)]
    
    mes_a_cambiar = st.selectbox("Selecciona mes", opciones_meses, format_func=lambda x: f"{meses_nombres[int(x.split('-')[1])-1]} {x.split('-')[0]}", key="sel_pay_month")
    
    ya_pagado = not df_pagados[(df_pagados['tarjeta'] == tarjeta_sel) & (df_pagados['mes_vencimiento'] == mes_a_cambiar)].empty
    
    if ya_pagado:
        if st.button("Marcar como PENDIENTE (Deshacer pago)"):
            cursor.execute("DELETE FROM meses_pagados WHERE tarjeta = ? AND mes_vencimiento = ?", (tarjeta_sel, mes_a_cambiar))
            conn.commit()
            st.rerun()
    else:
        if st.button("Marcar como PAGADO"):
            cursor.execute("INSERT OR IGNORE INTO meses_pagados (tarjeta, mes_vencimiento) VALUES (?, ?)", (tarjeta_sel, mes_a_cambiar))
            conn.commit()
            st.rerun()

# --- CÁLCULO DE PROYECCIONES ---
st.markdown("---")
st.subheader("📊 Proyección de Pagos Futuros")

meses_a_proyectar = st.slider("Cantidad de meses a proyectar", min_value=3, max_value=24, value=6)

fechas_minimas = [datetime.strptime(f, "%Y-%m-%d") for f in df_consumos['primer_vencimiento'].unique()]
fecha_inicio = min(fechas_minimas)

proyecciones = []
for m in range(meses_a_proyectar):
    f_mes = fecha_inicio + relativedelta(months=m)
    f_key = f_mes.strftime("%Y-%m")
    nombre_col = f"{meses_nombres[f_mes.month - 1]} {f_mes.year}"
    
    fila = {"Mes Vencimiento": nombre_col, "key": f_key}
    total_pendiente = 0.0
    
    for t in tarjetas_activas:
        df_t = df_consumos[df_consumos['tarjeta'] == t]
        v_base = datetime.strptime(df_t['primer_vencimiento'].iloc[0], "%Y-%m-%d")
        delta = (f_mes.year - v_base.year) * 12 + (f_mes.month - v_base.month)
        
        pagado = not df_pagados[(df_pagados['tarjeta'] == t) & (df_pagados['mes_vencimiento'] == f_key)].empty
        
        if delta < 0 or pagado:
            monto = 0.0
        else:
            activos = df_t[df_t['cuota_actual'] + delta <= df_t['cuota_total']]
            monto = activos['monto'].sum()
            
        fila[t] = monto
        total_pendiente += monto
        
    fila["Total Pendiente"] = total_pendiente
    proyecciones.append(fila)

df_resultado = pd.DataFrame(proyecciones)

# Agregar Fila de TOTAL ACUMULADO al final de la tabla
fila_total = {"Mes Vencimiento": "TOTAL ACUMULADO PROYECTADO"}
for col in df_resultado.columns:
    if col not in ("Mes Vencimiento", "key"):
        fila_total[col] = df_resultado[col].sum()

df_con_total = pd.concat([df_resultado, pd.DataFrame([fila_total])], ignore_index=True)

# Formatear la tabla con totales
df_mostrar = df_con_total.drop(columns=["key"]).copy()
for col in df_mostrar.columns:
    if col != "Mes Vencimiento":
        df_mostrar[col] = df_mostrar[col].apply(fmt_moneda)

st.dataframe(df_mostrar, use_container_width=True)

# --- TARJETAS RESUMEN DE TOTALES EN LA PARTE INFERIOR ---
st.markdown("### 💰 Totales Acumulados por Tarjeta")

cols_metricas = st.columns(len(tarjetas_activas) + 1)
with cols_metricas[0]:
    total_general_proyectado = df_resultado["Total Pendiente"].sum()
    st.metric("Total General Pendiente", fmt_moneda(total_general_proyectado))

for idx, t in enumerate(tarjetas_activas):
    with cols_metricas[idx + 1]:
        monto_tarjeta = df_resultado[t].sum() if t in df_resultado.columns else 0.0
        st.metric(f"Total {t}", fmt_moneda(monto_tarjeta))

# Gráfico de barras apiladas
st.markdown("---")
tarjetas_cols = [t for t in tarjetas_activas if t in df_resultado.columns]
fig = px.bar(
    df_resultado,
    x="Mes Vencimiento",
    y=tarjetas_cols,
    title="Montos Pendientes por Tarjeta",
    labels={"value": "Monto ($)", "variable": "Tarjeta"},
    barmode="stack"
)
st.plotly_chart(fig, use_container_width=True)
