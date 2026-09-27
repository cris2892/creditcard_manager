# 💳 Credit Card Manager (Gestor de Tarjetas y Cuotas)

Aplicación web desarrollada en **Python** y **Streamlit** para la gestión, control y proyección financiera de consumos en tarjetas de crédito a partir de resúmenes de cuenta en formato Excel.

---

## ✨ Características Principales

- 📥 **Carga Multi-Archivo y Multi-Pestaña**:
  - Soporte para subir uno o varios archivos Excel (`.xlsx`) simultáneamente.
  - Procesamiento automático de múltiples pestañas (hojas) dentro de un mismo libro de Excel.
- 🔍 **Escáner Inteligente de Encabezados**:
  - Detecta automáticamente dónde comienzan las columnas reales en el Excel aunque existan filas vacías o títulos al inicio.
  - Reconoce variantes y sinónimos de encabezados (`Monto`, `Importe`, `Concepto`, `Detalle`, `Cuota`, `Plan`).
- 💵 **Soporte de Formato de Moneda (Argentina / España)**:
  - Convierte automáticamente formatos como `$ 14.308,34`, `$ 5.903` o `$ -5.433,74` a valores numéricos precisos.
- 📅 **1er Vencimiento Personalizado por Tarjeta**:
  - Permite seleccionar individualmente el mes y año del primer vencimiento para cada tarjeta o pestaña cargada.
- 📊 **Proyección Futura de Cuotas y Totales**:
  - Simula el impacto de cuotas fijas a 3, 6, 12 o 24 meses.
  - Muestra una tabla detallada con fila de **Total Acumulado Proyectado**.
  - Paneles de métricas (`st.metric`) con el total pendiente general y por tarjeta.
  - Gráfico interactivo de barras apiladas generado con **Plotly**.
- ✅ **Gestión de Estado de Pagos**:
  - Posibilidad de marcar meses como **PAGADO** o **PENDIENTE** para pausar u omitir cuotas liquidadas.
- 🗑️ **Administración y Borrado de Datos**:
  - Permite eliminar registros de una tarjeta específica o realizar un reinicio completo de la base de datos local.

---

## 🛠️ Tecnologías Utilizadas

- **Lenguaje**: Python 3.10+
- **Interfaz Web**: [Streamlit](https://streamlit.io/)
- **Procesamiento de Datos**: [Pandas](https://pandas.pydata.org/), [OpenPyXL](https://openpyxl.readthedocs.io/)
- **Base de Datos**: SQLite3 (base local `finanzas_tarjetas.db`)
- **Visualización**: [Plotly Express](https://plotly.com/python/plotly-express/)
- **Control de Versiones**: Git & GitHub

---

## 🚀 Instalación y Ejecución

### 1. Clonar el repositorio
```bash
git clone https://github.com/cris2892/creditcard_manager.git
cd creditcard_manager
```

### 2. Crear y activar un entorno virtual (opcional pero recomendado)
```bash
# En Windows (PowerShell)
python -m venv .venv
.\.venv\Scripts\activate
```

### 3. Instalar las dependencias
```bash
pip install streamlit pandas openpyxl plotly python-dateutil
```

### 4. Iniciar la aplicación
```bash
streamlit run credit.py
```

---

## 📖 Modo de Uso

1. **Cargar Resúmenes**: Despliega la sección `📥 Cargar resúmenes` y sube tu(s) archivo(s) Excel.
2. **Configurar Vencimientos**: Para cada tarjeta/pestaña detectada, ajusta el nombre, el mes y el año del primer vencimiento.
3. **Procesar**: Haz clic en `🚀 Guardar / Procesar Todas las Tarjetas`.
4. **Analizar Proyecciones**: Ajusta el deslizador de meses proyectados para visualizar los montos a pagar futuros, los totales acumulados por tarjeta y el gráfico interactivo.
5. **Gestionar Pagos**: Marca los meses que vayas abonando para mantener tu proyección limpia y actualizada.
