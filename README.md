# Sistema de Gestión de Efectivo — L2 Tucumán

Aplicación interactiva en **Python + Streamlit** para el control operativo,
registro (ABM) y proyección semanal del flujo de pagos en **efectivo** a
contratistas y proveedores del proyecto L2 Tucumán.

Construida a partir de la estructura del archivo original
`SISTEMA_GESTION_EFECTIVO_L2_TUC.xlsx` (hojas `DB_Pagos`,
`Control Proveedores` y `Matriz Semanal`).

---

## 1. Contenido del proyecto

```
cash_app/
├── app.py             # Interfaz Streamlit (5 módulos / pestañas)
├── core.py            # Lógica de negocio (cálculos, pivote, KPIs, export)
├── data_seed.csv       # Datos iniciales (51 pagos, abril 2026 – enero 2027)
├── requirements.txt   # Dependencias
└── README.md          # Este documento
```

`core.py` está desacoplado de la interfaz para poder testearlo o
reutilizarlo en otros contextos (scripts, notebooks, otra UI).

---

## 2. Instalación local

Requisitos: **Python 3.10 o superior**.

```bash
# 1. Clonar/copiar la carpeta del proyecto y ubicarse dentro
cd cash_app

# 2. (Recomendado) crear un entorno virtual
python3 -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la aplicación
streamlit run app.py
```

La aplicación se abre automáticamente en `http://localhost:8501`.

---

## 3. Funcionalidades

### 📊 Panel Ejecutivo
- KPIs: total comprometido, total pagado, saldo pendiente, desembolso promedio.
- Alertas automáticas de **picos de efectivo** (semanas > $10.000.000 ARS),
  para planificar el retiro logístico de la sucursal bancaria.
- Distribución mensual proyectada y progreso por proveedor (comprometido vs. pagado).

### 🗓️ Matriz Semanal
- Tabla pivote: semanas (lunes a viernes) x proveedores, con importe en ARS.
- Totalizadores automáticos por fila y por columna.
- Resalta en rojo las semanas que superan el umbral de $10.000.000 ARS.
- Respeta los filtros de la barra lateral (proveedor, estado, mes, rango de fechas).

### 📝 Registro de Pagos (ABM)
- Alta de nuevos pagos y edición/eliminación de existentes.
- Cálculo automático de `Lunes de Semana` y `Semana Etiqueta` a partir de la fecha.
- Selector de proveedor existente o alta de un proveedor nuevo.
- Si la moneda es **USD**, calcula automáticamente el equivalente en ARS
  usando el Tipo de Cambio ingresado.
- Estados: `Programado`, `Aprobado`, `Pagado`.

### 🏗️ Control por Proveedor
- Compromiso total, pagado, saldo, % del flujo total y % de avance por proveedor.
- Barras de progreso individuales.

### ⬇️ Exportar
- Descarga un `.xlsx` actualizado, respetando la estructura de 4 hojas del
  archivo original: `Resumen Ejecutivo`, `DB_Pagos`, `Matriz Semanal` y
  `Control Proveedores` (con formato, totalizadores y resaltado de picos).

### 📁 Carga de datos
- Desde la barra lateral se puede subir un `.xlsx` con la misma estructura
  de `DB_Pagos` del archivo original para **reemplazar** la base cargada,
  o restaurar los datos originales de la semilla incluida.

---

## 4. Persistencia de datos

La aplicación mantiene los datos en la **sesión de Streamlit**
(`st.session_state`), lo que permite editar y navegar sin perder cambios
durante el uso. Para conservar los datos entre sesiones o reinicios del
servidor:

1. Usar el botón **"⬇️ Descargar SISTEMA_GESTION_EFECTIVO_L2_TUC.xlsx"** en
   la pestaña *Exportar* al finalizar la sesión de trabajo.
2. La próxima vez, subir ese mismo archivo desde
   **"📁 Cargar / reemplazar datos"** en la barra lateral.

> Nota: si se requiere persistencia automática en base de datos (sin
> exportar/importar manualmente), se puede extender `core.py` para
> leer/escribir en SQLite, PostgreSQL o Google Sheets sin modificar `app.py`,
> ya que toda la lógica de datos está centralizada allí.

---

## 5. Despliegue

### Streamlit Community Cloud (gratuito)
1. Subir la carpeta `cash_app/` a un repositorio de GitHub.
2. Ingresar a [share.streamlit.io](https://share.streamlit.io), conectar el
   repositorio y seleccionar `app.py` como archivo principal.
3. Streamlit instalará automáticamente `requirements.txt` y publicará la app
   con una URL pública.

### Servidor propio / Docker
```bash
# Dockerfile mínimo (opcional, no incluido por defecto)
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

```bash
docker build -t cash-app .
docker run -p 8501:8501 cash-app
```

---

## 6. Datos incluidos (semilla)

`data_seed.csv` contiene los 51 desembolsos originales:

| Proveedor      | Pagos | Compromiso (ARS) |
|----------------|-------|-------------------|
| RODAS SERGIO   | 18    | $105.347.500      |
| BARRIONUEVO    | 16    | $100.428.000      |
| ANDRADA JORGE  | 7     | $13.662.000 (USD 9.200 a TC 1.485) |
| RODA TS        | 3     | $5.000.000        |
| RAZURI         | 7     | $6.900.000        |
| **Total**      | **51**| **$231.337.500**  |

`ANDRADA JORGE` se modela con **Moneda = USD** y **Tipo de Cambio = 1.485**,
de modo que el sistema recalcula automáticamente el equivalente en ARS
—tal como lo haría con cualquier otro proveedor pactado en dólares—, en
lugar de almacenar el monto ya convertido como si fuera un pago en pesos.

---

## 7. Notas técnicas

- **Sin placeholders**: toda función en `core.py` y `app.py` está
  completamente implementada y probada (carga, cálculo, pivote, export).
- **Modularidad**: la lógica de negocio (`core.py`) no depende de Streamlit,
  por lo que puede probarse con `pytest` o reutilizarse en otra interfaz.
- **Semana operativa**: se define de lunes a viernes; el "Lunes de Semana"
  se calcula restando el día de la semana (`weekday()`) a la fecha de pago.
- **Umbral de alerta**: configurable en `core.py`
  (`UMBRAL_PICO_EFECTIVO = 10_000_000`).
