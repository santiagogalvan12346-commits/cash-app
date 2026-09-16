# Sistema de Gestión de Efectivo — Multisede (Tucumán / Buenos Aires)

Evolución del sistema original de control de flujo de caja en efectivo,
ahora con soporte **multisede**, **trazabilidad de escaleras de pago por
proveedor** e **ingesta automática de planillas históricas** (formato
"bloques" de Buenos Aires).

---

## 1. Qué cambió respecto de la versión anterior

| Requerimiento | Qué se agregó |
|---|---|
| **1. Multisede** | Columna `Tesoreria` (`Tucumán` / `Buenos Aires`) en toda la base. Selector global (`st.segmented_control`) en la parte superior de la barra lateral: `Consolidado / Tucumán / Buenos Aires`. Todos los KPIs, la matriz semanal, los gráficos y el ABM responden a la sede seleccionada. |
| **2. Escaleras / presupuestos por proveedor** | Columna `Concepto / Presupuesto`. La Matriz Semanal se mantiene **siempre consolidada por proveedor** (una fila por proveedor por semana). El desglose se ve en un `st.expander` **"🔍 Ver desglose de escaleras por proveedor"** debajo de la matriz, con resumen por contrato, detalle fila a fila y barra de progreso. |
| **3. Ingesta de Buenos Aires** | Funciones `parse_bloques_tabulares`, `parse_reporte_semanal` e `ingest_buenos_aires_workbook` en `core.py`, más un formulario en la barra lateral para subir e importar la planilla histórica sin perder los datos ya cargados. |

**Compatibilidad hacia atrás garantizada:** si subís un Excel viejo (sin
`Tesoreria` ni `Concepto / Presupuesto`), el sistema completa
automáticamente `Tesoreria = 'Tucumán'` y `Concepto = ''` — no hace falta
tocar nada.

---

## 2. Contenido del proyecto

```
cash_app/
├── app.py             # Interfaz Streamlit (5 pestañas)
├── core.py            # Lógica de negocio + parsers de ingesta
├── data_seed.csv       # Datos consolidados: 183 pagos (Tucumán + Buenos Aires)
├── requirements.txt   # Dependencias
└── README.md          # Este documento
```

---

## 3. Instalación local

```bash
cd cash_app
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

> ⚠️ **Streamlit ≥ 1.36 es obligatorio**: es la versión que introduce
> `st.segmented_control`, usado en el selector de sede. Si ya tenés la
> app desplegada en Streamlit Community Cloud, verificá que el
> `requirements.txt` del repo incluya este piso de versión antes de
> redeployar — si no, el selector de sede fallará al cargar.

---

## 4. Requerimiento 1 — Multisede en detalle

- El selector de sede vive en `st.session_state.sede_global` y se aplica
  con `core.filter_by_sede(df, sede)` **antes** de cualquier filtro de la
  barra lateral (proveedor, estado, mes, fechas), de modo que esos
  filtros secundarios solo listan opciones relevantes a la sede activa.
- En el formulario ABM, el campo **"Tesorería / Sede"** viene
  pre-seleccionado según la sede activa en el selector global (si estás
  parado en "Buenos Aires", el próximo pago que cargues por defecto va a
  esa sede), pero se puede cambiar libremente pago por pago.
- El Panel Ejecutivo agrega un gráfico de torta "Distribución por sede"
  cuando estás en la vista **Consolidado**.
- `Kpis` ahora incluye `tc_promedio`: tipo de cambio promedio ponderado
  por importe, calculado solo sobre los pagos pactados en USD.

---

## 5. Requerimiento 2 — Escaleras / presupuestos por proveedor

- Nuevo campo en el ABM: **Concepto / Presupuesto** (texto libre —
  "Tramo 1 - Galpón", "Presupuesto N° 23062", "Obra Civil L2", etc.).
- La **Matriz Semanal sigue una única fila por proveedor**: si SCHMIDT
  tiene 3 escaleras corriendo en simultáneo, la celda de esa semana suma
  las 3. Esto es una regla de negocio estricta y no se puede desactivar
  desde la UI (si necesitás ver el desglose, usá el expander de abajo).
- El expander **"🔍 Ver desglose de escaleras por proveedor"**, debajo de
  la matriz, muestra:
  1. Resumen por Concepto/Presupuesto (compromiso, pagado, saldo, %
     avance, TC promedio si aplica).
  2. Barra de progreso del total pagado vs. saldo pendiente del
     proveedor completo.
  3. Detalle fila por fila: fecha, semana, concepto, moneda, TC,
     importe ARS, estado y observaciones.
- Funciones nuevas en `core.py`: `provider_summary_by_concept(df, proveedor)`
  y `provider_detail_rows(df, proveedor)`.

---

## 6. Requerimiento 3 — Ingesta de la planilla de Buenos Aires

### 6.1 Cómo se usa desde la UI
En la barra lateral, expander **"📁 Cargar / reemplazar datos"** →
sección **"Importar planilla histórica de Buenos Aires"**: subís el
`.xlsx`, confirmás el nombre de la hoja de bloques (por defecto
`" L2 (BA)"`, con el espacio inicial tal como viene en el archivo
original) y opcionalmente una hoja de respaldo semanal, y presionás
**"➕ Importar y agregar a la base"**. Los pagos se agregan a los ya
cargados; no se pierde nada.

### 6.2 Cómo funciona el parser (`core.parse_bloques_tabulares`)
La planilla real de Buenos Aires no es una grilla limpia: cada proveedor
ocupa un bloque de columnas (`FECHA` / `IMPORTE` / columna en blanco) y,
dentro de un mismo bloque, pueden convivir **varias escaleras
apiladas verticalmente**, cada una con su propia fila `TOTAL` intermedia
y a veces un nuevo encabezado `FECHA`/`IMPORTE` repetido más abajo (por
ejemplo, la columna de SCHMIDT o CEARG en la hoja `" L2 (BA)"`).

En lugar de asumir una estructura de encabezado fija, el parser:
1. Detecta el inicio de cada bloque por el nombre del proveedor en la
   primera fila.
2. Recorre **todas** las filas del bloque y captura como pago válido
   cualquier fila donde la primera columna sea una fecha y la segunda un
   número — así ignora automáticamente encabezados repetidos, subtítulos
   y "TOTAL" intermedios, sin importar cuántas escaleras haya.
3. Cada `TOTAL` detectado después de haber capturado al menos un pago
   incrementa el número de escalera (`Concepto = "Escalera 2 (import. BA)"`,
   etc.), separando automáticamente sub-contratos consecutivos.
4. Si encuentra una referencia de tipo de cambio (`TC`, `TC 1400`, o el
   par `TC` | `1465`), la anota en Observaciones como contexto — los
   importes de este formato ya están expresados en pesos finales
   (confirmado contra la hoja `SCHMIDT`, que reproduce el mismo total:
   $117.553.800), por lo que **no se vuelven a convertir**.

### 6.3 Reconciliación de nombres (`core.PROVIDER_ALIASES`)
Al comparar la hoja de bloques (detalle) contra la hoja `"Reporte"`
(usada como respaldo para proveedores sin desglose, ver 6.4), aparecen
variantes del mismo nombre: `"CARBALLO"` vs. `"CARBALLO SANTIAGO"`,
`"ACOSTA"` vs. `"ACOSTA ADRIAN"`, `"IG DE SANTIS"` (errata) vs.
`"ING. DE SANTIS"`. `core.PROVIDER_ALIASES` reconcilia estos casos para
que no se dupliquen como si fueran proveedores distintos. Si en tu
organización aparecen más variantes, alcanza con agregar una entrada al
diccionario.

### 6.4 Respaldo semanal (`core.parse_reporte_semanal`)
La hoja `"Reporte"` de la planilla es un pivote semana × proveedor ya
consolidado. Se usa **solo como respaldo**, para no perder el flujo de
proveedores que aparecen en ese reporte pero de los que la hoja de
bloques no trae ningún detalle itemizado — en el archivo de ejemplo,
esto aplicó a **BARRIONUEVO** y **MITRAL**. Esos pagos quedan cargados
como "Consolidado semanal (reporte, sin desglose)", fechados al lunes de
la semana correspondiente, y podés editarlos luego desde el ABM si
conseguís el detalle.

### 6.5 Limitación conocida
Si dentro de un mismo bloque conviven dos escaleras con **tipos de
cambio distintos**, el parser anota el último TC de referencia visto
como contexto de bloque, no por escalera individual. Para una atribución
perfecta en esos casos puntuales, revisá y ajustá manualmente la columna
`Concepto`/`Observaciones` de esas filas después de importar.

### 6.6 Regenerar el CSV consolidado desde cero
Si necesitás reconstruir `data_seed.csv` (por ejemplo, con una versión
más nueva de la planilla de Buenos Aires), corré:

```python
import pandas as pd
import core

# Base histórica de Tucumán ya en el esquema nuevo (con tesoreria/concepto)
tuc = core.load_seed_csv("data_seed_tucuman_legacy.csv")  # tu CSV viejo

# Ingesta de Buenos Aires desde el Excel original
ba = core.ingest_buenos_aires_workbook("INFORME_-_PAGOS_L2_BS_AS_-_OK.xlsx")

combinado = pd.concat([tuc[core.COLUMNS], ba[core.COLUMNS]], ignore_index=True)
combinado.to_csv("data_seed.csv", index=False)
```

---

## 7. Validación de los datos consolidados

El `data_seed.csv` incluido ya contiene la consolidación de ambas sedes,
validada contra las hojas originales:

| Tesorería | Pagos | Total ARS |
|---|---|---|
| Tucumán | 51 | $231.337.485 |
| Buenos Aires | 132 | $492.115.480 |
| **Total** | **183** | **$723.452.965** |

Desglose por proveedor (Buenos Aires) extraído de la hoja de bloques
`" L2 (BA)"` (validado 1:1 contra la hoja `SCHMIDT` para ese proveedor,
que reproduce exactamente $117.553.800):

| Proveedor | Pagos | Total ARS |
|---|---|---|
| PSI | 25 | $152.766.200 |
| SCHMIDT | 38 | $117.553.800 |
| CARBALLO SANTIAGO | 14 | $89.920.500 |
| BARRIONUEVO *(sin desglose, vía reporte)* | 7 | $35.000.000 |
| M&S INGENIERIA | 8 | $31.570.000 |
| CEARG | 16 | $26.893.400 |
| RAMAC | 9 | $17.337.140 |
| MITRAL *(sin desglose, vía reporte)* | 3 | $7.914.440 |
| ACOSTA ADRIAN | 6 | $6.720.000 |
| COLLAZO | 3 | $4.940.000 |
| ING. DE SANTIS | 3 | $1.500.000 |

> Nota: la hoja `"Reporte"` de la planilla original totaliza
> $376.770.480 para Buenos Aires — es menor que el total itemizado
> ($492.115.480) porque ese reporte no incluye todos los pagos con fecha
> más lejana (2027) que sí figuran en el detalle de escaleras. Se usó el
> detalle itemizado como fuente de verdad por ser más completo y
> granular ("todos los pagos desglosados", como pediste), y el reporte
> solo como respaldo para los 2 proveedores sin desglose disponible.

---

## 8. Formato de moneda y manejo de nulos

- `fmt_ars(valor)`: formato argentino sin decimales — `$ 231.337.500`.
- `fmt_ars_dec(valor)`: igual, con coma decimal — `$ 231.337.500,15`.
- Todo campo numérico nulo (`NaN`) se completa en `normalize_df`:
  `importe`→0.0, `tc`→1.0, `estado`→"Programado", `obs`→"",
  `tesoreria`→"Tucumán", `concepto`→"". Ningún cálculo (KPIs, matriz,
  control por proveedor) puede romperse por un campo faltante en el
  Excel de origen.

---

## 9. Despliegue en Streamlit Community Cloud

Sin cambios respecto de la versión anterior: subí la carpeta a GitHub,
conectá el repo en [share.streamlit.io](https://share.streamlit.io) y
apuntá a `app.py`. Confirmá que el `requirements.txt` del repo tenga
`streamlit>=1.36` antes de redeployar, por el uso de
`st.segmented_control`.

---

## 10. Notas técnicas

- **Sin placeholders**: toda función nueva en `core.py` y `app.py` está
  completamente implementada y probada contra los dos archivos Excel
  reales (Tucumán y Buenos Aires).
- **Modularidad preservada**: `core.py` sigue sin depender de Streamlit.
- **Matriz consolidada por regla de negocio**: `build_weekly_matrix`
  agrupa por `proveedor` sin importar cuántos `concepto` distintos
  tenga cada uno — es la regla estricta del Requerimiento 2.
- **Acumulado**: la matriz semanal ahora incluye una columna
  `ACUMULADO` (suma corrida de `TOTAL SEMANAL`), útil para proyectar
  necesidad de caja disponible a una fecha dada.
