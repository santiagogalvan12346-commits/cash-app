"""
core.py — Sistema de Gestión de Efectivo (multisede)
--------------------------------------------------------------------
Lógica de negocio, independiente de la capa de UI (Streamlit).

Novedades de esta versión respecto de la v1 (Tucumán-only):
  1. Soporte multisede: columna `Tesoreria` ('Tucumán' | 'Buenos Aires'),
     con compatibilidad hacia atrás (archivos viejos sin la columna se
     asumen 'Tucumán').
  2. Trazabilidad por escalera/presupuesto: columna `Concepto / Presupuesto`,
     con utilidades de resumen y detalle por proveedor.
  3. Ingesta de planillas históricas en formato de "bloques" por proveedor
     (FECHA / IMPORTE / TC en columnas repetidas), típico de la planilla
     de Buenos Aires, y de planillas "pivote semanal" (semana x proveedor).

La Matriz de Tesorería Semanal se mantiene SIEMPRE consolidada por
proveedor (una fila por proveedor por semana, sumando todos los
conceptos/escaleras que caigan en esa semana), tal como exige la regla
de negocio. El desglose por escalera se expone en una capa aparte
(`provider_summary_by_concept` / `provider_detail_rows`).
--------------------------------------------------------------------
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# Constantes de negocio
# ---------------------------------------------------------------------------

UMBRAL_PICO_EFECTIVO = 10_000_000  # ARS — disparo de alerta logística
ESTADOS_VALIDOS = ["Programado", "Aprobado", "Pagado"]
MONEDAS_VALIDAS = ["ARS", "USD"]
PROYECTO_DEFAULT = "L2 TUC"
TESORERIA_DEFAULT = "Tucumán"
TESORERIAS_VALIDAS = ["Tucumán", "Buenos Aires"]
SEDE_CONSOLIDADO = "Consolidado"
OPCIONES_SEDE = [SEDE_CONSOLIDADO] + TESORERIAS_VALIDAS

# Orden canónico de columnas "de trabajo" (antes de calcular campos derivados)
COLUMNS = [
    "id",
    "tesoreria",
    "fecha",
    "proveedor",
    "proyecto",
    "concepto",
    "moneda",
    "importe",
    "tc",
    "estado",
    "obs",
]


# ---------------------------------------------------------------------------
# Carga y normalización de datos
# ---------------------------------------------------------------------------

def load_seed_csv(path: str | Path) -> pd.DataFrame:
    """Carga el CSV semilla y normaliza tipos."""
    df = pd.read_csv(path)
    return normalize_df(df)


def load_from_uploaded_xlsx(file_like) -> pd.DataFrame:
    """
    Lee un archivo .xlsx subido por el usuario, reconociendo la hoja
    DB_Pagos con el esquema del sistema (compatible con archivos viejos
    que no tengan las columnas `Tesoreria` o `Concepto / Presupuesto`).
    """
    raw = pd.read_excel(file_like, sheet_name="DB_Pagos", header=3)
    raw = raw.dropna(subset=["ID Pago"])
    raw = raw[raw["ID Pago"].astype(str).str.upper() != "TOTAL EFECTIVO"]

    df = pd.DataFrame(
        {
            "id": raw["ID Pago"].astype(str),
            "tesoreria": raw["Tesoreria"] if "Tesoreria" in raw.columns else TESORERIA_DEFAULT,
            "fecha": pd.to_datetime(raw["Fecha Pago"]),
            "proveedor": raw["Proveedor"].astype(str).str.strip(),
            "proyecto": raw["Proyecto / Destino"].fillna(PROYECTO_DEFAULT) if "Proyecto / Destino" in raw.columns else PROYECTO_DEFAULT,
            "concepto": raw["Concepto / Presupuesto"].fillna("") if "Concepto / Presupuesto" in raw.columns else "",
            "moneda": raw["Moneda"].fillna("ARS").astype(str).str.upper(),
            "importe": pd.to_numeric(raw["Importe Moneda"], errors="coerce"),
            "tc": pd.to_numeric(raw["Tipo Cambio"], errors="coerce").fillna(1),
            "estado": raw["Estado"].fillna("Programado"),
            "obs": raw["Observaciones"].fillna("") if "Observaciones" in raw.columns else "",
        }
    )
    return normalize_df(df)


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Asegura tipos correctos y agrega columnas calculadas.
    Compatibilidad hacia atrás: si faltan `tesoreria` o `concepto`
    (archivos/CSV antiguos), se completan con sus valores por defecto.
    """
    df = df.copy()

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0.0)
    df["tc"] = pd.to_numeric(df["tc"], errors="coerce").fillna(1.0)
    df["moneda"] = df["moneda"].fillna("ARS").astype(str).str.upper().str.strip()
    df["proveedor"] = df["proveedor"].astype(str).str.strip()
    df["proyecto"] = df["proyecto"].fillna(PROYECTO_DEFAULT).astype(str).str.strip()
    df["estado"] = df["estado"].fillna("Programado").astype(str).str.strip()
    df["obs"] = df["obs"].fillna("").astype(str)

    # --- Compatibilidad hacia atrás: Tesoreria y Concepto ---
    df["tesoreria"] = df["tesoreria"].fillna(TESORERIA_DEFAULT).astype(str).str.strip()
    df.loc[~df["tesoreria"].isin(TESORERIAS_VALIDAS), "tesoreria"] = TESORERIA_DEFAULT
    df["concepto"] = df["concepto"].fillna("").astype(str).str.strip()

    df = compute_calculated_fields(df)
    df = df.sort_values("fecha").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Campos calculados: semana e importe ARS
# ---------------------------------------------------------------------------

def monday_of(fecha: pd.Timestamp) -> pd.Timestamp:
    """Devuelve el lunes de la semana calendario (lunes a viernes) de `fecha`."""
    if pd.isna(fecha):
        return pd.NaT
    return fecha - pd.Timedelta(days=fecha.weekday())


def week_label(lunes: pd.Timestamp) -> str:
    """Etiqueta legible de semana: 'DD/MM - DD/MM (viernes)'."""
    if pd.isna(lunes):
        return ""
    viernes = lunes + pd.Timedelta(days=4)
    return f"{lunes.strftime('%d/%m')} - {viernes.strftime('%d/%m')}"


def compute_importe_ars(row) -> float:
    """Importe ARS = importe si Moneda ARS, o importe * TC si Moneda USD."""
    if str(row["moneda"]).upper() == "USD":
        return round(float(row["importe"]) * float(row["tc"]), 2)
    return round(float(row["importe"]), 2)


def compute_calculated_fields(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega lunes_semana, semana_etiqueta, mes e importe_ars al DataFrame."""
    df = df.copy()
    df["lunes_semana"] = df["fecha"].apply(monday_of)
    df["semana_etiqueta"] = df["lunes_semana"].apply(week_label)
    df["mes"] = df["fecha"].dt.to_period("M").astype(str)
    df["importe_ars"] = df.apply(compute_importe_ars, axis=1)
    return df


def next_id(df: pd.DataFrame) -> str:
    """Genera el próximo ID correlativo con formato EFC-###."""
    if df.empty:
        return "EFC-001"
    nums = df["id"].astype(str).str.extract(r"(\d+)$")[0].dropna().astype(int)
    nxt = (nums.max() + 1) if not nums.empty else 1
    return f"EFC-{nxt:03d}"


# ---------------------------------------------------------------------------
# Filtro global de sede (Consolidado / Tucumán / Buenos Aires)
# ---------------------------------------------------------------------------

def filter_by_sede(df: pd.DataFrame, sede: str) -> pd.DataFrame:
    """Aplica el filtro global de sede. 'Consolidado' no filtra nada."""
    if sede == SEDE_CONSOLIDADO or not sede:
        return df
    return df[df["tesoreria"] == sede]


# ---------------------------------------------------------------------------
# Matriz de Tesorería Semanal (pivote, SIEMPRE consolidada por proveedor)
# ---------------------------------------------------------------------------

def build_weekly_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """
    Construye la matriz semana x proveedor con importes ARS, consolidando
    TODOS los conceptos/escaleras de cada proveedor en una única celda
    por semana (regla de negocio estricta). Incluye totalizadores por
    fila (semana), por columna (proveedor) y una columna de acumulado.
    """
    if df.empty:
        return pd.DataFrame()

    pivot = pd.pivot_table(
        df,
        index=["lunes_semana", "semana_etiqueta"],
        columns="proveedor",
        values="importe_ars",
        aggfunc="sum",
        fill_value=0.0,
    )
    pivot = pivot.sort_index(level="lunes_semana")
    pivot["TOTAL SEMANAL"] = pivot.sum(axis=1)
    pivot["ACUMULADO"] = pivot["TOTAL SEMANAL"].cumsum()

    pivot = pivot.reset_index()
    pivot = pivot.drop(columns=["lunes_semana"])
    pivot = pivot.rename(columns={"semana_etiqueta": "Semana"})

    total_row = pivot.drop(columns=["Semana", "ACUMULADO"]).sum(numeric_only=True)
    total_row["Semana"] = "TOTAL POR PROVEEDOR"
    total_row["ACUMULADO"] = pivot["ACUMULADO"].iloc[-1] if not pivot.empty else 0.0
    pivot = pd.concat([pivot, total_row.to_frame().T], ignore_index=True)
    return pivot


def weeks_over_threshold(df: pd.DataFrame, threshold: float = UMBRAL_PICO_EFECTIVO) -> pd.DataFrame:
    """Semanas cuyo total de efectivo supera el umbral logístico."""
    if df.empty:
        return pd.DataFrame(columns=["lunes_semana", "semana_etiqueta", "total_ars"])
    weekly = (
        df.groupby(["lunes_semana", "semana_etiqueta"])["importe_ars"]
        .sum()
        .reset_index()
        .rename(columns={"importe_ars": "total_ars"})
        .sort_values("lunes_semana")
    )
    return weekly[weekly["total_ars"] > threshold].reset_index(drop=True)


# ---------------------------------------------------------------------------
# KPIs ejecutivos y control por proveedor
# ---------------------------------------------------------------------------

@dataclass
class Kpis:
    total_comprometido: float
    total_pagado: float
    saldo_pendiente: float
    desembolso_promedio: float
    cantidad_desembolsos: int
    cantidad_pagados: int
    tc_promedio: float


def compute_tc_promedio(df: pd.DataFrame) -> float:
    """
    Tipo de cambio promedio ponderado por importe (en USD) de los pagos
    pactados en moneda extranjera. Devuelve 1.0 si no hay pagos en USD
    (todo el flujo está en pesos).
    """
    usd = df[df["moneda"] == "USD"]
    if usd.empty or usd["importe"].sum() == 0:
        return 1.0
    return round((usd["tc"] * usd["importe"]).sum() / usd["importe"].sum(), 2)


def compute_kpis(df: pd.DataFrame) -> Kpis:
    if df.empty:
        return Kpis(0, 0, 0, 0, 0, 0, 1.0)

    total_comprometido = df["importe_ars"].sum()
    pagados = df[df["estado"] == "Pagado"]
    total_pagado = pagados["importe_ars"].sum()
    saldo_pendiente = total_comprometido - total_pagado
    desembolso_promedio = df["importe_ars"].mean()

    return Kpis(
        total_comprometido=round(total_comprometido, 2),
        total_pagado=round(total_pagado, 2),
        saldo_pendiente=round(saldo_pendiente, 2),
        desembolso_promedio=round(desembolso_promedio, 2),
        cantidad_desembolsos=len(df),
        cantidad_pagados=len(pagados),
        tc_promedio=compute_tc_promedio(df),
    )


def provider_control(df: pd.DataFrame) -> pd.DataFrame:
    """Tabla de control integral por proveedor (compromiso, pagado, % avance)."""
    if df.empty:
        return pd.DataFrame()

    grp = df.groupby("proveedor")
    out = grp.agg(
        cantidad_pagos=("id", "count"),
        primer_pago=("fecha", "min"),
        ultimo_pago=("fecha", "max"),
        moneda_base=("moneda", lambda s: s.mode().iat[0] if not s.mode().empty else "ARS"),
        compromiso_total_ars=("importe_ars", "sum"),
    ).reset_index()

    pagado = (
        df[df["estado"] == "Pagado"]
        .groupby("proveedor")["importe_ars"]
        .sum()
        .rename("pagado_ars")
    )
    out = out.merge(pagado, on="proveedor", how="left")
    out["pagado_ars"] = out["pagado_ars"].fillna(0.0)
    out["saldo_ars"] = out["compromiso_total_ars"] - out["pagado_ars"]

    total_general = out["compromiso_total_ars"].sum()
    out["pct_del_flujo"] = (
        (out["compromiso_total_ars"] / total_general * 100).round(2) if total_general else 0.0
    )
    out["pct_avance"] = (out["pagado_ars"] / out["compromiso_total_ars"] * 100).fillna(0.0).round(2)

    out = out.sort_values("compromiso_total_ars", ascending=False).reset_index(drop=True)
    return out


def monthly_distribution(df: pd.DataFrame) -> pd.DataFrame:
    """Distribución mensual proyectada del total en ARS."""
    if df.empty:
        return pd.DataFrame(columns=["mes", "importe_ars"])
    out = df.groupby("mes")["importe_ars"].sum().reset_index()
    return out.sort_values("mes").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Trazabilidad por Concepto / Presupuesto (escaleras) — Requerimiento 2
# ---------------------------------------------------------------------------

def provider_summary_by_concept(df: pd.DataFrame, proveedor: str) -> pd.DataFrame:
    """
    Resumen de los contratos/escaleras vigentes de un proveedor: una fila
    por Concepto/Presupuesto con compromiso, pagado, saldo y % de avance.
    """
    sub = df[df["proveedor"] == proveedor].copy()
    if sub.empty:
        return pd.DataFrame()

    sub["concepto_mostrar"] = sub["concepto"].replace("", "Sin concepto asignado")

    grp = sub.groupby("concepto_mostrar")
    out = grp.agg(
        cantidad_pagos=("id", "count"),
        moneda=("moneda", lambda s: s.mode().iat[0] if not s.mode().empty else "ARS"),
        tc_promedio=("tc", "mean"),
        primer_pago=("fecha", "min"),
        ultimo_pago=("fecha", "max"),
        compromiso_ars=("importe_ars", "sum"),
    ).reset_index().rename(columns={"concepto_mostrar": "concepto"})

    pagado = (
        sub[sub["estado"] == "Pagado"]
        .groupby("concepto_mostrar")["importe_ars"]
        .sum()
        .rename("pagado_ars")
    )
    out = out.merge(pagado, left_on="concepto", right_index=True, how="left")
    out["pagado_ars"] = out["pagado_ars"].fillna(0.0)
    out["saldo_ars"] = out["compromiso_ars"] - out["pagado_ars"]
    out["pct_avance"] = (out["pagado_ars"] / out["compromiso_ars"] * 100).fillna(0.0).round(2)

    out = out.sort_values("compromiso_ars", ascending=False).reset_index(drop=True)
    return out


def provider_detail_rows(df: pd.DataFrame, proveedor: str) -> pd.DataFrame:
    """Detalle fila por fila (cada tramo/escalera) de un proveedor, ordenado por fecha."""
    sub = df[df["proveedor"] == proveedor].copy()
    if sub.empty:
        return pd.DataFrame()
    sub["concepto"] = sub["concepto"].replace("", "Sin concepto asignado")
    cols = [
        "id", "tesoreria", "fecha", "semana_etiqueta", "concepto", "moneda",
        "importe", "tc", "importe_ars", "estado", "obs",
    ]
    return sub[cols].sort_values("fecha").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Requerimiento 3 — Ingesta / normalización de planillas históricas
# ---------------------------------------------------------------------------

_TOTAL_MARKER_RE = re.compile(r"TOTAL", re.IGNORECASE)
_TC_MARKER_RE = re.compile(r"^\s*TC\s*(\d{3,6})?\s*$", re.IGNORECASE)
_TC_NUMBER_RE = re.compile(r"(\d{3,6})")


def parse_bloques_tabulares(
    ws: Worksheet,
    tesoreria: str = "Buenos Aires",
    proyecto: str = "L2 BA",
    estado_default: str = "Programado",
    header_row: int = 1,
    block_width: int = 3,
) -> pd.DataFrame:
    """
    Normaliza el formato histórico "por bloques" en el que cada proveedor
    ocupa un grupo de columnas (típicamente FECHA / IMPORTE / [columna en
    blanco]), repetido horizontalmente a lo largo de la hoja — el formato
    de la planilla de Buenos Aires (hoja " L2 (BA)").

    Estrategia (deliberadamente tolerante a la variabilidad del formato
    real, que mezcla encabezados, sub-totales ("TOTAL", "TOTAL USD") y
    referencias de tipo de cambio en distintas filas dentro de un mismo
    bloque):
      - Se detectan los bloques a partir de las celdas no vacías de
        `header_row` (nombre del proveedor).
      - Dentro de cada bloque, se recorren TODAS las filas y se capturan
        como pago válido únicamente aquellas donde la primera columna del
        bloque es una fecha y la segunda es un número — así se ignoran
        automáticamente encabezados repetidos, subtítulos y filas "TOTAL"
        intermedias, sin importar cuántas "escaleras" (sub-contratos)
        haya apiladas verticalmente en el mismo bloque.
      - Cada vez que se detecta una fila con la palabra "TOTAL" después de
        haber capturado al menos un pago, se incrementa el número de
        "escalera" para el resto de las filas del bloque — esto separa
        automáticamente sub-contratos consecutivos dentro de una misma
        columna (p. ej. Concepto = "Escalera 1", "Escalera 2", ...).
      - Si se encuentra una celda de referencia de tipo de cambio (p. ej.
        "TC", "TC 1400", o el par "TC" | 1465), se anota como contexto en
        Observaciones. Los importes de este formato representan pesos ya
        liquidados (no se re-convierten), por lo que se cargan con
        Moneda='ARS' y TC=1 salvo que el valor puntual sí represente una
        conversión explícita fila a fila (no es el caso de este formato).

    Devuelve un DataFrame con el esquema de trabajo interno (ver
    `core.COLUMNS`), listo para concatenar y pasar por `normalize_df`.

    Limitación conocida: si dentro de un mismo bloque coexisten DOS
    escaleras con tipos de cambio distintos, el TC de referencia
    detectado se anota en observaciones a nivel de bloque, no por
    escalera individual — para una atribución perfecta por escalera se
    recomienda revisar y ajustar la columna `concepto`/`obs` generada.
    """
    max_row = ws.max_row
    max_col = ws.max_column

    block_starts = [
        c for c in range(1, max_col + 1)
        if ws.cell(row=header_row, column=c).value not in (None, "")
    ]

    records: list[dict] = []

    for i, sc in enumerate(block_starts):
        raw_name = str(ws.cell(row=header_row, column=sc).value).strip()
        # separa calificadores tipo "CARBALLO SANTIAGO - L2 - 50%" -> nombre + nota
        nombre_match = re.split(r"\s+-\s+", raw_name)
        proveedor_nombre = nombre_match[0].strip()
        nota_calificador = raw_name[len(proveedor_nombre):].strip(" -") if len(nombre_match) > 1 else ""

        col_a, col_b = sc, sc + 1
        escalera = 1
        seen_data_in_escalera = False
        tc_contexto: Optional[float] = None
        usd_contexto: Optional[float] = None

        for r in range(header_row + 1, max_row + 1):
            va = ws.cell(row=r, column=col_a).value
            vb = ws.cell(row=r, column=col_b).value

            if isinstance(va, dt.datetime) and isinstance(vb, (int, float)):
                obs_parts = []
                if nota_calificador:
                    obs_parts.append(nota_calificador)
                if tc_contexto:
                    obs_parts.append(f"Ref. TC planilla: {tc_contexto:g}")
                if usd_contexto:
                    obs_parts.append(f"Ref. contrato USD: {usd_contexto:g}")

                records.append({
                    "id": None,
                    "tesoreria": tesoreria,
                    "fecha": pd.Timestamp(va),
                    "proveedor": proveedor_nombre,
                    "proyecto": proyecto,
                    "concepto": f"Escalera {escalera} (import. BA)",
                    "moneda": "ARS",
                    "importe": float(vb),
                    "tc": 1.0,
                    "estado": estado_default,
                    "obs": " | ".join(obs_parts),
                })
                seen_data_in_escalera = True
                continue

            # referencia de TC: "TC" | valor, o "TC 1400" en una sola celda
            if isinstance(va, str):
                va_up = va.strip().upper()
                if _TC_MARKER_RE.match(va_up):
                    m = _TC_NUMBER_RE.search(va_up)
                    if m:
                        tc_contexto = float(m.group(1))
                    elif isinstance(vb, (int, float)):
                        tc_contexto = float(vb)
                elif va_up.startswith("TOTAL USD") and isinstance(vb, (int, float)):
                    usd_contexto = float(vb)
                elif _TOTAL_MARKER_RE.search(va_up) and seen_data_in_escalera:
                    escalera += 1
                    seen_data_in_escalera = False

    df = pd.DataFrame(records, columns=COLUMNS)
    if df.empty:
        return df
    df["id"] = [f"BA-{i+1:03d}" for i in range(len(df))]
    return df


def parse_reporte_semanal(
    ws: Worksheet,
    tesoreria: str = "Buenos Aires",
    proyecto: str = "L2 BA",
    estado_default: str = "Programado",
    header_row: int = 1,
    id_prefix: str = "BAW",
) -> pd.DataFrame:
    """
    Normaliza una hoja "pivote semanal" (columna `SEMANA` con etiquetas
    tipo '07-11 Sep', y una columna por proveedor con el importe ARS de
    esa semana; columna final opcional 'TOTAL SEMANAL' que se ignora).

    Cada celda no nula/no cero se transforma en un pago consolidado,
    fechado al lunes de esa semana (se toma el primer día mencionado en
    la etiqueta). Útil como *fallback* para proveedores de los que solo
    se dispone de un reporte agregado (sin desglose de escaleras),
    evitando perder el dato de flujo aunque no exista detalle.
    """
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    try:
        col_semana = headers.index("SEMANA") + 1
    except ValueError:
        col_semana = 1

    proveedor_cols = {
        c + 1: str(h).strip()
        for c, h in enumerate(headers)
        if h not in (None, "") and str(h).strip().upper() not in ("SEMANA", "TOTAL SEMANAL")
    }

    records: list[dict] = []
    year_hint = dt.date.today().year

    for r in range(header_row + 1, ws.max_row + 1):
        semana_label = ws.cell(row=r, column=col_semana).value
        if not semana_label or str(semana_label).strip().upper() == "TOTAL":
            continue
        fecha = _parse_week_label_to_monday(str(semana_label), year_hint)
        if fecha is None:
            continue

        for c, proveedor in proveedor_cols.items():
            val = ws.cell(row=r, column=c).value
            if isinstance(val, (int, float)) and val:
                records.append({
                    "id": None,
                    "tesoreria": tesoreria,
                    "fecha": fecha,
                    "proveedor": proveedor,
                    "proyecto": proyecto,
                    "concepto": "Consolidado semanal (reporte, sin desglose)",
                    "moneda": "ARS",
                    "importe": float(val),
                    "tc": 1.0,
                    "estado": estado_default,
                    "obs": f"Semana {semana_label}",
                })

    df = pd.DataFrame(records, columns=COLUMNS)
    if df.empty:
        return df
    df["id"] = [f"{id_prefix}-{i+1:03d}" for i in range(len(df))]
    return df


_MESES_ABBR = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    "dec": 12,
}


def _parse_week_label_to_monday(label: str, year_hint: int) -> Optional[pd.Timestamp]:
    """
    Convierte una etiqueta de semana tipo '07-11 Sep' o '28 Sep - 02 Oct'
    en la fecha del lunes correspondiente. Best-effort: usa el primer
    día/mes mencionado en la etiqueta.
    """
    m = re.search(r"(\d{1,2})\s*(?:-\s*\d{1,2}\s*)?([A-Za-z]{3,4})", label)
    if not m:
        return None
    day = int(m.group(1))
    mes_txt = m.group(2).strip(".").lower()[:3]
    month = _MESES_ABBR.get(mes_txt)
    if not month:
        return None
    try:
        return pd.Timestamp(year=year_hint, month=month, day=day)
    except ValueError:
        return None


# Distintas hojas de una misma planilla suelen nombrar al mismo proveedor
# de forma ligeramente distinta (nombre completo vs. apellido solo, o
# directamente con una errata). Este mapa reconcilia esos casos para
# evitar registrar al mismo proveedor dos veces bajo nombres distintos.
PROVIDER_ALIASES = {
    "CARBALLO": "CARBALLO SANTIAGO",
    "ACOSTA": "ACOSTA ADRIAN",
    "IG DE SANTIS": "ING. DE SANTIS",
}


def canonicalize_provider_name(name: str) -> str:
    """Resuelve alias conocidos de nombres de proveedor a una forma canónica."""
    if not isinstance(name, str):
        return name
    key = name.strip().upper()
    return PROVIDER_ALIASES.get(key, name.strip())


def ingest_buenos_aires_workbook(
    path_or_buffer,
    hoja_bloques: str = " L2 (BA)",
    hoja_reporte: Optional[str] = "Reporte",
    tesoreria: str = "Buenos Aires",
    proyecto: str = "L2 BA",
) -> pd.DataFrame:
    """
    Pipeline de ingesta de alto nivel para la planilla histórica de
    Buenos Aires:
      1. Extrae el detalle itemizado (escaleras) desde la hoja de bloques
         (`parse_bloques_tabulares`).
      2. Si se indica `hoja_reporte`, la usa como *fallback* para agregar
         proveedores que aparecen en el reporte semanal consolidado pero
         de los que NO hay detalle itemizado en la hoja de bloques
         (evita perder flujo de proveedores sin desglose disponible).
         Los nombres de proveedor se reconcilian con `PROVIDER_ALIASES`
         antes de comparar, para no duplicar a un mismo proveedor que
         aparece nombrado distinto en cada hoja (p. ej. "CARBALLO" en el
         reporte semanal vs. "CARBALLO SANTIAGO" en el detalle itemizado).
      3. Devuelve un único DataFrame combinado, en el esquema de trabajo
         interno, listo para pasar por `normalize_df` y concatenar con la
         base existente.
    """
    wb = load_workbook(path_or_buffer, data_only=True)

    detalle = pd.DataFrame(columns=COLUMNS)
    if hoja_bloques in wb.sheetnames:
        detalle = parse_bloques_tabulares(wb[hoja_bloques], tesoreria=tesoreria, proyecto=proyecto)
        if not detalle.empty:
            detalle["proveedor"] = detalle["proveedor"].apply(canonicalize_provider_name)

    proveedores_con_detalle = set(detalle["proveedor"].str.upper()) if not detalle.empty else set()

    fallback = pd.DataFrame(columns=COLUMNS)
    if hoja_reporte and hoja_reporte in wb.sheetnames:
        fallback_full = parse_reporte_semanal(wb[hoja_reporte], tesoreria=tesoreria, proyecto=proyecto)
        if not fallback_full.empty:
            fallback_full["proveedor"] = fallback_full["proveedor"].apply(canonicalize_provider_name)
            fallback = fallback_full[~fallback_full["proveedor"].str.upper().isin(proveedores_con_detalle)]

    combinado = pd.concat([detalle, fallback], ignore_index=True)
    return combinado


# ---------------------------------------------------------------------------
# Exportación a Excel (misma estructura de hojas, columnas nuevas incluidas)
# ---------------------------------------------------------------------------

_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
_ALERT_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
_THIN = Side(style="thin", color="B7B7B7")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _style_header_row(ws, row_idx: int, n_cols: int) -> None:
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _BORDER


def _autofit(ws, n_cols: int, min_width: int = 12, max_width: int = 42) -> None:
    for c in range(1, n_cols + 1):
        col_letter = get_column_letter(c)
        max_len = min_width
        for cell in ws[col_letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, max_width)


def export_to_excel(df: pd.DataFrame) -> bytes:
    """
    Genera un .xlsx con la estructura de 4 hojas del sistema:
      - Resumen Ejecutivo (KPIs, incluye TC promedio)
      - DB_Pagos (base normalizada, con Tesoreria y Concepto/Presupuesto)
      - Matriz Semanal (pivote consolidado por proveedor, con acumulado)
      - Control Proveedores (compromiso vs. desembolsado)
    Devuelve los bytes del archivo, listos para `st.download_button`.
    """
    wb = Workbook()

    # ---------- Resumen Ejecutivo ----------
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    ws["A1"] = "TABLERO EJECUTIVO - FLUJO DE EFECTIVO (MULTISEDE)"
    ws["A1"].font = _TITLE_FONT
    ws["A2"] = "Monitoreo financiero integral de erogaciones programadas en billetes físicos."
    ws["A2"].font = Font(italic=True, color="666666")
    ws["A3"] = f"Generado: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"

    kpis = compute_kpis(df)
    kpi_rows = [
        ("Total comprometido (ARS)", kpis.total_comprometido),
        ("Total pagado (ARS)", kpis.total_pagado),
        ("Saldo pendiente (ARS)", kpis.saldo_pendiente),
        ("Desembolso promedio (ARS)", kpis.desembolso_promedio),
        ("Tipo de cambio promedio (USD)", kpis.tc_promedio),
        ("Cantidad de desembolsos", kpis.cantidad_desembolsos),
        ("Desembolsos pagados", kpis.cantidad_pagados),
    ]
    start = 5
    ws.cell(row=start, column=1, value="Indicador").font = Font(bold=True)
    ws.cell(row=start, column=2, value="Valor").font = Font(bold=True)
    for i, (label, val) in enumerate(kpi_rows, start=start + 1):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=val)
    _autofit(ws, 2)

    # distribución por sede
    row_sede = start + len(kpi_rows) + 2
    ws.cell(row=row_sede, column=1, value="DISTRIBUCIÓN POR SEDE (ARS)").font = Font(bold=True)
    ws.cell(row=row_sede + 1, column=1, value="Tesorería")
    ws.cell(row=row_sede + 1, column=2, value="Importe ARS")
    por_sede = df.groupby("tesoreria")["importe_ars"].sum().reset_index() if not df.empty else pd.DataFrame()
    for i, (_, r) in enumerate(por_sede.iterrows(), start=row_sede + 2):
        ws.cell(row=i, column=1, value=r["tesoreria"])
        ws.cell(row=i, column=2, value=r["importe_ars"])

    # distribución mensual
    dist = monthly_distribution(df)
    row0 = row_sede + len(por_sede) + 4
    ws.cell(row=row0, column=1, value="DISTRIBUCIÓN MENSUAL PROYECTADA (ARS)").font = Font(bold=True)
    ws.cell(row=row0 + 1, column=1, value="Mes")
    ws.cell(row=row0 + 1, column=2, value="Importe ARS")
    for i, (_, r) in enumerate(dist.iterrows(), start=row0 + 2):
        ws.cell(row=i, column=1, value=r["mes"])
        ws.cell(row=i, column=2, value=r["importe_ars"])

    # ---------- DB_Pagos ----------
    ws2 = wb.create_sheet("DB_Pagos")
    ws2["A1"] = "BASE CENTRAL DE PAGOS EN EFECTIVO (MULTISEDE)"
    ws2["A1"].font = _TITLE_FONT
    ws2["A2"] = "Registro único normalizado. Todos los tableros y matrices se calculan a partir de esta hoja."
    ws2["A2"].font = Font(italic=True, color="666666")

    headers = [
        "ID Pago", "Tesoreria", "Fecha Pago", "Lunes Semana", "Semana Etiqueta", "Mes",
        "Proveedor", "Proyecto / Destino", "Concepto / Presupuesto", "Moneda", "Importe Moneda",
        "Tipo Cambio", "Importe ARS", "Estado", "Observaciones",
    ]
    header_row = 4
    for c, h in enumerate(headers, start=1):
        ws2.cell(row=header_row, column=c, value=h)
    _style_header_row(ws2, header_row, len(headers))

    for i, (_, r) in enumerate(df.iterrows(), start=header_row + 1):
        ws2.cell(row=i, column=1, value=r["id"])
        ws2.cell(row=i, column=2, value=r["tesoreria"])
        ws2.cell(row=i, column=3, value=r["fecha"].to_pydatetime() if pd.notna(r["fecha"]) else None)
        ws2.cell(row=i, column=4, value=r["lunes_semana"].to_pydatetime() if pd.notna(r["lunes_semana"]) else None)
        ws2.cell(row=i, column=5, value=r["semana_etiqueta"])
        ws2.cell(row=i, column=6, value=r["mes"])
        ws2.cell(row=i, column=7, value=r["proveedor"])
        ws2.cell(row=i, column=8, value=r["proyecto"])
        ws2.cell(row=i, column=9, value=r["concepto"])
        ws2.cell(row=i, column=10, value=r["moneda"])
        ws2.cell(row=i, column=11, value=r["importe"])
        ws2.cell(row=i, column=12, value=r["tc"])
        ws2.cell(row=i, column=13, value=r["importe_ars"])
        ws2.cell(row=i, column=14, value=r["estado"])
        ws2.cell(row=i, column=15, value=r["obs"])
    total_row = header_row + len(df) + 1
    ws2.cell(row=total_row, column=1, value="TOTAL EFECTIVO").font = Font(bold=True)
    ws2.cell(row=total_row, column=13, value=round(df["importe_ars"].sum(), 2) if not df.empty else 0).font = Font(bold=True)
    _autofit(ws2, len(headers))
    for col in ("C", "D"):
        for cell in ws2[col][header_row:total_row]:
            cell.number_format = "DD/MM/YYYY"

    # ---------- Matriz Semanal ----------
    ws3 = wb.create_sheet("Matriz Semanal")
    ws3["A1"] = "MATRIZ DE FLUJO SEMANAL DE EFECTIVO (CONSOLIDADA POR PROVEEDOR)"
    ws3["A1"].font = _TITLE_FONT
    ws3["A2"] = "Generación automática vinculada a DB_Pagos. Incluye acumulado."
    ws3["A2"].font = Font(italic=True, color="666666")

    matrix = build_weekly_matrix(df)
    if not matrix.empty:
        header_row3 = 4
        for c, col_name in enumerate(matrix.columns, start=1):
            ws3.cell(row=header_row3, column=c, value=col_name)
        _style_header_row(ws3, header_row3, len(matrix.columns))

        for i, (_, r) in enumerate(matrix.iterrows(), start=header_row3 + 1):
            for c, col_name in enumerate(matrix.columns, start=1):
                ws3.cell(row=i, column=c, value=r[col_name])
            total_val = r["TOTAL SEMANAL"]
            if r["Semana"] != "TOTAL POR PROVEEDOR" and total_val > UMBRAL_PICO_EFECTIVO:
                for c in range(1, len(matrix.columns) + 1):
                    ws3.cell(row=i, column=c).fill = _ALERT_FILL
        _autofit(ws3, len(matrix.columns))

    # ---------- Control Proveedores ----------
    ws4 = wb.create_sheet("Control Proveedores")
    ws4["A1"] = "CONTROL INTEGRAL POR PROVEEDOR"
    ws4["A1"].font = _TITLE_FONT
    ws4["A2"] = "Resumen consolidado de montos comprometidos vs. desembolsados."
    ws4["A2"].font = Font(italic=True, color="666666")

    ctrl = provider_control(df)
    headers4 = [
        "Proveedor", "Cantidad Pagos", "Primer Pago", "Último Pago",
        "Moneda Base", "Compromiso Total ARS", "Pagado ARS", "Saldo ARS",
        "% del Flujo", "% Avance",
    ]
    header_row4 = 4
    for c, h in enumerate(headers4, start=1):
        ws4.cell(row=header_row4, column=c, value=h)
    _style_header_row(ws4, header_row4, len(headers4))

    for i, (_, r) in enumerate(ctrl.iterrows(), start=header_row4 + 1):
        ws4.cell(row=i, column=1, value=r["proveedor"])
        ws4.cell(row=i, column=2, value=int(r["cantidad_pagos"]))
        ws4.cell(row=i, column=3, value=r["primer_pago"].to_pydatetime() if pd.notna(r["primer_pago"]) else None)
        ws4.cell(row=i, column=4, value=r["ultimo_pago"].to_pydatetime() if pd.notna(r["ultimo_pago"]) else None)
        ws4.cell(row=i, column=5, value=r["moneda_base"])
        ws4.cell(row=i, column=6, value=r["compromiso_total_ars"])
        ws4.cell(row=i, column=7, value=r["pagado_ars"])
        ws4.cell(row=i, column=8, value=r["saldo_ars"])
        ws4.cell(row=i, column=9, value=r["pct_del_flujo"])
        ws4.cell(row=i, column=10, value=r["pct_avance"])
    _autofit(ws4, len(headers4))
    for col in ("C", "D"):
        for cell in ws4[col][header_row4: header_row4 + len(ctrl)]:
            cell.number_format = "DD/MM/YYYY"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
