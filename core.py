"""
core.py
--------------------------------------------------------------------
Lógica de negocio del Sistema de Gestión de Efectivo L2 Tucumán.

Este módulo es independiente de la capa de UI (Streamlit) para que
pueda ser testeado o reutilizado (por ejemplo, en un script de
línea de comandos o en otra interfaz web).

Responsabilidades:
  - Normalización de tipos de datos (fechas, montos, moneda).
  - Cálculo de "Lunes de Semana" y "Semana Etiqueta" (lunes a viernes).
  - Cálculo del Importe ARS (conversión automática si Moneda = USD).
  - Construcción de la Matriz de Tesorería Semanal (pivote).
  - Cálculo de KPIs ejecutivos y detección de picos de efectivo.
  - Exportación de la base normalizada + matriz + control de
    proveedores a un archivo .xlsx con el mismo esquema de hojas
    del archivo original.
--------------------------------------------------------------------
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constantes de negocio
# ---------------------------------------------------------------------------

UMBRAL_PICO_EFECTIVO = 10_000_000  # ARS — disparo de alerta logística
ESTADOS_VALIDOS = ["Programado", "Aprobado", "Pagado"]
MONEDAS_VALIDAS = ["ARS", "USD"]
PROYECTO_DEFAULT = "L2 TUC"

COLUMNS = [
    "id",
    "fecha",
    "proveedor",
    "proyecto",
    "moneda",
    "importe",
    "tc",
    "estado",
    "obs",
]

DTYPES = {
    "id": "string",
    "proveedor": "string",
    "proyecto": "string",
    "moneda": "string",
    "estado": "string",
    "obs": "string",
}


# ---------------------------------------------------------------------------
# Carga y normalización de datos
# ---------------------------------------------------------------------------

def load_seed_csv(path: str | Path) -> pd.DataFrame:
    """Carga el CSV semilla (extraído de DB_Pagos) y normaliza tipos."""
    df = pd.read_csv(path)
    return normalize_df(df)


def load_from_uploaded_xlsx(file_like) -> pd.DataFrame:
    """
    Lee un archivo .xlsx subido por el usuario, reconociendo la hoja
    DB_Pagos con el mismo esquema del archivo original
    (SISTEMA_GESTION_EFECTIVO_L2_TUC.xlsx).
    """
    raw = pd.read_excel(file_like, sheet_name="DB_Pagos", header=3)
    raw = raw.dropna(subset=["ID Pago"])
    raw = raw[raw["ID Pago"].astype(str).str.upper() != "TOTAL EFECTIVO"]

    df = pd.DataFrame(
        {
            "id": raw["ID Pago"].astype(str),
            "fecha": pd.to_datetime(raw["Fecha Pago"]),
            "proveedor": raw["Proveedor"].astype(str).str.strip(),
            "proyecto": raw["Proyecto / Destino"].fillna(PROYECTO_DEFAULT),
            "moneda": raw["Moneda"].fillna("ARS").astype(str).str.upper(),
            "importe": pd.to_numeric(raw["Importe Moneda"], errors="coerce"),
            "tc": pd.to_numeric(raw["Tipo Cambio"], errors="coerce").fillna(1),
            "estado": raw["Estado"].fillna("Programado"),
            "obs": raw["Observaciones"].fillna(""),
        }
    )
    return normalize_df(df)


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Asegura tipos correctos y agrega columnas calculadas."""
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
    nums = (
        df["id"]
        .astype(str)
        .str.extract(r"(\d+)$")[0]
        .dropna()
        .astype(int)
    )
    nxt = (nums.max() + 1) if not nums.empty else 1
    return f"EFC-{nxt:03d}"


# ---------------------------------------------------------------------------
# Matriz de Tesorería Semanal (pivote)
# ---------------------------------------------------------------------------

def build_weekly_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """
    Construye la matriz semana x proveedor con importes ARS,
    totalizadores por fila (semana) y por columna (proveedor).
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

    pivot = pivot.reset_index()
    pivot = pivot.drop(columns=["lunes_semana"])
    pivot = pivot.rename(columns={"semana_etiqueta": "Semana"})

    total_row = pivot.drop(columns=["Semana"]).sum(numeric_only=True)
    total_row["Semana"] = "TOTAL POR PROVEEDOR"
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


def compute_kpis(df: pd.DataFrame) -> Kpis:
    if df.empty:
        return Kpis(0, 0, 0, 0, 0, 0)

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
    out["pct_avance"] = (
        (out["pagado_ars"] / out["compromiso_total_ars"] * 100)
        .fillna(0.0)
        .round(2)
    )

    out = out.sort_values("compromiso_total_ars", ascending=False).reset_index(drop=True)
    return out


def monthly_distribution(df: pd.DataFrame) -> pd.DataFrame:
    """Distribución mensual proyectada del total en ARS."""
    if df.empty:
        return pd.DataFrame(columns=["mes", "importe_ars"])
    out = df.groupby("mes")["importe_ars"].sum().reset_index()
    return out.sort_values("mes").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Exportación a Excel (misma estructura de hojas que el archivo original)
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
    Genera un .xlsx con la misma estructura del archivo original:
      - DB_Pagos (base normalizada)
      - Matriz Semanal (pivote con totalizadores)
      - Control Proveedores (compromiso vs. desembolsado)
      - Resumen Ejecutivo (KPIs)
    Devuelve los bytes del archivo, listos para `st.download_button`.
    """
    wb = Workbook()

    # ---------- Resumen Ejecutivo ----------
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    ws["A1"] = "TABLERO EJECUTIVO - FLUJO DE EFECTIVO L2 TUCUMÁN"
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

    # distribución mensual
    dist = monthly_distribution(df)
    row0 = start + len(kpi_rows) + 3
    ws.cell(row=row0, column=1, value="DISTRIBUCIÓN MENSUAL PROYECTADA (ARS)").font = Font(bold=True)
    ws.cell(row=row0 + 1, column=1, value="Mes")
    ws.cell(row=row0 + 1, column=2, value="Importe ARS")
    for i, (_, r) in enumerate(dist.iterrows(), start=row0 + 2):
        ws.cell(row=i, column=1, value=r["mes"])
        ws.cell(row=i, column=2, value=r["importe_ars"])

    # ---------- DB_Pagos ----------
    ws2 = wb.create_sheet("DB_Pagos")
    ws2["A1"] = "BASE CENTRAL DE PAGOS EN EFECTIVO"
    ws2["A1"].font = _TITLE_FONT
    ws2["A2"] = "Registro único normalizado. Todos los tableros y matrices se calculan a partir de esta hoja."
    ws2["A2"].font = Font(italic=True, color="666666")

    headers = [
        "ID Pago", "Fecha Pago", "Lunes Semana", "Semana Etiqueta", "Mes",
        "Proveedor", "Proyecto / Destino", "Moneda", "Importe Moneda",
        "Tipo Cambio", "Importe ARS", "Estado", "Observaciones",
    ]
    header_row = 4
    for c, h in enumerate(headers, start=1):
        ws2.cell(row=header_row, column=c, value=h)
    _style_header_row(ws2, header_row, len(headers))

    for i, (_, r) in enumerate(df.iterrows(), start=header_row + 1):
        ws2.cell(row=i, column=1, value=r["id"])
        ws2.cell(row=i, column=2, value=r["fecha"].to_pydatetime() if pd.notna(r["fecha"]) else None)
        ws2.cell(row=i, column=3, value=r["lunes_semana"].to_pydatetime() if pd.notna(r["lunes_semana"]) else None)
        ws2.cell(row=i, column=4, value=r["semana_etiqueta"])
        ws2.cell(row=i, column=5, value=r["mes"])
        ws2.cell(row=i, column=6, value=r["proveedor"])
        ws2.cell(row=i, column=7, value=r["proyecto"])
        ws2.cell(row=i, column=8, value=r["moneda"])
        ws2.cell(row=i, column=9, value=r["importe"])
        ws2.cell(row=i, column=10, value=r["tc"])
        ws2.cell(row=i, column=11, value=r["importe_ars"])
        ws2.cell(row=i, column=12, value=r["estado"])
        ws2.cell(row=i, column=13, value=r["obs"])
    total_row = header_row + len(df) + 1
    ws2.cell(row=total_row, column=1, value="TOTAL EFECTIVO").font = Font(bold=True)
    ws2.cell(row=total_row, column=11, value=round(df["importe_ars"].sum(), 2)).font = Font(bold=True)
    _autofit(ws2, len(headers))
    for col in ("B", "C"):
        for cell in ws2[col][header_row:total_row]:
            cell.number_format = "DD/MM/YYYY"

    # ---------- Matriz Semanal ----------
    ws3 = wb.create_sheet("Matriz Semanal")
    ws3["A1"] = "MATRIZ DE FLUJO SEMANAL DE EFECTIVO"
    ws3["A1"].font = _TITLE_FONT
    ws3["A2"] = "Generación automática vinculada a DB_Pagos."
    ws3["A2"].font = Font(italic=True, color="666666")

    matrix = build_weekly_matrix(df)
    if not matrix.empty:
        header_row3 = 4
        for c, col_name in enumerate(matrix.columns, start=1):
            ws3.cell(row=header_row3, column=c, value=col_name)
        _style_header_row(ws3, header_row3, len(matrix.columns))

        alert_col = list(matrix.columns).index("TOTAL SEMANAL") + 1
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
        for cell in ws4[col][header_row4 : header_row4 + len(ctrl)]:
            cell.number_format = "DD/MM/YYYY"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
