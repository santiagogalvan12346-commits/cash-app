"""
core.py — Sistema de Gestión de Efectivo (multisede)
--------------------------------------------------------------------
Lógica de negocio, cálculos de tesorería y conciliación de flujo.
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
ESTADOS_VALIDOS = ["Pendiente", "Aprobado", "Pagado"]
MONEDAS_VALIDAS = ["ARS", "USD"]
PROYECTO_DEFAULT = "L2 TUC"
TESORERIA_DEFAULT = "Tucumán"
TESORERIAS_VALIDAS = ["Tucumán", "Buenos Aires"]
SEDE_CONSOLIDADO = "Consolidado"
OPCIONES_SEDE = [SEDE_CONSOLIDADO] + TESORERIAS_VALIDAS

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
    df = pd.read_csv(path)
    return normalize_df(df)


def load_from_uploaded_xlsx(file_like) -> pd.DataFrame:
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
            "estado": raw["Estado"].fillna("Pendiente"),
            "obs": raw["Observaciones"].fillna("") if "Observaciones" in raw.columns else "",
        }
    )
    return normalize_df(df)


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce", dayfirst=True, format="mixed")
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0.0)
    df["tc"] = pd.to_numeric(df["tc"], errors="coerce").fillna(1.0)
    df["moneda"] = df["moneda"].fillna("ARS").astype(str).str.upper().str.strip()
    df["proveedor"] = df["proveedor"].astype(str).str.strip()
    df["proyecto"] = df["proyecto"].fillna(PROYECTO_DEFAULT).astype(str).str.strip()
    df["estado"] = df["estado"].fillna("Pendiente").astype(str).str.strip()
    df["estado"] = df["estado"].replace({"Programado": "Pendiente"})
    df["obs"] = df["obs"].fillna("").astype(str)

    df["tesoreria"] = df["tesoreria"].fillna(TESORERIA_DEFAULT).astype(str).str.strip()
    df.loc[~df["tesoreria"].isin(TESORERIAS_VALIDAS), "tesoreria"] = TESORERIA_DEFAULT
    df["concepto"] = df["concepto"].fillna("").astype(str).str.strip()

    df = compute_calculated_fields(df)
    df = df.sort_values("fecha").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Campos calculados
# ---------------------------------------------------------------------------

def monday_of(fecha: pd.Timestamp) -> pd.Timestamp:
    if pd.isna(fecha):
        return pd.NaT
    return fecha - pd.Timedelta(days=fecha.weekday())


def week_label(lunes: pd.Timestamp) -> str:
    if pd.isna(lunes):
        return ""
    viernes = lunes + pd.Timedelta(days=4)
    return f"{lunes.strftime('%d/%m')} - {viernes.strftime('%d/%m')}"


def compute_importe_ars(row) -> float:
    if str(row["moneda"]).upper() == "USD":
        return round(float(row["importe"]) * float(row["tc"]), 2)
    return round(float(row["importe"]), 2)


def compute_calculated_fields(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["lunes_semana"] = df["fecha"].apply(monday_of)
    df["semana_etiqueta"] = df["lunes_semana"].apply(week_label)
    df["mes"] = df["fecha"].dt.to_period("M").astype(str)
    df["importe_ars"] = df.apply(compute_importe_ars, axis=1)
    return df


def next_id(df: pd.DataFrame) -> str:
    if df.empty:
        return "EFC-001"
    nums = df["id"].astype(str).str.extract(r"(\d+)$")[0].dropna().astype(int)
    nxt = (nums.max() + 1) if not nums.empty else 1
    return f"EFC-{nxt:03d}"


def filter_by_sede(df: pd.DataFrame, sede: str) -> pd.DataFrame:
    if sede == SEDE_CONSOLIDADO or not sede:
        return df
    return df[df["tesoreria"] == sede]


# ---------------------------------------------------------------------------
# Matriz Semanal Consolidada
# ---------------------------------------------------------------------------

def build_weekly_matrix(df: pd.DataFrame) -> pd.DataFrame:
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
# KPIs y Resumen
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
    return out.sort_values("compromiso_total_ars", ascending=False).reset_index(drop=True)


def monthly_distribution(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["mes", "importe_ars"])
    out = df.groupby("mes")["importe_ars"].sum().reset_index()
    return out.sort_values("mes").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Asistente de Escaleras y Conciliación Rápida
# ---------------------------------------------------------------------------

def update_payments_status(df: pd.DataFrame, payment_ids: list[str], new_status: str = "Pagado") -> pd.DataFrame:
    """Actualiza en bloque el estado de los IDs seleccionados."""
    df = df.copy()
    mask = df["id"].isin(payment_ids)
    df.loc[mask, "estado"] = new_status
    return df


def generate_custom_ladder(
    df_existing: pd.DataFrame,
    tesoreria: str,
    proveedor: str,
    proyecto: str,
    concepto: str,
    moneda: str,
    tc: float,
    items: list[dict],
) -> pd.DataFrame:
    """
    Genera filas para una escalera con fechas e importes no lineales / personalizados.
    """
    rows = []
    base_id_num = 1
    if not df_existing.empty:
        nums = df_existing["id"].astype(str).str.extract(r"(\d+)$")[0].dropna().astype(int)
        if not nums.empty:
            base_id_num = nums.max() + 1

    total_items = len(items)
    for i, item in enumerate(items):
        prefix = "EFC" if tesoreria == "Tucumán" else "BA"
        new_id = f"{prefix}-{base_id_num + i:03d}"
        f_val = pd.Timestamp(item["fecha"])
        imp_val = float(item["importe"])

        rows.append({
            "id": new_id,
            "tesoreria": tesoreria,
            "fecha": f_val,
            "proveedor": proveedor.strip().upper(),
            "proyecto": proyecto.strip().upper(),
            "concepto": concepto.strip(),
            "moneda": moneda,
            "importe": imp_val,
            "tc": float(tc) if moneda == "USD" else 1.0,
            "estado": "Pendiente",
            "obs": f"Cuota {i + 1}/{total_items}",
        })

    new_df = pd.DataFrame(rows, columns=COLUMNS)
    return normalize_df(new_df)


# ---------------------------------------------------------------------------
# Exportación Excel
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
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    ws["A1"] = "TABLERO EJECUTIVO - FLUJO DE EFECTIVO"
    ws["A1"].font = _TITLE_FONT
    ws["A2"] = f"Generado: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"

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
    start = 4
    for i, (label, val) in enumerate(kpi_rows, start=start):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=val)
    _autofit(ws, 2)

    # Matriz Semanal
    ws3 = wb.create_sheet("Matriz Semanal")
    matrix = build_weekly_matrix(df)
    if not matrix.empty:
        for c, col_name in enumerate(matrix.columns, start=1):
            ws3.cell(row=1, column=c, value=col_name)
        _style_header_row(ws3, 1, len(matrix.columns))
        for i, (_, r) in enumerate(matrix.iterrows(), start=2):
            for c, col_name in enumerate(matrix.columns, start=1):
                ws3.cell(row=i, column=c, value=r[col_name])
        _autofit(ws3, len(matrix.columns))

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
import re
import requests

def clean_cuit(val: str) -> str:
    """Elimina guiones y caracteres no numéricos dejando 11 dígitos."""
    return re.sub(r"\D", "", str(val or ""))[:11]

def parse_cuits_input(text: str) -> list[str]:
    """Extrae CUITs únicos de un bloque de texto libre."""
    raw = re.split(r"[\n,; \t]+", str(text or ""))
    cleaned = [clean_cuit(x) for x in raw if clean_cuit(x)]
    return sorted(list({c for c in cleaned if len(c) == 11}))

def fetch_bcra_data(cuit: str) -> dict:
    """Consulta la API del BCRA para Central de Deudores y Cheques Rechazados."""
    base_headers = {"Accept": "application/json"}
    
    # 1. Central de Deudores
    deuda_data = None
    deuda_status = "ok"
    try:
        r_deuda = requests.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/{cuit}",
            headers=base_headers,
            timeout=8,
        )
        if r_deuda.status_code == 200:
            deuda_data = r_deuda.json().get("results", {})
        elif r_deuda.status_code == 404:
            deuda_status = "none"
        else:
            deuda_status = "error"
    except Exception:
        deuda_status = "error"

    # 2. Cheques Rechazados
    checks = []
    try:
        r_checks = requests.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/ChequesRechazados/{cuit}",
            headers=base_headers,
            timeout=8,
        )
        if r_checks.status_code == 200:
            cr = r_checks.json().get("results", {})
            causales = cr.get("causales", [])
            for c in causales:
                causal_nombre = c.get("causal", "Sin causal")
                for e in c.get("entidades", []):
                    entidad_nombre = e.get("entidad", "Entidad no informada")
                    for d in e.get("detalle", []):
                        checks.append({
                            "fechaRechazo": d.get("fechaRechazo", ""),
                            "fechaPago": d.get("fechaPago", ""),
                            "monto": float(d.get("monto", 0.0) or 0.0),
                            "causal": causal_nombre,
                            "entidad": entidad_nombre,
                        })
    except Exception:
        checks = []

    # Procesar perfil crediticio
    if deuda_status == "ok" and deuda_data:
        periods = deuda_data.get("periodos", [])
        latest = periods[0] if periods else {}
        ents = latest.get("entidades", [])

        worst = max([int(e.get("situacion", 0) or 0) for e in ents], default=0)
        debt = sum([(float(e.get("monto", 0) or 0) * 1000.0) for e in ents])
        rejected = len(checks)
        pending = len([c for c in checks if not c.get("fechaPago")])
        pending_amount = sum([c["monto"] for c in checks if not c.get("fechaPago")])

        # Criterio de semáforo
        if worst >= 3 or pending > 0 or rejected > 5:
            risk = "bad"
            rt = "ALERTA"
        elif worst == 2 or (1 <= rejected <= 5):
            risk = "warn"
            rt = "REVISAR"
        else:
            risk = "ok"
            rt = "SIN ALERTAS"

        alerts = []
        if worst >= 3:
            alerts.append(f"Peor situación informada: {worst} (Deterioro / Riesgo alto).")
        elif worst == 2:
            alerts.append("Registra entidad(es) en Situación 2 (Seguimiento especial).")

        if pending > 0:
            alerts.append(f"Posee {pending} cheque(s) rechazado(s) PENDIENTE(S) de pago.")
        if rejected > 5:
            alerts.append(f"Historial crítico de cheques: registra {rejected} rechazos en total.")
        elif rejected > 0 and pending == 0:
            alerts.append(f"Registra {rejected} rechazo(s) histórico(s), pero figuran cancelados/pagados.")

        if not alerts:
            alerts.append("Sin señales negativas: Situación normal y sin cheques rechazados.")

        # Ordenar últimos 3 cheques
        sorted_checks = sorted(checks, key=lambda x: str(x.get("fechaRechazo", "")), reverse=True)
        last3 = sorted_checks[:3]

        banks = [
            [e.get("entidad", ""), int(e.get("situacion", 0) or 0), (float(e.get("monto", 0) or 0) * 1000.0)]
            for e in ents
        ]

        return {
            "cuit": cuit,
            "denominacion": deuda_data.get("denominacion") or f"Librador {cuit}",
            "risk": risk,
            "risk_label": rt,
            "worst": worst,
            "debt": debt,
            "rejected": rejected,
            "pending": pending,
            "pending_amount": pending_amount,
            "alerts": alerts,
            "banks": banks,
            "last3": last3,
        }

    elif deuda_status == "none":
        rej_count = len(checks)
        pen_count = len([c for c in checks if not c.get("fechaPago")])
        r = "ok"
        rt = "SIN DEUDA BCRA"
        if pen_count > 0 or rej_count > 5:
            r, rt = "bad", "ALERTA"
        elif rej_count > 0:
            r, rt = "warn", "REVISAR"

        return {
            "cuit": cuit,
            "denominacion": f"CUIT {cuit}",
            "risk": r,
            "risk_label": rt,
            "worst": 0,
            "debt": 0.0,
            "rejected": rej_count,
            "pending": pen_count,
            "pending_amount": 0.0,
            "alerts": ["Sin deuda bancaria registrada en Central de Deudores."],
            "banks": [],
            "last3": checks[:3],
        }
    else:
        return {
            "cuit": cuit,
            "denominacion": f"CUIT {cuit}",
            "risk": "warn",
            "risk_label": "ERROR CONSULTA",
            "worst": "-",
            "debt": 0.0,
            "rejected": 0,
            "pending": 0,
            "pending_amount": 0.0,
            "alerts": ["Error de comunicación con los servidores del BCRA."],
            "banks": [],
            "last3": [],
        }
# ---------------------------------------------------------------------------
# Módulo de Consulta y Scoring BCRA
# ---------------------------------------------------------------------------
import re
import requests


def clean_cuit(val: str) -> str:
    """Elimina guiones y caracteres no numéricos dejando 11 dígitos."""
    return re.sub(r"\D", "", str(val or ""))[:11]


def parse_cuits_input(text: str) -> list[str]:
    """Extrae CUITs únicos de un bloque de texto libre."""
    raw = re.split(r"[\n,; \t]+", str(text or ""))
    cleaned = [clean_cuit(x) for x in raw if clean_cuit(x)]
    return sorted(list({c for c in cleaned if len(c) == 11}))


def fetch_bcra_data(cuit: str) -> dict:
    """Consulta la API del BCRA para Central de Deudores y Cheques Rechazados."""
    base_headers = {"Accept": "application/json"}

    # 1. Central de Deudores
    deuda_data = None
    deuda_status = "ok"
    try:
        r_deuda = requests.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/{cuit}",
            headers=base_headers,
            timeout=8,
        )
        if r_deuda.status_code == 200:
            deuda_data = r_deuda.json().get("results", {})
        elif r_deuda.status_code == 404:
            deuda_status = "none"
        else:
            deuda_status = "error"
    except Exception:
        deuda_status = "error"

    # 2. Cheques Rechazados
    checks = []
    try:
        r_checks = requests.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/ChequesRechazados/{cuit}",
            headers=base_headers,
            timeout=8,
        )
        if r_checks.status_code == 200:
            cr = r_checks.json().get("results", {})
            causales = cr.get("causales", [])
            for c in causales:
                causal_nombre = c.get("causal", "Sin causal")
                for e in c.get("entidades", []):
                    entidad_nombre = e.get("entidad", "Entidad no informada")
                    for d in e.get("detalle", []):
                        checks.append({
                            "fechaRechazo": d.get("fechaRechazo", ""),
                            "fechaPago": d.get("fechaPago", ""),
                            "monto": float(d.get("monto", 0.0) or 0.0),
                            "causal": causal_nombre,
                            "entidad": entidad_nombre,
                        })
    except Exception:
        checks = []

    # 3. Procesar perfil crediticio
    if deuda_status == "ok" and deuda_data:
        periods = deuda_data.get("periodos", [])
        latest = periods[0] if periods else {}
        ents = latest.get("entidades", [])

        worst = max([int(e.get("situacion", 0) or 0) for e in ents], default=0)
        debt = sum([(float(e.get("monto", 0) or 0) * 1000.0) for e in ents])
        rejected = len(checks)
        pending = len([c for c in checks if not c.get("fechaPago")])
        pending_amount = sum([c["monto"] for c in checks if not c.get("fechaPago")])

        # Criterio de semáforo
        if worst >= 3 or pending > 0 or rejected > 5:
            risk = "bad"
            rt = "ALERTA"
        elif worst == 2 or (1 <= rejected <= 5):
            risk = "warn"
            rt = "REVISAR"
        else:
            risk = "ok"
            rt = "SIN ALERTAS"

        alerts = []
        if worst >= 3:
            alerts.append(f"Peor situación informada: {worst} (Deterioro / Riesgo alto).")
        elif worst == 2:
            alerts.append("Registra entidad(es) en Situación 2 (Seguimiento especial).")

        if pending > 0:
            alerts.append(f"Posee {pending} cheque(s) rechazado(s) PENDIENTE(S) de pago.")
        if rejected > 5:
            alerts.append(f"Historial crítico de cheques: registra {rejected} rechazos en total.")
        elif rejected > 0 and pending == 0:
            alerts.append(f"Registra {rejected} rechazo(s) histórico(s), pero figuran cancelados/pagados.")

        if not alerts:
            alerts.append("Sin señales negativas: Situación normal y sin cheques rechazados.")

        # Ordenar últimos 3 cheques por fecha
        sorted_checks = sorted(checks, key=lambda x: str(x.get("fechaRechazo", "")), reverse=True)
        last3 = sorted_checks[:3]

        banks = [
            [e.get("entidad", ""), int(e.get("situacion", 0) or 0), (float(e.get("monto", 0) or 0) * 1000.0)]
            for e in ents
        ]

        return {
            "cuit": cuit,
            "denominacion": deuda_data.get("denominacion") or f"Librador {cuit}",
            "risk": risk,
            "risk_label": rt,
            "worst": worst,
            "debt": debt,
            "rejected": rejected,
            "pending": pending,
            "pending_amount": pending_amount,
            "alerts": alerts,
            "banks": banks,
            "last3": last3,
        }

    elif deuda_status == "none":
        rej_count = len(checks)
        pen_count = len([c for c in checks if not c.get("fechaPago")])
        r = "ok"
        rt = "SIN DEUDA BCRA"
        if pen_count > 0 or rej_count > 5:
            r, rt = "bad", "ALERTA"
        elif rej_count > 0:
            r, rt = "warn", "REVISAR"

        return {
            "cuit": cuit,
            "denominacion": f"CUIT {cuit}",
            "risk": r,
            "risk_label": rt,
            "worst": 0,
            "debt": 0.0,
            "rejected": rej_count,
            "pending": pen_count,
            "pending_amount": 0.0,
            "alerts": ["Sin deuda bancaria registrada en Central de Deudores."],
            "banks": [],
            "last3": checks[:3],
        }
    else:
        return {
            "cuit": cuit,
            "denominacion": f"CUIT {cuit}",
            "risk": "warn",
            "risk_label": "ERROR CONSULTA",
            "worst": "-",
            "debt": 0.0,
            "rejected": 0,
            "pending": 0,
            "pending_amount": 0.0,
            "alerts": ["Error de comunicación con los servidores del BCRA."],
            "banks": [],
            "last3": [],
        }
