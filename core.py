# -*- coding: utf-8 -*-
"""
Funciones centrales y modelos de datos — Suite Financiera (For Drink SA)
"""
from __future__ import annotations

import datetime as dt
import re
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Constantes del módulo de pagos / tesorería
SEDE_CONSOLIDADO = "Consolidado"
TESORERIAS_VALIDAS = ["Tucumán", "Buenos Aires"]
OPCIONES_SEDE = [SEDE_CONSOLIDADO, "Tucumán", "Buenos Aires"]
MONEDAS_VALIDAS = ["ARS", "USD"]
ESTADOS_VALIDOS = ["Pendiente", "Pagado", "Observado"]

COLUMNS = [
    "id", "tesoreria", "fecha", "proveedor", "proyecto",
    "concepto", "moneda", "importe", "tc", "estado", "obs"
]

COLUMNS_CHEQUES = [
    "Banco", "EMPRESA", "Cuenta Libradora", "Fecha Emisión", "Fecha Pago",
    "DIA", "DIA.1", "MES", "MES.2", "AÑO", "Nro. de Cheque", "Importe",
    "CUIT Beneficiario", "Razón Social Beneficiario", "Estado", "Motivo Anulación"
]

DIAS_ES = {0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves", 4: "viernes", 5: "sábado", 6: "domingo"}
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}
MESES_ES_UPPER = {k: v.upper() for k, v in MESES_ES.items()}
MESES_ABR = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
}


@dataclass
class KPIs:
    total_comprometido: float
    total_pagado: float
    saldo_pendiente: float
    desembolso_promedio: float


def normalize_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normaliza la base de pagos de tesorería preservando fechas y semanas en español."""
    if df_raw is None or df_raw.empty:
        df = pd.DataFrame(columns=COLUMNS)
    else:
        df = df_raw.copy()

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["id"] = df["id"].astype(str).str.strip()
    df["tesoreria"] = df["tesoreria"].astype(str).str.strip()
    
    # Parseo estricto dd/mm/aaaa
    df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, format="mixed", errors="coerce")
    
    df["proveedor"] = df["proveedor"].astype(str).str.strip()
    df["proyecto"] = df["proyecto"].astype(str).str.strip()
    df["concepto"] = df["concepto"].astype(str).str.strip()
    df["moneda"] = df["moneda"].astype(str).str.strip().str.upper()
    df["moneda"] = df["moneda"].apply(lambda m: m if m in MONEDAS_VALIDAS else "ARS")

    def _to_num(val, default=0.0):
        if pd.isna(val):
            return default
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace("$", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return default

    df["importe"] = df["importe"].apply(_to_num)
    df["tc"] = df["tc"].apply(lambda v: _to_num(v, default=1.0))
    df.loc[df["tc"] <= 0, "tc"] = 1.0

    df["estado"] = df["estado"].astype(str).str.strip()
    df["estado"] = df["estado"].apply(lambda e: e if e in ESTADOS_VALIDOS else "Pendiente")
    df["obs"] = df["obs"].fillna("").astype(str)

    df["importe_ars"] = df.apply(
        lambda r: r["importe"] * r["tc"] if r["moneda"] == "USD" else r["importe"],
        axis=1
    )

    df["mes"] = df["fecha"].dt.to_period("M").astype(str)
    df["lunes_semana"] = df["fecha"].apply(
        lambda d: (d - pd.Timedelta(days=d.weekday())).floor("D") if pd.notna(d) else pd.NaT
    )
    
    # Formato exacto semanal laboral: sem 21/09 - 25/09
    df["semana_etiqueta"] = df["lunes_semana"].apply(
        lambda d: f"sem {d.strftime('%d/%m')} - {(d + pd.Timedelta(days=4)).strftime('%d/%m')}" if pd.notna(d) else ""
    )

    return df


def load_seed_csv(path: Path) -> pd.DataFrame:
    try:
        df_raw = pd.read_csv(path)
        return normalize_df(df_raw)
    except Exception:
        return normalize_df(pd.DataFrame(columns=COLUMNS))


def filter_by_sede(df: pd.DataFrame, sede: str) -> pd.DataFrame:
    if sede == SEDE_CONSOLIDADO or not sede:
        return df.copy()
    return df[df["tesoreria"].str.lower() == sede.lower()].copy()


def compute_kpis(df: pd.DataFrame) -> KPIs:
    if df.empty:
        return KPIs(0.0, 0.0, 0.0, 0.0)
    total_comp = df["importe_ars"].sum()
    total_pag = df[df["estado"] == "Pagado"]["importe_ars"].sum()
    saldo_pend = df[df["estado"] != "Pagado"]["importe_ars"].sum()
    cant = len(df)
    prom = (total_comp / cant) if cant > 0 else 0.0
    return KPIs(total_comp, total_pag, saldo_pend, prom)


def monthly_distribution(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["mes", "importe_ars"])
    res = df.groupby("mes")["importe_ars"].sum().reset_index().sort_values("mes")
    return res


def provider_control(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["proveedor", "compromiso_total_ars", "pagado_ars", "saldo_pendiente_ars"])
    grouped = df.groupby("proveedor").agg(
        compromiso_total_ars=("importe_ars", "sum"),
        pagado_ars=("importe_ars", lambda s: s[df.loc[s.index, "estado"] == "Pagado"].sum()),
        saldo_pendiente_ars=("importe_ars", lambda s: s[df.loc[s.index, "estado"] != "Pagado"].sum()),
    ).reset_index()
    return grouped.sort_values("compromiso_total_ars", ascending=False)


def build_weekly_matrix(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "lunes_semana" not in df.columns:
        return pd.DataFrame()
    valid = df.dropna(subset=["lunes_semana"]).copy()
    if valid.empty:
        return pd.DataFrame()

    piv = valid.pivot_table(
        index=["lunes_semana", "semana_etiqueta"],
        columns="proveedor",
        values="importe_ars",
        aggfunc="sum",
        fill_value=0.0
    ).reset_index()

    piv = piv.sort_values("lunes_semana").reset_index(drop=True)
    prov_cols = [c for c in piv.columns if c not in ("lunes_semana", "semana_etiqueta")]
    piv["TOTAL SEMANAL"] = piv[prov_cols].sum(axis=1)
    piv["ACUMULADO"] = piv["TOTAL SEMANAL"].cumsum()

    tot_row = {"semana_etiqueta": "TOTAL POR PROVEEDOR"}
    for p in prov_cols:
        tot_row[p] = piv[p].sum()
    tot_row["TOTAL SEMANAL"] = piv["TOTAL SEMANAL"].sum()
    tot_row["ACUMULADO"] = piv["TOTAL SEMANAL"].sum()

    res = pd.concat([piv, pd.DataFrame([tot_row])], ignore_index=True)
    res = res.rename(columns={"semana_etiqueta": "Semana"})
    if "lunes_semana" in res.columns:
        res = res.drop(columns=["lunes_semana"])
    return res


def next_id(df: pd.DataFrame) -> str:
    if df.empty or "id" not in df.columns:
        return "PAG-0001"
    ids = df["id"].astype(str)
    nums = []
    for x in ids:
        m = re.search(r"(\d+)", x)
        if m:
            nums.append(int(m.group(1)))
    max_num = max(nums) if nums else 0
    return f"PAG-{max_num + 1:04d}"


def generate_custom_ladder(
    df_existing: pd.DataFrame,
    tesoreria: str,
    proveedor: str,
    proyecto: str,
    concepto: str,
    moneda: str,
    tc: float,
    items: list[dict]
) -> pd.DataFrame:
    rows = []
    start_id = 1
    if not df_existing.empty and "id" in df_existing.columns:
        ids = df_existing["id"].astype(str)
        nums = []
        for x in ids:
            m = re.search(r"(\d+)", x)
            if m:
                nums.append(int(m.group(1)))
        start_id = (max(nums) if nums else 0) + 1

    for idx, item in enumerate(items):
        new_id = f"PAG-{start_id + idx:04d}"
        rows.append({
            "id": new_id,
            "tesoreria": tesoreria,
            "fecha": pd.to_datetime(item["fecha"], dayfirst=True),
            "proveedor": proveedor,
            "proyecto": proyecto,
            "concepto": concepto,
            "moneda": moneda,
            "importe": float(item["importe"]),
            "tc": float(tc),
            "estado": "Pendiente",
            "obs": f"Tramo {idx + 1}/{len(items)}",
        })
    return normalize_df(pd.DataFrame(rows))


def update_payments_status(df: pd.DataFrame, payment_ids: list[str], new_status: str = "Pagado") -> pd.DataFrame:
    out = df.copy()
    mask = out["id"].isin(payment_ids)
    out.loc[mask, "estado"] = new_status
    return normalize_df(out)


def load_from_uploaded_xlsx(uploaded_file: Any) -> pd.DataFrame:
    df_raw = pd.read_excel(uploaded_file)
    return normalize_df(df_raw)


def export_to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[COLUMNS].to_excel(writer, index=False, sheet_name="Pagos")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Módulo de Clearing y Cheques Emitidos
# ---------------------------------------------------------------------------

def normalize_cheques_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Estandariza los cheques preservando estado y trazabilidad lógica."""
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=[
            "Banco", "EMPRESA", "Cuenta Libradora", "Fecha Emisión", "Fecha Pago",
            "DIA_TXT", "DIA_NUM", "MES_NUM", "MES_TXT", "AÑO", "Nro. de Cheque", "Importe",
            "CUIT Beneficiario", "Razón Social Beneficiario", "Estado", "Motivo Anulación",
            "FECHA_LABEL", "MES_KEY"
        ])

    df = df_raw.copy()
    cols = list(df.columns)
    col_mapping = {}
    dia_count = 0
    for c in cols:
        c_str = str(c).strip()
        if c_str.upper() == "DIA":
            dia_count += 1
            col_mapping[c] = "DIA_TXT" if dia_count == 1 else "DIA_NUM"
        elif c_str.upper() == "MES":
            col_mapping[c] = "MES_NUM"
        elif c_str.upper() == "MES.2":
            col_mapping[c] = "MES_TXT"
    df = df.rename(columns=col_mapping)

    def clean_currency_val(val):
        if pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace("$", "").replace(" ", "")
        if not s:
            return 0.0
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    if "Importe" in df.columns:
        df["Importe"] = df["Importe"].apply(clean_currency_val)

    for f_col in ["Fecha Emisión", "Fecha Pago"]:
        if f_col in df.columns:
            df[f_col] = pd.to_datetime(df[f_col], dayfirst=True, format="mixed", errors="coerce")

    if "DIA_NUM" in df.columns and "MES_NUM" in df.columns and "AÑO" in df.columns:
        mask_na = df["Fecha Pago"].isna()
        if mask_na.any():
            try:
                df.loc[mask_na, "Fecha Pago"] = pd.to_datetime({
                    "year": pd.to_numeric(df.loc[mask_na, "AÑO"], errors="coerce"),
                    "month": pd.to_numeric(df.loc[mask_na, "MES_NUM"], errors="coerce"),
                    "day": pd.to_numeric(df.loc[mask_na, "DIA_NUM"], errors="coerce"),
                })
            except Exception:
                pass

    if "Fecha Pago" in df.columns:
        df["DIA_TXT"] = df["Fecha Pago"].dt.dayofweek.map(DIAS_ES)
        df["DIA_NUM"] = df["Fecha Pago"].dt.day
        df["MES_NUM"] = df["Fecha Pago"].dt.month
        df["MES_TXT"] = df["Fecha Pago"].dt.month.map(MESES_ES)
        df["AÑO"] = df["Fecha Pago"].dt.year
        df["FECHA_LABEL"] = df["Fecha Pago"].apply(
            lambda d: f"{d.day:02d}-{MESES_ABR.get(d.month, '')}" if pd.notna(d) else ""
        )
        df["MES_KEY"] = df["Fecha Pago"].dt.strftime("%Y-%m")

    if "Banco" in df.columns:
        df["Banco"] = df["Banco"].astype(str).str.strip().str.upper()

    if "Nro. de Cheque" in df.columns:
        df["Nro. de Cheque"] = df["Nro. de Cheque"].astype(str).str.strip().replace("nan", "")

    # Estado de cheque (Baja Lógica): "Emitido" o "Anulado"
    if "Estado" not in df.columns:
        df["Estado"] = "Emitido"
    else:
        df["Estado"] = df["Estado"].fillna("Emitido").astype(str).str.strip()
        df.loc[df["Estado"] == "", "Estado"] = "Emitido"

    if "Motivo Anulación" not in df.columns:
        df["Motivo Anulación"] = ""
    else:
        df["Motivo Anulación"] = df["Motivo Anulación"].fillna("").astype(str).str.strip()

    return df


def prepare_cheques_to_save(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara el DataFrame para persistir en Google Sheets con trazabilidad."""
    df_out = pd.DataFrame()
    df_out["Banco"] = df["Banco"].astype(str).str.strip().str.upper()
    df_out["EMPRESA"] = df.get("EMPRESA", "FD")
    df_out["Cuenta Libradora"] = df.get("Cuenta Libradora", "")
    
    df_out["Fecha Emisión"] = pd.to_datetime(df["Fecha Emisión"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
    df_out["Fecha Pago"] = pd.to_datetime(df["Fecha Pago"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")

    fp = pd.to_datetime(df["Fecha Pago"], dayfirst=True, errors="coerce")
    df_out["DIA"] = fp.dt.dayofweek.map(DIAS_ES)
    df_out["DIA.1"] = fp.dt.day
    df_out["MES"] = fp.dt.month
    df_out["MES.2"] = fp.dt.month.map(MESES_ES)
    df_out["AÑO"] = fp.dt.year

    df_out["Nro. de Cheque"] = df.get("Nro. de Cheque", "")
    df_out["Importe"] = pd.to_numeric(df["Importe"], errors="coerce").fillna(0.0).round(2)
    df_out["CUIT Beneficiario"] = df.get("CUIT Beneficiario", "")
    df_out["Razón Social Beneficiario"] = df.get("Razón Social Beneficiario", "")
    df_out["Estado"] = df.get("Estado", "Emitido")
    df_out["Motivo Anulación"] = df.get("Motivo Anulación", "")

    return df_out


def check_cheque_duplicates(df_exist: pd.DataFrame, new_items: list[dict]) -> tuple[list[dict], list[dict]]:
    """Detecta cheques duplicados según Banco + Nro. de Cheque + Cuenta Libradora."""
    def make_k(b, n, c):
        return f"{str(b or '').strip().upper()}|{str(n or '').strip()}|{str(c or '').strip()}"

    existing_keys = set()
    if df_exist is not None and not df_exist.empty:
        for _, r in df_exist.iterrows():
            if str(r.get("Estado", "")).strip().lower() != "anulado":
                k = make_k(r.get("Banco"), r.get("Nro. de Cheque"), r.get("Cuenta Libradora", ""))
                existing_keys.add(k)

    valids, dups = [], []
    for item in new_items:
        k = make_k(item.get("Banco"), item.get("Nro. de Cheque"), item.get("Cuenta Libradora", ""))
        if k in existing_keys:
            dups.append(item)
        else:
            existing_keys.add(k)
            valids.append(item)

    return valids, dups


def compute_clearing_kpis(df: pd.DataFrame, feriados: list[dt.date] | None = None) -> dict:
    """Calcula métricas de clearing excluyendo valores anulados."""
    if df is None or df.empty or "Importe" not in df.columns:
        return {
            "total_comprometido": 0.0,
            "cant_cheques": 0,
            "dias_habiles": 0,
            "promedio_diario": 0.0,
            "pico_maximo": 0.0,
            "bancos_distribucion": {},
        }

    # Excluir cheques anulados de los cálculos de fondos
    activos = df[df.get("Estado", "Emitido").astype(str).str.lower() != "anulado"].copy()
    if activos.empty:
        return {
            "total_comprometido": 0.0,
            "cant_cheques": 0,
            "dias_habiles": 0,
            "promedio_diario": 0.0,
            "pico_maximo": 0.0,
            "bancos_distribucion": {},
        }

    feriados_set = set(feriados or [])
    total = float(pd.to_numeric(activos["Importe"], errors="coerce").fillna(0.0).sum())
    cant = len(activos)

    fechas_series = pd.to_datetime(activos.get("Fecha Pago"), dayfirst=True, errors="coerce").dropna()
    if fechas_series.empty:
        return {
            "total_comprometido": total,
            "cant_cheques": cant,
            "dias_habiles": 1,
            "promedio_diario": total,
            "pico_maximo": total,
            "bancos_distribucion": {},
        }

    fechas_unicas = fechas_series.dt.date.unique()
    dias_habiles = sum(1 for d in fechas_unicas if d.weekday() < 5 and d not in feriados_set)
    dias_divisor = dias_habiles if dias_habiles > 0 else (len(fechas_unicas) if len(fechas_unicas) > 0 else 1)
    promedio = total / dias_divisor

    df_temp = pd.DataFrame({
        "f": fechas_series.dt.date,
        "m": pd.to_numeric(activos["Importe"], errors="coerce").fillna(0.0)
    })
    pico = float(df_temp.groupby("f")["m"].sum().max()) if not df_temp.empty else 0.0

    bancos_dist = {}
    if total > 0 and "Banco" in activos.columns:
        by_banco = activos.groupby("Banco")["Importe"].sum().sort_values(ascending=False)
        for b, m in by_banco.items():
            bancos_dist[b] = {"monto": float(m), "pct": (float(m) / total) * 100}

    return {
        "total_comprometido": total,
        "cant_cheques": cant,
        "dias_habiles": dias_habiles,
        "promedio_diario": promedio,
        "pico_maximo": pico,
        "bancos_distribucion": bancos_dist,
    }


def build_clearing_matrix(df: pd.DataFrame, fecha_inicio: dt.date | None = None) -> tuple[pd.DataFrame, list[str]]:
    """Genera la sábana de clearing agrupada por día con subtotales mensuales (excluyendo anulados)."""
    if df is None or df.empty:
        return pd.DataFrame(), []

    sub = df[df.get("Estado", "Emitido").astype(str).str.lower() != "anulado"].copy()
    sub["Fecha_Pago_DT"] = pd.to_datetime(sub.get("Fecha Pago"), dayfirst=True, errors="coerce")
    sub = sub.dropna(subset=["Fecha_Pago_DT"])

    if fecha_inicio:
        sub = sub[sub["Fecha_Pago_DT"].dt.date >= fecha_inicio]

    if sub.empty:
        return pd.DataFrame(), []

    if "FECHA_LABEL" not in sub.columns or sub["FECHA_LABEL"].isna().any():
        sub["FECHA_LABEL"] = sub["Fecha_Pago_DT"].apply(
            lambda d: f"{d.day:02d}-{MESES_ABR.get(d.month, '')}" if pd.notna(d) else ""
        )

    bancos = sorted([b for b in sub["Banco"].dropna().unique() if str(b).strip()])

    pivot = sub.pivot_table(
        index=["Fecha_Pago_DT", "FECHA_LABEL"],
        columns="Banco",
        values="Importe",
        aggfunc="sum",
        fill_value=0.0,
    ).reset_index()

    banco_cols = [b for b in bancos if b in pivot.columns]
    pivot["TOTAL"] = pivot[banco_cols].sum(axis=1)
    pivot = pivot.sort_values("Fecha_Pago_DT").reset_index(drop=True)

    rows_with_subtotals = []
    pivot["PERIODO"] = pivot["Fecha_Pago_DT"].dt.to_period("M")

    for periodo, grp in pivot.groupby("PERIODO", sort=False):
        for _, r in grp.iterrows():
            item = r.to_dict()
            item["IS_SUBTOTAL"] = False
            item["MES_KEY"] = f"{periodo.year}-{periodo.month:02d}"
            rows_with_subtotals.append(item)

        mes_nombre = MESES_ES_UPPER.get(periodo.month, "")
        subtot_row = {
            "Fecha_Pago_DT": pd.NaT,
            "FECHA_LABEL": f"SUBTOTAL {mes_nombre} {periodo.year}",
            "IS_SUBTOTAL": True,
            "MES_KEY": f"{periodo.year}-{periodo.month:02d}",
        }
        for b in banco_cols:
            subtot_row[b] = grp[b].sum()
        subtot_row["TOTAL"] = grp["TOTAL"].sum()
        rows_with_subtotals.append(subtot_row)

    df_res = pd.DataFrame(rows_with_subtotals)
    df_res = df_res.rename(columns={"Fecha_Pago_DT": "Fecha Pago"})

    return df_res, banco_cols


# ---------------------------------------------------------------------------
# Módulo BCRA (Sesión Robusta)
# ---------------------------------------------------------------------------

def _get_bcra_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept-Language": "es-AR,es;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
    })
    retries = Retry(
        total=3,
        backoff_factor=0.8,
        status_forcelist=[429, 500, 502, 503, 504],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

BCRA_SESSION = _get_bcra_session()


def clean_cuit(val: str) -> str:
    return re.sub(r"\D", "", str(val or ""))[:11]


def parse_cuits_input(text: str) -> list[str]:
    raw = re.split(r"[\n,; \t]+", str(text or ""))
    cleaned = [clean_cuit(x) for x in raw if clean_cuit(x)]
    seen = set()
    result = []
    for c in cleaned:
        if len(c) == 11 and c not in seen:
            seen.add(c)
            result.append(c)
    return result


def fetch_bcra_data(cuit: str) -> dict:
    deuda_data = None
    deuda_status = "ok"
    try:
        r_deuda = BCRA_SESSION.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/{cuit}",
            timeout=12,
        )
        if r_deuda.status_code == 200:
            deuda_data = r_deuda.json().get("results", {})
        elif r_deuda.status_code == 404:
            deuda_status = "none"
        else:
            deuda_status = "error"
    except Exception:
        deuda_status = "error"

    time.sleep(0.3)

    checks = []
    try:
        r_checks = BCRA_SESSION.get(
            f"https://api.bcra.gob.ar/CentralDeDeudores/v1.0/Deudas/ChequesRechazados/{cuit}",
            timeout=12,
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

    if deuda_status == "ok" and deuda_data:
        periods = deuda_data.get("periodos", [])
        latest = periods[0] if periods else {}
        ents = latest.get("entidades", [])

        worst = max([int(e.get("situacion", 0) or 0) for e in ents], default=0)
        debt = sum([(float(e.get("monto", 0) or 0) * 1000.0) for e in ents])
        rejected = len(checks)
        pending = len([c for c in checks if not c.get("fechaPago")])
        pending_amount = sum([c["monto"] for c in checks if not c.get("fechaPago")])

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
            alerts.append(f"Historial crítico: registra {rejected} cheques rechazados en total.")
        elif rejected > 0 and pending == 0:
            alerts.append(f"Registra {rejected} rechazo(s) histórico(s), pero figuran cancelados/pagados.")

        if not alerts:
            alerts.append("Sin señales negativas: Situación normal y sin cheques rechazados.")

        sorted_checks = sorted(checks, key=lambda x: str(x.get("fechaRechazo", "")), reverse=True)
        last3 = sorted_checks[:3]

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
            "alerts": ["Error temporal de comunicación con los servidores del BCRA."],
            "last3": [],
        }
