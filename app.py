# -*- coding: utf-8 -*-
"""
Suite Financiera Corporativa — L2 For Drink SA
====================================================================
Módulos:
1. Gestión de Tesorería (Flujo de Caja, Escaleras y Conciliación)
2. Analizador de Libradores · BCRA (Scoring y Cheques Rechazados)
"""

from __future__ import annotations

import datetime as dt
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    from streamlit_gsheets import GSheetsConnection
    HAS_GSHEETS = True
except ImportError:
    HAS_GSHEETS = False

import core

# ---------------------------------------------------------------------------
# Configuración general y estilos
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Suite Financiera — L2 For Drink SA",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    [data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        white-space: nowrap !important;
        overflow: visible !important;
        text-overflow: unset !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #555555;
    }
    
    /* Estilos del Analizador BCRA fieles a la V2 */
    .bcra-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 14px 18px;
    }
    .bcra-label {
        font-size: 11px;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.05em;
        color: #64748b;
    }
    .bcra-val {
        font-size: 26px;
        font-weight: 700;
        margin-top: 4px;
        line-height: 1.1;
    }
    .badge-sit {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        font-weight: 700;
        font-size: 12px;
    }
    .badge-sit-1 { background-color: #dcfce7; color: #166534; }
    .badge-sit-2 { background-color: #fef9c3; color: #854d0e; }
    .badge-sit-bad { background-color: #fee2e2; color: #991b1b; }
    
    .pill-status {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 700;
    }
    .pill-alerta { background-color: #fee2e2; color: #991b1b; }
    .pill-revisar { background-color: #fef9c3; color: #854d0e; }
    .pill-ok { background-color: #dcfce7; color: #166534; }
    </style>
    """,
    unsafe_allow_html=True,
)

SEED_PATH = Path(__file__).parent / "data_seed.csv"

# ---------------------------------------------------------------------------
# Control de Acceso
# ---------------------------------------------------------------------------

params = st.query_params
is_admin_param = params.get("admin", "") == "1"

if "is_admin" not in st.session_state:
    st.session_state.is_admin = is_admin_param

# ---------------------------------------------------------------------------
# Conexión persistente a Google Sheets
# ---------------------------------------------------------------------------

conn = None
if HAS_GSHEETS:
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
    except Exception:
        conn = None


def load_data_source() -> pd.DataFrame:
    if conn is not None:
        try:
            df_cloud = conn.read(ttl="0s")
            if df_cloud is not None and not df_cloud.empty:
                return core.normalize_df(df_cloud)
        except Exception as e:
            st.sidebar.warning(f"Aviso Google Sheets: {e}")
    if SEED_PATH.exists():
        return core.load_seed_csv(SEED_PATH)
    return core.normalize_df(pd.DataFrame(columns=core.COLUMNS))


def save_data_source(df: pd.DataFrame) -> bool:
    if not st.session_state.is_admin:
        st.error("Acceso de solo lectura: no tiene permisos para modificar la base.")
        return False

    df_to_save = core.normalize_df(df)
    df_out = df_to_save[core.COLUMNS].copy()
    if "fecha" in df_out.columns:
        df_out["fecha"] = df_out["fecha"].dt.strftime("%Y-%m-%d")

    if conn is not None:
        try:
            conn.update(data=df_out)
            return True
        except Exception as e:
            st.error(f"Error al sincronizar con Google Sheets: {e}")
            return False
    return False


def init_state() -> None:
    if "df" not in st.session_state:
        st.session_state.df = load_data_source()
    if "editing_id" not in st.session_state:
        st.session_state.editing_id = None
    if "sede_global" not in st.session_state:
        st.session_state.sede_global = core.SEDE_CONSOLIDADO


init_state()


def get_df() -> pd.DataFrame:
    return st.session_state.df


def set_df(df: pd.DataFrame, sync_cloud: bool = True) -> None:
    norm = core.normalize_df(df)
    st.session_state.df = norm
    if sync_cloud and st.session_state.is_admin:
        save_data_source(norm)


def fmt_ars(value: float) -> str:
    try:
        return f"$ {value:,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "$ 0"


# ---------------------------------------------------------------------------
# Barra lateral: Selector de Módulos (Suite Financiera)
# ---------------------------------------------------------------------------

st.sidebar.title("💼 Suite Financiera")
st.sidebar.caption("L2 For Drink SA · Centro Operativo")

if conn is not None:
    st.sidebar.markdown(
        """
        <div style="display:flex; align-items:center; gap:6px; color:#2e7d32; font-size:0.85rem; font-weight:500; margin-bottom:12px;">
            <span style="height:8px; width:8px; background-color:#2e7d32; border-radius:50%; display:inline-block;"></span>
            Base central conectada
        </div>
        """,
        unsafe_allow_html=True,
    )

# Selector preponderante de módulos
if st.session_state.is_admin:
    modulo_activo = st.sidebar.radio(
        "Herramienta Activa",
        ["💵 Gestión de Tesorería", "🔍 Analizador de Libradores · BCRA"],
        index=0,
    )
else:
    modulo_activo = "💵 Gestión de Tesorería"

st.sidebar.divider()

# Login Admin
if not st.session_state.is_admin:
    with st.sidebar.expander("🔒 Modo Operador / Admin", expanded=False):
        admin_pass = st.text_input("Clave de edición", type="password", key="pass_admin_input")
        if st.button("Habilitar"):
            if admin_pass == "fordrink2026":
                st.session_state.is_admin = True
                st.rerun()
            else:
                st.error("Clave incorrecta")
else:
    st.sidebar.markdown("<small style='color:#1F4E78; font-weight:600;'>🔑 Perfil: Administrador</small>", unsafe_allow_html=True)
    if st.sidebar.button("Cerrar sesión Admin"):
        st.session_state.is_admin = False
        st.rerun()

# ===========================================================================
# MÓDULO 1: GESTIÓN DE TESORERÍA
# ===========================================================================
if modulo_activo == "💵 Gestión de Tesorería":
    st.sidebar.markdown("##### 🏢 Sede / Tesorería")
    sede_global = st.sidebar.segmented_control(
        "Sede",
        options=core.OPCIONES_SEDE,
        default=st.session_state.sede_global,
        label_visibility="collapsed",
    )
    if sede_global is None:
        sede_global = core.SEDE_CONSOLIDADO
    st.session_state.sede_global = sede_global

    umbral_tension = 30_000_000.0 if sede_global == core.SEDE_CONSOLIDADO else 15_000_000.0

    if st.session_state.is_admin:
        with st.sidebar.expander("📁 Importar / Actualizar base", expanded=False):
            uploaded = st.file_uploader("Subir archivo Excel", type=["xlsx"], key="uploader_replace")
            if uploaded is not None and st.button("Reemplazar base actual", type="primary"):
                try:
                    new_df = core.load_from_uploaded_xlsx(uploaded)
                    set_df(new_df, sync_cloud=True)
                    st.success(f"Cargados {len(new_df)} pagos.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

            st.markdown("---")
            if st.button("🔄 Recargar datos desde Google Sheets"):
                st.session_state.df = load_data_source()
                st.rerun()
        st.sidebar.divider()

    st.sidebar.subheader("🔎 Filtros")
    df_all_raw = get_df()
    df_all = core.filter_by_sede(df_all_raw, sede_global)

    proveedores_disponibles = sorted(df_all["proveedor"].dropna().unique().tolist())
    f_proveedores = st.sidebar.multiselect("Proveedor", proveedores_disponibles, default=proveedores_disponibles)
    f_estados = st.sidebar.multiselect("Estado", core.ESTADOS_VALIDOS, default=core.ESTADOS_VALIDOS)

    min_date = df_all["fecha"].min()
    max_date = df_all["fecha"].max()
    if pd.isna(min_date) or pd.isna(max_date):
        min_date = max_date = pd.Timestamp.today()

    f_rango = st.sidebar.date_input(
        "Rango de fechas",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )

    meses_disponibles = sorted(df_all["mes"].dropna().unique().tolist())
    f_meses = st.sidebar.multiselect("Mes", meses_disponibles, default=meses_disponibles)

    def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        if f_proveedores:
            out = out[out["proveedor"].isin(f_proveedores)]
        if f_estados:
            out = out[out["estado"].isin(f_estados)]
        if f_meses:
            out = out[out["mes"].isin(f_meses)]
        if isinstance(f_rango, tuple) and len(f_rango) == 2:
            d0, d1 = pd.Timestamp(f_rango[0]), pd.Timestamp(f_rango[1])
            out = out[(out["fecha"] >= d0) & (out["fecha"] <= d1)]
        return out

    df_filtered = apply_filters(df_all)

    st.title("Gestión de Tesorería — Planes de Pago")
    st.caption(f"L2 For Drink SA · Sede activa: **{sede_global}** · Flujo de caja en tiempo real.")

    if st.session_state.is_admin:
        tab_kpi, tab_matriz, tab_escaleras, tab_conciliar, tab_asistente, tab_abm, tab_export = st.tabs([
            "📊 Panel Ejecutivo",
            "🗓️ Matriz Semanal",
            "🪜 Detalle de Escaleras",
            "⚡ Conciliación Rápida",
            "➕ Cargar Nueva Escalera",
            "📝 Registro Manual (ABM)",
            "⬇️ Exportar",
        ])
    else:
        tab_kpi, tab_matriz, tab_escaleras = st.tabs([
            "📊 Panel Ejecutivo",
            "🗓️ Matriz Semanal",
            "🪜 Detalle de Escaleras",
        ])

    with tab_kpi:
        kpis = core.compute_kpis(df_filtered)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Comprometido", fmt_ars(kpis.total_comprometido))
        c2.metric("Total Pagado", fmt_ars(kpis.total_pagado))
        c3.metric("Saldo Pendiente", fmt_ars(kpis.saldo_pendiente))
        c4.metric("Desembolso Promedio", fmt_ars(kpis.desembolso_promedio))

        st.divider()
        left, right = st.columns([3, 2])
        with left:
            st.subheader("Distribución mensual proyectada")
            dist = core.monthly_distribution(df_filtered)
            if not dist.empty:
                fig = px.bar(
                    dist, x="mes", y="importe_ars", text_auto=".2s",
                    labels={"mes": "Mes", "importe_ars": "Importe ARS"},
                    color_discrete_sequence=["#1F4E78"],
                )
                fig.update_layout(yaxis_title="Importe ARS", xaxis_title="", margin=dict(t=10, b=10), height=340)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sin datos para graficar.")

            if sede_global == core.SEDE_CONSOLIDADO and not df_filtered.empty:
                st.subheader("Distribución por Sede")
                por_sede = df_filtered.groupby("tesoreria")["importe_ars"].sum().reset_index()
                fig_sede = px.pie(por_sede, names="tesoreria", values="importe_ars", hole=0.45,
                                  color_discrete_sequence=["#1F4E78", "#8FAADC"])
                fig_sede.update_layout(margin=dict(t=10, b=10), height=280)
                st.plotly_chart(fig_sede, use_container_width=True)

        with right:
            st.subheader("🚦 Carga Semanal y Semáforos de Pago")
            st.caption(f"Semáforo rojo (> {fmt_ars(umbral_tension)}) según sede activa ({sede_global}).")

            hoy_ts = pd.Timestamp(dt.date.today())
            lunes_actual = hoy_ts - pd.Timedelta(days=hoy_ts.weekday())

            df_futuro = df_filtered[df_filtered["lunes_semana"] >= lunes_actual].copy()

            if df_futuro.empty:
                st.info("No hay desembolsos pendientes programados desde la semana actual.")
            else:
                sem_futuras = (
                    df_futuro.groupby(["lunes_semana", "semana_etiqueta"])["importe_ars"]
                    .sum()
                    .reset_index()
                    .sort_values("lunes_semana")
                )
                promedio_futuro = sem_futuras["importe_ars"].mean()

                st.markdown(f"**Promedio semanal proyectado:** `{fmt_ars(promedio_futuro)}` ({len(sem_futuras)} semanas)")

                for _, r_sem in sem_futuras.iterrows():
                    monto_s = r_sem["importe_ars"]
                    sem_label = r_sem["semana_etiqueta"]

                    if monto_s > umbral_tension or monto_s > (promedio_futuro * 1.3):
                        icono = "🔴"
                        msj = f"**{sem_label}** — {fmt_ars(monto_s)} *(Tensión Alta)*"
                        st.error(msj, icon=icono)
                    elif monto_s >= (promedio_futuro * 0.8):
                        icono = "🟡"
                        msj = f"**{sem_label}** — {fmt_ars(monto_s)} *(En Promedio)*"
                        st.warning(msj, icon=icono)
                    else:
                        icono = "🟢"
                        msj = f"**{sem_label}** — {fmt_ars(monto_s)} *(Holgada)*"
                        st.success(msj, icon=icono)

        st.divider()
        st.subheader("Compromiso y Avance por Proveedor")
        ctrl = core.provider_control(df_filtered)
        if not ctrl.empty:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(y=ctrl["proveedor"], x=ctrl["compromiso_total_ars"], orientation="h",
                                   name="Comprometido", marker_color="#D9D9D9"))
            fig2.add_trace(go.Bar(y=ctrl["proveedor"], x=ctrl["pagado_ars"], orientation="h",
                                   name="Pagado", marker_color="#1F4E78"))
            fig2.update_layout(barmode="overlay", xaxis_title="Importe ARS", height=max(320, 26 * len(ctrl)),
                                margin=dict(t=10, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(fig2, use_container_width=True)

    with tab_matriz:
        st.subheader("Matriz Semanal de Flujo de Efectivo (Consolidada)")
        st.caption("Visión consolidada por proveedor y semana. Reportes listos para imprimir o auditar.")

        if df_filtered.empty:
            st.info("Sin datos para mostrar con los filtros aplicados.")
        else:
            df_sorted_weeks = df_filtered.sort_values("lunes_semana")
            semanas_ordenadas = (
                df_sorted_weeks[["lunes_semana", "semana_etiqueta"]]
                .drop_duplicates()
                .dropna()
            )

            opciones_semanas = semanas_ordenadas["semana_etiqueta"].tolist()

            hoy_ts = pd.Timestamp(dt.date.today())
            lunes_actual = hoy_ts - pd.Timedelta(days=hoy_ts.weekday())

            idx_default = 0
            semanas_desde_hoy = semanas_ordenadas[semanas_ordenadas["lunes_semana"] >= lunes_actual]
            if not semanas_desde_hoy.empty:
                primera_sem_label = semanas_desde_hoy["semana_etiqueta"].iloc[0]
                if primera_sem_label in opciones_semanas:
                    idx_default = opciones_semanas.index(primera_sem_label)

            col_filtro_sem, col_btn_excel, col_btn_pdf = st.columns([2, 1.2, 1.2])

            with col_filtro_sem:
                semana_inicio_sel = st.selectbox(
                    "📅 Mostrar semanas desde:",
                    options=opciones_semanas,
                    index=idx_default,
                    help="Abre por defecto en la semana en curso.",
                )

            lunes_corte = semanas_ordenadas.loc[
                semanas_ordenadas["semana_etiqueta"] == semana_inicio_sel, "lunes_semana"
            ].iloc[0]

            df_matriz_periodo = df_filtered[df_filtered["lunes_semana"] >= lunes_corte]
            matrix = core.build_weekly_matrix(df_matriz_periodo)

            if matrix.empty:
                st.info("No hay pagos para el período seleccionado.")
            else:
                def export_matrix_to_excel(matrix_df: pd.DataFrame) -> bytes:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Matriz Semanal"
                    ws["A1"] = f"MATRIZ SEMANAL DE FLUJO DE EFECTIVO — {sede_global.upper()}"
                    ws["A1"].font = Font(size=13, bold=True, color="1F4E78")
                    ws["A2"] = f"Período: Desde semana {semana_inicio_sel} · Generado el: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"
                    ws["A2"].font = Font(size=9, italic=True, color="666666")

                    headers = list(matrix_df.columns)
                    h_row = 4
                    for c_idx, h in enumerate(headers, 1):
                        cell = ws.cell(row=h_row, column=c_idx, value=h)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                    peak_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    total_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")

                    for r_idx, (_, row) in enumerate(matrix_df.iterrows()):
                        row_num = h_row + 1 + r_idx
                        is_total_row = (row.get("Semana") == "TOTAL POR PROVEEDOR")
                        total_semanal = row.get("TOTAL SEMANAL", 0)
                        is_peak = (not is_total_row and isinstance(total_semanal, (int, float)) and total_semanal > umbral_tension)

                        for c_idx, h in enumerate(headers, 1):
                            val = row[h]
                            cell = ws.cell(row=row_num, column=c_idx)

                            if is_total_row:
                                cell.font = Font(bold=True)
                                cell.fill = total_fill
                            elif is_peak and h == "TOTAL SEMANAL":
                                cell.font = Font(bold=True)
                                cell.fill = peak_fill

                            if h == "Semana":
                                cell.value = str(val)
                                cell.alignment = Alignment(horizontal="center")
                            else:
                                cell.value = float(val) if pd.notna(val) else 0.0
                                cell.number_format = "$ #,##0"

                    for col in ws.columns:
                        max_len = max(len(str(cell.value or "")) for cell in col)
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

                    buf = BytesIO()
                    wb.save(buf)
                    return buf.getvalue()

                def export_matrix_to_pdf(matrix_df: pd.DataFrame) -> bytes:
                    buf = BytesIO()
                    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=25)
                    styles = getSampleStyleSheet()
                    story = []

                    title_style = ParagraphStyle('MatTitle', parent=styles['Heading1'], fontSize=13, textColor=colors.HexColor('#1F4E78'), spaceAfter=3)
                    sub_style = ParagraphStyle('MatSub', parent=styles['Normal'], fontSize=8, textColor=colors.HexColor('#555555'), spaceAfter=10)

                    story.append(Paragraph(f"<b>L2 FOR DRINK SA — MATRIZ SEMANAL DE FLUJO DE EFECTIVO ({sede_global.upper()})</b>", title_style))
                    story.append(Paragraph(f"Período: Desde semana {semana_inicio_sel} | Fecha de emisión: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}", sub_style))

                    headers = list(matrix_df.columns)
                    table_data = [headers]
                    for _, r in matrix_df.iterrows():
                        row_vals = []
                        for h in headers:
                            v = r[h]
                            row_vals.append(str(v) if h == "Semana" else (fmt_ars(v) if pd.notna(v) else "$ 0"))
                        table_data.append(row_vals)

                    col_w = max(40, int(780 / len(headers)))
                    t = Table(table_data, colWidths=[col_w] * len(headers))
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 7),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#DDDDDD')),
                        ('FONTSIZE', (0, 1), (-1, -1), 6.5),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EFEFEF')),
                        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                    ]))
                    story.append(t)
                    doc.build(story)
                    return buf.getvalue()

                with col_btn_excel:
                    st.write("")
                    st.write("")
                    excel_matriz = export_matrix_to_excel(matrix)
                    st.download_button(
                        label="📥 Descargar Reporte en Excel",
                        data=excel_matriz,
                        file_name=f"Matriz_Semanal_{sede_global}_{dt.date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        key="dl_matriz_excel",
                        use_container_width=True,
                    )

                with col_btn_pdf:
                    st.write("")
                    st.write("")
                    pdf_matriz = export_matrix_to_pdf(matrix)
                    st.download_button(
                        label="📄 Descargar Reporte en PDF",
                        data=pdf_matriz,
                        file_name=f"Matriz_Semanal_{sede_global}_{dt.date.today().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key="dl_matriz_pdf",
                        use_container_width=True,
                    )

                st.markdown("---")
                provider_cols = [c for c in matrix.columns if c not in ("Semana", "TOTAL SEMANAL", "ACUMULADO")]
                def highlight_peaks(row):
                    styles = [""] * len(row)
                    if row["Semana"] == "TOTAL POR PROVEEDOR":
                        return ["font-weight: bold; background-color: #EFEFEF"] * len(row)
                    try:
                        if row["TOTAL SEMANAL"] > umbral_tension:
                            idx = list(row.index).index("TOTAL SEMANAL")
                            styles[idx] = "background-color:#FFC7CE; font-weight:bold;"
                    except Exception:
                        pass
                    return styles

                styled = matrix.style.apply(highlight_peaks, axis=1).format({c: fmt_ars for c in provider_cols + ["TOTAL SEMANAL", "ACUMULADO"]})
                st.dataframe(styled, use_container_width=True, height=520, hide_index=True)

    with tab_escaleras:
        st.subheader("🪜 Control Individual de Escaleras y Presupuestos")
        st.caption("Inspeccioná cada contrato y cuota de forma vertical con descarga en Excel y PDF para legajos.")

        prov_list = sorted(df_filtered["proveedor"].dropna().unique().tolist())
        if not prov_list:
            st.info("No hay proveedores para los filtros seleccionados.")
        else:
            col_sel_p, col_sel_op = st.columns([1, 1])
            with col_sel_p:
                p_sel = st.selectbox("Seleccionar Proveedor", prov_list, key="p_inspect")

            df_p_raw = df_filtered[df_filtered["proveedor"] == p_sel].copy()

            def clean_op_label(r):
                c = str(r["concepto"]).strip() if pd.notna(r["concepto"]) and str(r["concepto"]).strip() else ""
                return c if c else "Operación General / Presupuesto Base"

            df_p_raw["operacion_identificada"] = df_p_raw.apply(clean_op_label, axis=1)
            operaciones_encontradas = sorted(df_p_raw["operacion_identificada"].unique().tolist())

            with col_sel_op:
                if len(operaciones_encontradas) > 1:
                    opciones_operacion = ["Todas las Operaciones (Consolidado)"] + operaciones_encontradas
                    op_sel = st.selectbox(f"📌 Operación / Escalera ({len(operaciones_encontradas)} registradas)", opciones_operacion, key="op_inspect")
                else:
                    op_sel = "Todas las Operaciones (Consolidado)"
                    st.selectbox("📌 Operación / Escalera", ["Operación Única"], disabled=True, key="op_inspect_single")

            df_p = df_p_raw[df_p_raw["operacion_identificada"] == op_sel].copy() if op_sel != "Todas las Operaciones (Consolidado)" else df_p_raw.copy()
            subtitulo_op = f" — {op_sel}" if op_sel != "Todas las Operaciones (Consolidado)" else ""

            kpis_p = core.compute_kpis(df_p)
            st.markdown(f"#### Ficha: **{p_sel}**{subtitulo_op}")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Compromiso Total", fmt_ars(kpis_p.total_comprometido))
            m2.metric("Total Pagado", fmt_ars(kpis_p.total_pagado))
            m3.metric("Saldo Pendiente", fmt_ars(kpis_p.saldo_pendiente))
            pct_p = (kpis_p.total_pagado / kpis_p.total_comprometido * 100) if kpis_p.total_comprometido else 0
            m4.metric("% Cancelado", f"{pct_p:.1f}%")
            st.progress(min(pct_p / 100, 1.0))

            st.divider()

            def merge_concepto_obs(r):
                c = str(r["concepto"]).strip() if pd.notna(r["concepto"]) else ""
                o = str(r["obs"]).strip() if pd.notna(r["obs"]) else ""
                if c and o:
                    return f"{c} ({o})" if c != o else c
                return c or o or "Presupuesto Único"

            df_p["presupuesto_obs"] = df_p.apply(merge_concepto_obs, axis=1)

            def format_ladder_table(sub_df: pd.DataFrame) -> pd.DataFrame:
                t = sub_df.sort_values("fecha").copy()
                return pd.DataFrame({
                    "Fecha Vto.": t["fecha"],
                    "Presupuesto / Observaciones": t["presupuesto_obs"],
                    "Importe Pactado": t["importe"],
                    "TC": t["tc"],
                    "Importe ARS": t["importe_ars"],
                    "Estado": t["estado"],
                    "Moneda": t["moneda"],
                })

            def export_provider_to_excel(df_table: pd.DataFrame, prov_name: str, op_name: str) -> bytes:
                wb = Workbook()
                ws = wb.active
                ws.title = "Escalera de Pagos"
                titulo_excel = f"ESCALERA DE PAGOS — {prov_name.upper()}" + (f" ({op_name.upper()})" if op_name != "Todas las Operaciones (Consolidado)" else "")
                ws["A1"] = titulo_excel
                ws["A1"].font = Font(size=13, bold=True, color="1F4E78")
                ws["A2"] = f"Generado el: {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"
                ws["A2"].font = Font(size=9, italic=True, color="666666")

                headers = list(df_table.columns)
                h_row = 4
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=h_row, column=col_num, value=header)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                for r_idx, (_, row) in enumerate(df_table.iterrows()):
                    row_num = h_row + 1 + r_idx
                    for c_idx, h in enumerate(headers, 1):
                        val = row[h]
                        cell = ws.cell(row=row_num, column=c_idx)
                        if h == "Fecha Vto." and pd.notna(val):
                            cell.value = val.strftime("%d/%m/%Y") if hasattr(val, "strftime") else str(val)
                            cell.alignment = Alignment(horizontal="center")
                        elif h in ("Importe Pactado", "TC"):
                            cell.value = float(val) if pd.notna(val) else 0.0
                            cell.number_format = "#,##0.00"
                        elif h == "Importe ARS":
                            cell.value = float(val) if pd.notna(val) else 0.0
                            cell.number_format = "$ #,##0"
                        else:
                            cell.value = str(val) if pd.notna(val) else ""

                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                buf = BytesIO()
                wb.save(buf)
                return buf.getvalue()

            def export_provider_to_pdf(df_table: pd.DataFrame, prov_name: str, op_name: str, kpis_info) -> bytes:
                buf = BytesIO()
                doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                styles = getSampleStyleSheet()
                story = []

                title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor('#1F4E78'), spaceAfter=4)
                sub_style = ParagraphStyle('DocSub', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#444444'), spaceAfter=10)
                kpi_style = ParagraphStyle('DocKpi', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#1F4E78'), spaceAfter=14)

                titulo_txt = f"L2 FOR DRINK SA — PLAN DE PAGO: {prov_name.upper()}" + (f" - {op_name.upper()}" if op_name != "Todas las Operaciones (Consolidado)" else "")
                story.append(Paragraph(f"<b>{titulo_txt}</b>", title_style))
                story.append(Paragraph(f"Fecha de Emisión: <b>{dt.datetime.now().strftime('%d/%m/%Y %H:%M')}</b> | Sede: <b>{sede_global}</b>", sub_style))

                resumen_txt = f"<b>Compromiso Total:</b> {fmt_ars(kpis_info.total_comprometido)} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Total Pagado:</b> {fmt_ars(kpis_info.total_pagado)} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Saldo Pendiente:</b> {fmt_ars(kpis_info.saldo_pendiente)}"
                story.append(Paragraph(resumen_txt, kpi_style))

                pdf_data = [["Fecha Vto.", "Detalle / Concepto", "Importe ARS", "Estado"]]
                for _, r in df_table.iterrows():
                    f_str = r["Fecha Vto."].strftime("%d/%m/%Y") if hasattr(r["Fecha Vto."], "strftime") else str(r["Fecha Vto."])
                    c_str = str(r["Presupuesto / Observaciones"])[:45]
                    m_str = fmt_ars(r["Importe ARS"])
                    e_str = str(r["Estado"])
                    pdf_data.append([f_str, c_str, m_str, e_str])

                t = Table(pdf_data, colWidths=[80, 240, 115, 85])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
                    ('FONTSIZE', (0, 1), (-1, -1), 8.5),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ]))
                story.append(t)
                story.append(Spacer(1, 40))

                firmas_data = [
                    ["___________________________________", "___________________________________"],
                    ["Firma y Aclaración Tesorería", "Conformidad Contratista / Proveedor"],
                    ["L2 For Drink SA", f"{prov_name}"],
                ]
                t_firmas = Table(firmas_data, colWidths=[260, 260])
                t_firmas.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#555555')),
                    ('TOPPADDING', (0, 0), (-1, -1), 2),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ]))
                story.append(t_firmas)
                doc.build(story)
                return buf.getvalue()

            table_view = format_ladder_table(df_p)
            excel_bytes = export_provider_to_excel(table_view, p_sel, op_sel)
            pdf_bytes = export_provider_to_pdf(table_view, p_sel, op_sel, kpis_p)

            c_down_xl, c_down_pdf, _ = st.columns([1.3, 1.3, 2])
            with c_down_xl:
                st.download_button("📥 Descargar Reporte en Excel", excel_bytes, f"Escalera_{p_sel.replace(' ', '_')}_{dt.date.today().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_xl_{p_sel}_{op_sel}", type="primary", use_container_width=True)
            with c_down_pdf:
                st.download_button("📄 Descargar Reporte en PDF", pdf_bytes, f"Legajo_Escalera_{p_sel.replace(' ', '_')}_{dt.date.today().strftime('%Y%m%d')}.pdf", "application/pdf", key=f"dl_pdf_{p_sel}_{op_sel}", use_container_width=True)

            st.dataframe(table_view.style.format({"Fecha Vto.": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "", "Importe Pactado": "{:,.2f}", "TC": "{:,.2f}", "Importe ARS": fmt_ars}), use_container_width=True, hide_index=True)

    if st.session_state.is_admin:
        with tab_conciliar:
            st.subheader("⚡ Conciliación Rápida de Pagos")
            st.caption("Cambiá el estado de los desembolsos a 'Pagado' en un solo clic y sincronizalo con Google Drive.")
            df_prog = df_filtered[df_filtered["estado"] != "Pagado"].sort_values("fecha").copy()
            if df_prog.empty:
                st.success("🎉 ¡Excelente! No hay pagos pendientes en este rango.")
            else:
                st.write(f"Se encontraron **{len(df_prog)} pagos pendientes** con los filtros activos.")
                df_editor = df_prog[["id", "tesoreria", "fecha", "proveedor", "concepto", "importe_ars", "estado", "obs"]].copy()
                df_editor.insert(0, "Conciliar (Marcar Pagado)", False)
                df_editor["fecha"] = df_editor["fecha"].dt.date

                edited = st.data_editor(
                    df_editor,
                    column_config={
                        "Conciliar (Marcar Pagado)": st.column_config.CheckboxColumn("¿Pagar?", default=False),
                        "importe_ars": st.column_config.NumberColumn("Importe ARS", format="$ %d"),
                        "fecha": st.column_config.DateColumn("Fecha Vencimiento", format="DD/MM/YYYY"),
                    },
                    disabled=["id", "tesoreria", "fecha", "proveedor", "concepto", "importe_ars", "estado", "obs"],
                    use_container_width=True,
                    hide_index=True,
                    height=380,
                    key="conciliar_editor",
                )

                pagos_a_conciliar = edited[edited["Conciliar (Marcar Pagado)"] == True]["id"].tolist()
                c_btn, c_resumen = st.columns([1, 2])
                with c_btn:
                    if st.button("💾 Marcar Seleccionados como PAGADO", type="primary", disabled=len(pagos_a_conciliar) == 0):
                        df_current = get_df()
                        df_updated = core.update_payments_status(df_current, pagos_a_conciliar, new_status="Pagado")
                        set_df(df_updated, sync_cloud=True)
                        st.success(f"¡Se actualizaron {len(pagos_a_conciliar)} pagos a estado PAGADO en la nube!")
                        st.rerun()

                with c_resumen:
                    if pagos_a_conciliar:
                        suma_sel = edited[edited["Conciliar (Marcar Pagado)"] == True]["importe_ars"].sum()
                        st.info(f"Seleccionados **{len(pagos_a_conciliar)} pagos** por un total de **{fmt_ars(suma_sel)}**.")

        with tab_asistente:
            st.subheader("➕ Cargar Nueva Escalera de Pago")
            st.caption("Completá los datos del acuerdo y cargá los tramos en la grilla.")

            if "escalera_guardada_msj" in st.session_state:
                st.success(st.session_state.escalera_guardada_msj, icon="✅")
                st.toast(st.session_state.escalera_guardada_msj, icon="🚀")
                del st.session_state["escalera_guardada_msj"]

            c_cab1, c_cab2, c_cab3 = st.columns(3)
            with c_cab1:
                esc_tesoreria = st.selectbox("Tesorería / Sede", core.TESORERIAS_VALIDAS, index=0 if sede_global == "Tucumán" else 1, key="esc_sede")
                tipo_prov = st.radio("Proveedor", ["Existente", "Nuevo"], horizontal=True, key="esc_tipo_prov")
                prov_opts = sorted(get_df()["proveedor"].dropna().unique().tolist())
                esc_proveedor = st.selectbox("Seleccionar de la lista", prov_opts, key="esc_prov_sel") if tipo_prov == "Existente" and prov_opts else st.text_input("Escribir nombre del nuevo proveedor", key="esc_prov_new").strip().upper()

            with c_cab2:
                esc_proyecto = st.text_input("Proyecto / Obra", value="L2 TUC" if esc_tesoreria == "Tucumán" else "L2 BA", key="esc_proy")
                esc_concepto = st.text_input("Presupuesto / Identificador de Operación", placeholder="Ej: Carpintería Aluminio - Presupuesto N° 450", key="esc_conc")

            with c_cab3:
                esc_moneda = st.selectbox("Moneda del Acuerdo", core.MONEDAS_VALIDAS, key="esc_mon")
                esc_tc = st.number_input("Tipo de Cambio acordado (USD)", min_value=1.0, value=1400.0, step=10.0, key="esc_tc_usd") if esc_moneda == "USD" else 1.0
                if esc_moneda == "ARS":
                    st.text_input("Tipo de Cambio", value="1.00 (ARS)", disabled=True)

            st.markdown("---")
            cant_cuotas = st.number_input("Cantidad de tramos / cuotas", min_value=1, max_value=50, value=4, step=1, key="esc_n_cuotas")

            if "grid_ladder_data" not in st.session_state or len(st.session_state.grid_ladder_data) != cant_cuotas or st.session_state.get("reset_grid_flag", False):
                base_f = dt.date.today()
                st.session_state.grid_ladder_data = pd.DataFrame([{"Tramo": f"Cuota {i + 1}", "Fecha Vto.": base_f + dt.timedelta(days=7 * i), "Importe": 0.0} for i in range(int(cant_cuotas))])
                st.session_state["reset_grid_flag"] = False

            grid_edited = st.data_editor(
                st.session_state.grid_ladder_data,
                column_config={
                    "Tramo": st.column_config.TextColumn("Tramo", disabled=True),
                    "Fecha Vto.": st.column_config.DateColumn("Fecha Vto.", format="DD/MM/YYYY", required=True),
                    "Importe": st.column_config.NumberColumn(f"Importe en {esc_moneda}", min_value=0.0, format="%.2f", required=True),
                },
                use_container_width=True,
                hide_index=True,
                key="grid_cuotas_editor",
            )

            total_acordado = float(grid_edited["Importe"].sum())
            total_ars_calc = total_acordado * esc_tc if esc_moneda == "USD" else total_acordado
            st.info(f"📊 **Resumen:** {cant_cuotas} tramos · Total: **{esc_moneda} {total_acordado:,.2f}** {f'(Equivalente: **{fmt_ars(total_ars_calc)}**)' if esc_moneda == 'USD' else ''}")

            if st.button("🚀 Guardar Escalera Completa en la Base", type="primary", use_container_width=True):
                if not esc_proveedor or not esc_concepto or total_acordado <= 0:
                    st.error("Verificá los datos ingresados.")
                else:
                    with st.spinner("Guardando en Google Sheets..."):
                        items_ladder = [{"fecha": pd.to_datetime(r["Fecha Vto."]), "importe": float(r["Importe"])} for _, r in grid_edited.iterrows()]
                        df_curr = get_df()
                        new_rows = core.generate_custom_ladder(df_curr, esc_tesoreria, esc_proveedor, esc_proyecto, esc_concepto, esc_moneda, esc_tc, items_ladder)
                        df_total = pd.concat([df_curr[core.COLUMNS], new_rows[core.COLUMNS]], ignore_index=True)
                        if save_data_source(df_total):
                            st.session_state.df = core.normalize_df(df_total)
                            st.session_state["reset_grid_flag"] = True
                            st.session_state["escalera_guardada_msj"] = f"¡Escalera para {esc_proveedor} guardada con éxito ({len(new_rows)} cuotas)!"
                            time.sleep(0.5)
                            st.rerun()

        with tab_abm:
            st.subheader("Registro y edición manual de pagos individuales")
            col_form, col_table = st.columns([1, 2])
            with col_form:
                editing_id = st.session_state.editing_id
                df_current = get_df()
                record = df_current[df_current["id"] == editing_id].iloc[0] if editing_id is not None and not df_current[df_current["id"] == editing_id].empty else None

                st.markdown(f"##### {'✏️ Editar pago ' + editing_id if record is not None else '➕ Nuevo pago individual'}")
                proveedores_existentes = sorted(df_current["proveedor"].dropna().unique().tolist())

                with st.form("form_pago_manual", clear_on_submit=False):
                    tesoreria_sel = st.selectbox("Sede", core.TESORERIAS_VALIDAS, index=core.TESORERIAS_VALIDAS.index(record["tesoreria"]) if record is not None else 0)
                    fecha_pago = st.date_input("Fecha", value=record["fecha"].date() if record is not None and pd.notna(record["fecha"]) else dt.date.today())
                    modo_p = st.radio("Proveedor", ["Existente", "Nuevo"], horizontal=True)
                    proveedor = st.selectbox("Proveedor", proveedores_existentes, index=proveedores_existentes.index(record["proveedor"]) if record is not None and record["proveedor"] in proveedores_existentes else 0) if modo_p == "Existente" and proveedores_existentes else st.text_input("Nombre nuevo", value=record["proveedor"] if record is not None else "").strip().upper()
                    proyecto = st.text_input("Proyecto", value=record["proyecto"] if record is not None else "L2")
                    concepto = st.text_input("Operación / Concepto", value=record["concepto"] if record is not None else "")
                    c_mon, c_tc = st.columns(2)
                    moneda = c_mon.selectbox("Moneda", core.MONEDAS_VALIDAS, index=core.MONEDAS_VALIDAS.index(record["moneda"]) if record is not None else 0)
                    tc = c_tc.number_input("TC", min_value=0.0, value=float(record["tc"]) if record is not None else 1.0, disabled=(moneda == "ARS"))
                    importe = st.number_input("Importe", min_value=0.0, value=float(record["importe"]) if record is not None else 0.0, step=50000.0)
                    estado = st.selectbox("Estado", core.ESTADOS_VALIDOS, index=core.ESTADOS_VALIDOS.index(record["estado"]) if record is not None else 0)
                    obs = st.text_area("Notas", value=record["obs"] if record is not None else "")

                    b1, b2 = st.columns(2)
                    sub = b1.form_submit_button("💾 Guardar", type="primary", use_container_width=True)
                    can = b2.form_submit_button("✖️ Cancelar", use_container_width=True)

                    if sub and proveedor:
                        df_current = get_df()
                        if record is not None:
                            idx = df_current.index[df_current["id"] == editing_id][0]
                            df_current.loc[idx, ["tesoreria", "fecha", "proveedor", "proyecto", "concepto", "moneda", "importe", "tc", "estado", "obs"]] = [tesoreria_sel, pd.Timestamp(fecha_pago), proveedor, proyecto, concepto, moneda, importe, tc, estado, obs]
                            st.session_state.editing_id = None
                        else:
                            new_id = core.next_id(df_current)
                            new_row = pd.DataFrame([{"id": new_id, "tesoreria": tesoreria_sel, "fecha": pd.Timestamp(fecha_pago), "proveedor": proveedor, "proyecto": proyecto, "concepto": concepto, "moneda": moneda, "importe": importe, "tc": tc, "estado": estado, "obs": obs}])
                            df_current = pd.concat([df_current, new_row], ignore_index=True)
                        set_df(df_current, sync_cloud=True)
                        st.rerun()

                    if can:
                        st.session_state.editing_id = None
                        st.rerun()

            with col_table:
                disp = df_filtered[["id", "tesoreria", "fecha", "proveedor", "concepto", "importe_ars", "estado"]].rename(columns={"id": "ID", "tesoreria": "Sede", "fecha": "Fecha", "proveedor": "Proveedor", "concepto": "Operación", "importe_ars": "Importe ARS", "estado": "Estado"}).sort_values("Fecha")
                st.dataframe(disp.style.format({"Fecha": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "", "Importe ARS": fmt_ars}), use_container_width=True, height=360, hide_index=True)
                ids_disp = disp["ID"].tolist()
                if ids_disp:
                    c_sel, c_edit, c_del = st.columns([2, 1, 1])
                    s_id = c_sel.selectbox("Seleccionar ID", ids_disp, label_visibility="collapsed")
                    if c_edit.button("✏️ Editar", use_container_width=True):
                        st.session_state.editing_id = s_id
                        st.rerun()
                    if c_del.button("🗑️ Eliminar", use_container_width=True):
                        df_curr = get_df()
                        df_curr = df_curr[df_curr["id"] != s_id]
                        set_df(df_curr, sync_cloud=True)
                        st.rerun()

        with tab_export:
            st.subheader("⬇️ Exportar Base Actualizada")
            df_exp = get_df()
            xlsx_data = core.export_to_excel(df_exp)
            st.download_button("⬇️ Descargar Excel Completo", xlsx_data, f"base_pagos_efectivo_{dt.date.today().isoformat()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ===========================================================================
# MÓDULO 2: ANALIZADOR DE LIBRADORES · BCRA (ESPACIO DEDICADO)
# ===========================================================================
elif modulo_activo == "🔍 Analizador de Libradores · BCRA":
    import streamlit.components.v1 as components

    col_t_bcra, col_btns_bcra = st.columns([3, 1.2])
    with col_t_bcra:
        st.title("Analizador de Libradores BCRA")
        st.caption("Evaluación crediticia de CUITs en Central de Deudores y Cheques Rechazados en tiempo real.")

    with col_btns_bcra:
        st.write("")
        b_demo, b_clean = st.columns(2)
        if b_demo.button("Cargar demo", use_container_width=True):
            st.session_state["bcra_input_val"] = "30715316915\n20222645019\n30528011493"
            st.rerun()
        if b_clean.button("Limpiar", use_container_width=True):
            st.session_state["bcra_input_val"] = ""
            st.session_state["resultados_bcra"] = []
            st.rerun()

    input_text_val = st.session_state.get("bcra_input_val", "")

    c_box_in, c_box_btn = st.columns([3.5, 1])
    with c_box_in:
        cuits_raw = st.text_area(
            "CUITs a evaluar",
            value=input_text_val,
            placeholder="Pegá los CUITs (uno por línea o separados por espacio/coma)\nEj:\n30-71649553-8\n30-52801149-3",
            height=90,
            label_visibility="collapsed",
            key="area_cuits_bcra",
        )
    with c_box_btn:
        st.write("")
        st.write("")
        btn_consultar_bcra = st.button("🔎 Consultar BCRA", type="primary", use_container_width=True)

    if btn_consultar_bcra and cuits_raw.strip():
        lista_cuits = core.parse_cuits_input(cuits_raw)
        if not lista_cuits:
            st.warning("No se detectaron CUITs válidos de 11 dígitos.")
        else:
            prog = st.progress(0, text="Iniciando conexión con BCRA...")
            res_bcra = []
            for idx, c in enumerate(lista_cuits):
                prog.progress((idx + 1) / len(lista_cuits), text=f"Consultando CUIT {c} ({idx + 1}/{len(lista_cuits)})...")
                res_bcra.append(core.fetch_bcra_data(c))
            prog.empty()

            peso = {"bad": 0, "warn": 1, "ok": 2}
            res_bcra.sort(key=lambda x: peso.get(x["risk"], 3))
            st.session_state["resultados_bcra"] = res_bcra
            st.session_state["bcra_status_msg"] = f"✅ Consulta finalizada ({len(res_bcra)} libradores evaluados)."

    if "bcra_status_msg" in st.session_state:
        st.caption(st.session_state["bcra_status_msg"])

    data_bcra = st.session_state.get("resultados_bcra", [])
    n_tot = len(data_bcra)
    n_ok = len([x for x in data_bcra if x["risk"] == "ok"])
    n_warn = len([x for x in data_bcra if x["risk"] == "warn"])
    n_bad = len([x for x in data_bcra if x["risk"] == "bad"])

    # Tarjetas de Resumen con color explícito forzado
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"<div class='bcra-card'><div class='bcra-label'>Libradores Evaluados</div><div class='bcra-val' style='color:#0f172a;'>{n_tot}</div></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='bcra-card'><div class='bcra-label'>Sin Alertas</div><div class='bcra-val' style='color:#166534;'>{n_ok}</div></div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"<div class='bcra-card'><div class='bcra-label'>Revisar</div><div class='bcra-val' style='color:#854d0e;'>{n_warn}</div></div>", unsafe_allow_html=True)
    with c4:
        st.markdown(f"<div class='bcra-card'><div class='bcra-label'>Alertas / Rechazar</div><div class='bcra-val' style='color:#991b1b;'>{n_bad}</div></div>", unsafe_allow_html=True)

    st.write("")

    if data_bcra:
        col_grid_main, col_grid_drawer = st.columns([1.6, 1.1])

        with col_grid_main:
            st.markdown(
                """
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <b style="font-size:15px;">Listado de Libradores</b>
                    <small style="color:#64748b;">Prioridad: Alerta → Revisar → Sin Alertas</small>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # Generamos las filas HTML sin indentación que confunda al parser
            rows_html_list = []
            for x in data_bcra:
                sit_val = x["worst"]
                badge_bg = "#dcfce7" if sit_val in (0, 1) else ("#fef9c3" if sit_val == 2 else "#fee2e2")
                badge_tx = "#166534" if sit_val in (0, 1) else ("#854d0e" if sit_val == 2 else "#991b1b")
                
                pill_bg = "#fee2e2" if x["risk"] == "bad" else ("#fef9c3" if x["risk"] == "warn" else "#dcfce7")
                pill_tx = "#991b1b" if x["risk"] == "bad" else ("#854d0e" if x["risk"] == "warn" else "#166534")

                imp_text = f"<span style='color:#991b1b; font-weight:700;'>({x['pending']} impagos)</span>" if x["pending"] > 0 else "<span style='color:#64748b;'>(0 impagos)</span>"

                row = (
                    f"<tr style='border-bottom:1px solid #f1f5f9;'>"
                    f"<td style='padding:12px 10px;'><b>{x['denominacion']}</b><br><small style='font-family:monospace; color:#64748b;'>{x['cuit']}</small></td>"
                    f"<td style='text-align:center; padding:12px 6px;'><span style='display:inline-flex; align-items:center; justify-content:center; width:24px; height:24px; border-radius:50%; background:{badge_bg}; color:{badge_tx}; font-weight:700; font-size:12px;'>{sit_val}</span></td>"
                    f"<td style='padding:12px 10px; font-weight:700; color:#0f172a;'>{fmt_ars(x['debt'])}</td>"
                    f"<td style='padding:12px 10px; color:#0f172a;'>{x['rejected']} {imp_text}</td>"
                    f"<td style='padding:12px 10px; text-align:right;'><span style='display:inline-block; padding:3px 10px; border-radius:9999px; font-size:11px; font-weight:700; background:{pill_bg}; color:{pill_tx};'>{x['risk_label']}</span></td>"
                    f"</tr>"
                )
                rows_html_list.append(row)

            table_content = "".join(rows_html_list)
            full_table_html = (
                f"<div style='background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden; font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;'>"
                f"<table style='width:100%; border-collapse:collapse; font-size:13px;'>"
                f"<thead><tr style='border-bottom:1px solid #e2e8f0; color:#64748b; font-size:11px; text-transform:uppercase; background:#f8fafc;'>"
                f"<th style='padding:10px; text-align:left;'>Librador / Denominación</th>"
                f"<th style='padding:10px; text-align:center;'>Peor Sit.</th>"
                f"<th style='padding:10px; text-align:left;'>Deuda Bancaria</th>"
                f"<th style='padding:10px; text-align:left;'>Rechazos (Pend.)</th>"
                f"<th style='padding:10px; text-align:right;'>Estado</th>"
                f"</tr></thead>"
                f"<tbody>{table_content}</tbody>"
                f"</table></div>"
            )

            # Renderizado encapsulado garantizado
            calc_height = max(240, 48 * len(data_bcra) + 60)
            components.html(full_table_html, height=calc_height, scrolling=True)

        with col_grid_drawer:
            st.markdown(
                """
                <div style="margin-bottom:8px;">
                    <span class="bcra-label">Ficha del Librador</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            opciones_nombres = [f"{x['denominacion']} ({x['cuit']})" for x in data_bcra]
            sel_lib = st.selectbox("Seleccionar para inspeccionar:", opciones_nombres, label_visibility="collapsed", key="sel_drawer_lib")
            idx_sel = opciones_nombres.index(sel_lib)
            lib = data_bcra[idx_sel]

            pill_bg_d = "#fee2e2" if lib["risk"] == "bad" else ("#fef9c3" if lib["risk"] == "warn" else "#dcfce7")
            pill_tx_d = "#991b1b" if lib["risk"] == "bad" else ("#854d0e" if lib["risk"] == "warn" else "#166534")

            drawer_html = (
                f"<div style='background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:18px; font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;'>"
                f"<h3 style='margin:0 0 2px 0; font-size:17px; font-weight:700; color:#0f172a;'>{lib['denominacion']}</h3>"
                f"<div style='font-family:monospace; color:#64748b; font-size:13px; margin-bottom:12px;'>{lib['cuit']}</div>"
                f"<span style='display:inline-block; padding:3px 10px; border-radius:9999px; font-size:11px; font-weight:700; background:{pill_bg_d}; color:{pill_tx_d};'>{lib['risk_label']}</span>"
                f"<div style='display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-top:16px;'>"
                f"<div style='background:#f8fafc; padding:10px; border-radius:8px; border:1px solid #edf2f7;'><div style='font-size:10px; text-transform:uppercase; font-weight:600; color:#64748b;'>Peor Situación</div><b style='font-size:16px; color:#0f172a; margin-top:2px; display:block;'>{lib['worst']}</b></div>"
                f"<div style='background:#f8fafc; padding:10px; border-radius:8px; border:1px solid #edf2f7;'><div style='font-size:10px; text-transform:uppercase; font-weight:600; color:#64748b;'>Deuda Total</div><b style='font-size:16px; color:#0f172a; margin-top:2px; display:block;'>{fmt_ars(lib['debt'])}</b></div>"
                f"<div style='background:#f8fafc; padding:10px; border-radius:8px; border:1px solid #edf2f7;'><div style='font-size:10px; text-transform:uppercase; font-weight:600; color:#64748b;'>Total Rechazos</div><b style='font-size:16px; color:#0f172a; margin-top:2px; display:block;'>{lib['rejected']}</b></div>"
                f"<div style='background:#f8fafc; padding:10px; border-radius:8px; border:1px solid #edf2f7;'><div style='font-size:10px; text-transform:uppercase; font-weight:600; color:#64748b;'>Impagos / Pend.</div><b style='font-size:16px; margin-top:2px; display:block; color:{'#991b1b' if lib['pending'] > 0 else '#0f172a'};'>{lib['pending']}</b></div>"
                f"</div></div>"
            )
            components.html(drawer_html, height=210)

            # Alertas y diagnósticos
            for al in lib["alerts"]:
                bg_al = "#fee2e2" if lib["risk"] == "bad" else ("#fef9c3" if lib["risk"] == "warn" else "#dcfce7")
                tx_al = "#991b1b" if lib["risk"] == "bad" else ("#854d0e" if lib["risk"] == "warn" else "#166534")
                st.markdown(f"<div style='background:{bg_al}; color:{tx_al}; padding:8px 12px; border-radius:6px; font-size:12px; font-weight:500; margin-bottom:6px;'>{al}</div>", unsafe_allow_html=True)

            if lib["last3"]:
                st.markdown("<small style='font-weight:700; text-transform:uppercase; color:#64748b;'>Últimos 3 Cheques Rechazados</small>", unsafe_allow_html=True)
                for ch in lib["last3"]:
                    f_ch = str(ch.get("fechaRechazo", "—")).split("T")[0]
                    if "-" in f_ch:
                        p_f = f_ch.split("-")
                        if len(p_f) == 3 and len(p_f[0]) == 4:
                            f_ch = f"{p_f[2]}-{p_f[1]}-{p_f[0]}"
                    m_ch = fmt_ars(ch.get("monto", 0))
                    estado_ch = "<span style='color:#166534; font-weight:700;'>Pagado</span>" if ch.get("fechaPago") else "<span style='color:#991b1b; font-weight:700;'>Impago</span>"
                    st.markdown(f"<div style='font-size:12px; border-bottom:1px solid #f1f5f9; padding:4px 0;'>• <b>{f_ch}</b> — {m_ch} ({ch.get('causal')}) [{estado_ch}]</div>", unsafe_allow_html=True)

        # Módulo WhatsApp
        st.divider()
        st.subheader("📱 Mensaje de sugerencia para WhatsApp")
        rechazados = [x for x in data_bcra if x["risk"] == "bad"]
        revisar = [x for x in data_bcra if x["risk"] == "warn"]

        if rechazados:
            wa_txt = "Hola! Te paso el resultado de la tanda analizada:\n\n"
            wa_txt += "❌ *RECHAZAR los siguientes libradores:*\n"
            for r in rechazados:
                motivo = f"Sit: {r['worst']}"
                if r['pending'] > 0:
                    motivo += f" | {r['pending']} cheques impagos"
                if r['rejected'] > 5:
                    motivo += f" | {r['rejected']} rechazos totales"
                wa_txt += f"• *{r['denominacion']}* (CUIT {r['cuit']}) - {motivo}\n"
            if revisar:
                wa_txt += "\n⚠️ *A REVISAR antes de recibir:*\n"
                for rv in revisar:
                    wa_txt += f"• {rv['denominacion']} (Sit: {rv['worst']})\n"
        elif revisar:
            wa_txt = "Hola! De la tanda analizada no hay alertas críticas, pero sugiero *revisar*:\n\n"
            for rv in revisar:
                wa_txt += f"• *{rv['denominacion']}* (CUIT {rv['cuit']}) - Sit: {rv['worst']}\n"
        else:
            wa_txt = "Hola! Todos los libradores analizados están en condiciones *OK (Sin Alertas)*. ✅"

        st.text_area("Copia el texto directamente para enviar:", value=wa_txt, height=130)
