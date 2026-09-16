# -*- coding: utf-8 -*-
"""
Sistema de Gestión de Efectivo — Multisede (Tucumán / Buenos Aires)
====================================================================
Versión Optimizada: Detalle de Escaleras con Descarga Excel, Conciliador Rápido
y Carga Dinámica de Escaleras No Lineales.
"""

from __future__ import annotations

import datetime as dt
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import core

# ---------------------------------------------------------------------------
# Configuración general de la página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Gestión de Efectivo — Multisede",
    page_icon="💵",
    layout="wide",
    initial_sidebar_state="expanded",
)

SEED_PATH = Path(__file__).parent / "data_seed.csv"

# ---------------------------------------------------------------------------
# Estado de la aplicación
# ---------------------------------------------------------------------------

def init_state() -> None:
    if "df" not in st.session_state:
        if SEED_PATH.exists():
            st.session_state.df = core.load_seed_csv(SEED_PATH)
        else:
            st.session_state.df = core.normalize_df(pd.DataFrame(columns=core.COLUMNS))
    if "editing_id" not in st.session_state:
        st.session_state.editing_id = None
    if "sede_global" not in st.session_state:
        st.session_state.sede_global = core.SEDE_CONSOLIDADO

init_state()

def get_df() -> pd.DataFrame:
    return st.session_state.df

def set_df(df: pd.DataFrame) -> None:
    st.session_state.df = core.normalize_df(df)

def fmt_ars(value: float) -> str:
    try:
        return f"$ {value:,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "$ 0"

# ---------------------------------------------------------------------------
# Barra lateral
# ---------------------------------------------------------------------------

st.sidebar.title("💵 Gestión de Efectivo")
st.sidebar.caption("Control Operativo y Financiero Multisede")

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

st.sidebar.divider()

with st.sidebar.expander("📁 Cargar / Importar datos", expanded=False):
    st.markdown("**Reemplazar base central (DB_Pagos)**")
    uploaded = st.file_uploader("Subir archivo Excel", type=["xlsx"], key="uploader_replace")
    if uploaded is not None and st.button("Reemplazar base actual", type="primary"):
        try:
            new_df = core.load_from_uploaded_xlsx(uploaded)
            set_df(new_df)
            st.success(f"Cargados {len(new_df)} pagos.")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

    st.markdown("---")
    if st.button("↺ Restaurar datos originales (semilla)"):
        set_df(core.load_seed_csv(SEED_PATH))
        st.rerun()

st.sidebar.divider()
st.sidebar.subheader("🔎 Filtros Globales")

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

st.sidebar.divider()
st.sidebar.caption(f"Sede: **{sede_global}** · Pagos: **{len(df_filtered)}** / {len(df_all_raw)}")

# ---------------------------------------------------------------------------
# Pestañas Principales
# ---------------------------------------------------------------------------

st.title("💵 Sistema de Gestión de Efectivo")
st.caption(f"Unidad activa: **{sede_global}** · Flujo de caja, escaleras de pagos y conciliación operativa.")

tab_kpi, tab_matriz, tab_escaleras, tab_conciliar, tab_asistente, tab_abm, tab_export = st.tabs([
    "📊 Panel Ejecutivo",
    "🗓️ Matriz Semanal",
    "🪜 Detalle de Escaleras",
    "⚡ Conciliación Rápida",
    "➕ Cargar Nueva Escalera",
    "📝 Registro Manual (ABM)",
    "⬇️ Exportar",
])

# ---------------------------------------------------------------------------
# TAB 1 — Panel Ejecutivo
# ---------------------------------------------------------------------------
with tab_kpi:
    kpis = core.compute_kpis(df_filtered)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Comprometido", fmt_ars(kpis.total_comprometido))
    c2.metric("Total Pagado", fmt_ars(kpis.total_pagado))
    c3.metric("Saldo Pendiente", fmt_ars(kpis.saldo_pendiente))
    c4.metric("Desembolso Promedio", fmt_ars(kpis.desembolso_promedio))
    c5.metric("TC Promedio (USD)", f"{kpis.tc_promedio:,.2f}".replace(",", "."))

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
        st.subheader("🚨 Alertas de Picos Semanales")
        st.caption(f"Semanas con retiros superiores a {fmt_ars(core.UMBRAL_PICO_EFECTIVO)}")
        picos = core.weeks_over_threshold(df_filtered)
        if picos.empty:
            st.success("No hay semanas que superen el tope logístico.")
        else:
            for _, r in picos.iterrows():
                st.error(f"**Semana {r['semana_etiqueta']}** — {fmt_ars(r['total_ars'])}", icon="⚠️")

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

# ---------------------------------------------------------------------------
# TAB 2 — Matriz Semanal Consolidada (CON FILTRO DE FECHAS Y DESCARGA EXCEL)
# ---------------------------------------------------------------------------
with tab_matriz:
    st.subheader("Matriz Semanal de Flujo de Efectivo (Consolidada)")
    st.caption("Visión consolidada por proveedor y semana. Filtrá semanas pasadas y descargá el reporte en Excel listo para enviar.")

    if df_filtered.empty:
        st.info("Sin datos para mostrar con los filtros aplicados.")
    else:
        # Obtener todas las semanas disponibles ordenadas
        df_sorted_weeks = df_filtered.sort_values("lunes_semana")
        semanas_ordenadas = (
            df_sorted_weeks[["lunes_semana", "semana_etiqueta"]]
            .drop_duplicates()
            .dropna()
        )

        col_filtro_sem, col_btn_descarga = st.columns([2, 2])

        with col_filtro_sem:
            opciones_semanas = semanas_ordenadas["semana_etiqueta"].tolist()
            # Selector de semana de inicio
            semana_inicio_sel = st.selectbox(
                "📅 Mostrar semanas desde:",
                options=opciones_semanas,
                index=0,
                help="Ocultá semanas pasadas para enfocarte en las obligaciones vigentes y futuras.",
            )

        # Filtrar el DataFrame a partir del lunes de la semana seleccionada
        lunes_corte = semanas_ordenadas.loc[
            semanas_ordenadas["semana_etiqueta"] == semana_inicio_sel, "lunes_semana"
        ].iloc[0]

        df_matriz_periodo = df_filtered[df_filtered["lunes_semana"] >= lunes_corte]
        matrix = core.build_weekly_matrix(df_matriz_periodo)

        if matrix.empty:
            st.info("No hay pagos para el período seleccionado.")
        else:
            # Función para exportar la matriz a un Excel .xlsx formateado profesionalmente
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
                    is_peak = (not is_total_row and isinstance(total_semanal, (int, float)) and total_semanal > core.UMBRAL_PICO_EFECTIVO)

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

            with col_btn_descarga:
                st.write("")  # Espaciado visual
                st.write("")
                excel_matriz = export_matrix_to_excel(matrix)
                st.download_button(
                    label="📥 Descargar Matriz en Excel (.xlsx)",
                    data=excel_matriz,
                    file_name=f"Matriz_Semanal_{sede_global}_{dt.date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="dl_matriz_excel",
                )

            st.markdown("---")

            # Renderizado de la tabla en pantalla
            provider_cols = [c for c in matrix.columns if c not in ("Semana", "TOTAL SEMANAL", "ACUMULADO")]
            def highlight_peaks(row):
                styles = [""] * len(row)
                if row["Semana"] == "TOTAL POR PROVEEDOR":
                    return ["font-weight: bold; background-color: #EFEFEF"] * len(row)
                try:
                    if row["TOTAL SEMANAL"] > core.UMBRAL_PICO_EFECTIVO:
                        idx = list(row.index).index("TOTAL SEMANAL")
                        styles[idx] = "background-color:#FFC7CE; font-weight:bold;"
                except Exception:
                    pass
                return styles

            styled = matrix.style.apply(highlight_peaks, axis=1).format(
                {c: fmt_ars for c in provider_cols + ["TOTAL SEMANAL", "ACUMULADO"]}
            )
            st.dataframe(styled, use_container_width=True, height=520, hide_index=True)

# ---------------------------------------------------------------------------
# TAB 3 — Control Detallado de Escaleras (CON BOTÓN DE EXCEL NATIVO)
# ---------------------------------------------------------------------------
with tab_escaleras:
    st.subheader("🪜 Control Individual de Escaleras y Presupuestos")
    st.caption("Inspeccioná cada contrato y cuota de forma vertical. Podés descargar el archivo Excel limpio para compartir.")

    prov_list = sorted(df_filtered["proveedor"].dropna().unique().tolist())
    if not prov_list:
        st.info("No hay proveedores para los filtros seleccionados.")
    else:
        col_sel_p, col_info_p = st.columns([1, 2])
        with col_sel_p:
            p_sel = st.selectbox("Seleccionar Proveedor a Inspeccionar", prov_list, key="p_inspect")

        df_p = df_filtered[df_filtered["proveedor"] == p_sel].copy()
        kpis_p = core.compute_kpis(df_p)

        with col_info_p:
            st.markdown(f"#### Ficha: **{p_sel}**")
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
        df_p["concepto_clean"] = df_p["concepto"].fillna("").str.strip()
        conceptos = sorted([c for c in df_p["concepto_clean"].unique().tolist() if c])
        if not conceptos:
            conceptos = ["General / Único Presupuesto"]
            df_p["concepto_clean"] = "General / Único Presupuesto"

        def format_ladder_table(sub_df: pd.DataFrame) -> pd.DataFrame:
            t = sub_df.sort_values("fecha").copy()
            res = pd.DataFrame({
                "Fecha Vto.": t["fecha"],
                "Presupuesto / Observaciones": t["presupuesto_obs"],
                "Importe Pactado": t["importe"],
                "TC": t["tc"],
                "Importe ARS": t["importe_ars"],
                "Estado": t["estado"],
                "Moneda": t["moneda"],
            })
            return res

        def export_provider_to_excel(df_table: pd.DataFrame, prov_name: str) -> bytes:
            wb = Workbook()
            ws = wb.active
            ws.title = "Escalera de Pagos"

            ws["A1"] = f"ESCALERA DE PAGOS — {prov_name.upper()}"
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

        def render_ladder_ui(sub_df: pd.DataFrame, label: str):
            table_view = format_ladder_table(sub_df)
            excel_bytes = export_provider_to_excel(table_view, p_sel)

            c_down, _ = st.columns([2, 3])
            with c_down:
                st.download_button(
                    label="📥 Descargar Escalera en Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Escalera_{p_sel.replace(' ', '_')}_{dt.date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{label}_{p_sel}",
                    type="primary",
                )

            st.dataframe(
                table_view.style.format({
                    "Fecha Vto.": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                    "Importe Pactado": "{:,.2f}", "TC": "{:,.2f}", "Importe ARS": fmt_ars,
                }),
                use_container_width=True,
                hide_index=True,
            )

        if len(conceptos) > 1:
            st.markdown(f"##### Se detectaron **{len(conceptos)} escaleras/presupuestos** en simultáneo:")
            esc_tabs = st.tabs([f"📌 {c}" for c in conceptos] + ["📑 Todas las Escaleras Consolidadas"])
            for idx, c_name in enumerate(conceptos):
                with esc_tabs[idx]:
                    sub_esc = df_p[df_p["concepto_clean"] == c_name]
                    sub_kpis = core.compute_kpis(sub_esc)
                    e1, e2, e3 = st.columns(3)
                    e1.write(f"**Total Contrato:** {fmt_ars(sub_kpis.total_comprometido)}")
                    e2.write(f"**Pagado:** {fmt_ars(sub_kpis.total_pagado)}")
                    e3.write(f"**Saldo:** {fmt_ars(sub_kpis.saldo_pendiente)}")
                    render_ladder_ui(sub_esc, f"tab_{idx}")

            with esc_tabs[-1]:
                render_ladder_ui(df_p, "tab_all")
        else:
            render_ladder_ui(df_p, "tab_single")

# ---------------------------------------------------------------------------
# TAB 4 — Conciliador Rápido (De Pendiente a Pagado)
# ---------------------------------------------------------------------------
with tab_conciliar:
    st.subheader("⚡ Conciliador Rápido de Pagos")
    st.caption("Cambiá el estado de los desembolsos de 'Pendiente' a 'Pagado' en un solo paso.")

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
                "Conciliar (Marcar Pagado)": st.column_config.CheckboxColumn(
                    "¿Pagar?",
                    help="Tildá los pagos liquidados",
                    default=False,
                ),
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
                set_df(df_updated)
                st.success(f"¡Se actualizaron {len(pagos_a_conciliar)} pagos a estado PAGADO!")
                st.rerun()

        with c_resumen:
            if pagos_a_conciliar:
                suma_sel = edited[edited["Conciliar (Marcar Pagado)"] == True]["importe_ars"].sum()
                st.info(f"Seleccionados **{len(pagos_a_conciliar)} pagos** por un total de **{fmt_ars(suma_sel)}**.")

# ---------------------------------------------------------------------------
# TAB 5 — Cargar Nueva Escalera (INTERACTIVA, NO LINEAL Y SIN ST.FORM)
# ---------------------------------------------------------------------------
with tab_asistente:
    st.subheader("➕ Cargar Nueva Escalera de Pago")
    st.caption("Configurá la cabecera y completá las fechas y montos específicos de cada tramo en la grilla.")

    c_cab1, c_cab2, c_cab3 = st.columns(3)
    with c_cab1:
        esc_tesoreria = st.selectbox(
            "Tesorería / Sede", core.TESORERIAS_VALIDAS,
            index=0 if sede_global == "Tucumán" else 1, key="esc_sede"
        )
        tipo_prov = st.radio("Proveedor", ["Existente", "Nuevo"], horizontal=True, key="esc_tipo_prov")
        prov_opts = sorted(get_df()["proveedor"].dropna().unique().tolist())
        if tipo_prov == "Existente" and prov_opts:
            esc_proveedor = st.selectbox("Seleccionar de la lista", prov_opts, key="esc_prov_sel")
        else:
            esc_proveedor = st.text_input("Escribir nombre del nuevo proveedor", key="esc_prov_new").strip().upper()

    with c_cab2:
        esc_proyecto = st.text_input(
            "Proyecto / Obra",
            value="L2 TUC" if esc_tesoreria == "Tucumán" else "L2 BA",
            key="esc_proy"
        )
        esc_concepto = st.text_input(
            "Presupuesto / Observaciones",
            placeholder="Ej: Carpintería Aluminio - Presupuesto N° 450",
            key="esc_conc"
        )

    with c_cab3:
        esc_moneda = st.selectbox("Moneda del Acuerdo", core.MONEDAS_VALIDAS, key="esc_mon")
        if esc_moneda == "USD":
            esc_tc = st.number_input("Tipo de Cambio acordado (USD)", min_value=1.0, value=1400.0, step=10.0, key="esc_tc_usd")
        else:
            esc_tc = 1.0
            st.text_input("Tipo de Cambio", value="1.00 (ARS)", disabled=True)

    st.markdown("---")
    st.markdown("##### Cronograma de Pagos (Fechas e Importes Personalizados)")

    col_cant, _ = st.columns([1, 3])
    with col_cant:
        cant_cuotas = st.number_input("Cantidad de tramos / cuotas", min_value=1, max_value=50, value=4, step=1, key="esc_n_cuotas")

    if "grid_ladder_data" not in st.session_state or len(st.session_state.grid_ladder_data) != cant_cuotas:
        base_f = dt.date.today()
        st.session_state.grid_ladder_data = pd.DataFrame([
            {"Tramo": f"Cuota {i + 1}", "Fecha Vto.": base_f + dt.timedelta(days=7 * i), "Importe": 0.0}
            for i in range(int(cant_cuotas))
        ])

    grid_edited = st.data_editor(
        st.session_state.grid_ladder_data,
        column_config={
            "Tramo": st.column_config.TextColumn("Tramo", disabled=True),
            "Fecha Vto.": st.column_config.DateColumn("Fecha Vto.", format="DD/MM/YYYY", required=True),
            "Importe": st.column_config.NumberColumn(
                f"Importe en {esc_moneda}",
                help="Ingresá el monto correspondiente a esta cuota",
                min_value=0.0,
                format="%.2f",
                required=True,
            ),
        },
        use_container_width=True,
        hide_index=True,
        key="grid_cuotas_editor",
    )

    total_acordado = float(grid_edited["Importe"].sum())
    total_ars_calc = total_acordado * esc_tc if esc_moneda == "USD" else total_acordado

    st.info(
        f"📊 **Resumen del Acuerdo:** {cant_cuotas} tramos · Total: **{esc_moneda} {total_acordado:,.2f}** "
        f"{f'(Equivalente: **{fmt_ars(total_ars_calc)}** al TC {esc_tc:,.2f})' if esc_moneda == 'USD' else ''}"
    )

    if st.button("🚀 Guardar Escalera Completa en la Base", type="primary", use_container_width=True):
        if not esc_proveedor:
            st.error("Debés indicar el nombre del Proveedor.")
        elif not esc_concepto:
            st.error("Debés completar el campo Presupuesto / Observaciones.")
        elif total_acordado <= 0:
            st.error("El total de los importes debe ser mayor a 0.")
        else:
            items_ladder = []
            for _, r in grid_edited.iterrows():
                items_ladder.append({
                    "fecha": pd.to_datetime(r["Fecha Vto."]),
                    "importe": float(r["Importe"]),
                })

            df_curr = get_df()
            new_rows = core.generate_custom_ladder(
                df_existing=df_curr,
                tesoreria=esc_tesoreria,
                proveedor=esc_proveedor,
                proyecto=esc_proyecto,
                concepto=esc_concepto,
                moneda=esc_moneda,
                tc=esc_tc,
                items=items_ladder,
            )
            df_total = pd.concat([df_curr[core.COLUMNS], new_rows[core.COLUMNS]], ignore_index=True)
            set_df(df_total)
            st.success(f"¡Se incorporaron exitosamente los {len(new_rows)} pagos de la escalera para {esc_proveedor}!")
            st.rerun()

# ---------------------------------------------------------------------------
# TAB 6 — Registro Manual (ABM Individual)
# ---------------------------------------------------------------------------
with tab_abm:
    st.subheader("Registro y edición manual de pagos individuales")
    col_form, col_table = st.columns([1, 2])

    with col_form:
        editing_id = st.session_state.editing_id
        df_current = get_df()
        record = None
        if editing_id is not None:
            match = df_current[df_current["id"] == editing_id]
            if not match.empty:
                record = match.iloc[0]

        st.markdown(f"##### {'✏️ Editar pago ' + editing_id if record is not None else '➕ Nuevo pago individual'}")
        proveedores_existentes = sorted(df_current["proveedor"].dropna().unique().tolist())

        with st.form("form_pago_manual", clear_on_submit=False):
            tesoreria_sel = st.selectbox(
                "Tesorería / Sede", core.TESORERIAS_VALIDAS,
                index=core.TESORERIAS_VALIDAS.index(record["tesoreria"]) if record is not None else (
                    core.TESORERIAS_VALIDAS.index(sede_global) if sede_global in core.TESORERIAS_VALIDAS else 0
                ),
            )
            fecha_pago = st.date_input(
                "Fecha de Pago",
                value=record["fecha"].date() if record is not None and pd.notna(record["fecha"]) else dt.date.today(),
            )
            lunes_preview = core.monday_of(pd.Timestamp(fecha_pago))
            st.caption(f"Semana calculada: **{core.week_label(lunes_preview)}**")

            modo_p = st.radio("Proveedor", ["Seleccionar existente", "Agregar nuevo"], horizontal=True)
            if modo_p == "Seleccionar existente" and proveedores_existentes:
                default_idx = (
                    proveedores_existentes.index(record["proveedor"])
                    if record is not None and record["proveedor"] in proveedores_existentes else 0
                )
                proveedor = st.selectbox("Proveedor", proveedores_existentes, index=default_idx)
            else:
                proveedor = st.text_input(
                    "Nombre nuevo", value=record["proveedor"] if record is not None else ""
                ).strip().upper()

            proyecto = st.text_input("Proyecto", value=record["proyecto"] if record is not None else "L2")
            concepto = st.text_input("Presupuesto / Observaciones", value=record["concepto"] if record is not None else "")

            c_mon, c_tc = st.columns(2)
            moneda = c_mon.selectbox("Moneda", core.MONEDAS_VALIDAS, index=core.MONEDAS_VALIDAS.index(record["moneda"]) if record is not None else 0)
            tc = c_tc.number_input("TC", min_value=0.0, value=float(record["tc"]) if record is not None else 1.0, disabled=(moneda == "ARS"))

            importe = st.number_input("Importe", min_value=0.0, value=float(record["importe"]) if record is not None else 0.0, step=50000.0)
            estado = st.selectbox("Estado", core.ESTADOS_VALIDOS, index=core.ESTADOS_VALIDOS.index(record["estado"]) if record is not None else 0)
            obs = st.text_area("Notas internas", value=record["obs"] if record is not None else "")

            b1, b2 = st.columns(2)
            sub = b1.form_submit_button("💾 Guardar", type="primary", use_container_width=True)
            can = b2.form_submit_button("✖️ Cancelar", use_container_width=True)

            if sub:
                if not proveedor:
                    st.error("Indicá un proveedor.")
                else:
                    df_current = get_df()
                    if record is not None:
                        idx = df_current.index[df_current["id"] == editing_id][0]
                        df_current.loc[idx, [
                            "tesoreria", "fecha", "proveedor", "proyecto", "concepto",
                            "moneda", "importe", "tc", "estado", "obs",
                        ]] = [
                            tesoreria_sel, pd.Timestamp(fecha_pago), proveedor, proyecto, concepto,
                            moneda, importe, tc, estado, obs,
                        ]
                        st.session_state.editing_id = None
                    else:
                        new_id = core.next_id(df_current)
                        new_row = pd.DataFrame([{
                            "id": new_id, "tesoreria": tesoreria_sel, "fecha": pd.Timestamp(fecha_pago),
                            "proveedor": proveedor, "proyecto": proyecto, "concepto": concepto,
                            "moneda": moneda, "importe": importe, "tc": tc, "estado": estado, "obs": obs,
                        }])
                        df_current = pd.concat([df_current, new_row], ignore_index=True)
                    set_df(df_current)
                    st.success("Guardado correctamente.")
                    st.rerun()

            if can:
                st.session_state.editing_id = None
                st.rerun()

    with col_table:
        st.markdown("##### Listado de Pagos")
        disp = df_filtered[["id", "tesoreria", "fecha", "proveedor", "concepto", "importe_ars", "estado"]].rename(columns={
            "id": "ID", "tesoreria": "Sede", "fecha": "Fecha", "proveedor": "Proveedor",
            "concepto": "Presupuesto / Observaciones", "importe_ars": "Importe ARS", "estado": "Estado",
        }).sort_values("Fecha")

        st.dataframe(disp.style.format({
            "Fecha": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
            "Importe ARS": fmt_ars,
        }), use_container_width=True, height=360, hide_index=True)

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
                set_df(df_curr)
                st.warning(f"Pago {s_id} eliminado.")
                st.rerun()

# ---------------------------------------------------------------------------
# TAB 7 — Exportación Completa
# ---------------------------------------------------------------------------
with tab_export:
    st.subheader("⬇️ Exportar Base Actualizada")
    st.caption("Descargá la base completa con todas las sedes y escaleras consolidadas en formato Excel.")
    df_exp = get_df()
    xlsx_data = core.export_to_excel(df_exp)
    st.download_button(
        label="⬇️ Descargar Excel (base_pagos_efectivo_multisede.xlsx)",
        data=xlsx_data,
        file_name=f"base_pagos_efectivo_{dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
