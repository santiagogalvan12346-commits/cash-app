"""
Sistema de Gestión de Efectivo — L2 Tucumán
============================================
Aplicación Streamlit para el control operativo, registro (ABM) y
proyección del flujo de pagos en efectivo a contratistas y
proveedores de obra.

Ejecutar con:
    streamlit run app.py
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import core

# ---------------------------------------------------------------------------
# Configuración general de la página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Gestión de Efectivo — L2 Tucumán",
    page_icon="💵",
    layout="wide",
    initial_sidebar_state="expanded",
)

SEED_PATH = Path(__file__).parent / "data_seed.csv"


# ---------------------------------------------------------------------------
# Estado de la aplicación (persistencia en memoria de sesión)
# ---------------------------------------------------------------------------

def init_state() -> None:
    if "df" not in st.session_state:
        if SEED_PATH.exists():
            st.session_state.df = core.load_seed_csv(SEED_PATH)
        else:
            st.session_state.df = core.normalize_df(pd.DataFrame(columns=core.COLUMNS))
    if "editing_id" not in st.session_state:
        st.session_state.editing_id = None


init_state()


def get_df() -> pd.DataFrame:
    return st.session_state.df


def set_df(df: pd.DataFrame) -> None:
    st.session_state.df = core.normalize_df(df)


def fmt_ars(value: float) -> str:
    return f"$ {value:,.0f}".replace(",", ".")


# ---------------------------------------------------------------------------
# Sidebar — carga de archivo, filtros y navegación
# ---------------------------------------------------------------------------

st.sidebar.title("💵 Gestión de Efectivo")
st.sidebar.caption("Proyecto L2 Tucumán")

with st.sidebar.expander("📁 Cargar / reemplazar datos", expanded=False):
    uploaded = st.file_uploader(
        "Subir archivo Excel (hoja DB_Pagos)",
        type=["xlsx"],
        help="Debe contener una hoja 'DB_Pagos' con el mismo esquema del "
        "archivo SISTEMA_GESTION_EFECTIVO_L2_TUC.xlsx",
    )
    if uploaded is not None:
        if st.button("Reemplazar base actual con este archivo", type="primary"):
            try:
                new_df = core.load_from_uploaded_xlsx(uploaded)
                set_df(new_df)
                st.success(f"Base reemplazada: {len(new_df)} pagos cargados.")
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")

    if st.button("↺ Restaurar datos originales (semilla)"):
        set_df(core.load_seed_csv(SEED_PATH))
        st.rerun()

st.sidebar.divider()
st.sidebar.subheader("🔎 Filtros")

df_all = get_df()

proveedores_disponibles = sorted(df_all["proveedor"].dropna().unique().tolist())
estados_disponibles = core.ESTADOS_VALIDOS

f_proveedores = st.sidebar.multiselect(
    "Proveedor", proveedores_disponibles, default=proveedores_disponibles
)
f_estados = st.sidebar.multiselect(
    "Estado", estados_disponibles, default=estados_disponibles
)

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
st.sidebar.caption(
    f"Mostrando **{len(df_filtered)}** de **{len(df_all)}** pagos registrados."
)

# ---------------------------------------------------------------------------
# Encabezado
# ---------------------------------------------------------------------------

st.title("💵 Sistema de Gestión de Efectivo — L2 Tucumán")
st.caption(
    "Control operativo, registro y proyección del flujo de pagos en efectivo "
    "a contratistas y proveedores clave de obra/servicios."
)

tab_kpi, tab_matriz, tab_abm, tab_proveedores, tab_export = st.tabs(
    [
        "📊 Panel Ejecutivo",
        "🗓️ Matriz Semanal",
        "📝 Registro de Pagos (ABM)",
        "🏗️ Control por Proveedor",
        "⬇️ Exportar",
    ]
)

# ---------------------------------------------------------------------------
# TAB 1 — Panel Ejecutivo (KPIs + Alertas)
# ---------------------------------------------------------------------------

with tab_kpi:
    kpis = core.compute_kpis(df_filtered)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total comprometido", fmt_ars(kpis.total_comprometido))
    c2.metric("Total pagado", fmt_ars(kpis.total_pagado))
    c3.metric("Saldo pendiente", fmt_ars(kpis.saldo_pendiente))
    c4.metric(
        "Desembolso promedio",
        fmt_ars(kpis.desembolso_promedio),
        help=f"{kpis.cantidad_desembolsos} desembolsos totales, {kpis.cantidad_pagados} pagados",
    )

    st.divider()

    left, right = st.columns([3, 2])

    with left:
        st.subheader("Distribución mensual proyectada")
        dist = core.monthly_distribution(df_filtered)
        if not dist.empty:
            fig = px.bar(
                dist,
                x="mes",
                y="importe_ars",
                text_auto=".2s",
                labels={"mes": "Mes", "importe_ars": "Importe ARS"},
                color_discrete_sequence=["#1F4E78"],
            )
            fig.update_layout(
                yaxis_title="Importe ARS",
                xaxis_title="",
                margin=dict(t=10, b=10),
                height=360,
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos para los filtros seleccionados.")

    with right:
        st.subheader("🚨 Alertas de picos de efectivo")
        st.caption(
            f"Semanas con salida de billetes superior a {fmt_ars(core.UMBRAL_PICO_EFECTIVO)} "
            "(planificación logística / retiro bancario)."
        )
        picos = core.weeks_over_threshold(df_filtered)
        if picos.empty:
            st.success("No se detectaron picos por encima del umbral.")
        else:
            for _, r in picos.iterrows():
                st.error(
                    f"**Semana {r['semana_etiqueta']}** — {fmt_ars(r['total_ars'])}",
                    icon="⚠️",
                )

    st.divider()
    st.subheader("Progreso por proveedor (% pagado vs. comprometido)")
    ctrl = core.provider_control(df_filtered)
    if not ctrl.empty:
        fig2 = go.Figure()
        fig2.add_trace(
            go.Bar(
                y=ctrl["proveedor"],
                x=ctrl["compromiso_total_ars"],
                orientation="h",
                name="Comprometido",
                marker_color="#D9D9D9",
            )
        )
        fig2.add_trace(
            go.Bar(
                y=ctrl["proveedor"],
                x=ctrl["pagado_ars"],
                orientation="h",
                name="Pagado",
                marker_color="#1F4E78",
            )
        )
        fig2.update_layout(
            barmode="overlay",
            xaxis_title="Importe ARS",
            height=320,
            margin=dict(t=10, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

# ---------------------------------------------------------------------------
# TAB 2 — Matriz de Tesorería Semanal
# ---------------------------------------------------------------------------

with tab_matriz:
    st.subheader("Matriz de Tesorería Semanal (Cash Flow)")
    st.caption(
        "Filas: semanas cronológicas (lunes a viernes) · Columnas: proveedores · "
        "Totalizadores automáticos por fila y columna."
    )

    matrix = core.build_weekly_matrix(df_filtered)
    if matrix.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        provider_cols = [c for c in matrix.columns if c not in ("Semana", "TOTAL SEMANAL")]

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
            {c: fmt_ars for c in provider_cols + ["TOTAL SEMANAL"]}
        )
        st.dataframe(styled, use_container_width=True, height=520)

        st.caption(
            "🔴 Las celdas resaltadas en la columna TOTAL SEMANAL indican semanas que "
            f"superan el umbral de {fmt_ars(core.UMBRAL_PICO_EFECTIVO)}."
        )

# ---------------------------------------------------------------------------
# TAB 3 — Registro / ABM de Pagos
# ---------------------------------------------------------------------------

with tab_abm:
    st.subheader("Registro y edición de pagos")

    col_form, col_table = st.columns([1, 2])

    with col_form:
        editing_id = st.session_state.editing_id
        df_current = get_df()
        record = None
        if editing_id is not None:
            match = df_current[df_current["id"] == editing_id]
            if not match.empty:
                record = match.iloc[0]

        st.markdown(f"##### {'✏️ Editar pago ' + editing_id if record is not None else '➕ Nuevo pago'}")

        proveedores_existentes = sorted(df_current["proveedor"].dropna().unique().tolist())

        with st.form("form_pago", clear_on_submit=False):
            fecha_pago = st.date_input(
                "Fecha de Pago",
                value=record["fecha"].date() if record is not None and pd.notna(record["fecha"]) else dt.date.today(),
            )

            lunes_preview = core.monday_of(pd.Timestamp(fecha_pago))
            st.caption(
                f"📅 Lunes de semana: **{lunes_preview.strftime('%d/%m/%Y')}** · "
                f"Semana: **{core.week_label(lunes_preview)}**"
            )

            modo_proveedor = st.radio(
                "Proveedor", ["Seleccionar existente", "Agregar nuevo"], horizontal=True
            )
            if modo_proveedor == "Seleccionar existente" and proveedores_existentes:
                default_idx = (
                    proveedores_existentes.index(record["proveedor"])
                    if record is not None and record["proveedor"] in proveedores_existentes
                    else 0
                )
                proveedor = st.selectbox("Seleccione proveedor", proveedores_existentes, index=default_idx)
            else:
                proveedor = st.text_input(
                    "Nombre del nuevo proveedor",
                    value=record["proveedor"] if record is not None else "",
                ).strip().upper()

            proyecto = st.text_input(
                "Proyecto / Destino",
                value=record["proyecto"] if record is not None else core.PROYECTO_DEFAULT,
            )

            c_mon, c_tc = st.columns(2)
            moneda = c_mon.selectbox(
                "Moneda",
                core.MONEDAS_VALIDAS,
                index=core.MONEDAS_VALIDAS.index(record["moneda"]) if record is not None else 0,
            )
            tc = c_tc.number_input(
                "Tipo de Cambio",
                min_value=0.0,
                value=float(record["tc"]) if record is not None else 1.0,
                step=0.5,
                disabled=(moneda == "ARS"),
                help="Solo aplica si Moneda = USD.",
            )

            importe = st.number_input(
                "Importe pactado (en la moneda seleccionada)",
                min_value=0.0,
                value=float(record["importe"]) if record is not None else 0.0,
                step=1000.0,
            )

            importe_ars_preview = importe * tc if moneda == "USD" else importe
            st.caption(f"💰 Equivalente en ARS: **{fmt_ars(importe_ars_preview)}**")

            estado = st.selectbox(
                "Estado",
                core.ESTADOS_VALIDOS,
                index=core.ESTADOS_VALIDOS.index(record["estado"]) if record is not None else 0,
            )
            obs = st.text_area("Observaciones", value=record["obs"] if record is not None else "")

            b1, b2 = st.columns(2)
            submitted = b1.form_submit_button(
                "💾 Guardar cambios" if record is not None else "➕ Registrar pago",
                type="primary",
                use_container_width=True,
            )
            cancelled = b2.form_submit_button("✖️ Cancelar edición", use_container_width=True)

            if submitted:
                if not proveedor:
                    st.error("Debe indicar un proveedor.")
                else:
                    df_current = get_df()
                    if record is not None:
                        idx = df_current.index[df_current["id"] == editing_id][0]
                        df_current.loc[idx, ["fecha", "proveedor", "proyecto", "moneda", "importe", "tc", "estado", "obs"]] = [
                            pd.Timestamp(fecha_pago), proveedor, proyecto, moneda, importe, tc, estado, obs,
                        ]
                        st.success(f"Pago {editing_id} actualizado.")
                        st.session_state.editing_id = None
                    else:
                        new_id = core.next_id(df_current)
                        new_row = pd.DataFrame(
                            [{
                                "id": new_id, "fecha": pd.Timestamp(fecha_pago), "proveedor": proveedor,
                                "proyecto": proyecto, "moneda": moneda, "importe": importe, "tc": tc,
                                "estado": estado, "obs": obs,
                            }]
                        )
                        df_current = pd.concat([df_current, new_row], ignore_index=True)
                        st.success(f"Pago {new_id} registrado.")
                    set_df(df_current)
                    st.rerun()

            if cancelled:
                st.session_state.editing_id = None
                st.rerun()

    with col_table:
        st.markdown("##### Pagos registrados (según filtros de la barra lateral)")
        display_df = df_filtered[
            ["id", "fecha", "semana_etiqueta", "proveedor", "proyecto", "moneda", "importe", "tc", "importe_ars", "estado", "obs"]
        ].rename(
            columns={
                "id": "ID", "fecha": "Fecha", "semana_etiqueta": "Semana", "proveedor": "Proveedor",
                "proyecto": "Proyecto", "moneda": "Moneda", "importe": "Importe Pactado", "tc": "TC",
                "importe_ars": "Importe ARS", "estado": "Estado", "obs": "Observaciones",
            }
        ).sort_values("Fecha")

        st.dataframe(
            display_df.style.format({"Fecha": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                                      "Importe Pactado": "{:,.2f}", "TC": "{:,.2f}", "Importe ARS": fmt_ars}),
            use_container_width=True,
            height=380,
        )

        st.markdown("###### Acciones rápidas")
        ids_disponibles = display_df["ID"].tolist()
        if ids_disponibles:
            sel_col, edit_col, del_col = st.columns([2, 1, 1])
            sel_id = sel_col.selectbox("Seleccionar pago por ID", ids_disponibles, label_visibility="collapsed")
            if edit_col.button("✏️ Editar", use_container_width=True):
                st.session_state.editing_id = sel_id
                st.rerun()
            if del_col.button("🗑️ Eliminar", use_container_width=True):
                df_current = get_df()
                df_current = df_current[df_current["id"] != sel_id]
                set_df(df_current)
                st.warning(f"Pago {sel_id} eliminado.")
                st.rerun()
        else:
            st.info("No hay pagos para mostrar con los filtros actuales.")

# ---------------------------------------------------------------------------
# TAB 4 — Control por Proveedor
# ---------------------------------------------------------------------------

with tab_proveedores:
    st.subheader("Control integral por proveedor")
    ctrl = core.provider_control(df_filtered)
    if ctrl.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        show = ctrl.rename(
            columns={
                "proveedor": "Proveedor", "cantidad_pagos": "Cant. Pagos", "primer_pago": "Primer Pago",
                "ultimo_pago": "Último Pago", "moneda_base": "Moneda Base",
                "compromiso_total_ars": "Compromiso Total ARS", "pagado_ars": "Pagado ARS",
                "saldo_ars": "Saldo ARS", "pct_del_flujo": "% del Flujo", "pct_avance": "% Avance",
            }
        )
        st.dataframe(
            show.style.format({
                "Primer Pago": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                "Último Pago": lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "",
                "Compromiso Total ARS": fmt_ars, "Pagado ARS": fmt_ars, "Saldo ARS": fmt_ars,
                "% del Flujo": "{:.2f}%", "% Avance": "{:.2f}%",
            }),
            use_container_width=True,
            height=280,
        )

        st.markdown("##### % Avance por proveedor")
        for _, r in ctrl.iterrows():
            st.write(f"**{r['proveedor']}** — {r['pct_avance']:.1f}% pagado ({fmt_ars(r['pagado_ars'])} de {fmt_ars(r['compromiso_total_ars'])})")
            st.progress(min(r["pct_avance"] / 100, 1.0))

# ---------------------------------------------------------------------------
# TAB 5 — Exportación
# ---------------------------------------------------------------------------

with tab_export:
    st.subheader("Exportar base actualizada")
    st.write(
        "Descarga la base completa (sin aplicar los filtros de la barra lateral) "
        "manteniendo la estructura normalizada de 4 hojas: **Resumen Ejecutivo**, "
        "**DB_Pagos**, **Matriz Semanal** y **Control Proveedores**."
    )

    df_export = get_df()
    xlsx_bytes = core.export_to_excel(df_export)

    st.download_button(
        label="⬇️ Descargar SISTEMA_GESTION_EFECTIVO_L2_TUC.xlsx",
        data=xlsx_bytes,
        file_name=f"SISTEMA_GESTION_EFECTIVO_L2_TUC_{dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()
    st.caption(f"Base actual en memoria: {len(df_export)} pagos registrados.")
    st.dataframe(
        df_export[["id", "fecha", "proveedor", "moneda", "importe", "tc", "importe_ars", "estado"]],
        use_container_width=True,
        height=300,
    )
