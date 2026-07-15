# -*- coding: utf-8 -*-
"""Vista de Streamlit para el modo "Plan de Justicia".

`render()` dibuja toda la pantalla — el `app.py` de la raíz solo decide
cuándo llamarla (según el modo elegido en el selector superior).
"""

from __future__ import annotations

import io
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
import streamlit as st

from py.plan_justicia.pptx_report import crear_reporte_productividad, fmt_num
from py.plan_justicia.data_io import (
    cargar_clues_info,
    cargar_metas_clues,
    consultar_datos,
    opciones_selector_clues,
    procedimientos_personas_filtrado,
)

COLOR_GUINDA = "#611232"
COLOR_VERDE = "#1E5B4F"
COLOR_BEIGE = "#D9D2BE"
COLOR_BEIGE_2026 = "#B08D57"
COLOR_MUTED = "#6B7280"

MAPA_COL_ANUAL = {
    "consulta_general": "consulta_general_anual",
    "consulta_especialidad": "consulta_especialidad_anual",
    "procedimientos_qx": "procedimientos_qx_anual",
    "egresos": "egresos_anual",
}
MAPA_META = {
    "consulta_general": "meta_general_anual",
    "consulta_especialidad": "meta_especialidad_anual",
    "procedimientos_qx": "meta_cirugia_anual",
    "egresos": "meta_egresos_anual",
}


# ---------------------------------------------------------------------------
# Gráficas de avance vs. meta (en pantalla — matplotlib está bien aquí,
# no necesitan ser editables como las del PPTX)
# ---------------------------------------------------------------------------
def _grafica_avance_meta(datos, variable_sel, titulo, datos_anual_grafica, metas_filtrado):
    d = datos.copy()
    d["anio"] = d["fecha"].dt.year
    fecha_corte = d["fecha"].max()
    mes_corte, dia_corte = fecha_corte.month, fecha_corte.day

    anios = [2024, 2025, 2026]
    filas = []
    for anio in anios:
        try:
            corte_anio = pd.Timestamp(year=anio, month=mes_corte, day=dia_corte)
        except ValueError:
            corte_anio = pd.Timestamp(year=anio, month=mes_corte, day=1) + pd.offsets.MonthEnd(0)
        sub = d[d["anio"] == anio]
        avance = sub.loc[sub["fecha"] <= corte_anio, variable_sel].sum()
        filas.append({"anio": anio, "avance": avance})
    df_avance = pd.DataFrame(filas)

    hay_2026 = bool((d.loc[d["anio"] == 2026, variable_sel] > 0).any())

    col_anual = MAPA_COL_ANUAL[variable_sel]
    df_total = (
        datos_anual_grafica[datos_anual_grafica["anio"].isin(anios)][["anio", col_anual]]
        .rename(columns={col_anual: "total_anual"})
        .copy()
    )
    df_total = pd.DataFrame({"anio": anios}).merge(df_total, on="anio", how="left")
    df_total["total_anual"] = df_total["total_anual"].fillna(0)

    if hay_2026:
        meta_col = MAPA_META[variable_sel]
        meta_valor = 0.0
        if meta_col in metas_filtrado.columns:
            meta_valor = pd.to_numeric(metas_filtrado[meta_col], errors="coerce").fillna(0).sum()
        df_total.loc[df_total["anio"] == 2026, "total_anual"] = meta_valor

    df_plot = df_avance.merge(df_total, on="anio", how="left")
    df_plot["pendiente"] = (df_plot["total_anual"] - df_plot["avance"]).clip(lower=0)

    fig, ax = plt.subplots(figsize=(5.2, 4.2), dpi=150)
    x = np.arange(len(anios))
    avances = df_plot["avance"].to_numpy(dtype=float)
    pendientes = df_plot["pendiente"].to_numpy(dtype=float)
    totales = df_plot["total_anual"].to_numpy(dtype=float)
    colores_pendiente = [COLOR_BEIGE_2026 if a == 2026 else COLOR_BEIGE for a in anios]

    ax.bar(x, avances, width=0.62, color=COLOR_VERDE, zorder=2)
    ax.bar(x, pendientes, width=0.62, bottom=avances, color=colores_pendiente, zorder=2)

    for xi, tot in zip(x, totales):
        if tot > 0:
            ax.text(xi, tot, fmt_num(tot), ha="center", va="bottom",
                    fontweight="bold", fontsize=10.5, color="black")

    for xi, av, tot, pend in zip(x, avances, totales, pendientes):
        if tot > 0 and av > 0:
            ax.text(xi, av / 2, fmt_num(av), ha="center", va="center",
                    color="white", fontweight="bold", fontsize=10.5)
            pct = av / tot
            ax.text(xi, av + pend * 0.1, f"{round(pct * 100)}%", ha="center", va="bottom",
                    fontweight="bold", fontsize=10.5, color="black")

    ax.set_xticks(x)
    ax.set_xticklabels([str(a) for a in anios], fontweight="bold", color=COLOR_MUTED, fontsize=12)
    ymax = max(totales.max() if len(totales) else 1, 1)
    ax.set_ylim(0, ymax * 1.22)
    ax.tick_params(axis="y", colors=COLOR_MUTED, length=0, labelsize=8)
    ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: fmt_num(v)))
    ax.grid(axis="y", color="#E5E7EB", linewidth=0.6, zorder=0)
    ax.set_title(titulo, fontsize=15, fontweight="bold", color=COLOR_MUTED, pad=12)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(axis="x", length=0)

    legend_elems = [
        Patch(facecolor=COLOR_BEIGE, label="Resto del año"),
        Patch(facecolor=COLOR_VERDE, label="Avance al corte"),
        Patch(facecolor=COLOR_BEIGE_2026, label="Meta"),
    ]
    ax.legend(handles=legend_elems, loc="upper center", bbox_to_anchor=(0.5, -0.08),
              ncol=3, frameon=False, fontsize=8.5)

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def _crear_excel_bytes(clues_seleccionada, datos, clues_info):
    info_sel = clues_info[
        (clues_info["clues"] == clues_seleccionada) | (clues_info["id"] == clues_seleccionada)
    ].head(1)

    if len(info_sel) == 0:
        fila = {
            "nombre": clues_seleccionada,
            "entidad": "NACIONAL" if clues_seleccionada == "NACIONAL" else None,
            "nivel_atencion": None,
            "categoria_gerencial": None,
            "estatus_de_operacion": None,
        }
    else:
        fila = info_sel.iloc[0].to_dict()

    tabla_info = pd.DataFrame({
        "campo": [
            "Selección", "Nombre", "Entidad", "Nivel de atención",
            "Categoría gerencial", "Estatus de operación", "Fecha de corte",
        ],
        "valor": [
            clues_seleccionada,
            fila.get("nombre"),
            fila.get("entidad"),
            fila.get("nivel_atencion"),
            fila.get("categoria_gerencial"),
            fila.get("estatus_de_operacion"),
            date.today().strftime("%d/%m/%Y"),
        ],
    })

    d = datos.copy()
    d["anio"] = d["fecha"].dt.year
    num_cols = [c for c in d.select_dtypes(include=[np.number]).columns if c != "anio"]
    resumen_anual = d.groupby("anio", as_index=False)[num_cols].sum().sort_values("anio")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabla_info.to_excel(writer, sheet_name="resumen", index=False, startrow=0)
        resumen_anual.to_excel(writer, sheet_name="resumen", index=False, startrow=9)
        datos.to_excel(writer, sheet_name="productividad detalle", index=False)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Pantalla
# ---------------------------------------------------------------------------
def render():
    st.title("📊 Productividad por Plan de Justicia")
    st.caption("IMSS Bienestar · Información 2020 a la fecha")

    clues_info = cargar_clues_info()
    metas_clues = cargar_metas_clues()
    opciones = opciones_selector_clues()

    etiqueta_sel = st.selectbox(
        "Selecciona un Plan de Justicia:",
        options=list(opciones.keys()),
        index=0,
        placeholder="Busca por plan de justicia",
    )
    clues_seleccionada = opciones[etiqueta_sel]

    with st.spinner("Consultando datos..."):
        datos = consultar_datos(clues_seleccionada)

    if datos.empty:
        st.error(f"No se encontraron datos para: {clues_seleccionada}")
        st.stop()

    st.success(f"✅ Consulta exitosa: {len(datos):,} registros encontrados para: {clues_seleccionada}")

    datos_anual_grafica = (
        datos.assign(anio=datos["fecha"].dt.year)
        .query("anio in [2024, 2025, 2026]")
        .groupby("anio", as_index=False)
        .agg(
            consulta_general_anual=("consulta_general", "sum"),
            consulta_especialidad_anual=("consulta_especialidad", "sum"),
            procedimientos_qx_anual=("procedimientos_qx", "sum"),
            egresos_anual=("egresos", "sum"),
        )
    )

    if clues_seleccionada == "NACIONAL":
        metas_filtrado = metas_clues
    else:
        metas_filtrado = metas_clues[metas_clues["clues"] == clues_seleccionada]

    col1, col2 = st.columns(2)
    with col1:
        st.pyplot(_grafica_avance_meta(datos, "consulta_general", "Consulta general",
                                        datos_anual_grafica, metas_filtrado))
    with col2:
        st.pyplot(_grafica_avance_meta(datos, "consulta_especialidad", "Consulta de especialidad",
                                        datos_anual_grafica, metas_filtrado))

    col3, col4 = st.columns(2)
    with col3:
        st.pyplot(_grafica_avance_meta(datos, "procedimientos_qx", "Procedimientos quirúrgicos",
                                        datos_anual_grafica, metas_filtrado))
    with col4:
        st.pyplot(_grafica_avance_meta(datos, "egresos", "Egresos",
                                        datos_anual_grafica, metas_filtrado))

    st.divider()

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        excel_bytes = _crear_excel_bytes(clues_seleccionada, datos, clues_info)
        st.download_button(
            "⬇️ Descargar datos (Excel)",
            data=excel_bytes,
            file_name=f"datos_planes_justicia_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="pj_descargar_excel",
        )

    with dcol2:
        if st.button("📄 Generar informe (PowerPoint)", key="pj_generar_pptx"):
            with st.spinner("Generando informe en PowerPoint..."):
                presentacion = crear_reporte_productividad(
                    codigo_clues=clues_seleccionada,
                    clues_info=clues_info,
                    metas=metas_clues,
                    historicos=datos,
                    procedimientos_personas=procedimientos_personas_filtrado(clues_seleccionada),
                )
                pptx_buffer = io.BytesIO()
                presentacion.save(pptx_buffer)
                st.session_state["pj_pptx_bytes"] = pptx_buffer.getvalue()
                st.session_state["pj_pptx_clues"] = clues_seleccionada
            st.success("¡Informe generado exitosamente!")

        if (st.session_state.get("pj_pptx_bytes")
                and st.session_state.get("pj_pptx_clues") == clues_seleccionada):
            st.download_button(
                "⬇️ Descargar informe (.pptx)",
                data=st.session_state["pj_pptx_bytes"],
                file_name=f"datos_clues_{clues_seleccionada}_{date.today()}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="pj_descargar_pptx",
            )

    with st.expander("ℹ️ Ayuda / ¿Cómo usar esta app?"):
        st.markdown(
            """
            1. Selecciona un plan de justicia del buscador (o **NACIONAL** para el agregado nacional).
            2. Las gráficas de arriba muestran el avance de productividad contra la meta anual.
            3. Puedes descargar los datos en formato Excel.
            4. También puedes generar y descargar el informe completo en PowerPoint.
            """
        )
