# -*- coding: utf-8 -*-
"""Carga y consulta de los datos fuente (Excel) para la app de Streamlit.

Los 3 archivos en `data/` son la fuente de verdad y se actualizan a mano
(se reemplazan y se hace `git push`; Streamlit Community Cloud vuelve a
desplegar automáticamente con cada push, sin pasos manuales adicionales):

    cubos_completos_plan_justicia.xlsx         -> histórico de productividad
    metas_planes_justicia.xlsx                 -> catálogo de planes de justicia + metas anuales
    procedimientos_personas_plan_justicia.xlsx -> personas atendidas por procedimiento/año
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data" / "plan_justicia"

RUTA_CUBOS = DATA_DIR / "cubos_completos_plan_justicia.xlsx"
RUTA_METAS = DATA_DIR / "metas_planes_justicia.xlsx"
RUTA_PROCEDIMIENTOS_PERSONAS = DATA_DIR / "procedimientos_personas_plan_justicia.xlsx"


@st.cache_data(show_spinner=False)
def cargar_cubos() -> pd.DataFrame:
    df = pd.read_excel(RUTA_CUBOS)
    df["fecha"] = pd.to_datetime(df["fecha"])
    return df


@st.cache_data(show_spinner=False)
def cargar_metas() -> pd.DataFrame:
    return pd.read_excel(RUTA_METAS)


@st.cache_data(show_spinner=False)
def cargar_clues_info() -> pd.DataFrame:
    metas = cargar_metas()
    info = metas[[
        "clues_imb", "entidad", "categoria_gerencial",
        "estatus_de_operacion", "nombre_de_la_unidad", "nivel_atencion",
    ]].rename(columns={"clues_imb": "clues", "nombre_de_la_unidad": "nombre"})
    info["id"] = info["clues"]
    return info


@st.cache_data(show_spinner=False)
def cargar_metas_clues() -> pd.DataFrame:
    metas = cargar_metas()
    return metas[[
        "clues_imb", "meta_general_anual", "meta_especialidad_anual",
        "meta_cirugia_anual", "meta_egresos_anual",
    ]].rename(columns={"clues_imb": "clues"})


@st.cache_data(show_spinner=False)
def opciones_selector_clues() -> dict:
    """Devuelve {etiqueta: plan_de_justicia} ordenado por entidad/nombre, con NACIONAL primero.

    La etiqueta y el valor son el nombre del plan de justicia tal como
    aparece en metas_planes_justicia.xlsx (columna clues_imb) — no un
    código CLUES.
    """
    info = cargar_clues_info().copy()
    info["etiqueta"] = info.apply(
        lambda r: r["clues"] if pd.isna(r["nombre"]) or str(r["nombre"]).strip() == ""
        else f"{r['clues']} - {r['nombre']}",
        axis=1,
    )
    info = info.sort_values(["entidad", "clues"], na_position="last")

    opciones = {"NACIONAL": "NACIONAL"}
    for _, row in info.iterrows():
        opciones[row["etiqueta"]] = row["clues"]
    return opciones


@st.cache_data(show_spinner=False)
def consultar_datos(clues_seleccionada: str) -> pd.DataFrame:
    """Equivalente a `construir_consulta_clues` + `dbGetQuery` del módulo R."""
    df = cargar_cubos()
    df = df[df["fecha"].notna()]

    if clues_seleccionada != "NACIONAL":
        df = df[df["plan de justicia"] == clues_seleccionada]

    out = (
        df.groupby("fecha", as_index=False)
        .agg(
            consulta_total=("consultas_totales", "sum"),
            consulta_general=("consultas_generales", "sum"),
            consulta_especialidad=("consultas_de_especialidad", "sum"),
            procedimientos_qx=("procedimientos_quirurgicos", "sum"),
            egresos=("egresos", "sum"),
        )
        .sort_values("fecha")
        .reset_index(drop=True)
    )
    return out


@st.cache_data(show_spinner=False)
def cargar_procedimientos_personas() -> pd.DataFrame:
    df = pd.read_excel(RUTA_PROCEDIMIENTOS_PERSONAS)
    return df.dropna(subset=["anio_insert", "tipo_procedimiento"])


@st.cache_data(show_spinner=False)
def procedimientos_personas_filtrado(clues_seleccionada: str) -> pd.DataFrame:
    """Personas atendidas por tipo de procedimiento y año para el plan
    de justicia seleccionado (columna `personas` del Excel de
    procedimientos_personas), en el formato que espera
    `crear_reporte_productividad` como parámetro `procedimientos_personas`.
    """
    df = cargar_procedimientos_personas()

    if clues_seleccionada != "NACIONAL":
        df = df[df["id"] == clues_seleccionada]

    out = (
        df.groupby(["anio_insert", "tipo_procedimiento"], as_index=False)
        .agg(procedimientos=("procedimientos", "sum"), personas=("personas", "sum"))
        .rename(columns={"anio_insert": "fecha"})
    )
    return out[["fecha", "tipo_procedimiento", "procedimientos", "personas"]]
