# -*- coding: utf-8 -*-
"""Productividad IMSS Bienestar — app unificada.

Fusiona dos apps que hacían básicamente lo mismo (mismo motor de reporte en
PowerPoint) sobre dos niveles de datos distintos:

    - "Plan de Justicia": agrupa por los 32 planes de justicia.
    - "CLUES / Unidad médica": por unidad médica individual.

Cada fuente sigue jalando de sus propios datos (no se combinan), pero se
eligen desde un solo selector — no hay pantallas ni pestañas separadas por
modo. El único "NACIONAL" que se muestra es el del dataset CLUES; el
NACIONAL propio de Plan de Justicia se omite del selector para no
duplicar la etiqueta.
"""

import base64
from pathlib import Path

import streamlit as st

from py.clues.data_loader import load_choices, load_clues_info as cargar_clues_info_clues, load_metas as cargar_metas_clues
from py.clues.vista import render_seleccion as render_clues
from py.plan_justicia.data_io import (
    cargar_clues_info as cargar_clues_info_plan_justicia,
    cargar_metas_clues as cargar_metas_plan_justicia,
    opciones_selector_clues,
)
from py.plan_justicia.vista import render_seleccion as render_plan_justicia

st.set_page_config(
    page_title="Productividad IMSS Bienestar",
    page_icon="📊",
    layout="wide",
)

GUINDA = "#611232"
GUINDA_OSCURO = "#4E0D28"
DORADO = "#BC955C"
DORADO_OSCURO = "#A57F2C"
FONDO = "#F8F7F5"

ASSETS = Path(__file__).parent / "assets"


@st.cache_data(show_spinner=False)
def _img_b64(nombre: str) -> str:
    return base64.b64encode((ASSETS / nombre).read_bytes()).decode()


LOGO_GOB = _img_b64("logo_gobierno_mexico.png")
LOGO_IMSS = _img_b64("logo_imss_bienestar.png")

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: {FONDO}; }}
        div.stButton > button, div.stDownloadButton > button {{
            background-color: {DORADO};
            color: white;
            font-weight: 700;
            border: none;
        }}
        div.stButton > button:hover, div.stDownloadButton > button:hover {{
            background-color: {DORADO_OSCURO};
            color: white;
        }}
        #MainMenu, footer {{ visibility: hidden; }}
    </style>
    """,
    unsafe_allow_html=True,
)

def _hero():
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, {GUINDA} 0%, {GUINDA_OSCURO} 100%);
            border-radius: 10px;
            padding: 2.2rem 2.5rem;
            margin-bottom: 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1.5rem;
        ">
            <div>
                <div style="color: {DORADO}; font-weight: 700; letter-spacing: 2px;
                            text-transform: uppercase; font-size: 0.85rem; margin-bottom: 0.4rem;">
                    IMSS Bienestar
                </div>
                <div style="color: white; font-weight: 800; font-size: 2.4rem; line-height: 1.15;">
                    Productividad IMSS Bienestar
                </div>
                <div style="color: #E9DDCC; font-size: 1.05rem; margin-top: 0.5rem;">
                    Consulta de productividad médica · Información 2020 a la fecha
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 1.5rem; background: white;
                        padding: 0.9rem 1.4rem; border-radius: 8px;">
                <img src="data:image/png;base64,{LOGO_GOB}" style="height: 42px;">
                <div style="width: 1px; height: 42px; background: #D1D5DB;"></div>
                <img src="data:image/png;base64,{LOGO_IMSS}" style="height: 42px;">
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def _construir_opciones():
    """Combina las opciones de CLUES y de Plan de Justicia en un solo dict
    {etiqueta: (tipo, valor)}. El NACIONAL de Plan de Justicia se omite
    para no duplicar la etiqueta "NACIONAL" (se usa el de CLUES)."""
    opciones: dict[str, tuple[str, str]] = {}

    choices = load_choices()
    label_map = dict(zip(choices["value"], choices["label"]))
    for valor in choices["value"]:
        opciones[label_map[valor]] = ("clues", valor)

    for etiqueta, valor in opciones_selector_clues().items():
        if valor == "NACIONAL":
            continue
        opciones.setdefault(etiqueta, ("plan_justicia", valor))

    return opciones


def _pantalla_consulta():
    with st.sidebar:
        st.markdown("### 📊 Productividad IMSS Bienestar")
        st.caption("Información 2020 a la fecha")

    opciones = _construir_opciones()
    etiquetas = list(opciones.keys())
    default_idx = next(
        (i for i, e in enumerate(etiquetas) if opciones[e] == ("clues", "NACIONAL")), 0
    )

    etiqueta_sel = st.selectbox(
        "Selecciona una unidad médica o un plan de justicia:",
        options=etiquetas,
        index=default_idx,
        placeholder="Busca por CLUES o por plan de justicia",
        key="selector_unificado",
    )

    tipo, valor = opciones[etiqueta_sel]

    if tipo == "clues":
        render_clues(valor, cargar_clues_info_clues(), cargar_metas_clues())
    else:
        render_plan_justicia(valor, cargar_clues_info_plan_justicia(), cargar_metas_plan_justicia())


_hero()
_pantalla_consulta()
