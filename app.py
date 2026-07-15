# -*- coding: utf-8 -*-
"""Productividad IMSS Bienestar — app unificada.

Fusiona dos apps que hacían básicamente lo mismo (mismo motor de reporte en
PowerPoint) sobre dos niveles de datos distintos:

    - "Plan de Justicia": agrupa por los 32 planes de justicia.
    - "CLUES / Unidad médica": por unidad médica individual.

Cada modo sigue jalando de su propia fuente de datos (no se combinan) —
lo único unificado es la app y el motor de gráficas nativas del PPTX
(`py/formas_nativas.py`).
"""

import streamlit as st

st.set_page_config(
    page_title="Productividad IMSS Bienestar",
    page_icon="📊",
    layout="wide",
)

st.sidebar.markdown("### 📊 Productividad IMSS Bienestar")
modo = st.sidebar.radio(
    "Ver por:",
    ["Plan de Justicia", "CLUES / Unidad médica"],
    key="modo_app",
)
st.sidebar.divider()
st.sidebar.caption("Información 2020 a la fecha")

if modo == "Plan de Justicia":
    from py.plan_justicia.vista import render
else:
    from py.clues.vista import render

render()
