# Productividad IMSS Bienestar

App de Streamlit que fusiona dos reportes de productividad que compartían el
mismo motor de PowerPoint pero se habían mantenido como proyectos separados:

- **Plan de Justicia** — agrupa por los 32 planes de justicia.
- **CLUES / Unidad médica** — por unidad médica individual.

Cada modo sigue usando su propia fuente de datos; lo único que se comparte
es la app y el generador de gráficas nativas del PPTX.

## Estructura

```
app.py                     <- selector "Ver por" + despacha al modo elegido
py/
  formas_nativas.py         <- kit compartido: formas nativas de PowerPoint
                                (editables) y las gráficas de barras / serie
                                temporal que se construyen con ellas
  plan_justicia/
    data_io.py               <- carga y consulta de los 3 Excel
    pptx_report.py            <- tarjetas, tablas y armado del reporte
    vista.py                  <- pantalla de Streamlit
  clues/
    data_loader.py, queries.py, charts.py, excel_export.py,
    utils_comunes.py, pptx_report.py
    vista.py                  <- pantalla de Streamlit
data/
  master_presentacion.pptx   <- plantilla compartida por ambos modos
  plan_justicia/              <- 3 Excel fuente
  clues/                       <- 5 parquet fuente
```

## Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Actualizar datos

Cada modo se actualiza reemplazando sus archivos en `data/plan_justicia/` o
`data/clues/` y haciendo `git push` — Streamlit Community Cloud vuelve a
desplegar automáticamente.
