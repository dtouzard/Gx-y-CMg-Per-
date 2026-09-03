# COES Perú — App

App Streamlit para explorar CMg, Generación y Cruce de ingresos (COES Perú),
para el equipo de M&A de Colbún. Lee únicamente de la base publicada en este
mismo repo (`04. Base de datos/_base`, Parquet) — no consulta al COES en
vivo.

Desplegado en Streamlit Community Cloud. Archivo principal: `05. App/App_COES_Peru.py`.

## Actualizar la base

Este repo es una foto de la base al momento de subirla. Para actualizarla:
1. Correr los scripts de descarga en el proyecto original (fuera de este repo).
2. Reemplazar la carpeta `04. Base de datos/_base` acá.
3. `git add`, `git commit`, `git push` — Streamlit Cloud redespliega solo.
