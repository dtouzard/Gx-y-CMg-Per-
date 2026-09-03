# -*- coding: utf-8 -*-
"""
PERU - App interactiva (Streamlit) para explorar CMg, Generacion y el
Cruce de ingresos, sin tocar el COES en vivo.

Esta app SOLO lee de "04. Base de datos/_base" (Parquet, ver
Actualizar_Base_Peru.py) - nunca le pega a COES/BCCh/BLS. Por eso, a
diferencia de los otros 3 scripts, esta SI se podria alojar en un
servidor en la nube sin el problema de bloqueo geografico (ver CLAUDE.md)
... aunque el hosting todavia esta sin definir, por ahora corre local.

Si los datos que ves acá están desactualizados (falta el último mes,
etc.), no es un bug de la app: hay que correr primero
Generacion_lookup_Peru.py / CMg_lookup_Peru.py con el período nuevo en
Parametros_Comunes.xlsx, y después Actualizar_Base_Peru.py para
publicarlo a la base - esta app no descarga nada por si sola.

El Cruce (que central le corresponde a que barra) se arma AL VUELO aca,
eligiendo de dropdowns - no depende de que alguien haya precomputado ese
cruce de antemano (ver CLAUDE.md, "Objetivo final ampliado").

*** USD real: la base ya trae "cmg_usd_nominal" y "cmg_usd_real"
calculados con la FECHA_BASE_REAL que estaba configurada en
Parametros_Comunes.xlsx al momento de publicar - pero esta app NO usa esa
columna tal cual, porque la fecha base para "USD real" es justo algo que
el usuario quiere poder elegir en el momento (ej. "quiero ver todo en
dólares reales de diciembre 2025"), sin tener que volver a correr el
backfill. Se recalcula al vuelo con la misma formula que usa
CMg_lookup_Peru.py (cmg_usd_real = cmg_usd_nominal * cpi_base / cpi_mes),
usando la columna "cpi" que ya viene en la base (mismo valor para todas
las barras de un mes dado).

*** Empresa / Tecnología / Potencia (capacidad instalada) por central:
la base de Generación (04. Base de datos/_base/generacion_horaria) NO
guarda esos 3 datos - solo guarda [fecha_hora, central, generacion_mwh]
(ver Actualizar_Base_Peru.py, gx.cargar_horario_por_central agrupa por
central sola). Se sacan de "0. Inputs/Empresas_Centrales_COES.xlsx"
(hoja "Por central" - reemplazó a "Diccionario Potencia.xlsx" el
2026-09-03: se armó cruzando la cache RAW de Generacion_lookup_Peru.py
contra el diccionario viejo, y Dani lo curó a mano), el mismo archivo
que ya usa Cruce_Ingresos_Peru.py para la potencia - el nombre de
central ahí no siempre coincide exacto con el que devuelve COES, así que
una central de la base puede no tener match (se muestra sin agrupar/sin
factor de planta en vez de romper).

Instalar una vez (TERMINAL):
    python -m pip install streamlit pandas pyarrow openpyxl plotly

Correr (TERMINAL, desde esta carpeta o indicando la ruta completa):
    streamlit run "App_COES_Peru.py"
"""

import io
import os
import sys
import glob
import math
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
from _coes_common import periodos_rango  # noqa: E402

BASE_DIR = os.path.join(_RAIZ, "04. Base de datos", "_base")
CMG_BASE_DIR = os.path.join(BASE_DIR, "cmg_horario")
GX_BASE_DIR = os.path.join(BASE_DIR, "generacion_horaria")
POTENCIA_XLSX = os.path.join(_RAIZ, "0. Inputs", "Empresas_Centrales_COES.xlsx")

MONEDAS_CMG_SIMPLES = {
    "Soles (S/.)": "cmg_sol_mwh",
    "USD nominal": "cmg_usd_nominal",
}

# Paleta de "Gráficos usuales.pptx" (0. Inputs): azul = volumen/generación,
# rojo = precio/tasa/porcentaje. Se mantiene esa convención en toda la app.
COLOR_AZUL = "#7A99BF"
COLOR_ROJO = "#C00000"
COLOR_GRILLA = "#E7E6E6"

# Definición de Dani (2026-09-02, convención de hora 1-24): Base cruza la
# medianoche (24,1..8), Media 9-18, Punta 19-23.
HORAS_BASE = {24, 1, 2, 3, 4, 5, 6, 7, 8}
HORAS_MEDIA = set(range(9, 19))
HORAS_PUNTA = set(range(19, 24))


# --------------------------- helpers de carga (cacheados) ---------------------------

@st.cache_data(show_spinner=False)
def _meses_disponibles(carpeta):
    """(anio, mes) de cada particion mensual presente en la carpeta, a
    partir del nombre de archivo (AAAA-MM.parquet) - asi la app no
    depende de que le hardcodeemos un rango, se ajusta sola a lo que ya
    esta publicado en la base."""
    archivos = glob.glob(os.path.join(carpeta, "*.parquet"))
    meses = []
    for f in archivos:
        nombre = os.path.splitext(os.path.basename(f))[0]
        try:
            a, m = nombre.split("-")
            meses.append((int(a), int(m)))
        except ValueError:
            continue
    return sorted(meses)


@st.cache_data(show_spinner=False)
def _valores_distintos(carpeta, columna):
    """Valores distintos de una columna (ej. 'barra' o 'central') a
    traves de TODAS las particiones - se lee solo esa columna (Parquet
    es columnar, rapido) para poblar los dropdowns."""
    archivos = sorted(glob.glob(os.path.join(carpeta, "*.parquet")))
    valores = set()
    for f in archivos:
        valores.update(pd.read_parquet(f, columns=[columna])[columna].unique())
    return sorted(valores)


@st.cache_data(show_spinner=False)
def _cpi_por_mes():
    """{(anio, mes): cpi} a partir de la base de CMg - el CPI es el mismo
    para todas las barras de un mes dado, asi que alcanza con leer esas
    3 columnas (liviano) de cualquier particion. Sirve para recalcular
    USD real con la fecha base que el usuario elija en la app (ver
    docstring del modulo)."""
    archivos = sorted(glob.glob(os.path.join(CMG_BASE_DIR, "*.parquet")))
    piezas = []
    for f in archivos:
        piezas.append(pd.read_parquet(f, columns=["anio", "mes", "cpi"]).drop_duplicates())
    if not piezas:
        return {}
    df = pd.concat(piezas, ignore_index=True).drop_duplicates(subset=["anio", "mes"])
    return {(int(r.anio), int(r.mes)): r.cpi for r in df.itertuples()}


@st.cache_data(show_spinner=False)
def _cargar_diccionario_potencia():
    """Central -> {MW, Tecnología, Empresa Usado}, desde
    Empresas_Centrales_COES.xlsx (mismo archivo/hoja que usa
    Cruce_Ingresos_Peru.py)."""
    df = pd.read_excel(POTENCIA_XLSX, sheet_name="Por central")
    df = df[["Central (COES)", "Tecnología (actual)", "Empresa Usado (actual)", "Potencia MW (actual)"]]
    df = df.dropna(subset=["Central (COES)"])
    df = df.rename(columns={"Central (COES)": "Central", "Tecnología (actual)": "Tecnología",
                             "Empresa Usado (actual)": "Empresa Usado", "Potencia MW (actual)": "MW"})
    df["Central"] = df["Central"].astype(str).str.strip()
    return df.drop_duplicates(subset="Central").set_index("Central")


@st.cache_data(show_spinner="Leyendo CMg de la base...")
def _cargar_cmg(barras, anio_ini, mes_ini, anio_fin, mes_fin):
    piezas = []
    for anio, mes in periodos_rango(anio_ini, mes_ini, anio_fin, mes_fin):
        ruta = os.path.join(CMG_BASE_DIR, f"{anio}-{mes:02d}.parquet")
        if not os.path.exists(ruta):
            continue
        d = pd.read_parquet(ruta)
        d = d[d["barra"].isin(barras)]
        if not d.empty:
            piezas.append(d)
    if not piezas:
        return pd.DataFrame(columns=["fecha_hora", "anio", "mes", "dia", "hora", "barra",
                                      "cmg_sol_mwh", "tc", "cpi", "cmg_usd_nominal", "cmg_usd_real"])
    return pd.concat(piezas, ignore_index=True)


@st.cache_data(show_spinner="Leyendo Generación de la base...")
def _cargar_generacion(centrales, anio_ini, mes_ini, anio_fin, mes_fin):
    piezas = []
    for anio, mes in periodos_rango(anio_ini, mes_ini, anio_fin, mes_fin):
        ruta = os.path.join(GX_BASE_DIR, f"{anio}-{mes:02d}.parquet")
        if not os.path.exists(ruta):
            continue
        d = pd.read_parquet(ruta)
        d = d[d["central"].isin(centrales)]
        if not d.empty:
            piezas.append(d)
    if not piezas:
        return pd.DataFrame(columns=["fecha_hora", "anio", "mes", "dia", "hora", "central", "generacion_mwh"])
    return pd.concat(piezas, ignore_index=True)


@st.cache_data(show_spinner=False)
def _excel_descargable(hojas):
    """hojas: lista de (nombre_hoja, dataframe). Devuelve bytes de un
    .xlsx armado en memoria, para el boton de descarga.

    @st.cache_data ES IMPORTANTE ACA: streamlit re-ejecuta las 3 pestañas
    completas en CADA interaccion (tocar cualquier widget en cualquier
    pestaña reejecuta todo el script de arriba a abajo), y sin cache esto
    arma el .xlsx (openpyxl, lento) de las 3 pestañas en cada click,
    aunque nadie vaya a descargar nada - eso era la causa real de la
    lentitud reportada, sobre todo en Cruce/Generacion (las hojas de
    detalle horario son las mas grandes). Con cache, solo se reconstruye
    cuando cambian los datos de ESA pestaña en particular."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as xw:
        for nombre, df in hojas:
            df.to_excel(xw, sheet_name=nombre[:31], index=False)
    return buffer.getvalue()


def _selector_periodo(meses, key_prefix):
    """Dos selectbox (inicio/fin) sobre los (anio,mes) realmente
    disponibles en la base - evita que se pueda elegir un periodo que ni
    siquiera esta publicado."""
    etiquetas = [f"{a}-{m:02d}" for a, m in meses]
    c1, c2 = st.columns(2)
    ini = c1.selectbox("Desde", etiquetas, index=0, key=f"{key_prefix}_ini")
    fin = c2.selectbox("Hasta", etiquetas, index=len(etiquetas) - 1, key=f"{key_prefix}_fin")
    anio_ini, mes_ini = (int(x) for x in ini.split("-"))
    anio_fin, mes_fin = (int(x) for x in fin.split("-"))
    if (anio_ini, mes_ini) > (anio_fin, mes_fin):
        st.error("El período 'Desde' no puede ser posterior a 'Hasta'.")
        st.stop()
    return anio_ini, mes_ini, anio_fin, mes_fin


def _selector_granularidad(key_prefix):
    return st.radio("Ver por", ["Mes", "Año"], horizontal=True, key=f"{key_prefix}_gran")


def _selector_empresa(dicc, key_prefix):
    """Dropdown de Empresa (desde Empresas_Centrales_COES.xlsx, columna
    'Empresa Usado (actual)'). Se deja FUERA de los st.form de cada pestaña a
    propósito, para que filtrar Central siga siendo instantáneo aunque
    el resto de los campos esperen al botón "Calcular"."""
    empresas_disp = ["(Todas)"] + sorted(dicc["Empresa Usado"].dropna().unique()) + ["(Sin clasificar)"]
    return st.selectbox("Empresa", empresas_disp, key=f"{key_prefix}_empresa")


def _filtrar_centrales_por_empresa(centrales_disp, dicc, empresa):
    """Centrales de `centrales_disp` que pertenecen a `empresa` (según
    _selector_empresa). 'Sin clasificar' agrupa las centrales de la base
    que no están en el Diccionario de Potencia (no se pierden, solo
    quedan fuera de cualquier filtro por empresa)."""
    if empresa == "(Todas)":
        return centrales_disp
    if empresa == "(Sin clasificar)":
        clasificadas = set(dicc.index)
        return [c for c in centrales_disp if c not in clasificadas]
    de_esa_empresa = set(dicc[dicc["Empresa Usado"] == empresa].index)
    return [c for c in centrales_disp if c in de_esa_empresa]


def _agregar_periodo(df, granularidad):
    """Agrega la columna 'Período' (texto) segun la granularidad elegida
    - 'AAAA-MM' si es mensual, 'AAAA' si es anual. df debe tener 'anio' y
    (si es mensual) 'mes'."""
    df = df.copy()
    if granularidad == "Mes":
        df["Período"] = df["anio"].astype(str) + "-" + df["mes"].astype(str).str.zfill(2)
    else:
        df["Período"] = df["anio"].astype(str)
    return df


def _cols_grupo(granularidad):
    return ["anio", "mes"] if granularidad == "Mes" else ["anio"]


def _paleta(n, escala):
    if n <= 1:
        return [COLOR_AZUL if escala == "Blues" else COLOR_ROJO]
    return px.colors.sample_colorscale(escala, [0.35 + 0.55 * i / (n - 1) for i in range(n)])


def _paso_lindo(valor):
    """Redondea 'valor' hacia arriba al siguiente número 'prolijo' (1, 2
    o 5 por una potencia de 10 - ej. 0.2, 0.5, 2, 5, 20, 50, 500...).
    Reemplaza la regla de "múltiplo de 5 fijo" (pedida 2026-09-02): un
    5 fijo funciona para escalas grandes (cientos de GWh) pero no para
    chicas - ej. el perfil mensual-horario de una central da ~1 GWh por
    punto, y forzar el eje a multiplos de 5 dejaba el gráfico casi vacío
    (encontrado por Dani: "dice GWh pero llega hasta el 5, es raro")."""
    if valor <= 0:
        return 1
    exp = math.floor(math.log10(valor))
    for mult in (1, 2, 5):
        candidato = mult * (10 ** exp)
        if candidato >= valor:
            return candidato
    return 10 ** (exp + 1)


def _rango_eje(valores, pct, es_porcentaje=False, divisiones=6):
    """[min, max] para un eje: el máximo es el valor más alto +pct (20%
    para Generación/Precios, 10% para Factor de planta - regla de Dani,
    2026-09-02), redondeado HACIA AFUERA (arriba) a un paso "prolijo"
    (ver _paso_lindo) elegido para que el eje tenga unas `divisiones`
    marcas. El mínimo es 0 si todos los valores son >=0; si hay
    negativos, se extiende con la MISMA regla pero hacia abajo (más
    negativo), con el mismo paso. Devuelve None si no hay datos (deja
    que Plotly autoescale). `es_porcentaje` ya no cambia el cálculo (el
    paso "prolijo" se adapta solo a cualquier escala) - se deja el
    parámetro por compatibilidad con los llamados existentes."""
    valores = pd.Series(valores).dropna()
    if valores.empty:
        return None
    vmax, vmin = valores.max(), valores.min()
    vmax_pad = vmax * (1 + pct)
    paso = _paso_lindo(vmax_pad / divisiones)
    limite_max = math.ceil(vmax_pad / paso) * paso
    if vmin >= 0:
        limite_min = 0
    else:
        limite_min = math.floor((vmin * (1 + pct)) / paso) * paso
    return [limite_min, limite_max]


def _layout_base(fig, titulo_izq, titulo_der=None, formato_der=".0%", categorias=None,
                  rango_izq=None, rango_der=None):
    """categorias: lista ordenada de valores del eje X (ej. Períodos
    ordenados cronológicamente). Sin esto, Plotly ordena las categorías
    por "orden de aparición" mezclando las trazas -> si una barra/central
    tiene menos historia que otra, sus períodos aparecen primero y el
    eje queda desordenado (bug real encontrado 2026-09-02).

    rango_izq/rango_der: [min,max] explícitos (ver _rango_eje) - si no se
    pasan, Plotly autoescala con rangemode="tozero" (comportamiento
    anterior)."""
    xaxis_cfg = dict(type="category")
    if categorias is not None:
        xaxis_cfg["categoryorder"] = "array"
        xaxis_cfg["categoryarray"] = list(categorias)
    eje_izq = dict(title=titulo_izq, gridcolor=COLOR_GRILLA)
    if rango_izq is not None:
        eje_izq["range"] = rango_izq
    else:
        eje_izq["rangemode"] = "tozero"
    layout = dict(
        barmode="group",
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=70, l=10, r=10, b=10),
        yaxis=eje_izq,
        xaxis=xaxis_cfg,
    )
    if titulo_der:
        eje2 = dict(title=titulo_der, overlaying="y", side="right", showgrid=False)
        if rango_der is not None:
            eje2["range"] = rango_der
        else:
            eje2["rangemode"] = "tozero"
        if formato_der:
            eje2["tickformat"] = formato_der
        layout["yaxis2"] = eje2
    fig.update_layout(**layout)
    return fig


# --------------------------------- tabs ---------------------------------

def tab_cmg():
    meses = _meses_disponibles(CMG_BASE_DIR)
    if not meses:
        st.warning("Todavía no hay nada publicado en la base de CMg. Corré Actualizar_Base_Peru.py primero.")
        return

    barras_disp = _valores_distintos(CMG_BASE_DIR, "barra")
    etiquetas_meses = [f"{a}-{m:02d}" for a, m in meses]

    with st.form("cmg_form"):
        barras_sel = st.multiselect("Barra(s)", barras_disp, default=barras_disp[:1])
        moneda_sel = st.radio("Moneda", ["Soles (S/.)", "USD nominal", "USD real"], horizontal=True)
        fecha_base_sel = st.selectbox(
            "Fecha base (solo aplica si elegís 'USD real')",
            etiquetas_meses, index=len(etiquetas_meses) - 1,
            help="A qué mes se llevan los dólares reales. Ej. si elegís diciembre 2025, todos los "
                 "valores quedan expresados en poder adquisitivo de diciembre 2025 (deflactados con "
                 "el CPI de EE.UU. del BLS).",
        )
        granularidad_sel = st.radio("Ver por", ["Mes", "Año"], horizontal=True, key="cmg_gran_sel")
        c1, c2 = st.columns(2)
        ini_sel = c1.selectbox("Desde", etiquetas_meses, index=0, key="cmg_ini_sel")
        fin_sel = c2.selectbox("Hasta", etiquetas_meses, index=len(etiquetas_meses) - 1, key="cmg_fin_sel")
        enviado = st.form_submit_button("Calcular CMg", type="primary")

    if enviado:
        if not barras_sel:
            st.error("Elegí al menos una barra.")
        else:
            anio_ini_v, mes_ini_v = (int(x) for x in ini_sel.split("-"))
            anio_fin_v, mes_fin_v = (int(x) for x in fin_sel.split("-"))
            if (anio_ini_v, mes_ini_v) > (anio_fin_v, mes_fin_v):
                st.error("El período 'Desde' no puede ser posterior a 'Hasta'.")
            else:
                st.session_state["cmg_params"] = dict(
                    barras=barras_sel, moneda_label=moneda_sel, fecha_base_label=fecha_base_sel,
                    granularidad=granularidad_sel, anio_ini=anio_ini_v, mes_ini=mes_ini_v,
                    anio_fin=anio_fin_v, mes_fin=mes_fin_v,
                )

    params = st.session_state.get("cmg_params")
    if not params:
        st.info("Elegí Barra(s), moneda y período, y apretá 'Calcular CMg'.")
        return

    barras = params["barras"]
    moneda_label = params["moneda_label"]
    fecha_base_label = params["fecha_base_label"]
    fecha_base = tuple(int(x) for x in fecha_base_label.split("-"))
    granularidad = params["granularidad"]
    anio_ini, mes_ini = params["anio_ini"], params["mes_ini"]
    anio_fin, mes_fin = params["anio_fin"], params["mes_fin"]

    datos = _cargar_cmg(tuple(barras), anio_ini, mes_ini, anio_fin, mes_fin)
    if datos.empty:
        st.warning("No hay datos de CMg para esa combinación de barra(s) y período.")
        return

    if moneda_label == "USD real":
        cpi_mes = _cpi_por_mes()
        cpi_base = cpi_mes.get(fecha_base)
        if cpi_base is None or pd.isna(cpi_base):
            st.error(f"El BLS todavía no tiene CPI publicado para {fecha_base_label} — elegí otro mes base.")
            return
        datos = datos.copy()
        datos["valor"] = datos["cmg_usd_nominal"] * (cpi_base / datos["cpi"])
        etiqueta_moneda = "USD real/MWh (1)"
        nota_pie = f"(1) Precios reales a {fecha_base_label}"
    else:
        datos = datos.copy()
        datos["valor"] = datos[MONEDAS_CMG_SIMPLES[moneda_label]]
        etiqueta_moneda = moneda_label
        nota_pie = None

    agg = datos.groupby(_cols_grupo(granularidad) + ["barra"], as_index=False)["valor"].mean()
    agg = _agregar_periodo(agg, granularidad)

    etiqueta_gran = "mensual" if granularidad == "Mes" else "anual"
    st.subheader(f"Evolución CMg {etiqueta_gran} ({etiqueta_moneda})")

    colores = _paleta(len(barras), "Reds")
    fig = go.Figure()
    for i, barra in enumerate(barras):
        sub = agg[agg["barra"] == barra].sort_values("Período")
        fig.add_scatter(x=sub["Período"], y=sub["valor"], name=barra, mode="lines+markers",
                         line=dict(color=colores[i], width=3))
    _layout_base(fig, etiqueta_moneda, categorias=sorted(agg["Período"].unique()),
                 rango_izq=_rango_eje(agg["valor"], 0.20))
    st.plotly_chart(fig, width="stretch")
    if nota_pie:
        st.caption(nota_pie)

    tabla_descarga = agg.drop(columns=["anio"] + (["mes"] if "mes" in agg.columns else []))
    tabla_descarga = tabla_descarga.rename(columns={"valor": etiqueta_moneda})
    excel = _excel_descargable([(f"CMg {etiqueta_gran}", tabla_descarga), ("CMg horario", datos)])
    st.download_button("Descargar Excel", data=excel, file_name="CMg_Peru.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def tab_generacion():
    meses = _meses_disponibles(GX_BASE_DIR)
    if not meses:
        st.warning("Todavía no hay nada publicado en la base de Generación. Corré Actualizar_Base_Peru.py primero.")
        return

    centrales_disp = _valores_distintos(GX_BASE_DIR, "central")
    dicc = _cargar_diccionario_potencia()
    etiquetas_meses = [f"{a}-{m:02d}" for a, m in meses]

    # Empresa queda fuera del form para que filtrar Central siga siendo
    # instantáneo (mismo criterio que Cruce, ver comentario ahí).
    empresa = _selector_empresa(dicc, "gx")
    opciones_central = _filtrar_centrales_por_empresa(centrales_disp, dicc, empresa)
    if not opciones_central:
        st.info("Esa empresa no tiene centrales publicadas en la base.")
        return

    with st.form("gx_form"):
        centrales_sel = st.multiselect("Central(es)", opciones_central, default=opciones_central[:1])
        c1, c2 = st.columns(2)
        agrupar_por_sel = c1.radio(
            "Agrupar por", ["Central", "Tecnología", "Total"], horizontal=True, key="gx_agrupar_sel",
            help="'Central': un gráfico por central. 'Tecnología': suma las centrales elegidas por "
                 "tecnología (un gráfico por tecnología). 'Total': suma TODAS las centrales elegidas "
                 "en un solo gráfico.")
        granularidad_sel = c2.radio("Ver por", ["Mes", "Año"], horizontal=True, key="gx_gran_sel")
        c3, c4 = st.columns(2)
        ini_sel = c3.selectbox("Desde", etiquetas_meses, index=0, key="gx_ini_sel")
        fin_sel = c4.selectbox("Hasta", etiquetas_meses, index=len(etiquetas_meses) - 1, key="gx_fin_sel")
        enviado = st.form_submit_button("Calcular generación", type="primary")

    if enviado:
        if not centrales_sel:
            st.error("Elegí al menos una central.")
        else:
            anio_ini_v, mes_ini_v = (int(x) for x in ini_sel.split("-"))
            anio_fin_v, mes_fin_v = (int(x) for x in fin_sel.split("-"))
            if (anio_ini_v, mes_ini_v) > (anio_fin_v, mes_fin_v):
                st.error("El período 'Desde' no puede ser posterior a 'Hasta'.")
            else:
                st.session_state["gx_params"] = dict(
                    centrales=centrales_sel, agrupar_por=agrupar_por_sel, granularidad=granularidad_sel,
                    anio_ini=anio_ini_v, mes_ini=mes_ini_v, anio_fin=anio_fin_v, mes_fin=mes_fin_v,
                )

    params = st.session_state.get("gx_params")
    if not params:
        st.info("Elegí Empresa, Central(es) y el período, y apretá 'Calcular generación'.")
        return

    centrales = params["centrales"]
    agrupar_por = params["agrupar_por"]
    granularidad = params["granularidad"]
    anio_ini, mes_ini = params["anio_ini"], params["mes_ini"]
    anio_fin, mes_fin = params["anio_fin"], params["mes_fin"]

    datos = _cargar_generacion(tuple(centrales), anio_ini, mes_ini, anio_fin, mes_fin)
    if datos.empty:
        st.warning("No hay datos de Generación para esa combinación de central(es) y período.")
        return

    tecnologia_de = dicc["Tecnología"].to_dict()
    potencia_de = dicc["MW"].to_dict()

    sin_potencia = [c for c in centrales if c not in potencia_de]
    if sin_potencia:
        st.caption(f"Sin match en el Diccionario de Potencia (quedan fuera del factor de planta): "
                   f"{', '.join(sin_potencia)}")

    def _grupo_de_central(c):
        if agrupar_por == "Tecnología":
            return tecnologia_de.get(c, "Sin clasificar")
        if agrupar_por == "Total":
            return "Total"
        return c

    datos = datos.copy()
    datos["grupo"] = datos["central"].map(_grupo_de_central)

    potencia_por_grupo = {}
    for c in centrales:
        if c not in potencia_de:
            continue
        grupo = _grupo_de_central(c)
        potencia_por_grupo[grupo] = potencia_por_grupo.get(grupo, 0.0) + potencia_de[c]

    def _tecnologia_de_grupo(grupo):
        if agrupar_por == "Central":
            return tecnologia_de.get(grupo, "Sin clasificar")
        if agrupar_por == "Tecnología":
            return grupo
        tecnologias = sorted({tecnologia_de.get(c, "Sin clasificar") for c in centrales if c in potencia_de})
        return ", ".join(tecnologias) if tecnologias else "s/d"

    def _tabla_agregada(cols_grupo_tiempo):
        g = datos.groupby(cols_grupo_tiempo + ["grupo"])
        gen = g["generacion_mwh"].sum().rename("Generación (MWh)")
        horas = g.size().rename("horas")
        t = pd.concat([gen, horas], axis=1).reset_index()
        t["Generación (GWh)"] = t["Generación (MWh)"] / 1000.0
        t["Potencia (MW)"] = t["grupo"].map(potencia_por_grupo)
        t["Factor de planta"] = t["Generación (MWh)"] / (t["Potencia (MW)"] * t["horas"])
        return t

    tabla_mensual = _tabla_agregada(["anio", "mes"])
    tabla_anual = _tabla_agregada(["anio"])
    tabla = _agregar_periodo(tabla_mensual if granularidad == "Mes" else tabla_anual, granularidad)

    grupos = sorted(tabla["grupo"].unique())
    etiqueta_gran = "mensual" if granularidad == "Mes" else "anual"

    st.markdown("##### Resumen del período (promedios)")
    resumen = pd.DataFrame([{
        "Grupo": grupo,
        "Tecnología": _tecnologia_de_grupo(grupo),
        "Potencia (MW)": potencia_por_grupo.get(grupo),
        "Generación prom. mensual (GWh)": tabla_mensual.loc[tabla_mensual["grupo"] == grupo, "Generación (GWh)"].mean(),
        "Generación prom. anual (GWh)": tabla_anual.loc[tabla_anual["grupo"] == grupo, "Generación (GWh)"].mean(),
        "Factor de planta prom. mensual": tabla_mensual.loc[tabla_mensual["grupo"] == grupo, "Factor de planta"].mean(),
        "Factor de planta prom. anual": tabla_anual.loc[tabla_anual["grupo"] == grupo, "Factor de planta"].mean(),
    } for grupo in grupos])
    st.dataframe(
        resumen.style.format({
            "Potencia (MW)": "{:,.1f}",
            "Generación prom. mensual (GWh)": "{:,.0f}",
            "Generación prom. anual (GWh)": "{:,.0f}",
            "Factor de planta prom. mensual": "{:.0%}",
            "Factor de planta prom. anual": "{:.0%}",
        }, na_rep="s/d"),
        hide_index=True,
    )

    st.subheader(f"Generación {etiqueta_gran} (GWh) y Factor de planta (%)")
    st.caption("Un gráfico independiente por grupo (como en Gráficos usuales.pptx).")
    categorias_orden = sorted(tabla["Período"].unique())
    for grupo in grupos:
        sub = tabla[tabla["grupo"] == grupo].sort_values("Período")
        if agrupar_por == "Tecnología":
            centrales_del_grupo = sorted(c for c in centrales if _grupo_de_central(c) == grupo)
            st.markdown(f"**{grupo}** _(centrales: {', '.join(centrales_del_grupo)})_")
        else:
            st.markdown(f"**{grupo}**")
        fig = go.Figure()
        fig.add_bar(x=sub["Período"], y=sub["Generación (GWh)"], name="Generación (GWh)",
                    marker_color=COLOR_AZUL, yaxis="y",
                    text=sub["Generación (GWh)"], texttemplate="%{text:,.0f}", textposition="inside")
        if sub["Factor de planta"].notna().any():
            fig.add_scatter(x=sub["Período"], y=sub["Factor de planta"], name="Factor de planta",
                             mode="lines+markers", line=dict(color=COLOR_ROJO, width=3), yaxis="y2")
        _layout_base(fig, "Generación (GWh)", "Factor de planta", categorias=categorias_orden,
                     rango_izq=_rango_eje(sub["Generación (GWh)"], 0.20),
                     rango_der=_rango_eje(sub["Factor de planta"], 0.10, es_porcentaje=True))
        st.plotly_chart(fig, width="stretch")

    tabla_descarga = tabla.drop(columns=["anio", "horas"] + (["mes"] if "mes" in tabla.columns else []))
    excel = _excel_descargable([(f"Generación {etiqueta_gran}", tabla_descarga), ("Generación horaria", datos)])
    st.download_button("Descargar Excel", data=excel, file_name="Generacion_Peru.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _agregados_cruce(detalle, cols_grupo):
    """Mismas formulas que Cruce_Ingresos_Peru.py (_agregados_por), sin
    potencia ni factor de planta - eso quedó solo en la pestaña
    Generación (ver pedido de Dani 2026-09-02)."""
    g = detalle.groupby(cols_grupo)
    gen = g["generacion_mwh"].sum().rename("Generación (MWh)")
    iny = g["inyeccion_valorizada_usd"].sum().rename("Inyección valorizada (USD)")
    precio_simple = g["cmg_usd_real"].mean().rename("Precio promedio simple (USD real/MWh)")
    out = pd.concat([gen, iny, precio_simple], axis=1).reset_index()
    out["Precio capturado (USD real/MWh)"] = out["Inyección valorizada (USD)"] / out["Generación (MWh)"]
    out["Ratio de captura"] = (out["Precio capturado (USD real/MWh)"]
                                / out["Precio promedio simple (USD real/MWh)"])
    return out


def tab_cruce():
    meses_cmg = set(_meses_disponibles(CMG_BASE_DIR))
    meses_gx = set(_meses_disponibles(GX_BASE_DIR))
    meses = sorted(meses_cmg & meses_gx)
    if not meses:
        st.warning("Todavía no hay meses en común publicados entre CMg y Generación.")
        return

    centrales_disp = _valores_distintos(GX_BASE_DIR, "central")
    barras_disp = _valores_distintos(CMG_BASE_DIR, "barra")
    dicc = _cargar_diccionario_potencia()
    etiquetas_meses = [f"{a}-{m:02d}" for a, m in meses]

    # Empresa queda FUERA del form (para que filtrar Central siga siendo
    # instantáneo); el resto de los campos van adentro con un botón
    # "Calcular" - así elegir Central/Barra/Fecha base/Período no dispara
    # un recálculo completo (y una descarga de datos) por cada click,
    # solo al final cuando ya está todo elegido (pedido de Dani 2026-09-02
    # por la lentitud de ir cambiando campo por campo).
    empresa = _selector_empresa(dicc, "cruce")
    opciones_central = _filtrar_centrales_por_empresa(centrales_disp, dicc, empresa)
    if not opciones_central:
        st.info("Esa empresa no tiene centrales publicadas en la base.")
        return

    with st.form("cruce_form"):
        c1, c2 = st.columns(2)
        central_sel = c1.selectbox("Central", opciones_central, key="cruce_central_sel")
        barra_sel = c2.selectbox("Barra (CMg)", barras_disp, key="cruce_barra_sel")
        fecha_base_sel = st.selectbox(
            "Fecha base (USD real)", etiquetas_meses, index=len(etiquetas_meses) - 1,
            help="A qué mes se llevan los dólares reales del CMg (deflactados con el CPI del BLS). "
                 "El Ratio de captura NO cambia con esto - solo los niveles de precio en USD/MWh.",
            key="cruce_fecha_base_sel",
        )
        granularidad_sel = st.radio("Ver por", ["Mes", "Año"], horizontal=True, key="cruce_gran_sel")
        c3, c4 = st.columns(2)
        ini_sel = c3.selectbox("Desde", etiquetas_meses, index=0, key="cruce_ini_sel")
        fin_sel = c4.selectbox("Hasta", etiquetas_meses, index=len(etiquetas_meses) - 1, key="cruce_fin_sel")
        enviado = st.form_submit_button("Calcular cruce", type="primary")

    if enviado:
        anio_ini_v, mes_ini_v = (int(x) for x in ini_sel.split("-"))
        anio_fin_v, mes_fin_v = (int(x) for x in fin_sel.split("-"))
        if (anio_ini_v, mes_ini_v) > (anio_fin_v, mes_fin_v):
            st.error("El período 'Desde' no puede ser posterior a 'Hasta'.")
        else:
            st.session_state["cruce_params"] = dict(
                central=central_sel, barra=barra_sel, fecha_base_label=fecha_base_sel,
                granularidad=granularidad_sel, anio_ini=anio_ini_v, mes_ini=mes_ini_v,
                anio_fin=anio_fin_v, mes_fin=mes_fin_v,
            )

    params = st.session_state.get("cruce_params")
    if not params:
        st.info("Elegí Empresa, Central, Barra y el período, y apretá 'Calcular cruce'.")
        return

    central = params["central"]
    barra = params["barra"]
    fecha_base_label = params["fecha_base_label"]
    granularidad = params["granularidad"]
    anio_ini, mes_ini = params["anio_ini"], params["mes_ini"]
    anio_fin, mes_fin = params["anio_fin"], params["mes_fin"]
    fecha_base = tuple(int(x) for x in fecha_base_label.split("-"))

    if central in dicc.index:
        fila = dicc.loc[central]
        st.caption(f"Capacidad instalada: **{fila['MW']:.2f} MW** · Tecnología: **{fila['Tecnología']}**")
    else:
        st.caption("Capacidad instalada: sin match en el Diccionario de Potencia.")

    cmg_h = _cargar_cmg((barra,), anio_ini, mes_ini, anio_fin, mes_fin)[["fecha_hora", "cmg_usd_nominal", "cpi"]]
    gx_h = _cargar_generacion((central,), anio_ini, mes_ini, anio_fin, mes_fin)

    cpi_mes = _cpi_por_mes()
    cpi_base = cpi_mes.get(fecha_base)
    if cpi_base is None or pd.isna(cpi_base):
        st.error(f"El BLS todavía no tiene CPI publicado para {fecha_base_label} — elegí otro mes base.")
        return
    cmg_h = cmg_h.copy()
    cmg_h["cmg_usd_real"] = cmg_h["cmg_usd_nominal"] * (cpi_base / cmg_h["cpi"])

    detalle = gx_h.merge(cmg_h[["fecha_hora", "cmg_usd_real"]], on="fecha_hora", how="inner")
    if detalle.empty:
        st.warning("No hay horas en común entre la generación de esa central y el CMg de esa barra, "
                   "en este período.")
        return
    if len(detalle) < len(gx_h):
        st.info(f"{len(gx_h) - len(detalle):,} hora(s) de generación sin CMg correspondiente "
                f"en la barra (se excluyen del cruce).")

    detalle["inyeccion_valorizada_usd"] = detalle["generacion_mwh"] * detalle["cmg_usd_real"]
    st.caption(f"(1) Precios reales a {fecha_base_label}")

    cols_grupo = _cols_grupo(granularidad)
    tabla = _agregar_periodo(_agregados_cruce(detalle, cols_grupo), granularidad)

    horas_por_periodo = detalle.groupby(cols_grupo).size().rename("horas").reset_index()
    tabla = tabla.merge(horas_por_periodo, on=cols_grupo)
    if central in dicc.index:
        potencia_mw_central = float(dicc.loc[central, "MW"])
        tabla["Factor de planta"] = tabla["Generación (MWh)"] / (potencia_mw_central * tabla["horas"])
    else:
        tabla["Factor de planta"] = None

    etiqueta_gran = "mensual" if granularidad == "Mes" else "anual"
    st.subheader(f"{central} <-> {barra} — {etiqueta_gran}")

    fig = go.Figure()
    fig.add_bar(x=tabla["Período"], y=tabla["Generación (MWh)"] / 1000.0, name="Generación (GWh)",
                marker_color=COLOR_AZUL, yaxis="y")
    if tabla["Factor de planta"].notna().any():
        fig.add_scatter(x=tabla["Período"], y=tabla["Factor de planta"], name="Factor de planta",
                         mode="lines+markers+text", line=dict(color=COLOR_ROJO, width=3),
                         text=tabla["Factor de planta"].map(lambda v: f"{v:.0%}" if pd.notna(v) else ""),
                         textposition="top center", yaxis="y2")
    else:
        st.caption("Sin match en el Diccionario de Potencia — no se puede calcular el factor de planta "
                   "para esta central.")
    categorias_orden = sorted(tabla["Período"].unique())
    _layout_base(fig, "Generación (GWh)", "Factor de planta", categorias=categorias_orden,
                 rango_izq=_rango_eje(tabla["Generación (MWh)"] / 1000.0, 0.20),
                 rango_der=_rango_eje(tabla["Factor de planta"], 0.10, es_porcentaje=True))
    st.plotly_chart(fig, width="stretch")

    tabla_resumen = tabla.set_index("Período").reindex(categorias_orden)
    tabla_resumen_fmt = pd.DataFrame({
        "Precio promedio": tabla_resumen["Precio promedio simple (USD real/MWh)"].map(lambda v: f"{v:,.1f}"),
        "Precio capturado": tabla_resumen["Precio capturado (USD real/MWh)"].map(lambda v: f"{v:,.1f}"),
        "Ratio de captura": tabla_resumen["Ratio de captura"].map(lambda v: f"{v:.1%}"),
    }).T
    st.dataframe(tabla_resumen_fmt)

    st.markdown("##### Precio promedio simple vs. capturado")
    fig2 = go.Figure()
    fig2.add_scatter(x=tabla["Período"], y=tabla["Precio promedio simple (USD real/MWh)"],
                      name="Precio promedio simple", mode="lines+markers", line=dict(color=COLOR_AZUL, width=3))
    fig2.add_scatter(x=tabla["Período"], y=tabla["Precio capturado (USD real/MWh)"],
                      name="Precio capturado", mode="lines+markers", line=dict(color=COLOR_ROJO, width=3))
    rango_precios = _rango_eje(pd.concat([tabla["Precio promedio simple (USD real/MWh)"],
                                           tabla["Precio capturado (USD real/MWh)"]]), 0.20)
    _layout_base(fig2, "USD real/MWh", categorias=categorias_orden, rango_izq=rango_precios)
    st.plotly_chart(fig2, width="stretch")

    st.markdown("##### Perfil horario típico del período")
    st.caption("Misma lógica que el 'perfil horario típico' de Generacion_lookup_Peru.py: para cada "
               "(mes, hora) se promedia entre los años del período (ej. la hora 1 de enero de todos los "
               "años elegidos), y esos 12 valores típicos por hora (uno por mes) se suman para Generación "
               "(GWh, es energía -> se suma) y se promedian para el Precio promedio simple (USD real/MWh, "
               "no se suma un precio).")

    # Paso 1: sumar por dia dentro de cada (anio, mes, hora) - o promediar para precio, que no es aditivo.
    paso1_gen = detalle.groupby(["anio", "mes", "hora"], as_index=False)["generacion_mwh"].sum()
    paso1_precio = detalle.groupby(["anio", "mes", "hora"], as_index=False)["cmg_usd_real"].mean()
    # Paso 2: "tipico" - promedio entre los años del período, por (mes, hora).
    tipico_gen = paso1_gen.groupby(["mes", "hora"], as_index=False)["generacion_mwh"].mean()
    tipico_precio = paso1_precio.groupby(["mes", "hora"], as_index=False)["cmg_usd_real"].mean()
    # Paso 3: juntar los 12 meses en un solo perfil de 24 horas.
    perfil_gen = tipico_gen.groupby("hora", as_index=False)["generacion_mwh"].sum()
    perfil_gen["generacion_gwh"] = perfil_gen["generacion_mwh"] / 1000.0
    perfil_precio = tipico_precio.groupby("hora", as_index=False)["cmg_usd_real"].mean()
    perfil = perfil_gen.merge(perfil_precio, on="hora").sort_values("hora")

    fig3 = go.Figure()
    fig3.add_bar(x=perfil["hora"], y=perfil["generacion_gwh"], name="Generación (GWh)",
                 marker_color=COLOR_AZUL, yaxis="y")
    fig3.add_scatter(x=perfil["hora"], y=perfil["cmg_usd_real"], name="Precio promedio simple (USD real/MWh)",
                      mode="lines+markers", line=dict(color=COLOR_ROJO, width=3), yaxis="y2")
    fig3 = _layout_base(fig3, "Generación (GWh)", "Precio promedio simple (USD real/MWh)", formato_der=None,
                         rango_izq=_rango_eje(perfil["generacion_gwh"], 0.20),
                         rango_der=_rango_eje(perfil["cmg_usd_real"], 0.20))
    fig3.update_layout(xaxis=dict(title="Hora del día", type="linear", dtick=1, tick0=1, range=[0.5, 24.5]))
    st.plotly_chart(fig3, width="stretch")

    total_perfil = perfil["generacion_gwh"].sum()
    pct_base = perfil.loc[perfil["hora"].isin(HORAS_BASE), "generacion_gwh"].sum() / total_perfil
    pct_media = perfil.loc[perfil["hora"].isin(HORAS_MEDIA), "generacion_gwh"].sum() / total_perfil
    pct_punta = perfil.loc[perfil["hora"].isin(HORAS_PUNTA), "generacion_gwh"].sum() / total_perfil
    cb, cm, cp = st.columns(3)
    cb.metric("Base (24-8h)", f"{pct_base:.0%}")
    cm.metric("Media (9-18h)", f"{pct_media:.0%}")
    cp.metric("Punta (19-23h)", f"{pct_punta:.0%}")

    st.markdown("##### Perfil mensual-horario típico del período")
    st.caption("Mismo 'típico' (promedio entre años, por mes-hora) que el gráfico anterior, pero sin "
               "sumar/promediar entre meses: se grafican los 12 meses x 24 horas seguidos. "
               "Avenida: diciembre-abril · Estiaje: mayo-noviembre.")

    MESES_NOMBRE = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
                    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    perfil_mensual = tipico_gen.merge(tipico_precio, on=["mes", "hora"]).sort_values(["mes", "hora"])
    perfil_mensual["generacion_gwh"] = perfil_mensual["generacion_mwh"] / 1000.0
    perfil_mensual["posicion"] = (perfil_mensual["mes"] - 1) * 24 + perfil_mensual["hora"]

    fig4 = go.Figure()
    fig4.add_scatter(x=perfil_mensual["posicion"], y=perfil_mensual["generacion_gwh"], name="Generación (GWh)",
                      mode="lines", fill="tozeroy", line=dict(color=COLOR_AZUL, width=0.5), yaxis="y")
    fig4.add_scatter(x=perfil_mensual["posicion"], y=perfil_mensual["cmg_usd_real"],
                      name="Precio promedio simple (USD real/MWh)",
                      mode="lines", line=dict(color=COLOR_ROJO, width=2), yaxis="y2")
    for mes in range(1, 13):
        if mes % 2 == 0:
            fig4.add_vrect(x0=(mes - 1) * 24 + 0.5, x1=mes * 24 + 0.5,
                            fillcolor=COLOR_GRILLA, opacity=0.6, layer="below", line_width=0)
    fig4 = _layout_base(fig4, "Generación (GWh)", "Precio promedio simple (USD real/MWh)", formato_der=None,
                         rango_izq=_rango_eje(perfil_mensual["generacion_gwh"], 0.20),
                         rango_der=_rango_eje(perfil_mensual["cmg_usd_real"], 0.20))
    fig4.update_layout(xaxis=dict(
        title="Mes (24 horas por mes)",
        type="linear",
        tickmode="array",
        tickvals=[(m - 1) * 24 + 12.5 for m in range(1, 13)],
        ticktext=MESES_NOMBRE,
        range=[0.5, 288.5],
    ))
    st.plotly_chart(fig4, width="stretch")

    tabla_final = tabla.drop(columns=["anio"] + (["mes"] if "mes" in tabla.columns else [])).copy()
    tabla_final["Generación (GWh)"] = tabla_final.pop("Generación (MWh)") / 1000.0
    tabla_final["Inyección valorizada (US$ MM)"] = tabla_final.pop("Inyección valorizada (USD)") / 1_000_000.0
    tabla_final = tabla_final[["Período", "Generación (GWh)", "Inyección valorizada (US$ MM)",
                                "Precio promedio simple (USD real/MWh)", "Precio capturado (USD real/MWh)",
                                "Ratio de captura"]]
    st.dataframe(
        tabla_final.style.format({
            "Generación (GWh)": "{:,.1f}",
            "Inyección valorizada (US$ MM)": "{:,.1f}",
            "Precio promedio simple (USD real/MWh)": "{:,.1f}",
            "Precio capturado (USD real/MWh)": "{:,.1f}",
            "Ratio de captura": "{:.1%}",
        }),
        hide_index=True,
    )

    mensual_completo = _agregar_periodo(_agregados_cruce(detalle, ["anio", "mes"]), "Mes")
    anual_completo = _agregar_periodo(_agregados_cruce(detalle, ["anio"]), "Año")
    excel = _excel_descargable([
        ("Detalle horario", detalle.sort_values("fecha_hora")),
        ("Mensual", mensual_completo.drop(columns=["anio", "mes"])),
        ("Anual", anual_completo.drop(columns=["anio"])),
    ])
    st.download_button("Descargar Excel", data=excel,
                        file_name=f"Cruce_Ingresos_Peru_{central}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------- main ---------------------------------

st.set_page_config(page_title="Información Histórica Perú", layout="wide")
st.title("Información Histórica Perú — Gx y CMg")

tab1, tab2, tab3 = st.tabs(["CMg", "Generación", "Cruce"])
with tab1:
    tab_cmg()
with tab2:
    tab_generacion()
with tab3:
    tab_cruce()
