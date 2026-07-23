"""
pages/3_Detalle_Etapa_V2.py — Fase 2 · V2: Detalle por Formato
Vista agregada por formato/proceso con drill-down. Basado en
pages/3_Detalle_Etapa.py, usando el loader del Excel V2 (hoja "Etapas (2)").
"""

import pandas as pd
import streamlit as st

from utils.charts_v2 import FORMATO_SHORT, render_etapas_drilldown
from utils.data_loader_v2 import (
    FORMATOS_ORDEN,
    FORMATO_CLR,
    FAC_ABREV_INV,
    get_estadisticas_etapa,
    load_etapas_data,
)
from utils.f2_components_v2 import (
    apply_current_filters,
    render_f2_header,
    render_f2_sidebar,
    render_filter_bar,
)
from utils.poli_theme import TEXT_MUTED, TEXT_PRIMARY, phosphor_icon, streamlit_global_css

st.set_page_config(
    page_title="Detalle por Formato V2 · POLI",
    page_icon=":material/assignment:",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(streamlit_global_css(), unsafe_allow_html=True)

df_raw = load_etapas_data()
fac_abrev_inv = FAC_ABREV_INV
fac_ops = sorted(df_raw["FACULTAD_ABREV"].dropna().unique().tolist()) if "FACULTAD_ABREV" in df_raw.columns else []
mods_ops = sorted(df_raw["MODALIDAD"].dropna().unique().tolist()) if "MODALIDAD" in df_raw.columns else []
pers_ops = sorted(df_raw["PERIODO DE IMPLEMENTACIÓN"].dropna().unique().tolist()) if "PERIODO DE IMPLEMENTACIÓN" in df_raw.columns else []
niveles_ops = [n for n in ["Pregrado", "Posgrado"] if n in df_raw.get("NIVEL_HOMOLOGADO", pd.Series(dtype=str)).values]

render_f2_sidebar()
render_f2_header("Detalle por Formato")
render_filter_bar(
    df_raw, fac_abrev_inv, mods_ops, fac_ops, pers_ops, niveles_ops, key_prefix="detalle_v2"
)

df, *_ = apply_current_filters(df_raw, fac_abrev_inv, key_prefix="detalle_v2")

if len(df) == 0:
    st.warning("No hay programas con los filtros actuales.")
else:
    gen_avg = round(float(df["avance_general_vact"].mean()), 1) if len(df) else 0
    st.markdown(
        f'<div style="font-size:18px;font-weight:700;color:{TEXT_PRIMARY};margin:20px 0 12px">'
        f'{phosphor_icon("clipboard-text", size=22)} Detalle por Formato</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="background:#fff;border:1px solid rgba(15,56,90,.1);border-radius:12px;padding:14px;margin-bottom:16px">'
        f'<span style="font-size:13px;color:{TEXT_MUTED}">{len(df)} programa(s) · Avance general promedio '
        f'<b style="color:{TEXT_PRIMARY}">{gen_avg}%</b> (promedio de los {len(FORMATOS_ORDEN)} formatos)</span></div>',
        unsafe_allow_html=True,
    )

    n_cols = 4
    for row_start in range(0, len(FORMATOS_ORDEN), n_cols):
        row_formatos = FORMATOS_ORDEN[row_start:row_start + n_cols]
        cols = st.columns(n_cols)
        for i, formato in enumerate(row_formatos):
            stats = get_estadisticas_etapa(df, formato)
            clr = FORMATO_CLR.get(formato, "#6e7681")
            with cols[i]:
                st.markdown(
                    f'<div style="background:#fff;border-left:4px solid {clr};border-radius:10px;'
                    f'padding:12px;border:1px solid rgba(15,56,90,.08);margin-bottom:10px">'
                    f'<div style="font-size:9px;color:#6a8a9e;text-transform:uppercase">'
                    f'{FORMATO_SHORT.get(formato, formato)}</div>'
                    f'<div style="font-size:22px;font-weight:800;color:{clr}">{stats["pct_promedio"]}%</div>'
                    f'<div style="font-size:10px;color:{TEXT_MUTED}">{stats["n_programas"]} prog.</div>'
                    f"</div>",
                    unsafe_allow_html=True,
                )

    st.markdown("<div style='margin:20px 0 8px'></div>", unsafe_allow_html=True)
    render_etapas_drilldown(df, key_prefix="detalle_etapa_v2")
