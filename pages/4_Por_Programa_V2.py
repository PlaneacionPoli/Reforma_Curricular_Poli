"""
pages/4_Por_Programa_V2.py — Fase 2 · V2: Resumen Por Programa
Tabla maestra paginada por período, ficha gráfica por programa seleccionado.
Basado en pages/4_Por_Programa.py, usando el loader del Excel V2 (hoja
"Etapas (2)") y agregando los bloques de Producción de Contenidos y Aulas
Master en la ficha del programa.
"""

import pandas as pd
import streamlit as st

from utils.charts_v2 import render_program_ficha_grafica
from utils.data_loader_v2 import FAC_ABREV_INV, get_etapas_by_programa, load_etapas_data
from utils.f2_components_v2 import (
    apply_current_filters,
    render_f2_header,
    render_f2_sidebar,
    render_filter_bar,
)
from utils.poli_theme import PCT_HIGH, TEXT_MUTED, TEXT_PRIMARY, color_for_pct, phosphor_icon, streamlit_global_css
from utils.master_table_v2 import excel_export_bytes, render_master_table_by_periodo

st.set_page_config(
    page_title="Resumen Por Programa V2 · POLI",
    page_icon=":material/school:",
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
render_f2_header("Resumen Por Programa")
render_filter_bar(
    df_raw, fac_abrev_inv, mods_ops, fac_ops, pers_ops, niveles_ops, key_prefix="programa_v2"
)

df, *_ = apply_current_filters(df_raw, fac_abrev_inv, key_prefix="programa_v2")


def _mini_kpi(label: str, val, color: str) -> str:
    return (
        f'<div style="background:#fff;border:1px solid rgba(15,56,90,.10);border-radius:10px;'
        f'padding:10px 12px;text-align:center">'
        f'<div style="font-size:9px;color:#6a8a9e;text-transform:uppercase;letter-spacing:.4px">{label}</div>'
        f'<div style="font-size:18px;font-weight:800;color:{color};margin-top:3px">{val}</div>'
        f'</div>'
    )


def _n(v) -> float | None:
    """Convierte a float, devolviendo None si el valor es NaN/vacío — evita
    que 'No aplica' (NaN) se muestre como 0 o rompa int()/float() (NaN es
    'truthy', así que `v or 0` no protege contra NaN)."""
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else f


def _render_produccion_ficha(programa: str, df_scope: pd.DataFrame) -> None:
    """Bloques de Producción de Contenidos y Aulas Master para el programa seleccionado."""
    data = get_etapas_by_programa(df_scope, programa)
    prod = data.get("produccion") or {}
    aulas = data.get("aulas_master") or {}

    prod_total = _n(prod.get("total_modulos"))
    prod_proceso = _n(prod.get("modulos_proceso"))
    prod_pct = _n(prod.get("pct_avance"))
    prod_fecha = prod.get("fecha_entrega", "—")

    aulas_conf = _n(aulas.get("modulos_conformidad"))
    aulas_crear = _n(aulas.get("modulos_a_crear"))
    aulas_creadas = _n(aulas.get("total_creadas"))
    aulas_pct = _n(aulas.get("pct_avance"))

    has_prod = any(v not in (None, 0) for v in (prod_total, prod_proceso, prod_pct))
    has_aulas = any(v not in (None, 0) for v in (aulas_conf, aulas_crear, aulas_creadas, aulas_pct))
    if not has_prod and not has_aulas:
        return

    if has_prod:
        st.markdown(
            f'<div class="f2-ficha-section">'
            f'<p class="f2-ficha-section-title">{phosphor_icon("factory", size=14)} Producción de Contenidos</p>'
            f'</div>',
            unsafe_allow_html=True,
        )
        cols = st.columns(4)
        items = [
            ("Módulos a producir", int(prod_total) if prod_total is not None else "No aplica", "#0F385A"),
            ("En proceso", int(prod_proceso) if prod_proceso is not None else "No aplica", "#1FB2DE"),
            ("% Avance Producción", f"{prod_pct:.0f}%" if prod_pct is not None else "No aplica", color_for_pct(prod_pct) if prod_pct is not None else "#94a3b8"),
            ("Fecha proyectada de entrega", prod_fecha, "#6a8a9e"),
        ]
        for col, (label, val, color) in zip(cols, items):
            with col:
                st.markdown(_mini_kpi(label, val, color), unsafe_allow_html=True)

    if has_aulas:
        st.markdown(
            f'<div class="f2-ficha-section">'
            f'<p class="f2-ficha-section-title">{phosphor_icon("factory", size=14)} Aulas Master</p>'
            f'</div>',
            unsafe_allow_html=True,
        )
        cols = st.columns(4)
        items = [
            ("Módulos recibidos a conformidad", int(aulas_conf) if aulas_conf is not None else "No aplica", PCT_HIGH),
            ("Módulos a crear", int(aulas_crear) if aulas_crear is not None else "No aplica", "#1FB2DE"),
            ("Total Aulas Master creadas", int(aulas_creadas) if aulas_creadas is not None else "No aplica", "#0F385A"),
            ("% Avance Aulas Master", f"{aulas_pct:.0f}%" if aulas_pct is not None else "No aplica", color_for_pct(aulas_pct) if aulas_pct is not None else "#94a3b8"),
        ]
        for col, (label, val, color) in zip(cols, items):
            with col:
                st.markdown(_mini_kpi(label, val, color), unsafe_allow_html=True)


if len(df) == 0:
    st.warning("No hay programas con los filtros actuales.")
else:
    st.markdown(
        f'<div style="font-size:18px;font-weight:700;color:{TEXT_PRIMARY};margin:20px 0 12px">'
        f'{phosphor_icon("student", size=22)} Resumen Por Programa</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div style="font-size:14px;font-weight:700;color:{TEXT_PRIMARY};margin:8px 0 8px">'
        f'{phosphor_icon("table", size=16)} Detalle completo por programa</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Listado de todos los programas según los filtros superiores. "
        "Navegue por período en las pestañas. Pulse + en cada formato para ver los campos selector."
    )
    render_master_table_by_periodo(df, key_prefix="programa_tabla_v2")

    st.divider()

    def _build_label(row: pd.Series) -> str:
        nombre = str(row.get("NOMBRE DEL PROGRAMA", "")).strip()
        mod = str(row.get("MODALIDAD", "")).strip()
        sede = str(row.get("SEDE", "")).strip()
        partes = [p for p in [mod, sede] if p and p.lower() not in ("", "none", "nan")]
        return f"{nombre} · {' · '.join(partes)}" if partes else nombre

    df_sel = df.copy().reset_index(drop=True)
    df_sel["_label"] = df_sel.apply(_build_label, axis=1)

    seen: dict[str, int] = {}
    labels_uniq: list[str] = []
    for lbl in df_sel["_label"]:
        if lbl in seen:
            seen[lbl] += 1
            labels_uniq.append(f"{lbl} ({seen[lbl]})")
        else:
            seen[lbl] = 1
            labels_uniq.append(lbl)
    df_sel["_label_uniq"] = labels_uniq

    opciones = sorted(df_sel["_label_uniq"].tolist())
    sel_label = st.selectbox(
        "Buscar programa",
        opciones,
        index=0,
        placeholder="Seleccione un programa",
        key="programa_sel_v2",
    )

    if sel_label:
        st.caption(
            "Ficha del programa: avance por formato, estado de cada campo selector."
        )
        mask = df_sel["_label_uniq"] == sel_label
        df_prog = df_sel[mask].drop(columns=["_label", "_label_uniq"])
        nombre_prog = df_prog["NOMBRE DEL PROGRAMA"].iloc[0] if len(df_prog) else sel_label
        render_program_ficha_grafica(df_prog, nombre_prog)
        _render_produccion_ficha(nombre_prog, df_prog)

    st.markdown("<div style='margin:24px 0 8px'></div>", unsafe_allow_html=True)
    st.download_button(
        "Descargar Excel",
        data=excel_export_bytes(df),
        file_name="reforma_curricular_fase2_v2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
        key="dl_prog_v2",
    )
