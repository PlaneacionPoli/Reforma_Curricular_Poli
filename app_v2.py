"""
app_v2.py — Dashboard Fase 2 · V2: Reforma Curricular por Etapas
Fuente: hoja Etapas · CONTROL MAESTRO DE REFORMA CURRICULAR V2.xlsx
Página: Resumen Ejecutivo

Basado en app_act.py (misma lógica visual y de filtros), adaptado al nuevo
Excel V2: agrega columnas de info (SNIES, Compartidas/Institucional) y una
nueva sección de Producción de Contenidos (módulos a producir/entregados/
en proceso, % de avance, fecha proyectada de entrega).
"""

import html as html_lib
from pathlib import Path

import pandas as pd
import streamlit as st

from utils.data_loader_v2 import (
    FORMATOS_ORDEN,
    FORMATO_CLR,
    FORMATO_PCT_COL,
    FAC_ABREV_INV,
    STATUS_LABEL,
    _ensure_activities_meta,
    get_estadisticas_produccion,
    get_estadisticas_aulas_master,
    load_etapas_data,
)
from utils.charts_v2 import render_etapas_drilldown
from utils.f2_components_v2 import (
    apply_current_filters as f2_apply_filters,
    render_f2_header,
    render_f2_sidebar,
    render_filter_bar as f2_render_filter_bar,
)
from utils.poli_theme import (
    BG_ROW,
    BG_ROW_ALT,
    BG_TABLE,
    BORDER_ROW,
    BORDER_TABLE,
    BRAND_ACCENT,
    BRAND_HIGHLIGHT,
    BRAND_PRIMARY,
    BRAND_SECONDARY,
    MODALIDAD_CLR,
    PERIODO_CLR,
    TEXT_LIGHT,
    TEXT_MUTED,
    TEXT_NA,
    TEXT_PRIMARY,
    TEXT_SUBTLE,
    badge_html,
    p_bar_html,
    status_icon_html,
    streamlit_global_css,
    FACULTAD_CLR,
    ETAPA_CLR,
    STATUS_CLR,
    PCT_HIGH,
    PCT_CRITICAL,
    PCT_LOW,
    color_for_pct,
    phosphor_icon,
    phosphor_icon_kpi,
    phosphor_icon_nav,
    rgba_hex,
    PHOSPHOR_ICONS,
)

# Nivel de formación detallado (definido aquí para compatibilidad con despliegues sin poli_theme actualizado)
NIVEL_ORDEN = ["Maestría", "Especialización", "Profesional", "Tecnológico", "Técnico"]
NIVEL_CLR = {
    "Especialización": "#2563eb",
    "Maestría": "#7c3aed",
    "Profesional": "#059669",
    "Tecnológico": BRAND_ACCENT,
    "Técnico": "#0891b2",
}
try:
    from utils.poli_theme import NIVEL_CLR as _NIVEL_CLR_THEME, NIVEL_ORDEN as _NIVEL_ORDEN_THEME

    NIVEL_CLR = _NIVEL_CLR_THEME
    NIVEL_ORDEN = _NIVEL_ORDEN_THEME
except ImportError:
    pass

st.set_page_config(
    page_title="Reforma Curricular V2 · POLI",
    page_icon=":material/school:",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _esc(s):
    return html_lib.escape(str(s) if s is not None else "—")


st.markdown(streamlit_global_css(), unsafe_allow_html=True)

# ── Datos ─────────────────────────────────────────────────────────────────────
df_raw = load_etapas_data()

fac_abrev_inv = FAC_ABREV_INV
fac_ops = sorted(df_raw["FACULTAD_ABREV"].dropna().unique().tolist()) if "FACULTAD_ABREV" in df_raw.columns else []
mods_ops = sorted(df_raw["MODALIDAD"].dropna().unique().tolist()) if "MODALIDAD" in df_raw.columns else []
pers_ops = sorted(df_raw["PERIODO DE IMPLEMENTACIÓN"].dropna().unique().tolist()) if "PERIODO DE IMPLEMENTACIÓN" in df_raw.columns else []
niveles_ops = [n for n in ["Pregrado", "Posgrado"] if n in df_raw.get("NIVEL_HOMOLOGADO", pd.Series(dtype=str)).values]


# ══════════════════════════════════════════════════════════════════════════════
# RESUMEN EJECUTIVO
# ══════════════════════════════════════════════════════════════════════════════

def _arc(pct, color, r=22, sz=56):
    circ = 2 * 3.14159 * r
    dash = circ * min(pct, 100) / 100
    gap = circ - dash
    c = sz // 2
    return (
        f'<svg width="{sz}" height="{sz}" viewBox="0 0 {sz} {sz}">'
        f'<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="rgba(15,56,90,0.10)" stroke-width="5"/>'
        f'<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="{color}" stroke-width="5"'
        f' stroke-dasharray="{dash:.2f} {gap:.2f}" stroke-linecap="round"'
        f' transform="rotate(-90 {c} {c})"/>'
        f'</svg>'
    )


def _kpi_icon_badge(icon_name, color, sz=56):
    """Insignia con icono (sin anillo de %) para KPIs que muestran un conteo,
    no un porcentaje — un anillo de avance ahí no representa nada real."""
    bg = rgba_hex(color, 0.14)
    return (
        f'<div style="width:{sz}px;height:{sz}px;border-radius:50%;background:{bg};'
        f'display:flex;align-items:center;justify-content:center;flex-shrink:0">'
        f'{phosphor_icon(icon_name, size=24, color=color)}'
        f'</div>'
    )


def _kpi_card(label, val, sub, color, pct_bar=None, icon=None):
    if pct_bar is None:
        ring = _kpi_icon_badge(icon or "circle", color)
    else:
        ring = _arc(min(pct_bar * 100, 100), color)
    return (
        f'<div style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);'
        f'border-left:4px solid {color};border-radius:12px;'
        f'padding:14px 16px;display:flex;align-items:center;gap:12px;min-height:84px;'
        f'box-shadow:0 2px 8px rgba(15,56,90,0.07)">'
        f'<div style="flex-shrink:0">{ring}</div>'
        f'<div style="flex:1;min-width:0">'
        f'<div style="font-size:10px;color:#6a8a9e;text-transform:uppercase;'
        f'letter-spacing:.5px;margin-bottom:3px">{label}</div>'
        f'<div style="font-size:26px;font-weight:700;color:{color};line-height:1.1">{val}</div>'
        f'<div style="font-size:10px;color:#8aabb0;margin-top:2px">{sub}</div>'
        f'</div></div>'
    )


def _render_kpis(df: pd.DataFrame):
    n = len(df)
    avg_general = round(df["avance_general_vact"].mean(), 1) if n > 0 else 0
    presencial = int((df["MODALIDAD"] == "Presencial").sum()) if n > 0 else 0
    virtual = int((df["MODALIDAD"] == "Virtual").sum()) if n > 0 else 0
    hibrido = int((df["MODALIDAD"] == "Híbrido").sum()) if n > 0 else 0
    pct_presencial = round(presencial / n * 100, 1) if n > 0 else 0
    pct_virtual = round(virtual / n * 100, 1) if n > 0 else 0
    pct_hibrido = round(hibrido / n * 100, 1) if n > 0 else 0

    kpis = [
        ("Total Programas", str(n), "Programas activos", "#0F385A", None, "books"),
        ("Avance Promedio", f"{avg_general}%", "Avance general", color_for_pct(avg_general), avg_general / 100, None),
        ("Presencial", f"{pct_presencial}%", f"{presencial} programas", "#2980B9", pct_presencial / 100, None),
        ("Virtual", f"{pct_virtual}%", f"{virtual} programas", "#1FB2DE", pct_virtual / 100, None),
        ("Híbrido", f"{pct_hibrido}%", f"{hibrido} programas", "#FBAF17", pct_hibrido / 100, None),
    ]

    cols = st.columns(5)
    for i, (label, val, sub, color, pct_bar, icon) in enumerate(kpis):
        with cols[i]:
            st.markdown(_kpi_card(label, val, sub, color, pct_bar, icon), unsafe_allow_html=True)


_PROD_ACCENT = "#1FB2DE"
_AULAS_ACCENT = "#EC0677"


def _produccion_card_html(
    icon_name: str, title: str, subtitle: str, accent: str,
    activo: float, activo_lbl: str, pendiente: float, pendiente_lbl: str,
    total: float, pct: float, pct_sub: str,
) -> str:
    """Tarjeta compacta: barra de composición (activo/pendiente sobre el
    total) + cifras + % de avance, todo correlacionado en un mismo bloque
    visual. Pensada para ir lado a lado con otra del mismo tipo en una fila."""
    w_activo = round(activo / total * 100, 1) if total else 0
    w_pend = max(0.0, 100 - w_activo)
    bar = (
        f'<div style="height:9px;border-radius:5px;overflow:hidden;background:#eef3f8;display:flex;margin-top:8px">'
        f'<div style="width:{w_activo}%;background:{accent}"></div>'
        f'<div style="width:{w_pend}%;background:#d8e2ea"></div>'
        f'</div>'
    )
    return (
        f'<div style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);border-left:4px solid {accent};'
        f'border-radius:12px;padding:14px 16px;box-shadow:0 2px 8px rgba(15,56,90,0.07);height:100%">'
        f'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:10px">'
        f'<div style="min-width:0">'
        f'<div style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'
        f'{phosphor_icon(icon_name, size=14, color=accent)} {title}</div>'
        f'<div style="font-size:10px;color:{TEXT_MUTED};margin-top:1px">{subtitle}</div>'
        f'</div>'
        f'<div style="text-align:right;flex-shrink:0">'
        f'<div style="font-size:22px;font-weight:800;color:{color_for_pct(pct)};line-height:1">{pct}%</div>'
        f'</div></div>'
        f'{bar}'
        f'<div style="display:flex;justify-content:space-between;margin-top:6px;font-size:10px;color:{TEXT_SUBTLE};gap:6px">'
        f'<span><b style="color:{accent}">{int(activo)}</b> {activo_lbl}</span>'
        f'<span><b style="color:#8aabb0">{int(pendiente)}</b> {pendiente_lbl}</span>'
        f'<span><b style="color:{TEXT_PRIMARY}">{int(total)}</b> total</span>'
        f'</div>'
        f'<div style="font-size:9px;color:{TEXT_MUTED};margin-top:5px">{pct_sub}</div>'
        f'</div>'
    )


def _render_produccion_kpis(df: pd.DataFrame):
    """Sección V2: tarjetas de Producción de Contenidos y Aulas Master — dos
    formatos distintos (Gerencia de Educación Virtual / Gerencia de
    Operaciones Academicas), lado a lado en la misma fila."""
    prod = get_estadisticas_produccion(df)
    aulas = get_estadisticas_aulas_master(df)

    has_prod = prod["total_modulos"] or prod["modulos_proceso"] or prod["n_programas_con_dato"]
    has_aulas = aulas["total_creadas"] or aulas["modulos_a_crear"] or aulas["n_programas_con_dato"]
    if not has_prod and not has_aulas:
        return

    st.markdown(
        f'<div style="font-size:14px;font-weight:700;color:{TEXT_PRIMARY};margin:8px 0 4px">'
        f'{phosphor_icon("factory", size=16)} Producción de Contenidos y Aulas Master</div>'
        f'<div style="font-size:11px;color:{TEXT_MUTED};margin-bottom:8px">'
        "Nuevo en el Control Maestro V2.</div>",
        unsafe_allow_html=True,
    )

    cards = []
    if has_prod:
        total = prod["total_modulos"]
        proceso = prod["modulos_proceso"]
        pendiente = max(0.0, total - proceso)
        cards.append(_produccion_card_html(
            "monitor-play", "Producción de Contenidos", "Gerencia de Educación Virtual", _PROD_ACCENT,
            proceso, "en proceso", pendiente, "pendientes", total,
            prod["pct_avance_promedio"], f"avance prom. · {prod['n_programas_con_dato']} programas",
        ))
    if has_aulas:
        creadas = aulas["total_creadas"]
        a_crear = aulas["modulos_a_crear"]
        total_aulas = creadas + a_crear
        cards.append(_produccion_card_html(
            "chalkboard-teacher", "Aulas Master", "Gerencia de Operaciones Academicas", _AULAS_ACCENT,
            creadas, "creadas", a_crear, "por crear", total_aulas,
            aulas["pct_avance_promedio"], f"avance prom. · {aulas['n_programas_con_dato']} programas",
        ))

    cols = st.columns(len(cards))
    for col, card_html in zip(cols, cards):
        with col:
            st.markdown(card_html, unsafe_allow_html=True)


def _render_chart_facultad(df: pd.DataFrame):
    if "FACULTAD_ABREV" not in df.columns:
        return

    facs = df["FACULTAD_ABREV"].unique()
    colors_map = {"FSCC": "#EC0677", "FIDI": "#1FB2DE", "FNGS": "#A6CE38"}
    max_val = 100
    bar_h = 32
    gap = 12
    svg_w = 480
    svg_h = len(facs) * (bar_h + gap) + 20

    bars = ""
    for i, fac in enumerate(sorted(facs)):
        fac_df = df[df["FACULTAD_ABREV"] == fac]
        avg = round(fac_df["avance_general_vact"].mean(), 1) if len(fac_df) > 0 else 0
        width = (avg / max_val) * (svg_w - 80)
        color = colors_map.get(fac, "#6e7681")
        y = i * (bar_h + gap)
        fill_color = "#fff" if avg > 30 else "#1e293b"
        bars += (
            f'<g transform="translate(0,{y})">'
            f'<rect x="0" y="0" width="{svg_w}" height="{bar_h}" rx="6" fill="rgba(0,0,0,0.03)"/>'
            f'<rect x="0" y="0" width="{max(4, width)}" height="{bar_h}" rx="6" fill="{color}" opacity="0.85"/>'
            f'<text x="8" y="{bar_h/2+1}" dominant-baseline="middle" fill="{fill_color}" font-family="Segoe UI,sans-serif" font-size="11" font-weight="600">{fac}</text>'
            f'<text x="{svg_w-8}" y="{bar_h/2+1}" dominant-baseline="middle" text-anchor="end" fill="#475569" font-family="Segoe UI,sans-serif" font-size="11" font-weight="700">{avg}%</text>'
            f'</g>'
        )

    st.markdown(
        f'<div style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(15,56,90,0.07)">'
        f'<div style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};margin-bottom:12px">Avance por Facultad</div>'
        f'<svg viewBox="0 0 {svg_w} {svg_h}">{bars}</svg>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _render_chart_nivel_detalle(df: pd.DataFrame) -> None:
    """Barras horizontales por nivel de formación (columna NIVEL)."""
    if "NIVEL" not in df.columns:
        return

    raw = df["NIVEL"].dropna().astype(str).str.strip()
    raw = raw[raw != ""]
    if len(raw) == 0:
        return
    vc = raw.value_counts()
    items = [(n, int(vc[n])) for n in NIVEL_ORDEN if n in vc.index]
    items += [(n, int(vc[n])) for n in vc.index if n not in NIVEL_ORDEN]
    if not items:
        return
    max_c = max(c for _, c in items)

    rows = ""
    for nivel, count in items:
        color = NIVEL_CLR.get(nivel, "#6e7681")
        pct = (count / max_c * 100) if max_c else 0
        fill_color = "#fff" if pct >= 28 else "#1e293b"
        rows += (
            f'<TAG style="display:flex;align-items:center;gap:12px;margin-bottom:10px">'
            f'<TAG style="flex:1;position:relative;height:26px;background:rgba(15,56,90,0.05);border-radius:6px;overflow:hidden">'
            f'<TAG style="width:{max(5, pct):.1f}%;height:100%;background:{color};border-radius:6px"></TAG>'
            f'<span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);'
            f'font-size:11px;font-weight:600;color:{fill_color}">{nivel}</span>'
            f"</TAG>"
            f'<span style="font-size:12px;font-weight:700;color:#0f172a;min-width:24px;text-align:right">{count}</span>'
            f"</TAG>"
        )
    rows = rows.replace("TAG", "div")

    card = (
        f'<TAG style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);border-radius:12px;padding:16px;'
        f'box-shadow:0 2px 8px rgba(15,56,90,0.07)">'
        f'<TAG style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};margin-bottom:2px">'
        f"Distribución por Nivel Académico</TAG>"
        f'<TAG style="font-size:11px;color:#94a3b8;margin-bottom:14px">Programas por nivel de formación</TAG>'
        f"{rows}</TAG>"
    )
    st.markdown(card.replace("TAG", "div"), unsafe_allow_html=True)


def _render_rankings(df: pd.DataFrame):
    # Top programas
    top = df.nlargest(8, "avance_general_vact")
    rank_colors = [PCT_HIGH, BRAND_ACCENT, PCT_LOW]

    top_html = ""
    for i, (_, row) in enumerate(top.iterrows()):
        color = rank_colors[i] if i < 3 else "#94a3b8"
        nombre = row.get("NOMBRE DEL PROGRAMA", "—")[:35]
        fac = row.get("FACULTAD_ABREV", "—")
        mod = row.get("MODALIDAD", "—")
        pct = int(row.get("avance_general_vact", 0))

        top_html += (
            f'<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:#f8fafc;border-radius:6px;margin-bottom:4px">'
            f'<span style="font-family:Segoe UI,sans-serif;font-size:12px;font-weight:800;width:20px;text-align:center;color:{color}">{i+1}</span>'
            f'<div style="flex:1;min-width:0"><div style="font-size:12px;font-weight:600;color:{TEXT_PRIMARY};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{nombre}</div>'
            f'<div style="font-size:10px;color:{TEXT_MUTED}">{fac} · {mod}</div></div>'
            f'<span style="font-family:Segoe UI,sans-serif;font-size:13px;font-weight:800;color:{color}">{pct}%</span></div>'
        )

    # Programas críticos
    criticos = df.nsmallest(8, "avance_general_vact")
    criticos_html = ""
    for i, (_, row) in enumerate(criticos.iterrows()):
        nombre = row.get("NOMBRE DEL PROGRAMA", "—")[:35]
        fac = row.get("FACULTAD_ABREV", "—")
        mod = row.get("MODALIDAD", "—")
        pct = int(row.get("avance_general_vact", 0))

        criticos_html += (
            f'<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:#fef2f2;border-radius:6px;margin-bottom:4px">'
            f'<span style="font-family:Segoe UI,sans-serif;font-size:12px;font-weight:800;width:20px;text-align:center;color:{PCT_CRITICAL}">{i+1}</span>'
            f'<div style="flex:1;min-width:0"><div style="font-size:12px;font-weight:600;color:{TEXT_PRIMARY};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{nombre}</div>'
            f'<div style="font-size:10px;color:{TEXT_MUTED}">{fac} · {mod}</div></div>'
            f'<span style="font-family:Segoe UI,sans-serif;font-size:13px;font-weight:800;color:{PCT_CRITICAL}">{pct}%</span></div>'
        )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            f'<div style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(15,56,90,0.07)">'
            f'<div style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};margin-bottom:12px">{phosphor_icon("trophy", size=18)} Top Programas Destacados</div>'
            f'<div style="font-size:11px;color:{TEXT_MUTED};margin-bottom:10px">Mayor avance general</div>'
            f'{top_html}</div>',
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            f'<div style="background:#FFFFFF;border:1px solid rgba(15,56,90,0.10);border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(15,56,90,0.07)">'
            f'<div style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};margin-bottom:12px">{phosphor_icon("warning-circle", size=18, color=PCT_CRITICAL)} Programas Críticos</div>'
            f'<div style="font-size:11px;color:{TEXT_MUTED};margin-bottom:10px">Avance general menor al 20%</div>'
            f'{criticos_html}</div>',
            unsafe_allow_html=True,
        )


render_f2_sidebar()

render_f2_header("Resumen Ejecutivo")

f2_render_filter_bar(
    df_raw, fac_abrev_inv, mods_ops, fac_ops, pers_ops, niveles_ops, key_prefix="ejecutivo_v2"
)

df, *_ = f2_apply_filters(df_raw, fac_abrev_inv, key_prefix="ejecutivo_v2")
n = len(df)

if n == 0:
    st.warning("No hay programas que coincidan con los filtros seleccionados.")
else:
    st.markdown(
        f'<div style="font-size:18px;font-weight:700;color:{TEXT_PRIMARY};margin:20px 0 12px">{phosphor_icon("chart-bar", size=18)} Resumen Ejecutivo</div>',
        unsafe_allow_html=True,
    )

    _render_kpis(df)

    st.markdown("<div style='margin-bottom:24px'></div>", unsafe_allow_html=True)

    _render_produccion_kpis(df)

    st.markdown("<div style='margin-bottom:24px'></div>", unsafe_allow_html=True)

    col_chart1, col_chart2 = st.columns([1, 1])
    with col_chart1:
        _render_chart_nivel_detalle(df)
    with col_chart2:
        _render_chart_facultad(df)

    st.markdown("<div style='margin-bottom:24px'></div>", unsafe_allow_html=True)

    st.markdown(
        f'<div style="font-size:14px;font-weight:700;color:{TEXT_PRIMARY};margin:8px 0 4px">'
        f'{phosphor_icon("chart-bar-horizontal", size=16)} Avance por Etapa</div>'
        f'<div style="font-size:11px;color:{TEXT_MUTED};margin-bottom:8px">'
        "Seleccione una barra del gráfico para ver el detalle de actividades.</div>",
        unsafe_allow_html=True,
    )
    render_etapas_drilldown(df, key_prefix="resumen_v2")
