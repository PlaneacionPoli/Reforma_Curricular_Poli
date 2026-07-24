"""
Gráficos Plotly Fase 2 · V2 — drill-down avance por formato/proceso.
Basado en utils/charts_vact.py, adaptado a los 11 formatos de la hoja
"Etapas (2)" (en vez de las 4 etapas canónicas).
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from utils.data_loader_v2 import (
    FORMATOS_ORDEN,
    FORMATO_CLR,
    FORMATO_NUM,
    FORMATO_PCT_COL,
    _ensure_activities_meta,
    get_detalle_etapa,
)
from utils.poli_theme import BRAND_PRIMARY, STATUS_CLR, TEXT_MUTED, TEXT_PRIMARY

# Etiqueta corta por formato (sin numerar) — base para construir FORMATO_SHORT.
_FORMATO_SHORT_BASE = {
    "Aseguramiento de la Calidad": "Aseguramiento",
    "Formato Creación de Programas Banner": "Creación Banner",
    "Proyecciones Académicas": "Proyecciones Académicas",
    "Resultados de Aprendizaje RA": "Result. Aprendizaje",
    "Actas de Homologación": "Actas Homolog.",
    "Syllabus": "Syllabus",
    "Gerencia de Educación Virtual": "Educ. Virtual",
    "Gerencia de Operaciones Academicas": "Op. Académicas",
    "Gerencia de Operaciones Academicas-Banner": "Op. Banner",
    "Convenios y Homologaciones": "Convenios",
    "Dirección de Mercado": "Mercado",
}

# Etiqueta corta numerada ("1. Aseguramiento", "2. Creación Banner", ...).
# Al numerarse aquí, el número se propaga automáticamente a todas las vistas
# que usan FORMATO_SHORT: gráficos, tabla maestra, ficha de programa y export.
FORMATO_SHORT = {
    f: f"{FORMATO_NUM[f]}. {_FORMATO_SHORT_BASE.get(f, f)}" for f in FORMATOS_ORDEN
}

# Nombre completo numerado, para vistas que muestran el nombre sin abreviar.
FORMATO_LABEL = {f: f"{FORMATO_NUM[f]}. {f}" for f in FORMATOS_ORDEN}

STATUS_STACK = [
    ("done", "Finalizado / Aprobado", STATUS_CLR["done"]),
    ("inprog", "En proceso", STATUS_CLR["inprog"]),
    ("devuelto", "Devuelto", STATUS_CLR["devuelto"]),
    ("nostart", "Sin iniciar", STATUS_CLR["nostart"]),
    ("na", "No aplica", STATUS_CLR["na"]),
]

_STATUS_KEYS = [k for k, _, _ in STATUS_STACK]

_PLOTLY_CONFIG = {"displayModeBar": False}


def _formato_promedios(df: pd.DataFrame) -> list[float]:
    out = []
    for formato in FORMATOS_ORDEN:
        col = FORMATO_PCT_COL.get(formato)
        if col and col in df.columns and len(df):
            out.append(round(float(df[col].mean()), 1))
        else:
            out.append(0.0)
    return out


def _resolve_formato_from_point(pt: dict) -> str | None:
    cd = pt.get("customdata")
    if cd is not None:
        if isinstance(cd, str) and cd in FORMATOS_ORDEN:
            return cd
        if isinstance(cd, (list, tuple)) and cd:
            val = cd[0]
            if isinstance(val, (list, tuple)):
                val = val[0]
            if val in FORMATOS_ORDEN:
                return val
    idx = pt.get("point_index", pt.get("pointNumber"))
    if idx is not None and 0 <= int(idx) < len(FORMATOS_ORDEN):
        return FORMATOS_ORDEN[int(idx)]
    return None


def _fig_formatos_level(df: pd.DataFrame) -> go.Figure:
    promedios = _formato_promedios(df)
    labels = [FORMATO_SHORT.get(f, f) for f in FORMATOS_ORDEN]
    colors = [FORMATO_CLR.get(f, "#6e7681") for f in FORMATOS_ORDEN]
    x_max = max(105, max(promedios, default=0) + 8)

    fig = go.Figure(
        go.Bar(
            x=promedios,
            y=labels,
            orientation="h",
            marker=dict(color=colors, line=dict(width=1, color="white")),
            text=[f"{p}%" for p in promedios],
            textposition="outside",
            textfont=dict(size=11, color="#475569"),
            hovertemplate="<b>%{y}</b><br>Avance: %{x}%<extra></extra>",
            customdata=[[f] for f in FORMATOS_ORDEN],
        )
    )
    fig.update_layout(
        height=max(320, len(FORMATOS_ORDEN) * 34 + 40),
        margin=dict(l=4, r=48, t=8, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=[0, x_max], ticksuffix="%", showgrid=True, gridcolor="rgba(15,56,90,0.07)", zeroline=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        font=dict(family="Segoe UI", color=TEXT_PRIMARY),
        showlegend=False,
        clickmode="event",
    )
    return fig


def _activity_y_labels(acts: list[dict]) -> tuple[list[str], list[str]]:
    seen: dict[str, int] = {}
    shorts: list[str] = []
    fulls: list[str] = []
    for m in acts:
        full = m["name"]
        base = _short(full, 42)
        key = base.lower()
        if key in seen:
            seen[key] += 1
            shorts.append(f"{base} ({seen[key]})")
        else:
            seen[key] = 1
            shorts.append(base)
        fulls.append(full)
    return shorts, fulls


def _fig_actividades_level(df: pd.DataFrame, formato: str) -> go.Figure:
    meta = _ensure_activities_meta(df)
    acts = [m for m in meta if m["phase"] == formato]

    if not acts:
        fig = go.Figure()
        fig.update_layout(
            height=160,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            annotations=[dict(text="Este formato usa un % de avance directo (sin campos selector).", showarrow=False, font=dict(size=12, color=TEXT_MUTED))],
        )
        return fig

    # Se conserva el orden original de las columnas del Excel (no se reordena
    # por % de avance), para que coincida con el orden real del formato.
    acts_sorted = [m for m in acts if f"cl_act_{m['idx']}" in df.columns]
    names_short, names_full = _activity_y_labels(acts_sorted)

    fig = go.Figure()
    for cl_key, lbl, clr in STATUS_STACK:
        vals = []
        for m in acts_sorted:
            col = f"cl_act_{m['idx']}"
            vals.append(int((df[col] == cl_key).sum()) if col in df.columns else 0)
        txt_color = "#475569" if cl_key == "na" else "#ffffff"
        fig.add_trace(
            go.Bar(
                name=lbl, y=names_short, x=vals, orientation="h", marker_color=clr,
                customdata=names_full,
                text=[str(v) if v > 0 else "" for v in vals],
                textposition="inside", insidetextanchor="middle", constraintext="none", textangle=0,
                textfont=dict(size=10, color=txt_color),
                hovertemplate="<b>%{customdata}</b><br>" + lbl + ": %{x}<extra></extra>",
            )
        )

    # El título va como caption HTML fuera de la figura (ver render_etapas_drilldown)
    # para que nunca se superponga con la leyenda horizontal.
    fig.update_layout(
        barmode="stack",
        bargap=0.28,
        height=max(280, len(acts_sorted) * 28 + 80),
        margin=dict(l=4, r=24, t=50, b=8),
        uniformtext=dict(minsize=8, mode="hide"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, font=dict(size=10)),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        xaxis=dict(showgrid=True, gridcolor="rgba(15,56,90,0.06)", title="Programas"),
        font=dict(family="Segoe UI"),
    )
    return fig


def _short(text: str, n: int) -> str:
    t = str(text).strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def _render_panel_resumen(df: pd.DataFrame) -> None:
    promedios = _formato_promedios(df)
    avg_global = round(sum(promedios) / len(promedios), 1) if promedios else 0
    rows = ""
    for formato, pct in zip(FORMATOS_ORDEN, promedios):
        clr = FORMATO_CLR.get(formato, "#6e7681")
        short = FORMATO_SHORT.get(formato, formato)
        rows += (
            '<D style="display:flex;align-items:center;gap:8px;margin-bottom:6px">'
            f'<D style="width:8px;height:8px;border-radius:2px;background:{clr};flex-shrink:0"></D>'
            f'<span style="font-size:11px;color:#475569;flex:1">{short}</span>'
            f'<span style="font-size:11px;font-weight:700;color:#0f172a">{pct}%</span>'
            "</D>"
        )
    html = (
        '<D style="background:#fff;border:1px solid rgba(15,56,90,.1);border-radius:12px;padding:14px;max-height:520px;overflow-y:auto">'
        f'<D style="font-size:11px;font-weight:700;color:{TEXT_MUTED};text-transform:uppercase;letter-spacing:.4px">'
        "Resumen general</D>"
        f'<D style="font-size:26px;font-weight:800;color:{TEXT_PRIMARY};margin:4px 0">{avg_global}%</D>'
        f'<D style="font-size:10px;color:{TEXT_MUTED};margin-bottom:10px">'
        f"Promedio de {len(FORMATOS_ORDEN)} formatos · {len(df)} programas</D>"
        f"{rows}</D>"
    )
    st.markdown(html.replace("<D ", "<div ").replace("</D>", "</div>"), unsafe_allow_html=True)


def _render_panel_etapa(df: pd.DataFrame, formato: str | None) -> None:
    if not formato:
        _render_panel_resumen(df)
        return

    det = get_detalle_etapa(df, formato)
    color = FORMATO_CLR.get(formato, "#6e7681")
    n_prog = det.get("n_programas", len(df))
    n_acts = det.get("n_actividades", 0)
    total_cells = det.get("total_act") or 1

    status_rows = [
        ("done", "Finalizado / Aprobado", STATUS_CLR["done"], det.get("done", 0)),
        ("inprog", "En proceso", STATUS_CLR["inprog"], det.get("inprog", 0)),
        ("nostart", "Sin iniciar", STATUS_CLR["nostart"], det.get("nostart", 0)),
        ("na", "No aplica", STATUS_CLR["na"], det.get("na", 0)),
    ]
    if det.get("devuelto", 0):
        status_rows.append(("devuelto", "Devuelto", STATUS_CLR["devuelto"], det["devuelto"]))

    estado_html = ""
    for _, lbl, clr, cnt in status_rows:
        pct = round(cnt / total_cells * 100, 1) if total_cells else 0
        bar_w = min(100, pct)
        estado_html += (
            f'<D style="margin-bottom:7px">'
            f'<D style="display:flex;justify-content:space-between;font-size:10px;margin-bottom:2px">'
            f'<span style="color:#475569">{lbl}</span>'
            f'<span style="font-weight:700;color:#0f172a">{cnt} ({pct}%)</span>'
            f'</D>'
            f'<D style="height:4px;background:#e2e8f0;border-radius:2px">'
            f'<D style="width:{bar_w:.1f}%;height:100%;background:{clr};border-radius:2px"></D>'
            f'</D></D>'
        )

    acts_html = ""
    for act in det.get("actividades", []):
        pct = act["pct_done"]
        bar_w = min(100, pct)
        acts_html += (
            f'<D style="margin-bottom:7px">'
            f'<D style="display:flex;justify-content:space-between;font-size:10px;margin-bottom:2px">'
            f'<span style="color:#475569;max-width:75%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'
            f'{_short(act["nombre"], 30)}</span>'
            f'<span style="font-weight:700;color:#0f172a">{pct}%</span>'
            f'</D>'
            f'<D style="height:4px;background:#e2e8f0;border-radius:2px">'
            f'<D style="width:{bar_w:.1f}%;height:100%;background:{color};border-radius:2px;opacity:.85"></D>'
            f'</D></D>'
        )
    if not acts_html:
        acts_html = f'<p style="font-size:10px;color:{TEXT_MUTED}">Este formato usa un % de avance directo del Excel (sin campos selector).</p>'

    html = (
        f'<D style="background:#fff;border:1px solid rgba(15,56,90,.1);border-radius:12px;'
        f'padding:14px;border-left:4px solid {color};overflow-y:auto;max-height:520px">'
        f'<D style="font-size:12px;font-weight:700;color:{color}">{FORMATO_SHORT.get(formato, formato)}</D>'
        f'<D style="font-size:28px;font-weight:800;color:{color};line-height:1.1">{det["pct_promedio"]}%</D>'
        f'<D style="font-size:10px;color:{TEXT_MUTED};margin:4px 0 12px">'
        f'{n_prog} programas · {n_acts} campos selector</D>'
        f'<D style="font-size:10px;font-weight:700;color:{TEXT_MUTED};text-transform:uppercase;'
        f'letter-spacing:.4px;margin-bottom:6px">Por estado</D>'
        f'{estado_html}'
        f'<D style="font-size:10px;font-weight:700;color:{TEXT_MUTED};text-transform:uppercase;'
        f'letter-spacing:.4px;margin:10px 0 6px">Campos</D>'
        f'{acts_html}'
        f'</D>'
    )
    st.markdown(html.replace("<D ", "<div ").replace("</D>", "</div>"), unsafe_allow_html=True)


def _fig_programa_donut(pct: float) -> go.Figure:
    fig = go.Figure(
        go.Pie(
            values=[pct, max(0, 100 - pct)],
            labels=["Avance", ""],
            hole=0.62,
            marker_colors=[BRAND_PRIMARY if pct >= 30 else "#dc2626", "#e2e8f0"],
            textinfo="none",
            hoverinfo="skip",
        )
    )
    fig.update_layout(
        height=150,
        margin=dict(l=4, r=4, t=4, b=4),
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        annotations=[
            dict(text=f"<b>{pct:.0f}%</b>", x=0.5, y=0.55, font_size=20, showarrow=False),
            dict(text="Avance general", x=0.5, y=0.38, font_size=10, showarrow=False, font_color=TEXT_MUTED),
        ],
    )
    return fig


def _html_esc(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _render_ficha_resumen_block(programa: str, info: dict, gen_pct: float) -> None:
    from utils.poli_theme import MODALIDAD_CLR, badge_html

    fac = info.get("FACULTAD_ABREV", "—")
    mod = info.get("MODALIDAD", "—")
    per = info.get("PERIODO DE IMPLEMENTACIÓN", "—")
    mod_c = MODALIDAD_CLR.get(mod, "#6e7681")
    badges = (
        f'{badge_html(fac, info.get("FACULTAD_COLOR", "#6e7681"))}'
        f'{badge_html(mod, mod_c)}'
        f'{badge_html(per, "#94a3b8")}'
    )

    with st.container(border=True):
        col_info, col_donut = st.columns([2, 1], gap="small", vertical_alignment="center")
        with col_info:
            st.markdown(
                f'<div class="f2-prog-card f2-prog-card-inline">'
                f'<div class="f2-prog-info" style="min-width:0">'
                f"<h3>{_html_esc(programa)}</h3>"
                f'<div class="f2-prog-badges">{badges}</div>'
                f"</div></div>",
                unsafe_allow_html=True,
            )
        with col_donut:
            st.plotly_chart(_fig_programa_donut(gen_pct), width="stretch", config=_PLOTLY_CONFIG)


def _render_ficha_formatos_gantt(data: dict) -> None:
    rows: list[str] = [
        '<div class="f2-ficha-section">',
        '<p class="f2-ficha-section-title">Avance por formato</p>',
        '<div class="f2-gantt f2-ficha-gantt">',
    ]
    for formato in FORMATOS_ORDEN:
        pct_raw = data["etapas"].get(formato, {}).get("pct")
        short = FORMATO_SHORT.get(formato, formato)
        clr = FORMATO_CLR.get(formato, "#6e7681")
        if pct_raw is None:
            # El formato no aplica a este programa (Excel = "No aplica").
            rows.append(
                f'<div class="f2-gantt-row f2-gantt-row-compact">'
                f'<span class="f2-gantt-label">{_html_esc(short)}</span>'
                f'<div class="f2-gantt-track" style="display:flex;align-items:center;padding-left:10px">'
                f'<span style="font-size:11px;color:{TEXT_MUTED};font-style:italic">No aplica</span></div>'
                f'<span class="f2-gantt-pct-out"></span>'
                f"</div>"
            )
            continue
        pct = float(pct_raw)
        w = min(max(pct, 0), 100)
        pct_lbl = f"{pct:.0f}%"
        if w >= 18:
            fill = f'<div class="f2-gantt-fill" style="width:{w:.0f}%;background:{clr}"><span class="f2-gantt-pct">{pct_lbl}</span></div>'
            pct_out = '<span class="f2-gantt-pct-out"></span>'
        else:
            fill = f'<div class="f2-gantt-fill" style="width:{max(w, 2):.0f}%;background:{clr}"></div>'
            pct_out = f'<span class="f2-gantt-pct-out">{pct_lbl}</span>'
        rows.append(
            f'<div class="f2-gantt-row f2-gantt-row-compact">'
            f'<span class="f2-gantt-label">{_html_esc(short)}</span>'
            f'<div class="f2-gantt-track">{fill}</div>'
            f"{pct_out}"
            f"</div>"
        )
    rows.append("</div></div>")
    st.markdown("".join(rows), unsafe_allow_html=True)


def render_program_ficha_grafica(df: pd.DataFrame, programa: str) -> None:
    """Ficha del programa con donut, barras por formato y desglose de campos selector."""
    from utils.data_loader_v2 import get_etapas_by_programa
    from utils.program_ficha_detalle_v2 import render_program_actividades_detalle

    data = get_etapas_by_programa(df, programa)
    info = data.get("info", {})
    gen = float(data.get("avance_general", 0) or 0)

    _render_ficha_resumen_block(programa, info, gen)
    _render_ficha_formatos_gantt(data)
    render_program_actividades_detalle(data)


def render_etapas_drilldown(df: pd.DataFrame, *, key_prefix: str = "formatos") -> None:
    """Gráfica de barras con drill-down formato → campos selector."""
    if len(df) == 0 or FORMATO_PCT_COL[FORMATOS_ORDEN[0]] not in df.columns:
        st.info("Sin datos para mostrar avance por formato.")
        return

    sk = f"{key_prefix}_drill"
    if sk not in st.session_state:
        st.session_state[sk] = None

    formato_sel = st.session_state[sk]
    col_chart, col_panel = st.columns([2.2, 1])

    with col_chart:
        if formato_sel:
            if st.button("Volver a formatos", key=f"{key_prefix}_back", icon=":material/arrow_back:"):
                st.session_state[sk] = None
                st.rerun()
            st.markdown(
                f'<div style="font-size:13px;font-weight:700;color:{TEXT_PRIMARY};margin:10px 0 2px">'
                f'Campos selector — {FORMATO_SHORT.get(formato_sel, formato_sel)}</div>',
                unsafe_allow_html=True,
            )
            fig = _fig_actividades_level(df, formato_sel)
        else:
            fig = _fig_formatos_level(df)
            if all(p == 0 for p in _formato_promedios(df)):
                st.caption("Los porcentajes de avance por formato son 0 con los filtros actuales.")

        event = st.plotly_chart(
            fig, width="stretch", on_select="rerun", selection_mode="points",
            key=f"{key_prefix}_plotly", config=_PLOTLY_CONFIG,
        )

        if not formato_sel and event and getattr(event, "selection", None):
            points = getattr(event.selection, "points", None) or []
            if points:
                picked = _resolve_formato_from_point(points[0])
                if picked:
                    st.session_state[sk] = picked
                    st.rerun()

    with col_panel:
        _render_panel_etapa(df, formato_sel)
