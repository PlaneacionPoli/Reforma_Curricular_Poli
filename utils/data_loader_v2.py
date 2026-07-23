"""
utils/data_loader_v2.py
Carga y procesa la hoja 'Etapas (2)' del archivo V2 (Fase 2 — versión ampliada
con checklist de aprobaciones por formato/proceso).
Fuente: data/raw/CONTROL MAESTRO DE REFORMA CURRICULAR V2.xlsx

Modelo de cálculo (reemplaza el sistema de 4 etapas canónicas de la versión
anterior):
- Los programas y actividades se agrupan por "formato/proceso" (fila 9 del
  Excel): Aseguramiento de la Calidad, Formato Creación de Programas Banner,
  Proyecciones Académicas, Resultados de Aprendizaje RA, Actas de
  Homologación, Syllabus, Gerencia de Educación Virtual, Gerencia de
  Operaciones Academicas, Gerencia de Operaciones Academicas-Banner,
  Convenios y Homologaciones, Dirección de Mercado.
- Cada formato tiene "campos selectores" (checkboxes True/False o estados
  tipo Finalizado/En proceso/Sin Iniciar) que se promedian para obtener el
  % de avance del formato. Los campos puramente informativos (conteos,
  fechas, texto categórico como "Tipo de trámite") se excluyen del promedio.
- Dos formatos usan un campo de % explícito en vez de promediar checkboxes:
  Gerencia de Educación Virtual ("% de avance Producción") y Gerencia de
  Operaciones Academicas ("%Avance de Aulas Master").
- El avance general del programa es el promedio simple de los % de los 11
  formatos.
"""

from __future__ import annotations

import re
import warnings
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

ROOT = Path(__file__).parent.parent
DATA_PATH = ROOT / "data" / "raw" / "CONTROL MAESTRO DE REFORMA CURRICULAR V2.xlsx"
SHEET_NAME = "Etapas (2)"
DATA_SOURCE_NAME = "Control Maestro Reforma Curricular V2"
TZ_BOGOTA = ZoneInfo("America/Bogota")


def _parse_office_datetime(value: str) -> datetime | None:
    """Parsea fecha ISO de metadatos Office (UTC) a datetime con zona."""
    if not value or not str(value).strip():
        return None
    s = str(value).strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("UTC"))
    return dt


def _excel_core_modified(path: Path) -> datetime | None:
    """Última modificación guardada en Excel (docProps/core.xml), no mtime del SO."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "docProps/core.xml" not in zf.namelist():
                return None
            root = ET.fromstring(zf.read("docProps/core.xml"))
        for el in root.iter():
            if el.tag.endswith("}modified") and el.text:
                parsed = _parse_office_datetime(el.text)
                if parsed is not None:
                    return parsed
        return None
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return None


def _data_file_last_modified() -> datetime | None:
    """Última modificación del archivo: la más reciente entre disco y metadatos Excel."""
    if not DATA_PATH.is_file():
        return None
    mtime_dt = datetime.fromtimestamp(DATA_PATH.stat().st_mtime, tz=TZ_BOGOTA)
    excel_dt = _excel_core_modified(DATA_PATH)
    if excel_dt is None:
        return mtime_dt
    excel_dt = excel_dt.astimezone(TZ_BOGOTA)
    return max(mtime_dt, excel_dt)


def get_raw_data_updated_label() -> str:
    """Fecha/hora de última modificación del Excel, horario Bogotá (lectura en cada llamada)."""
    dt = _data_file_last_modified()
    if dt is None:
        return "Actualizado: —"
    return dt.strftime("Actualizado: %d/%m/%Y %H:%M")


def _data_file_cache_key() -> tuple[int, int, int]:
    """Clave para invalidar caché: mtime + tamaño del Excel y mtime del loader."""
    excel_mtime, excel_size = (0, 0)
    if DATA_PATH.is_file():
        s = DATA_PATH.stat()
        excel_mtime, excel_size = int(s.st_mtime), s.st_size
    loader_mtime = int(Path(__file__).stat().st_mtime) if Path(__file__).is_file() else 0
    return (excel_mtime, excel_size, loader_mtime)


# ── Estructura de filas (hoja "Etapas (2)") ─────────────────────────────────
PHASE_ROW = 7        # fila 8 Excel: fases generales (Alistamiento/Diseño/...)
GROUP_ROW = 8         # fila 9 Excel: formato / proceso (agrupación real)
HEADER_ROW = 9        # fila 10 Excel: nombre de campo
SUBLABEL_ROW = 10     # fila 11 Excel: sigla de aprobador (ED/AD/AC/AO...), informativa
DATA_START_ROW = 11   # fila 12 Excel: primer programa

INFO_COLS = [
    "FACULTAD",
    "ESCUELA",
    "NOMBRE DEL PROGRAMA",
    "MODALIDAD",
    "NIVEL",
    "COMPARTIDAS /INSTITUCIONAL",
    "SEDE",
    "SNIES VIGENTE",
    "PERIODO DE IMPLEMENTACIÓN",
]

FAC_ABREV = {
    "Facultad de Sociedad, Cultura y Creatividad": "FSCC",
    "Facultad de Ingeniería, Diseño e Innovación": "FIDI",
    "Facultad de Negocios, Gestión y Sostenibilidad": "FNGS",
}

from utils.poli_theme import FACULTAD_CLR, color_for_pct

FAC_COLORS = FACULTAD_CLR
FAC_ABREV_INV = {v: k for k, v in FAC_ABREV.items()}

STATUS_LABEL = {
    "done": "Finalizado / Aprobado",
    "inprog": "En proceso",
    "nostart": "Sin iniciar",
    "devuelto": "Devuelto",
    "info": "Informativo",
    "na": "No aplica",
}

# ── Formatos / procesos (reemplazan las 4 etapas canónicas) ────────────────
FORMATOS_ORDEN = [
    "Aseguramiento de la Calidad",
    "Formato Creación de Programas Banner",
    "Proyecciones Académicas",
    "Resultados de Aprendizaje RA",
    "Actas de Homologación",
    "Syllabus",
    "Gerencia de Educación Virtual",
    "Gerencia de Operaciones Academicas",
    "Gerencia de Operaciones Academicas-Banner",
    "Convenios y Homologaciones",
    "Dirección de Mercado",
]

FORMATO_CLR = {
    "Aseguramiento de la Calidad": "#FBAF17",
    "Formato Creación de Programas Banner": "#2980B9",
    "Proyecciones Académicas": "#1FB2DE",
    "Resultados de Aprendizaje RA": "#7c3aed",
    "Actas de Homologación": "#EC0677",
    "Syllabus": "#0891b2",
    "Gerencia de Educación Virtual": "#A6CE38",
    "Gerencia de Operaciones Academicas": "#059669",
    "Gerencia de Operaciones Academicas-Banner": "#0F385A",
    "Convenios y Homologaciones": "#F47B20",
    "Dirección de Mercado": "#6d28d9",
}

FORMATO_PCT_COL = {f: f"pct_fmt_{i}" for i, f in enumerate(FORMATOS_ORDEN)}
FORMATO_SLUG = {f: f"fmt{i}" for i, f in enumerate(FORMATOS_ORDEN)}

# Texto crudo de la fila 9 (normalizado) -> nombre canónico del formato
_FORMATO_CANON: dict[str, str] = {
    "aseguramiento de la calidad": "Aseguramiento de la Calidad",
    "formato creación de programas banner": "Formato Creación de Programas Banner",
    "proyecciones académicas- plan vigente - plan propuesto": "Proyecciones Académicas",
    "resultados de aprendizaje ra": "Resultados de Aprendizaje RA",
    "actas de homologación": "Actas de Homologación",
    "syllabus": "Syllabus",
    "gerencia de educación virtual": "Gerencia de Educación Virtual",
    "gerencia de operaciones academicas": "Gerencia de Operaciones Academicas",
    "gerencia de operaciones academicas-banner": "Gerencia de Operaciones Academicas-Banner",
    "convenios y homologaciones": "Convenios y Homologaciones",
    "dirección de mercado": "Dirección de Mercado",
}

# Campos con % explícito en el Excel: se usan directamente como % del
# formato (no se promedian checkboxes) — Producción de Contenidos y Aulas
# Master ya traen su propio % calculado en el Excel.
_OVERRIDE_PCT_KEYWORDS = (
    "% de avance producción",
    "%avance de aulas master",
    "% avance de aulas master",
)

# Campos informativos (conteos, fechas, texto categórico) que NO se
# consideran "selectores" y por lo tanto se excluyen del promedio de avance
# del formato.
_EXCLUDE_KEYWORDS = (
    "tipo de trámite",
    "total modulos a producir",
    "total módulos a producir",
    "módulos en proceso de producción",
    "modulos en proceso de producción",
    "fecha proyectada de entrega",
    "total módulos recibidos",
    "modulos a crear",
    "total aulas master",
    "estado del aula",
    "fecha de parametrización",
    "total de convenios",
    "nº parametrización",
    "n° parametrización",
    "% de parametrización de convenios",
)

_DONE_VALUES = {"finalizado", "aprobado", "publicado", "si", "sí", "true"}
_DEVUELTO_VALUES = {"devuelto", "devuelta"}
_INPROG_VALUES = {"en proceso", "en ajustes"}
_NOSTART_VALUES = {"sin iniciar", "sin publicar", "pendiente", "no", "false"}
_NA_VALUES = {"no aplica", "", "none", "nan", "—"}


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _cls_formato_option(v) -> float | None:
    """Puntaje 0-100 de un campo 'selector' (checkbox True/False o estado)."""
    s = _norm(v)
    if not s or s in _NA_VALUES:
        return None
    if s in _DONE_VALUES:
        return 100.0
    if s in _DEVUELTO_VALUES:
        return 25.0
    if s in _INPROG_VALUES:
        return 50.0
    if s in _NOSTART_VALUES:
        return 0.0
    return None  # valor no reconocido -> se excluye del promedio


def _score_to_class(score: float | None) -> str:
    if score is None:
        return "na"
    if score >= 100:
        return "done"
    if score >= 50:
        return "inprog"
    if score >= 25:
        return "devuelto"
    return "nostart"


def _cls_pct_value(v) -> float | None:
    """Convierte valor de celda a porcentaje 0-100."""
    s = str(v).strip().lower()
    if not s or s in ("none", "nan", "no aplica", "—", ""):
        return None
    try:
        f = float(s.replace("%", "").replace(",", "."))
        return round(f * 100 if 0 <= f <= 1 else f, 1)
    except Exception:
        return None


def _cls_numeric_value(v) -> float | None:
    """Convierte valor de celda a número (conteos de módulos, aulas, etc.)."""
    s = str(v).strip().lower()
    if not s or s in ("none", "nan", "no aplica", "—", ""):
        return None
    try:
        return float(s.replace(",", "."))
    except Exception:
        return None


def _fmt_excel_date(v) -> str:
    """Convierte número serial Excel o fecha ISO ya parseada a texto dd/mm/aaaa."""
    s = str(v).strip()
    if not s or s.lower() in ("none", "nan", "no aplica", "—", ""):
        return "—"
    try:
        n = float(s)
        if n <= 0:
            return "—"
        from datetime import date, timedelta

        dt = date(1899, 12, 30) + timedelta(days=int(n))
        return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass
    try:
        parsed = pd.to_datetime(s, errors="raise")
        return parsed.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return s


def homologar_nivel(val) -> str:
    v = str(val).strip().lower()
    if any(v.startswith(p) for p in ("profesional", "tecnico", "técnico", "tecnologico", "tecnológico", "tecn")):
        return "Pregrado"
    if any(v.startswith(p) for p in ("maestria", "maestría", "especializacion", "especialización", "maest", "espec")):
        return "Posgrado"
    return ""


def _field_role(formato: str, header_norm: str) -> str:
    """Clasifica un campo dentro de su formato: 'override' | 'exclude' | 'score'."""
    if any(k in header_norm for k in _OVERRIDE_PCT_KEYWORDS):
        return "override"
    if any(k in header_norm for k in _EXCLUDE_KEYWORDS):
        return "exclude"
    return "score"


def _build_group_by_col(raw: pd.DataFrame) -> dict[int, str | None]:
    """Asigna a cada columna su formato canónico (forward-fill de la fila 9),
    ignorando el rótulo genérico que cubre las columnas de información.
    """
    group_row = raw.iloc[GROUP_ROW]
    group_by_col: dict[int, str | None] = {}
    current: str | None = None
    info_end_col = 11  # última columna de info (L = 'Priorización preliminar')
    for j in range(raw.shape[1]):
        if j <= info_end_col:
            group_by_col[j] = None
            continue
        gv = group_row.iloc[j]
        if pd.notna(gv) and str(gv).strip():
            canon = _FORMATO_CANON.get(_norm(gv))
            if canon:
                current = canon
        group_by_col[j] = current
    return group_by_col


def _build_field_map(raw: pd.DataFrame) -> list[dict]:
    """Retorna lista de campos por columna: {col_idx, formato, name, role}."""
    group_by_col = _build_group_by_col(raw)
    header_row = raw.iloc[HEADER_ROW]

    fields: list[dict] = []
    for j in range(raw.shape[1]):
        formato = group_by_col.get(j)
        if not formato:
            continue
        hv = header_row.iloc[j]
        if pd.isna(hv) or not str(hv).strip():
            continue
        hname = str(hv).strip()
        role = _field_role(formato, _norm(hname))
        fields.append({"col_idx": j, "formato": formato, "name": hname, "role": role})
    return fields


def _build_activities_meta_list(fields: list[dict]) -> list[dict]:
    """Metadatos de campos 'selector' (role == score), para tablas de detalle."""
    built: list[dict] = []
    act_idx = 0
    for f in fields:
        if f["role"] != "score":
            continue
        built.append({"idx": act_idx, "phase": f["formato"], "name": f["name"], "responsable": "—"})
        act_idx += 1
    return built


def _build_etapas_df() -> pd.DataFrame:
    global _ACTIVITIES_META

    raw = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME, header=None, dtype=str)
    raw = raw.fillna("")

    fields = _build_field_map(raw)

    info_idx = {
        "FACULTAD": 2,
        "ESCUELA": 3,
        "NOMBRE DEL PROGRAMA": 4,
        "MODALIDAD": 5,
        "NIVEL": 6,
        "COMPARTIDAS /INSTITUCIONAL": 7,
        "SEDE": 8,
        "SNIES VIGENTE": 9,
        "PERIODO DE IMPLEMENTACIÓN": 10,
    }

    data_raw = raw.iloc[DATA_START_ROW:].reset_index(drop=True)
    n_rows = len(data_raw)

    prog_idx = info_idx["NOMBRE DEL PROGRAMA"]
    mask = data_raw.iloc[:, prog_idx].astype(str).str.strip() != ""
    data_raw = data_raw[mask].reset_index(drop=True)
    n_rows = len(data_raw)

    df = pd.DataFrame()
    for col_name, idx in info_idx.items():
        if idx < data_raw.shape[1]:
            df[col_name] = data_raw.iloc[:, idx].astype(str).str.strip().values

    if "PERIODO DE IMPLEMENTACIÓN" in df.columns:
        df["PERIODO DE IMPLEMENTACIÓN"] = (
            df["PERIODO DE IMPLEMENTACIÓN"].astype(str).str.split("\n").str[0].str.strip()
        )

    if "NIVEL" in df.columns:
        df["NIVEL_HOMOLOGADO"] = df["NIVEL"].apply(homologar_nivel)
    else:
        df["NIVEL_HOMOLOGADO"] = ""

    if "FACULTAD" in df.columns:
        df["FACULTAD_ABREV"] = df["FACULTAD"].map(FAC_ABREV).fillna("—")
        df["FACULTAD_COLOR"] = df["FACULTAD_ABREV"].map(FAC_COLORS).fillna("#6e7681")
    else:
        df["FACULTAD_ABREV"] = "—"
        df["FACULTAD_COLOR"] = "#6e7681"

    # ── Campos de Producción de Contenidos y Aulas Master (informativos) ───
    _RAW_FIELDS = {
        "prod_total_modulos": "total modulos a producir",
        "prod_modulos_proceso": "módulos en proceso de producción",
        "prod_pct_avance": "% de avance producción",
        "prod_fecha_entrega": "fecha proyectada de entrega",
        "aulas_modulos_conformidad": "total módulos recibidos a conformidad",
        "aulas_modulos_a_crear": "modulos a crear",
        "aulas_total_creadas": "total aulas master credos",
        "aulas_pct_avance": "%avance de aulas master",
    }
    header_row = raw.iloc[HEADER_ROW]
    for out_col, key in _RAW_FIELDS.items():
        col_idx = None
        for j in range(raw.shape[1]):
            hv = header_row.iloc[j]
            if pd.notna(hv) and key in _norm(hv):
                col_idx = j
                break
        if col_idx is None or col_idx >= data_raw.shape[1]:
            continue
        series = data_raw.iloc[:n_rows, col_idx].astype(str).str.strip()
        series.index = range(n_rows)
        if out_col.endswith("fecha_entrega"):
            df[out_col] = series.apply(_fmt_excel_date).values
        elif out_col.endswith("pct_avance"):
            df[out_col] = series.apply(lambda v: _cls_pct_value(v) if _cls_pct_value(v) is not None else np.nan).values
        else:
            df[out_col] = series.apply(_cls_numeric_value).values

    # ── Campos 'selector' por formato (checkboxes / estados) ───────────────
    act_idx = 0
    formato_scores: dict[str, list[pd.Series]] = {f: [] for f in FORMATOS_ORDEN}
    formato_override: dict[str, pd.Series] = {}

    for field in fields:
        col_idx = field["col_idx"]
        if col_idx >= data_raw.shape[1]:
            continue
        series = data_raw.iloc[:n_rows, col_idx].astype(str).str.strip()
        series.index = range(n_rows)
        formato = field["formato"]

        if field["role"] == "override":
            pct_series = series.apply(lambda v: _cls_pct_value(v) if _cls_pct_value(v) is not None else np.nan)
            formato_override[formato] = pct_series
            continue

        if field["role"] == "exclude":
            continue

        # role == "score": campo selector (checkbox o estado)
        score_series = series.apply(_cls_formato_option)
        formato_scores.setdefault(formato, []).append(score_series)

        cl_col = f"cl_act_{act_idx}"
        val_col = f"val_act_{act_idx}"
        df[cl_col] = score_series.apply(_score_to_class).values
        df[val_col] = series.apply(
            lambda v: ("Sí" if _norm(v) == "true" else "No" if _norm(v) == "false" else (
                str(v).strip() if str(v).strip() not in ("", "None", "nan") else "—"
            ))
        ).values
        df[f"act_phase_{act_idx}"] = formato
        df[f"act_name_{act_idx}"] = field["name"]
        df[f"act_owner_{act_idx}"] = "—"
        act_idx += 1

    df["_n_activities"] = act_idx

    # ── % de avance por formato ──────────────────────────────────────────
    for formato in FORMATOS_ORDEN:
        pct_key = FORMATO_PCT_COL[formato]
        if formato in formato_override:
            df[pct_key] = formato_override[formato].fillna(0).astype(float).values
            continue
        scores = formato_scores.get(formato) or []
        if scores:
            mat = pd.concat(scores, axis=1)
            calc = mat.mean(axis=1, skipna=True).round(1)
            df[pct_key] = calc.fillna(0).values
        else:
            df[pct_key] = 0.0

    # ── Avance general: promedio simple de los % de los 11 formatos ────────
    pct_cols = [FORMATO_PCT_COL[f] for f in FORMATOS_ORDEN]
    df["avance_general_vact"] = df[pct_cols].mean(axis=1).round(1)

    _ACTIVITIES_META = _build_activities_meta_list(fields)
    return df


_ACTIVITIES_META: list[dict] = []


def _ensure_activities_meta(df: pd.DataFrame) -> list[dict]:
    global _ACTIVITIES_META
    if _ACTIVITIES_META:
        return _ACTIVITIES_META
    n = int(df.get("_n_activities", pd.Series([0])).iloc[0]) if len(df) else 0
    meta = []
    for i in range(n):
        if f"act_phase_{i}" not in df.columns:
            break
        meta.append({
            "idx": i,
            "phase": df[f"act_phase_{i}"].iloc[0],
            "name": df[f"act_name_{i}"].iloc[0],
            "responsable": "—",
        })
    _ACTIVITIES_META = meta
    return meta


def load_etapas_data() -> pd.DataFrame:
    """Retorna DataFrame procesado de la hoja 'Etapas (2)' (Excel V2)."""
    global _ACTIVITIES_META
    try:
        import streamlit as st

        cache_key = _data_file_cache_key()

        @st.cache_data
        def _cached(_mtime: int, _size: int, _loader_mtime: int):
            return _build_etapas_df()

        df = _cached(cache_key[0], cache_key[1], cache_key[2])
    except Exception:
        df = _build_etapas_df()
    if not _ACTIVITIES_META:
        _ensure_activities_meta(df)
    return df


def apply_filters_vact(
    df: pd.DataFrame,
    modalidad=None,
    facultad=None,
    periodo=None,
    nivel=None,
) -> pd.DataFrame:
    """Filtros para Fase 2. facultad acepta nombres completos o abreviaturas."""

    def _to_list(v):
        if not v:
            return []
        return list(v) if not isinstance(v, str) else [v]

    mods = _to_list(modalidad)
    facs = _to_list(facultad)
    pers = _to_list(periodo)
    nivs = _to_list(nivel)

    out = df.copy()
    if mods and "MODALIDAD" in out.columns:
        out = out[out["MODALIDAD"].isin(mods)]
    if facs and "FACULTAD" in out.columns:
        full_facs = [FAC_ABREV_INV.get(f, f) for f in facs]
        out = out[out["FACULTAD"].isin(full_facs)]
    if pers and "PERIODO DE IMPLEMENTACIÓN" in out.columns:
        mask = out["PERIODO DE IMPLEMENTACIÓN"].isin(pers)
        if "2027-1" in pers:
            mask = mask | out["PERIODO DE IMPLEMENTACIÓN"].str.contains("2027-1", na=False)
        if any("oferta" in str(p).lower() for p in pers):
            mask = mask | out["PERIODO DE IMPLEMENTACIÓN"].str.contains("oferta", case=False, na=False)
        out = out[mask]
    if nivs and "NIVEL_HOMOLOGADO" in out.columns:
        out = out[out["NIVEL_HOMOLOGADO"].isin(nivs)]
    return out.reset_index(drop=True)


def get_etapas_by_programa(df: pd.DataFrame, nombre_programa: str) -> dict:
    """Actividades (campos selector) + % por formato de un programa."""
    meta = _ensure_activities_meta(df)
    row = df[df["NOMBRE DEL PROGRAMA"].astype(str).str.strip() == str(nombre_programa).strip()]
    if row.empty:
        return {"programa": nombre_programa, "etapas": {}}
    row = row.iloc[0]
    result: dict = {"programa": nombre_programa, "etapas": {}}
    for formato in FORMATOS_ORDEN:
        acts = []
        for m in meta:
            if m["phase"] != formato:
                continue
            i = m["idx"]
            cl = row.get(f"cl_act_{i}", "na")
            val = row.get(f"val_act_{i}", "—")
            score = _cls_formato_option(val) if val not in ("Sí", "No") else (100.0 if val == "Sí" else 0.0)
            acts.append({
                "nombre": m["name"],
                "estado": STATUS_LABEL.get(cl, cl),
                "estado_key": cl,
                "valor": val,
                "pct": score if score is not None else "—",
                "responsable": "—",
            })
        pct_col = FORMATO_PCT_COL[formato]
        result["etapas"][formato] = {
            "pct": float(row.get(pct_col, 0) or 0),
            "actividades": acts,
        }
    result["avance_general"] = float(row.get("avance_general_vact", 0) or 0)
    result["info"] = {
        "FACULTAD": row.get("FACULTAD", "—"),
        "FACULTAD_ABREV": row.get("FACULTAD_ABREV", "—"),
        "FACULTAD_COLOR": row.get("FACULTAD_COLOR", "#6e7681"),
        "ESCUELA": row.get("ESCUELA", "—"),
        "MODALIDAD": row.get("MODALIDAD", "—"),
        "NIVEL": row.get("NIVEL", "—"),
        "NIVEL_HOMOLOGADO": row.get("NIVEL_HOMOLOGADO", "—"),
        "SEDE": row.get("SEDE", "—"),
        "SNIES VIGENTE": row.get("SNIES VIGENTE", "—"),
        "COMPARTIDAS /INSTITUCIONAL": row.get("COMPARTIDAS /INSTITUCIONAL", "—"),
        "PERIODO DE IMPLEMENTACIÓN": row.get("PERIODO DE IMPLEMENTACIÓN", "—"),
    }
    result["produccion"] = {
        "total_modulos": row.get("prod_total_modulos", None),
        "modulos_proceso": row.get("prod_modulos_proceso", None),
        "pct_avance": row.get("prod_pct_avance", None),
        "fecha_entrega": row.get("prod_fecha_entrega", "—"),
    }
    result["aulas_master"] = {
        "modulos_conformidad": row.get("aulas_modulos_conformidad", None),
        "modulos_a_crear": row.get("aulas_modulos_a_crear", None),
        "total_creadas": row.get("aulas_total_creadas", None),
        "pct_avance": row.get("aulas_pct_avance", None),
    }
    return result


def get_estadisticas_etapa(df: pd.DataFrame, formato_name: str) -> dict:
    """Estadísticas agregadas de un formato sobre el df filtrado."""
    pct_col = FORMATO_PCT_COL.get(formato_name)
    if not pct_col or pct_col not in df.columns or len(df) == 0:
        return {
            "pct_promedio": 0, "done": 0, "inprog": 0, "devuelto": 0,
            "nostart": 0, "info": 0, "na": 0, "total_act": 0, "n_programas": 0,
        }

    meta = _ensure_activities_meta(df)
    acts_meta = [m for m in meta if m["phase"] == formato_name]
    done = inprog = devuelto = nostart = info = na = 0
    for m in acts_meta:
        col = f"cl_act_{m['idx']}"
        if col not in df.columns:
            continue
        for cl in df[col]:
            if cl == "done":
                done += 1
            elif cl == "inprog":
                inprog += 1
            elif cl == "devuelto":
                devuelto += 1
            elif cl == "nostart":
                nostart += 1
            elif cl == "info":
                info += 1
            else:
                na += 1

    return {
        "pct_promedio": round(float(df[pct_col].mean()), 1),
        "done": done, "inprog": inprog, "devuelto": devuelto,
        "nostart": nostart, "info": info, "na": na,
        "total_act": len(acts_meta) * len(df) if acts_meta else 0,
        "n_programas": len(df),
    }


def get_detalle_etapa(df: pd.DataFrame, formato_name: str) -> dict:
    """Desglose ampliado de un formato: estados, % y lista de campos selector."""
    stats = get_estadisticas_etapa(df, formato_name)
    meta = _ensure_activities_meta(df)
    acts_meta = [m for m in meta if m["phase"] == formato_name]
    n_prog = len(df)
    actividades = []
    for m in acts_meta:
        col = f"cl_act_{m['idx']}"
        if col not in df.columns:
            continue
        done = int((df[col] == "done").sum())
        inprog = int((df[col] == "inprog").sum())
        devuelto = int((df[col] == "devuelto").sum())
        nostart = int((df[col] == "nostart").sum())
        info = int((df[col] == "info").sum())
        na = int((df[col] == "na").sum())
        scores = df[col].map({"done": 100.0, "inprog": 50.0, "devuelto": 25.0, "nostart": 0.0}).astype(float)
        pct_avance = round(float(scores.dropna().mean()), 1) if scores.notna().any() else 0.0
        actividades.append({
            "nombre": m["name"], "done": done, "inprog": inprog, "devuelto": devuelto,
            "nostart": nostart, "info": info, "na": na,
            "pct_done": round(done / n_prog * 100, 1) if n_prog else 0,
            "pct_avance": pct_avance,
        })
    actividades.sort(key=lambda a: (-a["pct_done"], a["nombre"]))
    total_cells = stats.get("total_act") or 0
    pct_por_estado = {}
    for k in ("done", "inprog", "devuelto", "nostart", "info", "na"):
        pct_por_estado[k] = round(stats[k] / total_cells * 100, 1) if total_cells else 0
    return {
        **stats, "actividades": actividades, "pct_por_estado": pct_por_estado,
        "n_actividades": len(acts_meta),
    }


def get_estadisticas_produccion(df: pd.DataFrame) -> dict:
    """Estadísticas agregadas de Producción de Contenidos (Gerencia Educación Virtual)."""
    if len(df) == 0:
        return {"total_modulos": 0, "modulos_proceso": 0, "pct_avance_promedio": 0, "n_programas_con_dato": 0}
    total_modulos = float(df.get("prod_total_modulos", pd.Series(dtype=float)).sum())
    modulos_proceso = float(df.get("prod_modulos_proceso", pd.Series(dtype=float)).sum())
    pct_col = df.get("prod_pct_avance", pd.Series(dtype=float))
    n_con_dato = int((pct_col.fillna(0) > 0).sum()) if len(pct_col) else 0
    pct_avance = round(float(pct_col.mean()), 1) if len(pct_col) and pct_col.notna().any() else 0.0
    return {
        "total_modulos": total_modulos, "modulos_proceso": modulos_proceso,
        "pct_avance_promedio": pct_avance, "n_programas_con_dato": n_con_dato,
    }


def get_estadisticas_aulas_master(df: pd.DataFrame) -> dict:
    """Estadísticas agregadas de Aulas Master (Gerencia de Operaciones Academicas)."""
    if len(df) == 0:
        return {"total_creadas": 0, "modulos_a_crear": 0, "pct_avance_promedio": 0, "n_programas_con_dato": 0}
    total_creadas = float(df.get("aulas_total_creadas", pd.Series(dtype=float)).sum())
    modulos_a_crear = float(df.get("aulas_modulos_a_crear", pd.Series(dtype=float)).sum())
    pct_col = df.get("aulas_pct_avance", pd.Series(dtype=float))
    n_con_dato = int((pct_col.fillna(0) > 0).sum()) if len(pct_col) else 0
    pct_avance = round(float(pct_col.mean()), 1) if len(pct_col) and pct_col.notna().any() else 0.0
    return {
        "total_creadas": total_creadas, "modulos_a_crear": modulos_a_crear,
        "pct_avance_promedio": pct_avance, "n_programas_con_dato": n_con_dato,
    }
