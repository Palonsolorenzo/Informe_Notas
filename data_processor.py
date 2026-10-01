import pandas as pd
import numpy as np

# Nombres estándar de columnas requeridos
COL_ASIGNATURA = "Nombre de la Asignatura"
COL_MATRICULADOS = "Nº alumnos matriculados"
COL_CO_EV_CONT = "Nº Alumnos superan Evaluación continua en CO"
COL_CO_EV_FIN = "Nº Alumnos superan Evaluación final en CO"
COL_CO_TOTAL = "Nº alumnos que superan asignatura en CO"
COL_CE_EV_CONT = "Nº Alumnos superan Evaluación continua en CE"
COL_CE_EV_FIN = "Nº Alumnos superan Evaluación final en CE"
COL_CE_TOTAL = "Nº alumnos que superan asignatura en CE"

# Nuevas columnas calculadas
COL_PCT_CO_EV_CONT = "% Ev. Continua CO"
COL_PCT_CO_EV_FIN = "% Ev. Final CO"
COL_PCT_CO_TOTAL = "% Aprobados CO"
COL_PCT_CE_EV_CONT = "% Ev. Continua CE"
COL_PCT_CE_EV_FIN = "% Ev. Final CE"
COL_PCT_CE_TOTAL = "% Aprobados CE"
COL_TOTAL_ACUMULADO = "Total Aprobados (CO + CE)"
COL_PCT_TOTAL_ACUMULADO = "% Aprobados Total"

NUMERIC_COLS = [
    COL_MATRICULADOS,
    COL_CO_EV_CONT,
    COL_CO_EV_FIN,
    COL_CO_TOTAL,
    COL_CE_EV_CONT,
    COL_CE_EV_FIN,
    COL_CE_TOTAL,
]

def get_excel_sheets(file_source):
    """Obtiene la lista de nombres de pestañas/hojas de un archivo Excel."""
    try:
        xl = pd.ExcelFile(file_source)
        return xl.sheet_names
    except Exception:
        return ["Titulación"]

def find_matching_column(df_columns, target_name):
    """Busca una columna en el DataFrame ignorando mayúsculas, espacios y caracteres especiales."""
    import re
    norm_target = re.sub(r'[^a-zA-Z0-9]', '', target_name.lower())
    for col in df_columns:
        norm_col = re.sub(r'[^a-zA-Z0-9]', '', str(col).lower())
        if norm_target in norm_col or norm_col in norm_target:
            return col
    return None

def preprocess_dataframe(df):
    """Limpia y normaliza el DataFrame importado de Excel."""
    df_clean = df.copy()
    
    col_mapping = {}
    expected_cols = [
        (COL_ASIGNATURA, "asignatura"),
        (COL_MATRICULADOS, "matriculados"),
        (COL_CO_EV_CONT, "evaluacion continua en co"),
        (COL_CO_EV_FIN, "evaluacion final en co"),
        (COL_CO_TOTAL, "superan asignatura en co"),
        (COL_CE_EV_CONT, "evaluacion continua en ce"),
        (COL_CE_EV_FIN, "evaluacion final en ce"),
        (COL_CE_TOTAL, "superan asignatura en ce"),
    ]
    
    for std_name, _ in expected_cols:
        matched = find_matching_column(df_clean.columns, std_name)
        if matched:
            col_mapping[matched] = std_name
            
    df_clean = df_clean.rename(columns=col_mapping)
    
    for std_name, _ in expected_cols:
        if std_name not in df_clean.columns:
            if std_name == COL_ASIGNATURA:
                df_clean[std_name] = "Sin nombre"
            else:
                df_clean[std_name] = 0

    df_clean[COL_ASIGNATURA] = df_clean[COL_ASIGNATURA].astype(str).str.strip()
    df_clean = df_clean[df_clean[COL_ASIGNATURA].str.len() > 0]
    
    for col in NUMERIC_COLS:
        df_clean[col] = pd.to_numeric(
            df_clean[col].astype(str).str.replace('-', '0').str.replace(' ', ''),
            errors='coerce'
        ).fillna(0).astype(int)
        
    return df_clean

def enrich_data(df_clean):
    """Añade columnas de porcentaje y totales acumulados a cada asignatura."""
    df_enriched = df_clean.copy()
    
    matr = df_enriched[COL_MATRICULADOS].replace(0, np.nan)
    
    df_enriched[COL_PCT_CO_EV_CONT] = (df_enriched[COL_CO_EV_CONT] / matr).fillna(0)
    df_enriched[COL_PCT_CO_EV_FIN] = (df_enriched[COL_CO_EV_FIN] / matr).fillna(0)
    df_enriched[COL_PCT_CO_TOTAL] = (df_enriched[COL_CO_TOTAL] / matr).fillna(0)
    
    df_enriched[COL_PCT_CE_EV_CONT] = (df_enriched[COL_CE_EV_CONT] / matr).fillna(0)
    df_enriched[COL_PCT_CE_EV_FIN] = (df_enriched[COL_CE_EV_FIN] / matr).fillna(0)
    df_enriched[COL_PCT_CE_TOTAL] = (df_enriched[COL_CE_TOTAL] / matr).fillna(0)
    
    df_enriched[COL_TOTAL_ACUMULADO] = df_enriched[COL_CO_TOTAL] + df_enriched[COL_CE_TOTAL]
    df_enriched[COL_PCT_TOTAL_ACUMULADO] = (df_enriched[COL_TOTAL_ACUMULADO] / matr).fillna(0)
    
    return df_enriched

def get_global_metrics(df_enriched):
    """Calcula el resumen global de toda la titulación."""
    tot_matriculados = int(df_enriched[COL_MATRICULADOS].sum())
    tot_co_ev_cont = int(df_enriched[COL_CO_EV_CONT].sum())
    tot_co_ev_fin = int(df_enriched[COL_CO_EV_FIN].sum())
    tot_co_total = int(df_enriched[COL_CO_TOTAL].sum())
    
    tot_ce_ev_cont = int(df_enriched[COL_CE_EV_CONT].sum())
    tot_ce_ev_fin = int(df_enriched[COL_CE_EV_FIN].sum())
    tot_ce_total = int(df_enriched[COL_CE_TOTAL].sum())
    
    tot_acumulado = int(df_enriched[COL_TOTAL_ACUMULADO].sum())
    
    pct_co_ev_cont = tot_co_ev_cont / tot_matriculados if tot_matriculados > 0 else 0
    pct_co_ev_fin = tot_co_ev_fin / tot_matriculados if tot_matriculados > 0 else 0
    pct_co_total = tot_co_total / tot_matriculados if tot_matriculados > 0 else 0
    
    pct_ce_ev_cont = tot_ce_ev_cont / tot_matriculados if tot_matriculados > 0 else 0
    pct_ce_ev_fin = tot_ce_ev_fin / tot_matriculados if tot_matriculados > 0 else 0
    pct_ce_total = tot_ce_total / tot_matriculados if tot_matriculados > 0 else 0
    
    pct_acumulado = tot_acumulado / tot_matriculados if tot_matriculados > 0 else 0
    
    return {
        "num_asignaturas": len(df_enriched),
        "tot_matriculados": tot_matriculados,
        "tot_co_ev_cont": tot_co_ev_cont,
        "pct_co_ev_cont": pct_co_ev_cont,
        "tot_co_ev_fin": tot_co_ev_fin,
        "pct_co_ev_fin": pct_co_ev_fin,
        "tot_co_total": tot_co_total,
        "pct_co_total": pct_co_total,
        "tot_ce_ev_cont": tot_ce_ev_cont,
        "pct_ce_ev_cont": pct_ce_ev_cont,
        "tot_ce_ev_fin": tot_ce_ev_fin,
        "pct_ce_ev_fin": pct_ce_ev_fin,
        "tot_ce_total": tot_ce_total,
        "pct_ce_total": pct_ce_total,
        "tot_acumulado": tot_acumulado,
        "pct_acumulado": pct_acumulado,
    }
