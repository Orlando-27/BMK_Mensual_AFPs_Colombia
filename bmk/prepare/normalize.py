"""
Normalizacion de insumos (paso 2 del manual, y preparacion para clasificar).

- Extrae de Formato_351 la tabla de activos con las 15 columnas que el proceso
  usa (mapeo exacto tomado de BMO/Modulo2.Importar_activos_industria), aplicando
  el filtro Nombre Patrimonio <> "FONDO ALTERNATIVO".
- Normaliza nombres de AFP y de portafolio a los codigos internos.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# Mapeo Formato_351 (indice 0-based) -> nombre logico. Fuente: BMO Modulo2.
COL_351 = {
    "portafolio": 6,        # G  Nombre Patrimonio
    "afp": 2,               # C  Nombre de Entidad
    "nemo": 27,             # AB Nemotecnico
    "isin": 81,             # CD No. Iden Asignado Custodio
    "clas_sfc": 25,         # Z  Clase de Inversion (codigo SFC)
    "emisor": 17,           # R  Razon Social Emisor
    "f_compra": 33,         # AH Fecha Compra
    "f_vcto": 31,           # AF Fecha Vencim Titulo
    "moneda": 34,           # AI Codigo Moneda
    "valor_nominal": 36,    # AK Valor Nominal
    "valor_nominal_residual": 38,  # AM Vlor_Nominal_Residual_Capit
    "nro_acciones": 39,     # AN No. Acciones
    "vr_moneda_ext": 55,    # Vr_mrcd_pres_mnda_dif_peso (valor en moneda del activo)
    "vr_mercado": 53,       # BB Vr. mercado o Vr presente en $
    "tasa_facial_valor": 56,  # BE Tasa de negoc
    "tasa_facial_ind": 43,  # AR Tasa Facial Titulo
}

FONDO_ALTERNATIVO = "FONDO ALTERNATIVO"

# Portafolio SFC -> codigo interno (Diccionario SFC, bloque F:J).
PORTAFOLIO_COD = {
    "FONDO DE PENSIONES OBLIGATORIAS MODERADO": "PO",
    "FONDO DE PENSIONES OBLIGATORIAS CONSERVADOR": "PC",
    "FONDO DE PENSIONES OBLIGATORIAS MAYOR RIESGO": "PM",
    "FONDO DE PENSIONES OBLIGATORIAS PORVENIR RETIRO PROGRAMADO": "PR",
    "FONDO DE PENSIONES OBLIGATORIAS PROTECCION PORVENIR RETIRO PROGRAMADO": "PR",
    "FONDO DE CESANTIAS": "CS",
    "FONDO DE CESANTIAS LEY 50": "CS",
    "FONDO DE CESANTIAS SKANDIA": "CS",
    "FONDO ESPECIAL DE RETIRO PROGRAMADO": "PR",
}


# Normalizacion de AFP por palabra clave (robusta a comillas y sufijos del SFC).
# Nombres crudos SFC: '"PORVENIR"', '"PROTECCION"', 'SKANDIA AFP - ACCAI S.A.',
# '"COLFONDOS S.A." Y "COLFONDOS"'. Se llevan al nombre canonico interno.
AFP_CANON = [
    ("COLFONDOS", "COLFONDOS"),
    ("PORVENIR", "PORVENIR"),
    ("PROTECCION", "PROTECCION"),
    ("PROTECCIÓN", "PROTECCION"),
    ("SKANDIA", "SKANDIA"),
    ("OLD MUTUAL", "OLD MUTUAL"),
    ("ACCAI", "SKANDIA"),
]


def normalizar_afp(nombre: str) -> str:
    n = str(nombre).upper()
    for clave, canon in AFP_CANON:
        if clave in n:
            return canon
    return str(nombre).strip().strip('"')


def _cod_portafolio(nombre: str) -> str:
    n = str(nombre).strip().upper()
    if n in PORTAFOLIO_COD:
        return PORTAFOLIO_COD[n]
    # Heuristica por palabras clave (robusta a variantes de nombre).
    if "CESANTIAS" in n:
        return "CS"
    if "RETIRO PROGRAMADO" in n:
        return "PR"
    if "MAYOR RIESGO" in n:
        return "PM"
    if "CONSERVADOR" in n:
        return "PC"
    if "MODERADO" in n or "OBLIGATORIAS" in n:
        return "PO"
    return "ND"


def extraer_activos(formato_351: pd.DataFrame, ref_dir: str | Path = "bmk/config/reference") -> pd.DataFrame:
    """Construye la tabla de activos a clasificar a partir de Formato_351."""
    ref_dir = Path(ref_dir)
    cols = formato_351.columns
    out = pd.DataFrame({
        name: formato_351.iloc[:, idx] for name, idx in COL_351.items()
    })
    # Filtro: excluir FONDO ALTERNATIVO (se maneja aparte).
    out = out[out["portafolio"].astype(str).str.strip().str.upper() != FONDO_ALTERNATIVO]
    # Descartar filas totalmente vacias.
    out = out.dropna(how="all")
    # Limpieza de texto.
    for c in ("nemo", "isin", "clas_sfc", "emisor", "afp", "moneda", "portafolio"):
        out[c] = out[c].astype("string").fillna("").str.strip()
    out = out[out["clas_sfc"] != ""]  # sin clas_sfc no es un activo valido
    # Normalizacion AFP y codigo de portafolio.
    out["afp"] = out["afp"].map(normalizar_afp)
    out["cod_portafolio"] = out["portafolio"].map(_cod_portafolio)
    return out.reset_index(drop=True)
