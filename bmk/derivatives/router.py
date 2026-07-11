"""
Enrutamiento de derivados de Formato_415 por "Tipo de derivado" (col J).

Reimplementa BMO/Modulo1.Importar_derivados_industria e Importar_Swaps:
  J=1  -> forwards           J=4  -> futuros TRM (se agrupan con forwards)
  J=5  -> futuros locales    J=6  -> futuros internacionales
  J=16 -> swaps (IRS)        J=17 -> swaps (CCS)

Devuelve DataFrames con columnas economicas nombradas (no el layout interno del
BMO), listas para las etapas de exposicion y valoracion.
"""
from __future__ import annotations

import pandas as pd

from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp


def _col(df: pd.DataFrame, key: str) -> pd.Series:
    return df.iloc[:, COL_415[key]]


def _base(df415: pd.DataFrame) -> pd.DataFrame:
    """Proyeccion de las columnas economicas comunes de 415, filtrando vacios."""
    tipo = pd.to_numeric(_col(df415, "tipo_derivado"), errors="coerce")
    out = pd.DataFrame({
        "tipo_derivado": tipo,
        "entidad": _col(df415, "entidad"),
        "afp": _col(df415, "entidad").map(normalizar_afp),
        "portafolio": _col(df415, "codigo_patrimonio"),
        "nombre_patrimonio": _col(df415, "nombre"),
        "numero_contrato": _col(df415, "numero_contrato"),
        "moneda_liquidacion": _col(df415, "moneda_liquidacion"),
        "fecha_celebracion": _col(df415, "fecha_celebracion"),
        "fecha_vencimiento": _col(df415, "fecha_vencimiento"),
        "fecha_liquidacion": _col(df415, "fecha_liquidacion"),
        "posicion_compra_venta": _col(df415, "posicion_compra_venta"),
        "moneda_derecho": _col(df415, "moneda_derecho"),
        "nominal_derecho": pd.to_numeric(_col(df415, "nominal_derecho"), errors="coerce"),
        "moneda_obligacion": _col(df415, "moneda_obligacion"),
        "nominal_obligacion": pd.to_numeric(_col(df415, "nominal_obligacion"), errors="coerce"),
        "strike": pd.to_numeric(_col(df415, "strike"), errors="coerce"),
        "tipo_subyacente": _col(df415, "tipo_subyacente"),
        "id_subyacente": _col(df415, "id_subyacente"),
    })
    return out[out["tipo_derivado"].notna()].reset_index(drop=True)


def enrutar(formato_415: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Devuelve {forwards, futuros_trm, futuros_local, futuros_int, swaps}."""
    base = _base(formato_415)
    return {
        "forwards": base[base["tipo_derivado"] == 1].reset_index(drop=True),
        "futuros_trm": base[base["tipo_derivado"] == 4].reset_index(drop=True),
        "futuros_local": base[base["tipo_derivado"] == 5].reset_index(drop=True),
        "futuros_int": base[base["tipo_derivado"] == 6].reset_index(drop=True),
        "swaps": base[base["tipo_derivado"].isin([16, 17])].reset_index(drop=True),
    }
