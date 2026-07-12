"""
Output "Valoracion Futuros Industria": replica las hojas `FutLoc Industria`
(futuros locales de TES) y `FutInt Industria` (futuros internacionales) del
Benchmark (Detallado).

Fuente de futuros: Formato_415 del archivo SFC.
  - tipo_derivado = 5 -> futuros locales (FutLoc)
  - tipo_derivado = 6 -> futuros internacionales (FutInt)
Valoracion con el VECTOR DE PRECIOS del corte (ver bmk.pricing.extraer_vector_precios).

Escribe un Excel con hojas "FutLoc Industria" y "FutInt Industria".
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp, _cod_portafolio
from bmk.pricing.fut_valuator import Vector, valorar_local, valorar_int


def _norm_415(formato_415: pd.DataFrame, tipo: int, solo_industria: bool) -> pd.DataFrame:
    def c(k):
        return formato_415.iloc[:, COL_415[k]]
    td = pd.to_numeric(c("tipo_derivado"), errors="coerce")
    df = pd.DataFrame({
        "afp": c("entidad").map(normalizar_afp),
        "portafolio": c("nombre").map(_cod_portafolio),
        "posicion": pd.to_numeric(c("posicion_compra_venta"), errors="coerce"),
        "moneda": c("moneda_derecho"),
        "nominal": pd.to_numeric(c("nominal_derecho"), errors="coerce"),
        "strike": pd.to_numeric(c("strike"), errors="coerce"),
        "tipo": c("tipo_subyacente"),
        "subyacente": c("id_subyacente"),
        "fecha_vencimiento": c("fecha_vencimiento"),
    })
    df = df[td == tipo].reset_index(drop=True)
    if solo_industria:
        df = df[df["afp"] != "COLFONDOS"].reset_index(drop=True)
    return df


def valorar(formato_415: pd.DataFrame, vector_path: str | Path = "insumos/vector_precios.csv",
            fx: dict | None = None, solo_industria: bool = True):
    """Devuelve (locales, internacionales) valorados."""
    vector = Vector(vector_path)
    local = valorar_local(_norm_415(formato_415, 5, solo_industria), vector)
    intl = valorar_int(_norm_415(formato_415, 6, solo_industria), vector, fx)
    return local, intl


def escribir(formato_415: pd.DataFrame, out_dir: str | Path, corte: str,
             vector_path: str | Path = "insumos/vector_precios.csv",
             fx: dict | None = None, solo_industria: bool = True) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    local, intl = valorar(formato_415, vector_path, fx, solo_industria)
    dest = out_dir / f"valoracion_futuros_industria_{corte}.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        local.to_excel(xw, sheet_name="FutLoc Industria", index=False)
        intl.to_excel(xw, sheet_name="FutInt Industria", index=False)
    return str(dest)
