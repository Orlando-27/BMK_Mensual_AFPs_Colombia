"""
Clasificacion de swaps de la industria (Formato_415, J=16/17).

Reimplementa (de forma simplificada) el clasificador de tipo de swap de la
formula BMO Swaps!BX2, a partir de la moneda de cada pata. El resultado
(tipo_swap) es el mismo vocabulario que consume el motor de valoracion v6:
  IRS IBRCOP, IRS SOFUSD, CCS COPUSD, CCS COPUVR, CCS COPEUR, ...

La valoracion VPN/DM se hace con el motor v6 (bmk/pricing), que requiere la
Calculadora Swap del corte (SWAPIND + curvas). Aqui solo se etiqueta y agrega
exposicion nocional por tipo, para reporte y controles.
"""
from __future__ import annotations

import pandas as pd


def _m(v) -> str:
    s = "" if v is None else str(v).strip().upper()
    return "UVR" if s in ("COU", "UVR") else s


def clasificar_tipo(mon_der: str, mon_obl: str) -> str:
    d, o = _m(mon_der), _m(mon_obl)
    par = {d, o}
    if d == o == "COP":
        return "IRS IBRCOP"
    if par == {"COP", "USD"}:
        return "CCS COPUSD"
    if par == {"COP", "UVR"}:
        return "CCS COPUVR"
    if par == {"COP", "EUR"}:
        return "CCS COPEUR"
    if par == {"UVR", "USD"}:
        return "CCS UVRUSD"
    if par == {"UVR", "EUR"}:
        return "CCS UVREUR"
    if d == o == "USD":
        return "IRS SOFUSD"
    return f"OTRO {d}/{o}"


def procesar(swaps: pd.DataFrame) -> pd.DataFrame:
    if swaps.empty:
        return swaps.assign(tipo_swap=[])
    out = swaps.copy()
    out["tipo_swap"] = [
        clasificar_tipo(d, o)
        for d, o in zip(out["moneda_derecho"], out["moneda_obligacion"])
    ]
    return out


def resumen_por_tipo(swaps_clasificados: pd.DataFrame) -> pd.DataFrame:
    if swaps_clasificados.empty:
        return pd.DataFrame()
    g = swaps_clasificados.groupby("tipo_swap").agg(
        n=("tipo_swap", "count"),
        nominal_derecho=("nominal_derecho", "sum"),
        nominal_obligacion=("nominal_obligacion", "sum"),
    ).reset_index()
    return g
