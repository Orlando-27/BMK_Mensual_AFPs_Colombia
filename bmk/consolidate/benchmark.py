"""
Consolidacion del benchmark (posiciones de la industria por clase de inversion).

Toma los activos clasificados y produce:
  - valor de fondo por (AFP, portafolio)
  - peso (participacion) de cada activo en su fondo
  - agregacion por clase de inversion / clasificacion x AFP
  - vista industria y vista Colfondos

Reimplementa la esencia de la hoja Benchmark del Detallado (SUMIFS + peso Z),
sin las funciones de Excel; la exposicion detallada por region/moneda (CPRVI) y
la valoracion de derivados se integran en etapas dedicadas.
"""
from __future__ import annotations

import pandas as pd

COLFONDOS = "COLFONDOS"


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def con_pesos(clasificados: pd.DataFrame) -> pd.DataFrame:
    """Agrega valor de fondo por (AFP, portafolio) y el peso de cada activo."""
    out = clasificados.copy()
    out["vr_mercado_num"] = _num(out["vr_mercado"])
    fondo = out.groupby(["afp", "cod_portafolio"])["vr_mercado_num"].transform("sum")
    out["valor_fondo"] = fondo
    out["peso"] = out["vr_mercado_num"] / fondo.replace(0, pd.NA)
    return out


def valor_por_fondo(clasificados: pd.DataFrame) -> pd.DataFrame:
    out = clasificados.copy()
    out["vr_mercado_num"] = _num(out["vr_mercado"])
    g = out.groupby(["afp", "cod_portafolio"]).agg(
        valor_fondo=("vr_mercado_num", "sum"),
        n_titulos=("vr_mercado_num", "count"),
    ).reset_index()
    return g


def por_clase(clasificados: pd.DataFrame, dim: str = "clase_inversion") -> pd.DataFrame:
    """Pivot valor de mercado por dimension x AFP (vista industria)."""
    out = clasificados.copy()
    out["vr_mercado_num"] = _num(out["vr_mercado"])
    piv = pd.pivot_table(out, values="vr_mercado_num", index=dim, columns="afp",
                         aggfunc="sum", fill_value=0.0)
    piv["TOTAL_INDUSTRIA"] = piv.sum(axis=1)
    return piv.reset_index()


def separar_colfondos(clasificados: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Devuelve {industria, colfondos} (Colfondos = solo nosotros)."""
    ind = clasificados
    colf = clasificados[clasificados["afp"] == COLFONDOS].reset_index(drop=True)
    return {"industria": ind, "colfondos": colf}
