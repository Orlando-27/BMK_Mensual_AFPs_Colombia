"""
Derivacion de tasas de cambio (FX) desde el propio archivo SFC.

El Formato_351 reporta, para activos en moneda extranjera:
  - col 53 (BB): Vr. mercado o Vr presente en $  (COP)
  - col 55     : Vr_mrcd_pres_mnda_dif_peso        (valor en la moneda del activo)
Por lo tanto FX(moneda) = mediana( COP / moneda ) sobre los activos de esa moneda.

Esto evita depender de un vector externo para la TRM (paso "toma del archivo").
"""
from __future__ import annotations

import numpy as np
import pandas as pd

COL_VR_COP = 53
COL_VR_MONEDA = 55
COL_MONEDA = 34


def fx_por_moneda(formato_351: pd.DataFrame) -> dict[str, float]:
    """{moneda: tasa a COP} derivada del archivo SFC (mediana COP/moneda)."""
    cop = pd.to_numeric(formato_351.iloc[:, COL_VR_COP], errors="coerce")
    ext = pd.to_numeric(formato_351.iloc[:, COL_VR_MONEDA], errors="coerce")
    mon = formato_351.iloc[:, COL_MONEDA].astype("string").str.strip().str.upper()
    df = pd.DataFrame({"mon": mon, "cop": cop, "ext": ext})
    df = df[(df["ext"].abs() > 0) & (df["cop"].abs() > 0) & df["mon"].notna()]
    df["r"] = df["cop"] / df["ext"]
    out = {"COP": 1.0}
    for m, g in df.groupby("mon"):
        out[m] = float(g["r"].median())
    return out


def trm(formato_351: pd.DataFrame) -> float:
    """TRM (USD->COP) derivada del archivo SFC."""
    fx = fx_por_moneda(formato_351)
    return fx.get("USD", np.nan)
