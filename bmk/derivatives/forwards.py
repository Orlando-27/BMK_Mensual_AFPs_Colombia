"""
Procesamiento de forwards (y futuros TRM) de la industria.

Reimplementa las formulas de Fwd Industria (BMO, fila 23):
  - compra/venta segun que pata es COP
  - construccion del par de divisas
  - inversion de paridad (1/strike) para pares invertidos (USDJPY, USDMXN, ...)
  - exposicion (compra positiva / venta negativa)
"""
from __future__ import annotations

import pandas as pd

# Pares que en el mercado se cotizan invertidos respecto a como vienen (1/strike).
PARES_INVERTIDOS = {"USDJPY", "USDMXN", "USDBRL", "USDCAD"}
# Pares que se dejan en orden directo; el resto se invierte al construir el par.
PARES_DIRECTOS = {"USDCOP", "EURUSD", "GBPUSD", "AUDUSD"}


def _s(v) -> str:
    return "" if v is None else str(v).strip().upper()


def _compra_venta(mon_der: str, mon_obl: str) -> str:
    d, o = _s(mon_der), _s(mon_obl)
    if o == "COP":
        return "VENTA"
    if d == "COP":
        return "COMPRA"
    if d != "COP" and o == "USD":
        return "COMPRA"
    if o != "COP" and d == "USD":
        return "VENTA"
    return "REVISAR"


def _par(mon_der: str, mon_obl: str, lado: str) -> str:
    d, o = _s(mon_der), _s(mon_obl)
    directo = (o + d) if lado == "VENTA" else (d + o)
    if directo in PARES_DIRECTOS or directo == "USDCOP" or directo == "EURUSD":
        return directo
    # invertir
    return (d + o) if lado == "VENTA" else (o + d)


def procesar(forwards: pd.DataFrame) -> pd.DataFrame:
    """Agrega compra/venta, par, strike ajustado y exposicion a cada forward."""
    if forwards.empty:
        return forwards.assign(compra_venta=[], par=[], strike_aj=[], exposicion=[])
    out = forwards.copy()
    out["compra_venta"] = [
        _compra_venta(d, o) for d, o in zip(out["moneda_derecho"], out["moneda_obligacion"])
    ]
    out["par"] = [
        _par(d, o, cv) for d, o, cv in
        zip(out["moneda_derecho"], out["moneda_obligacion"], out["compra_venta"])
    ]
    # strike ajustado: 1/strike si el par esta invertido
    out["strike_aj"] = [
        (1.0 / s if (par in PARES_INVERTIDOS and s not in (None, 0) and pd.notna(s)) else s)
        for par, s in zip(out["par"], out["strike"])
    ]
    # exposicion: nominal (derecho si compra, obligacion si venta), con signo
    nominal = [
        (nd if cv == "COMPRA" else no)
        for cv, nd, no in zip(out["compra_venta"], out["nominal_derecho"], out["nominal_obligacion"])
    ]
    signo = out["compra_venta"].map({"COMPRA": 1, "VENTA": -1}).fillna(0)
    out["exposicion"] = pd.Series(nominal, index=out.index).fillna(0) * signo
    return out
