"""
Valorador de futuros de la industria — replica las hojas `FutLoc Industria`
(futuros locales, TES) y `FutInt Industria` (futuros internacionales) del
Benchmark (Detallado).

FUTUROS LOCALES (FutLoc Industria) — validado al peso contra la hoja:
  Cada futuro de TES tiene un subyacente (codigo de contrato, p. ej. COL17CT03342)
  con precio en el VECTOR DE PRECIOS (decimal, 0.92411 = 92,411%).
    signo = +1 compra (posicion 1), -1 venta (posicion 2)
    Valor de Mercado = NOMINAL * precio_vector * signo

FUTUROS INTERNACIONALES (FutInt Industria) — segun las formulas de la hoja
(AD/AF): indices y treasuries.
    signo = +1 compra, -1 venta
    precio = nivel del indice o precio del treasury (VECTOR DE PRECIOS); treasury /100
    Valor de Mercado (moneda) = NOMINAL * signo * precio
    Valor de Mercado (COP)    = valor_moneda * FX(moneda del contrato)
  (En el corte del archivo no hay futuros internacionales, por lo que este camino
  queda implementado segun formula, pendiente de validar con datos.)
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd


class Vector:
    """Vector de precios del corte (id -> precio decimal)."""

    def __init__(self, ruta: str | Path = "insumos/vector_precios.csv"):
        df = pd.read_csv(ruta, dtype={"id": str})
        self.precio = dict(zip(df["id"].str.strip(), df["precio"]))
        self.moneda = dict(zip(df["id"].str.strip(), df.get("moneda", "")))

    def get(self, ident: str):
        return self.precio.get(str(ident).strip())


def _signo(posicion) -> int:
    """1 = compra (+1); 2 o 'venta' = -1."""
    s = str(posicion).strip().upper()
    if s in ("1", "1.0", "COMPRA", "BUY"):
        return 1
    if s in ("2", "2.0", "VENTA", "SELL"):
        return -1
    return 1


def valorar_local(fut: pd.DataFrame, vector: Vector) -> pd.DataFrame:
    """Futuros locales (TES). `fut` con: subyacente, nominal, posicion.
    Devuelve el df con precio_vector, signo y valor_mercado.
    """
    out = fut.copy().reset_index(drop=True)
    precio, signo, vm = [], [], []
    for _, r in out.iterrows():
        sub = str(r["subyacente"]).strip()
        p = vector.get(sub)
        sg = _signo(r["posicion"])
        nom = float(r["nominal"] or 0)
        precio.append(p)
        signo.append(sg)
        vm.append(nom * p * sg if p is not None else None)
    out["precio_vector"] = precio
    out["signo"] = signo
    out["valor_mercado"] = vm
    return out


def valorar_int(fut: pd.DataFrame, vector: Vector, fx: dict | None = None) -> pd.DataFrame:
    """Futuros internacionales (indices/treasuries). `fut` con: subyacente,
    nominal, posicion, tipo (INDICE/TREASURY/...), moneda. `fx`: moneda->TRM/rate.
    """
    fx = fx or {}
    out = fut.copy().reset_index(drop=True)
    precio, signo, vm_mon, vm_cop = [], [], [], []
    for _, r in out.iterrows():
        sub = str(r["subyacente"]).strip()
        p = vector.get(sub)
        tipo = str(r.get("tipo", "")).strip().upper()
        if p is not None and tipo.startswith("TREAS"):
            p = p / 100.0 if p > 1 else p
        sg = _signo(r["posicion"])
        nom = float(r["nominal"] or 0)
        v_mon = nom * sg * p if p is not None else None
        rate = fx.get(str(r.get("moneda", "")).strip().upper(), 1.0)
        precio.append(p)
        signo.append(sg)
        vm_mon.append(v_mon)
        vm_cop.append(v_mon * rate if v_mon is not None else None)
    out["precio_vector"] = precio
    out["signo"] = signo
    out["valor_mercado_moneda"] = vm_mon
    out["valor_mercado_cop"] = vm_cop
    return out
