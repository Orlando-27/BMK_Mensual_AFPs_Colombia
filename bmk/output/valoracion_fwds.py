"""
Output "Valoracion Fwds Industria": valoracion forward por forward de la
industria, replicando la hoja `Fwd Industria` del Benchmark (Detallado).

Dos entradas posibles de forwards:
  - Formato_415 del archivo SFC (produccion, via normalizar_415).
  - La propia hoja `Fwd Industria` del .xlsb (validacion).

La valoracion la hace bmk.pricing.fwd_valuator con las curvas de insumos/fwd_curves
(extraidas de la hoja `Curvas`; ver bmk.pricing.extraer_curvas_fwd).

Escribe un Excel con:
  - hoja "Detalle": una fila por forward con tasa VPN/forward, valor COP de cada
    pata y P&G.
  - hoja "Resumen": P&G y posicion neta por AFP / portafolio / divisa.
"""
from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pandas as pd

_EPOCH = _dt.date(1899, 12, 30)  # dia base del serial de Excel


def yyyymmdd_a_serial(v) -> float | None:
    """Convierte una fecha AAAAMMDD (formato Formato_415) a serial de Excel."""
    try:
        n = int(v)
    except (TypeError, ValueError):
        return None
    if n < 10000101 or n > 99991231:
        return None
    try:
        return float((_dt.date(n // 10000, (n // 100) % 100, n % 100) - _EPOCH).days)
    except ValueError:
        return None


def corte_a_serial(corte: str) -> int:
    """'AAAA-MM-DD' -> serial de Excel."""
    d = _dt.date.fromisoformat(corte)
    return (d - _EPOCH).days

from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp, _cod_portafolio
from bmk.pricing.fwd_valuator import Curvas, valorar, PARES, SPOT_MONEDA

# Monedas COP-equivalentes: si una pata es COP el par es USDCOP-style local.
_MON = {"COU": "UVR"}


def _cur(s):
    return _MON.get(str(s).strip().upper(), str(s).strip().upper())


def _paridad(mder: str, mobl: str) -> str | None:
    """Construye la paridad (p. ej. USDCOP, USDBRL, EURUSD) desde las patas."""
    d, o = _cur(mder), _cur(mobl)
    for par in PARES:
        a, b = par[:3], par[3:]
        if {d, o} == {a, b}:
            return par
    return None


def normalizar_415(formato_415: pd.DataFrame, solo_industria: bool = True,
                   tipo_derivado: int = 1) -> pd.DataFrame:
    """Forwards de Formato_415 al esquema del valorador. tipo_derivado=1 forwards,
    =4 futuros de TRM (se valoran con el mismo motor)."""
    def c(k):
        return formato_415.iloc[:, COL_415[k]]
    tipo = pd.to_numeric(c("tipo_derivado"), errors="coerce")
    df = pd.DataFrame({
        "afp": c("entidad").map(normalizar_afp),
        "portafolio": c("nombre").map(_cod_portafolio),
        "mder": c("moneda_derecho").map(_cur),
        "nder": pd.to_numeric(c("nominal_derecho"), errors="coerce"),
        "mobl": c("moneda_obligacion").map(_cur),
        "nobl": pd.to_numeric(c("nominal_obligacion"), errors="coerce"),
        "strike": pd.to_numeric(c("strike"), errors="coerce"),
        # Formato_415 trae la fecha en AAAAMMDD -> convertir a serial de Excel.
        "fecha_cumplimiento": pd.to_numeric(c("fecha_liquidacion"), errors="coerce").map(yyyymmdd_a_serial),
    })
    df = df[tipo == tipo_derivado].reset_index(drop=True)
    if solo_industria:
        df = df[df["afp"] != "COLFONDOS"].reset_index(drop=True)
    df["paridad"] = [_paridad(d, o) for d, o in zip(df["mder"], df["mobl"])]
    # nominal = pata en la moneda extranjera del par (no la COP); operacion VENTA
    # si se entrega (obligacion) la divisa extranjera y se recibe COP/base.
    mons = df["paridad"].map(lambda p: SPOT_MONEDA.get(p) if p else None)
    def _nom(row):
        if row["paridad"] is None:
            return None
        base = SPOT_MONEDA.get(row["paridad"])  # divisa cuyo nominal reporta la hoja
        return row["nder"] if row["mder"] == base else row["nobl"]
    df["nominal"] = df.apply(_nom, axis=1)
    # OPERACION: COMPRA si el derecho es la divisa base (se recibe la divisa), VENTA si se entrega.
    df["operacion"] = [("COMPRA" if md == SPOT_MONEDA.get(p) else "VENTA") if p else None
                       for md, p in zip(df["mder"], df["paridad"])]
    return df[df["paridad"].notna()].reset_index(drop=True)


def generar(forwards: pd.DataFrame, curvas: Curvas, fecha_valoracion: int):
    """Devuelve (detalle, resumen) valorando cada forward."""
    val = valorar(forwards, curvas, fecha_valoracion)
    resumen = None
    if {"afp", "portafolio"}.issubset(val.columns):
        resumen = (val.groupby(["afp", "portafolio", "paridad"], dropna=False)
                   .agg(n=("pyg", "size"),
                        valor_der=("valor_cop_der", "sum"),
                        valor_obli=("valor_cop_obli", "sum"),
                        pyg=("pyg", "sum")).reset_index())
    return val, resumen


def escribir(forwards: pd.DataFrame, out_dir: str | Path, corte: str,
             dir_curvas: str | Path = "insumos/fwd_curves",
             fecha_valoracion: int | None = None) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    curvas = Curvas(dir_curvas)
    if fecha_valoracion is None:
        try:
            fecha_valoracion = corte_a_serial(corte)
        except (ValueError, TypeError):
            fv = pd.to_numeric(forwards["fecha_cumplimiento"], errors="coerce")
            fecha_valoracion = int(fv.min()) if fv.notna().any() else 0
    detalle, resumen = generar(forwards, curvas, fecha_valoracion)
    dest = out_dir / f"valoracion_fwds_industria_{corte}.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        detalle.to_excel(xw, sheet_name="Detalle", index=False)
        if resumen is not None:
            resumen.to_excel(xw, sheet_name="Resumen", index=False)
    return str(dest)
