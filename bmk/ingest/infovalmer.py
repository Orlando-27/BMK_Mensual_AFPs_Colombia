"""
Conversor de los planos diarios de INFOVALMER (carpetas insumos_diariosAAAA-MM-DD)
al formato que consumen los motores de valoracion.

Formatos de entrada (una carpeta por fecha):
  - SwapCC_<X>_Diaria_AAAAMMDD.txt : "plazo tasa" (espacio), tasa en decimal.
  - Fwd_<PAR>_Diaria_AAAAMMDD.txt  : "plazo puntos" (espacio).
  - AAAA-MM-DD IND.csv             : "cod;nombre;fecha;unidad;valor;..." (INFOVALMER).
  - Matriz_TC_AAAAMMDD.txt         : matriz de cruces (fila 1 = monedas; luego cada fila).

Produce:
  A) Insumos del motor de swaps v6 (curva_*.csv 'Dias,Tasa' + INICIO_*.csv).
  B) Curvas del valorador de forwards (fwd_curves/: FWPCOP, FWTCOP, LIBBTS,
     puntos por par y paridades.csv).
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# SwapCC_<X> (INFOVALMER) -> curva_<Y> (motor v6). Solo las que usa el motor.
SWAPCC_A_CURVA = {
    "IBRColateralUSD": "IBR_COLATERAL_USD", "COPColateralUSD": "COP_COLATERAL_USD",
    "USDOIS": "USDOIS", "USDCO": "USDCO", "EURCOP": "EURCOP", "IBRUVR": "IBRUVR",
    "DTF": "DTF", "LIBORUVR": "LIBORUVR", "LIBORCOP": "LIBORCOP", "IBR": "IBR",
}
# Fwd_<PAR> (INFOVALMER) -> curva de puntos del valorador de forwards.
FWD_A_PUNTOS = {
    "USDCOP": "FWPCOP", "USDBRL": "FWPBRL", "USDMXN": "FWPMXN", "USDJPY": "FWPJPY",
    "USDCLP": "FWPCLP", "USDCHF": "FWPCHF", "USDCAD": "FWPCAD", "EURUSD": "FWPEUR",
    "GBPUSD": "Fwd_GBPUSD_Diaria",
}
# Paridades (valorador de forwards): moneda -> como leer el spot de Matriz_TC.
#   ("USD", "X")   -> Matriz[USD][X]   (USDXXX: X por USD)
#   ("X",  "USD")  -> Matriz[X][USD]   (XXXUSD: USD por X)
PARIDAD_MATRIZ = {
    "USD": ("USD", "COP"), "BRL": ("USD", "BRL"), "MXN": ("USD", "MXN"),
    "JPY": ("USD", "JPY"), "CLP": ("USD", "CLP"), "CHF": ("USD", "CHF"),
    "CAD": ("USD", "CAD"), "EUR": ("EUR", "USD"), "GBP": ("GBP", "USD"),
    "AUD": ("AUD", "USD"),
}


def _yyyymmdd(fecha: str) -> str:
    return fecha.replace("-", "")


def _leer_txt_curva(path: Path) -> pd.DataFrame:
    """Lee un archivo de curva separado por espacios y toma la 1a columna como
    Dias (plazo) y la 2a como Tasa (mid). Los SwapCC_* traen 2 columnas
    (plazo, tasa); los Fwd_<par> traen 4 (plazo, mid, bid, ask) -> se ignoran
    bid/ask. (Antes se usaba names=[Dias,Tasa], que con 4 columnas tomaba las
    ULTIMAS dos = bid/ask y perdia el plazo, rompiendo la interpolacion.)"""
    df = pd.read_csv(path, sep=r"\s+", header=None, engine="python")
    out = pd.DataFrame({
        "Dias": pd.to_numeric(df.iloc[:, 0], errors="coerce"),
        "Tasa": pd.to_numeric(df.iloc[:, 1], errors="coerce"),
    })
    return out.dropna()


def _leer_matriz(path: Path) -> pd.DataFrame:
    """Lee Matriz_TC: fila 1 = monedas, luego filas 'MON v1 v2 ...'. Indexada por
    moneda de fila; columnas = monedas."""
    lineas = Path(path).read_text(encoding="latin-1").splitlines()
    monedas = lineas[0].split()[1:]  # el primer token es la etiqueta 'Mid'
    filas = {}
    for ln in lineas[1:]:
        p = ln.split()
        if not p:
            continue
        filas[p[0]] = [float(x) for x in p[1:]]
    m = pd.DataFrame.from_dict(filas, orient="index", columns=monedas)
    return m


def _indicadores(src: Path, fecha: str) -> dict:
    """TRM, UVR, EURCOP de una fecha (de Matriz_TC + IND.csv)."""
    ym = _yyyymmdd(fecha)
    m = _leer_matriz(src / f"Matriz_TC_{ym}.txt")
    trm = float(m.loc["USD", "COP"])
    eurcop = float(m.loc["EUR", "COP"]) if "EUR" in m.index else trm * float(m.loc["EUR", "USD"])
    # UVR del IND.csv (codigo 'UVR').
    uvr = None
    ind = src / f"{fecha} IND.csv"
    if ind.exists():
        for ln in ind.read_text(encoding="latin-1").splitlines():
            c = ln.split(";")
            if len(c) > 4 and c[0].strip().upper() == "UVR":
                try:
                    uvr = float(c[4])
                except ValueError:
                    pass
                break
    return {"TRM": trm, "UVR": uvr, "EURCOP": eurcop, "matriz": m}


def preparar_swaps(src_dir, fecha: str, out_dir, swapind_csv=None) -> Path:
    """Escribe curva_*.csv + INICIO_*.csv (y opcional SWAPIND) para el motor v6."""
    src, out = Path(src_dir), Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ym = _yyyymmdd(fecha)
    for cc, curva in SWAPCC_A_CURVA.items():
        f = src / f"SwapCC_{cc}_Diaria_{ym}.txt"
        if f.exists():
            _leer_txt_curva(f).to_csv(out / f"curva_{curva}_{ym}.csv", index=False)
    ind = _indicadores(src, fecha)
    inicio = pd.DataFrame({"key": ["FECHA_VAL", "TRM", "UVR", "EURCOP"],
                           "value": [fecha, ind["TRM"], ind["UVR"], ind["EURCOP"]]})
    inicio.to_csv(out / f"INICIO_{ym}.csv", index=False)
    if swapind_csv is not None:
        Path(swapind_csv).replace(out / f"SWAPIND_{ym}.csv")
    return out


def preparar_forwards(src_dir, fecha: str, out_dir) -> Path:
    """Escribe fwd_curves/ (FWPCOP, FWTCOP, LIBBTS, puntos por par, paridades)."""
    src, out = Path(src_dir), Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ym = _yyyymmdd(fecha)

    def _puntos(par, nombre):
        f = src / f"Fwd_{par}_Diaria_{ym}.txt"
        if f.exists():
            _leer_txt_curva(f).rename(columns={"Dias": "plazo_dias", "Tasa": "valor"}
                                      ).to_csv(out / f"{nombre}.csv", index=False)
    for par, nombre in FWD_A_PUNTOS.items():
        _puntos(par, nombre)
    # Curvas de descuento de forwards (en %). Decodificadas contract-level contra
    # la hoja `Curvas` del macro (corte junio, valorado 10-jul); cada una casa al
    # decimal con su fuente INFOVALMER:
    #   FWTCOP (descuento COP de USDCOP)      <- IBR    (macro FWTCOP == IBR exacto)
    #   FWTUSD (descuento USD de USDCOP)      <- USDCO  (tfwd = spot x DF(USDCO))
    #   LIBBTS (descuento USD de los cruces)  <- USDOIS (macro LIBBTS == USDOIS)
    #   FWTEUR (descuento EUR de EURUSD)      <- EUROIS (tfwd = spot x DF(EUROIS))
    # USDCOP no usa puntos: tvpn = K x DF(IBR), tfwd = spot x DF(USDCO) (paridad
    # cubierta bi-moneda). Verificado: reproduce la hoja Fwd Industria al peso.
    for cc, nombre in (("IBR", "FWTCOP"), ("USDCO", "FWTUSD"),
                       ("USDOIS", "LIBBTS"), ("EUROIS", "FWTEUR")):
        f = src / f"SwapCC_{cc}_Diaria_{ym}.txt"
        if f.exists():
            d = _leer_txt_curva(f)
            d["Tasa"] = d["Tasa"] * 100.0
            d.rename(columns={"Dias": "plazo_dias", "Tasa": "valor"}
                     ).to_csv(out / f"{nombre}.csv", index=False)
    # Paridades desde Matriz_TC.
    ind = _indicadores(src, fecha)
    m = ind["matriz"]
    filas = []
    for mon, (a, b) in PARIDAD_MATRIZ.items():
        try:
            filas.append((mon, float(m.loc[a, b])))
        except KeyError:
            pass
    filas.append(("COP", 1.0))
    pd.DataFrame(filas, columns=["par", "spot"]).to_csv(out / "paridades.csv", index=False)
    return out
