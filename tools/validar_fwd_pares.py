"""
Valida mi valoracion de forwards contra la hoja 'Fwd Industria' del Benchmark
macro, por PARIDAD, para derivar la escala correcta de los puntos forward.

El valorador arma Tasa Forward = (spot + puntos) x DF. Si los puntos de un par
vienen en otra unidad (pips 1e-4, sen 1e-2, absoluto 1), la Tasa Forward se
descuadra. Comparando mi Tasa Forward (y el P&G) contra la del macro, por par,
sale el factor de escala que falta.

Uso:
    python3 tools/validar_fwd_pares.py \
        --mio salidas/2026-05-31/valoracion_fwds_industria_2026-05-31.xlsx \
        --macro "/ruta/Benchmark junio.xlsx"
"""
from __future__ import annotations

import argparse

import pandas as pd


def _leer_fwd(path: str) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    hoja = next((h for h in xl.sheet_names if h.strip().lower() == "fwd industria"), None)
    if hoja is None:
        raise SystemExit(f"No hay hoja 'Fwd Industria' en {path}. Hojas: {xl.sheet_names}")
    # El macro trae PRIMERO una matriz de consolidacion (posicion por paridad x
    # portafolio) y MAS ABAJO la tabla de valoracion forward-por-forward. Se busca
    # el encabezado de esa tabla: la fila que tiene STRIKE (firma unica de la tabla
    # contrato por contrato, no de la matriz de consolidacion).
    crudo = pd.read_excel(path, sheet_name=hoja, header=None, nrows=120)
    hrow = None
    for i in range(len(crudo)):
        vals = [str(x).strip().upper() for x in crudo.iloc[i].tolist()]
        if any("STRIKE" in v for v in vals) and any(v == "PARIDAD" or "PARIDAD" in v for v in vals):
            hrow = i
            break
    if hrow is None:  # fallback: primera fila con PARIDAD
        for i in range(len(crudo)):
            vals = [str(x).strip().upper() for x in crudo.iloc[i].tolist()]
            if any("PARIDAD" in v for v in vals):
                hrow = i
                break
    df = pd.read_excel(path, sheet_name=hoja, header=hrow or 0)
    df.columns = [str(c).strip().upper() for c in df.columns]

    def col(*alts):
        for a in alts:
            for c in df.columns:
                if a in c:
                    return c
        return None
    m = {
        "paridad": col("PARIDAD"),
        "spot": col("SPOT"),
        "tasa_fwd": col("TASA FORWARD"),
        "tasa_vpn": col("TASA VPN"),
        "pyg": col("P&G", "PYG", "P Y G"),
        "nominal": col("NOMINAL", "VALOR NOMINAL"),
    }
    faltan = [k for k, v in m.items() if v is None]
    if faltan:
        raise SystemExit(f"Faltan columnas {faltan} en {path}. Columnas: {list(df.columns)}")
    out = pd.DataFrame({k: df[v] for k, v in m.items()})
    for c in ["spot", "tasa_fwd", "tasa_vpn", "pyg", "nominal"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out["paridad"] = out["paridad"].astype(str).str.strip().str.upper()
    return out.dropna(subset=["paridad", "spot"])


def _resumen(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["pts"] = df["tasa_fwd"] - df["spot"]   # puntos (aprox, sin desagregar DF)
    return df.groupby("paridad").agg(
        n=("paridad", "size"),
        spot=("spot", "median"),
        tasa_fwd_med=("tasa_fwd", "median"),
        pts_med=("pts", "median"),
        pyg_MM=("pyg", lambda s: s.sum() / 1e6),
    )


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mio", required=True)
    ap.add_argument("--macro", required=True)
    a = ap.parse_args()

    mio = _resumen(_leer_fwd(a.mio))
    mac = _resumen(_leer_fwd(a.macro))
    t = mio.join(mac, how="outer", lsuffix="_mio", rsuffix="_mac")
    # Factor de escala sugerido: cuanto hay que multiplicar MIS puntos para igualar
    # los del macro (si mis puntos estan en pips y los del macro en decimal, ~1e-4).
    t["escala_pts"] = t["pts_med_mac"] / t["pts_med_mio"]
    pd.options.display.float_format = lambda x: f"{x:,.6g}"
    print("=== Por PARIDAD: spot / Tasa Forward / puntos (mio vs macro) ===")
    cols = ["spot_mio", "tasa_fwd_med_mio", "pts_med_mio", "tasa_fwd_med_mac",
            "pts_med_mac", "escala_pts", "pyg_MM_mio", "pyg_MM_mac"]
    print(t[cols].to_string())
    print("\nLectura: 'escala_pts' ~ el factor que falta aplicar a mis puntos.")
    print("  ~1     -> par correcto")
    print("  ~1e-4  -> puntos en pips (dividir /10000)")
    print("  ~1e-2  -> puntos en sen/centavos (dividir /100)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
