"""
Valora los swaps de la industria en una fecha, con curvas INFOVALMER.

Encadena: convierte las curvas diarias (SwapCC_*.txt) al formato del motor v6,
arma la sabana SWAPIND desde el archivo SFC, y corre el motor v6 (VPN/DM por pata).

Uso:
    python3 tools/correr_swaps.py \
        --curvas insumos_diarios/2026-05-31 --fecha 2026-05-31 \
        --sfc insumos/sfc/2026_05portainvdeta.xls --out salidas/swaps/2026-05-31
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

# Permite correr el tool directo (python3 tools/correr_swaps.py) sin PYTHONPATH.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bmk.ingest import sfc, infovalmer
from bmk.pricing.swaps import sabana
from bmk.pricing import valuator


def correr(curvas_dir: str, fecha: str, sfc_path: str, out_dir: str,
           solo_industria: bool = True) -> pd.DataFrame:
    ym = fecha.replace("-", "")
    out = Path(out_dir)
    insumos = out / "insumos"
    # Se limpian INICIO/SWAPIND/curvas previos: si el dir acumula curvas de otra
    # fecha (p. ej. una corrida anterior con --fecha-val distinta), el motor v6
    # tomaria la fecha equivocada. Cada corrida deja SOLO los insumos de `fecha`.
    if insumos.exists():
        for p in insumos.glob("*.csv"):
            p.unlink()
    insumos.mkdir(parents=True, exist_ok=True)

    # 1. Sabana SWAPIND desde el SFC -> CSV con el layout del motor.
    arch = sfc.leer(sfc_path)
    sab = sabana.construir(arch.formato_415, solo_industria=solo_industria)
    swapind = sabana.a_swapind(sab)
    swapind_csv = insumos / f"SWAPIND_{ym}.csv"
    swapind.to_csv(swapind_csv, index=False)
    print(f"[swaps] SWAPIND: {len(swapind)} swaps -> {swapind_csv.name}")

    # 2. Convertir curvas INFOVALMER -> curva_*.csv + INICIO en la misma carpeta.
    infovalmer.preparar_swaps(curvas_dir, fecha, insumos)
    curvas = sorted(p.name for p in insumos.glob("curva_*.csv"))
    print(f"[swaps] curvas: {len(curvas)} -> {curvas}")

    # 3. Motor v6.
    res = valuator.valorar_swaps(insumos, out)
    print(f"[swaps] valorados: {len(res)} swaps")
    if "VPN_Derecho_Calc" in res.columns:
        vpn = pd.to_numeric(res["VPN_Derecho_Calc"], errors="coerce").sum() \
            - pd.to_numeric(res["VPN_Oblig_Calc"], errors="coerce").sum()
        print(f"[swaps] VPN neto total: {vpn:,.0f} COP")
    if "Error" in res.columns:
        n_err = res["Error"].notna().sum()
        if n_err:
            print(f"[swaps] con error: {n_err} (ver salida_full)")
    return res


def _mtm(res: pd.DataFrame) -> dict:
    """MtM neto por ISIN (VPN derecho - VPN obligacion)."""
    der = pd.to_numeric(res["VPN_Derecho_Calc"], errors="coerce")
    obl = pd.to_numeric(res["VPN_Oblig_Calc"], errors="coerce")
    return dict(zip(res["ISIN"].astype(str), (der - obl).fillna(0.0)))


def ajustar_camara(res: pd.DataFrame, res_prev: pd.DataFrame) -> pd.DataFrame:
    """Los IRS SOFR se compensan en camara (CRCC) y se reportan con 1 dia habil de
    rezago; por eso se valoran al DIA HABIL ANTERIOR. Se reemplaza el VPN COMPLETO
    por pata (derecho y obligacion) por el del dia anterior (res_prev). Los demas
    swaps quedan al corte, sin tocar. NO se usa la variacion diaria: el Benchmark
    lleva el VPN completo (como la Calculadora / la hoja del macro).

    res, res_prev: salidas del motor v6 (mismo universo de swaps) al corte y al
    dia habil anterior, respectivamente."""
    prev_der = dict(zip(res_prev["ISIN"].astype(str),
                        pd.to_numeric(res_prev["VPN_Derecho_Calc"], errors="coerce")))
    prev_obl = dict(zip(res_prev["ISIN"].astype(str),
                        pd.to_numeric(res_prev["VPN_Oblig_Calc"], errors="coerce")))
    out = res.copy()
    es_cam = out["Tipo_Swap"].astype(str).str.upper().str.contains("SOF")  # IRS SOFR
    n = 0
    for i in out.index[es_cam]:
        isin = str(out.at[i, "ISIN"])
        if isin in prev_der:
            out.at[i, "VPN_Derecho_Calc"] = prev_der[isin]
            out.at[i, "VPN_Oblig_Calc"] = prev_obl[isin]
            n += 1
    print(f"[swaps] camara (IRS SOFR): {n} swaps con VPN completo al dia anterior")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--curvas", required=True, help="Carpeta con SwapCC_*.txt / IND / Matriz_TC")
    ap.add_argument("--fecha", required=True, help="AAAA-MM-DD")
    ap.add_argument("--sfc", required=True, help="Archivo SFC del corte")
    ap.add_argument("--out", required=True, help="Carpeta de salida")
    a = ap.parse_args()
    res = correr(a.curvas, a.fecha, a.sfc, a.out)
    res.to_csv(Path(a.out) / f"valoracion_swaps_{a.fecha}.csv", index=False)
    print(f"[swaps] -> {a.out}/valoracion_swaps_{a.fecha}.csv")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
