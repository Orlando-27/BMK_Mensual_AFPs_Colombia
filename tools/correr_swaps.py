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
    """Ajusta los swaps de camara (IRS SOFR, CRCC liquidados a diario) a su
    VARIACION DIARIA = MtM(corte) - MtM(dia habil anterior), que es el valor que
    reporta la SFC (Formato_415) para estos. El resultado se coloca en la pata
    derecho si es positivo (obligacion=0) o en obligacion si es negativo, como lo
    reporta la SFC (una pata en cero). Los demas swaps quedan al corte, sin tocar.

    res, res_prev: salidas del motor v6 (mismo universo de swaps) al corte y al
    dia habil anterior, respectivamente."""
    prev = _mtm(res_prev)
    out = res.copy()
    tipo = out["Tipo_Swap"].astype(str).str.upper()
    es_cam = tipo.str.contains("SOF")  # IRS SOFR = camara
    der = pd.to_numeric(out["VPN_Derecho_Calc"], errors="coerce").fillna(0.0)
    obl = pd.to_numeric(out["VPN_Oblig_Calc"], errors="coerce").fillna(0.0)
    n = 0
    for i in out.index[es_cam]:
        isin = str(out.at[i, "ISIN"])
        var = (der[i] - obl[i]) - prev.get(isin, der[i] - obl[i])
        out.at[i, "VPN_Derecho_Calc"] = var if var >= 0 else 0.0
        out.at[i, "VPN_Oblig_Calc"] = 0.0 if var >= 0 else -var
        n += 1
    print(f"[swaps] camara (IRS SOFR): {n} swaps ajustados a variacion diaria")
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
