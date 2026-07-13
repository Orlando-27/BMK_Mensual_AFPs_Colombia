"""
Corre el proceso COMPLETO de un corte con las curvas reales de INFOVALMER:
  1. Convierte las curvas forward (Fwd_*.txt) -> insumos/fwd_curves (reales).
  2. Valora los swaps con el motor v6 (curvas SwapCC_*.txt + SWAPIND del SFC).
  3. Corre el pipeline completo, inyectando el VPN de swaps en la hoja Benchmark.

Uso:
    python3 tools/correr_corte.py --curvas insumos_diarios/2026-05-31 \
        --mes Mayo --anio 2026 --corte 2026-05-31 \
        --sfc insumos/sfc/2026_05portainvdeta.xls
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd

from bmk.config.settings import Config
from bmk.ingest import infovalmer
from bmk import pipeline
import correr_swaps  # noqa: E402  (mismo directorio tools/)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--curvas", required=True, help="Carpeta insumos_diariosAAAA-MM-DD del corte")
    ap.add_argument("--corte", required=True, help="AAAA-MM-DD (fecha de valoracion)")
    ap.add_argument("--mes", required=True)
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--sfc", required=True, help="Archivo SFC del corte")
    a = ap.parse_args()

    print("=== 1. Curvas forward reales -> insumos/fwd_curves ===")
    infovalmer.preparar_forwards(a.curvas, a.corte, "insumos/fwd_curves")
    print("    ok (FWPCOP/FWTCOP/LIBBTS + puntos por par + paridades)")

    print("=== 2. Valoracion de swaps (motor v6) ===")
    res = correr_swaps.correr(a.curvas, a.corte, a.sfc, f"salidas/swaps/{a.corte}")
    vpn = res[["ISIN", "VPN_Derecho_Calc", "VPN_Oblig_Calc"]] if "ISIN" in res.columns else None

    print("=== 3. Pipeline completo (con VPN de swaps inyectado) ===")
    cfg = Config(anio=a.anio, mes=a.mes)
    resumen = pipeline.correr(cfg, archivo=a.sfc, vpn_swaps=vpn)
    print(f"\n[corte] salidas en {cfg.dir_corte()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
