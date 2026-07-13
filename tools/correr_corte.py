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
    ap.add_argument("--curvas-pub", help="Carpeta de curvas de la fecha de PUBLICACION "
                    "(para revaluar forwards como el macro). Si se omite, se usa el corte.")
    ap.add_argument("--fecha-pub", help="AAAA-MM-DD de publicacion (fecha de valoracion de forwards)")
    a = ap.parse_args()

    # Forwards: el macro los revalua a la fecha de PUBLICACION con las curvas de ese
    # dia. Si se pasan --curvas-pub/--fecha-pub, se preparan esas; si no, el corte.
    fwd_curvas = a.curvas_pub or a.curvas
    fwd_fecha = a.fecha_pub or a.corte
    print(f"=== 1. Curvas forward ({fwd_fecha}) -> insumos/fwd_curves ===")
    infovalmer.preparar_forwards(fwd_curvas, fwd_fecha, "insumos/fwd_curves")
    print("    ok (FWTCOP/LIBBTS + paridades)")

    print("=== 2. Valoracion de swaps (motor v6) ===")
    res = correr_swaps.correr(a.curvas, a.corte, a.sfc, f"salidas/swaps/{a.corte}")
    vpn = res[["ISIN", "VPN_Derecho_Calc", "VPN_Oblig_Calc"]] if "ISIN" in res.columns else None

    print("=== 3. Pipeline completo (VPN swaps + forwards revaluados) ===")
    cfg = Config(anio=a.anio, mes=a.mes)
    resumen = pipeline.correr(cfg, archivo=a.sfc, vpn_swaps=vpn, fecha_fwd=fwd_fecha)
    print(f"\n[corte] salidas en {cfg.dir_corte()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
