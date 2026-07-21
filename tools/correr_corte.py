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
                    "(para revaluar solo forwards). Si se omite, se usa el corte.")
    ap.add_argument("--fecha-pub", help="AAAA-MM-DD de publicacion (fecha de valoracion de forwards)")
    ap.add_argument("--curvas-val", help="Curvas de la fecha de VALORACION de derivados "
                    "(swaps Y forwards). Usa esto cuando 'todo esta valorado' a esa fecha.")
    ap.add_argument("--fecha-val", help="AAAA-MM-DD de valoracion de derivados (swaps Y forwards)")
    ap.add_argument("--curvas-prev", help="Curvas del DIA HABIL ANTERIOR al corte, para "
                    "valorar la variacion diaria de los swaps de camara (IRS SOFR).")
    ap.add_argument("--fecha-prev", help="AAAA-MM-DD del dia habil anterior (swaps de camara)")
    a = ap.parse_args()

    # Fechas de valoracion de derivados. Con --curvas-val/--fecha-val, swaps Y
    # forwards se valoran a esa fecha (caso 'todo valorado al dia de publicacion').
    # Si no, comportamiento previo: swaps al corte; forwards a --fecha-pub o corte.
    swap_curvas = a.curvas_val or a.curvas
    swap_fecha = a.fecha_val or a.corte
    fwd_curvas = a.curvas_val or a.curvas_pub or a.curvas
    fwd_fecha = a.fecha_val or a.fecha_pub or a.corte
    print(f"=== 1. Curvas forward ({fwd_fecha}) -> insumos/fwd_curves ===")
    infovalmer.preparar_forwards(fwd_curvas, fwd_fecha, "insumos/fwd_curves")
    print("    ok (FWTCOP/LIBBTS + paridades)")

    print(f"=== 2. Valoracion de swaps (motor v6, fecha {swap_fecha}) ===")
    res = correr_swaps.correr(swap_curvas, swap_fecha, a.sfc, f"salidas/swaps/{a.corte}")
    # `res` = VPN COMPLETO por pata (igual que la Calculadora Swap del Excel).
    # Camara (CRCC): los IRS SOFR se liquidan a diario -> al Benchmark se inyecta su
    # variacion diaria MtM(corte) - MtM(dia habil anterior), NO el VPN completo. El
    # VPN completo se conserva para el detalle swap-por-swap.
    res_iny = res
    if a.curvas_prev and a.fecha_prev:
        print(f"    valorando dia habil anterior ({a.fecha_prev}) para variacion de camara ...")
        res_prev = correr_swaps.correr(a.curvas_prev, a.fecha_prev, a.sfc,
                                       f"salidas/swaps/{a.corte}_prev")
        res_iny = correr_swaps.ajustar_camara(res, res_prev)  # no muta res (copia)
    _cols_vpn = ["ISIN", "VPN_Derecho_Calc", "VPN_Oblig_Calc"]
    for _pc in ("Precio_Der", "Precio_Obl"):  # precio limpio por pata (para la hoja)
        if _pc in res_iny.columns:
            _cols_vpn.append(_pc)
    vpn = res_iny[_cols_vpn] if "ISIN" in res_iny.columns else None
    # Output swap-por-swap: VPN COMPLETO (res) + columnas con la variacion de camara
    # (res_iny) que es lo que se inyecta al Benchmark para los IRS SOFR.
    try:
        from bmk.ingest import sfc as _sfc
        from bmk.pricing.swaps import sabana as _sab
        from bmk.output import valoracion_swaps as _vsw
        _cfg0 = Config(anio=a.anio, mes=a.mes)
        _sabana = _sab.construir(_sfc.leer(a.sfc).formato_415)
        _cam = res_iny if res_iny is not res else None
        _ruta = _vsw.escribir(res, _sabana, _cfg0.dir_corte(), a.corte, camara=_cam)
        print(f"    valoracion swap-por-swap -> {_ruta}")
    except Exception as e:  # noqa: BLE001
        print(f"    (aviso) no se pudo escribir la valoracion detallada de swaps: {str(e)[:120]}")

    print("=== 3. Pipeline completo (VPN swaps + forwards revaluados) ===")
    cfg = Config(anio=a.anio, mes=a.mes)
    resumen = pipeline.correr(cfg, archivo=a.sfc, vpn_swaps=vpn, fecha_fwd=fwd_fecha)
    print(f"\n[corte] salidas en {cfg.dir_corte()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
