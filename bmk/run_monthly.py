"""
Punto de entrada mensual (F6 - orquestacion).

Pensado para dispararse el dia 10 de cada mes (cron / Cloud Scheduler). Descarga
el corte del mes anterior de la SFC, corre el pipeline completo y deja las
salidas + el archivo de alertas en la carpeta del corte.

Uso:
    python3 -m bmk.run_monthly                 # corte del mes que corresponde hoy
    python3 -m bmk.run_monthly --mes Junio --anio 2026
"""
from __future__ import annotations

import argparse
import datetime as _dt

from bmk.config.settings import Config
from bmk.ingest.sfc import MESES
from bmk import pipeline


def _mes_corte_hoy() -> tuple[str, int]:
    """El corte a procesar es el mes ANTERIOR al actual (se publica el dia 10)."""
    hoy = _dt.date.today()
    m = hoy.month - 1 or 12
    anio = hoy.year if hoy.month > 1 else hoy.year - 1
    return MESES[m - 1], anio


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mes")
    ap.add_argument("--anio", type=int)
    ap.add_argument("--archivo", help="Insumo SFC local (omite descarga)")
    ap.add_argument("--swaps-insumos", help="Carpeta de curvas para valorar swaps")
    args = ap.parse_args()

    if args.mes and args.anio:
        mes, anio = args.mes, args.anio
    else:
        mes, anio = _mes_corte_hoy()

    cfg = Config(anio=anio, mes=mes)
    if args.swaps_insumos:
        from pathlib import Path
        cfg.dir_swaps_insumos = Path(args.swaps_insumos)

    print(f"=== Benchmark Mensual: corte {mes} {anio} ({cfg.corte}) ===")
    resumen = pipeline.correr(cfg, archivo=args.archivo)
    a = resumen["alertas"]
    print(f"\nOK. Alertas: {a.get('total', 0)} "
          f"(instrumentos nuevos y controles en {resumen['salidas']['xlsx']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
