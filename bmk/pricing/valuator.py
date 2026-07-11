"""
Wrapper del motor de valoracion de swaps v6 (bmk/pricing/swaps).

El motor v6 es un artefacto validado (no se modifica); lee su configuracion de
variables de entorno al importarse, por lo que se ejecuta como subproceso con
un directorio de insumos (INICIO_*, SWAPIND_*, curva_*_*) y devuelve el
resultado (VPN/DM por swap) como DataFrame.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pandas as pd

MOTOR = Path(__file__).parent / "swaps" / "swap_valuator_v6.py"


def valorar_swaps(insumos_dir: str | Path, out_dir: str | Path = "salidas/swaps") -> pd.DataFrame:
    """Corre el motor v6 sobre un directorio de insumos de una fecha.

    insumos_dir debe contener INICIO_YYYYMMDD.csv, SWAPIND_YYYYMMDD.csv y las
    curvas curva_*_YYYYMMDD.csv (extraidas con extraer_insumos.py de la
    Calculadora Swap del corte). Devuelve el CSV detallado (salida_full).
    """
    insumos_dir = Path(insumos_dir)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_inicio = out_dir / "valoracion_inicio.csv"
    out_full = out_dir / "valoracion_full.csv"
    env = dict(os.environ)
    env["V6_INSUMOS_DIR"] = str(insumos_dir)
    env["V6_OUTPUT_CSV"] = str(out_inicio)
    env["V6_OUTPUT_FULL"] = str(out_full)
    r = subprocess.run([sys.executable, str(MOTOR)], env=env,
                       capture_output=True, text=True, timeout=900)
    if r.returncode != 0:
        raise RuntimeError(f"Motor v6 fallo:\n{r.stdout}\n{r.stderr}")
    return pd.read_csv(out_full)
