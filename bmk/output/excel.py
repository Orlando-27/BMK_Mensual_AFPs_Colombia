"""
Generacion de la salida Excel del consolidado del benchmark mensual.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd


def escribir_consolidado(out_dir: str | Path, corte: str, hojas: dict[str, pd.DataFrame]) -> str:
    """Escribe un Excel multi-hoja con las tablas del consolidado."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / f"consolidado_benchmark_{corte}.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        for nombre, df in hojas.items():
            if df is None or len(df) == 0:
                continue
            df.to_excel(xw, sheet_name=nombre[:31], index=False)
    return str(dest)
