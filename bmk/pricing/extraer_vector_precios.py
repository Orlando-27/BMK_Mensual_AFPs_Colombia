"""
Extrae la hoja `VECTOR DE PRECIOS` del Benchmark (Detallado) a un CSV swappable.

Es el vector de precios del corte: ISIN/NEMO/codigo de contrato -> precio (en
forma decimal, p. ej. TES 0.92411 = 92,411%). Lo usan los valoradores de futuros
locales (precio del subyacente) e internacionales.

Salida: <dir>/vector_precios.csv (id, precio, moneda).
Uso: python3 -m bmk.pricing.extraer_vector_precios "insumos/Benchmark (Detallado).xlsb"
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


def extraer(xlsb_path: str | Path, out_dir: str | Path = "insumos") -> int:
    import pyxlsb
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = pyxlsb.open_workbook(str(xlsb_path))
    filas = []
    with wb.get_sheet("VECTOR DE PRECIOS") as sh:
        for i, row in enumerate(sh.rows()):
            if i == 0:
                continue
            v = [c.v for c in row]
            if v and v[0] is not None and len(v) > 1 and isinstance(v[1], (int, float)):
                moneda = v[3] if len(v) > 3 and v[3] is not None else ""
                filas.append((str(v[0]).strip(), float(v[1]), str(moneda).strip()))
    df = pd.DataFrame(filas, columns=["id", "precio", "moneda"]).drop_duplicates("id")
    df.to_csv(out_dir / "vector_precios.csv", index=False)
    return len(df)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "insumos/Benchmark (Detallado).xlsb"
    n = extraer(src)
    print(f"VECTOR DE PRECIOS extraido: {n} precios -> insumos/vector_precios.csv")
