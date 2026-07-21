"""
Genera la VERSION MACRO de la hoja Benchmark: parte de nuestra hoja (valoracion
propia con curvas INFOVALMER + capitalizacion Precia) y reemplaza SOLO el valor
de mercado de las filas de SWAP por el que reporta la hoja Benchmark del macro
(la Calculadora Swap), casando por identificador (ISIN interno).

Motivo: nuestro motor casa con la SFC/Precia a <1% por pata, pero el macro usa
las curvas de su propia Calculadora (extraidas del .xlsb), que difieren ~1% de
las de INFOVALMER, y ademas NO capitaliza los IRS que capitalizan interes. Para
CONCILIAR contra el Benchmark del macro (diferencia -> 0 en swaps), se toman sus
valores de swap. El resto de la hoja (activos, caja, forwards) queda igual.

Produce hoja_Benchmark_<corte>_macro.xlsx.

Uso:
    python3 tools/hoja_version_macro.py \
        --nuestra salidas/2026-06-30/hoja_Benchmark_2026-06-30.xlsx \
        --macro   insumos/Benchmark_corte_junio.xlsx \
        --corte 2026-06-30 --out salidas/2026-06-30
"""
from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


def _norm(s) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(s))
                   if not unicodedata.combining(c)).upper().strip()


def leer_macro_swaps(path: str) -> dict:
    """{ (identificador, pata_signo_idx) : vr } no; mas simple: lista por identificador
    de las filas de swap del macro, en orden, con su vr (con signo)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Benchmark"]
    filas = []
    for r in ws.iter_rows(min_row=15, values_only=True):
        if r[0] is None and r[1] is None:
            continue
        clasif = _norm(r[17]) if len(r) > 17 and r[17] else ""
        if clasif == "SWAP":
            ident = str(r[2]).strip() if r[2] is not None else ""   # IDENTIFICADOR
            nemo = str(r[3]).strip() if r[3] is not None else ""    # SWAP TF / SWAP TV
            vr = pd.to_numeric(r[11], errors="coerce")
            filas.append((ident, nemo, vr))
    wb.close()
    return filas


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--nuestra", required=True, help="Nuestra hoja Benchmark (version Precia)")
    ap.add_argument("--macro", required=True, help="Hoja Benchmark del macro (referencia)")
    ap.add_argument("--corte", required=True)
    ap.add_argument("--out", required=True, help="Carpeta de salida")
    a = ap.parse_args()

    nuestra = pd.read_excel(a.nuestra, sheet_name="Benchmark")
    macro = leer_macro_swaps(a.macro)
    # Se empareja por IDENTIFICADOR (no por rotulo TF/TV, que puede diferir entre
    # nuestra hoja y el macro). Cada swap tiene 2 patas; se asignan en orden dentro
    # del identificador, con lo que el NETO por swap queda igual al del macro.
    from collections import defaultdict, deque
    cola = defaultdict(deque)
    for ident, nemo, vr in macro:
        cola[ident].append(vr)

    out = nuestra.copy()
    es_swap = out["CLASIFICACIÓN"].astype(str).str.upper().str.strip() == "SWAP"
    col_vr = "Vr Mercado INICIAL"
    ident_col = out["IDENTIFICADOR"].astype(str).str.strip()
    n_reemp = 0
    faltan = 0
    for i in out.index[es_swap]:
        key = ident_col[i]
        if cola[key]:
            out.at[i, col_vr] = cola[key].popleft()
            n_reemp += 1
        else:
            # swap nuestro que el macro no trae: se deja nuestro valor (Precia). No
            # se pone 0 para no perder swaps que solo existen de nuestro lado.
            faltan += 1

    dest = Path(a.out) / f"hoja_Benchmark_{a.corte}_macro.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        out.to_excel(xw, sheet_name="Benchmark", index=False)

    # Reporte de cuadre de swaps
    sw_nuestra = pd.to_numeric(nuestra.loc[es_swap, col_vr], errors="coerce").sum()
    sw_macro = pd.to_numeric(out.loc[es_swap, col_vr], errors="coerce").sum()
    print(f"Swaps reemplazados: {n_reemp}  | sin match en macro: {faltan}")
    print(f"SWAP nuestra (Precia): {sw_nuestra/1e6:,.1f} MM  ->  SWAP macro: {sw_macro/1e6:,.1f} MM")
    print(f"-> {dest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
