"""
Comparacion COMPLETA de la hoja Benchmark generada por el pipeline contra el
Benchmark de las macros (ground truth), incluyendo derivados.

Compara, por CLASIFICACION y por AFP: numero de filas y valor de mercado
(columna 'Vr Mercado INICIAL'). Cubre activos, CAJA, forwards, opciones y swaps.

Uso:
    python3 tools/comparar_todo.py --macro "Benchmark junio.xlsx" \
        --nuestro salidas/2026-05-31/hoja_Benchmark_2026-05-31.xlsx
"""
from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


def _afp(s) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).upper().strip()
    if any(k in s for k in ("OLD MUTUAL", "SKANDIA", "ACCAI")):
        return "SKANDIA"
    for k in ("PORVENIR", "PROTECCION", "COLFONDOS"):
        if k in s:
            return k
    return s


def leer_macro(path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Benchmark"]
    rows = [r for r in ws.iter_rows(min_row=15, values_only=True)
            if not (r[0] is None and r[1] is None)]
    wb.close()
    return pd.DataFrame({
        "afp": [_afp(r[1]) for r in rows],
        "clasif": [str(r[17]).upper().strip() if len(r) > 17 and r[17] else "" for r in rows],
        "vr": [pd.to_numeric(r[11], errors="coerce") if len(r) > 11 else None for r in rows],
    })


def leer_nuestro(path: str) -> pd.DataFrame:
    o = pd.read_excel(path, sheet_name="Benchmark")
    return pd.DataFrame({
        "afp": o["AFP"].map(_afp),
        "clasif": o["CLASIFICACIÓN"].astype(str).str.upper().str.strip(),
        "vr": pd.to_numeric(o["Vr Mercado INICIAL"], errors="coerce"),
    })


def _tabla(macro, nuestro, campo):
    m = macro.groupby(campo).agg(n_macro=("vr", "size"), vr_macro=("vr", "sum"))
    o = nuestro.groupby(campo).agg(n_nuestro=("vr", "size"), vr_nuestro=("vr", "sum"))
    t = m.join(o, how="outer").fillna(0)
    t["dif_n"] = t["n_nuestro"] - t["n_macro"]
    t["dif_vr_%"] = ((t["vr_nuestro"] - t["vr_macro"]) / t["vr_macro"].abs().replace(0, float("nan")) * 100)
    return t.reset_index()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--macro", required=True)
    ap.add_argument("--nuestro", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()
    macro = leer_macro(a.macro)
    nuestro = leer_nuestro(a.nuestro)
    print(f"Macro:   {len(macro)} filas | Nuestro: {len(nuestro)} filas\n")

    por_clasif = _tabla(macro, nuestro, "clasif")
    por_afp = _tabla(macro, nuestro, "afp")
    pd.options.display.float_format = lambda x: f"{x:,.1f}"
    print("=== Por CLASIFICACION (valor en millones) ===")
    p = por_clasif.copy()
    p["vr_macro"] /= 1e6
    p["vr_nuestro"] /= 1e6
    print(p.to_string(index=False))
    print("\n=== Por AFP (valor en millones) ===")
    q = por_afp.copy()
    q["vr_macro"] /= 1e6
    q["vr_nuestro"] /= 1e6
    print(q.to_string(index=False))
    tot_m = macro["vr"].sum()
    tot_o = nuestro["vr"].sum()
    print(f"\nTOTAL valor de mercado: macro {tot_m/1e12:,.3f} B | nuestro {tot_o/1e12:,.3f} B "
          f"| dif {100*(tot_o-tot_m)/tot_m:+.3f}%")
    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            por_clasif.to_excel(xw, sheet_name="Por clasificacion", index=False)
            por_afp.to_excel(xw, sheet_name="Por AFP", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
