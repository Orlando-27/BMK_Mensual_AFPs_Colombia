#!/usr/bin/env python3
"""
Compara la hoja Benchmark generada por el pipeline contra la hoja Benchmark de
un archivo Benchmark (Detallado) producido por las macros (ground truth).

Alinea nomenclatura de AFP (OLD MUTUAL/SKANDIA, tildes), separa derivados de
activos, y reporta:
  - conteo de activos por AFP (theirs vs ours)
  - concordancia de clasificacion por identificador comun (CLASE/UBIC/CLASIF/RIESGO)
  - activos solo en uno u otro

Uso:
    python3 tools/comparar_benchmark.py --macro insumos/Benchmark_mayo.xlsx \
        --nuestro salidas/2026-05-31/hoja_Benchmark_2026-05-31.xlsx \
        --out salidas/2026-05-31/comparacion_mayo.xlsx
"""
from __future__ import annotations

import argparse
import unicodedata
import warnings
from pathlib import Path

import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

CAMPOS = [("CLASE DE INVERSION", "clase"), ("UBICACION", "ubic"),
          ("CLASIFICACION", "clasif"), ("RIESGO", "riesgo")]


def na(s) -> str:
    s = "" if s is None else str(s).strip()
    return "" if s.upper() in ("NA", "NONE", "") else s


def afp_key(s) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).upper()
    if any(k in s for k in ("OLD MUTUAL", "SKANDIA", "ACCAI")):
        return "SKANDIA"
    for k in ("PORVENIR", "PROTECCION", "COLFONDOS"):
        if k in s:
            return k
    return s


def _ident(isin, nemo) -> str:
    return (na(isin) or na(nemo)).upper()


def _es_deriv(clase, clas_sfc, clasif) -> bool:
    c, cs, cl = clase.upper(), clas_sfc.upper(), clasif.upper()
    return (c in ("FORWARD", "OPCIONES") or cs.startswith(("FW", "OPT"))
            or "SWAP" in cl or "FORWARD" in cl or c.startswith("SWAP"))


def leer_macro(path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Benchmark"]
    rows = [r for r in ws.iter_rows(min_row=15, values_only=True)
            if not (r[0] is None and r[1] is None)]
    wb.close()
    ix = {"port": 0, "afp": 1, "nemo": 3, "isin": 4, "clas_sfc": 5,
          "clase": 13, "ubic": 16, "clasif": 17, "riesgo": 18}
    d = pd.DataFrame({k: [na(r[i]) for r in rows] for k, i in ix.items()})
    d["afp"] = d["afp"].map(afp_key)
    d["id"] = [_ident(i, n) for i, n in zip(d["isin"], d["nemo"])]
    d["deriv"] = [_es_deriv(a, b, c) for a, b, c in zip(d["clase"], d["clas_sfc"], d["clasif"])]
    return d


def leer_nuestro(path: str) -> pd.DataFrame:
    o = pd.read_excel(path, sheet_name="Benchmark")
    d = pd.DataFrame({
        "port": o["Portafolio"].map(na), "afp": o["AFP"].map(afp_key),
        "nemo": o["Nemo"].map(na), "isin": o["ISIN"].map(na),
        "clas_sfc": o["Clas SFC"].map(na), "clase": o["CLASE DE INVERSION"].map(na),
        "ubic": o["UBICACION"].map(na), "clasif": o["CLASIFICACIÓN"].map(na),
        "riesgo": o["RIESGO"].map(na),
    })
    d["id"] = [_ident(i, n) for i, n in zip(d["isin"], d["nemo"])]
    return d


def comparar(macro: pd.DataFrame, nuestro: pd.DataFrame) -> dict:
    T = macro[~macro["deriv"]].copy()  # activos del macro (excluye derivados)
    O = nuestro
    conteo = pd.DataFrame({"macro": T.groupby("afp").size(),
                           "nuestro": O.groupby("afp").size()}).fillna(0).astype(int)
    conteo["dif"] = conteo["nuestro"] - conteo["macro"]

    tt = T[T["id"] != ""].drop_duplicates("id").set_index("id")
    oo = O[O["id"] != ""].drop_duplicates("id").set_index("id")
    comunes = tt.index.intersection(oo.index)

    concord = {}
    difs = {}
    for label, col in CAMPOS:
        a = tt.loc[comunes, col].str.upper().str.strip()
        b = oo.loc[comunes, col].str.upper().str.strip()
        concord[label] = (int((a == b).sum()), len(comunes))
        d = comunes[(a != b).values]
        if len(d):
            difs[label] = pd.DataFrame({"id": d, "macro": a.loc[d].values,
                                        "nuestro": b.loc[d].values,
                                        "clas_sfc": tt.loc[d, "clas_sfc"].values})
    return {"conteo": conteo, "concord": concord, "difs": difs,
            "solo_macro": len(tt.index.difference(oo.index)),
            "solo_nuestro": len(oo.index.difference(tt.index)),
            "comunes": len(comunes)}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--macro", required=True)
    ap.add_argument("--nuestro", required=True)
    ap.add_argument("--out", default="comparacion.xlsx")
    args = ap.parse_args()

    print(f"Leyendo macro: {args.macro} ...")
    macro = leer_macro(args.macro)
    print(f"  {len(macro)} filas ({int(macro['deriv'].sum())} derivados)")
    nuestro = leer_nuestro(args.nuestro)
    print(f"Nuestro: {len(nuestro)} activos")

    r = comparar(macro, nuestro)
    print("\n=== Conteo activos por AFP (macro vs nuestro) ===")
    print(r["conteo"].to_string())
    print(f"\n=== Concordancia de clasificacion ({r['comunes']} identificadores comunes) ===")
    for label, (ok, tot) in r["concord"].items():
        print(f"  {label:22}: {ok}/{tot} = {100*ok/tot:.1f}%" if tot else f"  {label}: sin comunes")
    print(f"\nSolo en macro: {r['solo_macro']}  |  Solo en nuestro: {r['solo_nuestro']}")

    with pd.ExcelWriter(args.out, engine="xlsxwriter") as xw:
        r["conteo"].reset_index().to_excel(xw, sheet_name="Conteo por AFP", index=False)
        resumen = pd.DataFrame([{"campo": k, "concordancia": f"{v[0]}/{v[1]}",
                                 "pct": round(100*v[0]/v[1], 2) if v[1] else None}
                                for k, v in r["concord"].items()])
        resumen.to_excel(xw, sheet_name="Concordancia", index=False)
        for label, d in r["difs"].items():
            d.to_excel(xw, sheet_name=f"Dif {label[:20]}", index=False)
    print(f"\nReporte: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
