"""
Comparacion POR-SWAP (linea a linea) de la valoracion de swaps contra el
Benchmark de las macros (ground truth).

La hoja Benchmark del macro trae cada swap en DOS filas (una por pata), unidas
por el IDENTIFICADOR (col 2 = numero de contrato = ISIN interno). Cada fila
lleva su Vr Mercado INICIAL (col 11, con signo recibir+/pagar-) y su Precio
inicial (col 20). Este script:

  1. Lee del macro, por ISIN: valor neto (suma de las 2 patas, ya firmadas) y
     magnitud bruta (|pata1|+|pata2|), mas la clasificacion IRS/CCS.
  2. Lee de mi valoracion (valoracion_full.csv del motor v6): VPN_Derecho_Calc,
     VPN_Oblig_Calc por ISIN. Neto = Der - Obl (derecho se recibe, oblig se paga);
     bruto = Der + Obl.
  3. Casa por ISIN y reporta, por IRS/CCS y por swap, la diferencia de NETO y de
     magnitud, ordenando por mayor descuadre para ubicar los swaps que mueven el
     residual.

IMPORTANTE (regla M-1): el macro del mes M usa el corte SFC del mes M-1. Para
comparar valores REALES ambos deben ser del MISMO corte:
  - mi valoracion del corte 2026-05-31  <->  Benchmark junio.xlsx
  - mi valoracion del corte 2026-04-30  <->  Benchmark_mayo.xlsx
Si se cruzan fechas distintas, la mecanica del join es valida pero los valores
no tienen por que coincidir (son otro dia de mercado).

Uso:
    python3 tools/comparar_swaps.py \
        --macro "Benchmark junio.xlsx" \
        --mio salidas/swaps/2026-05-31/valoracion_full.csv \
        --top 25 --out salidas/comparacion_swaps.xlsx
"""
from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


def _norm(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(s))
                   if not unicodedata.combining(c)).upper().strip()


def leer_macro(path: str) -> pd.DataFrame:
    """Neto y bruto por ISIN de las filas de swap del Benchmark macro."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Benchmark"]
    filas = []
    for r in ws.iter_rows(min_row=15, values_only=True):
        if len(r) < 21:
            continue
        clasif = _norm(r[17]) if r[17] is not None else ""
        if clasif != "SWAP":
            continue
        ident = str(r[2]).strip() if r[2] is not None else ""
        clas = _norm(r[5])  # IRS / CCS
        nemo = _norm(r[3])  # SWAP TF / SWAP TV
        vr = pd.to_numeric(r[11], errors="coerce")
        precio = pd.to_numeric(r[20], errors="coerce")
        filas.append((ident, clas, nemo, vr, precio))
    wb.close()
    d = pd.DataFrame(filas, columns=["isin", "clas", "nemo", "vr", "precio"])
    d = d[d["isin"] != ""]
    g = d.groupby("isin").agg(
        clas=("clas", "first"),
        n_patas=("vr", "size"),
        macro_neto=("vr", "sum"),
        macro_bruto=("vr", lambda s: s.abs().sum()),
        precio_min=("precio", "min"),
        precio_max=("precio", "max"),
    ).reset_index()
    return g


def leer_mio(path: str) -> pd.DataFrame:
    o = pd.read_csv(path)
    isin = o["ISIN"].astype(str).str.strip()
    der = pd.to_numeric(o["VPN_Derecho_Calc"], errors="coerce")
    obl = pd.to_numeric(o["VPN_Oblig_Calc"], errors="coerce")
    tipo = o["Tipo_Swap"].astype(str) if "Tipo_Swap" in o.columns else ""
    return pd.DataFrame({
        "isin": isin,
        "tipo_swap": tipo,
        "mio_neto": der - obl,   # derecho se recibe (+), oblig se paga (-)
        "mio_bruto": der.abs() + obl.abs(),
        "mio_der": der,
        "mio_obl": obl,
    })


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--macro", required=True)
    ap.add_argument("--mio", required=True)
    ap.add_argument("--top", type=int, default=25)
    ap.add_argument("--out")
    a = ap.parse_args()

    macro = leer_macro(a.macro)
    mio = leer_mio(a.mio)
    t = macro.merge(mio, on="isin", how="outer", indicator=True)

    solo_macro = (t["_merge"] == "left_only").sum()
    solo_mio = (t["_merge"] == "right_only").sum()
    ambos = (t["_merge"] == "both").sum()
    print(f"Swaps  macro: {len(macro)} | mio: {len(mio)} | "
          f"casan: {ambos} | solo macro: {solo_macro} | solo mio: {solo_mio}\n")

    b = t[t["_merge"] == "both"].copy()
    b["dif_neto"] = b["mio_neto"] - b["macro_neto"]
    b["dif_bruto_%"] = (b["mio_bruto"] - b["macro_bruto"]) / b["macro_bruto"].abs().replace(0, float("nan")) * 100

    # Resumen por clasificacion (IRS/CCS).
    pd.options.display.float_format = lambda x: f"{x:,.1f}"
    res = b.groupby("clas").agg(
        n=("isin", "size"),
        macro_neto_MM=("macro_neto", lambda s: s.sum() / 1e6),
        mio_neto_MM=("mio_neto", lambda s: s.sum() / 1e6),
        macro_bruto_B=("macro_bruto", lambda s: s.sum() / 1e12),
        mio_bruto_B=("mio_bruto", lambda s: s.sum() / 1e12),
    ).reset_index()
    res["dif_neto_MM"] = res["mio_neto_MM"] - res["macro_neto_MM"]
    print("=== Resumen por clasificacion (neto en millones, bruto en billones) ===")
    print(res.to_string(index=False))

    print(f"\n=== Top {a.top} swaps por |dif de neto| (millones COP) ===")
    top = b.reindex(b["dif_neto"].abs().sort_values(ascending=False).index).head(a.top)
    vis = top[["isin", "clas", "tipo_swap", "macro_neto", "mio_neto", "dif_neto",
               "macro_bruto", "mio_bruto", "dif_bruto_%", "precio_min", "precio_max"]].copy()
    for c in ["macro_neto", "mio_neto", "dif_neto", "macro_bruto", "mio_bruto"]:
        vis[c] = vis[c] / 1e6
    print(vis.to_string(index=False))

    tot_m = b["macro_neto"].sum() / 1e6
    tot_o = b["mio_neto"].sum() / 1e6
    print(f"\nNETO total swaps  macro: {tot_m:,.0f} MM | mio: {tot_o:,.0f} MM | "
          f"dif: {tot_o - tot_m:,.0f} MM")

    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            res.to_excel(xw, sheet_name="Resumen", index=False)
            b.sort_values("dif_neto", key=lambda s: s.abs(), ascending=False) \
                .to_excel(xw, sheet_name="Por swap", index=False)
            t[t["_merge"] != "both"].to_excel(xw, sheet_name="Sin casar", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
