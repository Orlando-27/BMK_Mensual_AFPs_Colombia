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


def _leer_macro_patas(path: str) -> pd.DataFrame:
    """Filas crudas (una por pata) de los swaps del Benchmark macro."""
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
        filas.append((
            ident,
            _norm(r[5]),                       # clas: IRS / CCS
            _norm(r[3]),                       # nemo: SWAP TF / SWAP TV
            _norm(r[9]),                       # moneda de la pata
            pd.to_numeric(r[10], errors="coerce"),   # nominal
            pd.to_numeric(r[11], errors="coerce"),   # Vr Mercado INICIAL (con signo)
            pd.to_numeric(r[20], errors="coerce"),   # Precio inicial
        ))
    wb.close()
    d = pd.DataFrame(filas, columns=["isin", "clas", "nemo", "moneda", "nominal", "vr", "precio"])
    return d[d["isin"] != ""].reset_index(drop=True)


def _subtipo(monedas: set[str]) -> str:
    """IBRCOP si ambas patas COP; SOFUSD si hay USD; si no, mezcla de la clas."""
    m = {x for x in monedas if x}
    if m == {"COP"}:
        return "IBRCOP"
    if "USD" in m and "COP" not in m:
        return "SOFUSD"
    if m == {"COP", "USD"}:
        return "COPUSD"
    return "/".join(sorted(m)) or "?"


def leer_macro(path: str) -> pd.DataFrame:
    """Neto y bruto por ISIN de las filas de swap del Benchmark macro."""
    d = _leer_macro_patas(path)
    g = d.groupby("isin").agg(
        clas=("clas", "first"),
        n_patas=("vr", "size"),
        macro_neto=("vr", "sum"),
        macro_bruto=("vr", lambda s: s.abs().sum()),
        macro_tf=("vr", lambda s: s[d.loc[s.index, "nemo"] == "SWAP TF"].sum()),
        macro_tv=("vr", lambda s: s[d.loc[s.index, "nemo"] == "SWAP TV"].sum()),
        precio_min=("precio", "min"),
        precio_max=("precio", "max"),
        subtipo=("moneda", lambda s: _subtipo(set(s))),
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


def drilldown(macro_path: str, mio_path: str, isin: str) -> int:
    """Detalle pata a pata de un contrato: macro (2 filas) vs mi der/obl."""
    d = _leer_macro_patas(macro_path)
    md = d[d["isin"].astype(str) == str(isin)]
    print(f"=== MACRO (Benchmark) contrato {isin} ===")
    if md.empty:
        print("  (no encontrado en el macro)")
    else:
        pd.options.display.float_format = lambda x: f"{x:,.2f}"
        print(md[["nemo", "clas", "moneda", "nominal", "vr", "precio"]].to_string(index=False))
        print(f"  neto (suma patas firmadas): {md['vr'].sum():,.0f} COP")

    o = pd.read_csv(mio_path)
    mo = o[o["ISIN"].astype(str) == str(isin)]
    print(f"\n=== MIO (motor v6) contrato {isin} ===")
    if mo.empty:
        print("  (no encontrado en mi valoracion)")
    else:
        cols = ["Tipo_Swap", "VPN_Derecho_Calc", "Precio_Der", "Duracion_Mod_Der", "Curva_Disc_Der",
                "VPN_Oblig_Calc", "Precio_Obl", "Duracion_Mod_Obl", "Curva_Disc_Obl"]
        cols = [c for c in cols if c in mo.columns]
        for c in cols:
            print(f"  {c:20} = {mo.iloc[0][c]}")
        der = pd.to_numeric(mo.iloc[0].get("VPN_Derecho_Calc"), errors="coerce")
        obl = pd.to_numeric(mo.iloc[0].get("VPN_Oblig_Calc"), errors="coerce")
        print(f"  neto (der - obl)     = {der - obl:,.0f} COP")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--macro", required=True)
    ap.add_argument("--mio", required=True)
    ap.add_argument("--top", type=int, default=25)
    ap.add_argument("--isin", help="Drill-down: imprime las 2 patas del macro y mi der/obl para ese contrato")
    ap.add_argument("--out")
    a = ap.parse_args()

    if a.isin:
        return drilldown(a.macro, a.mio, a.isin)

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

    # Resumen por subtipo (IBRCOP / SOFUSD / COPUSD / ...).
    pd.options.display.float_format = lambda x: f"{x:,.1f}"
    res = b.groupby("subtipo").agg(
        n=("isin", "size"),
        macro_neto_MM=("macro_neto", lambda s: s.sum() / 1e6),
        mio_neto_MM=("mio_neto", lambda s: s.sum() / 1e6),
        macro_bruto_B=("macro_bruto", lambda s: s.sum() / 1e12),
        mio_bruto_B=("mio_bruto", lambda s: s.sum() / 1e12),
    ).reset_index()
    res["dif_neto_MM"] = res["mio_neto_MM"] - res["macro_neto_MM"]
    res["dif_bruto_%"] = (res["mio_bruto_B"] - res["macro_bruto_B"]) / res["macro_bruto_B"].abs() * 100
    print("=== Resumen por subtipo (neto en millones, bruto en billones) ===")
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
