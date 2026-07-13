"""
Comparacion de mi valoracion de swaps contra el PROPIO informe SFC.

El Formato_415 trae, por contrato, el valor presente que reporta la entidad:
  - col 53 'Valor del derecho (COP)'
  - col 54 'Valor de la obligacion (COP)'
Mi motor v6 produce VPN_Derecho_Calc / VPN_Oblig_Calc por contrato. La idea es
que la diferencia tienda a cero.

Patrones de reporte en el SFC (ver analisis): no todas las entidades reportan el
valor presente de las dos patas.
  - "ambas VP": las dos columnas traen valor presente  -> comparacion por-pata.
  - "una pata en 0" (tipico PORVENIR en SOFUSD): reportan el P&G en una sola
    columna, no el VP de cada pata -> se compara mi NETO (der-obl) contra el
    valor no-cero, y se marca aparte (requiere valoracion del dia anterior para
    reconstruir el P&G con exactitud; aqui solo se reporta como referencia).
  - "ambas en 0" (parte de PROTECCION): no reportan valor -> se excluyen.

Uso:
    python3 tools/comparar_sfc.py \
        --sfc insumos/sfc/2026_05portainvdeta.xls \
        --mio salidas/swaps/2026-05-31/valoracion_full.csv \
        --top 25 --out salidas/comparacion_sfc_swaps.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd

from bmk.ingest import sfc as sfc_mod
from bmk.prepare.normalize import normalizar_afp


def leer_sfc(path: str) -> pd.DataFrame:
    arch = sfc_mod.leer(path)
    f = arch.formato_415
    tipo = pd.to_numeric(f.iloc[:, 9], errors="coerce")
    sw = f[tipo.isin([16, 17])].copy()
    d = pd.DataFrame({
        "isin": sw.iloc[:, 11].astype(str).str.strip(),
        "afp": sw.iloc[:, 0].map(normalizar_afp),
        "tipoD": pd.to_numeric(sw.iloc[:, 9], errors="coerce"),
        "mon_der": sw.iloc[:, 26].astype(str).str.strip(),
        "nom_der": pd.to_numeric(sw.iloc[:, 27], errors="coerce"),
        "mon_obl": sw.iloc[:, 28].astype(str).str.strip(),
        "nom_obl": pd.to_numeric(sw.iloc[:, 29], errors="coerce"),
        "sfc_der": pd.to_numeric(sw.iloc[:, 53], errors="coerce"),
        "sfc_obl": pd.to_numeric(sw.iloc[:, 54], errors="coerce"),
    })
    d["subtipo"] = d.apply(
        lambda r: "IRS" if r["tipoD"] == 16 and r["mon_der"] == r["mon_obl"]
        else ("CCS" if r["tipoD"] == 17 else "IRS-XCCY"), axis=1)
    # categoria de reporte
    a0 = d["sfc_der"].abs() < 1
    b0 = d["sfc_obl"].abs() < 1
    d["cat"] = "ambas_vp"
    d.loc[a0 & b0, "cat"] = "ambas0"
    d.loc[a0 & ~b0, "cat"] = "solo_der0"
    d.loc[~a0 & b0, "cat"] = "solo_obl0"
    return d


def leer_mio(path: str) -> pd.DataFrame:
    o = pd.read_csv(path)
    return pd.DataFrame({
        "isin": o["ISIN"].astype(str).str.strip(),
        "tipo_swap": o["Tipo_Swap"].astype(str) if "Tipo_Swap" in o.columns else "",
        "mio_der": pd.to_numeric(o["VPN_Derecho_Calc"], errors="coerce"),
        "mio_obl": pd.to_numeric(o["VPN_Oblig_Calc"], errors="coerce"),
    })


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--sfc", required=True)
    ap.add_argument("--mio", required=True)
    ap.add_argument("--top", type=int, default=25)
    ap.add_argument("--out")
    a = ap.parse_args()

    s = leer_sfc(a.sfc)
    m = leer_mio(a.mio)
    t = s.merge(m, on="isin", how="inner")
    print(f"SFC swaps: {len(s)} | mios: {len(m)} | casan: {len(t)}")
    print("\n=== Categoria de reporte SFC (casados) ===")
    print(t.groupby(["afp", "cat"]).size().unstack(fill_value=0).to_string())

    # --- SET LIMPIO: ambas patas con valor presente -> comparacion por-pata ---
    cl = t[t["cat"] == "ambas_vp"].copy()
    cl["dif_der"] = cl["mio_der"] - cl["sfc_der"]
    cl["dif_obl"] = cl["mio_obl"] - cl["sfc_obl"]
    cl["dif_der_%"] = cl["dif_der"] / cl["sfc_der"].abs().replace(0, float("nan")) * 100
    cl["dif_obl_%"] = cl["dif_obl"] / cl["sfc_obl"].abs().replace(0, float("nan")) * 100
    cl["sfc_neto"] = cl["sfc_der"] - cl["sfc_obl"]
    cl["mio_neto"] = cl["mio_der"] - cl["mio_obl"]
    cl["dif_neto"] = cl["mio_neto"] - cl["sfc_neto"]

    pd.options.display.float_format = lambda x: f"{x:,.2f}"
    print(f"\n=== SET LIMPIO (ambas VP): {len(cl)} swaps — dif por pata ===")
    res = cl.groupby("subtipo").agg(
        n=("isin", "size"),
        der_med_pct=("dif_der_%", "median"),
        obl_med_pct=("dif_obl_%", "median"),
        der_abs_med=("dif_der_%", lambda x: x.abs().median()),
        obl_abs_med=("dif_obl_%", lambda x: x.abs().median()),
        sfc_neto_MM=("sfc_neto", lambda x: x.sum() / 1e6),
        mio_neto_MM=("mio_neto", lambda x: x.sum() / 1e6),
    ).reset_index()
    print(res.to_string(index=False))
    print(f"\n  |dif pata derecho|  mediana global: {cl['dif_der_%'].abs().median():.2f}%")
    print(f"  |dif pata oblig.|   mediana global: {cl['dif_obl_%'].abs().median():.2f}%")

    print(f"\n=== Top {a.top} del set limpio por |dif neto| (MM COP) ===")
    top = cl.reindex(cl["dif_neto"].abs().sort_values(ascending=False).index).head(a.top)
    v = top[["isin", "afp", "subtipo", "sfc_der", "mio_der", "dif_der_%",
             "sfc_obl", "mio_obl", "dif_obl_%", "dif_neto"]].copy()
    for c in ["sfc_der", "mio_der", "sfc_obl", "mio_obl", "dif_neto"]:
        v[c] = v[c] / 1e6
    print(v.to_string(index=False))

    # --- P&G / parciales: solo referencia ---
    pg = t[t["cat"].isin(["solo_der0", "solo_obl0"])].copy()
    if len(pg):
        pg["sfc_reporta"] = pg["sfc_der"].where(pg["cat"] == "solo_obl0", pg["sfc_obl"])
        pg["mio_neto"] = pg["mio_der"] - pg["mio_obl"]
        print(f"\n=== P&G/parcial: {len(pg)} swaps (una pata en 0) — referencia ===")
        print("  (requieren valoracion del dia anterior para reconstruir el P&G)")
        print(pg.groupby("afp").agg(n=("isin", "size"),
              sfc_report_MM=("sfc_reporta", lambda x: x.sum() / 1e6),
              mio_neto_MM=("mio_neto", lambda x: x.sum() / 1e6)).to_string())

    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            res.to_excel(xw, sheet_name="Resumen limpio", index=False)
            cl.sort_values("dif_neto", key=lambda s: s.abs(), ascending=False) \
                .to_excel(xw, sheet_name="Set limpio", index=False)
            if len(pg):
                pg.to_excel(xw, sheet_name="PyG parciales", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
