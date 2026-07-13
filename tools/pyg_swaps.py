"""
Reconstruye el P&G (utilidad/perdida) diario de los swaps valorando en DOS
fechas y comparando contra lo que reporta el SFC.

Algunas entidades (tipico PORVENIR en SOFUSD) no reportan el valor presente de
cada pata en el Formato_415, sino el P&G en una sola columna (col 53 o 54). El
manual indica valorar en tres fechas; para el P&G del corte se necesita el dia
habil anterior:

    P&G(corte) = neto(corte) - neto(dia_anterior)
    neto = VPN_Derecho - VPN_Oblig  (valor presente del swap)

Este script valora los swaps con el motor v6 en ambas fechas (curvas INFOVALMER
de cada dia), calcula el P&G por contrato y lo compara contra el valor no-cero
que reporta el SFC para los swaps de reporte parcial.

Uso:
    python3 tools/pyg_swaps.py \
        --curvas-ant insumos_diarios/2026-05-30 --fecha-ant 2026-05-30 \
        --curvas-cor insumos_diarios/2026-05-31 --fecha-cor 2026-05-31 \
        --sfc insumos/sfc/2026_05portainvdeta.xls \
        --out salidas/pyg_swaps_2026-05-31.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd

from bmk.ingest import sfc as sfc_mod
from bmk.prepare.normalize import normalizar_afp
import correr_swaps  # noqa: E402  (mismo directorio tools/)


def _neto(curvas_dir: str, fecha: str, sfc_path: str) -> pd.DataFrame:
    res = correr_swaps.correr(curvas_dir, fecha, sfc_path, f"salidas/swaps/{fecha}")
    der = pd.to_numeric(res["VPN_Derecho_Calc"], errors="coerce")
    obl = pd.to_numeric(res["VPN_Oblig_Calc"], errors="coerce")
    return pd.DataFrame({"isin": res["ISIN"].astype(str).str.strip(),
                         "neto": der - obl})


def _sfc_reporte(sfc_path: str) -> pd.DataFrame:
    arch = sfc_mod.leer(sfc_path)
    f = arch.formato_415
    tipo = pd.to_numeric(f.iloc[:, 9], errors="coerce")
    sw = f[tipo.isin([16, 17])].copy()
    d = pd.DataFrame({
        "isin": sw.iloc[:, 11].astype(str).str.strip(),
        "afp": sw.iloc[:, 0].map(normalizar_afp),
        "sfc_der": pd.to_numeric(sw.iloc[:, 53], errors="coerce"),
        "sfc_obl": pd.to_numeric(sw.iloc[:, 54], errors="coerce"),
    })
    a0 = d["sfc_der"].abs() < 1
    b0 = d["sfc_obl"].abs() < 1
    d["cat"] = "ambas_vp"
    d.loc[a0 & b0, "cat"] = "ambas0"
    d.loc[a0 & ~b0, "cat"] = "solo_der0"
    d.loc[~a0 & b0, "cat"] = "solo_obl0"
    # Para reporte parcial, el P&G reportado es la columna no-cero.
    d["sfc_pyg"] = d["sfc_der"].where(d["cat"] == "solo_obl0", d["sfc_obl"])
    d.loc[d["cat"] == "ambas_vp", "sfc_pyg"] = d["sfc_der"] - d["sfc_obl"]
    return d


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--curvas-ant", required=True)
    ap.add_argument("--fecha-ant", required=True)
    ap.add_argument("--curvas-cor", required=True)
    ap.add_argument("--fecha-cor", required=True)
    ap.add_argument("--sfc", required=True)
    ap.add_argument("--top", type=int, default=15)
    ap.add_argument("--out")
    a = ap.parse_args()

    print(f"=== Valorando dia anterior {a.fecha_ant} ===")
    ant = _neto(a.curvas_ant, a.fecha_ant, a.sfc).rename(columns={"neto": "neto_ant"})
    print(f"=== Valorando corte {a.fecha_cor} ===")
    cor = _neto(a.curvas_cor, a.fecha_cor, a.sfc).rename(columns={"neto": "neto_cor"})

    m = ant.merge(cor, on="isin", how="inner")
    m["pyg_calc"] = m["neto_cor"] - m["neto_ant"]

    sfc = _sfc_reporte(a.sfc)
    t = sfc.merge(m, on="isin", how="inner")

    # Foco: los que reportan P&G parcial (una pata en 0).
    pg = t[t["cat"].isin(["solo_der0", "solo_obl0"])].copy()
    pg["dif"] = pg["pyg_calc"] - pg["sfc_pyg"]
    pg["dif_%"] = pg["dif"] / pg["sfc_pyg"].abs().replace(0, float("nan")) * 100

    pd.options.display.float_format = lambda x: f"{x:,.2f}"
    print(f"\n=== P&G reconstruido vs SFC — swaps de reporte parcial ({len(pg)}) ===")
    print(pg.groupby("afp").agg(
        n=("isin", "size"),
        sfc_pyg_MM=("sfc_pyg", lambda x: x.sum() / 1e6),
        calc_pyg_MM=("pyg_calc", lambda x: x.sum() / 1e6),
        dif_abs_med_pct=("dif_%", lambda x: x.abs().median()),
    ).to_string())

    print(f"\n=== Top {a.top} por |dif| (MM COP) ===")
    top = pg.reindex(pg["dif"].abs().sort_values(ascending=False).index).head(a.top)
    v = top[["isin", "afp", "cat", "sfc_pyg", "neto_ant", "neto_cor", "pyg_calc", "dif", "dif_%"]].copy()
    for c in ["sfc_pyg", "neto_ant", "neto_cor", "pyg_calc", "dif"]:
        v[c] = v[c] / 1e6
    print(v.to_string(index=False))

    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            pg.sort_values("dif", key=lambda s: s.abs(), ascending=False) \
                .to_excel(xw, sheet_name="PyG parcial", index=False)
            t.to_excel(xw, sheet_name="Todos", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
