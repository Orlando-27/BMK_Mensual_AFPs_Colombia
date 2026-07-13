"""
Deduce EMPIRICAMENTE contra que fecha de referencia calcula PORVENIR el P&G que
reporta en el SFC para sus swaps SOFUSD de reporte parcial.

Valora los swaps con el motor v6 en varias fechas y, por contrato, compara el
valor reportado por el SFC (col 53/54 no-cero) contra varios CANDIDATOS de P&G:
  - neto(corte)                         -> P&G "desde el inicio" si se pacto a la par
  - neto(corte) - neto(ref)  por cada fecha de referencia dada (p.ej. 30-abr = mes
                                           anterior, 30-may = dia calendario previo)
Reporta, por candidato: mediana del |error %|, correlacion y ratio medio contra el
SFC, para ver cual reproduce mejor el reporte.

Uso (pasar una o varias referencias como 'carpeta:fecha'):
    python3 tools/pyg_diagnostico.py \
        --corte insumos_diarios/2026-05-31:2026-05-31 \
        --ref insumos_diarios/2026-04-30:2026-04-30 \
        --ref insumos_diarios/2026-05-30:2026-05-30 \
        --sfc insumos/sfc/2026_05portainvdeta.xls \
        --out salidas/pyg_diagnostico.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import numpy as np
import pandas as pd

from bmk.ingest import sfc as sfc_mod
from bmk.prepare.normalize import normalizar_afp
import correr_swaps  # noqa: E402


def _neto(spec: str, sfc_path: str) -> pd.DataFrame:
    curvas, fecha = spec.split(":")
    res = correr_swaps.correr(curvas, fecha, sfc_path, f"salidas/swaps/{fecha}")
    der = pd.to_numeric(res["VPN_Derecho_Calc"], errors="coerce")
    obl = pd.to_numeric(res["VPN_Oblig_Calc"], errors="coerce")
    return pd.DataFrame({"isin": res["ISIN"].astype(str).str.strip(),
                         f"neto_{fecha}": der - obl}), fecha


def _sfc(sfc_path: str) -> pd.DataFrame:
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
    d["cat"] = np.where(a0 & b0, "ambas0", np.where(a0, "solo_der0",
                        np.where(b0, "solo_obl0", "ambas_vp")))
    d["sfc_val"] = d["sfc_der"].where(d["cat"] == "solo_obl0", d["sfc_obl"])
    return d[d["cat"].isin(["solo_der0", "solo_obl0"])]


def _score(sfc, cand):
    m = pd.notna(sfc) & pd.notna(cand)
    if m.sum() < 3:
        return None
    s, c = sfc[m], cand[m]
    err = (c - s) / s.abs().replace(0, np.nan) * 100
    corr = np.corrcoef(s, c)[0, 1]
    ratio = (c / s.replace(0, np.nan)).median()
    return {"n": int(m.sum()), "err_abs_med_%": err.abs().median(),
            "corr": corr, "ratio_medio": ratio,
            "suma_sfc_MM": s.sum() / 1e6, "suma_cand_MM": c.sum() / 1e6}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--corte", required=True, help="carpeta:fecha del corte")
    ap.add_argument("--ref", action="append", default=[], help="carpeta:fecha de referencia (repetible)")
    ap.add_argument("--sfc", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()

    print("=== Valorando corte ===")
    cor, fcor = _neto(a.corte, a.sfc)
    netos = cor
    fechas_ref = []
    for spec in a.ref:
        print(f"=== Valorando referencia {spec} ===")
        r, fr = _neto(spec, a.sfc)
        netos = netos.merge(r, on="isin", how="outer")
        fechas_ref.append(fr)

    sfc = _sfc(a.sfc)
    t = sfc.merge(netos, on="isin", how="inner")
    print(f"\nSwaps de reporte parcial casados: {len(t)}")

    # Candidatos de P&G.
    cands = {f"neto_corte ({fcor})": t[f"neto_{fcor}"]}
    for fr in fechas_ref:
        cands[f"corte - {fr}"] = t[f"neto_{fcor}"] - t[f"neto_{fr}"]
        cands[f"-(corte - {fr})"] = -(t[f"neto_{fcor}"] - t[f"neto_{fr}"])
    cands[f"-neto_corte"] = -t[f"neto_{fcor}"]
    cands[f"|neto_corte|"] = t[f"neto_{fcor}"].abs()

    print("\n=== Ajuste de cada candidato de P&G contra el valor SFC ===")
    filas = []
    for name, c in cands.items():
        sc = _score(t["sfc_val"], c)
        if sc:
            sc = {"candidato": name, **sc}
            filas.append(sc)
    rep = pd.DataFrame(filas)
    pd.options.display.float_format = lambda x: f"{x:,.2f}"
    print(rep.to_string(index=False))

    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            rep.to_excel(xw, sheet_name="Ajuste candidatos", index=False)
            t.to_excel(xw, sheet_name="Netos por swap", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
