"""
Valida los precios/VPN de la Calculadora de swaps (motor v6) contra los valores
que reporta el propio archivo de la SFC (Formato_415), swap por swap, al MISMO
corte.

El Formato_415 reporta, por contrato de swap (tipo 16/17):
  - col 53 = Valor presente pata derecho (COP)
  - col 54 = Valor presente pata obligacion (COP)
Estos son los "precios reportados por la Super" al corte. El motor v6 produce su
propio VPN por pata; este script los enfrenta uno a uno (por numero de contrato
= col 11 = ISIN interno).

Camara de riesgo (CRCC): los IRS SOFR se liquidan a diario, por lo que la SFC no
reporta su MtM completo sino el valor liquidado del dia (pata obligacion = 0,
pata derecho = variacion). Se separan del cuadre "full" y se marcan como camara:
su validacion correcta es la variacion diaria = MtM(corte) - MtM(dia anterior),
que requiere valorar tambien el dia habil anterior.

Uso:
    python3 tools/comparar_swaps_sfc.py \
        --sfc insumos/sfc/2026_06portainvdeta.xls \
        --valoracion salidas/2026-06-30/valoracion_swaps_industria_2026-06-30.xlsx \
        --out salidas/2026-06-30/comparacion_swaps_vs_sfc_2026-06-30.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import numpy as np
import pandas as pd

from bmk.ingest import sfc as _sfc
from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp

CAMARA = "IRS SOFUSD"  # swaps de camara (CRCC): reportados liquidados a diario


def leer_sfc(path: str) -> pd.DataFrame:
    """VPN por pata reportado por la SFC (Formato_415 col 53/54), industria."""
    arch = _sfc.leer(path)
    f = arch.formato_415
    tipo = pd.to_numeric(f.iloc[:, COL_415["tipo_derivado"]], errors="coerce")
    afp = f.iloc[:, 0].map(normalizar_afp)
    df = pd.DataFrame({
        "ISIN": f.iloc[:, 11].astype(str),
        "afp": afp,
        "sfc_vp_der": pd.to_numeric(f.iloc[:, 53], errors="coerce"),
        "sfc_vp_obl": pd.to_numeric(f.iloc[:, 54], errors="coerce"),
    })
    df = df[tipo.isin([16, 17]) & (afp != "COLFONDOS")].reset_index(drop=True)
    df["sfc_mtm"] = df["sfc_vp_der"] - df["sfc_vp_obl"]
    return df


def leer_valoracion(path: str) -> pd.DataFrame:
    d = pd.read_excel(path, sheet_name="Valoracion Swaps")
    d = d.rename(columns={"VPN DER": "my_vp_der", "VPN OBL": "my_vp_obl",
                          "MtM neto (DER-OBL)": "my_mtm", "Tipo Swap": "tipo"})
    d["ISIN"] = d["ISIN"].astype(str)
    return d[["ISIN", "tipo", "afp" if "afp" in d.columns else "AFP",
              "my_vp_der", "my_vp_obl", "my_mtm"]].rename(columns={"AFP": "afp"})


def _rel(a, b):
    return (a - b) / b.replace(0, np.nan)


def comparar(sfc_df: pd.DataFrame, val_df: pd.DataFrame):
    j = sfc_df.merge(val_df[["ISIN", "tipo", "my_vp_der", "my_vp_obl", "my_mtm"]],
                     on="ISIN", how="inner")
    j["rel_der"] = _rel(j["my_vp_der"], j["sfc_vp_der"])
    j["rel_obl"] = _rel(j["my_vp_obl"], j["sfc_vp_obl"])
    j["dif_mtm"] = j["my_mtm"] - j["sfc_mtm"]
    j["es_camara"] = j["tipo"] == CAMARA

    def _resumen(sub: pd.DataFrame) -> dict:
        leg_tot = sub["sfc_vp_der"].abs().sum() + sub["sfc_vp_obl"].abs().sum()
        leg_err = ((sub["my_vp_der"] - sub["sfc_vp_der"]).abs().sum()
                   + (sub["my_vp_obl"] - sub["sfc_vp_obl"]).abs().sum())
        return {
            "n": len(sub),
            "SFC_der_MM": sub["sfc_vp_der"].sum() / 1e6,
            "mio_der_MM": sub["my_vp_der"].sum() / 1e6,
            "SFC_obl_MM": sub["sfc_vp_obl"].sum() / 1e6,
            "mio_obl_MM": sub["my_vp_obl"].sum() / 1e6,
            "err_patas_%": 100 * leg_err / leg_tot if leg_tot else np.nan,
            "SFC_MtM_MM": sub["sfc_mtm"].sum() / 1e6,
            "mio_MtM_MM": sub["my_mtm"].sum() / 1e6,
            "corr_MtM": sub["my_mtm"].corr(sub["sfc_mtm"]),
        }

    filas = []
    for tp, sub in j.groupby("tipo"):
        filas.append({"grupo": tp, **_resumen(sub)})
    nc = j[~j["es_camara"]]
    filas.append({"grupo": "TOTAL no-camara", **_resumen(nc)})
    filas.append({"grupo": "TOTAL camara (IRS SOFR)", **_resumen(j[j["es_camara"]])})
    resumen = pd.DataFrame(filas)
    return j, resumen


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--sfc", required=True)
    ap.add_argument("--valoracion", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()
    sfc_df = leer_sfc(a.sfc)
    val_df = leer_valoracion(a.valoracion)
    detalle, resumen = comparar(sfc_df, val_df)
    pd.options.display.float_format = lambda x: f"{x:,.2f}"
    pd.options.display.width = 200
    print("=== Calculadora (motor v6) vs SFC reportado (Formato_415 col 53/54) ===\n")
    print(resumen.to_string(index=False))
    print("\nNota: los IRS SOFR son de camara (CRCC, liquidacion diaria); la SFC "
          "reporta su valor liquidado del dia (obl=0), no el MtM completo. Su "
          "validacion es la variacion diaria = MtM(corte) - MtM(dia anterior).")
    if a.out:
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            resumen.to_excel(xw, sheet_name="Resumen", index=False)
            detalle.sort_values("dif_mtm", key=lambda s: s.abs(), ascending=False
                                ).to_excel(xw, sheet_name="Detalle por swap", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
