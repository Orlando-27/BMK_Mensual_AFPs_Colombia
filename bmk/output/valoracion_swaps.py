"""
Output de valoracion de swaps uno a uno (replica el detalle de la Calculadora
Swap): por contrato y pata, VPN, precio, duracion modificada, convexidad, curva
de descuento, y el MtM neto (derecho - obligacion).

Une la salida del motor v6 (valoracion_full: VPN/precio/duracion/convexidad por
pata) con la sabana de condiciones faciales (AFP, portafolio, nominales, monedas,
tasas faciales) por numero de contrato.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd


COLS = [
    ("ISIN", "ISIN"), ("AFP", "afp"), ("Portafolio", "portafolio"),
    ("Tipo Swap", "Tipo_Swap"), ("F.Compra", "Fecha_Compra"), ("F.Vcto", "Fecha_Vcto"),
    # Pata derecho
    ("Moneda DER", "der_moneda"), ("Nominal DER", "der_nominal"),
    ("Indice DER", "der_indice_cod"), ("Tasa Facial DER", "der_tasa_facial"),
    ("VPN DER", "VPN_Derecho_Calc"), ("Precio DER", "Precio_Der"),
    ("Dur.Mod DER", "Duracion_Mod_Der"), ("Convex DER", "Convexidad_Der"),
    ("Curva Desc DER", "Curva_Disc_Der"),
    # Pata obligacion
    ("Moneda OBL", "obl_moneda"), ("Nominal OBL", "obl_nominal"),
    ("Indice OBL", "obl_indice_cod"), ("Tasa Facial OBL", "obl_tasa_facial"),
    ("VPN OBL", "VPN_Oblig_Calc"), ("Precio OBL", "Precio_Obl"),
    ("Dur.Mod OBL", "Duracion_Mod_Obl"), ("Convex OBL", "Convexidad_Obl"),
    ("Curva Desc OBL", "Curva_Disc_Obl"),
    # Resultado
    ("MtM neto (DER-OBL)", "mtm"),
]


def construir(valoracion_full: pd.DataFrame, sabana: pd.DataFrame,
              camara: pd.DataFrame | None = None) -> pd.DataFrame:
    """Une la valoracion v6 con la sabana por numero de contrato (ISIN).

    valoracion_full trae el VPN COMPLETO por pata (igual que la Calculadora Swap).
    Si se pasa `camara` (la misma valoracion con la variacion diaria ya aplicada a
    los IRS SOFR), se agregan columnas con esa variacion, que es el valor que se
    inyecta al Benchmark para los swaps de camara. Asi el detalle muestra AMBOS:
    el VPN completo (comparable con el Excel) y la variacion de camara."""
    v = valoracion_full.copy()
    v["ISIN"] = v["ISIN"].astype(str).str.strip()
    s = sabana.copy()
    s["id_contrato"] = s["id_contrato"].astype(str).str.strip()
    meta = s[["id_contrato", "afp", "portafolio", "der_moneda", "der_nominal",
              "der_indice_cod", "der_tasa_facial", "obl_moneda", "obl_nominal",
              "obl_indice_cod", "obl_tasa_facial"]].drop_duplicates("id_contrato")
    d = v.merge(meta, left_on="ISIN", right_on="id_contrato", how="left")
    d["mtm"] = pd.to_numeric(d["VPN_Derecho_Calc"], errors="coerce") \
        - pd.to_numeric(d["VPN_Oblig_Calc"], errors="coerce")
    out = pd.DataFrame({etq: d[col] if col in d.columns else "" for etq, col in COLS})
    if camara is not None:
        c = camara.copy()
        c["ISIN"] = c["ISIN"].astype(str).str.strip()
        cam = c[["ISIN", "VPN_Derecho_Calc", "VPN_Oblig_Calc"]].rename(columns={
            "VPN_Derecho_Calc": "VPN DER camara", "VPN_Oblig_Calc": "VPN OBL camara"})
        out = out.merge(cam, on="ISIN", how="left")
        der = pd.to_numeric(out["VPN DER camara"], errors="coerce")
        obl = pd.to_numeric(out["VPN OBL camara"], errors="coerce")
        out["MtM camara (usado en Benchmark)"] = der - obl
        # Marca solo donde la camara cambia el valor (IRS SOFR); en el resto el
        # valor de Benchmark = VPN completo.
        cambia = (out["MtM camara (usado en Benchmark)"].round(0)
                  != pd.to_numeric(out["MtM neto (DER-OBL)"], errors="coerce").round(0))
        out["es_camara"] = cambia.map({True: "SI (variacion diaria)", False: ""})
    return out


def escribir(valoracion_full: pd.DataFrame, sabana: pd.DataFrame,
             out_dir: str | Path, corte: str, camara: pd.DataFrame | None = None) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    df = construir(valoracion_full, sabana, camara=camara)
    dest = out_dir / f"valoracion_swaps_industria_{corte}.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        df.to_excel(xw, sheet_name="Valoracion Swaps", index=False)
        # Resumen por tipo de swap.
        res = df.copy()
        res["MtM neto (DER-OBL)"] = pd.to_numeric(res["MtM neto (DER-OBL)"], errors="coerce")
        res["VPN DER"] = pd.to_numeric(res["VPN DER"], errors="coerce")
        res["VPN OBL"] = pd.to_numeric(res["VPN OBL"], errors="coerce")
        g = res.groupby("Tipo Swap").agg(
            **{"# swaps": ("ISIN", "size"),
               "VPN DER (MM)": ("VPN DER", lambda s: s.sum() / 1e6),
               "VPN OBL (MM)": ("VPN OBL", lambda s: s.sum() / 1e6),
               "MtM neto (MM)": ("MtM neto (DER-OBL)", lambda s: s.sum() / 1e6)}).reset_index()
        g.to_excel(xw, sheet_name="Resumen por tipo", index=False)
    return str(dest)
