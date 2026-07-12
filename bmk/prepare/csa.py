"""
Cuentas CSA (cuentas de liquidez internacional asociadas al MtM de derivados OTC).

Las trae la hoja "Cuentas CSA" del archivo SFC (paso Importar_CSA_BMK de las
macros). Vienen en USD (SALDO CSA ACTIVO); se convierten a COP con la TRM del
corte (derivada del propio archivo SFC, ver bmk.prepare.fx).

Cada CSA con saldo activo > 0 se vuelve una fila de CAJA internacional en el
Benchmark, con el mismo esquema de columnas que los activos clasificados.
"""
from __future__ import annotations

import unicodedata

import pandas as pd

from bmk.prepare.normalize import normalizar_afp

# Tipo de portafolio SFC (hoja CSA) -> codigo interno.
PORT_CSA = {
    "FPO MODERADO": "PO", "FPO CONSERVADOR": "PC", "FPO MAYOR RIESGO": "PM",
    "FPO RETIRO PROGRAMADO": "PR", "CESANTIAS LP": "CS", "CES LP": "CS",
    "CESANTIAS CP": "CS", "CES CP": "CS",
}


def _cod_port_csa(nombre: str) -> str:
    n = "".join(c for c in unicodedata.normalize("NFKD", str(nombre)) if not unicodedata.combining(c)).upper().strip()
    if n in PORT_CSA:
        return PORT_CSA[n]
    if "MAYOR RIESGO" in n:
        return "PM"
    if "CONSERVADOR" in n:
        return "PC"
    if "RETIRO" in n:
        return "PR"
    if "CESANT" in n or n.startswith("CES"):
        return "CS"
    if "MODERADO" in n:
        return "PO"
    return "ND"


def extraer_csa(cuentas_csa: pd.DataFrame, trm: float, solo_industria: bool = True) -> pd.DataFrame:
    """Devuelve filas CSA (CAJA internacional) en el esquema de activos clasificados."""
    if cuentas_csa is None or cuentas_csa.empty:
        return pd.DataFrame()
    # Encabezado real en la fila con 'FECHA' (idx ~2); datos debajo.
    raw = cuentas_csa.reset_index(drop=True)
    hdr_row = None
    for i in range(min(6, len(raw))):
        fila = [str(x).strip().upper() for x in raw.iloc[i].tolist()]
        if "FECHA" in fila:
            hdr_row = i
            break
    if hdr_row is None:
        return pd.DataFrame()
    d = raw.iloc[hdr_row + 1:].copy()
    d.columns = [str(x).strip() for x in raw.iloc[hdr_row].tolist()]
    col_admin = next(c for c in d.columns if "administrador" in c.lower())
    col_port = next(c for c in d.columns if "portafolio" in c.lower())
    col_mon = next(c for c in d.columns if "moneda" in c.lower())
    col_act = next(c for c in d.columns if "activo" in c.lower())

    d["afp"] = d[col_admin].map(normalizar_afp)
    d["cod_portafolio"] = d[col_port].map(_cod_port_csa)
    d["saldo"] = pd.to_numeric(d[col_act], errors="coerce").fillna(0)
    d["moneda"] = d[col_mon].astype("string").str.strip().str.upper().fillna("USD")
    d = d[d["saldo"] > 0]
    if solo_industria:
        d = d[d["afp"] != "COLFONDOS"]
    if d.empty:
        return pd.DataFrame()

    fx = d["moneda"].map(lambda m: trm if str(m).upper() == "USD" else trm)
    vr_cop = d["saldo"] * fx
    out = pd.DataFrame({
        "cod_portafolio": d["cod_portafolio"].values,
        "afp": d["afp"].values,
        "nemo": "", "isin": "", "clas_sfc": "CSA",
        "emisor": "CSA " + d["moneda"].astype(str).values,
        "f_compra": "", "f_vcto": "", "moneda": d["moneda"].values,
        "valor_nominal": d["saldo"].values,
        "tasa_facial_ind": "", "tasa_facial_valor": "",
        "vr_mercado": vr_cop.values,
        "clase_inversion": "MONEY MARKET", "ubicacion": "INTERNACIONAL",
        "clasificacion": "CAJA", "riesgo": "No Reporta",
        "is_clasificado": True, "fuente_clasif": "csa",
    })
    return out.reset_index(drop=True)
