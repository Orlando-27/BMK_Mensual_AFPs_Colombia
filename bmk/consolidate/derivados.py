"""
Tabla unificada de derivados para controles.

Consolida en un solo esquema todos los derivados valorados (o pendientes de
valorar) para que el archivo de controles pueda revisar, por tipo:
  - cantidad de instrumentos, valor de mercado, nominal y precio.

Esquema de salida (una fila por instrumento/pata segun el tipo):
  tipo, afp, portafolio, subtipo, moneda, nominal, valor_mercado, precio, valorado
"""
from __future__ import annotations

import pandas as pd

_COLS = ["tipo", "afp", "portafolio", "subtipo", "moneda", "nominal",
         "valor_mercado", "precio", "valorado"]


def _fila(tipo, d, subtipo, moneda, nominal, vm, precio, valorado):
    return {"tipo": tipo, "afp": d.get("afp", ""), "portafolio": d.get("portafolio", ""),
            "subtipo": subtipo, "moneda": moneda, "nominal": nominal,
            "valor_mercado": vm, "precio": precio, "valorado": valorado}


def construir(fwd_val: pd.DataFrame | None = None,
              trm_val: pd.DataFrame | None = None,
              fut_local: pd.DataFrame | None = None,
              fut_int: pd.DataFrame | None = None,
              swaps_sabana: pd.DataFrame | None = None,
              vpn_swaps: pd.DataFrame | None = None) -> pd.DataFrame:
    """Une los derivados valorados en un esquema comun. Cada argumento es opcional.
    vpn_swaps (ISIN, VPN_Derecho_Calc, VPN_Oblig_Calc) llena el valor de mercado de
    las patas de swap; si no se pasa, los swaps quedan pendientes de valorar."""
    m_der = m_obl = {}
    if vpn_swaps is not None and len(vpn_swaps):
        m_der = dict(zip(vpn_swaps["ISIN"].astype(str),
                         pd.to_numeric(vpn_swaps["VPN_Derecho_Calc"], errors="coerce")))
        m_obl = dict(zip(vpn_swaps["ISIN"].astype(str),
                         pd.to_numeric(vpn_swaps["VPN_Oblig_Calc"], errors="coerce")))
    filas = []
    if fwd_val is not None and len(fwd_val):
        for _, r in fwd_val.iterrows():
            filas.append(_fila("FORWARD", r, r.get("paridad", ""), r.get("paridad", ""),
                               r.get("nominal", 0), r.get("pyg", 0), r.get("tasa_forward", None), True))
    if trm_val is not None and len(trm_val):
        for _, r in trm_val.iterrows():
            filas.append(_fila("FUT-TRM", r, "TRM", r.get("paridad", "USDCOP"),
                               r.get("nominal", 0), r.get("pyg", 0), r.get("tasa_forward", None), True))
    if fut_local is not None and len(fut_local):
        for _, r in fut_local.iterrows():
            filas.append(_fila("FUT-LOCAL", r, str(r.get("subyacente", "")), "COP",
                               r.get("nominal", 0), r.get("valor_mercado", 0),
                               r.get("precio_vector", None), pd.notna(r.get("valor_mercado"))))
    if fut_int is not None and len(fut_int):
        for _, r in fut_int.iterrows():
            valorado = pd.notna(r.get("precio_vector"))
            filas.append(_fila("FUT-INT", r, str(r.get("subyacente", "")), str(r.get("moneda", "")),
                               r.get("nominal", 0), r.get("valor_mercado_cop", 0),
                               r.get("precio_vector", None), valorado))
    if swaps_sabana is not None and len(swaps_sabana):
        # Una fila por pata (derecho/obligacion). El valor de mercado sale del VPN v6:
        # pata derecho se recibe (+VPN_Der), pata obligacion se paga (-VPN_Obl), de
        # modo que las dos patas netean al MtM del swap (como en la hoja Benchmark).
        for _, r in swaps_sabana.iterrows():
            tipo = r.get("tipo_swap", "SWAP")
            etq = "SWAP-CCS" if str(tipo).startswith("CCS") else "SWAP-IRS"
            idc = str(r.get("id_contrato", ""))
            vd = m_der.get(idc)
            vo = m_obl.get(idc)
            vm_der = float(vd) if vd is not None and pd.notna(vd) else None
            vm_obl = -float(vo) if vo is not None and pd.notna(vo) else None
            filas.append(_fila(etq, r, tipo, r.get("der_moneda", ""),
                               r.get("der_nominal", 0), vm_der, r.get("der_tasa_facial", None),
                               vm_der is not None))
            filas.append(_fila(etq, r, tipo, r.get("obl_moneda", ""),
                               r.get("obl_nominal", 0), vm_obl, r.get("obl_tasa_facial", None),
                               vm_obl is not None))
    if not filas:
        return pd.DataFrame(columns=_COLS)
    df = pd.DataFrame(filas)[_COLS]
    for c in ("nominal", "valor_mercado", "precio"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.reset_index(drop=True)
