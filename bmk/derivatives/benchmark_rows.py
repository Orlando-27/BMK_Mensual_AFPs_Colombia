"""
Filas de derivados para la hoja Benchmark (Detallado).

La hoja Benchmark completa tiene 18.638 filas: 15.714 activos/CSA + 2.924
derivados. Este modulo construye las filas de derivados en el MISMO esquema que
los activos clasificados, para anexarlas a la hoja (ver bmk.output.hoja_benchmark).

Estructura replicada del Benchmark de las macros (corte abril, industria):
  - FORWARDS  : 210 filas = 15 (afp,port) x 14 divisas. Una fila por divisa
    (FWUSD..FWKRW) con el valor de mercado NETO de esa pata (derecho +, obligacion -).
    clase=clasif="FORWARD".
  - OPCIONES  : 30 filas = 15 (afp,port) x 2 (OPT USD, OPT EUR). clase=clasif="OPCIONES".
  - SWAPS     : 2 filas por contrato (pata derecho "SWAP TF" y pata obligacion
    "SWAP TV"), clas_sfc IRS/CCS, clase=clasif="SWAP".

Valor de mercado:
  - Forwards: se usa el valor presente de cada pata reportado en el propio archivo
    SFC (Formato_415 col 53 derecho / col 54 obligacion), neteado por divisa. Es el
    valor AL CORTE. Las macros lo revaluan a la fecha de corrida (~10 dias despues)
    con la curva de puntos forward (hoja Curvas / Sensibilidad_Fwds), lo que da
    ~5-16% mas; esa curva es un insumo externo que no viaja en el archivo SFC.
  - Swaps: el VPN por pata lo produce el motor de valoracion v6 (bmk.pricing) con
    las curvas del corte. Si no hay insumos de curvas del corte, la fila queda con
    nominal y sin VPN (vr_mercado NaN) para que la etapa de valoracion lo complete.

Nota de etiquetas: se usa la normalizacion de AFP del pipeline (PORVENIR,
PROTECCION, SKANDIA). El Benchmark de las macros rotula SKANDIA como "OLD MUTUAL"
y PROTECCION con tilde; es un tema de homologacion cosmetica de etiqueta.
"""
from __future__ import annotations

import pandas as pd

from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp, _cod_portafolio

# Plantilla fija de 14 divisas de forward, en el orden del Benchmark de las macros.
FWD_DIVISAS = ["USD", "COP", "EUR", "JPY", "BRL", "CAD", "MXN",
               "GBP", "CLP", "AUD", "CHF", "PEN", "NOK", "KRW"]
# Plantilla fija de opciones.
OPT_DIVISAS = ["USD", "EUR"]


def _col(df: pd.DataFrame, key: str) -> pd.Series:
    return df.iloc[:, COL_415[key]]


def _norm_cur(s: pd.Series) -> pd.Series:
    up = s.astype(str).str.strip().str.upper()
    return up.replace({"COU": "UVR"})


def _base_415(f415: pd.DataFrame) -> pd.DataFrame:
    """Proyeccion economica de Formato_415 (una fila por contrato)."""
    tipo = pd.to_numeric(_col(f415, "tipo_derivado"), errors="coerce")
    out = pd.DataFrame({
        "tipo": tipo,
        "id_contrato": _col(f415, "numero_contrato").astype(str),
        "afp": _col(f415, "entidad").map(normalizar_afp),
        "port": _col(f415, "nombre").map(_cod_portafolio),
        "mder": _norm_cur(_col(f415, "moneda_derecho")),
        "nder": pd.to_numeric(_col(f415, "nominal_derecho"), errors="coerce"),
        "mobl": _norm_cur(_col(f415, "moneda_obligacion")),
        "nobl": pd.to_numeric(_col(f415, "nominal_obligacion"), errors="coerce"),
        "vp_der": pd.to_numeric(f415.iloc[:, 53], errors="coerce"),
        "vp_obl": pd.to_numeric(f415.iloc[:, 54], errors="coerce"),
    })
    return out[out["tipo"].notna()].reset_index(drop=True)


def _combos(base: pd.DataFrame, solo_industria: bool) -> list[tuple[str, str]]:
    """(afp, port) con al menos un derivado; industria excluye COLFONDOS."""
    b = base[base["afp"] != "COLFONDOS"] if solo_industria else base
    combos = b[["afp", "port"]].drop_duplicates()
    combos = combos[(combos["afp"] != "ND") & (combos["port"] != "ND")]
    return sorted(map(tuple, combos.to_numpy()))


def _fila(afp, port, clas, clase, nemo, vr_mercado, nominal=0.0, moneda="",
          id_contrato="", pata="") -> dict:
    return {
        "cod_portafolio": port, "afp": afp, "nemo": nemo, "isin": "",
        "clas_sfc": clas, "emisor": clas, "f_compra": "", "f_vcto": "",
        "moneda": moneda, "valor_nominal": nominal,
        "tasa_facial_ind": "", "tasa_facial_valor": "",
        "vr_mercado": vr_mercado, "clase_inversion": clase,
        "ubicacion": "INTERNACIONAL", "clasificacion": clase,
        "riesgo": "No Reporta", "is_clasificado": True, "fuente_clasif": "derivado",
        "id_contrato": str(id_contrato), "pata": pata,
    }


def _forwards(base: pd.DataFrame, combos) -> pd.DataFrame:
    fw = base[base["tipo"] == 1]
    # valor presente neto por (afp, port, divisa): +pata derecho, -pata obligacion.
    der = fw.groupby(["afp", "port", "mder"])["vp_der"].sum()
    obl = fw.groupby(["afp", "port", "mobl"])["vp_obl"].sum()
    filas = []
    for afp, port in combos:
        for cur in FWD_DIVISAS:
            v = float(der.get((afp, port, cur), 0.0)) - float(obl.get((afp, port, cur), 0.0))
            filas.append(_fila(afp, port, f"FW{cur}", "FORWARD", f"FW{cur}", v, moneda=cur))
    return pd.DataFrame(filas)


def _opciones(base: pd.DataFrame, combos) -> pd.DataFrame:
    # Sin insumo de valoracion de opciones industria en el archivo -> valor 0
    # (en el Benchmark de las macros la mayoria son 0; se revaluan aparte).
    filas = []
    for afp, port in combos:
        for cur in OPT_DIVISAS:
            filas.append(_fila(afp, port, f"OPT {cur}", "OPCIONES", f"OPT {cur}", 0.0, moneda=cur))
    return pd.DataFrame(filas)


def _swaps(base: pd.DataFrame, solo_industria: bool) -> pd.DataFrame:
    sw = base[base["tipo"].isin([16, 17])].copy()
    if solo_industria:
        sw = sw[sw["afp"] != "COLFONDOS"]
    sw = sw[(sw["afp"] != "ND") & (sw["port"] != "ND")].reset_index(drop=True)
    # tipo IRS si ambas patas misma moneda, si no CCS (BMO: hoja SWAP col X).
    clas = pd.Series(["IRS" if d == o else "CCS" for d, o in zip(sw["mder"], sw["mobl"])])
    filas = []
    for i, r in sw.iterrows():
        cs = clas.iloc[i]
        # Pata derecho (SWAP TF) y pata obligacion (SWAP TV). VPN pendiente de
        # valoracion (motor v6 con curvas del corte) -> vr_mercado NaN.
        filas.append(_fila(r["afp"], r["port"], cs, "SWAP", "SWAP TF",
                           float("nan"), nominal=r["nder"], moneda=r["mder"],
                           id_contrato=r["id_contrato"], pata="DER"))
        filas.append(_fila(r["afp"], r["port"], cs, "SWAP", "SWAP TV",
                           float("nan"), nominal=r["nobl"], moneda=r["mobl"],
                           id_contrato=r["id_contrato"], pata="OBL"))
    return pd.DataFrame(filas)


def inyectar_vpn_swaps(der: pd.DataFrame, vpn_df: pd.DataFrame) -> pd.DataFrame:
    """Llena vr_mercado de las filas de swap con el VPN del motor v6.
    vpn_df: columnas ISIN, VPN_Derecho_Calc, VPN_Oblig_Calc (por contrato)."""
    if der is None or der.empty or vpn_df is None or vpn_df.empty:
        return der
    m_der = dict(zip(vpn_df["ISIN"].astype(str), pd.to_numeric(vpn_df["VPN_Derecho_Calc"], errors="coerce")))
    m_obl = dict(zip(vpn_df["ISIN"].astype(str), pd.to_numeric(vpn_df["VPN_Oblig_Calc"], errors="coerce")))
    out = der.copy()
    es_swap = out["clasificacion"].astype(str).str.upper() == "SWAP"
    idc = out.get("id_contrato", pd.Series("", index=out.index)).astype(str)
    pata = out.get("pata", pd.Series("", index=out.index)).astype(str)
    nueva = out["vr_mercado"].copy()
    for i in out.index[es_swap]:
        m = m_der if pata[i] == "DER" else m_obl
        v = m.get(idc[i])
        if v is not None and pd.notna(v):
            nueva[i] = v
    out["vr_mercado"] = nueva
    return out


def generar(formato_415: pd.DataFrame, solo_industria: bool = True) -> pd.DataFrame:
    """Filas de derivados (forwards + opciones + swaps) en el esquema de activos."""
    if formato_415 is None or formato_415.empty:
        return pd.DataFrame()
    base = _base_415(formato_415)
    combos = _combos(base, solo_industria)
    partes = [_forwards(base, combos), _opciones(base, combos),
              _swaps(base, solo_industria)]
    out = pd.concat([p for p in partes if not p.empty], ignore_index=True)
    return out.reset_index(drop=True)
