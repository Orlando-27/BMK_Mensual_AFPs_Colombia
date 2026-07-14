"""
Clasificacion de activos (paso 6c del manual) + deteccion de instrumentos nuevos.

Reimplementa la cadena de formulas de la hoja Clasificacion del BMO:
  S  = nombre limpio (ISIN -> Nemo -> Emisor)
  AQ = nombre limpio (Nemo -> ISIN -> Emisor)
  R  = clave de clasificacion = si Clas SFC in {TDPE,FINDI} -> S
                                sino -> nemos_especiales(S) o AQ
  AB = CLASE DE INVERSION = VLOOKUP(R, clasificacion.csv)   [col clase_inversion]
  AE/AF/AG = ubicacion / clasificacion / riesgo             [misma tabla]

Un activo cuyo R no aparece en la tabla de referencia queda sin clasificar (ND)
= instrumento nuevo -> alerta.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

# Nemos especiales (Diccionario SFC BD:BE): ISIN -> NEMO.
NEMOS_ESPECIALES = {
    "COR42PT00014": "FONINMOBAL", "COV34PT00120": "PEI", "CORJ8PA00013": "HCOLSEL",
    "CORB6PA00015": "ICOLCAP", "CL0002892397": "NUAMCO", "CA1348083025": "CNEC",
    "COT80PT00036": "COT80PT00036", "CORB5PT05263": "FONDARRER",
}
CLAS_USA_ISIN = {"TDPE", "FINDI"}  # clas SFC que usan ISIN directo como clave


def _clean(v) -> str:
    return "" if v is None else str(v).strip()


def _key_S(isin, nemo, emisor) -> str:
    isin, nemo, emisor = _clean(isin), _clean(nemo), _clean(emisor)
    return isin or nemo or emisor


def _key_AQ(isin, nemo, emisor) -> str:
    isin, nemo, emisor = _clean(isin), _clean(nemo), _clean(emisor)
    return nemo or isin or emisor


def key_R(clas_sfc, isin, nemo, emisor) -> str:
    s = _key_S(isin, nemo, emisor)
    if _clean(clas_sfc).upper() in CLAS_USA_ISIN:
        return s
    return NEMOS_ESPECIALES.get(s, _key_AQ(isin, nemo, emisor))


def _collapse(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().upper()


def cargar_referencia(ref_dir: str | Path = "bmk/config/reference") -> dict:
    """Carga clasificacion.csv en un dict {clave_upper: fila}."""
    df = pd.read_csv(Path(ref_dir) / "clasificacion.csv", dtype=str).fillna("")
    ref = {}
    for _, r in df.iterrows():
        k = str(r["nemo_o_isin"]).strip().upper()
        if k and k not in ref:
            ref[k] = {
                "clase_inversion": r.get("clase_inversion", ""),
                "moneda": r.get("moneda", ""),
                "ubicacion": r.get("ubicacion", ""),
                "clasificacion": r.get("clasificacion", ""),
                "riesgo": r.get("riesgo", ""),
                "clase": r.get("clase", ""),
            }
    return ref


def cargar_fallback(ref_dir: str | Path = "bmk/config/reference") -> dict:
    """Carga la tabla de respaldo (Diccionario SFC AM:AY).

    La llave original es de ancho fijo: clas_sfc + emisor(padded) + moneda + tasa.
    Se reconstruye una clave robusta: collapse(clas_sfc+emisor) + '|' + moneda.
    """
    path = Path(ref_dir) / "clasificacion_fallback.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path, dtype=str).fillna("")
    fb = {}
    for _, r in df.iterrows():
        llave, moneda, tasa = r["llave"], r["moneda"], r["tasa_facial"]
        sufijo = _collapse(moneda) + _collapse(tasa)  # ej. 'EURNA'
        pref = _collapse(llave)
        if sufijo and pref.endswith(sufijo):
            pref = pref[: -len(sufijo)].strip()
        clave = f"{pref}|{_collapse(moneda)}"
        if clave not in fb:
            fb[clave] = {
                "clase_inversion": r.get("clase_inversion", ""),
                "moneda": r.get("moneda", ""),
                "ubicacion": r.get("ubicacion", ""),
                "clasificacion": r.get("clasificacion", ""),
                "riesgo": r.get("riesgo", ""),
            }
    return fb


def cargar_clase_a_bucket(ref_dir: str | Path = "bmk/config/reference") -> tuple[dict, dict]:
    """clase_inversion -> (clasificacion, ubicacion) mas frecuente en la referencia.

    Permite completar los buckets (R FIJA / R VARIABLE / CAJA y NACIONAL/INTERNAC.)
    de los activos resueltos solo por la leyenda de clas_sfc.
    """
    df = pd.read_csv(Path(ref_dir) / "clasificacion.csv", dtype=str).fillna("")
    clasif, ubic = {}, {}
    for col_dst, target in (("clasificacion", clasif), ("ubicacion", ubic)):
        modo = (df[df[col_dst] != ""]
                .groupby("clase_inversion")[col_dst]
                .agg(lambda s: s.value_counts().idxmax()))
        target.update(modo.to_dict())
    return clasif, ubic


def cargar_legend_unica(ref_dir: str | Path = "bmk/config/reference") -> dict:
    """clas_sfc -> clase_inversion cuando el codigo mapea a UNA sola clase.

    Tercer nivel de respaldo (para caja/depositos y codigos 1:1). Los codigos
    ambiguos (ej. BOEVS que depende de moneda) se excluyen adrede.
    """
    df = pd.read_csv(Path(ref_dir) / "clas_sfc_legend.csv", dtype=str).fillna("")
    por_codigo: dict[str, set] = {}
    for _, r in df.iterrows():
        c = str(r["clas_sfc"]).strip().upper()
        por_codigo.setdefault(c, set()).add(str(r["clase_inversion"]).strip())
    return {c: next(iter(s)) for c, s in por_codigo.items() if len(s) == 1}


def clasificar(activos: pd.DataFrame, ref_dir: str | Path = "bmk/config/reference") -> pd.DataFrame:
    """Devuelve activos con columnas de clasificacion resueltas + flag is_clasificado."""
    ref = cargar_referencia(ref_dir)
    fb = cargar_fallback(ref_dir)
    legend = cargar_legend_unica(ref_dir)
    out = activos.copy()
    out["clave_R"] = [
        key_R(cs, i, n, e)
        for cs, i, n, e in zip(out["clas_sfc"], out["isin"], out["nemo"], out["emisor"])
    ]
    # 1) Lookup primario por clave R.
    primario = out["clave_R"].str.upper().map(ref)
    # 2) Fallback por (clas_sfc+emisor, moneda) para caja/depositos/casos especiales.
    clave_fb = [f"{_collapse(str(cs) + str(e))}|{_collapse(m)}"
                for cs, e, m in zip(out["clas_sfc"], out["emisor"], out["moneda"])]
    secundario = pd.Series(clave_fb, index=out.index).map(fb)

    # 3) Tercer nivel: clas_sfc con mapeo unico en la leyenda.
    terciario = out["clas_sfc"].str.upper().map(legend)  # str o NaN

    def pick(p, s, t, campo):
        if isinstance(p, dict):
            return p.get(campo, "")
        if isinstance(s, dict):
            return s.get(campo, "")
        if isinstance(t, str) and campo == "clase_inversion":
            return t
        if isinstance(t, str):
            return ""  # clase conocida pero sin ubicacion/riesgo detallado
        return "ND"

    out["is_clasificado"] = [isinstance(p, dict) or isinstance(s, dict) or isinstance(t, str)
                             for p, s, t in zip(primario, secundario, terciario)]
    out["fuente_clasif"] = ["primario" if isinstance(p, dict)
                            else ("fallback" if isinstance(s, dict)
                                  else ("legend" if isinstance(t, str) else "ND"))
                            for p, s, t in zip(primario, secundario, terciario)]
    for campo in ("clase_inversion", "ubicacion", "clasificacion", "riesgo"):
        out[campo] = [pick(p, s, t, campo)
                      for p, s, t in zip(primario, secundario, terciario)]

    # Completar bucket (clasificacion/ubicacion) para los resueltos por leyenda,
    # derivandolo de la clase_inversion segun la referencia.
    clasif_map, ubic_map = cargar_clase_a_bucket(ref_dir)
    falta_clasif = out["is_clasificado"] & (out["clasificacion"].isin(["", "ND"]))
    out.loc[falta_clasif, "clasificacion"] = out.loc[falta_clasif, "clase_inversion"].map(
        clasif_map).fillna(out.loc[falta_clasif, "clasificacion"])
    falta_ubic = out["is_clasificado"] & (out["ubicacion"].isin(["", "ND"]))
    out.loc[falta_ubic, "ubicacion"] = out.loc[falta_ubic, "clase_inversion"].map(
        ubic_map).fillna(out.loc[falta_ubic, "ubicacion"])

    # Ultimo recurso: clas_sfc de renta fija (bonos, tesoros, CDs...) cuyas
    # clase_inversion mapean TODAS a R FIJA. Aunque la clase_inversion sea ambigua
    # (EM CORPORATIVOS vs CORPORATIVO LOCAL, etc.), la CLASIFICACION (R FIJA) no lo
    # es. Evita que instrumentos nuevos de renta fija queden en ND. La ubicacion se
    # deriva de la moneda (COP/UVR = NACIONAL; el resto INTERNACIONAL).
    ldf = pd.read_csv(Path(ref_dir) / "clas_sfc_legend.csv", dtype=str).fillna("")
    lm: dict[str, set] = {}
    for _, r in ldf.iterrows():
        lm.setdefault(str(r["clas_sfc"]).strip().upper(), set()).add(str(r["clase_inversion"]).strip())
    rf_codes = {c for c, cls in lm.items() if cls and all(clasif_map.get(ci) == "R FIJA" for ci in cls)}
    nd = ~out["is_clasificado"] | out["clasificacion"].astype(str).isin(["", "ND"])
    es_rf = nd & out["clas_sfc"].astype(str).str.upper().str.strip().isin(rf_codes)
    if es_rf.any():
        out.loc[es_rf, "clasificacion"] = "R FIJA"
        out.loc[es_rf, "is_clasificado"] = True
        out.loc[es_rf, "fuente_clasif"] = "clas_sfc_rf"
        if "clase" in out.columns:
            out.loc[es_rf, "clase"] = "R FIJA"
        if "moneda" in out.columns:
            mon = out.loc[es_rf, "moneda"].astype(str).str.upper().str.strip()
            out.loc[es_rf, "ubicacion"] = mon.map(
                lambda m: "NACIONAL" if m in ("COP", "COU", "UVR") else "INTERNACIONAL")
        sin_riesgo = es_rf & out["riesgo"].astype(str).isin(["", "ND"])
        out.loc[sin_riesgo, "riesgo"] = "No Reporta"
    return out
