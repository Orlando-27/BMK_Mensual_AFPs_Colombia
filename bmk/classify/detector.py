"""
Detector de instrumentos nuevos / no contemplados -> alertas.

Un activo es "nuevo" cuando su clave de clasificacion no resuelve contra la
tabla de referencia (clase_inversion = ND). Ademas detecta fondos/notas sin
mapeo de moneda-region (CPRVI) y codigos Clas SFC nunca vistos.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from bmk.alerts.registry import RegistroAlertas


def _clas_sfc_conocidos(ref_dir: Path) -> set[str]:
    df = pd.read_csv(ref_dir / "clas_sfc_legend.csv", dtype=str).fillna("")
    return {str(x).strip().upper() for x in df["clas_sfc"]}


def detectar(clasificados: pd.DataFrame, registro: RegistroAlertas,
             ref_dir: str | Path = "bmk/config/reference") -> dict:
    """Agrega alertas al registro. Devuelve conteos."""
    ref_dir = Path(ref_dir)
    colmap = {"portafolio": "portafolio", "afp": "afp", "isin": "isin",
              "nemo": "nemo", "clas_sfc": "clas_sfc", "moneda": "moneda",
              "valor_mercado": "vr_mercado"}

    # 1) Instrumentos sin clasificar (ND).
    nd = clasificados[~clasificados["is_clasificado"]]
    n_nd = registro.agregar_lote(
        nd, etapa="clasificacion", tipo="INSTRUMENTO_NUEVO_SIN_CLASIFICAR",
        severidad="WARN",
        descripcion="Activo cuyo ISIN/Nemo/Clas SFC no esta en la tabla de clasificacion de referencia",
        accion_sugerida="Clasificar el titulo (clase de inversion, moneda, region, riesgo) y agregar a clasificacion.csv",
        colmap=colmap)

    # 2) Clas SFC nunca visto (codigo de clasificacion nuevo de la SFC).
    conocidos = _clas_sfc_conocidos(ref_dir)
    nuevos_cs = clasificados[~clasificados["clas_sfc"].str.upper().isin(conocidos)]
    nuevos_cs_u = nuevos_cs.drop_duplicates("clas_sfc")
    n_cs = registro.agregar_lote(
        nuevos_cs_u, etapa="clasificacion", tipo="CODIGO_CLAS_SFC_DESCONOCIDO",
        severidad="WARN",
        descripcion="Codigo Clas SFC no presente en el diccionario conocido",
        accion_sugerida="Validar con el equipo el nuevo codigo SFC y agregarlo al diccionario",
        colmap=colmap)

    return {"instrumentos_nuevos_ND": n_nd, "clas_sfc_desconocidos": n_cs,
            "isines_ND_unicos": int(nd["clave_R"].nunique())}
