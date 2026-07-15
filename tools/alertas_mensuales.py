"""
Alertas mensuales del Benchmark: compara el corte actual contra el mes anterior
(Formato_351 de ambos SFC) y produce las alertas que la mesa necesita revisar:

  1) INSTRUMENTOS NUEVOS: ISINs que aparecen este mes y no existian el mes
     anterior. Para cada uno se SUGIERE la region (GLOBAL / US EQUITY /
     EUROPA+UK / JAPON EQUITY / EM ASIA / EM GLOBAL / LATAM / ALTERNATIVE) y la
     moneda, para que se clasifiquen en las columnas de regiones/monedas de la
     hoja Benchmark. Se prioriza renta variable (es lo que exige region).

  2) CAMBIOS SIGNIFICATIVOS: ISINs presentes ambos meses cuyo valor de mercado
     cambio mas de un umbral (default 30%), con nominal, precio y clasificacion
     (p. ej. una nota estructurada que se amortizo a la mitad).

Uso:
    python3 tools/alertas_mensuales.py \
        --actual insumos/sfc/2026_06portainvdeta.xls \
        --anterior insumos/sfc/2026_05portainvdeta.xls \
        --out salidas/2026-06-30/alertas_mensuales_2026-06-30.xlsx
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import numpy as np
import pandas as pd

from bmk.ingest import sfc
from bmk.prepare import normalize
from bmk.classify import classifier

# ── Sugeridor de REGION (cols 44-51 de la hoja Benchmark) por palabras clave ──
# El orden importa: la primera regla que hace match gana.
REGLAS_REGION = [
    ("ALTERNATIVE", r"\bFCP\b|CAPITAL PRIVADO|PRIVATE EQUITY|INMOBILIAR|REAL ESTATE|INFRAESTRUCT|HEDGE"),
    ("JAPON EQUITY", r"JAP[OÓ]N|JAPAN|NIKKEI|TOPIX"),
    ("EM ASIA", r"CHINA|CSI|ASIA|KOREA|COREA|INDIA|TAIWAN|EMERGING ASIA|HANG SENG|EM ASIA"),
    ("US EQUITY", r"USA\b|\bUS\b|ESTADOS UNIDOS|S&P|SP500|SPX|NASDAQ|RUSSELL|\bDOW\b|MSCI ?USA|US EQUITY"),
    ("EUROPA + UK EQUITY", r"EUROPE|EUROPA|EURO STOXX|STOXX|\bUK\b|REINO UNIDO|DAX|FTSE|CAC"),
    ("LATAM", r"LATAM|LATIN|BRAZIL|BRASIL|MEXICO|M[EÉ]XICO|COLOMBIA|CHILE|PERU|ECOPETROL|ICOLCAP|COLCAP"),
    ("EM GLOBAL", r"EMERGING|EMERGENTE|\bEM\b|MSCI EM"),
    ("GLOBAL", r"GLOBAL|WORLD|MUNDO|ACWI|DEVELOPED|MSCI WORLD"),
]

# Moneda -> pista de region cuando el nombre no alcanza.
MONEDA_REGION = {"JPY": "JAPON EQUITY", "BRL": "LATAM", "MXN": "LATAM", "CLP": "LATAM",
                 "COP": "LATAM", "COU": "LATAM", "EUR": "EUROPA + UK EQUITY",
                 "GBP": "EUROPA + UK EQUITY", "CHF": "EUROPA + UK EQUITY",
                 "USD": "US EQUITY"}


def sugerir_region(nemo: str, emisor: str, moneda: str, clase: str) -> str:
    txt = f"{nemo} {emisor} {clase}".upper()
    for region, patron in REGLAS_REGION:
        if re.search(patron, txt):
            return region
    return MONEDA_REGION.get(str(moneda).upper().strip(), "GLOBAL") + " (por moneda, revisar)"


def cargar(path: str) -> pd.DataFrame:
    arch = sfc.leer(path)
    act = normalize.extraer_activos(arch.formato_351).reset_index(drop=True)
    clas = classifier.clasificar(act).reset_index(drop=True)
    clas["precio"] = pd.to_numeric(arch.formato_351.iloc[:, 63], errors="coerce").reset_index(drop=True)
    clas["vr_mercado"] = pd.to_numeric(clas["vr_mercado"], errors="coerce")
    clas["valor_nominal"] = pd.to_numeric(clas["valor_nominal"], errors="coerce")
    clas["isin_u"] = clas["isin"].astype(str).str.upper().str.strip()
    return clas


def _agg(df: pd.DataFrame) -> pd.DataFrame:
    """Consolida por ISIN (suma valor, primer nemo/emisor/clas)."""
    return df.groupby("isin_u").agg(
        nemo=("nemo", "first"), emisor=("emisor", "first"),
        moneda=("moneda", "first"), clasificacion=("clasificacion", "first"),
        clase_inversion=("clase_inversion", "first"),
        vr=("vr_mercado", "sum"), nominal=("valor_nominal", "sum"),
        precio=("precio", "mean"), n=("isin_u", "size")).reset_index()


def alertas(actual: pd.DataFrame, anterior: pd.DataFrame, umbral: float = 0.30):
    a = _agg(actual[actual["isin_u"].ne("") & actual["isin_u"].ne("NAN")])
    p = _agg(anterior[anterior["isin_u"].ne("") & anterior["isin_u"].ne("NAN")])
    prev_isins = set(p["isin_u"])
    prev_vr = dict(zip(p["isin_u"], p["vr"]))

    # 1) Nuevos
    nuevos = a[~a["isin_u"].isin(prev_isins)].copy()
    nuevos["region_sugerida"] = [sugerir_region(n, e, m, c) for n, e, m, c in
                                 zip(nuevos["nemo"], nuevos["emisor"], nuevos["moneda"],
                                     nuevos["clase_inversion"])]
    nuevos["vr_MM"] = (nuevos["vr"] / 1e6).round(1)
    # renta variable primero, luego por valor
    nuevos["_rv"] = nuevos["clasificacion"].astype(str).str.upper().eq("R VARIABLE")
    nuevos = nuevos.sort_values(["_rv", "vr"], ascending=[False, False])
    nuevos = nuevos[["isin_u", "nemo", "emisor", "clasificacion", "clase_inversion",
                     "moneda", "region_sugerida", "vr_MM", "n"]]

    # 2) Cambios significativos
    comun = a[a["isin_u"].isin(prev_isins)].copy()
    comun["vr_prev"] = comun["isin_u"].map(prev_vr)
    comun["delta_%"] = ((comun["vr"] - comun["vr_prev"]) /
                        comun["vr_prev"].abs().replace(0, np.nan) * 100)
    cambios = comun[comun["delta_%"].abs() >= umbral * 100].copy()
    cambios["vr_actual_MM"] = (cambios["vr"] / 1e6).round(1)
    cambios["vr_prev_MM"] = (cambios["vr_prev"] / 1e6).round(1)
    cambios = cambios.sort_values("delta_%", key=lambda s: s.abs(), ascending=False)
    cambios = cambios[["isin_u", "nemo", "emisor", "clasificacion", "moneda",
                       "vr_prev_MM", "vr_actual_MM", "delta_%"]]
    return nuevos, cambios


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--actual", required=True, help="SFC del corte actual")
    ap.add_argument("--anterior", required=True, help="SFC del mes anterior")
    ap.add_argument("--umbral", type=float, default=0.30, help="Umbral de cambio (0.30 = 30%)")
    ap.add_argument("--out")
    a = ap.parse_args()
    act = cargar(a.actual)
    prev = cargar(a.anterior)
    nuevos, cambios = alertas(act, prev, a.umbral)
    pd.options.display.width = 220
    pd.options.display.max_rows = 80
    print(f"=== INSTRUMENTOS NUEVOS ({len(nuevos)}) — renta variable primero ===")
    print(nuevos.to_string(index=False))
    print(f"\n=== CAMBIOS SIGNIFICATIVOS (|Δ| >= {a.umbral*100:.0f}%) — {len(cambios)} ===")
    print(cambios.to_string(index=False))
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
            nuevos.to_excel(xw, sheet_name="Instrumentos nuevos", index=False)
            cambios.to_excel(xw, sheet_name="Cambios significativos", index=False)
        print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
