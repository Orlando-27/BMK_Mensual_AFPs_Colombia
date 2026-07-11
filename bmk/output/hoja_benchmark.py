"""
Genera la hoja "Benchmark" del archivo Benchmark (Detallado) — el output que se
importa en la herramienta de actualizacion diaria.

Solo se llenan las columnas de DATOS (info de los activos consolidados). Las
columnas FORMULADAS se dejan vacias: en el archivo destino ya existen y se
arrastran. La estructura (encabezado fila 14) se replica exacta.

Diagnostico de la hoja (Benchmark del Detallado):
  DATOS   : A Portafolio, B AFP, C IDENTIFICADOR, D Nemo, E ISIN, F Clas SFC,
            G Emisor, H F.Compra, I F.Vcto, J Moneda, K Valor Nominal,
            M TASA FACIAL, N CLASE DE INVERSION, O MONEDA, P TASA FACIAL,
            Q UBICACION, R CLASIFICACION, S RIESGO, T VALOR TASA FACIAL,
            U Precio inicial, V Periodicidad, W Precio Actual
  FORMULA : L Vr Mercado INICIAL, X..BU (valor mercado, rentabilidad,
            participacion, duracion, sensibilidad, plazos, exposicion por
            moneda AD-AR y region AS-AZ, notas BE-BL, DM/D REAL, cupones ...)
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# Especificacion de columnas de la hoja Benchmark, en orden.
# (letra, encabezado, campo_origen | None si es formula/vacia)
COLUMNAS = [
    ("A", "Portafolio", "cod_portafolio"),
    ("B", "AFP", "afp"),
    ("C", "IDENTIFICADOR", "identificador"),
    ("D", "Nemo", "nemo"),
    ("E", "ISIN", "isin"),
    ("F", "Clas SFC", "clas_sfc"),
    ("G", "Emisor", "emisor"),
    ("H", "F.Compra", "f_compra"),
    ("I", "F.Vcto", "f_vcto"),
    ("J", "Moneda", "moneda"),
    ("K", "Valor Nominal", "valor_nominal"),
    ("L", "Vr Mercado INICIAL", None),            # FORMULA
    ("M", "TASA FACIAL", "tasa_facial_ind"),
    ("N", "CLASE DE INVERSION", "clase_inversion"),
    ("O", "MONEDA", "moneda"),
    ("P", "TASA FACIAL", None),
    ("Q", "UBICACION", "ubicacion"),
    ("R", "CLASIFICACIÓN", "clasificacion"),
    ("S", "RIESGO", "riesgo"),
    ("T", "VALOR TASA FACIAL", "tasa_facial_valor"),
    ("U", "Precio inicial", "precio_proxy"),
    ("V", "Periodicidad", None),
    ("W", "Precio Actual", None),
]
# De X en adelante todas son formuladas -> se omiten (se dejan vacias).
COLUMNAS_FORMULA = [
    ("X", "Valor de mercado actual"), ("Y", "Rentabilidad acumulada"),
    ("Z", "Participación"), ("AA", "Duración"), ("AB", "sensibilidad"),
    ("AC", "Clasificación por plazos"),
    ("AD", "USD"), ("AE", "COP"), ("AF", "UVR"), ("AG", "EUR"), ("AH", "JPY"),
    ("AI", "BRL"), ("AJ", "CAD"), ("AK", "MXN"), ("AL", "GBP"), ("AM", "CLP"),
    ("AN", "AUD"), ("AO", "CHF"), ("AP", "PEN"), ("AQ", "NOK"), ("AR", "KRW"),
    ("AS", "GLOBAL"), ("AT", "US EQUITY"), ("AU", "EUROPA + UK EQUITY"),
    ("AV", "JAPON EQUITY"), ("AW", "EM GLOBAL"), ("AX", "EM ASIA"),
    ("AY", "LATAM"), ("AZ", "ALTERNATIVE"), ("BA", "Class Acciones"),
    ("BB", "Class caja"), ("BC", "Clasificación RFI"), ("BD", "Valorados"),
    ("BE", "NOTAS RV"), ("BF", "NOTAS RF"), ("BG", "Expo Notas"),
    ("BH", "Plazo Notas"), ("BJ", "Años al vencimiento"),
    ("BK", "Subyacente Nota estructurada"), ("BL", "Delta Nota estructurada"),
    ("BN", "DM"), ("BO", "D REAL"), ("BQ", "Fecha cupón"), ("BR", "Cupón"),
    ("BS", "Festivos"), ("BT", "Fecha cupón hábil"), ("BU", "% Participación"),
]


def _num(s):
    return pd.to_numeric(s, errors="coerce")


def generar(clasificados: pd.DataFrame, incluir_columnas_formula: bool = True) -> pd.DataFrame:
    """Devuelve la hoja Benchmark: filas = activos consolidados, columnas en
    orden; datos llenos, formuladas vacias."""
    df = clasificados.copy()
    # Campos derivados que la hoja necesita.
    isin = df["isin"].astype(str).str.strip()
    nemo = df["nemo"].astype(str).str.strip()
    df["identificador"] = isin.where(isin != "", nemo.where(nemo != "", df["emisor"]))
    nominal = _num(df["valor_nominal"])
    vr = _num(df["vr_mercado"])
    # Precio inicial (proxy) = Vr mercado / Valor nominal (limpio); si no aplica, vacio.
    df["precio_proxy"] = (vr / nominal).where(nominal > 0)

    out = pd.DataFrame()
    for _letra, header, campo in COLUMNAS:
        out[header] = df[campo] if (campo and campo in df.columns) else ""
    if incluir_columnas_formula:
        for _letra, header in COLUMNAS_FORMULA:
            # Encabezado repetido (TASA FACIAL) ya existe; sufijo para no colisionar.
            col = header if header not in out.columns else f"{header} "
            out[col] = ""
    return out.reset_index(drop=True)


def escribir(clasificados: pd.DataFrame, out_dir: str | Path, corte: str) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / f"hoja_Benchmark_{corte}.xlsx"
    hoja = generar(clasificados)
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        hoja.to_excel(xw, sheet_name="Benchmark", index=False)
    return str(dest)
