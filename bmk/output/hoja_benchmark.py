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
    ("L", "Vr Mercado INICIAL", "vr_mercado"),    # valor de mercado del corte (SFC)
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


_NOM_MAP: dict | None = None


def _cargar_nominal_map(ref_dir: str | Path = "bmk/config/reference") -> dict:
    """clas_sfc -> columna de nominal (Diccionario SFC Z:AB de las macros)."""
    global _NOM_MAP
    if _NOM_MAP is None:
        path = Path(ref_dir) / "nominal_por_clas_sfc.csv"
        if path.exists():
            d = pd.read_csv(path, dtype=str).fillna("")
            _NOM_MAP = {str(a).strip().upper(): str(b).strip()
                        for a, b in zip(d["clas_sfc"], d["nominal_col"])}
        else:
            _NOM_MAP = {}
    return _NOM_MAP


def _nominal_benchmark(df: pd.DataFrame, vr: pd.Series,
                       ref_dir: str | Path = "bmk/config/reference") -> pd.Series:
    """Valor Nominal segun la tabla 'NOMINAL A USAR' de las macros:
    Valor nominal (AK) / Nominal residual (AM) / No. Unidades (AN) por clas_sfc.
    Para depositos (caja) con nominal 0 usa el monto: Vr Mercado / FX = valor en
    moneda del activo (COP para locales, moneda extranjera para foraneos).
    """
    nmap = _cargar_nominal_map(ref_dir)
    def _c(name):
        return _num(df[name]) if name in df.columns else pd.Series(0.0, index=df.index)
    cols = {"valor_nominal": _c("valor_nominal"),
            "valor_nominal_residual": _c("valor_nominal_residual"),
            "nro_acciones": _c("nro_acciones")}
    ext = _c("vr_moneda_ext")
    cs = df["clas_sfc"].astype(str).str.upper()
    # columna base segun el mapa (default: valor_nominal)
    colname = cs.map(nmap).fillna("valor_nominal")
    nominal = pd.Series(0.0, index=df.index)
    for name, serie in cols.items():
        nominal = nominal.where(colname != name, serie)
    # Fallback depositos/caja con nominal 0: monto en moneda del activo.
    clasif = df["clasificacion"].astype(str).str.upper()
    dep = (nominal.fillna(0) == 0) & (clasif == "CAJA")
    monto = ext.where(ext.abs() > 0, vr)
    nominal = nominal.where(~dep, monto)
    return nominal


def generar(clasificados: pd.DataFrame, incluir_columnas_formula: bool = True,
            solo_industria: bool = True, umbral_min: float = 1000.0,
            csa: pd.DataFrame | None = None,
            derivados: pd.DataFrame | None = None) -> pd.DataFrame:
    """Devuelve la hoja Benchmark: filas = activos consolidados, columnas en
    orden; datos llenos, formuladas vacias.

    Para replicar la hoja Benchmark de la herramienta diaria:
      - solo_industria=True excluye COLFONDOS (va en su propia hoja).
      - umbral_min filtra activos con |Vr Mercado| < umbral (paso "Mercado<1000"
        del manual; esos van a Eliminados, no al Benchmark).
      - csa: filas de Cuentas CSA (caja internacional) a anexar (ver bmk.prepare.csa).
      - derivados: filas de forwards/opciones/swaps a anexar (ver
        bmk.derivatives.benchmark_rows). Se anexan despues del filtro <1000
        porque incluyen plantillas de divisa con valor 0 que deben conservarse.
    """
    df = clasificados.copy()
    if solo_industria:
        df = df[df["afp"] != "COLFONDOS"]
    if umbral_min:
        vr_ = pd.to_numeric(df["vr_mercado"], errors="coerce").fillna(0)
        df = df[vr_.abs() >= umbral_min]
    extras = [df]
    if csa is not None and not csa.empty:
        extras.append(csa)
    if derivados is not None and not derivados.empty:
        extras.append(derivados)
    df = pd.concat(extras, ignore_index=True) if len(extras) > 1 else df
    df = df.reset_index(drop=True)
    # Campos derivados que la hoja necesita.
    isin = df["isin"].astype(str).str.strip()
    nemo = df["nemo"].astype(str).str.strip()
    df["identificador"] = isin.where(isin != "", nemo.where(nemo != "", df["emisor"]))
    vr = _num(df["vr_mercado"])
    nominal = _nominal_benchmark(df, vr)
    df["valor_nominal"] = nominal
    # Precio inicial = Vr mercado / Nominal (=1 en caja, precio por unidad en RF/RV).
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


def escribir(clasificados: pd.DataFrame, out_dir: str | Path, corte: str,
             csa: pd.DataFrame | None = None,
             derivados: pd.DataFrame | None = None) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / f"hoja_Benchmark_{corte}.xlsx"
    hoja = generar(clasificados, csa=csa, derivados=derivados)
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        hoja.to_excel(xw, sheet_name="Benchmark", index=False)
    return str(dest)
