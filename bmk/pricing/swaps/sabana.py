"""
Sabana de condiciones faciales de los swaps (formato SWAPIND del motor v6).

Construye, desde Formato_415 (tipo 16/17), la tabla de condiciones faciales que
consume el motor de valoracion v6 (swap_valuator_v6). Las condiciones faciales
son las MISMAS para las 3 fechas de valoracion (el contrato no cambia); solo
cambian los insumos de mercado (curvas, TRM/UVR/EURCOP) por fecha.

Mapeo de columnas Formato_415 (0-based) -> condicion facial (nombres reales del
formato de la SFC):
  0  entidad (AFP)                 11 numero de contrato (ISIN interno)
  4  cod. patrimonio               5  nombre patrimonio (-> portafolio)
  22 fecha celebracion (F.Compra)  23 fecha vencimiento (F.Vcto)
  24 fecha liquidacion             45 base de calculo (UNA para ambas patas)
  Pata DERECHO (flujos a RECIBIR): 26 moneda, 27 nominal,
                 37 codigo tasa (FS/SOF/IBR), 38 indicador tasa,
                 39 valor tasa fija o spread, 40 PERIODICIDAD
  Pata OBLIG. (flujos a PAGAR):    28 moneda, 29 nominal,
                 41 codigo tasa, 42 indicador tasa,
                 43 valor tasa fija o spread, 44 PERIODICIDAD

Codigos:
  - Codigo tasa (col 37/41): FS = tasa fija, SOF = SOFR, IBR = IBR.
  - Periodicidad (col 40/44): decodificada empiricamente contra la Period. del
    macro (SWAPIND, 1436 swaps): 3=Trimestral, 5=Semestral, 6=Anual, 7=Anual.
  - Base (col 45): el motor v6 usa day-count por moneda (DAY_BASE), no esta
    columna; se expone solo para la sabana. base 1/2 -> ACT/360 / ACT/365 aprox.

NOTA: en la version anterior periodicidad se leia de col 38/42 (que son el
INDICADOR fija/variable) y base de col 40/44 (que son la PERIODICIDAD real), lo
que daba un schedule de cupones equivocado e inflaba la pata flotante ~10%.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from bmk.ingest.sfc import COL_415
from bmk.prepare.normalize import normalizar_afp, _cod_portafolio

# Indice de la tasa facial (col 37/41). Fija vs flotantes.
INDICE = {"FS": "Fija", "SOF": "SOFR", "IBR": "IBR", "DTF": "DTF", "IPC": "IPC", "UVR": "UVR"}

# Periodicidad (col 40/44) — decodificada contra la Period. del macro (SWAPIND).
PERIODICIDAD = {3: "Trimestral", 5: "Semestral", 6: "Anual", 7: "Anual"}
# Base de calculo (col 45) — cosmetica (el motor v6 usa DAY_BASE por moneda).
BASE = {1: "ACT/360", 2: "ACT/365"}


def _tipo_swap(ind_der, ind_obl, mon_der, mon_obl) -> str:
    """TIPO SWAP en el vocabulario del motor v6 (IRS IBRCOP, CCS COPUSD, ...)."""
    md, mo = str(mon_der).upper().strip(), str(mon_obl).upper().strip()
    md = "UVR" if md in ("COU", "UVR") else md
    mo = "UVR" if mo in ("COU", "UVR") else mo
    inds = {str(ind_der).upper().strip(), str(ind_obl).upper().strip()}
    if md == mo == "COP":
        return "IRS IBRCOP"
    if md == mo == "USD":
        return "IRS SOFUSD"
    par = {md, mo}
    if par == {"COP", "USD"}:
        return "CCS COPUSD"
    if par == {"COP", "UVR"}:
        return "CCS COPUVR"
    if par == {"UVR", "USD"}:
        return "CCS UVRUSD"
    if par == {"COP", "EUR"}:
        return "CCS COPEUR"
    return f"OTRO {md}/{mo}"


def construir(formato_415: pd.DataFrame, solo_industria: bool = True) -> pd.DataFrame:
    """Sabana de condiciones faciales (una fila por swap) desde Formato_415."""
    def c(idx):
        return formato_415.iloc[:, idx]
    tipo = pd.to_numeric(c(COL_415["tipo_derivado"]), errors="coerce")
    def num(s):
        return pd.to_numeric(s, errors="coerce")
    df = pd.DataFrame({
        "afp": c(0).map(normalizar_afp),
        "portafolio": c(5).map(_cod_portafolio),
        "cod_patrimonio": c(4),
        "id_contrato": c(11).astype(str),
        "f_compra": num(c(22)),
        "f_vcto": num(c(23)),
        "f_liquidacion": num(c(24)),
        # Pata derecho (flujos a RECIBIR)
        "der_moneda": c(26).astype(str).str.upper().str.strip(),
        "der_nominal": num(c(27)),
        "der_indice_cod": c(37).astype(str).str.upper().str.strip(),
        "der_periodicidad_cod": num(c(40)),
        "der_tasa_facial": num(c(39)),
        "der_base_cod": num(c(45)),
        # Pata obligacion (flujos a PAGAR)
        "obl_moneda": c(28).astype(str).str.upper().str.strip(),
        "obl_nominal": num(c(29)),
        "obl_indice_cod": c(41).astype(str).str.upper().str.strip(),
        "obl_periodicidad_cod": num(c(44)),
        "obl_tasa_facial": num(c(43)),
        "obl_base_cod": num(c(45)),
    })
    df = df[tipo.isin([16, 17])].reset_index(drop=True)
    if solo_industria:
        df = df[df["afp"] != "COLFONDOS"].reset_index(drop=True)
    # Decodificaciones legibles.
    df["tipo_swap"] = [_tipo_swap(a, b, m, n) for a, b, m, n in
                       zip(df["der_indice_cod"], df["obl_indice_cod"],
                           df["der_moneda"], df["obl_moneda"])]
    df["der_indice"] = df["der_indice_cod"].map(lambda x: INDICE.get(x, x))
    df["obl_indice"] = df["obl_indice_cod"].map(lambda x: INDICE.get(x, x))
    df["der_periodicidad"] = df["der_periodicidad_cod"].map(PERIODICIDAD)
    df["obl_periodicidad"] = df["obl_periodicidad_cod"].map(PERIODICIDAD)
    df["der_base"] = df["der_base_cod"].map(BASE)
    df["obl_base"] = df["obl_base_cod"].map(BASE)
    return df


# Codigo de frecuencia (Mo) segun periodicidad, como en SWAPIND (AV/SV/TV/MV).
_MO = {"Anual": "AV", "Semestral": "SV", "Trimestral": "TV", "Mensual": "MV",
       "Bimestral": "BV", "Cuatrimestral": "CV", "Al vencimiento": "AV"}


def _clase(indice_cod: str) -> str:
    """IND CLASE de SWAPIND: SWAP TF (tasa fija) si FS, SWAP TV si flotante."""
    return "SWAP TF" if str(indice_cod).upper().strip() == "FS" else "SWAP TV"


def _fecha_str(serie) -> pd.Series:
    """AAAAMMDD (del SFC) -> texto ISO 'AAAA-MM-DD'. Con guiones pandas NO lo
    re-infiere como entero al leer el CSV (que rompia el parseo de fecha del
    motor: un entero 20211021 da 1970, el ISO da 2021-10-21)."""
    n = pd.to_numeric(serie, errors="coerce")

    def _fmt(x):
        if pd.isna(x):
            return ""
        v = int(x)
        return f"{v // 10000:04d}-{(v // 100) % 100:02d}-{v % 100:02d}"
    return n.map(_fmt)


def a_swapind(sab: pd.DataFrame) -> pd.DataFrame:
    """Mapea la sabana al layout exacto de la hoja SWAPIND (Calculadora Swap).
    Las columnas de VPN/Utilidad quedan vacias (las llena el motor v6)."""
    def col(campo):
        return sab[campo] if campo in sab.columns else ""
    f_compra = _fecha_str(sab["f_compra"])
    f_vcto = _fecha_str(sab["f_vcto"])
    facial_der = pd.to_numeric(sab["der_tasa_facial"], errors="coerce") / 100.0
    facial_obl = pd.to_numeric(sab["obl_tasa_facial"], errors="coerce") / 100.0
    # Nombres EXACTOS (con espacios) que lee el motor v6 desde SWAPIND.
    out = pd.DataFrame({
        "TIPO SWAP": col("tipo_swap"), "ISIN": col("id_contrato"), "EMISOR": col("afp"),
        "F.Compra": f_compra, "Emision ": f_compra, "F.Vcto  ": f_vcto,
        "IND CLASE": sab["der_indice_cod"].map(_clase), "IND. Tfacial": col("der_indice_cod"),
        "Facial": facial_der, "Period.": col("der_periodicidad"),
        "Mo": sab["der_periodicidad"].map(_MO), "Vr Nominal DER": col("der_nominal"),
        "Moneda": col("der_moneda"), "VPN Derecho ME": "",
        "IND CLASE.1": sab["obl_indice_cod"].map(_clase), "IND. Tfacial.1": col("obl_indice_cod"),
        "Facial         ": facial_obl, "Period..1": col("obl_periodicidad"),
        "Mo.1": sab["obl_periodicidad"].map(_MO), "Vr Nominal": col("obl_nominal"),
        "Moneda.1": col("obl_moneda"), "VPN Obligacion ME": "",
        "Utilidad/Perdida": "", "POR": col("portafolio"), "IBR o/n": "",
        "Base Derecho": col("der_base"), "Base Obligación": col("obl_base"),
        "FECHA": "", "FCD": "", "FCO": "",
    })
    return out


def escribir(formato_415: pd.DataFrame, out_dir: str | Path, corte: str,
             solo_industria: bool = True) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    df = construir(formato_415, solo_industria)
    dest = out_dir / f"sabana_condiciones_faciales_swaps_{corte}.xlsx"
    with pd.ExcelWriter(dest, engine="xlsxwriter") as xw:
        a_swapind(df).to_excel(xw, sheet_name="SWAPIND", index=False)
    return str(dest)
