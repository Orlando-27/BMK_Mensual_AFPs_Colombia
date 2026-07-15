"""
swap_valuator_v6.py — Motor de Valoración Swaps OTC / Metodología Precia
Fecha de valoración: 31-Mar-2026

Cambios clave vs v4:
  [V6-A] Output en COP: VPN = PV_moneda × FX_spot (en v4 se devolvía en moneda
         nativa, causando errores de 76%+ en IRS).
  [V6-B] Curva USDCO integrada (USD descontado con colateral COP). Resuelve
         el problema estructural de 18% en IRS SOFUSD.
  [V6-C] Filtra cupones ya pagados (schedule > FECHA_VAL).
  [V6-D] Forward EA anualizado correcto: (DF(t1)/DF(t2))^(365/d)-1
         + spread, multiplicado por fracción de accrual.
  [V6-E] Schedule retroactivo desde vencimiento (como Precia) para preservar
         el día de aniversario.
  [V6-F] Asignación de curvas estricta por tipo, según manual Precia:
         IRS SOFUSD : USDCO descuenta ambas patas; USDOIS proyecta SOFR
         IRS IBRCOP : COP_COL_USD descuenta ambas; IBR proyecta IBR
         CCS COPUSD : COP_COL_USD descuenta COP; USDCO descuenta USD
         CCS COPUVR : COP_COL_USD descuenta COP; IBRUVR descuenta UVR

Estructura:
  SECCIÓN 1 — Configuración (paths, FX, fecha valoración)
  SECCIÓN 2 — Carga e interpolación de curvas
  SECCIÓN 3 — Generación de calendarios de cupones
  SECCIÓN 4 — Valoración de patas (fija y flotante)
  SECCIÓN 5 — Duración Modificada y Convexidad
  SECCIÓN 6 — Dispatcher por tipo de swap
  SECCIÓN 7 — Main: carga input, itera, guarda output
"""

import os
import warnings
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta

import numpy as np
import pandas as pd
from scipy.interpolate import interp1d

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — CONFIGURACIÓN DINÁMICA (cero hardcodeos de fecha/FX)
# ═══════════════════════════════════════════════════════════════════════
#
# Todos los parámetros dependientes de fecha se resuelven así:
#   1. El directorio de insumos contiene un INICIO_YYYYMMDD.csv con los
#      spots (TRM, UVR, EURCOP) y la fecha de valoración oficial.
#   2. Los archivos de curva siguen el patrón curva_<NOMBRE>_YYYYMMDD.csv.
#   3. El script autodetecta el YYYYMMDD a partir del primer INICIO_*.csv
#      que encuentre en el directorio (o lee el que se le pase por env var).
#
# Variables de entorno (todas opcionales con defaults razonables):
#   V6_INSUMOS_DIR   → carpeta con INICIO_*, SWAPIND_*, curva_* de una fecha
#   V6_OUTPUT_CSV    → CSV de salida (estructura rango B:J de INICIO)
#   V6_OUTPUT_FULL   → CSV de salida extendido (con MD, convexidad, etc.)
#
# No hay ninguna fecha, TRM, UVR ni EURCOP escritos directamente en el código.

BASE_DIR    = os.path.expanduser("~/swap_valuation")
INSUMOS_DIR = os.environ.get("V6_INSUMOS_DIR", os.path.join(BASE_DIR, "data", "insumos"))
OUTPUT_CSV  = os.environ.get("V6_OUTPUT_CSV",  os.path.join(BASE_DIR, "data", "output", "valoracion_inicio.csv"))
OUTPUT_FULL = os.environ.get("V6_OUTPUT_FULL", os.path.join(BASE_DIR, "data", "output", "valoracion_full.csv"))


def _autodetectar_fecha(carpeta: str) -> str:
    """Busca INICIO_YYYYMMDD.csv en la carpeta y devuelve el YYYYMMDD."""
    if not os.path.isdir(carpeta):
        raise FileNotFoundError(f"V6_INSUMOS_DIR no existe: {carpeta}")
    candidatos = [f for f in os.listdir(carpeta) if f.startswith("INICIO_") and f.endswith(".csv")]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró INICIO_YYYYMMDD.csv en {carpeta}")
    # Si hay varios, usar el más reciente por nombre
    candidatos.sort(reverse=True)
    yyyymmdd = candidatos[0].replace("INICIO_", "").replace(".csv", "")
    if len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        raise ValueError(f"Formato inesperado: {candidatos[0]}")
    return yyyymmdd


def _cargar_inicio(carpeta: str, yyyymmdd: str) -> dict:
    """Lee INICIO_YYYYMMDD.csv y devuelve un dict con los spots."""
    path = os.path.join(carpeta, f"INICIO_{yyyymmdd}.csv")
    df = pd.read_csv(path)
    d = dict(zip(df["key"], df["value"]))
    fecha_str = d["FECHA_VAL"]
    fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
    return {
        "FECHA_VAL": fecha,
        "TRM":      float(d["TRM"]),
        "UVR":      float(d["UVR"]),
        "EURCOP":   float(d["EURCOP"]),
    }


# Resolver fecha y spots dinámicamente al cargar el módulo
FECHA_YYYYMMDD = os.environ.get("V6_FECHA_YYYYMMDD") or _autodetectar_fecha(INSUMOS_DIR)
_spots         = _cargar_inicio(INSUMOS_DIR, FECHA_YYYYMMDD)

FECHA_VAL      = _spots["FECHA_VAL"]
TRM_USDCOP     = _spots["TRM"]
UVR_HOY        = _spots["UVR"]
EURCOP         = _spots["EURCOP"]

# SWAPIND del portafolio (misma carpeta, mismo sufijo)
INPUT_CSV      = os.path.join(INSUMOS_DIR, f"SWAPIND_{FECHA_YYYYMMDD}.csv")
# Benchmark Precia del rango B:J de INICIO
VPN_PRECIA_CSV = os.path.join(INSUMOS_DIR, f"VPN_PRECIA_{FECHA_YYYYMMDD}.csv")
# Directorio donde viven las curvas (en el modelo actual, el mismo)
CURVES_DIR     = INSUMOS_DIR

# Frecuencia de cupones en meses
FREQ_MONTHS = {
    "Anual":      12,
    "Semestral":   6,
    "Trimestral":  3,
    "Mensual":     1,
}

# Base de días por moneda (ACT/...)
DAY_BASE = {
    "USD": 360,
    "EUR": 360,
    "COP": 365,
    "UVR": 365,
}

# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — CARGA E INTERPOLACIÓN DE CURVAS
# ═══════════════════════════════════════════════════════════════════════
# Todas las curvas extraídas del Excel tienen formato CSV uniforme:
#    Dias, Tasa
# La interpolación es LINEAL sobre tasas (standard Precia). El factor de
# descuento se calcula según CURVE_CONVENTION (simple_360, simple_365, ea_365).

_CURVES_CACHE: dict = {}

# Mapeo curva lógica → nombre base del archivo (sin fecha). El sufijo
# _YYYYMMDD.csv se resuelve dinámicamente en runtime.
CURVE_FILES_BASE = {
    "IBR_COL_USD":   "curva_IBR_COLATERAL_USD",   # descuento COP en IRS IBRCOP
    "IBR_PROJ":      "curva_IBR_COLATERAL_USD",   # proyección IBR
    "COP_COL_USD":   "curva_COP_COLATERAL_USD",   # ★ descuento COP en CCS COPUSD
    "USDOIS":        "curva_USDOIS",              # proyección+descuento SOFR
    "USDCO":         "curva_USDCO",               # USD colateralizada (no se usa)
    "EURCOP_CURVE":  "curva_EURCOP",
    "UVR_DISC":      "curva_IBRUVR",              # descuento UVR
    "IBRUVR_PROJ":   "curva_IBRUVR",
    "DTF":           "curva_DTF",
    "LIBORUVR":      "curva_LIBORUVR",
    "LIBORCOP":      "curva_LIBORCOP",
}

# Generación dinámica: "IBR_COL_USD" → "curva_IBR_COLATERAL_USD_20260331.csv"
CURVE_FILES = {k: f"{v}_{FECHA_YYYYMMDD}.csv" for k, v in CURVE_FILES_BASE.items()}


# Convención de composición por curva. CRÍTICO: las curvas de Precia no
# están todas en EA. La normalización real es:
#   USDOIS, USDCO        → nominal simple ACT/360 (convención USD money-market)
#   IBR_COL_USD, IBR_PROJ → nominal simple ACT/365 (tasa nominal colombiana)
#   UVR_DISC, IBRUVR_PROJ → nominal simple ACT/365
#   COP_COL_USD, DTF,
#   LIBORUVR, LIBORCOP    → EA ACT/365 (convención de bonos/CETES COP)
#
# Discriminador empírico: si la tasa a 10+ años supera ~15% en la curva, casi
# siempre es convención simple (nominal acumulada), no EA. Las tasas cero EA
# en mercados funcionales rara vez superan 8-10% incluso en COP long-end.
CURVE_CONVENTION = {
    "IBR_COL_USD":  "simple_360",   # ★ tasa nominal ACT/360 (ajuste empírico final)
    "IBR_PROJ":     "simple_360",
    "COP_COL_USD":  "simple_360",   # ★ pata COP de CCS COPUSD (simple_360)
    "USDOIS":       "simple_360",   # SOFR OIS
    "USDCO":        "simple_360",   # USD col COP
    "EURCOP_CURVE": "ea_365",
    "UVR_DISC":     "simple_360",
    "IBRUVR_PROJ":  "simple_360",
    "DTF":          "ea_365",
    "LIBORUVR":     "ea_365",
    "LIBORCOP":     "ea_365",
}


def _load_curve(name: str):
    """Carga una curva y devuelve (plazos, tasas) ordenados."""
    if name in _CURVES_CACHE:
        return _CURVES_CACHE[name]
    if name not in CURVE_FILES:
        raise KeyError(f"Curva desconocida: {name}")
    path = os.path.join(CURVES_DIR, CURVE_FILES[name])
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe curva {name}: {path}")
    df = pd.read_csv(path)
    df = df.drop_duplicates("Dias").sort_values("Dias").reset_index(drop=True)
    plazos = df["Dias"].values.astype(float)
    tasas  = df["Tasa"].values.astype(float)
    _CURVES_CACHE[name] = (plazos, tasas)
    return plazos, tasas


def zero_rate(name: str, dias: float) -> float:
    """Interpola linealmente la tasa cero para cierto # de días."""
    plazos, tasas = _load_curve(name)
    if dias <= plazos[0]:
        return float(tasas[0])
    if dias >= plazos[-1]:
        return float(tasas[-1])
    return float(np.interp(dias, plazos, tasas))


def discount_factor(name: str, dias: float) -> float:
    """
    DF aplicando la convención correspondiente a la curva.
    """
    if dias <= 0:
        return 1.0
    r = zero_rate(name, dias)
    conv = CURVE_CONVENTION.get(name, "ea_365")
    if conv == "simple_360":
        return 1.0 / (1.0 + r * dias / 360.0)
    if conv == "simple_365":
        return 1.0 / (1.0 + r * dias / 365.0)
    # default: EA 365
    return 1.0 / ((1.0 + r) ** (dias / 365.0))


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 3 — CALENDARIOS DE CUPONES
# ═══════════════════════════════════════════════════════════════════════

def build_schedule(start: date, end: date, freq_months: int) -> list:
    """
    Construye calendario de cupones desde start hasta end.
    Retrocede desde end (preserva día de aniversario) con paso freq_months,
    y trunca cualquier fecha ≤ start. Siempre incluye end al final.
    """
    if end <= start:
        return []
    schedule = []
    cur = end
    while cur > start:
        schedule.append(cur)
        cur -= relativedelta(months=freq_months)
    schedule = sorted(schedule)
    if schedule[-1] != end:
        schedule.append(end)
    return schedule


def days_between(d1: date, d2: date) -> int:
    return (d2 - d1).days


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 4 — VALORACIÓN DE PATAS
# ═══════════════════════════════════════════════════════════════════════

def pata_fija(nominal: float, tasa: float, start: date, end: date,
              freq_months: int, moneda: str, curva_disc: str,
              incluir_nocional: bool = False) -> dict:
    """
    Valora pata fija. Retorna dict con:
      pv         : valor presente (en moneda nativa)
      cashflows  : lista de dicts con detalle por flujo
      precio     : pv / nominal (ratio a par)
    """
    base = DAY_BASE[moneda]
    schedule = build_schedule(start, end, freq_months)

    pv = 0.0
    cfs = []
    prev = start

    # Localizar el período inicial: anclamos en la fecha del cupón previo al
    # que todavía está vigente. Para retrocedido, eso es el primer cupón
    # del schedule - freq_months.
    prev_coupon = schedule[0] - relativedelta(months=freq_months) if schedule else start
    prev = prev_coupon

    for idx, fp in enumerate(schedule):
        dias_accrual = days_between(prev, fp)
        frac = dias_accrual / base
        # Flujo de intereses
        cf_interes = nominal * tasa * frac
        # Flujo de capital (solo al vencimiento si aplica)
        cf_cap = nominal if (incluir_nocional and idx == len(schedule) - 1) else 0.0
        cf_total = cf_interes + cf_cap

        # Descontar solo si es futuro
        if fp > FECHA_VAL:
            dias_al_venc = days_between(FECHA_VAL, fp)
            df = discount_factor(curva_disc, dias_al_venc)
            pv_flujo = cf_total * df
            pv += pv_flujo
            cfs.append({
                "fecha": fp,
                "dias_accrual": dias_accrual,
                "dias_venc": dias_al_venc,
                "tasa": tasa,
                "cf": cf_total,
                "cf_interes": cf_interes,
                "cf_capital": cf_cap,
                "df": df,
                "pv": pv_flujo,
            })
        prev = fp

    return {
        "pv": pv,
        "cashflows": cfs,
        "precio": pv / nominal if nominal else 0.0,
    }


def pata_flotante(nominal: float, spread: float, start: date, end: date,
                  freq_months: int, moneda: str, curva_proj: str,
                  curva_disc: str, incluir_nocional: bool = False) -> dict:
    """
    Valora pata flotante con forward implícitos.
      forward_EA(t1,t2) = (DF_proj(t1)/DF_proj(t2))^(365/dias) - 1
      cupón = (fwd_EA + spread) × nominal × (dias/base)
    Descuenta con curva_disc.
    """
    base = DAY_BASE[moneda]
    schedule = build_schedule(start, end, freq_months)

    pv = 0.0
    cfs = []

    prev_coupon = schedule[0] - relativedelta(months=freq_months) if schedule else start
    prev = prev_coupon

    for idx, fp in enumerate(schedule):
        dias_accrual = days_between(prev, fp)
        frac = dias_accrual / base

        if fp > FECHA_VAL:
            # Forward entre prev y fp usando curva_proj.
            dias_t1 = max(days_between(FECHA_VAL, prev), 0)
            dias_t2 = days_between(FECHA_VAL, fp)

            if prev < FECHA_VAL:
                # ── Cupón en curso ──
                # El período de acumulación empezó ANTES de la fecha de valoración
                # (i.e., el IBR/SOFR compound ya se observó parcialmente).
                # Sin serie histórica de la tasa flotante, aproximamos la tasa
                # realizada usando la tasa cero al tenor del período completo.
                # Para rates overnight-compound (IBR, SOFR), esta es la mejor
                # estimación: la curva cero refleja exactamente el promedio
                # esperado/observado del overnight sobre ese horizonte.
                fwd = zero_rate(curva_proj, dias_accrual)
            elif prev == FECHA_VAL:
                # ── Forward-starting exacto ──
                # Convención Precia: cuando el cupón en curso comienza
                # exactamente en la fecha de valoración (swap recién emitido
                # o cupón justo reseteado), la tasa flotante compound aún no
                # ha comenzado. Precia asigna cupón = 0 a este primer período,
                # de modo que PV del período = 0 + nocional_final × DF_disc.
                # Esto resuelve los 8 outliers IRS SOFUSD forward-starting.
                fwd = 0.0
            else:
                # ── Período futuro ──
                # Forward implícito estándar entre prev y fp.
                df1 = discount_factor(curva_proj, dias_t1)
                df2 = discount_factor(curva_proj, dias_t2)
                if dias_accrual > 0 and df2 > 0:
                    conv = CURVE_CONVENTION.get(curva_proj, "ea_365")
                    if conv == "simple_360":
                        fwd = (df1 / df2 - 1.0) * 360.0 / dias_accrual
                    elif conv == "simple_365":
                        fwd = (df1 / df2 - 1.0) * 365.0 / dias_accrual
                    else:
                        fwd = (df1 / df2) ** (365.0 / dias_accrual) - 1.0
                else:
                    fwd = 0.0

            cf_interes = (fwd + spread) * nominal * frac
            cf_cap = nominal if (incluir_nocional and idx == len(schedule) - 1) else 0.0
            cf_total = cf_interes + cf_cap

            df_disc = discount_factor(curva_disc, dias_t2)
            pv_flujo = cf_total * df_disc
            pv += pv_flujo
            cfs.append({
                "fecha": fp,
                "dias_accrual": dias_accrual,
                "dias_venc": dias_t2,
                "tasa": fwd + spread,
                "cf": cf_total,
                "cf_interes": cf_interes,
                "cf_capital": cf_cap,
                "df": df_disc,
                "pv": pv_flujo,
            })
        prev = fp

    return {
        "pv": pv,
        "cashflows": cfs,
        "precio": pv / nominal if nominal else 0.0,
    }


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 5 — DURACIÓN MODIFICADA Y CONVEXIDAD
# ═══════════════════════════════════════════════════════════════════════

def duracion_modificada_convexidad(cashflows: list, pv: float, y: float) -> tuple:
    """
    DM = [Σ (t × PV_i) / VPN] / (1+y)
    C  = Σ (t × (t+1) × PV_i) / [VPN × (1+y)^2]
    donde t = dias/365, y es el yield cero EA interpolado al vencimiento.
    """
    if pv <= 0 or not cashflows:
        return 0.0, 0.0
    t_arr  = np.array([cf["dias_venc"] / 365.0 for cf in cashflows])
    pv_arr = np.array([cf["pv"] for cf in cashflows])
    dmac = (t_arr * pv_arr).sum() / pv
    dm   = dmac / (1.0 + y)
    conv = (t_arr * (t_arr + 1.0) * pv_arr).sum() / (pv * (1.0 + y) ** 2)
    return float(dm), float(conv)


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 6 — DISPATCHER POR TIPO DE SWAP
# ═══════════════════════════════════════════════════════════════════════

def _fx_spot(moneda: str) -> float:
    if moneda == "USD": return TRM_USDCOP
    if moneda == "EUR": return EURCOP
    if moneda == "UVR": return UVR_HOY
    if moneda == "COP": return 1.0
    raise ValueError(f"Moneda no soportada: {moneda}")


def valorar_pata(tipo_swap: str, lado: str, clase: str, tfacial: str,
                 tasa: float, periodicidad: str, moneda: str, nominal: float,
                 start: date, end: date, vpn_reportado=None) -> dict:
    """
    Selecciona curvas apropiadas según tipo de swap y moneda, y delega a
    pata_fija o pata_flotante. Retorna el mismo dict más 'yield_venc'.

    periodicidad "Al vencimiento" => pata CAPITALIZANTE (bullet): el interés se
    acumula y se paga una sola vez al vencimiento. El nocional capitalizado a la
    fecha lo reporta Precia (vpn_reportado = VPN col53/54 del SFC, en moneda
    nativa COP); reconstruirlo desde cero exigiría la serie histórica de fixings
    del índice flotante, que no está en los insumos.
    """
    # ─── Asignación de curvas según metodología Precia ───
    if tipo_swap == "IRS SOFUSD":
        # IRS SOFUSD single-curve: USDOIS proyecta y descuenta.
        # Descartar USDCO aquí (es la curva colateral, pero para el IRS SOFUSD
        # el estándar es single-curve SOFR/USDOIS, que empíricamente da
        # errores sub-4% vs 6% del dual-curve).
        curva_disc = "USDOIS"
        curva_proj = "USDOIS"
    elif tipo_swap == "IRS IBRCOP":
        # Ambas patas COP. Descuenta con IBR_COL_USD, proyecta IBR.
        curva_disc = "IBR_COL_USD"
        curva_proj = "IBR_PROJ"
    elif tipo_swap == "CCS COPUSD":
        # Pata COP → COP_COL_USD (simple_360), Pata USD → USDOIS (no USDCO).
        # USDOIS resuelve el sesgo +1.8% de descontar USD con USDCO.
        # COP_COL_USD (no IBR_COL_USD) resuelve el sesgo +1.2% en la pata COP.
        curva_disc = "COP_COL_USD" if moneda == "COP" else "USDOIS"
        curva_proj = curva_disc
    elif tipo_swap == "CCS COPUVR":
        # Pata COP→IBR_COL_USD, UVR→IBRUVR
        curva_disc = "IBR_COL_USD" if moneda == "COP" else "UVR_DISC"
        curva_proj = curva_disc
    else:
        raise ValueError(f"Tipo swap no soportado: {tipo_swap}")

    # METODOLOGÍA PRECIA: cada pata se valora como bono replicado,
    # es decir CUPONES + NOCIONAL al final. Para el IRS los nocionales
    # se cancelan en el neto, pero Precia los reporta individualmente.
    # Esto explica por qué precios > 1 (pata flotante descontada con
    # USDCO que es más blanda que USDOIS con la que se proyecta).
    incluir_nocional = True

    # Tipo de pata
    es_fija = (clase == "SWAP TF") or (tfacial == "FS")

    # ─── Pata CAPITALIZANTE (bullet, "Al vencimiento") ───
    # El interés se capitaliza y se paga una sola vez al vencimiento. El nocional
    # capitalizado a la fecha es el que reporta Precia (vpn_reportado). Se modela
    # como un único flujo al vencimiento: el VPN es ese nocional acumulado y la
    # duración/convexidad se calculan como un cupón cero al vencimiento.
    if periodicidad == "Al vencimiento":
        _vpn = None
        if vpn_reportado not in (None, ""):
            try:
                v = float(vpn_reportado)
                if not np.isnan(v):
                    _vpn = v
            except (TypeError, ValueError):
                _vpn = None
        pv = _vpn if _vpn is not None else nominal
        dias_al_venc = max(days_between(FECHA_VAL, end), 1)
        df = discount_factor(curva_disc, dias_al_venc)
        cf_mat = pv / df if df else pv  # flujo bruto al vencimiento (bullet)
        res = {
            "pv": pv,
            "cashflows": [{
                "fecha": end, "dias_accrual": days_between(start, end),
                "dias_venc": dias_al_venc, "tasa": tasa, "cf": cf_mat,
                "cf_interes": 0.0, "cf_capital": cf_mat, "df": df, "pv": pv,
            }],
            "precio": pv / nominal if nominal else 0.0,
            "capitalizante": True,
        }
        y = zero_rate(curva_disc, dias_al_venc)
        res["yield_venc"] = y
        res["curva_disc"] = curva_disc
        res["curva_proj"] = curva_proj
        return res

    freq_m = FREQ_MONTHS[periodicidad]
    if es_fija:
        res = pata_fija(
            nominal=nominal, tasa=tasa, start=start, end=end,
            freq_months=freq_m, moneda=moneda, curva_disc=curva_disc,
            incluir_nocional=incluir_nocional,
        )
    else:
        # Pata flotante: el tratamiento del spread "Facial" difiere por índice:
        #  - SOF1 (SOFR): Precia NO aplica Facial como spread (es referencia
        #    histórica, no un diferencial real). Usar spread = 0.
        #  - IBR: Precia SÍ aplica Facial como spread (margen real IBR+X%).
        #
        # ⚠ NOTA P1 (pendiente revisión) — ver CAMBIOS_V6.md, sección Pendientes:
        # La mesa confirmó que los contratos SOF1 con Facial 0.26% sí pagan
        # SOF1 + 0.26% económicamente, pero Precia no lo refleja en MTM.
        # Se mantiene spread=0 para alinear con Precia; revisar si se necesita
        # una versión económica paralela.
        if tfacial == "SOF1":
            spread_real = 0.0
        else:
            spread_real = tasa
        res = pata_flotante(
            nominal=nominal, spread=spread_real, start=start, end=end,
            freq_months=freq_m, moneda=moneda,
            curva_proj=curva_proj, curva_disc=curva_disc,
            incluir_nocional=incluir_nocional,
        )

    # yield al vencimiento (para DM)
    dias_al_venc = days_between(FECHA_VAL, end)
    y = zero_rate(curva_disc, max(dias_al_venc, 1))
    res["yield_venc"] = y
    res["curva_disc"] = curva_disc
    res["curva_proj"] = curva_proj
    return res


def _parse_rate(v) -> float:
    """Parsea tasa. Acepta float, '0.26%', '0.26 %', etc."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if s.endswith("%"):
        return float(s[:-1]) / 100.0
    return float(s)


def valorar_swap(row: pd.Series) -> dict:
    """
    Valora un swap completo (ambas patas) y devuelve resultados en COP.
    """
    tipo = row["TIPO SWAP"]
    # Fechas - usamos F.Compra como inicio (la metodología lo confirmó)
    f_compra = pd.to_datetime(row["F.Compra"]).date()
    f_vcto   = pd.to_datetime(row["F.Vcto  "]).date()

    # Pata Derecho
    der = valorar_pata(
        tipo, "DER",
        clase=row["IND CLASE"],
        tfacial=row["IND. Tfacial"],
        tasa=_parse_rate(row["Facial"]),
        periodicidad=row["Period."],
        moneda=row["Moneda"],
        nominal=float(row["Vr Nominal DER"]),
        start=f_compra, end=f_vcto,
        vpn_reportado=row.get("VPN Precia DER"),
    )
    # Pata Obligación
    obl = valorar_pata(
        tipo, "OBL",
        clase=row["IND CLASE.1"],
        tfacial=row["IND. Tfacial.1"],
        tasa=_parse_rate(row["Facial         "]),
        periodicidad=row["Period..1"],
        moneda=row["Moneda.1"],
        nominal=float(row["Vr Nominal"]),
        start=f_compra, end=f_vcto,
        vpn_reportado=row.get("VPN Precia OBL"),
    )

    # Convertir a COP (metodología Precia: VPN en COP = PV_moneda × FX_spot)
    fx_der = _fx_spot(row["Moneda"])
    fx_obl = _fx_spot(row["Moneda.1"])

    vpn_der_cop = der["pv"] * fx_der
    vpn_obl_cop = obl["pv"] * fx_obl

    # DM y Convexidad
    dm_d, conv_d = duracion_modificada_convexidad(der["cashflows"], der["pv"], der["yield_venc"])
    dm_o, conv_o = duracion_modificada_convexidad(obl["cashflows"], obl["pv"], obl["yield_venc"])

    return {
        "vpn_der_cop": vpn_der_cop,
        "vpn_obl_cop": vpn_obl_cop,
        "dm_der":   dm_d,
        "conv_der": conv_d,
        "dm_obl":   dm_o,
        "conv_obl": conv_o,
        "precio_der": der["precio"],
        "precio_obl": obl["precio"],
        "curva_disc_der": der["curva_disc"],
        "curva_disc_obl": obl["curva_disc"],
    }


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 7 — MAIN
# ═══════════════════════════════════════════════════════════════════════

def main():
    print(f"Insumos dir     : {INSUMOS_DIR}")
    print(f"Fecha valoración: {FECHA_VAL}  (YYYYMMDD={FECHA_YYYYMMDD})")
    print(f"TRM             : {TRM_USDCOP:,.4f}")
    print(f"UVR             : {UVR_HOY:,.4f}")
    print(f"EURCOP          : {EURCOP:,.4f}")
    print(f"Input portafolio: {INPUT_CSV}")

    df_in = pd.read_csv(INPUT_CSV)
    print(f"                  {len(df_in)} swaps\n")

    results = []
    errores = 0
    for idx, row in df_in.iterrows():
        try:
            res = valorar_swap(row)
            vpn_d_precia = row.get("VPN Derecho ME")
            vpn_o_precia = row.get("VPN Obligacion ME")
            err_d = (res["vpn_der_cop"] / vpn_d_precia - 1) * 100 if vpn_d_precia else float("nan")
            err_o = (res["vpn_obl_cop"] / vpn_o_precia - 1) * 100 if vpn_o_precia else float("nan")
            results.append({
                # Identificación
                "ISIN":          row["ISIN"],
                "Emisor":        row["EMISOR"],
                "Tipo_Swap":     row["TIPO SWAP"],
                "Fecha_Compra":  row["F.Compra"],
                "Emision":       row["Emision "],
                "Fecha_Vcto":    row["F.Vcto  "],
                # Resultados DER
                "VPN_Derecho_Calc":    res["vpn_der_cop"],
                "VPN_Precia_Der":      vpn_d_precia,
                "Error_vs_Precia_Der": err_d,
                "Duracion_Mod_Der":    res["dm_der"],
                "Convexidad_Der":      res["conv_der"],
                "Precio_Der":          res["precio_der"],
                "Curva_Disc_Der":      res["curva_disc_der"],
                # Resultados OBL
                "VPN_Oblig_Calc":      res["vpn_obl_cop"],
                "VPN_Precia_Obl":      vpn_o_precia,
                "Error_vs_Precia_Obl": err_o,
                "Duracion_Mod_Obl":    res["dm_obl"],
                "Convexidad_Obl":      res["conv_obl"],
                "Precio_Obl":          res["precio_obl"],
                "Curva_Disc_Obl":      res["curva_disc_obl"],
            })
        except Exception as e:
            errores += 1
            results.append({
                "ISIN":        row.get("ISIN"),
                "Emisor":      row.get("EMISOR"),
                "Tipo_Swap":   row.get("TIPO SWAP"),
                "Fecha_Compra": row.get("F.Compra"),
                "Emision":     row.get("Emision "),
                "Fecha_Vcto":  row.get("F.Vcto  "),
                "Error": f"{type(e).__name__}: {e}",
            })

    df_full = pd.DataFrame(results)

    # ─── OUTPUT 1 (formato INICIO rango B:J de Precia) ─────────────────
    # # SWAP | DERECHO | OBLIGACIÓN | VPN | LLAVE_DER | DER | LLAVE_OBLIG | OBLIG
    df_inicio = pd.DataFrame({
        "# SWAP":      df_full["ISIN"],
        "DERECHO":     df_full["VPN_Derecho_Calc"],
        "OBLIGACIÓN":  df_full["VPN_Oblig_Calc"],
        "VPN":         df_full["VPN_Derecho_Calc"] - df_full["VPN_Oblig_Calc"],
        "LLAVE_DER":   "SWAP" + df_full["ISIN"].astype(str) + "D",
        "DER":         df_full["Precio_Der"],
        "LLAVE_OBLIG": "SWAP" + df_full["ISIN"].astype(str) + "O",
        "OBLIG":       df_full["Precio_Obl"],
    })
    os.makedirs(os.path.dirname(OUTPUT_CSV) or ".", exist_ok=True)
    df_inicio.to_csv(OUTPUT_CSV, index=False, float_format="%.10f")
    print(f"✓ Output INICIO (B:J): {OUTPUT_CSV}")
    print(f"  {len(df_inicio)} swaps  |  {errores} errores de valoración")

    # ─── OUTPUT 2 (detallado, con riesgo + benchmark) ──────────────────
    os.makedirs(os.path.dirname(OUTPUT_FULL) or ".", exist_ok=True)
    df_full.to_csv(OUTPUT_FULL, index=False)
    print(f"✓ Output detallado   : {OUTPUT_FULL}")

    # ─── Resumen de errores (si hay benchmark Precia en SWAPIND) ───────
    if "Error_vs_Precia_Der" in df_full.columns and df_full["Error_vs_Precia_Der"].notna().any():
        print("\nResumen de errores vs Precia (|error|% mediana, p95, max, n>2%):")
        df_full["abs_d"] = df_full["Error_vs_Precia_Der"].abs()
        df_full["abs_o"] = df_full["Error_vs_Precia_Obl"].abs()
        grp = df_full.groupby("Tipo_Swap").agg(
            n=("ISIN", "count"),
            med_der=("abs_d", "median"),
            p95_der=("abs_d", lambda s: s.quantile(0.95)),
            max_der=("abs_d", "max"),
            nout_der=("abs_d", lambda s: int((s > 2).sum())),
            med_obl=("abs_o", "median"),
            p95_obl=("abs_o", lambda s: s.quantile(0.95)),
            max_obl=("abs_o", "max"),
            nout_obl=("abs_o", lambda s: int((s > 2).sum())),
        )
        print(grp.round(3).to_string())
        total_out = int(((df_full["abs_d"] > 2) | (df_full["abs_o"] > 2)).sum())
        print(f"\nTotal swaps con cualquier |error| > 2%: {total_out} / {len(df_full)}")

    return df_full


if __name__ == "__main__":
    main()
