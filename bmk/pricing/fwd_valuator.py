"""
Valorador de forwards de la industria — replica la hoja `Fwd Industria` del
Benchmark (Detallado), forward por forward.

Modelo (decodificado y validado al decimal contra la propia hoja, corte junio):
  Para cada forward, con t = FECHA CUMPLIMIENTO - FECHA VALORACION (dias), cada
  pata se descuenta con la curva de SU moneda (paridad cubierta bi-moneda):
    DF_strike = 1 / (1 + curva_disc_strike[t]/100 * t/360)
    DF_spot   = 1 / (1 + curva_disc_spot[t]/100   * t/360)
    puntos    = FWP<par>[t]  (solo BRL/MXN; 0 en el resto)
    Tasa VPN     = strike          * DF_strike   -> pata "pactada" traida a hoy
    Tasa Forward = (spot + puntos) * DF_spot     -> pata "mercado" traida a hoy
  USDCOP: DF_strike<-FWTCOP(IBR), DF_spot<-FWTUSD(USDCO), sin puntos.
  EURUSD: DF_strike<-LIBBTS(USDOIS), DF_spot<-FWTEUR(EUROIS).

  Valor en COP de cada pata (segun como cotiza el par):
    - COP-quoted (USDCOP): la tasa YA es COP por unidad ->
        valor = nominal * tasa
    - Cross USD/XXX (USDBRL, USDMXN, USDJPY, USDCLP, USDCAD, USDCHF): la tasa es
      XXX por USD -> COP via TRM:  valor = nominal * TRM / tasa
    - Cross XXX/USD (EURUSD, AUDUSD, GBPUSD): la tasa es USD por XXX ->
        valor = nominal * TRM * tasa

  P&G del forward:
    VENTA : der = nominal*TasaVPN(pata pactada), obl = pata mercado
    COMPRA: der = pata mercado, obl = pata pactada
    P&G = VALOR DER - VALOR OBLI

Validacion: USDCOP casa al peso (P&G exacto en 1.367/1.367 filas del archivo).
Los pares cruzados usan la misma mecanica; su exactitud depende de la curva de
puntos por par disponible en `Curvas` (ver validar_contra_sabana).
"""
from __future__ import annotations

import bisect
import datetime as _dt
import math
from pathlib import Path

import pandas as pd

_EPOCH = _dt.date(1899, 12, 30)  # dia base del serial de Excel


def _mas_un_mes(serial: int) -> int:
    """serial de Excel -> serial del mismo dia del mes siguiente (roll mensual del
    macro para forwards ya vencidos a la fecha de valoracion)."""
    d = _EPOCH + _dt.timedelta(days=int(serial))
    y, m = (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
    try:
        nd = d.replace(year=y, month=m)
    except ValueError:  # dia inexistente (p.ej. 31) -> ultimo dia del mes
        import calendar
        nd = d.replace(year=y, month=m, day=calendar.monthrange(y, m)[1])
    return (nd - _EPOCH).days

# par -> (tipo_cotizacion, curva_descuento_strike, curva_descuento_spot, curva_puntos)
#   "COP" : tasa en COP por unidad extranjera (USDCOP)
#   "USDX": USD/XXX -> COP = TRM*tasa (EURUSD, AUDUSD, GBPUSD)
#   "XUSD": XXX/USD -> COP = TRM/tasa (USDBRL, USDMXN, USDJPY, USDCLP, USDCAD, USDCHF)
#
# Metodologia (decodificada al decimal contra la hoja Fwd Industria del macro,
# corte junio valorado 10-jul):
#     Tasa VPN     = strike           x DF(curva_descuento_strike)
#     Tasa Forward = (spot + puntos)  x DF(curva_descuento_spot)
# Cada pata se descuenta con la curva de SU moneda (paridad cubierta bi-moneda):
#   - USDCOP: pata strike (COP) <- FWTCOP(=IBR); pata spot (USD) <- FWTUSD(=USDCO).
#     Sin puntos: el forward outright sale de DF_usd/DF_cop (paridad cubierta).
#   - EURUSD: pata strike (USD) <- LIBBTS(=USDOIS); pata spot (EUR) <- FWTEUR(=EUROIS).
#   - AUDUSD/GBPUSD: ambas patas USD (no hay OIS de AUD/GBP en el insumo) -> LIBBTS.
#   - USDBRL/USDMXN: descuento USD (LIBBTS) + puntos de mercado (FWPBRL/FWPMXN).
#   - USDJPY/CLP/CAD/CHF: descuento USD (LIBBTS), sin puntos (premio pequeno).
PARES = {
    "USDCOP": ("COP", "FWTCOP", "FWTUSD", None),
    "USDBRL": ("XUSD", "LIBBTS", "LIBBTS", "FWPBRL"),
    "USDMXN": ("XUSD", "LIBBTS", "LIBBTS", "FWPMXN"),
    "USDJPY": ("XUSD", "LIBBTS", "LIBBTS", None),
    "USDCLP": ("XUSD", "LIBBTS", "LIBBTS", None),
    "USDCAD": ("XUSD", "LIBBTS", "LIBBTS", None),
    "USDCHF": ("XUSD", "LIBBTS", "LIBBTS", None),
    "EURUSD": ("USDX", "LIBBTS", "FWTEUR", None),
    "AUDUSD": ("USDX", "LIBBTS", "LIBBTS", None),
    "GBPUSD": ("USDX", "LIBBTS", "LIBBTS", None),
}
# par -> moneda extranjera cuyo spot (de paridades.csv) usar como spot del par.
SPOT_MONEDA = {"USDCOP": "USD", "EURUSD": "EUR", "AUDUSD": "AUD", "GBPUSD": "GBP",
               "USDBRL": "BRL", "USDMXN": "MXN", "USDJPY": "JPY", "USDCLP": "CLP",
               "USDCAD": "CAD", "USDCHF": "CHF"}


class Curvas:
    """Curvas forward cargadas de CSV (insumos/fwd_curves)."""

    def __init__(self, dir_curvas: str | Path = "insumos/fwd_curves"):
        self.dir = Path(dir_curvas)
        self._cache: dict[str, tuple[list, list]] = {}
        par = pd.read_csv(self.dir / "paridades.csv")
        self.spot = dict(zip(par["par"], par["spot"]))
        self.trm = self.spot.get("USD")

    def _curva(self, nombre: str):
        if nombre not in self._cache:
            df = pd.read_csv(self.dir / f"{nombre}.csv").sort_values("plazo_dias")
            self._cache[nombre] = (df["plazo_dias"].tolist(), df["valor"].tolist())
        return self._cache[nombre]

    def interp(self, nombre: str, t: float) -> float:
        xs, ys = self._curva(nombre)
        if t <= xs[0]:
            return ys[0]
        if t >= xs[-1]:
            return ys[-1]
        i = bisect.bisect_left(xs, t)
        if xs[i] == t:
            return ys[i]
        x0, x1, y0, y1 = xs[i - 1], xs[i], ys[i - 1], ys[i]
        return y0 + (y1 - y0) * (t - x0) / (x1 - x0)


def _valor_pata(tipo: str, nominal: float, tasa: float, trm: float) -> float:
    if tipo == "COP":
        return nominal * tasa
    if tipo == "USDX":
        return nominal * trm * tasa
    # XUSD
    return nominal * trm / tasa if tasa else 0.0


def valorar(fwd: pd.DataFrame, curvas: Curvas, fecha_valoracion: int) -> pd.DataFrame:
    """Valora cada forward. `fwd` con columnas: paridad, operacion (COMPRA/VENTA),
    nominal, strike, fecha_cumplimiento (serial Excel), [spot opcional].
    Devuelve el df con tasa_vpn, tasa_forward, valor_cop_der, valor_cop_obli, pyg.
    """
    out = fwd.copy().reset_index(drop=True)
    tv, tf, der, obl, pyg, plazo, spots = [], [], [], [], [], [], []
    for _, r in out.iterrows():
        par = str(r["paridad"]).strip().upper()
        cfg = PARES.get(par)
        nom = float(r["nominal"] or 0)
        K = float(r["strike"] or 0)
        cumpl = r.get("fecha_cumplimiento")
        venc = r.get("fecha_vencimiento")
        if cfg is None or not cumpl or pd.isna(cumpl):
            tv.append(None); tf.append(None); der.append(None); obl.append(None); pyg.append(None); plazo.append(None)
            continue
        # Roll mensual del macro: un forward vencido (o que vence justo el dia de
        # valoracion) se renueva a fecha_valoracion + 1 mes (cobertura que rueda mes
        # a mes). Verificado en la hoja Fwd Industria del corte junio: los forwards
        # con fecha_vencimiento <= 10-jul ruedan todos al 10-ago (=10-jul+1M).
        if venc is not None and pd.notna(venc) and int(venc) <= int(fecha_valoracion):
            cumpl = _mas_un_mes(int(fecha_valoracion))
        tipo, curva_disc_strike, curva_disc_spot, curva_pts = cfg
        t = int(cumpl) - int(fecha_valoracion)
        if t <= 0:  # forward ya cumplido/liquidado: sin valor de mercado.
            tv.append(0.0); tf.append(0.0); der.append(0.0); obl.append(0.0)
            pyg.append(0.0); plazo.append(t); spots.append(None)
            continue
        # Metodologia del macro (paridad cubierta bi-moneda): cada pata se descuenta
        # con la curva de SU moneda.
        #     Tasa VPN     = strike           x DF(curva_disc_strike)
        #     Tasa Forward = (spot + puntos)  x DF(curva_disc_spot)
        #     P&G          = (VPN - Forward) x nominal   (VENTA; se invierte en COMPRA)
        # Verificado contract-level contra la hoja 'Fwd Industria' del corte junio.
        r_strike = curvas.interp(curva_disc_strike, t) / 100.0
        df_strike = 1.0 / (1.0 + r_strike * t / 360.0)
        r_spot = curvas.interp(curva_disc_spot, t) / 100.0
        df_spot = 1.0 / (1.0 + r_spot * t / 360.0)
        spot = r["spot"] if "spot" in out.columns and pd.notna(r.get("spot")) else curvas.spot.get(SPOT_MONEDA.get(par, ""), 0.0)
        # Normalizar el strike a la convencion de cotizacion del par: si viene
        # invertido (p. ej. USDJPY 0.0064 en vez de 156), 1/strike queda mas cerca
        # del spot -> invertir. No afecta strikes ya correctos (SABANA validada).
        if K > 0 and spot > 0 and abs(math.log(K / spot)) > abs(math.log((1.0 / K) / spot)):
            K = 1.0 / K
        pts = curvas.interp(curva_pts, t) if curva_pts else 0.0
        tasa_vpn = K * df_strike
        tasa_fwd = (spot + pts) * df_spot
        v_pact = _valor_pata(tipo, nom, tasa_vpn, curvas.trm)
        v_mkt = _valor_pata(tipo, nom, tasa_fwd, curvas.trm)
        op = str(r["operacion"]).strip().upper()
        if op.startswith("V"):  # VENTA / SELL
            d, o = v_pact, v_mkt
        else:                    # COMPRA / BUY
            d, o = v_mkt, v_pact
        tv.append(tasa_vpn); tf.append(tasa_fwd)
        der.append(d); obl.append(o); pyg.append(d - o); plazo.append(t); spots.append(spot)
    out["plazo_dias"] = plazo
    out["spot"] = spots
    out["tasa_vpn"] = tv
    out["tasa_forward"] = tf
    out["valor_cop_der"] = der
    out["valor_cop_obli"] = obl
    out["pyg"] = pyg
    return out
