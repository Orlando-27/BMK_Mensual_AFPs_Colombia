"""
Valorador de forwards de la industria — replica la hoja `Fwd Industria` del
Benchmark (Detallado), forward por forward.

Modelo (decodificado y validado contra la propia hoja, corte del archivo):
  Para cada forward, con t = FECHA CUMPLIMIENTO - FECHA VALORACION (dias):
    r_local = FWTCOP[t] (%)          -> tasa de descuento COP
    DF      = 1 / (1 + r_local/100 * t/360)
    puntos  = FWP<par>[t]            -> puntos forward del par
    mkt     = spot_par + puntos      -> forward de mercado del par
    Tasa VPN     = strike * DF       -> pata "pactada" traida a hoy
    Tasa Forward = mkt    * DF       -> pata "mercado" traida a hoy

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
from pathlib import Path

import pandas as pd

# par -> (tipo_cotizacion, curva_de_puntos)
#   "COP" : tasa en COP por unidad extranjera (USDCOP)
#   "USDX": USD/XXX -> COP = TRM*tasa (EURUSD, AUDUSD, GBPUSD)
#   "XUSD": XXX/USD -> COP = TRM/tasa (USDBRL, USDMXN, USDJPY, USDCLP, USDCAD, USDCHF)
PARES = {
    "USDCOP": ("COP", "FWPCOP"),
    "EURUSD": ("USDX", "FWPEUR"),
    "AUDUSD": ("USDX", None),   # sin puntos (el forward = spot descontado)
    "GBPUSD": ("USDX", None),   # sin puntos (el forward = spot descontado)
    "USDBRL": ("XUSD", "FWPBRL"),
    "USDMXN": ("XUSD", "FWPMXN"),
    "USDJPY": ("XUSD", "FWPJPY"),
    "USDCLP": ("XUSD", "FWPCLP"),
    "USDCAD": ("XUSD", "FWPCAD"),
    "USDCHF": ("XUSD", "FWPCHF"),
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
    tv, tf, der, obl, pyg, plazo = [], [], [], [], [], []
    for _, r in out.iterrows():
        par = str(r["paridad"]).strip().upper()
        cfg = PARES.get(par)
        nom = float(r["nominal"] or 0)
        K = float(r["strike"] or 0)
        cumpl = r.get("fecha_cumplimiento")
        if cfg is None or not cumpl or pd.isna(cumpl):
            tv.append(None); tf.append(None); der.append(None); obl.append(None); pyg.append(None); plazo.append(None)
            continue
        tipo, curva_pts = cfg
        t = int(cumpl) - int(fecha_valoracion)
        if t <= 0:  # forward ya cumplido/liquidado: sin valor de mercado.
            tv.append(0.0); tf.append(0.0); der.append(0.0); obl.append(0.0); pyg.append(0.0); plazo.append(t)
            continue
        # COP-quoted (USDCOP) descuenta con la tasa COP (FWTCOP); los cruzados
        # USD/XXX y XXX/USD descuentan con la tasa foranea USD (LIBBTS).
        curva_desc = "FWTCOP" if tipo == "COP" else "LIBBTS"
        rl = curvas.interp(curva_desc, t) / 100.0
        df = 1.0 / (1.0 + rl * t / 360.0)
        spot = r["spot"] if "spot" in out.columns and pd.notna(r.get("spot")) else curvas.spot.get(SPOT_MONEDA.get(par, ""), 0.0)
        pts = curvas.interp(curva_pts, t) if curva_pts else 0.0
        mkt = spot + pts
        tasa_vpn = K * df
        tasa_fwd = mkt * df
        v_pact = _valor_pata(tipo, nom, tasa_vpn, curvas.trm)
        v_mkt = _valor_pata(tipo, nom, tasa_fwd, curvas.trm)
        op = str(r["operacion"]).strip().upper()
        if op.startswith("V"):  # VENTA / SELL
            d, o = v_pact, v_mkt
        else:                    # COMPRA / BUY
            d, o = v_mkt, v_pact
        tv.append(tasa_vpn); tf.append(tasa_fwd)
        der.append(d); obl.append(o); pyg.append(d - o); plazo.append(t)
    out["plazo_dias"] = plazo
    out["tasa_vpn"] = tv
    out["tasa_forward"] = tf
    out["valor_cop_der"] = der
    out["valor_cop_obli"] = obl
    out["pyg"] = pyg
    return out
