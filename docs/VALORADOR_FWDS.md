# Valorador de Forwards de la Industria

Replica la hoja `Fwd Industria` del Benchmark (Detallado): valora cada forward de
divisas de la industria, forward por forward, y produce el output
**"Valoracion Fwds Industria"**.

## Modelo (decodificado y validado contra la propia hoja)

Para cada forward, con `t = FECHA CUMPLIMIENTO - FECHA VALORACION` (dias):

```
DF   = 1 / (1 + tasa_descuento(t)/100 * t/360)
mkt  = spot_par + puntos_fwd(t)          # forward de mercado del par
Tasa VPN     = strike * DF               # pata pactada, traida a hoy
Tasa Forward = mkt    * DF               # pata mercado, traida a hoy
```

Tasa de descuento segun el par:
- **USDCOP** (cotiza en COP): descuenta con **FWTCOP** (tasa forward COP, ~11-13%).
- **Cruzados** (USD/XXX y XXX/USD): descuentan con **LIBBTS** (tasa foranea USD, ~3.6%).

Valor en COP de cada pata (segun como cotiza el par):
- **USDCOP** (COP por USD): `valor = nominal * tasa`
- **USD/XXX** (XXX por USD: BRL, MXN, JPY, CLP, CAD, CHF): `valor = nominal * TRM / tasa`
- **XXX/USD** (USD por XXX: EUR, AUD, GBP): `valor = nominal * TRM * tasa`

P&G del forward:
- VENTA : `der = pata pactada`, `obl = pata mercado`
- COMPRA: `der = pata mercado`, `obl = pata pactada`
- `P&G = VALOR DER - VALOR OBLI`

Forwards ya cumplidos (`t <= 0`): sin valor de mercado (0).

## Insumos (hoja `Curvas` del Detallado)

Extraidos a `insumos/fwd_curves/` (CSV swappables) por `bmk.pricing.extraer_curvas_fwd`:

| Insumo | Contenido |
|---|---|
| `paridades.csv` | spot por moneda (USD=TRM, EUR, BRL, ...) |
| `FWTCOP.csv` | tasa de descuento COP por plazo |
| `LIBBTS.csv` | tasa de descuento foranea (USD) por plazo |
| `FWPCOP.csv` | puntos forward USDCOP por plazo |
| `FWPBRL/MXN/JPY/CLP/CHF/CAD.csv`, `Fwd_GBPUSD_Diaria.csv`, `FWPEUR.csv` | puntos forward por par cruzado |

> Estas curvas son un **placeholder** (corte del .xlsb, ~2026-07-09). Para cada
> corte real se reemplazan por las curvas del dia (mismo formato de CSV).

## Validacion (contra la hoja Fwd Industria del propio archivo)

2.343 forwards, 9 pares. Valorando con las curvas del archivo y su fecha de
valoracion, el P&G reproduce el de la hoja **al peso**:

| Metrica | Resultado |
|---|---|
| P&G total (excl. 5 anomalas) | **-0,00045%** vs hoja |
| Filas dentro de 1 COP | 2.284 / 2.338 |
| USDCOP (1.372 filas) | P&G exacto 1.372/1.372 |
| Cruzados (BRL, MXN, JPY, EUR, AUD, CAD, CLP) | exacto al peso |

Residuales conocidos:
1. **5 forwards de PROTECCION** con nominal 353-400 MM USD: la hoja los fuerza a
   P&G=0 (exclusion especifica del macro, aun por confirmar la regla). Son el
   unico gap material (~847 mil MM COP si se valoran).
2. **GBPUSD** (55 filas): residual ~0,2% por fila (curva de puntos GBP); < 0,15%
   del P&G total. Por afinar el escalado de `Fwd_GBPUSD_Diaria`.

## Salida

`bmk.output.valoracion_fwds.escribir` genera
`valoracion_fwds_industria_<corte>.xlsx`:
- hoja **Detalle**: una fila por forward (par, operacion, nominal, strike, plazo,
  Tasa VPN, Tasa Forward, valor COP derecho/obligacion, P&G).
- hoja **Resumen**: P&G y valor por AFP / portafolio / par.

Fuente de forwards: Formato_415 del archivo SFC (`normalizar_415`) en produccion,
o la propia hoja `Fwd Industria` para validacion.

## Pendiente

1. Confirmar la regla de exclusion de los 5 forwards grandes de PROTECCION.
2. Afinar el escalado de la curva de puntos GBP.
3. Reemplazar las curvas placeholder por las del corte real (mismo CSV).
4. Validar el normalizador Formato_415 -> paridad/operacion/nominal contra un
   corte con hoja Fwd Industria del mismo mes.
