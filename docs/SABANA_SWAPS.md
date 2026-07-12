# Sabana de condiciones faciales de swaps + insumos de valoracion

La valoracion de swaps (IRS/CCS) usa el **motor v6** (`bmk/pricing/swaps/swap_valuator_v6.py`,
validado <2% vs Precia). El motor consume, por fecha de valoracion:

1. **SWAPIND** — la sabana de condiciones faciales (misma para las 3 fechas).
2. **INICIO** — FECHA_VAL + TRM, UVR, EURCOP de esa fecha.
3. **12 curvas cero cupon** (`curva_*`) de esa fecha.

## 1. Sabana de condiciones faciales (la arma el pipeline)

`bmk/pricing/swaps/sabana.py` la construye desde Formato_415 (tipo 16/17) del
archivo SFC. Output: `sabana_condiciones_faciales_swaps_<corte>.xlsx`.

Corte abril: **1.348 swaps industria** — IRS SOFUSD 719, IRS IBRCOP 497,
CCS COPUSD 128, CCS COPUVR 4.

Campos por swap (de Formato_415):

| Campo | Origen (col 415) |
|---|---|
| TIPO SWAP (IRS/CCS ...) | derivado de indices + monedas de ambas patas |
| id_contrato | 11 |
| AFP / portafolio | 0 / 5 |
| F.Compra / F.Vcto / F.Liquidacion | 22 / 23 / 24 |
| Pata DER: moneda, nominal, indice, tasa facial, periodicidad, base | 26,27,37,39,38,40 |
| Pata OBL: moneda, nominal, indice, tasa facial, periodicidad, base | 28,29,41,43,42,44 |

Indice tasa facial: `FS`=Fija, `SOF`=SOFR, `IBR`=IBR (decodificado).

> **Por confirmar con el diccionario SFC**: los codigos numericos de **periodicidad**
> (col 38/42: valores 2/4/5) y **base/day-count** (col 40/44: valores 3/5/6/7). Hoy
> se mapean con una tabla tentativa (`PERIODICIDAD`, `BASE` en sabana.py). Si el
> mapeo esta mal, la frecuencia de cupon / conteo de dias afecta el VPN.

## 2. Insumos que faltan (los provee el usuario) — por cada una de las 3 fechas

El motor v6 necesita, para CADA fecha:

- **TRM, UVR, EURCOP** de la fecha (hoja INICIO).
- **12 curvas cero cupon** (`curva_*`): IBR, USDOIS/SOFR, USD-CO, LIBOR-UVR/COP,
  IPC, DTF, colateral, etc. — se extraen del **.xlsb de Precia "Calculadora Swap"**
  de esa fecha con `bmk/pricing/swaps/extraer_insumos.py`.

En el repo solo estan los insumos de **2026-03-31** (fecha de validacion del motor).

## 3. Las 3 fechas (manual, paso 5)

Los swaps se valoran en **3 fechas** para el "Benchmark del mes M" (corte = ultimo
dia del mes M-1):

1. **Dia habil anterior al cierre** (p. ej. 2026-04-29).
2. **Ultimo dia del mes / fecha de corte** (p. ej. 2026-04-30).
3. **Dia anterior a la publicacion** (~dia 9 del mes siguiente, p. ej. 2026-05-09) —
   este reemplaza la Calculadora Swap del proceso diario.

## Resumen: que pasar para valorar los swaps de abril

Por cada una de las **3 fechas** (2026-04-29, 2026-04-30, ~2026-05-09):
- El **.xlsb de Precia "Calculadora Swap"** de esa fecha (de ahi salen las 12 curvas
  + TRM/UVR/EURCOP), o directamente los CSV `curva_*` + INICIO.

La sabana de condiciones faciales ya la genera el pipeline desde el archivo SFC;
solo hay que confirmar el mapeo de codigos de periodicidad/base.
