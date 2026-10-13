# Valoración de Swaps — Metodología a Fondo

> Motor v6 + ajustes de cámara, capitalizantes y presentación en la hoja Benchmark.
> Corte de referencia: junio 2026.

---

## 1. Universo y fuente

Los swaps vienen del `Formato_415` del SFC (tipo 16 = IRS misma moneda, 17 = CCS
cross-currency). Cada **fila** del Formato_415 es **un swap** con dos patas (derecho
y obligación). El `numero_contrato` (col 11) es el identificador interno (p. ej.
2317); **no** es un ISIN (los swaps privados no tienen ISIN).

Columnas clave del Formato_415 para swaps:

| Col | Campo | Nota |
|---|---|---|
| 11 | numero_contrato | identificador (2317, 2318…) |
| 26/28 | moneda der/obl | COP, USD, COU/UVR |
| 27/29 | nominal der/obl | |
| 37/41 | índice tasa der/obl | FS=fija, SOF=SOFR, IBR=IBR |
| 38/42 | amortizaciones der/obl | **4 = amortiza capital + intereses especiales (capitaliza)** |
| 39/43 | tasa fija / spread | |
| 40/44 | **periodicidad** | 3=Trim, 5=Sem, 6=Anual, **7=Al vencimiento (capitaliza)** |
| 45 | base | 1=ACT/360, 2=ACT/365 |
| 53/54 | **VP der/obl en COP** | valor presente reportado por la SFC/Precia |

## 2. Tipos de swap y curvas (motor v6)

| Tipo | Patas | Descuento | Proyección |
|---|---|---|---|
| IRS IBRCOP | COP fija vs COP IBR | IBR_COL_USD | IBR |
| IRS SOFUSD | USD SOFR vs USD fija | USDOIS | USDOIS |
| CCS COPUSD | COP vs USD | COP: COP_COL_USD / USD: USDOIS | = descuento |
| CCS COPUVR | COP vs UVR | COP: IBR_COL_USD / UVR: IBRUVR | = descuento |

Cada pata se valora como bono replicado (cupones + nocional); los nocionales se
cancelan en el neto pero Precia los reporta por separado. La conversión a COP de
las patas foráneas usa la **TRM del corte** (no la de publicación).

## 3. Validación contra la SFC/Precia

El motor casa con el VP que reporta la SFC (col 53/54), pata por pata:

| Tipo | Error de pata | Correlación MtM |
|---|---|---|
| CCS COPUSD | 0.72% | 1.00 |
| CCS COPUVR | 1.10% | 1.00 |
| IRS IBRCOP | 0.31% | 1.00 |
| No-cámara (total) | 0.82% | 0.99 |

> El MtM **neto** (der − obl) amplifica el error de pata cuando las dos patas son
> casi iguales: un 1% sobre patas de ~17.000 MM se vuelve 20–40% sobre un neto de
> ~450 MM. Por eso se lee el **error de pata** y la **correlación**, no el % del neto.

Validación entregada del motor contra la **Calculadora Swap** (corte marzo, mismos
insumos): **0.05% en SOFR, <0.7% en todo**. El ~1% de junio contra el macro viene
de la **fuente de curvas** (INFOVALMER .txt vs las de la Calculadora .xlsb).

## 4. Cámara de riesgo (IRS SOFR)

Los IRS SOFR se compensan en la **cámara CRCC**, que reporta con **1 día hábil de
rezago**. Por eso se valoran al **día hábil anterior al corte** (29-jun para el
corte de 30-jun), con el **VPN completo** (no la variación).

- En `run_corte.sh`: `--curvas-prev` / `--fecha-prev` (curvas del día anterior).
- `correr_swaps.ajustar_camara(res, res_prev)` reemplaza el VPN por pata de los
  swaps SOFR por el del día anterior; los demás quedan al corte.

> Nota histórica: primero se interpretó "día anterior" como **variación diaria**
> (MtM(corte) − MtM(día anterior)), que casa con la col 53/54 de la SFC pero
> **no** con el macro — el macro lleva el **VPN completo**. Se corrigió a VPN
> completo del día anterior.

## 5. Swaps capitalizantes (periodicidad 7)

Algunos IRS IBRCOP (vto 2027–2031, receive-fixed) tienen **periodicidad 7 = "al
vencimiento"** y **amortizaciones 4**: capitalizan el interés (lo acumulan y lo
pagan una sola vez al vencimiento). La pata flotante acumula el IBR desde la
compra; a ~5 años vale ~1.5× el nominal (verificado: nominal × (1+IBR)^años casa
el `vp_obl` de la SFC a ~2%).

- El motor los modela como **bullet**: VPN = nocional capitalizado (tomado del
  valor Precia col 53/54, observable diario) valorado a un único flujo al
  vencimiento, con precio/duración/convexidad.
- Resultado: los 22 IRS IBRCOP capitalizantes casan **al peso** con la SFC
  (error de pata IRS IBRCOP 0.84% → 0.31%).
- **El macro NO capitaliza** estos swaps (los valora ~par). Por eso la versión
  `_precia` (capitaliza, = Precia) difiere de la `_macro` (no capitaliza) en estos.

## 6. Presentación en la hoja Benchmark

Correcciones para que la hoja coincida con el macro:

- **Identificador/ISIN visible**: las filas de swap muestran el `numero_contrato`
  (2317, 2318…) en IDENTIFICADOR/ISIN, no el rótulo "SWAP TV".
- **VPN completo**: cada pata muestra su VPN completo (p. ej. 2317 TV 17.480 M /
  TF −15.556 M), no la variación de cámara.
- **Nominal de la pata obligación NEGATIVO** (−5.000.000), como el macro.
- **Precio limpio por pata** = ratio a par (~1.0) que calcula el motor; **0 precios
  negativos** (antes la hoja hacía vr/nominal y salía ~3.496 o negativo).

## 7. Resumen de resultados (corte junio)

| Versión | SWAP vs macro | Uso |
|---|---|---|
| `_precia` | +5.2% | producción (casa con SFC/Precia <1% por pata) |
| `_macro` | ~0% | conciliación contra el Benchmark histórico |

Descomposición del +5.2% (versión precia): IBRCOP −60k (capitalización; el macro
no capitaliza), SOFR +55k, CCS +31k (patas ~1% altas amplificadas). Se compensan
parcialmente. Estructura (nominal, precio, identificador): 100% correcta.

## 8. Archivos y herramientas

- `bmk/pricing/swaps/swap_valuator_v6.py` — motor de valoración.
- `bmk/pricing/swaps/sabana.py` — sábana de condiciones faciales (SWAPIND).
- `bmk/derivatives/benchmark_rows.py` — filas de swap para la hoja + inyección.
- `tools/correr_swaps.py` — corre el motor + `ajustar_camara`.
- `tools/comparar_swaps_sfc.py` — valida motor vs SFC/Precia.
- `tools/hoja_version_macro.py` — genera la versión conciliada vs el macro.
- `bmk/output/valoracion_swaps.py` — output swap por swap.
