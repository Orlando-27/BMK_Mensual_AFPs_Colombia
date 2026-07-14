# Valoración de swaps — hallazgos y validación

Corte de referencia: **2026-05-31** (Benchmark junio). Motor: `swap_valuator_v6`.
Fuente de verdad: el propio informe SFC (Formato_415, col 53 "Valor del derecho
COP" / col 54 "Valor de la obligación COP") y, como cruce, el Benchmark macro.

## 1. Corrección de periodicidad (el fix grande)

La sábana leía la periodicidad de las columnas equivocadas del Formato_415:

| Campo | Columna correcta | Antes leía |
|-------|------------------|------------|
| Periodicidad recibir | **40** | 38 (indicador fija/variable) |
| Periodicidad pagar   | **44** | 42 (indicador fija/variable) |
| Base de cálculo      | **45** (una, ambas patas) | 40/44 (la periodicidad real) |

Códigos de periodicidad (col 40/44), decodificados contra la `Period.` del macro
en 1436 swaps: **3=Trimestral, 5=Semestral, 6=Anual, 7=Anual**.

El motor usa `base = DAY_BASE[moneda]` (COP=365, USD=360) e ignora la columna de
base; solo importa la periodicidad (define el número de cupones).

**Efecto** (diferencia por pata contra el informe SFC, set limpio 787 swaps con
ambas patas en valor presente):

| | pre-fix | post-fix |
|---|---|---|
| IRS \|dif derecho\| | 7.05% | **0.30%** |
| IRS \|dif oblig.\|  | 6.66% | **0.29%** |
| CCS \|dif der/obl\| | 2.24% / 0.61% | **0.82% / 0.18%** |

La diferencia tiende a cero: mediana global 0.31% / 0.28% por pata, centrada en 0.

## 2. Periodicidad código 7 (pendiente de confirmar)

26 swaps (24 PORVENIR + 2 PROTECCION) tienen periodicidad código 7. Se mapea a
"Anual" (igual que el macro). Contra el informe SFC quedan con residual de
dirección OPUESTA (PROTECCION +7% alto, PORVENIR −32% bajo), así que **no es un
tema puro de frecuencia**; los de PORVENIR parecen capitalizar intereses
(factor ~1.57× nominal). Se mantiene "Anual" (reproduce el Benchmark macro).
Pendiente confirmar el código 7 con el diccionario/mesa.

## 3. Swaps de cámara (CRCC) con reporte parcial — PORVENIR SOFUSD

541 swaps SOFUSD de PORVENIR reportan una sola columna (53 o 54) con un número
pequeño; la otra en 0. NO es el valor presente de la pata (sería ~nocional×TRM).

Estos swaps son **compensados en cámara de riesgo (CRCC) → liquidación diaria**:
el MtM se salda en efectivo cada día hábil, así que en libros solo queda el
**interés causado / carry de pocos días de la pata fija**, no el MtM acumulado.

Evidencia empírica (541 swaps):

| Candidato | Correlación con el valor SFC |
|-----------|------------------------------|
| Valor presente de la pata | no (miles de veces mayor) |
| MtM neto (desde inicio / mensual / diario) | **~0** (−0.07 a 0.13) |
| Interés causado (accrued) | 0.64 |
| **Nocional** | **0.85** |

El valor equivale a ~0.047% del nocional ≈ **~4.4 días de carry** de la pata fija
(523 de 541 reportan la pata FS). Reproducirlo exacto requiere la convención de
días de causación de PORVENIR (varía por contrato: 2.8–8.4 días).

**Importante:** esto NO afecta el benchmark, que usa el MtM de mi motor (validado
a 0.3% por pata en el set limpio). Los 541 no se pueden cruzar contra el número
propio de PORVENIR porque reporta una magnitud distinta (causado de cámara).

## 4. PROTECCION con ambas patas en 0

175 swaps de PROTECCION reportan ambas columnas en 0 (no reportan valor). Se
excluyen del cruce.

## 5. Forwards (resuelto)

La hoja 'Fwd Industria' del macro trae PRIMERO una matriz de consolidacion y
MAS ABAJO (fila ~21) la tabla de valoracion contrato por contrato. Descifrando
esa tabla se replico la metodologia exacta del macro:

  - **NO usa puntos forward.** Descuento bi-moneda (paridad cubierta):
    - USDCOP: `Tasa VPN = strike x DF_COP` (FWTCOP = COPColateralUSD),
      `Tasa Forward = spot x DF_USD` (LIBBTS = USDOIS). Verificado al peso
      (P&G contrato 1 = (VPN-Forward) x nominal, exacto).
    - Cruces (USDX/XUSD): el macro NO descuenta ni aplica puntos (DF~1);
      `Tasa VPN = strike`, `Tasa Forward = spot`. P&G = nom x TRM x (1/spot-1/strike).
  - **Se valora a la FECHA DE PUBLICACION (~8 dias post-corte), no al corte.** Por
    eso el spot del macro (3581.46) != el del corte (3678.15 TRM del 31-may).

Implementado en `bmk/pricing/fwd_valuator.py` (metodologia bi-moneda) y cableado
en el pipeline (`fecha_fwd`) / `correr_corte.py` (`--curvas-pub`, `--fecha-pub`).
La inyeccion del MtM revaluado en la hoja Benchmark quedo REACTIVADA.

Tambien se corrigio un bug de parseo en `_leer_txt_curva` (los Fwd_<par> traen 4
columnas plazo/mid/bid/ask y tomaba bid/ask).

Los CRUCES tambien se descuentan (DF_USD, LIBBTS) en las dos patas:
  - `Tasa VPN = strike x DF_USD`, `Tasa Forward = (spot + puntos) x DF_USD`.
  - Puntos solo en BRL/MXN (premio forward grande por rate-diff alto + contratos
    deep-OTM; sus puntos INFOVALMER estan bien escalados). El resto sin puntos
    (spot x DF_USD): CAD/JPY/CLP casi exactos; los puntos de EUR/CAD/JPY vienen
    en pips (mal escalados) y su premio es pequeno.

**Validacion final** (contra Benchmark junio con las curvas REALES del 8-jun):

| Clasificacion | dif vs macro |
|---|---|
| CAJA, R VARIABLE, NOTAS, OPCIONES | 0.0% |
| R FIJA | -0.1% |
| FORWARD | +0.8% (de -27.5%) |
| **TOTAL** | **+0.026%** |
| PORVENIR / PROTECCION / SKANDIA | 0.0% / 0.0% / 0.0% |

El benchmark reproduce el macro a **+0.026%**, todas las AFP en 0.0%.
`tools/validar_fwd_pares.py` compara mi Fwd Industria vs la del macro por par.

## Herramientas

- `tools/comparar_sfc.py` — mi VPN vs col 53/54 del SFC (set limpio + parciales).
- `tools/comparar_swaps.py` — mi VPN vs Benchmark macro, por swap y por subtipo.
- `tools/pyg_diagnostico.py` — deduce empíricamente la referencia del valor
  reportado (candidatos de P&G / carry) valorando en varias fechas.
