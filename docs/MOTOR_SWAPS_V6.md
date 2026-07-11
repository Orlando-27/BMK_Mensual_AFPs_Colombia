# Integración del motor de valoración de swaps (v6)

> El equipo ya tenía un **motor de valoración de swaps OTC en Python (v6)** validado contra Precia
> (<2% de error en 1.492 swaps al 31-Mar-2026). Se **vendoriza tal cual** en `bmk/pricing/swaps/` y
> reemplaza el módulo que iba a reconstruirse desde la `Calculadora Swap.xlsb`. **No se modifica** sin
> re-validar (es un artefacto validado). Documentación original: `bmk/pricing/swaps/README.md` y
> `CAMBIOS_V6.md`.

## Qué es

- `swap_valuator_v6.py` — motor. Lee CSVs de una fecha y produce VPN (por pata, en COP), Duración
  Modificada y Convexidad. Dinámico por fecha (cero hardcodeos).
- `extraer_insumos.py` — ETL que extrae de la `Calculadora Swap.xlsb` los insumos a CSV:
  - `INICIO_YYYYMMDD.csv` (FECHA_VAL, TRM, UVR, EURCOP, IPC)
  - `SWAPIND_YYYYMMDD.csv` (portafolio de swaps + VPN Precia benchmark)
  - `curva_<NOMBRE>_YYYYMMDD.csv` (11 curvas cero cupón desde la hoja `DATOS`)
  - `VPN_PRECIA_YYYYMMDD.csv` (benchmark oficial, rango B:J de `INICIO`)

## Verificación en este entorno

Ejecutado sobre los insumos incluidos (31-Mar-2026): **1.492 swaps, 0 errores, todos <2% vs Precia**
(máx: CCS COPUSD 1.89% DER, IRS IBRCOP 1.59% DER). Reproduce la tabla de validación del handoff.

## Metodología (resumen)

- **Curvas** con convención por tipo: `simple_360` (USDOIS, IBR_COL_USD, COP_COL_USD, UVR),
  `ea_365` (DTF, LIBORCOP, LIBORUVR, EURCOP). Interpolación lineal sobre tasas.
- **Asignación de curvas por tipo de swap:** IRS SOFUSD→USDOIS; IRS IBRCOP→IBR_COL_USD/IBR_PROJ;
  CCS COPUSD→COP_COL_USD (pata COP) + USDOIS (pata USD); CCS COPUVR→IBR_COL_USD/UVR.
- **Réplica de bono** por pata (cupones + nocional); VPN en COP = `PV_moneda × FX_spot`.
- **Schedule** retrocedido desde vencimiento; filtra cupones ya pagados.
- **DM** = Macaulay/(1+y); **Convexidad** análoga.

## Cómo encaja en el pipeline mensual

Etapa **5 (pricing)** de la arquitectura. Sustituye el paso 7 del manual (correr Calculadora Swap):

```
Calculadora Swap.xlsb (corte)  ──extraer_insumos.py──►  CSVs (INICIO, SWAPIND, curvas)
                                                              │
                                                   swap_valuator_v6.py
                                                              ▼
                                            VPN/DM por swap  ─►  consolidación (Detallado)
```

En producción, las **curvas** hoy salen de la hoja `DATOS` (poblada por las macros `EXPORTAR_*` desde
INFOVALMER). El pipeline puede: (a) seguir extrayendo de la `Calculadora Swap.xlsb` del corte con el
ETL actual, o (b) alimentar las curvas directamente desde los archivos INFOVALMER en GCS (evita Excel).
Se decidirá en la fase de pricing; el ETL ya desacopla el cálculo del origen.

## Puntos abiertos heredados (del handoff)

- **P1 — Spread SOF1 (0.26%)**: el motor lo ignora en pata SOF1 para alinear con MTM Precia; la mesa
  dice que económicamente sí se paga. Pendiente: confirmar con Precia y, si aplica, emitir columnas
  `Precio_Economico_*` paralelas. No bloquea el benchmark (se usa la cifra MTM Precia).
- **P2 — Validación cruzada vs SFC** (INICIO cols L,M): no implementada. Para SOFUSD el valor L parece
  ser PyG, no VPN; requiere definir la métrica exacta.
- **P3 — Notas estructuradas**: existe `condiciones_faciales_notas.csv` (payoffs, vols) para valorar
  notas (Call Spread, etc.); es el insumo del módulo NE del benchmark. Proyecto paralelo, no iniciado.

## Conciliación con el proceso del benchmark

El benchmark mensual usa la **Calculadora Swap** interna, cuyos VPN/DM el motor v6 reproduce (extrae de
la misma `Calculadora Swap.xlsb` y replica el rango `INICIO!B:J`). Por tanto, integrar v6 es equivalente
a correr la Calculadora Swap, pero en Python y sin Excel. En la fase de pricing se validará que los
`VPN/DM SWAP dd-mm-aaaa` que consume el Detallado coincidan con la salida del motor para un corte real.
