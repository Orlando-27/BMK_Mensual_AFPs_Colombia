# Entregables de referencia

Archivos de referencia generados a partir del corte **junio 2026**
(`insumos/sfc/2026_06portainvdeta.xls`).

## Archivos

- **`diccionario_nomenclaturas_SFC.xlsx`** — Diccionario de nomenclaturas de la
  SFC para entender la información de la Super:
  - `Formato_351 activos` y `Formato_415 derivados`: columnas clave (índice,
    campo, descripción).
  - `Cod tipo derivado`, `Cod indice tasa`, `Cod periodicidad`,
    `Cod amortizaciones`, `Cod base dias`: códigos y su significado.

- **`COLFONDOS_fondos_capital_privado_internacional.xlsx`** — Fondos de capital
  privado internacional (en el exterior) de COLFONDOS:
  - `Resumen`: definición, totales (70 fondos, 57 gestores, Vr mercado).
  - `Por clase`: FCPE (private equity intl) + FCPIE (real estate intl).
  - `Gestores`: raíz del nombre (HARBOURVEST = 1 gestor; HDS VII y VIII = 2 fondos).
  - `Fondos`: lista completa fondo por fondo, con gestor, clase, Vr mercado.

  > Definición usada: "capital privado internacional" = clases **FCPE + FCPIE**
  > (en el exterior, sin ISIN colombiano). FCP / FCPI / FCPAPP son locales y
  > quedan fuera. El "gestor" se extrae por heurística (corta el nombre en el
  > primer número romano / LP / FUND).

## Documentación completa

La metodología y el detalle del proceso están en `../docs/`:
`GUIA_COMPLETA.md`, `VALORACION_SWAPS.md`, `PROCESO_BENCHMARK_MENSUAL.md`,
`VALIDACION_vs_MACROS.md`, entre otros.
