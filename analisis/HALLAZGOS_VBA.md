# Hallazgos del análisis de los 5 archivos Excel (VBA + estructura)

> Destilado del código VBA extraído (~10.000 líneas) y del inventario de hojas.
> Fuente de verdad para reimplementar la lógica en Python. Archivos originales
> NO versionados (ver `.gitignore`); el VBA crudo está en `analisis/vba/`.

## Inventario de hojas por archivo

| Archivo | Hojas |
|---|---|
| **Benchmark Mensual Optimizado.xlsb** (BMO) | Parámetros · Columnas · Diccionario SFC · Eliminados · Clasificación · FutLoc Industria · FutInt Industria · Fwd Industria · Swaps · VECTOR · R.Fija Int · SX · CSA BMK |
| **Benchmark (Detallado).xlsb** | Procedimiento · Controles · Colfondos · Benchmark · Voluntarias · Formato de Apuestas PO-CS · Formato de Apuestas · Parametros · Formato de Apuestas CP · Opt Colfondos · Opt Ind · Cupónes Swaps · SWAP · FutLoc Industria · FutInt Industria · Fut Colf · Fwd Colfondos · Fwd Industria · Curvas · Curvas_2 · NE · TASAS · VECTOR DE PRECIOS · CUF |
| **Calculadora Swap.xlsb** | Instrucciones · RUTAS · LIB-IBR on · INICIO · SWAPIND · SWAPCC_COPUVR · SWAPIRS_SOF · SWAPCC_COPUSD · SWAPCC_IBRLIBOR · SWAPIRS_IBRoncop · SWAPIRS_IBRME · SWAPIRS_LIBOR · MONEDA · DATOS |
| **Sensibilidad_Fwds.xlsb** | Control · Curvas · 572 · Ind · Ejercicio Individual |
| **INICIO MACRO MULTIFONDOS.xlsm** | PROCEDIMIENTO · CONTROL · 597 · T620 · T597 · PORTAFOLIOS · 620 · CLASIFICACION · TITULOS NUEVOS · VECTOR DE PRECIOS · IND2 · IND · CLAS ACCIONES · Resumen CV · VENTAS · CPRVI |

## Inventario de macros VBA (módulos con lógica real)

| Archivo | Módulo | Líneas | Rol (inferido) |
|---|---|---:|---|
| BMO | Módulo1 | 464 | Importación de derivados/swaps + reemplazo diccionario SFC/AFP |
| BMO | Módulo4 | 287 | (por analizar) |
| BMO | Módulo3 | 268 | (por analizar) |
| BMO | Módulo2 | 118 | (por analizar) |
| BMO | Módulo5 | 86 | (por analizar) |
| Detallado | Módulo1 | 2711 | Núcleo del Benchmark Detallado (importación/consolidación) |
| Detallado | Módulo2 | 557 | (por analizar) |
| Detallado | Módulo11 | 393 | (por analizar) |
| Detallado | Módulo3 | 210 | (por analizar) |
| Detallado | Módulo4 | 103 | `traer_sensibilidad` (integración con Sensibilidad_Fwds) |
| Calculadora Swap | Módulo1 | 621 | Valoración swaps (VPN/DM) |
| Calculadora Swap | Módulo3 | 499 | (por analizar) |
| Calculadora Swap | Módulo4 | 123 | (por analizar) |
| INICIO MACRO MULTIFONDOS | IMPORTACION | 488 | Importación/clasificación de activos |
| INICIO MACRO MULTIFONDOS | Module1 | 122 | (por analizar) |
| Sensibilidad_Fwds | Módulo1 | 588 | Cálculo de sensibilidad de forwards |

## Lógica clave ya decodificada (BMO / Módulo1)

### `Parámetros!C17` = ruta del archivo SFC (`AAAA_MMportainvdeta.xlsx`)
Toda la importación abre ese libro y lee las hojas `Fmto-351` y `Fmto-415`.

### `DiccionarioSFC()` — construcción de diccionarios de control
Copia valores únicos (RemoveDuplicates) desde el archivo SFC a la hoja `Diccionario SFC`:
- `Fmto-351!C` → tipos (dic. B9) · `Fmto-415!A` → (dic. A9)
- `Fmto-351!G`, `Fmto-415!F` → (dic. G26/F26)
- `Fmto-351!Z` → (dic. O2) · `Fmto-415!J` → (dic. S3)
Luego llama `Columnas` e `Importar_derivados_industria`.

### `Importar_derivados_industria()` — clasificación de derivados por **columna J de Fmto-415**
La **columna J** (campo 10) del formato 415 es el **tipo de derivado**. Filtra y enruta:

| Código J | Destino | Tipo |
|---:|---|---|
| **5** | `FutLoc Industria` | Futuros locales |
| **6** | `FutInt Industria` | Futuros internacionales |
| **1** | `Fwd Industria` | Forwards |
| **4** | `Fwd Industria` (al final, marcado `TRM`) | Futuros de TRM |
| **16** | `Swaps` | Swaps CCS (Cross Currency) |
| **17** | `Swaps` | Swaps IRS (tasa de interés) |

**Mapeo de columnas Fmto-415 → hoja destino (valores, no fórmulas):**

*FutLoc Industria (J=5):* B→A, Z→B, F→F, AG→I, AE→J, AB→K, AC→L, X→M, A→N.

*FutInt Industria (J=6):* F→E, AF→F, AG→G, Z→H, P→J, W→K, X→L, Z→M, AE→N, AD→O, A→R.

*Fwd Industria (J=1 y J=4):* W→AK, AA→AL, AC→AM, AB→AO, AD→AP, AE→AR, F→AS, X→AT, A→AU.
Tras pegar TRM (J=4), se marca `AZ = "TRM"` y se arrastran fórmulas en AN, AQ, AV, AW, AX, AY
(cálculo de compra/venta y nominal según la posición).

*Swaps (J=16 y J=17):* se copia el bloque completo `A2:BV` de las filas visibles.
Luego se arrastran fórmulas en `BX2:CU` (columnas de valoración que después van a la Calculadora Swap).

### `ReemplazarDiccionarioSFC()` — normalización de nombres
Recorre `Diccionario SFC`:
- Filas 9+: col A (buscar) → col C (reemplazo) = **nombres de AFP**.
- Filas 4+: col F (buscar) → col E (reemplazo) = **nombres de portafolios**.
Aplica el reemplazo (solo celdas con valor, no fórmulas) en: `FutLoc Industria`, `FutInt Industria`,
`Fwd Industria`, `Clasificación`.

### Celda de control de importación
`Parámetros!F19` debe ser `True` para que la importación proceda (control "Positivo Importación").
Tiempos de cada macro se registran en `Parámetros!L3, L5, L9`.

## Implicación para la reimplementación en Python
- La importación de derivados es **determinista**: filtro por código + mapeo fijo de columnas →
  reimplementable 1:1 con `pandas` leyendo `Fmto-415`/`Fmto-351`.
- Las "fórmulas arrastradas" (AN/AQ/AV.., BX:CU) encapsulan cálculos financieros que hay que
  **decodificar de las celdas** (leer el texto de la fórmula en una copia `.xlsx`) y replicar.
- Los "diccionarios" (AFP, portafolios, tipos SFC) se vuelven **tablas de configuración** en el
  proyecto Python.

## Pendiente de analizar (siguiente iteración)
1. BMO: Módulo2/3/4/5 (Columnas, Vector Completo, Mercado<1000/FA, CSA).
2. Detallado: Módulo1 (2711 líneas) — consolidación final.
3. Calculadora Swap: Módulo1/3 — modelo de valoración VPN/DM.
4. INICIO MACRO MULTIFONDOS: IMPORTACION.bas — clasificación de activos.
5. Sensibilidad_Fwds: Módulo1 — cálculo de sensibilidad.
6. Decodificar las fórmulas de celda clave (requiere copia `.xlsx` de BMO y Detallado).
