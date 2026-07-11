# Diagnóstico de fórmulas de celda (extraídas de los .xlsx)

> Cierra el contrato de datos: la lógica que **no** estaba en VBA sino en fórmulas de celda.
> Fuente: `BMO_VERSION2.xlsx` y `Detallado_VERSION2.xlsx`. Fórmulas crudas en `BMO_formulas.md` y
> `Detallado_formulas.md`. Aquí la interpretación para reimplementar en Python.

## 1. Cadena de clasificación de activos (hoja `Clasificación` del BMO)

Columnas auxiliares (fila N por activo). `[2]CLASIFICACION` = tabla externa de INICIO MACRO
MULTIFONDOS (ya extraída a `config/reference/clasificacion.csv`).

| Col | Fórmula (interpretada) | Significado |
|---|---|---|
| `S` | `IF(D="",IF(C="",TRIM(F),TRIM(C)),TRIM(D))` | **Nombre/nemo** limpio (prioridad D→C→F). |
| `R` | `IF(T∈{TDPE,FINDI}, S, IFERROR(VLOOKUP(S, Dic!BD:BE,2), AQ))` | **Clave de clasificación** (con casos especiales por nemo). |
| `X` | `IFERROR(VLOOKUP(S, Dic!BJ:BK,2), VLOOKUP(I, VECTOR!M:N,2))` | **Moneda** (excepción por ISIN → moneda del vector). |
| `AB` | `IFERROR(VLOOKUP(R, CLASIFICACION!A:B,2), VLOOKUP(CONCAT(E,F,I,AA), Dic!AM:AY,2))` | **CLASE DE INVERSIÓN**. Si ambos fallan → **ND = instrumento nuevo**. ⭐ |
| `AE` | `…CLASIFICACION!A:Y col 5 …` | **Moneda** clasificada. |
| `AF` | `…CLASIFICACION!A:Y col 6 …` | **Ubicación/Clasificación** (R FIJA / R VARIABLE / CAJA). |
| `AG` | `…CLASIFICACION!A:Y col 7 …` | **Riesgo**. |
| `AH` | `IF(AF="R FIJA", VLOOKUP(S, VECTOR!A:C,3), "NA")` | Tasa/periodicidad si es renta fija. |
| `AN` | `XLOOKUP(S,'R.Fija Int'!D,I) ▸ XLOOKUP(S,SX!D,J) ▸ "NA"` | **Tasa facial** (RF internacional/local). |
| `AJ` | `Z / Y / VLOOKUP(X, VECTOR!N:O,2)` | **Precio calculado** = mercado/nominal/FX. |
| `AK` | `VLOOKUP(S ó R, VECTOR!A:C,2)` | **Precio Infovalmer** (vector). |
| `AL` | `IFERROR(AK-AJ, 0)` | **Diferencia de precio** → control de precios atípicos (alerta). |
| `AS` | flag `Dic!BG:BH` (valor nominal vs residual) | Tratamiento de nominal por excepción de ISIN. |

**Detector de instrumento nuevo (confirmado):** un activo es nuevo/no contemplado cuando `AB` (y por
ende `AE/AF/AG`) no resuelve por ninguno de los dos VLOOKUP → queda `#N/A`/ND. En Python: anti-join de
la clave `R` contra `clasificacion.csv` y de la clave concatenada `E&F&I&AA` contra el diccionario
`AM:AY`; si ninguno matchea → alerta `INSTRUMENTO_NUEVO_SIN_CLASIFICAR`.

## 2. Clasificador de tipo de swap (hoja `Swaps` del BMO, `BX2`) ⭐ puente con el motor v6

`BX2` asigna el tipo de swap a partir de moneda de cada pata (CJ/CR), índice (CE/CM) y facial:

- Ambas COP → `IRS IBR+PB` (si facial empieza "IBR o/n+") o `IRS IBRCOP`.
- USD con SOF1 → `IRS SOFUSD`; si no `IRS LIBORUSD`.
- COP↔USD → `CCS COPUSD`; COP↔UVR → `CCS COPUVR`; COP↔EUR → `CCS COPEUR`; UVR↔EUR/USD → `CCS UVR…`;
  COP↔USD con LIBOR e IBR → `CCS IBRLIB`.

**Estos son exactamente los tipos que consume `swap_valuator_v6.py`** (IRS SOFUSD, IRS IBRCOP, CCS
COPUSD, CCS COPUVR). El puente es directo: la columna BX del BMO = columna `TIPO SWAP` de `SWAPIND`.

### Campos de exportación de swap (BY..CU → Calculadora / motor v6)
| Col | Contenido |
|---|---|
| `BY` | Identificador con prefijo por AFP (`SK`/`Col`/id) — `BZ` = AFP (VLOOKUP Dic A9:C12). |
| `CE`/`CM` | Índice pata der/obl (`SOF`→`SOF1`). |
| `CF`/`CN` | **Facial/spread**: si SOF1/IBR y facial=0 → 0; si ≠0 → `ROUND(%,2)`; else `/100`. (Coincide con la regla P1 del motor v6.) |
| `CG`/`CO` → `CH`/`CP` | Periodicidad → código: Trimestral→`TV`, Semestral→`SV`, Anual→`AV`, Mensual→`MV`, else `Vnto`. |
| `CI`/`CQ` | Clase (fija/flotante) pata der/obl (`AB`/`AD`). |
| `CJ`/`CR` | Moneda pata (`COU`→`UVR`). |
| `CK`/`CS` | Nominal pata der/obl; `CT = CK-CS`. |

## 3. Forwards (hoja `Fwd Industria`, fila 23)

- `AN` **compra/venta**: si pata COP en obligación → VENTA; en derecho → COMPRA; con USD según lado;
  else "REVISAR". → base de la alerta `PARIDAD_INVERTIDA`/exposición.
- `AV`/`AY` construyen el **par** (orden USDCOP/EURUSD/GBPUSD/AUDUSD directo, resto invertido).
- `AW`/`AX` aplican **`1/strike`** a pares invertidos (USDJPY, USDMXN, USDBRL, USDCAD) — la corrección
  de paridad que el manual hace a mano.

## 4. Tablas de excepción (hoja `Diccionario SFC`, bloques auxiliares)

- `BD:BE` — **NEMOs especiales** (ISIN→NEMO): FONINMOBAL, PEI, HCOLSEL, ICOLCAP, NUAMCO, CNEC, FONDARRER.
- `BG:BH` — **VM×Nominal** excepción (ISIN→1): MX0MGO0000H9, MX0MGO0000P2.
- `BJ:BK` — **Moneda** excepción (ISIN→moneda): IE00B53QG562→EUR.
- `BB:BC` — **Periodicidad** normalizada: SEMESTRAL→S, AL VENCIMIENTO→P, ANUAL→A, MENSUAL→M, TRIMESTRAL→T.

## 5. Tabla de monedas/FX (hoja `VECTOR`, M:P)

`Moneda | ICO | Valor COP`: COP=1; USD/EUR/CAD/MXN/GBP/CLP/JPY/BRL/AUD = `VLOOKUP` al vector; `COU`
(UVR) y `PEN` con reglas especiales (UVR manual si `Parámetros!I17=TRUE` → el control de UVR de día no
hábil). → se vuelve tabla de FX del pipeline; la regla UVR alimenta la alerta `UVR_DIA_NO_HABIL`.

## 6. Implicaciones para Python

- La **clasificación** es un doble VLOOKUP determinista → reimplementable con `merge`/`map` contra
  `clasificacion.csv` + diccionario concatenado; el fallo de match es la señal de instrumento nuevo.
- El **tipo de swap** (BX) se calcula con las mismas reglas → produce la columna `TIPO SWAP` que el
  motor v6 ya sabe valorar. Integración limpia.
- Los **controles de precio** (AL) y **paridad** (AW/AX) son cálculos simples → alertas.
- Las **tablas de excepción** son pequeñas → se versionan como CSV de referencia.

## 7. Consolidación (hoja `Benchmark` del Detallado) — fase F4

Referencias externas: `[3]` = INICIO MACRO MULTIFONDOS (`CPRVI`, `CLASIFICACION`, `CLAS ACCIONES`),
`[4]` = `Duracion_Indexado_Precia`, `[5]` = `DICCIONARIO` (divisas).

| Col(s) | Lógica | Uso |
|---|---|---|
| `I:K` | `SUMIFS(X, A=portafolio, B=AFP) + cupones swaps` | Valor de mercado por portafolio/AFP. |
| `S:U` | `N/I` … | Participaciones (peso). |
| `X` | valor de mercado del activo | Base de exposición. |
| `Z` | `X / valor_fondo(portafolio, AFP)` | **Peso** del activo en el fondo. |
| `AA` | precio: SWAP→`VECTOR DE PRECIOS!G:J`; resto→vector | Precio unitario. |
| `AB` | `AA*Z` (o notas: `BJ*BF/valor_fondo`) | **Exposición ponderada**. |
| `AC` | `IF(BJ<CP,"CORTO",IF(BJ>=LP,"LARGO","MEDIANO"))` | **Plazo** (umbrales `CP`/`LP`). |
| `AD:AZ` | `VLOOKUP(D, [3]CPRVI!A:BA, MATCH(encabezado)) * X` | **Exposición por región y moneda** (fondos/notas/acciones). |
| `BA` | `VLOOKUP(D,'[3]CLAS ACCIONES'!A:B,2)` | Clasificación de acciones Colombia. |
| `BC` | `VLOOKUP(E ó D,[3]CLASIFICACION!A:O,15)` | Duración RF internacional. |
| `BF`/`BG`/`BK`/`BL` | `VLOOKUP` a hoja `NE` | **Notas estructuradas** (capital protegido, exposición). |
| `BJ` | `VLOOKUP(D,Parametros!AW:AX)` ó `(vcto-hoy)/365` | **Duración** del activo. |
| `BN`/`BO` | `Duracion_Indexado_Precia` ó `SWAPTF(...)` (VBA) | Duración/valoración; **swaps vía `SWAPTF`** + hoja `Curvas`. |
| `BQ`/`BR`/`BT` | `COUPPCD`/`WORKDAY` sobre `Fecha_Valoración` | Fechas de cupón corrido (sensibilidad). |

**Notas para Python:**
- La **exposición por región/moneda** (AD:AZ) es un producto matricial `peso_CPRVI(ISIN) × valor` →
  reimplementable como merge con `cprvi.csv` + broadcast. Fallo de match en CPRVI → alerta (fondo/nota
  sin mapeo).
- La **valoración de swaps del Detallado** (`SWAPTF`, hoja `Curvas`) es una vía **interna** paralela a
  la Calculadora Swap; para el pipeline se usa el **motor v6** (que ya reproduce la Calculadora). En la
  fase F4 se concilia que el Detallado consuma los VPN/DM del motor en lugar de recomputar con `SWAPTF`.
- Fechas de cupón usan funciones financieras de Excel (`COUPPCD`, `COUPNCD`, `WORKDAY`) → replicar con
  utilidades de calendario/festivos Colombia.

_(Fórmulas crudas completas en `BMO_formulas.md` y `Detallado_formulas.md`.)_
