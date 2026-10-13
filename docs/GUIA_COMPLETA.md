# Guía Completa — Benchmark Mensual AFPs Colombia

> Reimplementación en Python (corre en la nube, sin Excel) del proceso mensual de
> Benchmark de los fondos de pensiones colombianos, que replica las macros
> `Benchmark (Detallado).xlsb` / `Benchmark Mensual Optimizado.xlsb` /
> `Calculadora Swap.xlsb`.
>
> Última actualización: corte **junio 2026** (SFC 30-jun, valorado 10-jul).

---

## 1. Qué hace el proceso

A partir del archivo de la Superintendencia Financiera (SFC) `AAAA_MMportainvdeta.xls`
(portafolio de inversiones por AFP y portafolio), el proceso:

1. **Ingesta** los dos formatos: `Formato_351` (activos) y `Formato_415` (derivados).
2. **Clasifica** cada activo (R FIJA / R VARIABLE / CAJA / NOTAS ESTRUCTURADAS) y
   detecta instrumentos nuevos.
3. **Valora los derivados**: forwards, futuros y swaps (motor propio, sin Excel).
4. **Arma la hoja Benchmark** importable (activos + CAJA + derivados valorados).
5. **Genera controles, validaciones y alertas**.
6. **Empaqueta** todas las salidas en un `.zip`.

### Regla M-1 y fechas de valoración

- El Benchmark del mes **M** usa el **corte SFC del mes M-1**.
- **Activos**: se valoran al **corte** (fin de mes).
- **Forwards**: se revalúan a la **fecha de publicación** (~10 días hábiles después),
  porque la curva de puntos forward es un insumo externo de esa fecha.
- **Swaps**: se valoran al **corte** (la TRM de las patas USD es la del corte).
- **Swaps de cámara (IRS SOFR)**: al **día hábil anterior al corte** (rezago de 1
  día de la cámara CRCC).

Ejemplo junio 2026: SFC corte **30-jun**, forwards a **10-jul**, swaps a **30-jun**,
IRS SOFR a **29-jun**.

---

## 2. Cómo correr

Todo el proceso es un solo comando (`run_corte.sh`). En Cloud Shell:

```bash
# Primera vez
cd ~ && git clone -b claude/benchmark-automation-setup-5nv2b6 \
  https://github.com/Orlando-27/BMK_Mensual_AFPs_Colombia.git
cd BMK_Mensual_AFPs_Colombia && bash run_corte.sh

# Siguientes veces
cd ~/BMK_Mensual_AFPs_Colombia && git pull && bash run_corte.sh

# Descargar el zip de salidas
cloudshell download ~/BMK_Mensual_AFPs_Colombia/salidas_2026-06-30.zip
```

`run_corte.sh` corre con **junio por defecto**. Para otro corte se sobre-escriben
las variables (ver cabecera del script):

```bash
CORTE=2026-07-31 MES=Julio \
SFC=insumos/sfc/2026_07portainvdeta.xls SFC_PREV=insumos/sfc/2026_06portainvdeta.xls \
CURVAS=insumos_diarios/2026-07-31 \
CURVAS_PUB=insumos_diarios/2026-08-10 FECHA_PUB=2026-08-10 \
CURVAS_PREV=insumos_diarios/2026-07-30 FECHA_PREV=2026-07-30 \
MACRO=insumos/Benchmark_corte_julio.xlsx \
bash run_corte.sh
```

Insumos que necesita un corte:
- SFC del corte (`insumos/sfc/AAAA_MMportainvdeta.xls`) y del mes anterior (alertas).
- Curvas INFOVALMER del corte, de publicación y del día anterior
  (`insumos_diarios/AAAA-MM-DD/`: `SwapCC_*.txt`, `Fwd_*.txt`, `Matriz_TC`, `IND.csv`).
- (Opcional) El Benchmark del macro para conciliar (`insumos/Benchmark_corte_<mes>.xlsx`).

---

## 3. Metodología por clase de activo

### 3.1 Activos (Formato_351) — clasificación y regiones

- Clasificación por cadena de reglas (hoja `Clasificación` del BMO): clave R →
  tabla de referencia → clase de inversión / ubicación / clasificación / riesgo.
- Regla de último recurso: códigos `clas_sfc` de renta fija cuyas clases mapean
  todas a R FIJA → R FIJA (evita ND en instrumentos nuevos).
- **Regiones y monedas**: la hoja Benchmark tiene columnas de moneda (USD, COP,
  EUR, JPY, BRL…) y de región/equity (GLOBAL, US EQUITY, EUROPA+UK, JAPON EQUITY,
  EM ASIA, EM GLOBAL, LATAM, ALTERNATIVE). Los instrumentos **nuevos** no están en
  esa tabla y hay que clasificarlos; el output de **alertas** los sugiere (ver §6).

### 3.2 Forwards (Formato_415, tipo 1)

Decodificado al decimal contra la hoja `Fwd Industria` del macro. Cada pata se
descuenta con la curva de **su** moneda (paridad cubierta bi-moneda):

```
Tasa VPN     = strike          × DF(curva_descuento_strike)
Tasa Forward = (spot + puntos)  × DF(curva_descuento_spot)
P&G          = (VPN − Forward) × nominal   (VENTA; se invierte en COMPRA)
```

| Par | Descuento strike | Descuento spot | Puntos |
|---|---|---|---|
| USDCOP | FWTCOP = **IBR** | FWTUSD = **USDCO** | no |
| EURUSD | LIBBTS = **USDOIS** | FWTEUR = **EUROIS** | no |
| AUDUSD/GBPUSD | USDOIS | USDOIS | no |
| USDBRL/USDMXN | USDOIS | USDOIS | sí (FWPBRL/FWPMXN) |
| resto | USDOIS | USDOIS | no |

**Roll mensual**: un forward vencido (o que vence el día de valoración) se renueva
a `fecha_valoración + 1 mes` (cobertura que rueda mes a mes). Verificado: los 122
forwards vencidos de junio ruedan todos al 10-ago.

Resultado corte junio: **FORWARD −0.3%** vs el macro.

### 3.3 Swaps (Formato_415, tipo 16/17) — motor v6

Ver `docs/VALORACION_SWAPS.md` para el detalle completo. Resumen:

- **Motor v6** valora cada pata como bono (cupones + nocional), con curvas
  INFOVALMER. Casa con la SFC/Precia (Formato_415 col 53/54) a **<1% por pata**.
- **Curvas de descuento/proyección** por tipo:
  - IRS IBRCOP: descuenta IBR_COL_USD, proyecta IBR.
  - IRS SOFUSD: USDOIS proyecta y descuenta.
  - CCS COPUSD: pata COP → COP_COL_USD, pata USD → USDOIS.
  - CCS COPUVR: pata COP → IBR_COL_USD, pata UVR → IBRUVR.
- **Cámara (IRS SOFR)**: se compensan en CRCC con 1 día de rezago → se valoran al
  **día hábil anterior** (VPN completo, no variación).
- **Capitalizantes (periodicidad código 7 = "al vencimiento")**: swaps que
  capitalizan el interés (amortizaciones código 4). La pata flotante acumula el
  índice desde la compra y a ~5 años vale ~1.5× el nominal. El motor los modela
  como bullet tomando el nocional capitalizado del valor reportado por Precia.
- **Presentación en la hoja**: nominal de la pata obligación **negativo** (como el
  macro), precio limpio por pata (~par, sin negativos), ISIN/identificador visible.

### 3.4 Futuros y CSA

- Futuros locales (TES) e internacionales: valorados con el VECTOR DE PRECIOS.
- CSA (caja internacional): extraída con la TRM del propio archivo SFC.

---

## 4. Las DOS versiones de la hoja Benchmark

Por decisión del equipo, se generan **dos archivos**:

| Archivo | Metodología | vs macro | Para qué |
|---|---|---|---|
| `hoja_Benchmark_<corte>_precia.xlsx` | Curvas INFOVALMER + capitalización | SWAP +5% | **Producción** — casa con la SFC/Precia a <1% por pata (regulador) |
| `hoja_Benchmark_<corte>_macro.xlsx` | Swaps tomados del macro por identificador | SWAP ~0% | **Conciliación** contra el Benchmark histórico |

**Por qué difieren ~1% en swaps**: el macro usa las curvas de su propia
Calculadora (`.xlsb`), que difieren ~1% de las de INFOVALMER; además el macro NO
capitaliza los IRS que capitalizan. El motor v6 validó a **<0.1%** contra la
Calculadora cuando se le dan SUS mismos insumos (validación de marzo entregada);
el ~1% de junio es puramente la fuente de curvas.

---

## 5. Validaciones

- **vs SFC/Precia** (`comparar_swaps_sfc.py`): el motor casa con los valores que
  reporta la SFC (col 53/54) a **0.31–0.72% por pata**, correlación de MtM 0.99–1.00.
- **vs macro** (`comparar_todo.py`): por clasificación y por AFP. Corte junio:
  FORWARD −0.3%, R FIJA/R VARIABLE/CAJA/NOTAS 0.0%, SWAP +5.2% (versión precia) /
  ~0% (versión macro). **TOTAL ≈ 0%**.
- **Consolidado de precios** por ISIN × mes (`consolidar_precios.py`).

---

## 6. Alertas mensuales

`alertas_mensuales.py` compara el corte actual vs el mes anterior:

1. **Instrumentos nuevos** (ISIN que no existía), renta variable primero, con
   **región y moneda sugeridas** (para clasificarlos en la hoja Benchmark).
2. **Cambios significativos** (|Δ valor de mercado| ≥ 30%): amortizaciones,
   redenciones, rebalanceos grandes.

Ejemplo junio: `LU3300334797` (JPMorgan Japan) → JAPON EQUITY nueva;
`XS2343802141` (nota Skandia) −50% por amortización del nominal residual.

---

## 7. Nomenclaturas de la SFC

`diccionario_nomenclaturas_SFC.xlsx` consolida: layout de columnas del Formato_351
(activos) y Formato_415 (derivados), códigos de clasificación `clas_sfc`,
periodicidad/base/tipo de derivado, y las 26 hojas de la tabla de codificaciones.

---

## 8. Salidas (el zip curado)

`run_corte.sh` empaqueta en `salidas_<corte>.zip` los entregables:

- `hoja_Benchmark_<corte>_precia.xlsx` y `..._macro.xlsx`
- `controles_<corte>.xlsx` (controles formulados, todo OK)
- `valoracion_swaps_industria_<corte>.xlsx` (swap por swap: VPN completo, precio,
  duración modificada, convexidad; columnas de cámara aparte)
- `valoracion_fwds_industria_<corte>.xlsx`, `valoracion_futuros_industria_<corte>.xlsx`
- `sabana_condiciones_faciales_swaps_<corte>.xlsx`
- `comparacion_vs_macro_<corte>.xlsx`, `comparacion_swaps_vs_sfc_<corte>.xlsx`
- `alertas_mensuales_<corte>.xlsx`, `precios_consolidados_<anio>.xlsx`, `resumen.json`

---

## 9. Cabos sueltos (pendientes)

1. **~1% de las patas de swap vs el macro**: viene de que el motor usa curvas
   INFOVALMER (.txt) y el macro usa las de su Calculadora (.xlsb). Para cerrarlo,
   extraer las curvas de swap de la Calculadora del corte (como `extraer_insumos.py`),
   a costa de atar el proceso al Excel.
2. **Capitalizantes desde cero**: hoy el nocional capitalizado se toma del valor
   Precia (observable diario). Para reconstruirlo sin Precia haría falta la serie
   histórica de fixings de IBR (2021→fecha).
3. **Consolidado de precios año completo**: necesita los SFC de ene–abr en
   `insumos/sfc/`.

---

## 10. Mapa de documentos

- `docs/GUIA_COMPLETA.md` — este documento (visión end-to-end).
- `docs/PROCESO_BENCHMARK_MENSUAL.md` — proceso según el manual de la SFC.
- `docs/ARQUITECTURA_PIPELINE.md` — arquitectura del código.
- `docs/VALORACION_SWAPS.md` — metodología de swaps a fondo (motor v6, cámara,
  capitalizantes, precia vs macro).
- `docs/VALORADOR_FWDS.md` — forwards.
- `docs/VALORADOR_FUTUROS.md` — futuros.
- `docs/MOTOR_SWAPS_V6.md` — internals del motor v6.
- `docs/SABANA_SWAPS.md` — sábana de condiciones faciales.
- `docs/HOJA_BENCHMARK.md` — estructura de la hoja importable.
- `docs/VALIDACION_vs_MACROS.md` — resultados de validación.
- `docs/valoracion_swaps_hallazgos.md` — bitácora de hallazgos de swaps.
