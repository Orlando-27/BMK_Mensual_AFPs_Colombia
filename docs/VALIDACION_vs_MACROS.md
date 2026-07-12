# Validación: pipeline Python vs Benchmark de las macros

Comparación de la hoja `Benchmark` que genera el pipeline contra el archivo
`Benchmark (Detallado)` producido por las macros (ground truth del usuario).

## Hallazgo principal: alinear el CORTE

El archivo del usuario `Benchmark mayo.xlsx` se genera **con el corte de ABRIL**
(2026-04-30), no de mayo. Motivo: la SFC publica el día 10; el "benchmark de mayo"
se corre a inicios de mayo usando el último corte disponible (abril).

Evidencia: para `PORVENIR / MAYOR RIESGO / TUVT07220131` el corte de abril tiene
**12 filas** (= archivo del usuario); el de mayo-31 tiene 19. Comparar contra el
corte equivocado producía ~60 diferencias falsas (eran meses de datos distintos).

> Regla: el "Benchmark del mes M" usa el corte SFC del mes **M-1**.

## Resultado con el corte correcto (abril)

| Métrica | Macros | Pipeline | Diferencia |
|---|---|---:|---|
| Activos industria (excl. Colfondos, <1000) | 15.714 | 15.708 | **−6** |
| Valor de mercado industria | 503.33 B | 502.90 B | **−0.08%** |
| Concordancia de clasificación (CLASE/UBIC/CLASIF/RIESGO) | — | — | **100%** |

Conteo por AFP: Porvenir −1, Protección −5, Skandia 0.
Por clasificación: R VARIABLE exacto, NOTAS exacto, R FIJA +4, CAJA −11, ND +1.

## Descomposición de las diferencias residuales (todas explicadas)

Aislando fila por fila (clave AFP+portafolio+clas_sfc+emisor+valor):

1. **6 filas CSA** (cuentas de liquidez internacional, hoja `Cuentas CSA`): su
   Benchmark las incluye; el pipeline aún no. **Explican exactamente el neto −6.**
   Son SALDO CSA ACTIVO en USD → COP (requieren TRM del corte). Paso de importación
   aparte (Importar_CSA_BMK en las macros).
2. **7 filas TES** que difieren por **1 peso** (redondeo de punto flotante) — son
   las mismas filas, sin diferencia real.
3. **3 filas FCP** con **emisor de etiqueta distinta** (macros: "PARTNERS GROUP…";
   SFC crudo: código/"SILVER LAKE…") pero **idéntico valor de mercado** — misma
   posición, distinto rótulo de emisor.

**Conclusión:** salvo las 6 CSA (paso separado), el pipeline reproduce el Benchmark
de industria **posición por posición y al peso**. La lógica de clasificación coincide
al 100%.

## Validación de valores

- **Valor de mercado**: casa al peso en las posiciones (diferencia agregada −0.08%
  atribuible a las 6 CSA faltantes).
- **Valor nominal**: casa exacto en renta fija (ej. TES 1.208.039.000 = 1.208.039.000).
  El gap agregado de nominal proviene de acciones/FCP donde el "nominal" se expresa
  en unidades distintas (# acciones); no afecta el valor de mercado.
- **Precio**: en depósitos = 1 (par) en ambos; en renta fija nuestro precio es proxy
  (Vr mercado / Nominal) hasta integrar el vector.

## Cierre (con CSA incluida) — MATCH AL PESO

Tras incluir las CSA (TRM 3.637,51 derivada del propio archivo SFC):

| Métrica | Macros | Pipeline | Diferencia |
|---|---:|---:|---|
| Activos industria | 15.714 | 15.714 | **0** |
| Valor de mercado | 503,3260 B | 503,3260 B | **0,0000%** |

**Cero diferencias reales.** Las 6 CSA cuadran al peso. El único residual es
cosmético: etiqueta de emisor en 6 CSA (macros: vacío; pipeline: "CSA USD") y
en 3 FCP (macros: "PARTNERS GROUP…"; SFC crudo: otro rótulo) — misma posición y
mismo valor.

**Conclusión: el pipeline reproduce la hoja Benchmark de industria exactamente
(mismo número de filas y valor de mercado al peso), con clasificación 100%.**

## Validación de NOMINAL y PRECIO

El "Valor Nominal" del SFC significa algo distinto por tipo de activo; se replicó
la convención de las macros:

| Tipo | Nominal |
|---|---|
| Renta fija (TES, bonos) | Valor Nominal (col AK) |
| Renta variable (acciones, fondos, ETF) | No. Acciones (col AN) |
| Caja — carteras colectivas (CCA) | No. Acciones (# unidades) |
| Caja — depósitos (DEPVN) | Valor de mercado (monto del depósito) |

La regla exacta está en la tabla **`Diccionario SFC` Z:AB ("NOMINAL A USAR")** de
las macros (67 códigos), extraída a `config/reference/nominal_por_clas_sfc.csv`:
cada `clas_sfc` → `Valor nominal` (AK) / `Nominal residual` (AM) / `No. Unidades`
(AN). Para depósitos con nominal 0 se usa el monto (Vr Mercado / FX = valor en la
moneda del activo).

Resultado (corte abril, 2.531 grupos instrumento-emisor casados):
- **Nominal cuadra: 99.92%** (agregado **+0.0000%**, exacto)
- **Precio (Vr Mercado/Nominal) cuadra: 100.00%**
- Residual: 2 grupos FCPE (diferencia de etiqueta de emisor, valor idéntico).

## ⚠️ Derivados — NO incluidos aún

Las 15.714 filas comparadas son **solo activos + CSA**. La hoja Benchmark de las
macros tiene **18.638 filas**: 15.714 activos/CSA **+ 2.924 derivados** (IRS 2.430,
CCS 254, forwards ~240). La hoja que genera el pipeline **aún no incluye los
derivados**; es el siguiente paso para replicar la hoja completa.

## Opcionales / pendientes

1. **Incluir derivados** (forwards, futuros, swaps) como filas de la hoja Benchmark.
2. Afinar nominal de DEPVE (moneda extranjera) y bonos especiales.
3. Reemplazar el precio proxy por el del vector del corte.
4. Homologar etiquetas de emisor (CSA/FCP) — cosmético.

## Herramienta

`tools/comparar_benchmark.py` genera el reporte de diferencias para cualquier mes.
Recordar usar el corte M-1 para el "Benchmark del mes M".
