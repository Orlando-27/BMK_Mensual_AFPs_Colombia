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

### Por qué 99.92% y no 100% — explicado y descartado

El 99.92% **no es un error de cálculo de nominal**. Es un artefacto de la clave
de comparación, que agrupa por **etiqueta de emisor**. Las macros homologan el
emisor de los fondos de capital privado (FCPE) al nombre del gestor
(p. ej. "PARTNERS GROUP PGCS I SECONDARY", "SILVER LAKE SLP V"); el SFC crudo
trae la etiqueta por posición. Al agrupar por esa clave, 2 buckets no alinean:

| Grupo | Filas ellos / nosotros | Nominal ellos | Nominal nosotros |
|---|---|---:|---:|
| PORVENIR/PO FCPE "PARTNERS GROUP…" | 3 / 1 | 2.965.224 | 1.429.035 |
| PORVENIR/PM FCPE "SILVER LAKE SLP V" | 44 / 45 | 37.749,53 | 37.827,05 |

Las posiciones son las mismas; solo caen en un bucket de emisor distinto (a una
posición le pusieron otro rótulo). La prueba de que **no falta ni sobra nominal**:

- **FCPE en total**: 4.686 filas ellos = 4.686 nosotros; nominal **1.171.129.551,50
  = 1.171.129.551,50** (idéntico al centavo).
- **Nominal total industria**: 178.614.864.806.443,75 (ellos) vs
  178.614.864.806.443,81 (nosotros) → diferencia **0,06 COP** sobre 178 billones
  (**+0,000000%**), puro redondeo de punto flotante.

**Conclusión: el nominal cuadra al 100% en valor.** El "99,92%" mide *cuántos
buckets de emisor alinean*, no cuánto nominal coincide; cada fila FCPE calcula
`nominal = No. Acciones` correctamente. Es cosmético (rótulo de emisor), como las
3 filas FCP y las 6 CSA ya documentadas. Queda pendiente (opcional) homologar
etiquetas de emisor para que el 99,92% suba a 100% también a nivel de bucket.

## Derivados — INCLUIDOS (estructura al 100%)

La hoja Benchmark de las macros tiene **18.638 filas**: 15.714 activos/CSA **+
2.924 derivados**. El pipeline ahora genera esas filas (`bmk.derivatives.benchmark_rows`)
y la hoja completa **cuadra en número de filas exacto (18.638 = 18.638, diff 0)**.

Estructura de derivados replicada al detalle (corte abril, industria):

| Tipo | Macros | Pipeline | Detalle |
|---|---:|---:|---|
| FORWARD | 210 | 210 | 15 (afp,port) × 14 divisas (FWUSD..FWKRW) |
| OPCIONES | 30 | 30 | 15 (afp,port) × 2 (OPT USD, OPT EUR) |
| SWAP | 2.684 | 2.684 | 2 patas/contrato; **IRS 2.430 / CCS 254 exacto** |

Los nocionales de swap salen del archivo (`nominal_derecho/obligacion`) y casan
al peso (ej. 3.000.000, 10.000.000). La regla IRS/CCS (misma moneda ambas patas =
IRS) reproduce el split 2.430/254 **exacto**.

### Valor de mercado de derivados — requiere insumos de curvas del corte

La **estructura** (filas, nocionales, tipo) está al 100%. El **valor de mercado**
de cada derivado necesita insumos que NO viajan en el archivo SFC:

- **Forwards**: las macros revaluan cada pata a la fecha de corrida (~10 días
  después del corte) con la curva de puntos forward (hoja `Curvas` /
  `Sensibilidad_Fwds`). Usando el valor presente por pata que reporta el propio
  archivo (Formato_415 col 53/54) se obtiene el valor **al corte**, ~5-16% por
  debajo del de las macros (signos y divisas cero coinciden). Cerrar al peso
  requiere la curva de puntos forward del corte.
- **Swaps**: el VPN por pata lo produce el motor v6 (`bmk.pricing`) con las curvas
  del corte. El repo trae curvas de 2026-03-31; para el corte de abril hacen falta
  las curvas de 2026-04-30 (IBR, USDOIS, LIBOR*, USDCO, etc.).
- **Opciones**: en el Benchmark de las macros la mayoría son 0; se revaluan aparte.

Nota: la hoja se importa en la herramienta diaria donde las columnas formuladas se
arrastran; el valor de mercado de derivados es dato (no re-derivable por fórmula,
porque los forwards se colapsan a divisa) y se completa en la etapa de valoración.

## Opcionales / pendientes

1. **Valor de mercado de derivados**: curva de puntos forward + curvas de swap del
   corte (2026-04-30) para cerrar al peso.
2. Etiqueta de pata swap TF/TV por pata fija/flotante (hoy derecho=TF/obligación=TV;
   totales exactos, sólo cambia el rótulo de ~12 filas).
3. Afinar nominal de DEPVE (moneda extranjera) y bonos especiales.
4. Reemplazar el precio proxy por el del vector del corte.
5. Homologar etiquetas de emisor/AFP (CSA/FCP; SKANDIA↔OLD MUTUAL) — cosmético.

## Herramienta

`tools/comparar_benchmark.py` genera el reporte de diferencias para cualquier mes.
Recordar usar el corte M-1 para el "Benchmark del mes M".
