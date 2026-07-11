# Proceso Benchmark Mensual — AFPs Colombia (Especificación de referencia)

> **Propósito de este documento:** Servir como fuente de verdad del proceso mensual de actualización
> del *Benchmark Detallado* de portafolios de pensión obligatoria, tal como está descrito en el
> "Manual de Benchmark Mensual – Nuevo Proceso" (V1.0, 16/04/2026, Dirección de Estrategia de
> Inversiones, Colfondos). Este MD se usará como contexto permanente para diseñar la automatización.
> Cualquier decisión de automatización debe ser trazable a un paso de este documento.

---

## 1. Objetivo y contexto

- **Qué es:** Actualización mensual del archivo **Benchmark Detallado**, que contiene el detalle de las
  posiciones de la *industria* (todas las AFP) en los diferentes activos de los portafolios de pensión
  obligatoria (y cesantías), con base en la información reportada por la **Superintendencia Financiera
  de Colombia (SFC)**.
- **Para qué:** Contar con la información más actualizada de la industria para hacer seguimiento de
  posiciones, identificar posiciones largas/cortas y estimar un PyG esperado por entidad.
- **Fuente oficial:** SFC — *Portafolio de inversión detallado*, publicado el **día 10 calendario de
  cada mes**, correspondiente al **último corte mensual** (ej.: el 10 de abril se publica el corte de
  marzo).
- **Entidad que ejecuta:** Colfondos (Dirección de Estrategia de Inversiones / Análisis Cuantitativo).

### AFPs de la industria (nombres estandarizados)
| Nombre en fuente SFC | Nombre estandarizado (interno) |
|---|---|
| SKANDIA AFP - ACCAI S.A. | SKANDIA PENSIONES Y CESANTÍAS S.A. |
| "PROTECCION" | PROTECCION |
| "PORVENIR" | PORVENIR |
| "COLFONDOS S.A." y "COLFONDOS" | COLFONDOS S.A. PENSIONES Y CESANTIAS |
| (OLD MUTUAL / SKANDIA aparece como OLD MUTUAL en algunas hojas) | OLD MUTUAL |

> Nota: en el proceso, **Colfondos** es "nosotros"; el resto de AFP son "la industria". Las posiciones
> de Colfondos NO deben variar en la revisión final porque solo se actualiza información de industria.

---

## 2. Cadencia y disparadores

- **Frecuencia:** Mensual.
- **Disparador:** Día 10 calendario → validar publicación en la web de la SFC.
- **Corte de datos:** Último día del mes anterior (fecha de corte del BMK mensual).
- **Relación con el BMK diario:** Varios insumos vienen del proceso de benchmark **diario** (vector de
  precios, calculadora swap, notas estructuradas, etc.). En el mensual se trabaja con:
  - el corte del **último día del mes** (fecha de corte), y
  - el corte del **día anterior a la publicación** (último BMK diario corrido, ej. 09 de abril).

### Enlaces SFC
- **Pensión obligatoria (portafolio detallado):**
  `https://www.superfinanciera.gov.co/publicaciones/10097246/...portafolio-de-inversion-detallado-10097246/`
- **Cesantías corto plazo (CS-CP):**
  `https://www.superfinanciera.gov.co/publicaciones/61175/...fondos-de-cesantias-61175/`

---

## 3. Rutas de referencia (unidad de red `M:\`)

Base: `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\`

| Alias | Ruta |
|---|---|
| Carpeta mensual | `...\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Actualización Mensual\AAAA\MMMM` |
| Vector de precios | `...\Benchmark\Historicos Macros\Vector Precios\AAAA` |
| Backup SWAP | `...\BDatos\Backup Archivos\AAAA\SWAP` |
| Notas estructuradas (valoración diaria) | `...\BDatos\Backup Archivos\AAAA\Notas Estructuradas` |
| Consolidados publicados | `...\Benchmark\Publicados\MMMM\Consolidado` |

> Convención de nombres de mes: si el mes es de 1 dígito (ene–sep) se antepone `0`
> (ej.: `2026_3portainvdeta.xlsx` → `2026_03portainvdeta.xlsx`).

---

## 4. Inventario de archivos (insumos y herramientas)

### 4.1 Insumo primario (descarga SFC)
- **`AAAA_MMportainvdeta.xlsx`** — descarga SFC del corte del mes. Contiene, entre otras, las hojas:
  - **Formato_351** → renombrar a **`Fmto-351`** (activos y notas estructuradas de la industria).
  - **Formato_415** → renombrar a **`Fmto-415`** (derivados: futuros, forwards, swaps, etc.).
- **Insumo CS-CP:** descarga aparte para Cesantías Corto Plazo (paso 10).

### 4.2 Archivos-herramienta (con macros VBA) — se hace copia de seguridad antes de modificar
| Archivo | Rol |
|---|---|
| **`Benchmark Mensual Optimizado.xlsb`** (BMO) | Núcleo. Importa/clasifica/valida activos y derivados de la industria mediante botones (macros). |
| **`Calculadora Swap.xlsb`** | Valoración de swaps (VPN y DM). |
| **`Sensibilidad_Fwds.xlsb`** | Sensibilidad de forwards de la industria. |
| **`INICIO MACRO MULTIFONDOS.xlsm`** | Clasificación de activos (hoja Clasificación), mapeo moneda/región. |
| **`Benchmark (Detallado).xlsb`** | **Salida final** (mismo archivo del proceso diario). Recibe activos nuevos de la industria. |
| **`EWMA MENSUAL.xlsb`** | Cálculo EWMA (volatilidad) para futuros/renta variable. |
| **`Revisión Benchmark.xlsb`** | Revisión t-1 anterior vs t-1 nuevo y envío de correo. |
| **`Revisión SWAPS.xlsb`** | Revisión de swaps. |
| **`Formato Swaps Detallado.xlsb`** | Formatea swaps de industria para el Benchmark Detallado. |
| **Vector de Precios (Informalmer)** | Vector diario de precios del último día del mes. |
| **Notas Estructuradas — `Valoración NE ...csv`** | Valoración diaria de notas estructuradas. |

---

## 5. Diccionario de hojas relevantes (por archivo)

**Benchmark Mensual Optimizado (BMO):** `Parámetros`, `Columnas`, `Diccionario SFC`, `Eliminados`,
`Clasificación`, `FutLoc Industria`, `FutInt Industria`, `Fwd Industria`, `Swaps`, `Vector`.

**INICIO MACRO MULTIFONDOS:** `PROCEDIMIENTO`, `CONTROL`, `BAT`, `TDAT`, `620`, `PORTAFOLIOS`,
`CLASIFICACION`, `TITULOS NUEVOS`, `VECTOR DE PRECIOS`, `INDJ`, `IND`, `CUAS ACCIONES`, `CPRVI`.

**Benchmark (Detallado):** `Procedimiento`, `Controles`, `Colfondos`, `Benchmark`, `Voluntarias`,
`Formato de Apuestas PO-CS`, `Formato de Apuestas`, `Parámetros`, `Opt Colfondos`, `Opt Ind`,
`Cupones Swaps`, `SWAP`, `Fwd Colfondos`, `FutLoc Industria`, `FutInt Industria`, `Fwd Industria`,
`Fut Colf`, `Curvas`, `NE`, `Curvas_2`, `VECTOR DE PRECIOS`, `TASAS`, `CUF`, `CLASE DE MONEDA;TASA;CLASE TI`.

**Calculadora Swap:** `RUTAS`, `LIB-IBR on`, `INICIO`, `SWAPIND`, `SWAPCC_COPUVR`, `SWAPIRS_SOF`,
`SWAPCC_COPUSD`, `SWAPCC_IBRLIBOR`, `SWAPIRS_IBRoncop`, `Derecho`, `Obligación`, `PlantillaD`,
`PlantillaO`, `Parametros`.

**EWMA MENSUAL:** `S&P`, `Tesoros`, `EWMA JHON`, `Parametros`.

---

## 6. Flujo del proceso (paso a paso)

> Numeración alineada con el manual. Marcado: 🟩 = macro/botón · ✋ = manual · ⚖️ = requiere juicio.

### Paso 1 — Validar publicación SFC ✋
Día 10: verificar en la web SFC que esté publicado el corte del mes. Descargar
`AAAA_MMportainvdeta.xlsx` (contiene Formato_351 y Formato_415).

### Paso 2 — Preparación de insumos ✋
1. Guardar el `.xlsx` en la carpeta mensual `AAAA\MMMM` (con `0` en meses de 1 dígito).
2. Renombrar hojas: `Formato_351`→`Fmto-351`, `Formato_415`→`Fmto-415`.
3. Reemplazar nombres de AFP (ver tabla §1).
4. En hoja `Fmto-415`, garantizar que columnas **W, X, Y** (Fecha celebración/vencimiento/liquidación
   contrato) queden en **formato fecha** (usualmente vienen como texto): *Datos → Texto en columnas →
   Ancho fijo → Fecha AMD → Finalizar*. Guardar.

### Paso 3 — Copias de seguridad ✋
Antes de modificar, guardar en la carpeta mensual la última versión de: BMO, Calculadora Swap,
Sensibilidad_Fwds, INICIO MACRO MULTIFONDOS y Benchmark (Detallado).

### Paso 4 — Importación en BMO (correr botones 1 a 1) 🟩
Abrir `Benchmark Mensual Optimizado` y verificar que la **Fecha SFC (celda C5)** esté al corte.
Correr los botones en orden (la tabla de tiempos muestra la duración estimada):

| # | Botón | Qué hace | Controles a validar |
|---|---|---|---|
| a | **Controles e Importación de Derivados** | Valida orden/tipos de derivados y nombres de portafolios; luego importa. | Panel *Controles de Importación*: `Nombres AFPs`, `Portafolios`, `Clasificación Derivados`, `Nuevas Clases de Activos`, `Valor Mercado`, `Números Títulos`, `Columnas`, `Swaps Revisión`, `NOMINAL<>CAJA`, **`Positivo Importación`** = **VERDADERO**. |
| b | **Importación Activos** | Importa activos y notas estructuradas de `Fmto-351`. | |
| c | **Nombres AFPs y PO** | Reemplaza nombres AFP en todas las hojas de derivados y activos. | |
| d | **Vector Completo** | ⚠️ **Antes de correr**, copiar el vector de precios del último día del mes a `Vector Precios\AAAA`. Trae vector + 2 hojas de RF Internacional y RF Local (~±400.000 títulos c/u) para extraer frecuencia y tasa facial. En FDS trae tasas faciales del viernes pero últimas monedas de valoración del mes. | |
| e | **Mercado < 1000 y FA** | Limpieza en hoja Clasificación: activos con valor < 1000 COP se eliminan (se dejan en hoja `Eliminados`). Borra el FA (Fondo Alternativo). | |
| f | **Fórmulas Clasificación** | Arrastra fórmulas de la hoja `Clasificación` (se hace aparte por velocidad). | |
| g | **Importar Swaps** | Importa swaps (que vienen en hojas independientes). | |
| h | **CSA BMK** | Trae cuentas internacionales (caja). Último botón de importación. | |

### Paso 5 — Cambiar vínculo SFC en BMO ✋
En BMO, cambiar el vínculo del archivo SFC por el del mes que se está actualizando
(`AAAA_MMportainvdeta.xlsx`) en *Vínculos del libro*.

### Paso 6 — Ajuste de controles en hoja Parámetros
- **a. Control de Swaps** ✋ — En hoja `Swaps` verificar estructura; al final del 2º bloque (columna
  **CU**) borrar por completo las filas marcadas con **BORRAR**.
- **b. UVR manual** ✋ — Si el último día del mes es fin de semana/festivo, la macro trae la UVR del
  último día hábil; colocar manualmente la UVR del cierre del mes en hoja `VECTOR` **celda R9**.
- **c. Clasificación de nuevos títulos** ⚖️🟩 — En BMO hoja `Clasificación` filtrar columna **Q** por
  todas las AFP menos COLFONDOS, y columna **AB** (Clase de Inversión) por **ND** (no clasificados).
  Con `INICIO MACRO MULTIFONDOS` (en escritura) y un Consolidado de apoyo, clasificar cada título por
  su **Clas SFC** (columna F). Para cada código SFC (ej. BOEVS → "CORPORATIVO LOCAL" / "Inflación
  CORPORATIVO LOCAL" según moneda) registrar ISIN + clasificación en la hoja `CLASIFICACION` de INICIO
  MACRO MULTIFONDOS. Repetir hasta que **no quede ningún activo sin clasificar (ND)**.
  - **Fondos Mutuos / ETF nuevos** ⚖️ — Solicitar clasificación/diccionario al equipo de Renta
    Variable (para mapear MONEDA y REGIÓN). Agregar en hojas `TITULOS NUEVOS` y `CPRVI`.
  - **Notas nuevas** — agregar en `CPRVI` debajo de los activos de RVI (si no, no mapea moneda/región).
  - **d. Tabla de precios (Top Diferencias)** ⚖️ — Verificar que no haya diferencias grandes; casos
    especiales (NUAMCO, CAMACOL, ACCIONES) se clasifican en hoja `Diccionario SFC`.
  - **e. Forwards — paridades** ⚖️ — En hoja `Fwd Industria` validar Strike Promedio/MAX/MIN por
    moneda. Paridades invertidas (ej. USDJPY viene como JPYUSD): usar `1/…` en columna STRIKE JHON,
    pegar solo en visibles (Alt+;), F9. Correr macro **ORDEN FWDS** (agrupa Fwd y deja TRM a los
    futuros de TRM).
  - **f. FUT locales e internacionales** ✋ — Revisar hojas `FutLoc Industria` y `FutInt Industria`:
    validar quién creó/eliminó futuros. `ND` en rojo = esa AFP no tiene futuros.

### Paso 7 — Correr Calculadora Swap 🟩✋
1. Abrir `Calculadora Swap`.
2. En hoja `SWAPIND` borrar todos los swaps **excepto** los de Colfondos de pensión **voluntaria**
   (esos no vienen en el reporte SFC). Filtrar por portafolios y dejar los creados **después** de la
   fecha de corte (incluye compras hasta el día 10).
3. Copiar de BMO hoja `Swap` columnas **BX→CU** y pegar en `SWAPIND` debajo de los activos que se
   mantienen. Arrastrar fórmulas hacia abajo (columnas A, K, S, W, Z en adelante).
4. Copiar ISIN de columna B → pegar en columna B de hoja `INICIO` después del título SWAP; arrastrar
   fórmulas desde columna C.
5. Correr **3 fechas**: día anterior al último día del mes, último día del mes (fecha de corte) y día
   anterior a la publicación (ej. 09 de abril). Guardar VPN y DM (renombrar los ya creados para no
   eliminarlos) y guardar copia en `Backup Archivos\AAAA\SWAP` y en la carpeta mensual.
   - El del "09" (día anterior) reemplaza la Calculadora Swap del proceso diario.
6. Revisión con `Revisión SWAPS.xlsb`; en BMO hoja `Parámetros` correr macro **Swaps Nuevos y
   Fórmulas**; recamino de vínculos (archivo SFC + calculadoras del último día y del día anterior).
7. Validar tabla ACTUAL vs PERMITIDO (Promedio/Max/Min por tipo de swap) — diferencias deben ser
   pequeñas. Guardar.

### Paso 8 — Exportar industria (BMO → Benchmark Detallado)
Con BMO, `Divisas` (INICIO MACRO MULTIFONDOS) abiertos, correr el último botón de `Parámetros`:
- **A. Exportar industria** 🟩 — Abre el `Benchmark (Detallado)` y pega los nuevos activos de la
  industria en la hoja `Benchmark`.
- Cerrar BMO **sin guardar** (ya se guardó). Validar controles de la hoja `Controles` del Detallado:
  - **a. FWD industria** ✋ — Verificar que estén todos los activos, mismos FUTUROS DE TRM.
  - **b. FutInt Industria** ✋ — Reemplazar por texto los códigos: F (Tipo Subyacente 1–8 →
    Moneda/IFD/INDICE/Deuda pública/Deuda privada/Valores Participativos/Tasas de interés/Materias
    Primas), G (Subyacente = S&P 500), AE (Duración = 0), M (posición 1=COMPRA, 2=VENTA). La
    **exposición** debe quedar equivalente (compra positiva, venta negativa).
  - **c. FutLoc Industria** ✋ — Mismos activos que BMO y fórmulas arrastradas.
  - **d. Control Notas Estructuradas (NE)** ✋⚖️ — En hoja `NE`, donde la columna **X** = `1` significa
    que no encuentra la nota (la borraron): eliminar toda la fila (registrar portafolio, AFP y proxy de
    peso).
  - **e. Actualizar NE al corte** ✋ — Abrir `Valoración NE DD-MM-AAAA.csv`, separar por comas, pegar
    en hoja `NE` (zona Colfondos). Notas en **BRL**: dejar Capital Protegido, Exposición, RV, RF y Vlr
    Mdo en BRL → multiplicar por el valor de la BRL (del vector) y reemplazar. Cambiar la fecha al
    último BMK generado.
  - **f. Notas nuevas** ⚖️ — Agregar ISINes en hoja `NE` respetando estructura; si falta info, buscar
    en term sheet. Notas en otra moneda (BRL) → mismo procedimiento; todo debe quedar en COP.
  - **g. Swaps nuevos de industria** 🟩✋ — Abrir `Formato Swaps Detallado.xlsb`, actualizar vínculos
    con la Calculadora Swap del corte. Copiar ISIN a columna IDENTIFICADOR en hojas Derecho y
    Obligación; F9. Filtrar AFP y quitar COLFONDOS. Pegar en `Benchmark` (Detallado): primero derechos,
    debajo obligaciones. Traer **precio inicial** (del VPN calculadora, columna B→A). Precios iniciales
    = los del BMK reportado por SFC → dejar como **VALOR** en columna U (no fórmula). Ajustar
    periodicidad (Trimestral→T, Semestral→S, Anual→A). Arrastrar fórmulas desde W; F9.
  - **VECTOR DE PRECIOS (Detallado):** traer curva de precios de Swap y curva de duraciones:
    - PRECIOS: en `VECTOR DE PRECIOS` filtrar columna A por SWAP → seleccionar todo → Ctrl+, → Supr.
      Abrir VPN del 09 de abril, apilar 2 columnas, pegar VPN en columna A y arrastrar fórmulas de C/E.
    - DURACIONES: filtrar columna G por SWAP → borrar → mismo ejercicio con DM en columna G.
  - **h. Control Swaps vencidos** ✋ — Igual que BMK diario: dejarlos en la caja (traer del último
    Consolidado de BMK diario).
  - **i. Control sensibilidad (forwards)** 🟩✋ — Traer sensibilidad del forward de la industria:
    guardar en Benchmark Detallado, abrir `Sensibilidad_Fwds.xlsb`, verificar fecha del último BMK,
    correr botón 3 (Traer industria y calcular). En Detallado F11 → correr `Sub traer_sensibilidad`
    (Módulo 4).
  - **j. Cambiar fechas para generación de PDF/Excel** ✋ — Hoja `Formato de Apuestas PO-CS` celda
    **B84** = fecha de corte; hoja `Procedimiento` celda **D43** = fecha de corte.
  - **k. Calcular EWMA** 🟩✋ — Abrir `EWMA MENSUAL.xlsb`: copiar fila anterior y pegar abajo (S&P,
    Tesoros), arrastrar fórmula columna M, dejar como valor el 20%, F9, copiar celda O198 → BK5;
    repetir en EWMA JHON. Guardar.
  - **l. Limpiar Cupones Swap** 🟩 — macro **SWAP MES** (botón naranja).
  - **m. Guardar Benchmark Detallado.** ✋

### Paso 9 — Generación de insumos y revisión final ✋⚖️
1. Generar el Consolidado y los consolidados de portafolios PO-CS (con nombre alterno, sin reemplazar
   el anterior).
2. Revisión con `Revisión Benchmark.xlsb`: comparación **t-1 anterior vs t-1 nuevo**. Las posiciones de
   Colfondos **no deben variar** (solo cambia industria); es más fácil detectar errores/novedades por
   valores atípicos.
3. Enviar por correo la revisión y los PDF del Benchmark (Fondo Conservador, Moderado, Mayor Riesgo,
   Cesantías Largo Plazo, etc.).

### Paso 10 — Cesantías Corto Plazo (CS-CP) ✋
Si el mismo día está publicado el BMK de CS-CP se actualiza; si no, se hace después. Descargar insumo
CS-CP (Portafolio corto plazo AAAA → mes), guardar en carpeta mensual. En Benchmark Detallado hoja
`Formato de Apuestas CP` pegar en la clasificación los valores portados en `CLASE DE MONEDA;TASA;CLASE
TI` del archivo SFC. **Verificar que los valores totales sean iguales.**

---

## 7. Controles / validaciones clave (todos deben ser VERDADERO salvo indicado)

**Controles de Importación (BMO):** Nombres AFPs · Portafolios · Clasificación Derivados · Nuevas
Clases de Activos · Valor Mercado · Números Títulos · Columnas · Swaps Revisión · NOMINAL<>CAJA ·
**Positivo Importación**.

**Controles de Exportación (BMO):** Futuros TES · Futuros TRM · Futuros INT · Forwards · Activos ·
Swaps (puede ser FALSO temporalmente) · Pares FWDs · Nombres AFPs · Portafolios · UVR (a veces FALSO) ·
CSA · **Positivo Exportación**.

**Controles Detallado (extracto):** Composición BMK PR · Opciones Ind · Opciones Colfondos · CSA · NE
Importación · Act. Rentab. Mín. · Futuros RFI/RFL · Swaps · VOLUNTARIAS · Notas Industria · Cupones
Swaps · Control BMK Monedas PV · Control Futuros Internacionales · Sensibilidad Forwards + Fut TRM.

**Validación final:** totales iguales (industria vs consolidado); posiciones Colfondos sin variación.

---

## 8. Puntos que requieren juicio humano / insumos externos (⚖️)

Estos son los cuellos de botella de la automatización — no son mecánicos:
1. **Clasificación de títulos nuevos** por Clas SFC (mapeo a Clase de Inversión, moneda, región).
2. **Fondos Mutuos / ETF / Notas nuevas** → requieren diccionario del equipo de Renta Variable (correo)
   o term sheet.
3. **Diferencias de precio atípicas** (NUAMCO, CAMACOL, ACCIONES) → clasificación especial.
4. **Paridades de divisas invertidas** en forwards → corrección `1/x`.
5. **UVR manual** cuando el cierre es fin de semana/festivo.
6. **Swaps a mantener vs borrar** (voluntaria Colfondos, fechas de creación).
7. **Notas estructuradas borradas** (columna X = 1) y en moneda BRL.
8. **Interpretación de la revisión final** (valores atípicos, novedades).

---

## 9. Salidas / entregables

- **Benchmark (Detallado).xlsb** actualizado (archivo maestro del proceso diario).
- **Consolidados** (total y por portafolio PO-CS) en `Publicados\MMMM\Consolidado`.
- **PDFs** por fondo (Conservador, Moderado, Mayor Riesgo, Cesantías Largo Plazo, CS-CP).
- **Correos:** (i) revisión interna del cambio de benchmark; (ii) NUEVO BMK a usuarios finales.
- **Backups** de VPN/DM Swap y calculadoras en las rutas correspondientes.

---

## 10. Diagnóstico de automatización (borrador — a validar)

| Bloque | Automatizable | Observaciones |
|---|---|---|
| Descarga SFC (paso 1) | Media/Alta | Scraping o descarga directa del `.xlsx`; validar que exista. |
| Preparación insumo (paso 2) | Alta | Renombrar hojas, normalizar AFP, parsear fechas W/X/Y. Trivial en Python (openpyxl/pandas). |
| Copias de seguridad (paso 3) | Alta | Copiar archivos con timestamp. |
| Importación/limpieza activos (paso 4b,e) | Alta | Reimplementable en Python (reglas de filtrado, <1000 COP, FA, etc.). |
| Clasificación nuevos títulos (paso 6c,d) | **Baja/Media** | Requiere juicio; automatizable un motor de reglas + "cola de pendientes" para lo desconocido. |
| Derivados: FUT/FWD/SWAP (pasos 6e-f, 7) | Media | Valoración swap (VPN/DM) y paridades son cálculos financieros reimplementables; requiere curvas/vector. |
| Notas estructuradas (paso 8d-f) | Media | Parseo de CSV + conversión BRL→COP automatizable; notas nuevas requieren term sheet. |
| Sensibilidad forwards (paso 8i) | Media | Cálculo (posición × duración) reimplementable. |
| EWMA (paso 8k) | Alta | Fórmula EWMA estándar. |
| Revisión t-1 vs t-1 (paso 9) | Alta | Comparación automática + reporte de atípicos. |
| Generación PDF/Excel/correos | Media | Plantillas + envío. |
| CS-CP (paso 10) | Media | Análogo al principal. |

> **Riesgo transversal:** todo el proceso hoy vive en Excel + VBA con vínculos entre libros y rutas de
> red `M:\`. La estrategia de automatización debe decidir si se **reimplementa la lógica** (ej. Python)
> o se **orquesta/controla Excel** (mantener macros y automatizar disparadores). Ver preguntas abiertas.

---

## 11. Glosario

- **SFC:** Superintendencia Financiera de Colombia.
- **AFP:** Administradora de Fondos de Pensiones.
- **PO / CS / PV:** Pensión Obligatoria / Cesantías / Pensión Voluntaria.
- **CS-CP:** Cesantías Corto Plazo.
- **BMK:** Benchmark.
- **BMO:** Benchmark Mensual Optimizado.
- **NE:** Notas Estructuradas.
- **VPN / DM:** Valor Presente Neto / Duración Modificada.
- **FA:** Fondo Alternativo.
- **RFI / RFL:** Renta Fija Internacional / Renta Fija Local.
- **RVI:** Renta Variable Internacional.
- **CSA:** Cuentas internacionales (caja).
- **UVR:** Unidad de Valor Real.
- **EWMA:** Exponentially Weighted Moving Average (volatilidad).
- **Fmto-351:** Formato SFC de activos/notas. **Fmto-415:** Formato SFC de derivados.

---

## 12. Decisiones de diseño (acordadas)

| Tema | Decisión |
|---|---|
| **Estrategia** | **Reimplementar en Python** toda la lógica del proceso, independiente de Excel/VBA. |
| **Entorno de ejecución** | **Servidor/nube, sin Excel ni unidad de red `M:\`.** Las rutas `M:\...` se sustituyen por almacenamiento configurable (local/objeto). No se orquesta Excel. |
| **Alcance** | **Pipeline end-to-end** (bloques 1 a 10). Se construirá por módulos pero apuntando al proceso completo. |
| **Fuente de la lógica** | Los 5 Excel-herramienta (con macros VBA) son la **especificación**: se extrae el VBA y la estructura de hojas para replicar reglas y cálculos en Python. |
| **Artefactos en el repo** | En GitHub van los **artefactos destilados** (código Python, dumps de VBA en texto, mapas de columnas, CSV de muestra), **no** los `.xlsb` pesados. |

### Implicaciones técnicas
- Sin `M:\` ni Excel: toda lectura/escritura de datos se hace con librerías Python
  (`pandas`, `openpyxl`, `pyxlsb`, `xlsxwriter`) y rutas parametrizadas por configuración.
- `.xlsb` → valores legibles con `pyxlsb` (solo valores, no evalúa fórmulas); VBA extraíble con
  `oletools`/`olevba`. Para leer **fórmulas** como texto conviene una copia `.xlsm`/`.xlsx`.
- Las macros VBA definen la lógica real (importación, clasificación, valoración): **primer objetivo de
  análisis** antes de codificar.

## 13. Preguntas abiertas / pendientes
- Confirmar los **5 archivos** exactos y su rol (hipótesis: BMO, Calculadora Swap, Sensibilidad_Fwds,
  INICIO MACRO MULTIFONDOS, Benchmark Detallado).
- Definir el **almacenamiento** que reemplaza `M:\` en el entorno de nube (GCS bucket, etc.).
- Método de **entrega de insumos** SFC en producción (¿descarga automática o carga manual?).
- Cómo se resolverán en producción los **insumos externos** (diccionario Renta Variable, term sheets).

_Última actualización: decisiones de estrategia/entorno/alcance acordadas._
