# Benchmark Mensual Optimizado (BMO) — Lógica de módulos VBA

> Análisis de los módulos VBA del archivo `Benchmark Mensual Optimizado.xlsb` (BMO).
> Fuente de verdad para la reimplementación en Python.
> Módulo1 ya está documentado en `analisis/HALLAZGOS_VBA.md` (importación de derivados + diccionario SFC);
> aquí se cubre el **orquestador (Módulo4)**, la **importación de activos y fórmulas (Módulo2)**, el
> **vector de precios / limpieza <1000 / FA / reorganización de Fwds (Módulo3)**, la **exportación al
> Detallado, CSA y utilidades (Módulo4)**, y las **copias de seguridad (Módulo5)**. Se referencia Módulo1
> solo donde el flujo lo invoca.

## Celdas de configuración de la hoja `Parámetros` (transversales)

| Celda | Contenido | Usado por |
|---|---|---|
| **C5** | Fecha SFC / fecha de corte del mes (Date) | Módulo5 (nombre de carpeta AAAA\MMMM) |
| **C14** | Ruta del **Vector de precios** (`.xlsb`/`.xlsx`) del último día del mes | Módulo3 `V_PRECIOS_IMPORTAR_COMPLETO` |
| **C16** | Ruta del archivo **Benchmark (Detallado)** de salida | Módulo4 `Actualizar_detallado` |
| **C17** | **Ruta del archivo SFC** `AAAA_MMportainvdeta.xlsx` (insumo primario) | Módulo1, Módulo2, Módulo3 `Columnas`, Módulo4 `Importar_CSA_BMK` |
| **F19** | Control **"Positivo Importación"** (debe ser `True`) | Módulo1, Módulo2 (activos), Módulo1 `Importar_Swaps` |
| **I19** | Control **"Positivo Exportación"** (debe ser `True`) | Módulo4 `Actualizar_detallado` |
| **L3** | Tiempo (s) `DiccionarioSFC` | Módulo1 |
| **L4** | Tiempo (s) `Importar_activos_industria` | Módulo2 |
| **L5** | Tiempo (s) `ReemplazarDiccionarioSFC` | Módulo1 |
| **L6** | Tiempo (s) `V_PRECIOS_IMPORTAR_COMPLETO` | Módulo3 |
| **L7** | Tiempo (s) `F_Alternativo_Mercado_1000` | Módulo3 |
| **L8** | Tiempo (s) `Formulas_Clasificación` | Módulo2 |
| **L9** | Tiempo (s) `Importar_Swaps` | Módulo1 |
| **L14** | Tiempo (s) `Actualizar_detallado` | Módulo4 |
| **L19** | Tiempo (s) total `Importar` (orquestador) | Módulo4 |
| **Nombre `R_IND`** | Ruta base (sin extensión) del CSV de indicadores; `Importar_indicadores` le concatena `.CSV` | Módulo3 |

---

## Módulo4 — Orquestación, exportación al Detallado, CSA, utilidades

### `Sub Importar()` — Orquestador maestro de importación (los "7 pasos")
1. **Propósito.** Ejecuta en secuencia toda la importación mensual en BMO, mostrando el UserForm de
   progreso `frmProgreso` (7 pasos). Equivale a correr los botones a–h del Paso 4 del manual.
2. **Entradas.** Ninguna directa; delega en las sub-macros (que leen `Parámetros!C17`, `C14`, `F19`).
3. **Salidas.** Escribe `Parámetros!L19` (tiempo total). Cada sub escribe sus propias hojas.
4. **Secuencia** (con etiqueta del progreso):
   - Paso 1/7: `DiccionarioSFC` (Módulo1) → construye diccionarios + `Columnas` + `Importar_derivados_industria`.
   - Paso 2/7: `Importar_activos_industria` (Módulo2).
   - Paso 3/7: `ReemplazarDiccionarioSFC` (Módulo1) → normaliza nombres AFP/portafolios.
   - Paso 4/7: `V_PRECIOS_IMPORTAR_COMPLETO` (Módulo3) → vector + SX + R.Fija Int + indicadores.
   - Paso 5/7: `F_Alternativo_Mercado_1000` (Módulo3) → limpieza FA y <1000.
   - Paso 6/7: `Formulas_Clasificación` (Módulo2) → arrastra fórmulas `P:AS`.
   - Paso 7/7: `Importar_Swaps` (Módulo1) → swaps CCS/IRS + `Importar_CSA_BMK`.
5. **Controles.** No valida directamente; cada sub aborta con `MsgBox` si `F19<>True` o si falta el archivo.
6. **Riesgo Python.** Es el pipeline: en Python es una función `importar()` que llama a los pasos en el
   mismo orden. `Calculation=xlManual`/`ScreenUpdating` son irrelevantes fuera de Excel.

### `Sub Actualizar_detallado()` — Exportación de industria al Benchmark Detallado
1. **Propósito.** Abre el `Benchmark (Detallado)` y vuelca los derivados y activos de **la industria**
   (todas las AFP menos Colfondos) a sus hojas. Corresponde al Paso 8-A del manual ("Exportar industria").
2. **Entradas.**
   - `Parámetros!C16` = ruta del Benchmark Detallado (se abre como `wbDestino`).
   - `Parámetros!I19` = control de exportación; si `<>True` → aborta con MsgBox.
   - Hojas origen (BMO): `Fwd Industria`, `FutInt Industria`, `FutLoc Industria`, `Clasificación`, `CSA BMK`.
3. **Salidas** (en el Detallado): hojas `Fwd Industria`, `FutInt Industria`, `FutLoc Industria`, `Benchmark`.
4. **Filtro de industria (clave).** Se copian **solo** las filas cuyo nombre de AFP ∈
   **`{"PROTECCIÓN", "PORVENIR", "OLD MUTUAL"}`** (excluye COLFONDOS y SKANDIA/otros). El campo del filtro:
   - `Fwd Industria`: AutoFilter **Field 33** (fila 22 encabezado, datos desde fila 23).
   - `FutLoc Industria`: AutoFilter **Field 14** (encabezado fila 1, datos fila 2).
   - `Clasificación`: AutoFilter **Field 17** (columna Q).
   - `CSA BMK`: AutoFilter **Field 16** (AFP) **y Field 25 ≠ 0** (encabezado fila 2, datos fila 3).
5. **Mapeo de columnas por hoja (valores, no fórmulas):**
   - **Fwds:** origen `A23:F`→destino `A23`; `H`→`H`; `AG`→`AG`; `AH`→`AH`. Luego AutoFill de fórmulas
     `G23`→`G`, `I23:AF23`→`I:AF`, y copia de formato `A23:AH23` hacia abajo.
     Antes limpia en destino: `A23:F`, `H23`, `AG23`, `AH23`, `G24:G`, `I24:AF`.
   - **FutInt:** origen `E2:R`→destino `E2`; luego AutoFill fórmulas `S2:AK2`→`S:AK`.
     Limpia destino `A2:R` y `S3:AK1000`.
   - **FutLoc:** origen (filtrado) `A2:N`→destino `A2`; AutoFill fórmulas `O2:AF2`→`O:AF`.
     Limpia destino `A2:N` y `O3:AF1000`.
   - **Benchmark (activos):** de `Clasificación` filtrada: `P2:Q`→`Benchmark!A255`; `R2:AH`→`D255`;
     `AJ`→`U255`; `AO`→`V255`. Los activos se apilan **a partir de la fila 255**. Limpia antes
     `Benchmark!A255:V` y `W256:BS`.
   - **CSA:** de `CSA BMK` filtrada, `O3:AJ`→ se apila en `Benchmark` en la primera fila libre
     (`última fila col A + 1`).
   - **Fórmulas finales:** AutoFill `Benchmark!W255:BS255`→hasta última fila con datos en A.
6. **Salidas de tiempo.** `Parámetros!L14`.
7. **Instrumentos nuevos / juicio.** El manual (Paso 8b–f) indica revisiones manuales posteriores en el
   Detallado (FutInt códigos F/G/AE/M a texto, NE, swaps nuevos). La macro solo hace el volcado mecánico.
8. **Riesgo Python.** El **offset fijo fila 255** en `Benchmark` es una convención rígida (filas 1–254 son
   la parte de Colfondos que no se toca). Replicar apilado industria→CSA en ese bloque. El conjunto AFP
   industria está **hardcodeado** `{PROTECCIÓN, PORVENIR, OLD MUTUAL}` → parametrizar como config. Nótese
   inconsistencia: en `FutInt` no se filtra por AFP (copia todo `E2:R`), a diferencia de Fwd/FutLoc/Clasif/CSA.

### `Sub Importar_CSA_BMK()` — Importar cuentas internacionales (caja) desde SFC
1. **Propósito.** Trae las "Cuentas CSA" (caja internacional) del archivo SFC a la hoja `CSA BMK`.
   Último paso de importación (Paso 4-h). Llamada al final de `Importar_Swaps`.
2. **Entradas.** `Parámetros!C17` (archivo SFC) → hoja **`Cuentas CSA`**, rango `A1:G200`.
3. **Salidas.** `CSA BMK!A1` (PasteSpecial valores). Activa `Parámetros`.
4. **Riesgo Python.** Rango fijo `A1:G200` (asume ≤199 cuentas). La hoja del SFC se llama `Cuentas CSA`
   (distinta de `Fmto-351`/`Fmto-415`).

### `Sub exportar_swaps_para_julián()` — Exportar swaps con formato calculadora
1. **Propósito.** Crea un `.xlsx` nuevo con el bloque de valoración de swaps para entregar a otra persona
   (formato para pegar en la Calculadora Swap).
2. **Entradas.** BMO hoja `Swaps`, rango **`BX1:CU100000`** (columnas de valoración generadas por fórmulas).
3. **Salidas.** Archivo `Swaps Formato Calculadora.xlsx`, hoja `Datos`, en ruta **hardcodeada**:
   `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Actualización Mensual\`. Guarda como `xlOpenXMLWorkbook`.
4. **Riesgo Python.** Utilidad de export; ruta M:\ hardcodeada. En Python = escribir un `.xlsx` con las
   columnas BX:CU (valores) al almacenamiento configurable.

---

## Módulo2 — Importación de activos y fórmulas de Clasificación

### `Sub Importar_activos_industria()` — Importa activos (Fmto-351) a `Clasificación`
1. **Propósito.** Importa los **activos y notas estructuradas** del formato 351 a la hoja `Clasificación`
   (Paso 4-b del manual). No filtra por AFP: trae toda la industria, excluyendo solo el Fondo Alternativo.
2. **Entradas.**
   - `Parámetros!C17` = archivo SFC → hoja **`Fmto-351`**.
   - `Parámetros!F19` = control; si `<>True` → MsgBox "Revisar Controles Activos antes de importar" y aborta.
   - Valida `Dir(rutaArchivo)` (existencia del archivo).
3. **Salidas.** Hoja `Clasificación` (limpia antes `A3:AL100000`). Escribe `Parámetros!L4` (tiempo).
4. **Filtro de limpieza.** AutoFilter en `Fmto-351` **Field 7 (columna G) ≠ `"FONDO ALTERNATIVO"`**
   (excluye FA en el origen). Copia solo `SpecialCells(xlCellTypeVisible)`.
5. **Mapeo de columnas Fmto-351 → Clasificación (valores):**

   | Origen (Fmto-351) | Destino (Clasificación) | Contenido inferido |
   |---|---|---|
   | G | A | (portafolio / tipo — el filtrado por FA) |
   | C | B | AFP |
   | AB | C | Clase de inversión |
   | CD | D | |
   | Z | E | |
   | R | F | **Clas SFC** (código de clasificación) |
   | AH | G | |
   | AF | H | |
   | AI | I | |
   | AK | J | |
   | AM | K | |
   | AN | L | |
   | BB | M | **Valor de mercado** (usado en filtro <1000) |
   | BE | N | |
   | AR | O | |

6. **Reglas de limpieza.** Excluye FA vía filtro Field 7. (El filtro <1000 se hace después, en Módulo3.)
7. **Riesgo Python.** Mapeo determinista → replicable con pandas. La columna destino **M** = valor de
   mercado es la que luego dispara la regla <1000; la columna destino **F** = Clas SFC alimenta la
   clasificación de títulos nuevos (juicio). Todas las columnas destino A–O se pegan como valores; las
   fórmulas están en `P:AS` (ver siguiente sub).

### `Sub Formulas_Clasificación()` — Arrastra fórmulas de la hoja Clasificación
1. **Propósito.** Rellena las columnas calculadas de `Clasificación` arrastrando la fila 2 (Paso 4-f).
   Se hace en macro aparte "por velocidad" (nota del manual).
2. **Entradas.** Hoja `Clasificación`; determina `filaFinal` por columna A.
3. **Salidas.** AutoFill de **`P2:AS2` → `P2:AS{filaFinal}`**. Escribe `Parámetros!L8`.
4. **Cálculos.** Las fórmulas de `P:AS` encapsulan la lógica de clasificación (clase de inversión, moneda,
   región, precio/tasa facial vía búsquedas en VECTOR/SX/R.Fija Int, etc.).
5. **Riesgo Python — ALTO.** El contenido real de las fórmulas `P2:AS2` **no está en el VBA**; hay que
   extraerlo de una copia `.xlsx`/`.xlsm` del BMO. Son ~30 columnas de fórmulas que definen la
   clasificación final exportada (`P:Q, R:AH, AJ, AO` van al Detallado). **Bloqueante** para reimplementar.

---

## Módulo3 — Vector de precios, indicadores, limpieza FA/<1000 y reorganización de Fwds

### `Sub V_PRECIOS_IMPORTAR_COMPLETO()` — Importa vector de precios + hojas SX y R.Fija Int
1. **Propósito.** Trae el vector de precios del último día del mes: hojas `SX` y `R.Fija Int` completas
   (frecuencia/tasa facial de RFI y RFL, "~±400.000 títulos c/u" según manual) más el par ISIN/PRECIO del
   vector consolidado (Paso 4-d).
2. **Entradas.** `Parámetros!C14` = ruta del vector; se abre **ReadOnly**. Hojas origen: `SX`,
   `R.Fija Int`, `Vector`.
3. **Salidas** (en BMO):
   - `SX` y `R.Fija Int`: se limpian (`Cells.ClearContents`) y se copia el `UsedRange` completo
     (SX desde fila 2, R.Fija Int desde fila 1). Ceros → cadena vacía `""`.
   - `VECTOR`: limpia `A4:B800000`; copia `Vector!A4:B{lastRow}` (ISIN, PRECIO) como valores.
   - `Parámetros!L6` (tiempo).
4. **Cálculos.** Al terminar llama `FormulaVECTOR` (arrastra col C) que a su vez llama `Importar_indicadores`.
5. **Advertencia operativa.** El manual: **antes** de correr hay que copiar el vector del último día a
   `Vector Precios\AAAA`. En FDS trae tasas faciales del viernes pero últimas monedas de valoración del mes.
6. **Riesgo Python.** `UsedRange`/`lastRow` dinámicos. La conversión 0→"" es una regla de limpieza a
   replicar. Volúmenes grandes (~400k filas) → usar lectura por chunks.

### `Sub FormulaVECTOR()` — Arrastra fórmula de tasa facial en VECTOR y llama indicadores
1. **Propósito.** Rellena `VECTOR!C` (búsqueda de tasa facial de RF) arrastrando `C4`.
2. **Entradas/Salidas.** Hoja `VECTOR`; `lastRow` por columna B. Limpia `C5:C{lastRow}`, AutoFill
   `C4→C4:C{lastRow}`. Luego `Calculate` y `Call Importar_indicadores`.
3. **Riesgo Python.** La fórmula de `C4` (tasa facial) hay que extraerla de una copia; busca contra
   `SX`/`R.Fija Int`.

### `Sub Importar_indicadores()` — Carga CSV de indicadores a VECTOR (U:AI)
1. **Propósito.** Importa un CSV de indicadores (monedas/tasas de valoración) vía QueryTable.
2. **Entradas.** Nombre definido **`R_IND`** (ruta base) + `".CSV"`. Delimitadores: **tab y punto y coma**
   (`TabDelimiter=True`, `SemicolonDelimiter=True`), codificación **1252**, todas las columnas como texto
   (Array de 46 unos), sin comillas.
3. **Salidas.** Hoja `VECTOR`, destino `$U$1`; limpia antes `U:AI`.
4. **Riesgo Python.** El CSV usa `;`/tab como separador y Windows-1252. Ruta viene del nombre `R_IND`.
   En Python = `pd.read_csv(sep=';', encoding='cp1252')` (o tab), pegado en el bloque U:AI equivalente.

### `Sub F_Alternativo_Mercado_1000()` — Limpieza: valor <1000 COP → Eliminados; FA en Fwds
1. **Propósito.** (Paso 4-e) Elimina de `Clasificación` los activos con **valor de mercado < 1000 COP**
   (moviéndolos a `Eliminados`) y borra el **Fondo Alternativo** en `Fwd Industria`.
2. **Entradas/Salidas.** Hojas `Clasificación`, `Eliminados`, `Fwd Industria`. Escribe `Parámetros!L7`.
3. **Reglas de limpieza (exactas):**
   - **<1000:** limpia `Eliminados!A2:O10000`; `lastRow` por columna **M (13)**;
     AutoFilter `Clasificación` **Field 13 (columna M) Criteria `"<1000"`**; copia visibles `A2:O` a
     `Eliminados!A2` (valores) y **borra las filas visibles** (`EntireRow.Delete`).
   - **FA en Fwds:** `lastRow` por columna **AS (45)**; AutoFilter `Fwd Industria` rango `AK22:AU`
     **Field 9 (= columna AS, novena desde AK) Criteria `"alternativo"`**; borra filas visibles
     `A23:AU` (`EntireRow.Delete`).
4. **Riesgo Python.** Umbral **1000 COP hardcodeado** sobre columna M (valor de mercado). El filtro FA se
   basa en el texto `"alternativo"` (contiene) en la col AS de la zona de Fwds. Los eliminados hay que
   conservarlos (hoja/tabla `Eliminados`) para trazabilidad. `On Error Resume Next` protege el caso
   "ninguna fila visible".

### `Sub ReorganizarColumnasFwdIndustria()` — Reordena Fwds al formato del Detallado
1. **Propósito.** Tras corregir paridades (Paso 6-e, macro "ORDEN FWDS"), copia el bloque de trabajo
   `AK:AZ` a las columnas finales `A:AH` con el layout del Detallado. **No** está en el orquestador
   `Importar` (se corre manualmente después del juicio de paridades).
2. **Entradas/Salidas.** Hoja `Fwd Industria`. Limpia `A23:AH100000`; `lastRow` por columna AK.
3. **Mapeo (valor→valor, desde fila 23):**

   | Origen | Destino | | Origen | Destino |
   |---|---|---|---|---|
   | AK | A | | AS | F |
   | AY | B | | AT | H |
   | AN | C | | AU | AG |
   | AQ | D | | AZ | AH |
   | AX | E | | | |

4. **Riesgo Python.** Determinista. AZ→AH lleva la marca `"TRM"` (futuros de TRM) que puso Módulo1.

### `Sub Columnas()` — Volcado de encabezados SFC a hoja `columnas` (validación)
1. **Propósito.** Copia los **encabezados** de `Fmto-351` y `Fmto-415` a la hoja `columnas` para validar
   que el SFC no cambió el orden/nombre de columnas (control "Columnas"). Llamada por `DiccionarioSFC`.
2. **Entradas.** `Parámetros!C17` → `Fmto-351` fila 1 y `Fmto-415` fila 1 (última col por `xlToLeft`).
3. **Salidas.** Hoja `columnas`: encabezados de 351 en **fila 2**, encabezados de 415 en **fila 5**.
4. **Riesgo Python.** Es un **control de esquema**: en Python = validar que las cabeceras esperadas
   coinciden posición-a-posición; si cambian, todos los mapeos de columnas fallan silenciosamente.

---

## Módulo5 — Copias de seguridad

### `Sub CrearCarpetaYCopiarArchivos()` — Backup mensual de los 5 archivos-herramienta
1. **Propósito.** (Paso 3 del manual) Crea la carpeta `AAAA\MMMM` y copia allí la última versión de los
   5 archivos-herramienta antes de modificarlos.
2. **Entradas.** `Parámetros!C5` (fecha) → `anio = Year(C5)`, `mes = UCase(Format(C5,"mmmm"))`
   (nombre de mes en mayúsculas, español). Ruta base **hardcodeada**.
3. **Salidas.** Copia (con confirmación de reemplazo vía MsgBox Yes/No) a
   `...\Actualización Mensual\{AAAA}\{MMMM}\`. Crea carpetas año/mes si no existen (FileSystemObject).
4. **Archivos copiados (rutas M:\ hardcodeadas):**
   - `...\Actualización Mensual\Benchmark Mensual Optimizado.xlsb`
   - `...\Benchmark\Calculadora Swap.xlsb`
   - `...\Benchmark\Sensibilidad_Fwds.xlsb`
   - `...\Benchmark\INICIO MACRO MULTIFONDOS.xlsm`
   - `...\Benchmark\Benchmark (Detallado).xlsb`

   Base común: `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\`
5. **Riesgo Python.** Nombre de mes en español mayúsculas (`ABRIL`, no `04`) — ojo: distinto de la
   convención `AAAA_MM` del archivo SFC. Rutas M:\ → almacenamiento configurable. En Python = `shutil.copy`
   con estructura `{base}/{año}/{MES}`.

---

## frmProgreso — UserForm de progreso

### `ActualizarProgreso(paso, totalPasos, mensaje)`
- Barra de progreso simple: `lblProgreso.Width = (paso/totalPasos) * Me.Width`; `Me.Caption = mensaje`.
- Solo UI; sin lógica de negocio. Irrelevante para Python (reemplazar por logging/progreso de consola).

---

## Resumen de rutas hardcodeadas M:\ y patrones de fecha

- **Base:** `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\`
- **Carpeta mensual backup:** `{base}Actualización Mensual\{AAAA}\{MES_MAYÚSCULAS}` (Módulo5, `C5`).
- **Export swaps:** `{base}Actualización Mensual\Swaps Formato Calculadora.xlsx` (Módulo4).
- **Rutas por celda (no hardcodeadas):** SFC = `Parámetros!C17`; Vector = `C14`; Detallado = `C16`;
  CSV indicadores = nombre `R_IND`.
- **Patrones de fecha:** archivo SFC `AAAA_MMportainvdeta.xlsx` (mes con 0 a la izquierda, del manual);
  carpeta backup usa **mes en texto español mayúsculas** (`Format(...,"mmmm")`) — convención distinta.

## Reglas de limpieza/filtrado (exactas)

1. **Fondo Alternativo (activos):** `Fmto-351` Field 7 (col G) `<> "FONDO ALTERNATIVO"` al importar
   (Módulo2).
2. **Fondo Alternativo (Fwds):** `Fwd Industria` `AK22:AU` Field 9 (col AS) `= "alternativo"` → borrar
   filas (Módulo3).
3. **Valor de mercado <1000 COP:** `Clasificación` Field 13 (col M) `"<1000"` → mover a `Eliminados` y
   borrar (Módulo3).
4. **Ceros del vector → cadena vacía** en SX / R.Fija Int (Módulo3).
5. **Filtro de industria para exportar:** AFP ∈ `{PROTECCIÓN, PORVENIR, OLD MUTUAL}`
   (Fwd Field 33, FutLoc Field 14, Clasificación Field 17, CSA Field 16 **y** Field 25 ≠ 0) — Módulo4.

## Puntos de instrumentos nuevos / juicio humano

- **Clasificación ND:** columna destino **F** de `Clasificación` = Clas SFC; las fórmulas `P:AS` (Módulo2
  `Formulas_Clasificación`) mapean a clase/moneda/región; los títulos que quedan `ND` requieren
  clasificación manual con INICIO MACRO MULTIFONDOS (Paso 6-c).
- **Fórmulas no visibles en VBA (bloqueantes):** `Clasificación!P2:AS2`, `VECTOR!C4`, `Fwd Industria`
  AN/AQ/AV/AW/AX/AY, `Swaps!BX2:CU2`, y las fórmulas de FutInt/FutLoc/Fwd/Benchmark en el Detallado.
  Deben extraerse de una copia `.xlsx`/`.xlsm`.
- **Paridades / UVR / swaps a mantener:** manejados fuera del orquestador (`ReorganizarColumnasFwdIndustria`
  se corre tras corregir paridades manualmente; UVR manual en `VECTOR!R9`; swaps a mantener en Calculadora).
