# Análisis VBA — `INICIO MACRO MULTIFONDOS.xlsm`

> **Rol del archivo en el proceso mensual:** CLASIFICACIÓN de activos por código **Clas SFC → Clase de
> Inversión** y mapeo de **MONEDA / REGIÓN / indexador**. Corresponde al **paso 6c** de
> `docs/PROCESO_BENCHMARK_MENSUAL.md` (clasificación de nuevos títulos, cola de ND). Este libro se usa
> "en escritura" durante el mensual: el analista añade filas a las hojas-diccionario y la macro re-mapea.
>
> Autoría VBA: José Orlando Bobadilla (Nov-2015), con añadidos de J. Franco / F. Escala (2015),
> J. Higuera, K. Taborda (09/2025).

---

## 0. Mapa de módulos y orquestación

| Módulo | Subs | Función macro |
|---|---|---|
| **IMPORTACION** | `Importar`, `Importar_Informes`, `HOJA_IND`, `GENERAR_TXT_VECTORPRECIOS`, `GENERAR_TXT_DURACION`, `guardar`, `Limpiar_Datos_externos`, `Limpiar_Datos_externos2` | Descarga CSV 597/620/IND, arma IND, exporta TXT vector/duración, guarda |
| **Module1** | `cargardatos` (**orquestador**), `CxC`, `ReposNegativos` | Pipeline principal + moneda de títulos en tránsito + signos de repos |
| **Module2** | `Ventas`, `Compras` | Extrae ventas; inserta compras nuevas usando diccionario `TITULOS NUEVOS` |
| **Module3** | `PORTAFOLIOS` | Vuelca T597 → hoja PORTAFOLIOS y llama a la clasificación |
| **Module4** | `copia`, `portafolio` | Backup 597/620 → T597/T620; estandariza portafolios y fechas |
| **Module5** | `limpiar` | Limpieza genérica de rango |
| **Module6** | **`clasificacionPO`** | **NÚCLEO de clasificación**: Clas→Clase Inversión + parse moneda/indexador |
| **Module8** | `borrar` | Limpia hoja PORTAFOLIOS |
| **Módulo2** | `Importar_CSA` (viejo), `CSA_balance` (2025) | Trae balances CSA (caja internacional) a `Resumen CV` |
| **Módulo7** | `Copy` | Arma hoja VECTOR DE PRECIOS con factor de conversión de moneda (IND2) |

**Orden de ejecución real** (`Module1.cargardatos`):
`copia` → `portafolio` → `Ventas` → `Compras` → `PORTAFOLIOS` (→ `clasificacionPO`) → `Módulo7.Copy` →
`RefreshAll` → `CxC` → `HOJA_IND` → `ReposNegativos`.
La importación de CSV se hace por separado con `IMPORTACION.Importar` (llama `Importar_Informes` + `CSA_balance`).

---

## 1. `Module6.clasificacionPO` — NÚCLEO DE CLASIFICACIÓN ⚖️ (el más importante)

**Propósito:** asignar a cada título de la hoja `PORTAFOLIOS` su **Clase de Inversión, moneda, región y
demás atributos** por *lookup* contra la hoja `CLASIFICACION`, y extraer el indexador/moneda del código.

**Entradas:**
- Hoja `PORTAFOLIOS`, columna **C** (para contar filas) y columna **B** = identificador de búsqueda
  (ISIN o código; ver Module3 abajo cómo se construye), columna **F** (col 6) = código/serie SFC de
  respaldo, columna **S** (col 19) = facial/código con indexador.
- **TABLA DE MAPEO PRINCIPAL:** `Sheets("CLASIFICACION").Columns("A:M")` (`RNG`).
  - Columna **A** = clave de búsqueda (ISIN / código).
  - Columnas **B–L** (índices 2–12) = atributos que incluyen **Clase de Inversión, moneda, región**, etc.

**Lógica de clasificación (paso a paso):**
```
Ultimaf = CountA(PORTAFOLIOS!C:C)
' 1) Resolución de la clave: si el identificador de la col B no está en CLASIFICACION,
'    se cae al código SFC de la col F (col 6).
For i = 2 To Ultimaf
    If IsError(VLOOKUP(Cells(i,2), CLASIFICACION!A:M, 2, 0)) Then
        Cells(i,2) = Trim(Cells(i,6))          ' fallback: usa Clas SFC (col F)
    End If
Next
' 2) Volcado de atributos (VLOOKUP columna A:M de CLASIFICACION):
For i = 2 To Ultimaf
    Cells(i,13) = VLOOKUP(clave, RNG, 2, 0)     ' Clase de Inversión (col B de CLASIFICACION)
    Cells(i,14) = VLOOKUP(clave, RNG, 3, 0)
    Cells(i,15) = VLOOKUP(clave, RNG, 4, 0)
    Cells(i,16) = VLOOKUP(clave, RNG, 5, 0)
    Cells(i,17) = VLOOKUP(clave, RNG, 6, 0)
    Cells(i,18) = VLOOKUP(clave, RNG, 7, 0)
    Cells(i,23) = VLOOKUP(clave, RNG, 8, 0)
    Cells(i,24) = VLOOKUP(clave, RNG, 9, 0)
    Cells(i,25) = VLOOKUP(clave, RNG, 10, 0)
    Cells(i,26) = VLOOKUP(clave, RNG, 11, 0)
    Cells(i,27) = VLOOKUP(clave, RNG, 12, 0)   ' moneda/región están dentro de estas 11 columnas
```
- **Columna 29 (AC)** = fórmula de participación de renta fija:
  `=IF(RC[-6]="r fija", IF(OR(RC[-26]="NA",RC[-26]="COB51CB00452"),1, COUNTIFS(...)/COUNTIFS(...)),"NA")`
  → reparte proporcionalmente los títulos "r fija" por emisor.

**Parse de moneda/indexador (col 19 → col 20):** a partir del código facial de la columna **S** (col 19)
se extrae por **longitud del string** el segmento de moneda/indexador a la columna **T** (col 20):
```
If col19 = "No aplica" Or "Varias"  → col20 = col19
Else Select Case Len(col19)
   Case 6  → Mid(col19, 5, 1)
   Case 7  → Mid(col19, 6, 1)
   Case 8  → Mid(col19, 5, 3)
   Case 9  → If Left(col19,4)="IBR0" Then Mid(col19,6,3) Else Mid(col19,5,4)   ' caso especial IBR
   Case 10 → Mid(col19, 5, 5)
   Case 0  → col19="0.00000001 - Ef" ; col20=0.00000001   ' efectivo/caja sin código
   Case Else → Left(col19, 9)
```
Este bloque es **heurística frágil basada en posiciones fijas**; el caso `IBR0` está *hardcodeado*.

**Salidas:** hoja `PORTAFOLIOS`, columnas 13–18, 20, 23–27, 29 (clasificación resultante por fila).

**Manejo de ND (título no clasificado) — CLAVE:**
- No existe escritura explícita de la cadena "ND". Si tras el fallback a Clas SFC el código **sigue sin
  estar en `CLASIFICACION`**, los `VLOOKUP` devuelven **`#N/A`** (error) en las columnas 13–27.
- **Detección de instrumento nuevo = celda `#N/A` en las columnas de clasificación de PORTAFOLIOS.**
  En el proceso BMO paralelo esto se ve como **`ND`** en la columna AB de la hoja `Clasificación` (paso 6c).
- Resolución (manual): el analista agrega la fila (ISIN/Clas SFC → Clase de Inversión + moneda + región)
  a la hoja **`CLASIFICACION`** y re-corre, hasta que **no quede ningún `#N/A`/ND**.

---

## 2. `Module3.PORTAFOLIOS` — construcción de la tabla a clasificar

**Propósito:** volcar cada fila de `T597` a la hoja `PORTAFOLIOS` y **construir el identificador de
búsqueda** (columna B) que usará `clasificacionPO`.

**Lógica del identificador (col B):**
```
If T597.col8 vacío o "NA" Then
    If T597.col9 = "NA" o vacío Then  col2 = T597.col3   ' serie/Clas SFC
    Else                              col2 = T597.col9    ' (2º id, p.ej. nemotécnico)
Else                                  col2 = T597.col8    ' id preferente (ISIN)
```
**Prioridad de clave:** `col8 (ISIN)` → `col9` → `col3 (serie/Clas SFC)`.
**Salida:** hoja `PORTAFOLIOS` cols 1–12, 19, 21, 22. Al final llama `Module6.clasificacionPO`.

---

## 3. `Module2.Compras` — diccionario de TÍTULOS NUEVOS ⚖️

**Propósito:** insertar en `T597` las **compras** (`T620`) que aún no existen en cartera, tomando sus
atributos de un diccionario manual.

**Entradas / TABLA DE MAPEO:** `Sheets("TITULOS NUEVOS").Columns("A:D")` (`TN`); clave = título (col A).
Recorre `T620`, y por cada `"Compra"` busca el título en `T597`; si **no lo encuentra** (`k=Ultimaftaa`),
lo da de alta usando `VLOOKUP` contra `TITULOS NUEVOS`:
- `TN col2` → `T597 col3` (serie / Clas SFC)
- `TN col3` → `T597 col16`
- `TN col4` → `T597 col20` (**código de moneda/indexador**, luego parseado en Module6)
- resto de campos vienen de `T620` (nominal, fechas, valor de mercado, portafolio…).

**Punto de juicio:** un título comprado que **no esté en `TITULOS NUEVOS`** produce `#N/A` en cols 3/16/20
→ hay que agregarlo manualmente a `TITULOS NUEVOS` (y, según doc, a `CPRVI` para Fondos Mutuos/ETF/Notas).

**`Module2.Ventas`:** extrae a la hoja `VENTAS` las filas de `T597` marcadas con `"X"` en col 36 y luego
las elimina de `T597`. No clasifica; solo mueve ventas.

---

## 4. `Module1.CxC` — MONEDA de títulos en tránsito (diccionario `Resumen CV`)

**Propósito (comentario original):** traer la **moneda** de los títulos en tránsito (Cxc/CxP/T620).

**TABLA DE MAPEO:** hoja **`Resumen CV`** — `VLOOKUP` del título contra `'Resumen CV'!C[6]:C[7]`
(la clave y la "Clase de Moneda" en columnas contiguas de esa hoja).
**Entrada:** `T620` col D (identificador; `RC[-21]` desde col Y). **Salida:** `T620` col **Y** = "Clase de
Moneda" (fórmula VLOOKUP arrastrada de la fila 4 a la última).
> El comentario del código lo dice literalmente: *"El diccionario está ubicado en la hoja Resumen CV"*,
> que también relaciona Compras/Ventas por portafolio.

---

## 5. `Módulo7.Copy` — VECTOR DE PRECIOS + factor de conversión de moneda (IND2)

**Propósito:** armar la hoja `VECTOR DE PRECIOS` desde `T597` y calcular el precio en moneda base.

**Lógica de moneda:** col 8 (H) =
`=IF(RC[-2]="$", 1, VLOOKUP(RC[-2], 'IND2'!C1:C9, 5, FALSE))`
→ si la moneda es `"$"` (COP) el factor es 1; si no, busca en **`IND2`** (col 5) la **tasa de cambio**.
Col 9 = valor de mercado / factor / nominal (precio unitario en base). Col 4 = identificador con la misma
prioridad ISIN→nemo→serie que Module3.
**`IND2`** es por tanto la tabla de **conversión de divisa** (moneda → factor/TRM).

---

## 6. `IMPORTACION` — importación de CSV y armado de IND

- **`Importar_Informes`:** descarga por `QueryTables` los CSV **597** (Fmto activos), **620** (derivados/
  compras) e **IND** desde rutas parametrizadas en la hoja `PROCEDIMIENTO`. Aplica un filtro
  *hardcodeado* al 597: `Field:=3, Criteria1:="=AODF5 AH BBH USD              "` (elimina esas filas).
  Borra columnas condicionalmente (`"   Valor CVA/DVA"` en V2).
- **`HOJA_IND`:** abre el **Vector de Precios** externo (`.xlsm`, con password `"INVERSIONES"`), copia
  columnas A:D, F, I:K a la hoja `IND`; añade tasa/frecuencia vía `VLOOKUP(... IND2 ...)`. Marca filas
  con literal `"PORFIN"`.
- **`GENERAR_TXT_VECTORPRECIOS` / `GENERAR_TXT_DURACION`:** concatenan columnas de `IND` con `;` y las
  exportan a `.txt` (formato `xlText`) a rutas de `PROCEDIMIENTO`.
- **`guardar`:** solo guarda si `CONTROL!B19 = "GUARDAR"`; hace `SaveAs` a `R_GUARDAR\Nom_GUARDAR.xlsm`.
- **`Limpiar_Datos_externos` / `_externos2`:** borran Names con `"DatosExternos"` y conexiones externas.

---

## 7. `Módulo2` — CSA (caja internacional)

- **`CSA_balance`** (Katherin Taborda, 09/2025 — versión vigente): abre el archivo CSA
  (`R_CSA\N_CSA` de `PROCEDIMIENTO`), lee hoja **`Balances Esperados`** cols A,B,D,E desde fila 3, y las
  pega en `Resumen CV!AQ3`. Cierra sin guardar.
- **`Importar_CSA`** (Higuera, versión antigua): variante que valida `B3="Portafolio"` en hoja
  `Balance Esperado` y pega en `Resumen CV!E47`; si no, `MsgBox("Revisar CSA")` ✋. **Riesgo:** dos
  versiones con hojas/celdas destino distintas (`Balances Esperados`/AQ3 vs `Balance Esperado`/E47).

---

## 8. `Module4` y `ReposNegativos` — estandarización y signos

- **`Module4.copia`:** copia hojas `597`→`T597` y `620`→`T620` (respaldo de trabajo).
- **`Module4.portafolio`:** normaliza el nombre de portafolio (col 37): si el 4º carácter es `"-"` toma 2,
  si no 3 caracteres. Convierte fechas texto `AAAAMMDD` a fecha (`DateSerial`); vencimiento vacío →
  `1900-01-01`. Borra filas `"NV"` de `T620`.
- **`Module1.ReposNegativos`:** para filas `"REPO PASIVO"` invierte el signo (`*-1`) de columnas de
  nominal/valor (14,18,21,23,25 en 597/T597; 11,12 en PORTAFOLIOS). **Ajuste *hardcodeado*:** si
  `col3 = "COC04CB00137"` fuerza `col9 = 46992` (una fecha serial concreta).

---

## 9. Rutas, nombres y patrones de fecha

- **No hay rutas `M:\` literales en el VBA.** Todas se leen de **rangos con nombre** en las hojas
  `PROCEDIMIENTO` / `CONTROL`: `R_597, N_597, R_620, N_620, R_IND, N_IND, R_VECTOR, N_VECTOR,
  R_VARIABLE, R_BLOOMBERG, R_CSA, N_CSA, R_MACRO, N_MACRO, R_GUARDAR, Nom_GUARDAR,
  R_EXPO_VECTOR, R_EXPO_DUR, MES, Fecha_Actualizacion`. Las rutas físicas `M:\COB\...` viven en esas
  celdas (ver §3 del doc de proceso), **no** en el código → al reimplementar hay que extraerlas del libro.
- **Password del Vector externo:** `"INVERSIONES"` (hardcodeado en `HOJA_IND`).
- **Nombre de libro hardcodeado:** `"INICIO MACRO MULTIFONDOS.xlsm"` (activaciones de ventana).
- **Filtro hardcodeado 597:** `"AODF5 AH BBH USD              "` (con padding de espacios exacto).
- **Códigos hardcodeados:** `"COC04CB00137"`→46992, `"COB51CB00452"` (excepción en fórmula r fija),
  prefijo `"IBR0"`, literal `"PORFIN"`, `"0.00000001 - Ef"`.
- **Patrón de fecha:** insumos `AAAAMMDD` como texto → `DateSerial(Left,Mid,Right)`. Mes de 1 dígito se
  antepone `0` (convención de nombres de archivo, ver doc).

---

## 10. Controles / validaciones

- `CONTROL!B19 = "GUARDAR"` habilita el guardado.
- `CONTROL!B1:F37` y `C3` se recalculan tras importar (paneles de control).
- Barra de progreso (`frm_lcf_ProgressBar`) en `cargardatos`.
- Validación CSA: `MsgBox("Revisar CSA")` si la estructura no coincide (solo versión vieja).
- **No hay validación explícita de ND**: la ausencia de clasificación se manifiesta como `#N/A` y debe
  ser detectada visualmente/por filtro (paso 6c del manual).

---

## 11. Dónde viven los mapeos (resumen operativo)

| Mapeo | Hoja diccionario | Clave → Valor | Usado por |
|---|---|---|---|
| **Clas SFC / ISIN → Clase de Inversión + moneda + región (+8 atributos)** | **`CLASIFICACION`** cols A:M | col A (ISIN/código) → cols B–L | `Module6.clasificacionPO` |
| **Título nuevo → serie/Clas + moneda** | **`TITULOS NUEVOS`** cols A:D | col A (título) → cols B,C,D | `Module2.Compras` |
| **Título en tránsito → Clase de Moneda** | **`Resumen CV`** cols ~G:H | título → moneda | `Module1.CxC` (T620 col Y) |
| **Moneda → factor de conversión (TRM)** | **`IND2`** cols 1:9 | moneda → col 5 | `Módulo7.Copy` (vector) |
| **Fondos Mutuos / ETF / Notas nuevas → moneda/región** | **`CPRVI`** (+ `TITULOS NUEVOS`, `CLAS ACCIONES`) | manual, del equipo Renta Variable | insumo del paso 6c (fuera de estas macros) |
| **Parse indexador desde código facial** | *hardcodeado* en `Module6` (por longitud) | col 19 → col 20 | `Module6.clasificacionPO` |

---

## 12. Puntos de juicio humano / detección de instrumentos nuevos ⚖️ (para el detector de alertas)

1. **ND en CLASIFICACION** — clave (ISIN, luego Clas SFC) ausente de la hoja `CLASIFICACION` →
   `VLOOKUP` = `#N/A` en cols 13–27 de PORTAFOLIOS (≡ `ND` en col AB del BMO). **Señal primaria del
   detector.** Resolución manual: alta de fila en `CLASIFICACION`.
2. **Compra sin diccionario** — título `"Compra"` de T620 ausente de `T597` **y** de `TITULOS NUEVOS`
   → `#N/A` en cols 3/16/20. Alta manual en `TITULOS NUEVOS`.
3. **Fondos Mutuos / ETF / Notas nuevas** — requieren **diccionario del equipo de Renta Variable** (correo/
   term sheet) para mapear MONEDA y REGIÓN; se agregan a `TITULOS NUEVOS` y `CPRVI`. No hay lógica
   automática; es un insumo externo.
4. **Título en tránsito sin moneda** — ausente de `Resumen CV` → `#N/A` en T620!Y (Clase de Moneda).
5. **Parse de indexador frágil** — el `Select Case Len()` de Module6 puede clasificar mal códigos con
   longitud/estructura no prevista (solo `IBR0` está contemplado como excepción).
6. **CSA** — `MsgBox("Revisar CSA")` cuando la estructura del archivo no coincide.
7. **Códigos especiales hardcodeados** (`COC04CB00137`, `COB51CB00452`, filtro `AODF5 AH BBH USD`) —
   reglas *ad hoc* que rompen si cambia el universo de títulos.

---

## 13. Riesgos para reimplementar en Python

- **Toda la clasificación es VLOOKUP contra hojas-diccionario que crecen manualmente** (`CLASIFICACION`,
  `TITULOS NUEVOS`, `Resumen CV`, `IND2`, `CPRVI`). En Python: tablas/CSV versionados + **cola de
  pendientes** para ND (no simplemente `#N/A` silencioso). El detector de instrumentos nuevos debe
  reproducir la resolución de clave: **ISIN (col8) → nemo (col9) → Clas SFC (col3)** con *fallback* a Clas
  SFC, y marcar como "nuevo" todo lo que no case en ninguna tabla.
- **Semántica de columnas B–L de `CLASIFICACION` no está en el código** (solo índices numéricos); hay que
  leerla del libro (encabezados) para saber cuál columna es Clase de Inversión vs moneda vs región.
- **Parse de moneda por posición fija** (`Mid`/`Len`) — sustituir por un parser robusto de códigos SFC /
  ISIN; el actual es específico del formato de facial de esa fuente.
- **Rutas y parámetros no están en el VBA**, sino en celdas nombradas de `PROCEDIMIENTO`/`CONTROL` → hay
  que volcarlas a configuración.
- **Dos versiones de CSA** con destinos distintos (usar `CSA_balance` 2025 → `Resumen CV!AQ3`).
- **Reglas hardcodeadas** (filtros de string con padding, códigos puntuales, password `INVERSIONES`,
  `COC04CB00137`→46992) deben inventariarse como excepciones explícitas, no perderse.
- **Signos de REPO PASIVO** (`*-1` en columnas específicas) y **normalización de portafolio** (2 vs 3
  chars según 4º carácter `-`) son reglas de negocio a preservar.
- Fórmula de participación `r fija` (COUNTIFS) es un cálculo de prorrateo que debe replicarse.
