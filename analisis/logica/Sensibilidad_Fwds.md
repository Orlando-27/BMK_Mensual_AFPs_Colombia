# Análisis de lógica — `Sensibilidad_Fwds.xlsb` / `Módulo1`

> Fuente: `analisis/vba/Sensibilidad_Fwds__Módulo1.bas.vba` (588 líneas).
> Contexto: `docs/PROCESO_BENCHMARK_MENSUAL.md`, **paso 8i — Control sensibilidad (forwards)**.
> Objetivo del archivo: calcular la **sensibilidad (Key Rate Duration, KRD)** de los *forwards* de
> divisas de la industria (todas las AFP) y del ejercicio individual (Colfondos), separando la
> sensibilidad a la **tasa local (IBR, COP)** y a la **tasa foránea (implícita USD/EUR)**.

---

## 0. Visión general del archivo

**Hojas:** `Control` (parámetros/rutas/botones), `Curvas` (curvas y curva forward con shocks ±100pb),
`572` (ejercicio individual Colfondos, forwards del plano 572, solo USDCOP), `Ind` (forwards de la
industria traídos del Benchmark Detallado, USDCOP + EURCOP), `Ejercicio Individual`.

**Subs del Módulo1 (5):**

| # | Sub | Rol | Hoja destino |
|---|-----|-----|--------------|
| 1 | `ImportarCurvas` | Importa curvas (Fwd points, SWAP IBR, implícitas USD/EUR, indicadores) y construye la **curva forward con shocks ±100pb** | `Curvas` |
| 2 | `Limpiar_Datos_externos` | Borra los nombres `*DatosExternos*` que dejan las QueryTables | (workbook) |
| 3 | `Importar572` | Importa el plano **572** (forwards Colfondos, filtra `US$`/USDCOP) y calcula valoración + KRD | `572` |
| 4 | `Actualizar_Industria` | Abre `Benchmark (Detallado)`, copia hoja `Fwd Industria`, y llama a `CalcularInd` | `Ind` |
| 5 | `CalcularInd` | Motor de valoración y KRD de los forwards de la **industria** (USDCOP + EURCOP) | `Ind` |

**Mapa de botones (inferido del proceso, paso 8i):**

| Botón | Macro | Acción |
|-------|-------|--------|
| **1. Importar curvas** | `ImportarCurvas` | Trae insumos de mercado y arma la curva forward + escenarios de shock. |
| **2. (Ejercicio individual)** | `Importar572` → `CalcularInd` | Trae forwards de Colfondos (plano 572) y calcula su KRD. |
| **3. Traer industria y calcular** | `Actualizar_Industria` → `CalcularInd` | Trae forwards de la industria del Benchmark Detallado y calcula su KRD. Es el botón citado en paso 8i. |

**Integración externa:** tras correr el botón 3, en `Benchmark (Detallado)` celda **F11** se corre
`Sub traer_sensibilidad` (Módulo4 de ese archivo), que recoge los KRD calculados aquí.

**Cadena de dependencia:** botón 1 (curvas) **debe** correr antes que 2 y 3, porque las valoraciones
de `572`/`Ind` hacen `VLOOKUP` contra las columnas de la hoja `Curvas`.

---

## 1. `ImportarCurvas()`  (botón 1)

### Propósito
Cargar las curvas de mercado y **reconstruir la curva de puntos forward** (USDCOP y EURCOP) más los
cuatro escenarios de shock (±100 pb local / ±100 pb foráneo). Estas columnas son la base de todos los
cálculos de sensibilidad posteriores.

### Entradas
Rutas y nombres de archivo se leen de **named ranges** (celdas de `Control`), no están hardcodeadas:
`Ruta_Inf`, `Ruta_Indica`, `Pt_Fwd`, `Pt_FwdEUR`, `SWP_IBR`, `SWP_Impli`, `SWP_ImpliEUR`,
`Indicadores`. Se concatenan `Ruta & Nombre` para formar la ruta completa (todas apuntan a `M:\...`,
ver §7).

Archivos externos importados (todos vía `QueryTables.Add "TEXT;"...` = archivos planos):

| Archivo (named range) | Destino en `Curvas` | Contenido |
|---|---|---|
| `Pt_Fwd` (Puntos Fwd USDCOP) | `A:B` | A=plazo, B=puntos fwd |
| `SWP_IBR` (Curva SWAP IBR) | `D:E` | D=plazo, E=tasa **local COP** |
| `SWP_Impli` (Implícita USDCOP) | `G:H` | G=plazo, H=tasa **foránea USD** |
| `Indicadores` | `R:Z` (desde `R1`) | indicadores del día (incluye TRM, EUR) — delimitado `;`, code page 1252 |
| `Pt_FwdEUR` (Puntos Fwd EURCOP) | `AC:AD` | AC=plazo, AD=puntos |
| `SWP_ImpliEUR` (Implícita EURCOP) | `AF:AG` | se **abre como workbook** (`Workbooks.Open`), `TextToColumns` por comas, copia `G2:H2↓`, pega valores |

Named cells de mercado usados por las fórmulas: **`TRM`** (spot USDCOP), **`EUR`** (spot EURCOP),
**`Fecha`** (fecha del benchmark). Provienen de los indicadores / `Control`.

### Salidas — hoja `Curvas`
Se limpia `A:AZ` al inicio. Layout resultante (C, F, I = separadores en blanco):

**Bloque USDCOP (`J:P`), rótulos en fila 1, fórmulas desde fila 2 hasta `rango`:**

| Col | Rótulo | Fórmula R1C1 | Interpretación |
|-----|--------|--------------|----------------|
| J | Tasa FWD | `=TRM*((1+E·D/360)/(1+H·G/360))` | Fwd = TRM·(1+r_loc·t/360)/(1+r_for·t/360) |
| K | Puntos Fwd | `=J-TRM` | puntos forward reconstruidos |
| L | Loc +100pb | `=TRM*((1+(1%+E)·D/360)/(1+H·G/360))-TRM` | puntos con IBR (local) +100pb |
| M | Loc -100pb | `=TRM*((1+(-1%+E)·D/360)/(1+H·G/360))-TRM` | puntos con IBR local -100pb |
| N | For +100pb | `=TRM*((1+E·D/360)/(1+(1%+H)·G/360))-TRM` | puntos con implícita (foránea) +100pb |
| O | For -100pb | `=TRM*((1+E·D/360)/(1+(-1%+H)·G/360))-TRM` | puntos con implícita foránea -100pb |
| P | Validador | `=ROUND(K,5)=ROUND(B,5)` | verifica que el fwd reconstruido = fwd observado (col B) |

> Notación: E=tasa local (col5), D=plazo local (col4), H=tasa foránea (col8), G=plazo foráneo (col7).
> Los offsets R1C1 fueron verificados en el fuente (`RC[-7]`,`RC[-8]` en L, etc.) — son consistentes.

**Bloque EURCOP (`AI:AO`)** análogo, usando `EUR` en lugar de `TRM`, puntos EUR en `AC:AD` e implícita
EUR en `AF:AG`. La tasa local sigue siendo la IBR (`D:E`). Validador `AO` a 3 decimales.

`rango = MIN(últimaFila(A), últimaFila(D), últimaFila(G))`; `rangoEUR = MIN(col29, col4, col32)`.
Las fórmulas se pegan como valores (`xlPasteFormulas`→`Calculate`→`xlPasteValues`).

Al final llama a `Limpiar_Datos_externos` y selecciona `Control`.

### Modelo de cálculo
Paridad cubierta de tasas de interés (forward FX). Los cuatro escenarios shock ±100pb sobre la tasa
local (IBR) o foránea (implícita) generan las curvas de puntos forward que luego se usan para revaluar
cada operación y estimar la KRD por diferencias centrales.

### Controles / validaciones
- Columna **P / AO = Validador**: `VERDADERO` si el forward reconstruido reproduce el observado.
- `rango`/`rangoEUR` = mínimo de longitudes para no calcular sobre filas sin datos.

---

## 2. `Limpiar_Datos_externos()`
Recorre `ThisWorkbook.Names` y borra todo nombre que contenga `"DatosExternos"` (`InStr`). Limpieza de
los rangos con nombre que crean las `QueryTables`. Sin entradas/salidas de datos. Reimplementación en
Python: no aplica (artefacto de Excel).

---

## 3. `Importar572()`  (botón 2 — ejercicio individual Colfondos)

### Propósito
Traer los forwards de Colfondos desde el **plano 572** (formato regulatorio), filtrar solo **USDCOP**
(`US$`), y calcular valoración de patas + KRD por operación.

### Entradas
- Named ranges `Ruta_572`, `Plano572` → `File_572 = Ruta_572 & Plano572` (archivo plano `M:\...`).
- Named cells `Fecha`, `TRM`. Curvas de la hoja `Curvas` (dependencia del botón 1).

### Proceso
1. Limpia hoja `572`. Crea workbook nuevo, importa el plano vía `QueryTable` (delimitado `;`+Tab).
2. Borra columnas `AA:AD`. **Filtra columna 12 (`L`) por `"US$   "`** (con espacios) y elimina lo demás
   → deja solo forwards USDCOP.
3. Copia las primeras ~29 columnas (A a AC) y pega en hoja `572` del archivo Sensibilidad.
4. Escribe rótulos y fórmulas `AB3:AQ3`, las arrastra hasta `rango572` y pega como valores.
5. `Range("AC3").Select : Call CalcularInd` — **ojo:** `CalcularInd` opera sobre la hoja `Ind`, no
   sobre `572`. La llamada al final refresca la industria (o es residual); ver §7 Riesgos.

### Salidas — hoja `572`, columnas `AB:AQ`
| Col | Campo | Fórmula (resumen) |
|-----|-------|-------------------|
| AB | Fondo | `MID(RC[-4],2,2)` |
| AC | Tipo Operación | `TRIM(...)` (Compra/Venta) |
| AD | Plazo (días) | `= venc - Fecha` |
| AE | Strike | `= strike` |
| AF | **Valoración Derecho** | Compra: `(TRM+VLOOKUP(plazo,Curvas!G:K,5,0))/(1+r_loc·plazo/360)`; Venta: `strike/(1+r_loc·plazo/360)` |
| AG | **Valoración Obligación** | Venta: `(TRM+fwd_base)/(1+r_loc·…)`; Compra: `strike/(1+r_loc·…)` |
| AH/AI | Sube IBR Derecho/Oblig | usa `Curvas!G:O col6` (L=Loc+100) **y** discount `(1%+r_loc)` |
| AJ/AK | Baja IBR Derecho/Oblig | `col7` (M=Loc-100) y discount `(-1%+r_loc)` |
| AL/AM | Sube Impli Derecho/Oblig | `col8` (N=For+100), discount con `r_loc` **sin** shockear |
| AN/AO | Baja Impli Derecho/Oblig | `col9` (O=For-100), discount con `r_loc` sin shockear |
| AP | **KRD Local** | Compra: `-((AK-AI)/(2·1%·AG))`; Venta: `((AJ-AH)/(2·1%·AF))` |
| AQ | **KRD Foran** | Venta: `-((AO-AM)/(2·1%·AG))`; Compra: `((AN-AL)/(2·1%·AF))` |

`r_loc = VLOOKUP(plazo, Curvas!D:E, 2, 0)` (tasa IBR); `fwd_base = VLOOKUP(plazo, Curvas!G:K, 5, 0)`.

---

## 4. `Actualizar_Industria()`  (botón 3 — traer industria)

### Propósito
Traer los forwards de la **industria** desde `Benchmark (Detallado)` y disparar su cálculo de KRD.

### Entradas
- Named ranges `Ruta_Ind`, `BMK` → `FileBMK = Ruta_Ind & BMK` (ruta del `Benchmark (Detallado).xlsb`).
- Abre `FileBMK`, hoja **`Fwd Industria`**, copia **`A23:H23` hacia abajo** (8 columnas: A–H).

### Proceso
1. Limpia `Ind!A3:` hacia abajo. 2. Abre el Benchmark Detallado, copia `Fwd Industria` A23↓.
3. Pega valores+formatos en `Ind!A3`. 4. Cierra el Benchmark **sin guardar**. 5. `Call CalcularInd`.

### Salidas
Hoja `Ind` columnas `A:H` (datos crudos de industria) y luego `L:AA` (cálculo, ver §5).

### Controles
Cierra `Workbooks(BMK).Close SaveChanges:=False` (no altera el Benchmark). Depende del botón 1 (curvas).

---

## 5. `CalcularInd()`  (motor de KRD de la industria — multi-moneda)

### Propósito
Valorar cada forward de la industria y calcular su **KRD Local** y **KRD Foran**, soportando **USDCOP
y EURCOP**.

### Entradas
- Hoja `Ind` (datos de `Actualizar_Industria`, o de `Importar572` si se llamó desde ahí).
- Named cells `Fecha`, `TRM`, `EUR`. Curvas de la hoja `Curvas` (USD `G:O`/`D:E`; EUR `AF:AN`).
- `rangoIND = últimaFila(Ind col A)`.

### Salidas — hoja `Ind`, columnas `J` y `L:AA` (se limpia `L:AA` antes)
| Col | Campo | Fórmula (resumen) |
|-----|-------|-------------------|
| J | Venc. ajustado | `=IF(venc-Fecha<=0, Fecha+31, venc)` (si venció, +31 días) |
| L | Fondo | `=RC[-6]` |
| M | Tipo Operación | `=IF("SELL","Venta",IF("Buy","Compra","Error"))` |
| N | Plazo (días) | `= venc - Fecha` |
| O | Strike | `=RC[-10]` |
| P–Y | Valoraciones (base / IBR± / Impli±, Derecho y Obligación) | igual patrón que `572`, pero con `IF(moneda="EURCOP", EUR & Curvas!AF..AN, TRM & Curvas!G..O)` |
| Z | **KRD Local** | `=IF(OR(moneda=USDCOP,EURCOP), IF(Compra,-((U-W)/(2·1%·AA_valOblig)), IF(Venta,(...)/…)), "")` |
| AA | **KRD Foran** | análogo con las columnas de shock foráneo |

`Z`/`AA` solo se calculan si la moneda es `USDCOP` o `EURCOP` (los demás → vacío).
Selección de curva por moneda: EURCOP usa `EUR` + `Curvas!C32:C36`/`C32:C40`; resto usa `TRM` +
`Curvas!C7:C11`/`C7:C15`. La tasa de descuento local (IBR, `Curvas!C4:C5`) es común a ambas monedas.

### Modelo de cálculo (idéntico a `572`, generalizado a EUR)
Por operación se revalúa la pata relevante (Derecho para Venta / Obligación para Compra, en la tasa
local; e inverso para la foránea) bajo shocks ±100pb, y la KRD es la **diferencia central**
normalizada:

```
KRD ≈ -(V(+100pb) - V(-100pb)) / (2 · 0.01 · V_base)
```

- **KRD Local**: shock a la IBR (COP) → afecta puntos forward **y** factor de descuento local.
- **KRD Foran**: shock a la implícita (USD/EUR) → afecta **solo** los puntos forward (descuento local
  sin cambiar).
- La operación **Compra** toma la sensibilidad local por la pata **Obligación** (paga COP) y la foránea
  por la pata **Derecho**; la **Venta** al revés.

---

## 6. Relación con "sensibilidad relativa" (contexto paso 8i)
El manual indica: para **renta fija** la sensibilidad relativa usa **diferencia de sensibilidades**
(KRD), y para los **demás activos** usa **diferencia en posiciones**. Este módulo produce el insumo de
sensibilidad de forwards (KRD Local/Foran por operación) que luego consume `traer_sensibilidad`
(Módulo4 del Benchmark Detallado) para armar la diferencia industria vs. posición.

---

## 7. Rutas hardcodeadas y patrones de fecha

- **No hay rutas `M:\...` literales dentro del VBA.** Todas se leen de named ranges en `Control`:
  `Ruta_Inf`, `Ruta_Indica`, `Ruta_572`, `Ruta_Ind` (carpetas) + `Pt_Fwd`, `Pt_FwdEUR`, `SWP_IBR`,
  `SWP_Impli`, `SWP_ImpliEUR`, `Indicadores`, `Plano572`, `BMK` (nombres de archivo). La ruta física
  `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\...` vive en las celdas de `Control`, no en el
  código → para reimplementar hay que **volcar esos valores de celda**, no basta el VBA.
- **Patrones de fecha:** todo el fechado pasa por el named cell **`Fecha`** (fecha del benchmark /
  último BMK diario). El plazo siempre es `venc - Fecha` (base **360**). Regla de vencidos en `Ind`:
  si `venc - Fecha <= 0` ⇒ `Fecha + 31`. El nombre del archivo del mes (patrón `AAAA_MMportainvdeta`,
  `0` en meses de 1 dígito) se resuelve fuera de este módulo (en el named range `Plano572`/`BMK`).
- Code pages: `850` para planos de curvas/fwd; `1252` para `Indicadores`.

---

## 8. Riesgos para reimplementar en Python

1. **Curvas como dependencia oculta:** `572` e `Ind` no calculan nada sin la hoja `Curvas` poblada por
   el botón 1. En Python hay que construir primero la curva forward + 4 escenarios de shock y usarla
   como tabla de lookup por plazo.
2. **VLOOKUP por plazo con coincidencia exacta (`,0`)**: los planos deben traer plazos que existan en la
   curva. Si el plazo de la operación no está en `Curvas`, `VLOOKUP` da `#N/A`. Replicar con
   interpolación explícita o merge exacto + control de faltantes.
3. **Base 360 y convención de descuento simple** `(1 + r·t/360)` (no compuesto). Mantener.
4. **Semántica Derecho/Obligación por Compra/Venta cruzada** (local en una pata, foráneo en la otra).
   Fácil de invertir el signo; la KRD lleva un `-1*` en una de las ramas.
5. **Filtro USDCOP frágil**: `Criteria1:="US$   "` incluye **espacios de padding**; `Ind` filtra por
   texto `"USDCOP"`/`"EURCOP"` y mapea `SELL/Buy`. Normalizar strings (trim, mayúsculas) al portar.
6. **`Importar572` llama a `CalcularInd`** que opera sobre `Ind`, no sobre `572` — posible efecto
   colateral / residual. En Python separar claramente el pipeline individual (572) del de industria (Ind).
7. **Regla de vencidos `Fecha+31`** (columna J de `Ind`) es un parche que altera el plazo de forwards
   ya vencidos; decidir si se mantiene o se excluyen esas filas.
8. **Named cells `TRM`, `EUR`, `Fecha`** deben extraerse del archivo (no están en el VBA); dependen de
   la importación de `Indicadores`.
9. **Escala/signo de KRD** sigue la convención interna del área (no es duración modificada estándar);
   validar contra `traer_sensibilidad` del Benchmark Detallado antes de dar por buena la cifra.
10. **Multi-moneda parcial:** el bloque EUR de `Curvas` abre otro workbook (`SwapCC2_EURCOP`) — origen
    de datos distinto (no plano TEXT), a considerar en la ingesta.
