# Calculadora Swap — Análisis del modelo de valoración (VPN y DM)

> Fuente: módulos VBA `analisis/vba/Calculadora Swap__Módulo1..7.bas.vba` + estructura de celdas
> extraída de `insumos/Calculadora Swap.xlsb` (pyxlsb, solo valores). Autor original del libro:
> José Manuel Ibáñez Delgado ("Ingeniería Financiera — Calculadora Swap 2016-1").
> Objetivo: reimplementar en Python el cálculo de **VPN** (Valor Presente Neto) y **DM** (Duración
> Modificada) de swaps CCS/IRS del benchmark mensual de AFPs (Colombia). Paso 7 del proceso.

---

## 0. Hallazgo central (leer antes que nada)

**El VBA NO contiene el modelo de valoración.** Los 7 módulos son solo **orquestación**: importan
curvas desde archivos `.txt`/`.csv` de la unidad `M:\`, mueven insumos entre hojas, disparan
`Calculate` y copian resultados (VPN/DM) de vuelta. **El modelo financiero vive íntegramente en las
fórmulas de celda** de las hojas de valoración `SWAPCC_*` / `SWAPIRS_*` (columnas D→O de la tabla de
flujos, y celdas resumen B15/B16/E15/E16). Para reimplementar en Python hay que **replicar esas
fórmulas de celda**, decodificadas aquí a partir de la hoja `RUTAS` (que documenta las fórmulas como
texto) y de los valores computados de las hojas.

Arquitectura de ejecución (una fecha de valoración):

```
vpn_1  (Módulo2)  ─┬─ EXPORTARDATOS   (Módulo1) → importa curvas M:\ → hoja DATOS/MONEDA
                   ├─ CALCULARVPN     (Módulo3) → itera swaps de SWAPIND, valora en hoja motor, recoge VPN/DM
                   ├─ PEGADO          (Módulo4) → consolida columnas en hoja INICIO
                   └─ Limpiar_Datos_externos (Módulo2) → borra conexiones QueryTable
GUARDAR (Módulo4) → guarda 2 archivos: "VPN SWAP dd-mm-aaaa.xlsb" y "DM SWAP dd-mm-aaaa.xlsb"
```

---

## 1. Módulos VBA — Sub/Function por Sub/Function

### Módulo2 — orquestador principal

#### `vpn_1()` — MACRO PRINCIPAL (botón)
- **Propósito:** correr el ciclo completo de valoración para **una** fecha (la de `Fecha_valoracion`).
- **Entradas:** hoja INICIO/RUTAS (fechas, rutas); dispara todo el resto.
- **Flujo:** `EXPORTARDATOS` → `Calculate` → `CALCULARVPN` → `PEGADO` → `Calculate` →
  `Limpiar_Datos_externos`. Cronometra y escribe segundos en `RUTAS!I25`.
- **Salidas:** deja VPN/DM poblados en SWAPIND e INICIO (no guarda archivo; eso lo hace `GUARDAR`).
- Nota: contiene comentada `ActiveCell.FormulaR1C1="=+TODAY()-85"` (viejo hardcode de fecha).

#### `Limpiar_Datos_externos()`
- Borra todos los `Names` del libro que contengan `"DatosExternos"` (conexiones QueryTable creadas al
  importar los `.txt`). Higiene, sin efecto en el cálculo.

---

### Módulo1 — importación de curvas (611+ líneas, ~13 subs casi idénticas)

Patrón común de cada `EXPORTAR_*`: lee una **ruta base** de un named range, arma el nombre de archivo
concatenando **`AAAA` + `MM` + `DD`** de `Fecha_valoracion`, limpia un bloque de la hoja `DATOS` y lo
re-importa con `QueryTables.Add` (TEXT, tab/space-delimited, `.txt`). Cada curva cae en un par de
columnas de `DATOS` (plazo en días, tasa).

#### `EXPORTARDATOS()` — dispatcher
Llama en orden: `EXPORTAR_IBR`, `EXPORTAR_IBRUVR`, `EXPORTAR_DTF`, `EXPORTAR_USDOIS`,
`EXPORTAR_LIBORUVR`, `EXPORTAR_LIBORCOP`, `EXPORTAR_USDCO`, `EXPORTAR_MONEDA`, `EXPORTAR_EURCOP`,
`EXPORTAR_CCS_COLATERAL`. (Comentados/inactivos: `EXPORTAR_IPC`, `EXPORTAR_USDIBR`,
`EXPORTAR_LIBORUSD3M`.)

| Sub | Named range ruta | Destino en `DATOS` | Curva |
|---|---|---|---|
| `EXPORTAR_IBR` | `ruta_ibr` (IBR_COLATERAL) | D:E | IBR diaria colateral |
| `EXPORTAR_IBRUVR` | `ruta_ibruvr` | G:H | IBR-UVR |
| `EXPORTAR_DTF` | `ruta_dtf` | J:K | DTF |
| `EXPORTAR_USDOIS` | `ruta_usdois` | M:N | USD OIS (SOFR/OIS) |
| `EXPORTAR_LIBORUVR` | `ruta_liboruvr` | S:T | LIBOR-UVR |
| `EXPORTAR_LIBORCOP` | `ruta_liborcop` | V:W | LIBOR-COP |
| `EXPORTAR_USDCO` | `ruta_usdco` | Y:Z | USD-CO |
| `EXPORTAR_USDIBR` (inactivo) | `ruta_usdibr` | AB:AC | USD-IBR (`.csv`) |
| `EXPORTAR_LIBORUSD3M` (inactivo) | `ruta_liborusd3M` | AR:AS | LIBOR USD 3M (Bloomberg) |
| `EXPORTAR_CCS_COLATERAL` | `CCS_COLATERAL` | AZ:BA | COP colateral USD |
| `EXPORTAR_IPC` (inactivo) | `ruta_ipc` | P:Q | IPC |
| `EXPORTAR_MONEDA` | `ruta_moneda` | hoja `MONEDA` A:U | Indicadores (TRM, UVR, EUR…) `.CSV` |
| `EXPORTAR_EURCOP` | `ruta_eurcop` | AE:AF | EUR-COP (abre `.csv`, TextToColumns, pega valores) |

Detalles:
- **`EXPORTAR_MONEDA`**: nombre de archivo distinto → `<ruta>\AAAA-MM-DD IND.CSV` (delimitador `;`).
- **`EXPORTAR_EURCOP`**: no usa QueryTable; abre el `.csv` `SwapCC2_EURCOP_Diaria_AAAAMMDD.csv`, separa
  por comas, copia G:H, pega valores en `DATOS!AE2` y cierra sin guardar.
- Estas rutas son **dependencia externa dura**: sin los `.txt` diarios de INFOVALMER no hay curvas.

---

### Módulo3 — `CALCULARVPN()` (motor de iteración)

- **Propósito:** para cada swap listado en `SWAPIND`, colocarlo en la hoja-motor correspondiente,
  recalcular y **recoger VPN y DM** de derecho y obligación.
- **Lógica:**
  1. Cuenta filas de `SWAPIND` (col A). Itera `i = 1..UFV-3`.
  2. Toma el **tipo de swap** de la fila (`SWAPIND!B1` = `A3.Offset(i,0)`), y `C1` se recalcula a la
     **hoja-motor destino** (`hoja = C1.Value`) según la tabla de mapeo tipo→hoja.
  3. Copia el identificador/parámetros del swap (`A3.Offset(i,1)`) y lo **pega en `<hoja>!B4`**
     (celda de entrada del motor). Luego `Range("A1:AO150").Calculate`.
  4. Recoge resultados de la hoja-motor y los pega (valores) en `SWAPIND`:
     - `B15` (VP derecho) → `SWAPIND` col offset **+13** (columna N).
     - `B16` (DM derecho) → offset **+45** (columna AT).
     - `E15` (VP obligación) → offset **+21** (columna V).
     - `E16` (DM obligación) → offset **+46** (columna AU).
  5. Escribe tiempo en `RUTAS!I24`.
- **Mapeo tipo de swap → hoja-motor** (hoja `RUTAS` F23:G33):

  | Tipo (SWAPIND) | Hoja-motor |
  |---|---|
  | `IRS IBRCOP` | `SWAPIRS_IBRoncop` |
  | `CCS COPUSD` | `SWAPCC_COPUSD` |
  | `CCS IBRLIB` | `SWAPCC_IBRLIBOR` |
  | `CCS COPUVR` | `SWAPCC_COPUVR` |
  | `CCS COPEUR` / `IRS IBR+PB` | `SWAPIRS_IBRME` |
  | `IRS SOFUSD` | `SWAPIRS_SOF` |
  | `IRS LIBORUSD` | `SWAPIRS_LIBOR` |
  | `CCS UVREUR` / `CCS UVRUSD` | `SWAPCC_COPUVR` |
- **Nota:** ~380 líneas del módulo son una versión antigua `Select Case` **comentada** (misma lógica
  con `Case Is = "CCS COPUVR"`, etc.). La versión viva usa la tabla `C1` en vez del `Select Case`.

---

### Módulo4

#### `PEGADO()`
Consolida columnas de `SWAPIND` (todas las filas) en la hoja `INICIO` (a partir de fila 26), pegando
solo valores:
- `SWAPIND!B4↓` (# swap) → `INICIO!B26`
- `SWAPIND!N4↓` (VP derecho) → `INICIO!C26`
- `SWAPIND!V4↓` (VP obligación) → `INICIO!D26`
- `SWAPIND!W4↓` (VPN neto) → `INICIO!E26`
- `SWAPIND!AT4↓` (DM derecho) → `INICIO!W26`
- `SWAPIND!AU4↓` (DM obligación) → `INICIO!Y26`

#### `GUARDAR()` — genera los archivos entregables
- Lee named ranges: `RUTA_GUARDAR` (=`INICIO!S1`), `N_GUARDAR` (=`INICIO!S2`), `N_GUARDAR_DM`
  (=`INICIO!S4`).
- **Precios (VPN):** copia `INICIO!G26:J26↓`, crea libro nuevo, pega valores en `A2`, guarda como
  `RUTA_GUARDAR\<N_GUARDAR>.xlsb` (xlExcel12).
- **Duración modificada (DM):** copia `INICIO!V26:Y26↓`, libro nuevo, pega en `A2`, guarda como
  `RUTA_GUARDAR\<N_GUARDAR_DM>.xlsb`.
- Con los valores actuales del libro:
  - `RUTA_GUARDAR` = `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\BDatos\Backup Archivos\2026\SWAP`
  - `N_GUARDAR` = **`VPN SWAP 09-07-2026`**, `N_GUARDAR_DM` = **`DM SWAP 09-07-2026`**
  - Patrón de nombre: `VPN SWAP dd-mm-aaaa` y `DM SWAP dd-mm-aaaa`.

---

### Módulos 5, 6, 7 — utilidades menores (grabadas)
- **Módulo5 `Macro1()`**: fuerza `Calculate` columna por columna (f,A,K,S,W,Z,AA..AD). Workaround de
  recálculo. Sin lógica financiera.
- **Módulo6 `Macro2()`**: AutoFilter sobre `$A$3:$AD$101`, campo 8 = `"IBR"`. Utilidad de filtrado.
- **Módulo7 `Macro3()`**: TextToColumns de un CSV, copia G:H y pega en `DATOS!AE2` de
  `Calculadora Swap.xlsm` (variante manual de la importación EUR-COP).

---

## 2. MODELO DE VALORACIÓN (reconstruido de las hojas motor)

Cada tipo de swap tiene su hoja motor (`SWAPCC_*`, `SWAPIRS_*`) con **estructura idéntica**: dos
patas espejo, **DERECHO** (columnas A–O) y **OBLIGACIÓN** (columnas P–AD). Un solo swap por hoja
(entrada en `B4`).

### 2.1 Cabecera del motor (celdas de parámetros, filas 1–18)

| Celda | Significado |
|---|---|
| `B1` | Fecha de valoración (serial Excel) |
| `B4/E4` | N° de contrato |
| `B5/E5` | Frecuencia (`AV`/`SV`/`TV`/`MV` = anual/semestral/trimestral/mensual vencido) |
| `B6/E6` | Meses entre cupones (12/6/3/1) |
| `B7/E7` | Fecha cupón previo |
| `B8/E8` | Fecha vencimiento |
| `B9/E9` | Base derecho/obligación (day-count: `ACT/365`, `ACT/360`, `30/360`) |
| `B10/E10`,`B11/E11` | Base días flujos (`ACT`) y base años (360/365) |
| `B12/E12` | Tasa fija de la pata (cupón); `0` si es flotante |
| `B13/E13` | Moneda de la pata (COP/USD/UVR/EUR) |
| `B14/E14` | Nominal |
| `B15/E15` | **VP DERECHO / VP OBLIGACIÓN** = Σ(Valor Presente) → **VPN de cada pata** |
| `B16/E16` | **DURACIÓN MODIFICADA** derecho / obligación |
| `B17/E17` | Convexidad |
| `B18/E18` | Tipo de tasa (`FS`=fija, `IBR`, `LIBOR`, `SOF1`, …) |
| `J10/J12/J13` | Bloque de control: PyG 30, comparación Calculadora vs SFC |

### 2.2 Tabla de flujos (filas 20+, una fila por cupón)

Columnas de la pata DERECHO (P–AD replican para OBLIGACIÓN):

| Col | Campo | Cálculo |
|---|---|---|
| D | Fecha del flujo | Cronograma de cupones desde `FECHA CUPON PREVIO` a `VENCIMIENTO` (paso = frecuencia) |
| E | Días entre flujos | `D(n) − D(n−1)` (base ACT) |
| F | FRA / tasa del período | tasa fija (`B12`) si pata fija; **tasa forward** de la curva si flotante |
| G | Flujo capital en % | 0 salvo intercambio de nocional (CCS: en vencimiento) |
| H | Flujo intereses en $ | `Nominal × F × (E / BaseAños)` (o factor 30/360 según base) |
| I | Flujo capital en $ | Nominal en fecha de intercambio (CCS al vencimiento; IRS = 0) |
| J | Días al vencimiento | `D(n) − FechaValoración` |
| K | **Tasa de descuento** | **interpolada de la curva** (hoja `DATOS`) por plazo `J` |
| L | **Factor de descuento** | `L = (1 + K)^(−J / BaseAños)` (capitalización compuesta anual) |
| M | **Valor presente** | `M = (H + I) × L` |
| N | Aporte a duración | `≈ (J/BaseAños) × M` (tiempo × VP; base para Macaulay) |
| O | Aporte a convexidad | `≈ (J/BaseAños)(J/BaseAños+1) × M / (1+K)^2` |

**Verificación numérica (COPUVR, fila 22):** `K=0.1142`, `J=51`, base 365 →
`L=(1.1142)^(−51/365)=0.9841` ✓; `H=99,446,103`, `I=0` → `M=97,863,177` ✓.

### 2.3 Fórmulas resumen

- **VPN de la pata** (`B15`/`E15`): `VP = Σ_n M_n = Σ_n (H_n + I_n) · (1+K_n)^(−J_n/base)`.
- **VPN neto del swap** = `VP_DERECHO − VP_OBLIGACIÓN` (se calcula en `SWAPIND!W` / `INICIO!E`).
- **Duración modificada** (`B16`/`E16`):
  `Macaulay = Σ_n (t_n · M_n) / Σ_n M_n` con `t_n = J_n/base`; luego
  `DM = Macaulay / (1 + y)` (`y` = tasa/rendimiento de la pata). Es la duración **de cada pata por
  separado** (derecho y obligación), no una DM neta.
- **Convexidad** (`B17`/`E17`): `Σ O_n / Σ M_n`.

### 2.4 Curvas de descuento por tipo de swap

Todas las tasas de descuento `K` y las tasas forward `F` salen de la hoja `DATOS` (interpolación por
plazo). Mapeo curva↔pata (de `RUTAS` I:N y `DATOS`):

| Tipo swap | Pata derecho (curva) | Pata obligación (curva) |
|---|---|---|
| `CCS COPUSD` | COP: DTF/IBR-COP · USD: LIBOR/USD-CO | según patas FS/LIBOR/IBR |
| `CCS COPUVR` | COP (IBR/DTF) | UVR (IBR-UVR / LIBOR-UVR) |
| `CCS IBRLIB` | IBR-COP | LIBOR-USD |
| `IRS IBRCOP` | IBR-COP colateral | Fija COP |
| `IRS IBR+PB` | IBR + spread | Fija |
| `IRS LIBORUSD` | LIBOR-USD | Fija USD |
| `IRS SOFUSD` | SOFR/USD-OIS (col M:N) | Fija USD |

Curvas físicas en `DATOS` (pares columna plazo/tasa): D:E IBR-colateral · G:H IBR-UVR · J:K DTF ·
M:N USD-OIS(SOFR) · S:T LIBOR-UVR · V:W LIBOR-COP · Y:Z USD-CO · AB:AC USD-IBR · AE:AF EUR-COP ·
AR:AS LIBOR-USD-3M · AZ:BA COP-colateral-USD · AU UVR proyectada. Hoja `MONEDA`: TRM/UVR/EUR spot.

### 2.5 Convenciones (de `RUTAS`, fórmulas documentadas como texto)

- **Clasificación del tipo de swap** (`RUTAS!A22`): decide `IRS IBR+PB` / `IRS IBRCOP` /
  `IRS LIBORUSD` / `CCS COPUVR` / `CCS COPUSD` / `REVISAR` según (tasa, moneda) de cada pata (H/I vs
  P/Q, prefijos `IBR o/n`, `IBR o/n+`, `LIBOR 3M`, etc.).
- **Frecuencia → periodo** (`A28`/`A32`): `Trimestral→TV`, `semestral→SV`, `anual→AV`, `mensual→MV`.
- **Fechas de cupón** (`A38`/`A42`): función Excel **`CUPON.FECHA.L1` (=`COUPNCD`)** con
  `frecuencia` (AV=1, SV=2, TV=4) y `base` (ACT/ACT=1, ACT/360=2, ACT/365=3, 30/360=4).
- **Base day-count por combinación** (`A45`/`A54`): tabla condicional según tipo de swap y monedas
  (ej. `CCS COPUSD` COP/USD → derecho `30/360`; USD/COP → `ACT/365`; `IRS IBRCOP` → `ACT/360`;
  `IRS LIBORUSD` USD/USD → `ACT/360`). Derecho y obligación pueden tener bases distintas.
- **SOFR compuesto** (hoja `SWAPIRS_SOF`, bloque H:Q): calcula tasa compuesta overnight (dias
  acumulados, FD, "SOFR acumulado", "SOFR próximo cupón"). La pata SOFR flotante se proyecta con la
  curva USD-OIS y capitalización compuesta diaria.
- **IBR overnight compuesto / LIBOR** (hoja `LIB-IBR on`): serie histórica IBR y LIBOR diaria +
  columna "Tasa Compuesta" para la pata flotante ya devengada del cupón corriente.
- **Excepciones Non Delivery** (`RUTAS!T2`, `SWAPIRS_SOF!I5`): flag para swaps non-delivery.

---

## 3. Entradas / salidas (resumen operativo)

- **Fecha de valoración:** named range `Fecha_valoracion` (celda en INICIO/RUTAS). El VBA corre **una**
  fecha por ejecución; las **3 fechas** del paso 7 (día previo al cierre, cierre de mes, día previo a la
  publicación p.ej. 09) se corren **manualmente** cambiando la fecha y volviendo a ejecutar `vpn_1` +
  `GUARDAR`. El "09" reemplaza la Calculadora Swap del proceso diario.
- **Insumos externos (obligatorios):** archivos diarios INFOVALMER en `M:\...\BDatos\INFOVALMER\` (un
  `.txt`/`.csv` por curva, sufijo `AAAAMMDD`) + `INDICADORES\AAAA-MM-DD IND.CSV` (monedas).
- **Salidas:** hojas `SWAPIND` e `INICIO` pobladas con VPN (der/obl/neto) y DM (der/obl) por swap;
  y dos archivos `.xlsb`: **`VPN SWAP dd-mm-aaaa.xlsb`** y **`DM SWAP dd-mm-aaaa.xlsb`** en
  `M:\...\BDatos\Backup Archivos\AAAA\SWAP` (y copia en carpeta mensual). Estos alimentan la hoja
  `VECTOR DE PRECIOS` del Benchmark Detallado (curva de precios de Swap y curva de duraciones).

---

## 4. Controles / validaciones

- `INICIO!K:L` panel de control: `Valor <>0` OK · `Valor = #N/A` OK · `# SWAP` True ·
  DERECHO/OBLIGACIÓN/VPN True · `LIB-IBR` True · `DATE IND` True. Todos deben ser True.
- `INICIO!N:O`: `FESTIVOS`=True, control de `FECHAS`.
- Cada hoja motor tiene bloque **Calculadora vs SFC** (`H10/H11/H12`, `J10`): compara VP calculado vs
  el reportado por SFC y muestra Diferencia y Dif %. Sirve para validar el modelo swap por swap.
- `INICIO!L24` en adelante: `PRECIOS VECTOR DE PRECIOS` e `PRECIOS VECTOR DE DUR.MODIF` con columnas
  `Diferencia DER/OBL` para cotejar contra el vector.

---

## 5. Puntos de juicio humano

- **Swaps a mantener vs borrar** (paso 7.2): en `SWAPIND` se **borran todos los swaps excepto los de
  Colfondos de pensión voluntaria** (no vienen en el reporte SFC) y los creados **después de la fecha
  de corte** (compras hasta el día 10). Filtrado manual por portafolio/fecha de creación.
- **`REVISAR`**: si `RUTAS!A22` no clasifica el par (tasa,moneda), marca `REVISAR` → intervención.
- **UVR día no hábil** (`INICIO!C4/D4`) y **UVR día 15 vigente** (`E5/F5`): ajuste manual si el cierre
  es fin de semana/festivo.
- Selección de las **3 fechas** y renombrado de archivos previos para no sobreescribir.

---

## 6. Rutas hardcodeadas y patrones de fecha

- **Base curvas:** `M:\COB\Gerencia Estrategia\Analisis Cuantitativo\BDatos\INFOVALMER\` + prefijo por
  curva (`SwapCC_DTF_Diaria_`, `SwapCC_IBRColateralUSD_Diaria_`, `SwapCC_IBRUVR_Diaria_`,
  `SwapCC_USDOIS_Diaria_`, `SwapCC_LIBORUVR_Diaria_`, `SwapCC_LIBORCOP_Diaria_`, `SwapCC_USDCO_Diaria_`,
  `TasasEstim_USDIBR_Diaria_`, `SwapCC2_EURCOP_Diaria_`, `SwapCCba_LIBORUSD3M_Diaria_Bloomberg_`,
  `SwapCC_COPColateralUSD_Diaria_`) + **`AAAAMMDD`** + `.txt` (EUR-COP y USD-IBR = `.csv`).
- **Monedas:** `...\BDatos\Backup Archivos\AAAA\INDICADORES\AAAA-MM-DD IND.CSV`.
- **Guardado:** `...\BDatos\Backup Archivos\AAAA\SWAP\VPN SWAP dd-mm-aaaa.xlsb` y `DM SWAP dd-mm-aaaa.xlsb`.
- **Patrones de fecha:** importación usa `AAAAMMDD` (sin separador); nombres de salida usan
  `dd-mm-aaaa` (con guiones); indicadores usan `AAAA-MM-DD`. El VBA arma AAAA/MM/DD con
  `Format(Year/Month/Day(Fecha_valoracion),"0000"/"00")`.

---

## 7. Riesgos para reimplementar en Python

1. **La lógica está en fórmulas de celda, no en VBA.** Hay que decodificar las fórmulas reales de
   `SWAPCC_*`/`SWAPIRS_*` (columnas D→O) desde una copia `.xlsx`; `pyxlsb` solo da valores. Este doc
   reconstruye el modelo por ingeniería inversa de valores + la documentación de `RUTAS`, pero
   conviene confirmar las fórmulas exactas de F (forward), K (interpolación) y N/O.
2. **Funciones financieras de Excel:** `CUPON.FECHA.L1` (`COUPNCD`), `CUPON.DIAS`, etc. para generar
   el cronograma de cupones con la base correcta. Reimplementar con una librería de calendarios/day-count
   (ej. `python-dateutil` + funciones day-count propias o `QuantLib`).
3. **Interpolación de curvas:** `K` se interpola por plazo (días) sobre pares (plazo,tasa) de `DATOS`.
   Definir el método exacto (lineal en tasa vs en factor) — no está explícito.
4. **Convención de descuento:** compuesta anual `(1+K)^(−t)` con `t=días/BaseAños` (365 o 360 según
   pata). Bases day-count distintas por pata y por combinación de monedas (tablas `A45`/`A54`).
5. **Patas flotantes complejas:** SOFR compuesto (bloque H:Q de `SWAPIRS_SOF`) e IBR/LIBOR overnight
   compuesto (`LIB-IBR on`) — devengo del cupón corriente con capitalización diaria. Es la parte más
   delicada de replicar.
6. **Dependencia de curvas diarias:** sin los `.txt` INFOVALMER (11+ curvas) no hay valoración. En la
   nube hay que sustituir `M:\` por almacenamiento configurable y garantizar el feed de curvas.
7. **DM por pata + convexidad:** confirmar si Macaulay se divide por `(1+y)` con `y` = tasa de la pata
   o TIR; y la fórmula exacta de la columna N (aporte a duración).
8. **CCS (cross currency):** intercambio de nocional al vencimiento (col G/I) y conversión de moneda
   (TRM/UVR/EUR de hoja `MONEDA`) para expresar VPN en COP. Verificar en qué moneda queda cada
   `VP DERECHO`/`VP OBLIGACIÓN` y dónde se convierte a COP.
9. **Non-delivery:** flag de excepciones que altera el flujo de capital.
```

---
_Análisis basado en VBA (orquestación) + estructura/valores del `.xlsb`. Las fórmulas de celda exactas
deben confirmarse sobre copia `.xlsx` antes de codificar el motor en Python._
