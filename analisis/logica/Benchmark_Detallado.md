# Análisis VBA — Benchmark (Detallado).xlsb

> Documento destilado del código VBA del archivo maestro **Benchmark (Detallado)** (salida final del
> proceso de benchmark mensual/diario de AFPs — Colfondos). Fuente de verdad para reimplementar la
> lógica en Python. Cubre 11 módulos (`Módulo1..Módulo12`, sin Módulo6 duplicado). El VBA crudo está
> en `analisis/vba/Benchmark (Detallado)__Módulo*.bas.vba`.

## Convenciones y contexto general

- Casi todo el flujo se orquesta con **rangos con nombre** (Named Ranges) definidos en la hoja
  `Procedimiento` / `Parametros` (p.ej. `Fecha_Valoración`, `Ruta_PF`, `N_IND`, `RUTA_PUBLICACION`).
  Las rutas físicas no están hardcodeadas en el VBA salvo en los correos (Módulo11); vienen de celdas.
- El libro depende de tener **abiertos simultáneamente** otros libros: `INICIO MACRO MULTIFONDOS.xlsm`,
  `divisas.xlsm`, `Sensibilidad_Fwds.xlsb`, y abre bajo demanda archivos de Notas Estructuradas (CSV),
  Calculadora Swap (VPN/DM), Colfondos diario, etc.
- Patrón recurrente: `ClearContents` de destino → `Windows("X").Activate` + `Copy` → `PasteSpecial
  xlPasteValues` → arrastrar fórmulas (`AutoFill`/`PasteSpecial xlPasteFormulas`) → `Calculate` →
  convertir a valores. Esto es puro Excel/ActiveX y debe reimplementarse como transformaciones pandas.
- Hoja `Colfondos` = base consolidada de posiciones propias; hoja `Benchmark` = industria; los
  "Formato de Apuestas" son las plantillas de reporte (PDF); `Curvas` = curvas forward/descuento;
  `NE` = notas estructuradas; `SWAP`, `Fwd/Fut *` = derivados.

---

## Módulo1 — Núcleo de importación y generación (2711 líneas)

Es el corazón del archivo. Contiene el macro `Benchmark()` (orquestador diario), la importación de
curvas forward, indicadores, y toda la generación de PDF/carpetas.

### `Sub Benchmark()` (línea 2) — Orquestador principal de actualización
**Propósito:** Reconstruye la hoja `Colfondos` (posiciones propias) y hojas satélite a partir de los
libros abiertos, valora swaps, importa notas estructuradas, y encadena futuros + sensibilidad.

- **Entradas (lee):**
  - `INICIO MACRO MULTIFONDOS.xlsm` → hojas `IND` (A1:D + O3:R3, vector de precios industria), `IND2`
    (A1:E, vector RF), `PORTAFOLIOS` (A2:AA, títulos Colfondos).
  - `divisas.xlsm` → hojas `577` (futuros Colfondos, A2:Q), `DPO` (forwards Colfondos, A5:Q + T5:U5),
    `572` (opciones Colfondos), `T695` (swaps, A2:U), `DICCIONARIO` (contrapartes, vía VLOOKUP).
  - Archivo de opciones industria: se abre el libro cuya ruta+nombre son `R_OpInd & N_OpInd`
    (rangos con nombre), hoja `Hoja1` cols A:J.
  - Archivos CSV de Notas Estructuradas: `R_NE\N_NE.CSV` y `R_NE\N_NE_BMK.CSV` (importados con
    `QueryTables.Add TEXT;...`, delimitados por coma/tab).
  - Rangos con nombre: `Limit`, `Fecha_Valoración`, `Fecha_Valoración_ANT`, etc.
- **Salidas (escribe):** hojas `VECTOR DE PRECIOS`, `Fut Colf`, `Fwd Colfondos`, `Opt Colfondos`,
  `Opt Ind`, `SWAP`, `NE`, `Colfondos` (append de swaps derecho/obligación + títulos + fórmulas).
- **Cálculos financieros clave (fórmulas arrastradas en hoja `SWAP` W:AN):**
  - `X` TIPO DE SWAP = `IRS` si moneda pata activa = pata pasiva, si no `CCS`.
  - `Y`/`Z` VPN Derecho/Obligación: para IRS descuenta el nominal traído a valor de mercado por el
    factor de curva (`VECTOR DE PRECIOS` R1C19:R13C20) con tasa de la curva `FWTCOP` (COP) o `LIBBTS`
    (otras) de la hoja `Curvas`, interpolada por plazo `(vencimiento - Fecha_Valoración)`; para CCS usa
    el valor de mercado directo.
  - `AC`/`AD` RIESGO = TASA FIJA / VARIABLE / INFLACIÓN según `MID(nemo,6,2)` = TF/TV/(otro).
  - `AE`/`AF` valor tasa facial: parsea dígitos del nemotécnico `/100`.
  - `AG`/`AH` VPN SWAP TF, `AI`/`AJ` frecuencia (VLOOKUP a `Parametros!E1:F7`), `AK`/`AL` contraparte
    (VLOOKUP a `divisas.xlsm` DICCIONARIO). `Am`/`An` = controles (diferencia contra `Colfondos`).
  - En `Colfondos` se pegan por separado la **pata derecho** (cols SWAP c,d,ak,a,b,h,e,y,f,aa,ac,x,ae,
    ag,ai) y la **pata obligación** (cols k,l,al,a,b,p,m,z,n,ab,ad,x,af,ah,aj); clase = `SWAP` /
    `DERIVADOS`.
  - Fórmulas de participación→sector se arrastran desde `Colfondos!AB(Limit-1):BO(Limit-1)` hacia abajo.
- **Controles/validaciones:** `If Cells(4,1) <> "Opc No"` limpia opciones (evita datos corruptos).
  `MsgBox` "¿Desea actualizar el archivo de Notas Estructuradas" (Yes/No) — decisión humana.
- **Encadena:** `Call futuros` (Módulo6), `Call traer_sensibilidad` (Módulo4),
  `Call ReemplazarPalabraNotas` (Módulo9).
- **Riesgo Python:** depende de 3 libros abiertos por nombre de ventana; orden de append sensible al
  último renglón (`End(xlUp).Row`); QueryTables sobre CSV con formato regional (coma decimal, cp850).

### `Sub IMPORTACION_CURVAS()` (611) — Wrapper de importación de curvas
Recalcula `Procedimiento!A:F`, llama `IMP_CURVAS_HOY`, `Importar_indicadores`, `Limpiar_CONEXIONES`,
`Limpiar_Datos_externos`, recalcula `Curvas!AA:AP`. (`IMP_CURVAS_ANT` está comentado/desactivado.)

### `Sub IMP_CURVAS_HOY()` (627) — Importa curvas forward del día
**Propósito:** Poblar la hoja `Curvas` (cols A:E) con las curvas forward/devaluación por par de divisa
del día de valoración, leyendo archivos `.txt` planos.
- **Entradas:** archivos `Ruta_PF\<Nombre>.txt` donde los nombres vienen de rangos con nombre:
  `N_Fwd_USDCOP_T`, `N_Fwd_USDCAD_T`, `N_Fwd_USDBRL_T`, `N_Fwd_USDMXN_T`, `N_Fwd_USDJPY_T`,
  `N_Fwd_EURUSD_T`, `N_SWAP_LBRUSD_T`, `N_SWAP_LBRCOP_T`, `N_Fwd_EURCOP_T`, `N_Fwd_EUR2COP_T`,
  `N_Fwd_GBPUSD_T`, `N_Fwd_USDCLP_T`, `N_Fwd_USDCHF_T`, `N_Fwd_USDKRW_T`. Ruta base = `Ruta_PF`.
- **Salidas:** hoja `Curvas` con bloques etiquetados `FWPCOP, FWPCAD, FWPBRL, FWPMXN, FWPJPY, FWPEUR...`
  (Orden, Fecha_Valoración_ANT, Nombre_Indicador, Plazo_en_Dias, Valor).
- **Cálculos:** normalización de puntos forward: CAD y EUR `/10000`, JPY `/100` (escalado de pips).
  Fecha de operación = `Fecha_Valoración_ANT`.
- **Riesgo Python:** parsing de TXT con `TextFileSpaceDelimiter`; 14 bloques concatenados con
  offsets calculados dinámicamente (`End(xlUp)`); trivial de reimplementar como `read_csv` + concat.

### `Sub IMP_CURVAS_ANT()` (1320) — Igual que HOY pero para T-1 (curvas anteriores)
Mismo patrón, importa las mismas curvas para la fecha anterior (para comparar/valorar t-1).
Actualmente **desactivado** en `IMPORTACION_CURVAS` (llamada comentada). Reimplementación análoga.

### `Sub Importar_indicadores()` (1867) — Importa indicadores macro (IPC, DTF, IBR, TRM...)
- **Entradas:** 3 archivos CSV `R_IND_1\<N_IND_1|N_IND_2|N_IND_3>.CSV` (delimitados por `;`, cp1252),
  correspondientes a las fechas `Fecha_Valoración`, `-1`, `-2`.
- **Salidas:** hoja `Curvas` cols R:AA; col T se rellena con la fecha de cada bloque. Limpia X:Z.
- **Riesgo Python:** CSV con separador `;`; 3 fechas apiladas.

### `Sub Limpiar_CONEXIONES()` (1985) / `Sub Limpiar_Datos_externos()` (1996)
Borran todas las conexiones de datos externas y los nombres `DatosExternos*` que dejan las QueryTables.
Housekeeping de Excel; **no requiere equivalente en Python** (no habrá QueryTables).

### `Sub Generar_PDF_POCS_Relativo()` (2009) — Genera PDFs de portafolios PO-CS
**Propósito:** Para los 5 portafolios PO-CS genera un PDF por portafolio pegando imágenes de rangos de
`Formato de Apuestas PO-CS` + tablas del Colfondos diario dentro de un documento Word, y lo exporta.
- **Entradas:** rangos con nombre `RUTA_GUARDAR`, `CONTROL_1`; libro Colfondos diario
  (`Procedimiento!E38 & D38` ruta+archivo), hoja `Resumen diario` rango `Resumen_rojosazules_<PORT>`.
  Recorre portafolios en `Parametros!B22:B26` (offset 1..5).
- **Salidas:** PDF en `RUTA_GUARDAR` (vía Word `ExportAsFixedFormat ... , 17`). Usa archivo temporal
  **hardcodeado `C:\IT\MyFirstSave8.docx`**.
- **Controles:** solo genera si `CONTROL_1 = True`; si no, escribe "Error generando PDF de <PORT>" en
  hoja `Controles`.
- **Encadena:** `Generar_PDF_hoja_Voluntarias`, `Generar_PDF_VOLUNTARIAS`.
- **Riesgo Python:** usa **Word ActiveX** + `CopyPicture` (imágenes de rangos) → reimplementar como
  render de tablas a PDF (reportlab/plantilla); ruta temporal `C:\IT\...` hardcodeada.

### `Sub Generar_PDF_VOLUNTARIAS()` (2293) — PDFs de pensión voluntaria
Recorre `Parametros!B22` offset 6..44. Para i=7..11 usa vía Word (imágenes); resto exporta directo con
`ExportAsFixedFormat Type:=xlTypePDF`. Lee Colfondos diario de `Procedimiento!E39 & D39`, hoja
`Resumen` rango `rojosazules_<PORT>`. Control `CONTROL_1`. Mismo temporal `C:\IT\MyFirstSave8.docx`.

### `Sub Generar_PDFCP()` (2414) — PDF de Cesantías Corto Plazo
Toma portafolio de `Parametros!C50`, pone en `Formato de Apuestas PO CS!B12`, exporta rangos fijos
(`$A$1:$l$79,$n$1:$X$80,$CU$1:$DF$107`) a PDF si `CONTROL_2 = True`.

### `Sub Generar_PDF_hoja_Voluntarias()` (2667)
Exporta `Voluntarias!$A$1:$k$87` a PDF (`infVol`) si `CONTROL_2 And CONTROL_VOLUNTARIAS = True`;
si no, `MsgBox "Revisar la hoja voluntarias"`.

### Utilitarios de Módulo1
- `btnMensaje_Click()` (2454): activa/desactiva autofiltros en todas las hojas (contiene comillas
  tipográficas “ ” que romperían la compilación — código con bug menor).
- `Guardar()` (2467): `SaveAs` a `RUTA_PUBLICACION\N_PUBLICACION.xlsb` (xlExcel12) + `Copiar_paste`.
- `Copiar_paste()` (2482) / `Copiar_paste2()` (2640): convierten **todas las hojas a valores**
  (Copy → PasteSpecial xlPasteValues sobre `A1:CA1048576`). Congela fórmulas antes de publicar.
- `GenerarCarpeta()` (2536): crea carpetas mensuales/consolidado (`MkDir`) según rangos
  `C_M, C_P, N_M2, N_P` (recorre 20 portafolios).
- `Generar_Carpeta_Diaria()` (2599): crea carpeta diaria (`C_D\N_D`).
- `PDF()` (2626): orquesta `Generar_Carpeta_Diaria` + varias `Generar_PDF*`.

---

## Módulo2 — Valoración financiera de SWAPS (funciones VPN / Duración) (557 líneas)

Librería de **funciones UDF** que valoran flujos de swap y su duración. Estas son el modelo financiero
que debe replicarse fielmente en Python. Todas reciben `VPN0DUR1` (0=VPN, 1=Duración/DV01).

### `Function Interp(plazos, curva, T)` (4)
Interpolación **lineal** de una curva de tasas por plazo `T` (con extrapolación plana en los extremos).
Base para todas las valoraciones.

### `Function Tiempo(modalidad)` (24)
Mapea la periodicidad textual a meses: `A`→12, `S`→6, `T`→3, `M`→1, `P`→0 (cupón único / cero cupón).

### `Function SWAPTF(...)` (44) — Swap de tasa fija
Descuenta cupones `NOMINAL*CUPON/frecuencia` (más nominal al final) usando curva `Dias_Libor/Tasa_Libor`
(si MONEDA=USD) o `Dias_IBR/Tasa_IBR` (otra). VPN = Σ flujo·factor; Duración por diferencias ±1pb
`(VPN(-1pb)-VPN(+1pb))/(2·VPN·0.0001)`. Caso `T=0` (bullet) usa `(vencimiento-datum)/365`.
Retorna `"Vencido"` si `vencimiento < datum`.

### `Function EmisionInicial(vencimiento, modalidad, datum)` (195)
Calcula la fecha del primer cupón vivo retrocediendo desde vencimiento en pasos de `T` meses.

### `Function Numcupones(origen, vencimiento, modalidad)` (217)
Cuenta cupones entre origen y vencimiento.

### `Function SWAPIPC(...)` (243) — Swap indexado a inflación (IPC)
Cupón compuesto `(1+IPC)*(1+CUPON)-1`; descuenta con tasa `IPC` (rango con nombre). Duración escalada
`*10000`.

### `Function SWAPDTF(...)` (314) — Swap indexado a DTF
Cupón `(1+DTF)*(1+CUPON)-1`; primer cupón usa DTF interpolada de `Range("fecha")`/`Range("DTFSERIE")`.
Parámetro `base` (0=365, 1=trimestral 1/4). Descuenta con `DTF`.

### `Function SWAPUVR(...)` (402) — Swap en UVR (real)
Cupón real `(1+CUPON)^(dias/365)-1`; descuenta con curva `Dias_UVR/Tasa_UVR`.

### `Function SWAPIBRPF(...)` (477) — Swap IBR pata fija
Descuenta con curva `IBRF1/IBRF2`. `SWAPLIBOR`/`SWAPIBRPV` (238/472) devuelven `1/365` (patas flotantes
valuadas a la par).

- **Riesgo Python:** dependen de **rangos con nombre** (curvas: `Dias_Libor`, `Tasa_IBR`, `IPC`, `DTF`,
  `DTFSERIE`, `Dias_UVR`, `IBRF1/2`) que hay que cargar como vectores. Convención año 365, cálculo DV01
  por bump ±1pb. Son reimplementables 1:1 en numpy.

---

## Módulo3 — Generación de Benchmark y portafolios publicados (210 líneas)

### `Sub Generar_Benchmark()` (2)
**Propósito:** Publica el archivo maestro. Desactiva autofiltros de todas las hojas, congela
`Fecha_Valoración` a valor, llama `Copiar_paste` (todo a valores), y guarda `SaveAs` en
`RUTA_PUBLICACION\N_PUBLICACION.xlsb`. Luego `Vector_duraciones_SWAPs` (Módulo10) y `Generar_Portafolios`.

### `Sub Generar_Portafolios()` (38)
Bucle i=1..5 (portafolios PO-CS): setea `Parametros!S21`, `Formato de Apuestas!B9` y `PO-CS!B9` al
código de portafolio (`Parametros!B22..B26`), recalcula, y copia un subconjunto de hojas a un libro
nuevo guardado en `R_IND\N_IND.xlsb`. Encadena `Generar_Portafolios_Voluntarias`.

### `Sub Generar_Portafolios_Voluntarias()` (80)
Bucle i=6..45 (voluntarias): igual, guarda en `R_IND_VOL\N_IND_VOL.xlsb`. i=6 exporta set de hojas
distinto (incluye `Formato de Apuestas CP`).

### `Sub Borrar_Exceles()` (162)
Limpieza: recorre `Parametros!B22` offset 6..40, resuelve `R_BORR\N_BORR.xlsb` y borra el archivo con
`FileSystemObject.DeleteFile`; registra en `Procedimiento!C91:D200` los "No existe"/"Error al borrar".

- **Salidas:** N archivos `.xlsb` por portafolio (PO-CS y voluntarias) en rutas de rangos con nombre.
- **Riesgo Python:** SaveAs/Copy de hojas concretas → generar Excel por portafolio con openpyxl/
  xlsxwriter; borrado con `os.remove`.

---

## Módulo4 — Sensibilidad de forwards (integración con Sensibilidad_Fwds) (103 líneas)

### `Sub Macro1()` (4) — trivial (autofill de la selección hacia abajo).

### `Sub traer_sensibilidad()` (15)
**Propósito:** Trae la sensibilidad (KRD/duración por clave de riesgo) del libro externo
`Sensibilidad_Fwds.xlsb` y calcula la sensibilidad de los forwards de Colfondos e industria.
- **Entradas:** abre `Ruta_Bmk & Sensibilidad` (rangos con nombre) → hoja `Ind` cols Z3:AA (KRD local/
  foráneo). Rangos `CP`, `LP` (umbrales corto/largo plazo).
- **Salidas:** hoja `Fwd Industria` cols S:W (Plazo, KRD Local/Foráneo, Sens Local/Foráneo) y hoja
  `Fwd Colfondos` cols U:Y.
- **Cálculos:** clasificación de plazo (CORTO/MEDIANO/LARGO PLAZO según `CP`/`LP`); sensibilidad =
  `KRD * (exposición Compra/Venta / KRD del Benchmark)` vía `INDEX/MATCH` sobre `Benchmark!I2:K6` y
  `HLOOKUP` sobre `Colfondos!G1:...`. Para Colfondos usa VLOOKUP a `Sensibilidad_Fwds.xlsb!572` col 38/39.
- **Riesgo Python:** vínculo a libro externo `Sensibilidad_Fwds.xlsb`; fórmulas INDEX/MATCH/HLOOKUP a
  reproducir como merge/lookup.

---

## Módulo5 — Exportación a CSV (557... 86 líneas)

### `Sub calcular()` (2): `Selection.Calculate` (utilitario).

### `Public Sub SaveWorksheetsAsCsv()` (9)
**Propósito:** Genera los CSV de las hojas del benchmark (para consumo externo/consolidado).
Llama `Copiar_paste`, borra un set de hojas de trabajo, y guarda **cada hoja restante como CSV** en la
carpeta `R_1` (rango con nombre). Nombre = nombre de hoja.
- **Salidas:** múltiples archivos `.CSV` en `R_1`.
- **Riesgo Python:** `df.to_csv` por hoja; el set de hojas a excluir está hardcodeado.

---

## Módulo6 — `Sub futuros()` (13 líneas)
Escribe en `Curvas_2` la fórmula/valor tomada de `Procedimiento!G36` (celda destino) y
`Procedimiento!H36` (fórmula que apunta a un libro externo de PYG/vector), y la convierte a valor.
Mecanismo genérico de "traer un valor externo por fórmula". Llamado desde `Benchmark()`.

---

## Módulo7 — `Sub futuros_trm()` (117 líneas)
**Propósito:** Agrega al final de la hoja `Fwd Industria` los **futuros de TRM** tomados de la hoja
`Fut Colf` (filtrando col 21 = "TRM"). Copia columnas concretas (C→plazo, BE, AH, K, G, R, D, B) a los
bloques destino y arrastra fórmulas (G, I:AF, AH). Deja los futuros de TRM junto a los forwards para
que el cálculo de exposición cambiaria los incluya. Termina en hoja `Controles`.
- **Riesgo Python:** filtro por "TRM" + append + fórmulas de exposición; determinista.

---

## Módulo8 — `Sub ReemplazarPalabraEnHojasEspecificas()` (101 líneas)
**Propósito:** Normaliza nombres de portafolios/AFP a códigos cortos en `FutLoc Industria`,
`FutInt Industria`, `Fwd Industria`. Diccionario de reemplazos **hardcodeado**:
- "FONDO DE PENSIONES OBLIGATORIAS ... CONSERVADOR" → `PC`; "... MAYOR RIESGO" → `PM`;
  "... RETIRO PROGRAMADO" → `PR`; "... MODERADO" / "MODERADO" → `PO`; "FONDO DE CESANTIAS[ SKANDIA]" → `CS`.
- Cubre variantes de PORVENIR, PROTECCIÓN, SKANDIA. También `PROTECCION`→`PROTECCIÓN` y
  `SKANDIA PENSIONES Y CESANTÍAS S.A.`→`OLD MUTUAL`.
- **Punto de juicio/mantenimiento:** este diccionario debe mantenerse si cambian nombres de fondos SFC.
- **Riesgo Python:** tabla de reemplazo de strings; trivial pero frágil ante cambios de nombre.

---

## Módulo9 — Reemplazos y cupones de swap (42 líneas)

### `Sub ReemplazarPalabraNotas()` (2)
En hoja `NE` corrige `PROTECCIËN`→`PROTECCIÓN` (mojibake de codificación). Llamado desde `Benchmark()`.

### `Sub Consolidar_Cupones_Swaps()` (25)
Recalcula y **apila** (append valores) `Cupónes Swaps!A3:P3` al final de la hoja. Acumula histórico de
cupones de swap generados en cada corrida.

### `Sub Eliminar_Cupones_Swaps()` (34)
Limpia `Cupónes Swaps!A4:P10000` (reset del histórico). Botón "SWAP MES".

---

## Módulo10 — `Sub Vector_duraciones_SWAPs()` (78 líneas)
**Propósito:** Genera dos CSV de apoyo (Colfondos y Benchmark) con la información reducida usada para el
vector de duraciones de swaps.
- Copia hojas `Colfondos` y `Benchmark` a un libro nuevo; filtra por tipo de activo
  (`CAJA, FORWARD, Notas estructuradas, OPCIONES, R FIJA, R VARIABLE`) para **eliminar** esas filas y
  dejar el resto (swaps/duración); borra columnas intermedias dejando las claves.
- **Salidas:** `R_1\A_dur_c.CSV` (Colfondos) y `R_1\A_dur_b.CSV` (Benchmark).
- **Riesgo Python:** filtrado de filas por categoría + recorte de columnas; determinista.

---

## Módulo11 — Envío de correos (Outlook) (393 líneas)

Cuatro Subs que crean correos vía `CreateObject("Outlook.Application")`. Cuerpo HTML con hipervínculo e
imagen a rutas **M:\ hardcodeadas**. Todos los destinatarios/asuntos/adjuntos vienen de la hoja
`Parametros`.

### `Sub Enviar_Email_BMK_PO()` (2)
Correo de Pensión Obligatoria. To/CC/BCC/Subject = `Parametros!C78/C79/C80/C81`. Adjunta 5 archivos
(`Parametros!C84:C88`). Cuerpo cita `FechaBMK` (`Procedimiento!D7`) y `fechaIND` (rango `fechaIND`).
Encadena `Enviar_Email_BMK_PV` y `Enviar_Email_Caldas_Manizales`.

### `Sub Enviar_Email_BMK_PV()` (63)
Pensión Voluntaria. To/CC/BCC/Subject = `Parametros!C96:C99`. Adjunta ~40 archivos (`C102:C116` +
`E166:E191`).

### `Sub Enviar_Email_Caldas_Manizales()` (192)
Variante para Caldas/Manizales; Subject = `Parametros!D99`; adjunta `C115` y `E192`.

### `Sub Enviar_Email_BMK_PVant()` (236)
Versión previa (todo el bloque de adjuntos PV + Caldas). Mantenida como respaldo.

- **Salidas:** correos Outlook con PDFs adjuntos.
- **Rutas M:\ hardcodeadas:** en el cuerpo HTML (link e imagen `Imagencorreo.png` / `Publicados.png`).
- **Riesgo Python:** reimplementar con SMTP/API de correo; lista de adjuntos definida por celdas de
  `Parametros` (no en el código).

---

## Módulo12 — `Sub Copiarvalores()` (13 líneas)
Copia el valor del rango con nombre `PORTAFOLION` al rango `PORTAFOLIO` (selector de portafolio activo).
Utilitario de UI.

---

## Rutas M:\ hardcodeadas encontradas (literales en el VBA)
Solo aparecen en los cuerpos HTML de los correos (Módulo11), todas bajo la misma base:

```
M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados
M:\COB\Gerencia Estrategia\Analisis Cuantitativo\...\Benchmark\Publicados\Imagencorreo.png
M:\COB\Gerencia Estrategia\Analisis Cuantitativo\...\Benchmark\Imagen publicados\Publicados.png
```

Ruta local temporal (no M:\): `C:\IT\MyFirstSave8.docx` (Módulo1, generación de PDF vía Word).

Todas las demás rutas (`Ruta_PF`, `R_NE`, `R_IND`, `R_IND_VOL`, `RUTA_PUBLICACION`, `R_1`, `R_BORR`,
`Ruta_Bmk`, `R_OpInd`, `R_IND_1`, `C_M`, `C_P`, `C_D`) provienen de **rangos con nombre** en las hojas
`Procedimiento`/`Parametros` — no están literales en el código; su valor real vive en el `.xlsb`.

## Libros externos que abre / requiere abiertos
- Abiertos por ventana (deben estar abiertos): `INICIO MACRO MULTIFONDOS.xlsm`, `divisas.xlsm`.
- Abiertos bajo demanda: `Sensibilidad_Fwds.xlsb`, opciones industria (`R_OpInd\N_OpInd`), Colfondos
  diario (PDF), Calculadora Swap (vía vínculos), CSV de Notas Estructuradas (`N_NE`, `N_NE_BMK`),
  indicadores (`N_IND_1/2/3`), curvas forward (`.txt`).
- Vínculos de fórmula: `[divisas.xlsm]DICCIONARIO`, `[Sensibilidad_Fwds.xlsb]572`, `[...]620 Contable`.

## Outputs generados (PDF / Excel / CSV / correo)
- **PDF** por portafolio: PO-CS (`Generar_PDF_POCS_Relativo`), Voluntarias
  (`Generar_PDF_VOLUNTARIAS`, `Generar_PDF_hoja_Voluntarias`), Cesantías CP (`Generar_PDFCP`) →
  a `RUTA_GUARDAR`/`RUTA_GUARDAR_VOL`/`infVol`.
- **Excel (.xlsb)**: archivo maestro publicado (`RUTA_PUBLICACION\N_PUBLICACION.xlsb`) y un `.xlsb` por
  portafolio PO-CS (`R_IND\N_IND`) y por voluntaria (`R_IND_VOL\N_IND_VOL`).
- **CSV**: hojas del benchmark (`SaveWorksheetsAsCsv`→`R_1`); vector de duraciones swaps
  (`Vector_duraciones_SWAPs`→`R_1\A_dur_c.CSV`, `A_dur_b.CSV`).
- **Correo (Outlook)**: BMK PO, BMK PV, Caldas/Manizales, PVant — con PDFs adjuntos definidos en
  `Parametros`.

## Puntos que requieren juicio humano / insumos externos detectados
1. **Notas Estructuradas** — `MsgBox` Yes/No en `Benchmark()` para actualizar; se importan 2 CSV
   (`N_NE`, `N_NE_BMK`) que un humano prepara/valida; corrección de mojibake y conversión de moneda
   (BRL→COP) fuera del VBA (paso manual del proceso).
2. **Diccionario de nombres de fondos** (Módulo8) — tabla hardcodeada de fondos SFC→código; requiere
   mantenimiento cuando aparecen fondos/nombres nuevos.
3. **Contrapartes de swaps** — VLOOKUP a `divisas.xlsm!DICCIONARIO`; contrapartes nuevas deben existir
   en ese diccionario o devuelven error.
4. **Sensibilidad de forwards** — depende del libro `Sensibilidad_Fwds.xlsb` corrido aparte (paso
   manual con botón) y de umbrales `CP`/`LP`.
5. **Controles** (`CONTROL_1`, `CONTROL_2`, `CONTROL_VOLUNTARIAS`, `Positivo...`) — celdas booleanas que
   abortan la generación de PDF/portafolio y registran errores en la hoja `Controles`; su evaluación
   depende de fórmulas de validación de la hoja (composición, cuadre industria vs consolidado).
6. **Rangos con nombre de rutas y fechas** — el operador fija `Fecha_Valoración`, `Fecha_Valoración_ANT`
   y las rutas mensuales en `Procedimiento`/`Parametros` antes de correr (paso manual del proceso).

## Riesgos transversales para reimplementar en Python
- **Dependencia fuerte de Excel/ActiveX:** Word (`CreateObject("Word.Application")`) y Outlook para
  PDF y correo; QueryTables para CSV/TXT; múltiples libros abiertos referenciados por nombre de ventana.
- **Lógica dispersa entre VBA y fórmulas de celda:** gran parte del cálculo (participaciones, sectores,
  controles, VPN swap en hoja `SWAP`) está en **fórmulas arrastradas**, no en el VBA → hay que
  decodificar las fórmulas desde una copia `.xlsx` (varias ya transcritas arriba: SWAP W:AN, sensibilidad).
- **Orden y estado:** append por `End(xlUp).Row`, `ClearContents` previos, y encadenamiento estricto
  (`Benchmark`→`futuros`→`traer_sensibilidad`→`ReemplazarPalabraNotas`). El orden importa.
- **Formato/codificación:** CSV con coma decimal, `cp850`/`cp1252`, delimitadores `;`/`,`/tab/espacio;
  fechas AMD; escalado de pips por divisa (CAD/EUR /10000, JPY /100).
- **Rutas y credenciales de red `M:\`** y `C:\IT\` que no existirán en el entorno destino (sustituir por
  almacenamiento configurable).
