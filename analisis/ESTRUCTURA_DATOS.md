# Estructura de datos y tablas de configuración (complemento al análisis VBA)

> Extraído leyendo directamente las hojas de datos de los 5 archivos (no el VBA).
> Define las **tablas de mapeo** que en Python se vuelven configuración, el **esquema de columnas
> SFC**, y el **mecanismo de detección de instrumentos nuevos** (base del sistema de alertas).

## 1. Esquema de columnas SFC (hoja `Columnas` del BMO)

El BMO documenta el orden de columnas de los dos formatos de la descarga SFC:

**Fmto-351 (activos)** — cols: Tipo de Entidad · Código Tipo Entidad · Nombre de Entidad ·
Fecha de Corte · Código Tipo Patrimonio · Código Patrimonio · Nombre Patrimonio · Subtipo Patrimonio ·
Price Vendor · Desc Price Vendor · Unidad de Captura · Renglón · Secuencia · Nro_Asignado · Código PUC ·
Tipo_Id_Emisor · No. ID Emisor · Razón Social Emisor · … · Clase de Inversión · Descripción Clase ·
Nemotécnico · Cupón Principal · …

**Fmto-415 (derivados)** — cols: Entidad · Fecha Corte · Tipo Patrimonio · SubTipo Patrimonio ·
Código Patrimonio · Nombre · Unidad captura · Subcuenta · Categoría derivado · **Tipo de derivado (col J)** ·
Tipo de opción · Número de contrato · Finalidad de la operación · Forma de liquidación · Intercambio prima ·
Moneda liquidación · Tipo de estrategia · Identificación estrategia · Ramo de seguros · Tipo contraparte ·
Identificación paramétrica · Relación contraparte · **Fecha celebración (W)** · **Fecha vencimiento (X)** ·
**Fecha liquidación (Y)** · Posición de compra o venta · Código moneda derecho · Nominal derecho ·
Código moneda obligación · Nominal obligación · …

> Confirma que **col J = "Tipo de derivado"** (el discriminador 1/4/5/6/16/17) y W/X/Y = fechas.

## 2. Diccionario SFC (hoja `Diccionario SFC` del BMO) — tablas de mapeo

Una sola hoja con varios bloques de mapeo:

### 2.1 Nombres de AFP (cols A–E)  — normalización
Original (col A o C) → estándar. Incluye control "DEBE ESTAR / ESTA".

### 2.2 Portafolios (cols F–J) — SFC → código BMK
Nombre patrimonio SFC → código: **CS** (cesantías), **PO** (moderado/obligatorio),
**PC** (conservador), **PM** (mayor riesgo), **PR** (retiro programado), **ALTERNATIVO** (FA).

### 2.3 Clas SFC → Clase de Inversión (cols K–M)  ⭐ clave para clasificación
Mapea el código de clasificación SFC a la Clase de Inversión interna. Ejemplos reales:

| Clas SFC | Clase de Inversión |
|---|---|
| FWUSD/FWCOP/FWEUR/… | FORWARD |
| OPT USD / OPT EUR | OPCIONES |
| DEPVN | VISTA · DEPVE → MONEY MARKET · DEPVO → VISTA |
| TSUV | Inflacion-TES · TSTF → TF TES |
| BOEVS | EM CORPORATIVOS / Inflacion-CORPORATIVO LOCAL / CORPORATIVO LOCAL (según moneda) |
| TDPIT | Inflacion-CORPORATIVO / CORPORATIVO LOCAL / EM CORPORATIVOS |
| FINDI / PFMUIA | Fondo Mutuo o Fondo Índice |
| PEEBE | Notas estructuradas |
| FCP / FCPI / FCPAPP | Fondo de Capital Privado |
| AAEVS / CRAAAVS | Acciones Colombia |
| TDPE | EM BONDS |

> El mismo código (p.ej. BOEVS) mapea a distinta clase según **moneda** → la clasificación
> depende de (Clas SFC, moneda). Casos especiales (NUAMCO, CAMACOL, acciones) se tratan aparte.

### 2.4 Códigos de derivados (cols R–U y W–X) — taxonomía SFC
Código → Tipo: 1 Forward divisas · 2 Forward títulos · 3 Forward valores · 4 Futuro divisas ·
5 Futuro títulos · 6 Futuro valores · 7–15 Opciones (europea/americana/barrera/…) · 16 Swap IRS ·
17 Swap CCS · 18 CDS · 19 Asset Swap · 20 Futuros de tasa · 21 FRA · 22–26 Opciones/Forwards/Futuros
sobre materias primas. (Nota: el enrutamiento del código J en la importación usa 1/4/5/6/16/17.)

### 2.5 Control SFC (cols N–P: "REPORTADOS / ACTUALES / TIPO / CONTROL") ⭐ detector
Lista de códigos SFC **conocidos**. La macro compara los códigos **reportados** en el archivo SFC
del mes contra los **actuales** (conocidos); un código nuevo dispara el control → **instrumento nuevo**.

## 3. Clasificación de activos (INICIO MACRO MULTIFONDOS)

### 3.1 Hoja `CLASIFICACION` — tabla maestra ISIN/NEMO → atributos
Columnas: **NEMO · CLASE DE INVERSION · MONEDA · TASA FACIAL · UBICACION · CLASIFICACION · RIESGO ·
CLASE** (+ columnas auxiliares). `CONTROL = COUNTIF(A:A, A_n)`.
UBICACION ∈ {NACIONAL, INTERNACIONAL}; CLASIFICACION ∈ {R VARIABLE, R FIJA, CAJA, …};
RIESGO ∈ {VARIABLE, TASA FIJA, …}.

### 3.2 Hoja `TITULOS NUEVOS` — ISIN · NEMO · FRECUENCIA · MONEDA · CONTROL(COUNTIF)
Registro de títulos nuevos ya resueltos (frecuencia de cupón + moneda).

### 3.3 Hoja `CPRVI` — mapeo de moneda/región de fondos RVI y notas ⭐
Por ISIN: pesos por **región** (GLOBAL, US EQUITY, EUROPEAN EQUITY, JAPAN, EM GLOBAL, EM ASIAN,
LATIN AMERICAN, ALTERNATIVE, ORO, PLATA, REAL ESTATE) y por **moneda** (USD/EUR/JPY/CAD/GBP/MXN/
BRL/AUD/COP). RIESGO via `XLOOKUP`. `CONTROL = COUNTIF`. Para Fondos Mutuos/ETF/Notas nuevos hay
que agregar la fila (diccionario del equipo de Renta Variable).

### 3.4 Hoja `PORTAFOLIOS` — detalle de activos con clasificación resuelta
Fondo · Nemo · ISIN · Título · Emisor · Especie · F.Compra · Emisión · F.Vcto · Moneda ·
Vr Nominal · Vr Mercado · CLASE DE INVERSION · MONEDA · TASA FACIAL · UBICACION · CLASIFICACION ·
RIESGO · VALOR TASA FACIAL · TASA DE DESCUENTO · FRECUENCIA · …

## 4. Mecanismo de detección de instrumentos nuevos (⭐ diseño de alertas)

La clasificación se resuelve por búsqueda (`XLOOKUP`/`COUNTIF`) del ISIN/código contra las tablas de
referencia. Un instrumento es **nuevo / no contemplado** cuando:

1. Su **Clas SFC** no está en `Diccionario SFC` cols K–M ni en el control SFC cols N–P → clase = **ND**.
2. Su **ISIN** no está en `CLASIFICACION` (COUNTIF = 0) → sin moneda/región/riesgo.
3. Es Fondo Mutuo/ETF/Nota y su ISIN no está en `CPRVI` → no mapea moneda/región (posición no suma).

En el proceso manual esto se ve filtrando **Clase de Inversión = ND** (paso 6c del manual). En la
automatización, **cada uno de estos "no-match" alimenta el archivo de alertas** con: portafolio, AFP,
ISIN, NEMO, Clas SFC, moneda, valor de mercado, y motivo ("Clas SFC desconocido", "ISIN sin
clasificar", "fondo/nota sin mapeo moneda-región").

## 5. Calculadora Swap — insumos de valoración

- Hoja `MONEDA`: tabla de tasas/valores por código, incl. **betas Nelson-Siegel** de la curva TES
  (`B01TF/B02TF/B03TF`) y UVR (`B0UVR/B1UVR/B2UVR`), TRM y monedas (AUD, BRL, CAD, CHF, CLP, …).
- Hoja `INICIO`: `Fecha`, `TRM`, `UVR`, `EURL`; curvas con nombre de archivo y ruta
  `M:\COB\BDatos\INFOVALMER\...` (p.ej. `SwapCC_DTF_Diaria_...`); controles (Valor<>0, Valor=#N/A);
  `RUTA GUARDAR` + `NOMBRE GUARDAR` = `VPN SWAP DD-MM-AAAA` y `DM SWAP DD-MM-AAAA`.
- Hojas `SWAPIRS_*` / `SWAPCC_*`: una por combinación curva/moneda; campos FECHA VALORACION, FONDO,
  N° contrato, FRECUENCIA (AV=anual vencido, SV=semestral vencido, TV=trimestral vencido), curvas
  Compuesto SOFR / IBR, cálculo de días (ACUM/FUT/LIQ). Detalle del modelo → ver
  `analisis/logica/Calculadora_Swap.md`.

## 6. Rutas y parámetros centralizados (hoja `Parámetros` del BMO)

Todas las rutas `M:\` y nombres con patrón de fecha viven en `Parámetros` (col B etiqueta, col C valor):
- `Fecha Benchmark` (C3), `Fecha SFC` (C5) — fechas de corte.
- `Ruta BenchmarkIND`, `Ruta Benchmark mes`, `Ruta V. Precios swaps`, `Ruta V.P.`,
  `Benchmark Detallado`, `Nueva Industria SFC` — todas `M:\COB\Gerencia Estrategia\...`.
- `C17` = nombre del archivo SFC del mes (`AAAA_MMportainvdeta.xlsx`).
- `V. Precios Swaps` = `VPN SWAP DD-MM-AAAA.xlsx`; `Benchmark IND` = `Benchmark DD-MM-AAAA.xlsx`;
  `Indicadores Monedas` = `AAAA-MM-DD IND`.
- **Reconciliación**: `Valor de fondo MACRO` vs `Valor del Fondo SFC` por portafolio (PM/CS/PO/PC/PR);
  `Número títulos MACRO` vs `SFC`; `Número Eliminados` + `Mercado Eliminados`;
  `Top +/- Diferencias` de precio con NEMO (control de precios atípicos).
- Paneles de control **Importación** y **Exportación** (todos deben ser `True`; UVR puede ser `False`).

> En Python, `Parámetros` se reemplaza por un archivo de **configuración** (rutas de almacenamiento,
> fechas de corte, umbrales de control) y las tablas de mapeo por **datos de referencia versionados**.
