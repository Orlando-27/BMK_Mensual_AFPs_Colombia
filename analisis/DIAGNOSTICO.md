# Diagnóstico integral — Automatización del Benchmark Mensual (AFPs Colombia)

> Síntesis del análisis completo del manual + ingeniería inversa de los 5 archivos Excel
> (~10.000 líneas de VBA + estructura de hojas). Objetivo: **proceso mensual automático** que
> descargue, procese siguiendo el manual, genere el output y produzca un **archivo de alertas** ante
> fallos o instrumentos nuevos no contemplados. Referencias detalladas en `analisis/logica/*.md` y
> `analisis/ESTRUCTURA_DATOS.md`.

---

## 1. Resumen ejecutivo

- El proceso es **reimplementable en Python** de punta a punta. La lógica de importación, limpieza,
  clasificación y valoración es **determinista**; los puntos de juicio humano son acotados y
  canalizables a un **archivo de alertas**.
- Hoy vive en **5 libros Excel encadenados por vínculos + macros VBA + rutas de red `M:\`**, con
  dependencia de **Word (PDF)** y **Outlook (correo)**. Nada de eso es portable a nube tal cual → se
  reimplementa.
- **Hallazgo crítico:** una parte importante de los cálculos **NO está en VBA sino en fórmulas de
  celda arrastradas** (clasificación, valoración de swaps, sensibilidad, participaciones). El VBA
  mayormente **orquesta**. Para replicar con exactitud hace falta **extraer esas fórmulas** de una
  copia `.xlsx` de los libros (ver §7 y §11).
- **Factibilidad confirmada en el entorno nube:** la web de la SFC es **alcanzable** (HTTP 302); NO
  hay Excel, NO hay `M:\`, y LibreOffice está inoperativo (no convierte archivos) → la extracción de
  fórmulas requiere copias `.xlsx` generadas desde Excel.

---

## 2. Arquitectura actual (cómo funciona hoy)

```
                 ┌─────────────────────────── SFC (web, día 10) ───────────────────────────┐
                 │  AAAA_MMportainvdeta.xlsx  → hojas Fmto-351 (activos) y Fmto-415 (derivados) │
                 │  CS-CP (portafolio corto plazo)                                             │
                 └──────────────────────────────────┬──────────────────────────────────────┘
                                                     │
   Insumos externos (INFOVALMER / BDatos):          ▼
   - Vector de precios diario            ┌───────────────────────────────┐
   - ~11 curvas swap diarias             │  BENCHMARK MENSUAL OPTIMIZADO  │  (BMO)
   - Indicadores (TRM/UVR/IPC/DTF/IBR)   │  importa, limpia, clasifica    │
   - Valoración Notas Estructuradas CSV  │  derivados+activos de industria│
   - Diccionario RVI (equipo R.Variable) └───────────────┬───────────────┘
                                                          │ exporta industria
        ┌──────────────────────┐   VPN/DM   ┌─────────────▼───────────────┐   sensibilidad  ┌───────────────────┐
        │  CALCULADORA SWAP     │──────────►│   BENCHMARK (DETALLADO)      │◄────────────────│ SENSIBILIDAD_FWDS │
        │  valora swaps (curvas)│           │  consolida todo, genera      │                 │  KRD forwards     │
        └──────────────────────┘           │  Excel + PDF + correos       │                 └───────────────────┘
        ┌──────────────────────┐           │                              │
        │ INICIO MACRO MULTIF.  │──────────►│  clasificación de activos    │
        │ clasifica (VLOOKUP)   │           └──────────────┬───────────────┘
        └──────────────────────┘                          │
                                          Salidas: Consolidados .xlsb + PDFs por fondo + correos
```

### Rol de cada archivo
| Archivo | Rol | Lógica principal |
|---|---|---|
| **Benchmark Mensual Optimizado** (BMO) | Importación/limpieza/clasificación de industria | `Importar()` orquesta 7 pasos; enruta derivados por `Tipo de derivado` (col J: 1/4/5/6/16/17); limpia <1000 COP y Fondo Alternativo; exporta industria al Detallado. |
| **INICIO MACRO MULTIFONDOS** | Clasificación de activos y mapeo moneda/región | `clasificacionPO` (VLOOKUP contra hoja `CLASIFICACION` A:M); nuevo = `#N/A`. |
| **Calculadora Swap** | Valoración VPN y DM de swaps | `VP = Σ(H+I)·(1+K)^(−J/base)` por pata; `DM = Macaulay/(1+y)`; curvas por tipo de swap. Fórmulas en hojas motor `SWAP*`. |
| **Sensibilidad_Fwds** | Sensibilidad (KRD) de forwards | `KRD ≈ −(V(+100pb)−V(−100pb))/(2·0.01·V_base)`; fwd = `TRM·(1+r_loc·t/360)/(1+r_for·t/360)`. |
| **Benchmark (Detallado)** | Consolidación final + salidas | `Benchmark()` reconstruye Colfondos, valora swaps (bump DV01 ±1pb), importa NE, encadena futuros/sensibilidad; genera Excel/PDF (Word)/correos (Outlook). |

---

## 3. Insumos externos (lo que hay que "descargar/traer")

| Insumo | Fuente | Frecuencia | Uso |
|---|---|---|---|
| Portafolio detallado SFC (351/415) | Web SFC (día 10) | Mensual | Posiciones de industria (activos + derivados). |
| Portafolio CS-CP | Web SFC | Mensual | Cesantías corto plazo. |
| Vector de precios | INFOVALMER (`M:\...\Vector Precios`) | Diario (último del mes) | Precios/duraciones/tasas faciales. |
| ~11 curvas swap | INFOVALMER (`M:\...\BDatos\INFOVALMER\Swap*_Diaria_AAAAMMDD.txt`) | Diario | Descuento de swaps (IBR, SOFR, LIBOR, DTF, UVR, EUR…). |
| Indicadores (TRM/UVR/EUR/IPC/DTF/IBR) | `M:\...\Backup Archivos\AAAA\INDICADORES\AAAA-MM-DD IND.CSV` | Diario | Spots y tasas de referencia. |
| Valoración Notas Estructuradas | `M:\...\Notas Estructuradas\Valoración NE DD-MM-AAAA.csv` | Diario | Valoración de NE (con conversión BRL→COP). |
| Diccionario RVI (moneda/región) | Equipo Renta Variable (correo) | Ad-hoc (instrumentos nuevos) | Mapeo de FM/ETF/notas nuevos. |

> **Implicación:** en nube hay que definir de dónde se toman el vector, curvas, indicadores y NE que
> hoy están en `M:\`/INFOVALMER. Sin esos insumos no hay valoración (§11).

---

## 4. Modelos financieros a reimplementar (con base ya decodificada)

1. **Clasificación de activos** — resolución por clave (ISIN → nemo → Clas SFC) contra tablas de
   diccionario → Clase de Inversión + moneda + región + riesgo + frecuencia. (`INICIO`, `Diccionario SFC`).
2. **Enrutamiento de derivados** — por `Tipo de derivado` (col J de 415): forwards/futuros/swaps con
   mapeo fijo de columnas. (`BMO/Módulo1`, ya documentado 1:1).
3. **Valoración de swaps (VPN/DM)** — flujos descontados por curva según tipo (IRS/CCS, IBR/SOFR/
   LIBOR/DTF/UVR), day-count ACT/365·ACT/360·30/360, frecuencias AV/SV/TV/MV; duración Macaulay
   modificada por pata. (`Calculadora Swap`, `Detallado/Módulo2`).
4. **Sensibilidad de forwards (KRD)** — diferencias centrales ±100pb; separa KRD local/foráneo por
   pata; fwd con descuento simple base 360. (`Sensibilidad_Fwds`).
5. **EWMA** — volatilidad exponencial (S&P/Tesoros) — fórmula estándar (`EWMA MENSUAL`, no está entre
   los 5 pero es simple; ver manual paso 8k).
6. **Reconciliación/controles** — Valor de fondo MACRO vs SFC, número de títulos, eliminados,
   diferencias de precio atípicas (Top +/−). (`Parámetros` del BMO).

---

## 5. Dónde vive la lógica: VBA vs fórmulas de celda (el gap principal)

| Cálculo | ¿En VBA? | ¿En fórmulas de celda? | Acción |
|---|---|---|---|
| Enrutamiento derivados 415 | ✅ Sí (mapeo columnas) | — | Replicar directo. |
| Clasificación activos | Parcial (VLOOKUP orquestado) | ✅ Sí (tablas + fórmulas aux) | Extraer tablas + reglas. |
| VPN/DM swaps | Orquesta | ✅ **Sí** (hojas motor `SWAP*` cols D→O) | **Requiere .xlsx** para fórmulas exactas. |
| Sensibilidad KRD | Orquesta | ✅ **Sí** (`572`/`Ind`) | **Requiere .xlsx**. |
| Participaciones/consolidación Detallado | Parcial | ✅ Sí (W:AN arrastradas) | **Requiere .xlsx**. |
| Clasificación P2:AS2, Fwd AN/AQ, Swaps BX:CU | — | ✅ Sí | **Requiere .xlsx**. |

> **Conclusión:** el VBA da el **flujo y los mapeos**; los **coeficientes y fórmulas financieras**
> están en celdas. Para una réplica fiel al centavo hay que volcar esas fórmulas desde copias `.xlsx`
> de **BMO**, **Benchmark (Detallado)** y **Calculadora Swap** (las 3 con más fórmulas motor).

---

## 6. Rutas y parámetros (reemplazo de `M:\`)

- BMO: rutas `M:\` **literales** en hoja `Parámetros` (col C) → se vuelven **configuración**.
- INICIO / Sensibilidad / Detallado: rutas en **celdas con nombre** (hojas Control/Procedimiento/
  Parametros), no en VBA → también se vuelven configuración.
- Nombres con patrón de fecha: `AAAA_MMportainvdeta.xlsx`, `VPN/DM SWAP dd-mm-aaaa`, `Benchmark
  dd-mm-aaaa`, `Valoración NE dd-mm-aaaa`, `AAAA-MM-DD IND`. Convención mixta (mes con `0`, mes en
  español mayúsculas para carpetas de backup).

---

## 7. Puntos de juicio humano / instrumentos nuevos → **alertas**

Todos estos hoy son manuales; en la automatización se convierten en **entradas del archivo de alertas**
(no bloquean el resto del proceso, se marcan para revisión):

| # | Situación | Detección automática |
|---|---|---|
| 1 | **Activo con Clas SFC / ISIN no clasificado** (ND) | Clave no está en `CLASIFICACION`/`Diccionario SFC` → `#N/A`. |
| 2 | **Fondo Mutuo / ETF / Nota nueva** sin mapeo moneda/región | ISIN ausente en `CPRVI`/`TITULOS NUEVOS`. |
| 3 | **Código de derivado SFC nuevo** | Código J fuera de {1,4,5,6,16,17} o no en control SFC. |
| 4 | **Contraparte de swap nueva** | No está en diccionario de divisas. |
| 5 | **UVR de día no hábil** | Cierre en fin de semana/festivo → requiere valor manual. |
| 6 | **Paridad de divisa invertida** en forwards | Strike fuera de rango vs promedio/max/min. |
| 7 | **Nota estructurada borrada / en BRL** | Marca en control NE / moneda ≠ COP. |
| 8 | **Diferencia de precio atípica** | Top +/− diferencias sobre umbral. |
| 9 | **Descuadre de reconciliación** | Valor fondo MACRO ≠ SFC; nº títulos ≠. |

---

## 8. Automatizable vs juicio — panorama

| Bloque | Automatizable | Notas |
|---|---|---|
| Descarga SFC (351/415/CS-CP) | 🟢 Alta | Web alcanzable; parsear xlsx. |
| Preparación (renombrar, normalizar AFP, fechas W/X/Y) | 🟢 Alta | Trivial en pandas. |
| Importación/enrutamiento derivados | 🟢 Alta | Mapeo fijo ya decodificado. |
| Limpieza (<1000 COP, FA, eliminados) | 🟢 Alta | Reglas exactas conocidas. |
| Clasificación de activos | 🟡 Media | Determinista con diccionarios; los ND → alerta. |
| Valoración swaps (VPN/DM) | 🟡 Media | Modelo claro; requiere curvas + fórmulas exactas. |
| Sensibilidad forwards (KRD) | 🟡 Media | Modelo claro; requiere curvas. |
| Notas estructuradas | 🟡 Media | Parseo CSV + BRL→COP; notas nuevas → alerta. |
| EWMA | 🟢 Alta | Fórmula estándar. |
| Reconciliación / controles | 🟢 Alta | Comparaciones + umbrales. |
| Generación Excel | 🟢 Alta | openpyxl/xlsxwriter. |
| Generación PDF | 🟡 Media | Reemplazar Word por plantilla (reportlab/HTML→PDF). |
| Correos | 🟢 Alta | SMTP/API en vez de Outlook. |
| Detección instrumentos nuevos (alertas) | 🟢 Alta | Anti-join contra diccionarios. |

---

## 9. Riesgos

1. **Fórmulas de celda no visibles en VBA** → sin copias `.xlsx` la réplica de swaps/sensibilidad/
   consolidación sería por reverse-engineering aproximado. *(Mitiga: pedir 3 copias .xlsx.)*
2. **Insumos externos en `M:\`/INFOVALMER** (vector, curvas, indicadores, NE) → definir origen en nube.
3. **Convenciones frágiles heredadas**: parse de moneda por posición/longitud, filtros con padding
   (`"US$   "`), reglas hardcodeadas (ISINs especiales, NUAMCO/CAMACOL). *(Mitiga: portarlas explícitas + tests.)*
4. **Exactitud numérica**: day-count, overnight compuesto (SOFR/IBR), interpolación de curvas → validar
   contra un corte real (paridad al centavo).
5. **Dependencia de orden**: el proceso tiene precedencias (swaps antes de exportar, etc.) → el pipeline
   debe respetarlas y ser idempotente/reproducible.

---

## 10. Hallazgos de factibilidad en el entorno de ejecución (nube)

| Prueba | Resultado |
|---|---|
| Alcance web SFC | ✅ HTTP 302 (alcanzable → descarga automática viable). |
| Excel / `M:\` disponibles | ❌ No (confirmado; por eso reimplementación Python). |
| LibreOffice para `.xlsb→.xlsx` | ❌ Inoperativo (no carga ningún archivo). |
| Lectura de valores `.xlsb` (pyxlsb) | ✅ Sí. |
| Lectura de **fórmulas** `.xlsb` | ❌ No (pyxlsb solo da valores) → se necesitan copias `.xlsx`. |
| Extracción de VBA (olevba) | ✅ Sí (hecho, 41 módulos). |

---

## 11. Qué necesito para construir (insumos y decisiones)

**Insumos técnicos:**
1. **Copias `.xlsx`** (desde Excel: Guardar como) de: **Benchmark Mensual Optimizado**, **Benchmark
   (Detallado)** y **Calculadora Swap** → para extraer las fórmulas motor. *(Súbelas al bucket como
   los otros.)*
2. **Un ejemplo real de cada insumo externo** de un corte cerrado (ideal marzo/mayo 2026):
   `AAAA_MMportainvdeta.xlsx` (SFC), vector de precios, 1 juego de curvas swap, `AAAA-MM-DD IND.CSV`,
   `Valoración NE …csv`, y el `Diccionario RVI` — para tener el contrato de datos y validar cifras.
3. Acceso/mecanismo para obtener en producción el **vector, curvas, indicadores y NE** que hoy salen
   de INFOVALMER/`M:\` (¿API, SFTP, carpeta sincronizada, GCS?).

**Decisiones de producto:** ver `docs/ARQUITECTURA_PIPELINE.md` §Decisiones (formato de salida, canal
de correo, orquestación mensual, dónde viven los diccionarios de referencia).
