# Arquitectura propuesta — Pipeline Benchmark Mensual (Python)

> Diseño para automatizar el proceso mensual end-to-end en Python (nube, sin Excel/`M:\`).
> Basado en `analisis/DIAGNOSTICO.md`, `analisis/ESTRUCTURA_DATOS.md` y `analisis/logica/*.md`.
> **Aún es diseño** — no se ha escrito código de producción. Requiere validar las decisiones (§10)
> y recibir los insumos (§11 del diagnóstico) antes de construir.

---

## 1. Principios

1. **Python puro, sin Excel ni `M:\`.** Toda la lógica se reimplementa; las rutas se parametrizan.
2. **Determinista y reproducible.** Mismo insumo → mismo output. Corridas idempotentes.
3. **Config y datos de referencia versionados.** Los "diccionarios" de Excel (AFP, portafolios,
   Clas SFC→Clase, moneda/región) se vuelven tablas versionadas en el repo/almacén.
4. **Las excepciones no detienen el proceso: alimentan alertas.** Todo lo que hoy es juicio humano se
   detecta y se registra en un **archivo de alertas**; el pipeline continúa con lo que sí puede.
5. **Trazabilidad.** Cada etapa deja logs, controles de reconciliación y artefactos intermedios.

---

## 2. Componentes (módulos Python)

```
bmk/
├── config/               # Configuración (reemplaza hoja Parámetros y rutas M:\)
│   ├── settings.py       #   - fechas de corte, rutas de almacenamiento, umbrales de control
│   └── reference/        #   - datos de referencia (ex-diccionarios Excel), versionados:
│       ├── afp.csv               # normalización nombres AFP
│       ├── portafolios.csv       # SFC -> código (CS/PO/PC/PM/PR/ALTERNATIVO)
│       ├── clas_sfc.csv          # Clas SFC (+moneda) -> Clase de Inversión
│       ├── derivados_sfc.csv     # códigos tipo de derivado (1..26)
│       ├── clasificacion.csv     # ISIN/NEMO -> clase, moneda, región, riesgo, frecuencia
│       ├── cprvi.csv             # ISIN RVI/notas -> pesos región + moneda
│       └── curvas_swap.yml       # qué curva usa cada tipo de swap
├── ingest/               # (1) Adquisición de insumos
│   ├── sfc.py            #   descarga SFC 351/415 y CS-CP (web) -> DataFrames
│   ├── infovalmer.py    #   vector, curvas, indicadores (origen a definir)
│   └── notas.py         #   valoración notas estructuradas (CSV)
├── prepare/              # (2) Normalización
│   ├── normalize.py     #   nombres AFP, portafolios, fechas W/X/Y, tipos
│   └── schema.py        #   validación de esquema de columnas SFC (control "Columnas")
├── classify/             # (3) Clasificación de activos + detección de nuevos
│   ├── classifier.py    #   resolución ISIN->nemo->Clas SFC contra reference/
│   └── detector.py      #   genera alertas de instrumentos nuevos (ND / #N/A)
├── derivatives/          # (4) Derivados
│   ├── router.py        #   enruta 415 por 'Tipo de derivado' (1/4/5/6/16/17)
│   ├── forwards.py      #   forwards + paridades + futuros TRM
│   └── futures.py       #   FutLoc / FutInt (exposición, subyacente, duración)
├── pricing/              # (5) Valoración
│   ├── curves.py        #   construcción/interpolación de curvas
│   ├── swaps.py         #   VPN y DM (por pata; IRS/CCS; day-count)
│   └── sensitivity.py   #   KRD de forwards (±100pb)
├── consolidate/          # (6) Consolidación (ex Benchmark Detallado)
│   ├── benchmark.py     #   arma el consolidado industria + Colfondos
│   ├── ewma.py          #   volatilidad EWMA
│   └── reconcile.py     #   controles: valor fondo, nº títulos, eliminados, precios atípicos
├── output/               # (7) Salidas
│   ├── excel.py         #   consolidados .xlsx por fondo (openpyxl/xlsxwriter)
│   ├── pdf.py           #   PDFs por fondo (plantilla HTML->PDF / reportlab)
│   └── notify.py        #   correo (SMTP/API) con adjuntos
├── alerts/               # Sistema de alertas transversal
│   └── registry.py      #   acumula y escribe el archivo de alertas (Excel/JSON)
├── pipeline.py           # Orquestador: encadena etapas en orden, respeta precedencias
└── run_monthly.py        # Punto de entrada mensual (scheduler)
```

---

## 3. Configuración (reemplazo de `Parámetros` y rutas `M:\`)

Un único `settings` (env vars + YAML) con:
- **Corte:** `fecha_benchmark` (último día del mes), `fecha_publicacion` (día anterior, ej. 09).
- **Almacenamiento:** `raiz_insumos`, `raiz_salidas`, `raiz_referencia` (GCS/carpeta local en vez de `M:\`).
- **Fuentes:** URLs SFC, origen de vector/curvas/indicadores/NE.
- **Umbrales de control:** tolerancia de reconciliación, umbral de diferencias de precio, etc.
- **Correo:** remitente, destinatarios PO / PV, plantillas.

---

## 4. Datos de referencia (ex-diccionarios) — versionados

Se extraen una vez de los libros y se versionan como CSV/YAML; se actualizan cuando el negocio agrega
un mapeo (idealmente vía PR, con historial). Fuente inicial: hojas `Diccionario SFC`, `CLASIFICACION`,
`CPRVI`, `TITULOS NUEVOS` (ver `analisis/ESTRUCTURA_DATOS.md`).

> Cuando el detector encuentra un instrumento nuevo, además de alertar, puede **proponer** la fila
> nueva para que un humano la confirme y se agregue a `reference/` (ciclo de mantenimiento controlado).

---

## 5. Orden del pipeline (respetando precedencias del manual)

```
1. ingest.sfc            → 351, 415, CS-CP
2. prepare.normalize     → AFP/portafolios/fechas + validación de esquema (control Columnas)
3. classify              → clasificación de activos ─┐
   classify.detector     → alertas instrumentos nuevos│ (no bloquea)
4. derivatives.router    → forwards / futuros / swaps ─┘
5. ingest.infovalmer     → vector, curvas, indicadores
6. pricing.swaps         → VPN/DM (3 fechas)
   pricing.sensitivity   → KRD forwards
7. ingest.notas + consolidate.benchmark → arma consolidado (industria + Colfondos)
   consolidate.ewma      → EWMA
8. consolidate.reconcile → controles (valor fondo, nº títulos, precios atípicos)  → alertas
9. output.excel + output.pdf → entregables por fondo
10. output.notify        → correos (revisión interna + BMK a usuarios)
11. alerts.registry      → escribe archivo de alertas consolidado
```

Cada etapa: valida precondiciones, registra tiempos/controles y persiste su salida intermedia.

---

## 6. Sistema de alertas + detector de instrumentos nuevos (núcleo del requerimiento)

**Objetivo:** un **archivo de alertas** (Excel + JSON) por corrida, con severidad y contexto suficiente
para actuar sin abrir el proceso.

**Esquema de cada alerta:**
`corte · etapa · severidad {INFO|WARN|ERROR} · tipo · portafolio · AFP · ISIN · NEMO · Clas SFC ·
moneda · valor_mercado · descripcion · accion_sugerida`

**Tipos de alerta (mapa a §7 del diagnóstico):**
- `INSTRUMENTO_NUEVO_SIN_CLASIFICAR` (ND) — ISIN/Clas SFC no está en referencia.
- `FONDO_O_NOTA_SIN_MAPEO_MONEDA_REGION` — falta en CPRVI.
- `CODIGO_DERIVADO_DESCONOCIDO` — tipo J fuera de lo contemplado.
- `CONTRAPARTE_SWAP_NUEVA`, `PARIDAD_INVERTIDA`, `UVR_DIA_NO_HABIL`, `NOTA_BORRADA_O_BRL`.
- `DIFERENCIA_PRECIO_ATIPICA`, `DESCUADRE_RECONCILIACION`.
- `FALLO_ETAPA` — excepción técnica (insumo faltante, curva sin datos, timeout).

**Detección (determinista):** anti-join del universo SFC del mes contra las tablas de referencia +
validaciones de umbral. Si `severidad=ERROR` en un insumo crítico (p.ej. curva faltante), la etapa se
marca fallida pero el resto del pipeline sigue donde sea posible y el correo final lo refleja.

---

## 7. Salidas (entregables)

| Salida | Hoy | Reemplazo Python |
|---|---|---|
| Consolidado + por fondo | `.xlsb` (VBA) | `.xlsx` (openpyxl/xlsxwriter) |
| PDFs por fondo | Word ActiveX | plantilla HTML→PDF (WeasyPrint) o reportlab |
| Correos (revisión + BMK) | Outlook | SMTP o API de correo |
| **Archivo de alertas** | *(no existe hoy)* | Excel + JSON por corrida ⭐ nuevo |

---

## 8. Orquestación mensual

- **Disparador:** mensual el día 10 (o al detectar publicación SFC). Opciones en este entorno:
  un **trigger/cron** (Routine) que ejecuta `run_monthly.py`; reintentos con backoff si la SFC aún no
  publicó.
- **Idempotencia:** una corrida por corte; si se re-ejecuta, sobrescribe artefactos del mismo corte.
- **Observabilidad:** log estructurado + el archivo de alertas como resumen accionable.

---

## 9. Stack tecnológico propuesto

- **Datos:** `pandas`, `numpy`. **Excel:** `openpyxl`/`xlsxwriter` (escritura), `pyxlsb`/`openpyxl`
  (lectura). **Fechas/curvas:** `numpy`/`scipy` (interpolación), utilidades de day-count propias.
- **PDF:** `WeasyPrint` (HTML+CSS→PDF) o `reportlab`. **Correo:** `smtplib`/API.
- **Descarga:** `requests`/`httpx`. **Config:** `pydantic-settings` + YAML. **Tests:** `pytest`
  (validación contra un corte real, paridad numérica).

---

## 10. Decisiones a confirmar (producto)

| Tema | Opciones | Recomendación |
|---|---|---|
| Origen de insumos INFOVALMER en nube | GCS / SFTP / carpeta sincronizada / API | Definir con IT; GCS encaja con lo actual. |
| Formato de salida Excel | `.xlsx` (recomendado) vs replicar `.xlsb` | `.xlsx`. |
| PDF | WeasyPrint (HTML) vs reportlab | WeasyPrint (más fácil de estilizar como los actuales). |
| Correo | SMTP corporativo vs API | El que permita IT. |
| Dónde viven los diccionarios de referencia | Repo (CSV, con PR) vs base de datos | Repo al inicio; DB si crece. |
| Orquestación | Routine/cron en este entorno vs Cloud Scheduler | Empezar con cron; migrar si se productiviza. |
| Alcance de "auto-clasificar" nuevos | Solo alertar vs proponer clasificación | Alertar + proponer; confirma humano. |

---

## 11. Roadmap de construcción (por fases, tras aprobar diseño)

1. **F0 – Contrato de datos:** extraer fórmulas de las 3 copias `.xlsx` + fijar esquemas de insumos con
   ejemplos reales de un corte cerrado. Construir `reference/` desde los diccionarios.
2. **F1 – Ingesta + preparación:** descarga SFC + normalización + validación de esquema. (Salida:
   DataFrames limpios + primeras alertas de esquema.)
3. **F2 – Clasificación + detector de nuevos:** clasificar activos y emitir alertas de ND. (Entrega
   temprana de valor: el archivo de alertas de instrumentos nuevos.)
4. **F3 – Derivados + valoración:** enrutamiento, VPN/DM swaps, KRD forwards (contra curvas reales).
   Validar paridad numérica vs un corte histórico.
5. **F4 – Consolidación + controles:** Detallado + EWMA + reconciliación.
6. **F5 – Salidas:** Excel + PDF + correos + archivo de alertas final.
7. **F6 – Orquestación:** cron mensual, reintentos, observabilidad, pruebas end-to-end.

Cada fase entrega algo ejecutable y verificable contra un corte real.
