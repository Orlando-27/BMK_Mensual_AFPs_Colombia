# BMK Mensual AFPs Colombia

Automatización del proceso mensual de actualización del **Benchmark Detallado** de portafolios de
pensión obligatoria (y cesantías) de las AFP de Colombia, a partir de la información publicada por la
Superintendencia Financiera de Colombia (SFC).

**Objetivo:** un proceso mensual **automático** que (1) descargue la información de la SFC,
(2) la procese siguiendo el manual, (3) genere el output (consolidados, PDFs, correos) y
(4) produzca un **archivo de alertas** ante fallos o instrumentos nuevos no contemplados.

## Estado

Fase de **análisis y diseño** (no hay código de producción todavía). Se reimplementará en Python,
para correr en nube (sin Excel ni unidad de red `M:\`).

## Documentación

| Documento | Contenido |
|---|---|
| [`docs/PROCESO_BENCHMARK_MENSUAL.md`](docs/PROCESO_BENCHMARK_MENSUAL.md) | El proceso manual completo (10 bloques), rutas, controles, decisiones. |
| [`analisis/DIAGNOSTICO.md`](analisis/DIAGNOSTICO.md) | Diagnóstico integral: arquitectura actual, modelos financieros, riesgos, factibilidad. |
| [`docs/ARQUITECTURA_PIPELINE.md`](docs/ARQUITECTURA_PIPELINE.md) | Arquitectura propuesta del pipeline Python + roadmap de construcción. |
| [`analisis/ESTRUCTURA_DATOS.md`](analisis/ESTRUCTURA_DATOS.md) | Esquema de columnas SFC, tablas de mapeo, detección de instrumentos nuevos. |
| [`analisis/HALLAZGOS_VBA.md`](analisis/HALLAZGOS_VBA.md) | Destilado inicial del VBA (enrutamiento de derivados, normalización). |
| [`analisis/logica/`](analisis/logica/) | Mapa de lógica macro por macro de los 5 archivos Excel. |
| [`analisis/vba/`](analisis/vba/) | Código VBA extraído (41 módulos) — especificación de referencia. |

## Herramientas

- `tools/inspect_excel.py` — extrae VBA e inventario de hojas de archivos `.xlsb/.xlsm/.xlsx`.

> Los archivos Excel originales (insumos pesados) **no se versionan** (ver `.gitignore`).
