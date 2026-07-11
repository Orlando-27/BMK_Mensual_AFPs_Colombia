# BMK Mensual AFPs Colombia

Automatización del proceso mensual de actualización del **Benchmark Detallado** de portafolios de
pensión obligatoria (y cesantías) de las AFP de Colombia, a partir de la información publicada por la
Superintendencia Financiera de Colombia (SFC).

**Objetivo:** un proceso mensual **automático** que (1) descargue la información de la SFC,
(2) la procese siguiendo el manual, (3) genere el output (consolidados, PDFs, correos) y
(4) produzca un **archivo de alertas** ante fallos o instrumentos nuevos no contemplados.

## Estado

Pipeline **end-to-end funcionando** en Python (nube, sin Excel ni `M:\`), verificado contra el corte
real de **junio 2026** descargado de la SFC.

## Cómo correr

```bash
pip install -r requirements.txt

# Corte del mes (descarga de la SFC + procesa + genera salidas y alertas)
python3 -m bmk.run_monthly --mes Junio --anio 2026

# O procesando un insumo ya descargado, y valorando swaps con un corte de curvas:
python3 -m bmk.pipeline --archivo insumos/sfc/2026_06portainvdeta.xls \
    --swaps-insumos <carpeta con INICIO_*, SWAPIND_*, curva_*>
```

Salidas por corte en `salidas/AAAA-MM-DD/`:
- `consolidado_benchmark_*.xlsx` — valor por fondo, por clase de inversión, por clasificación,
  swaps por tipo, forwards por par, valoración de swaps.
- `alertas_*.xlsx` / `.json` — instrumentos nuevos, códigos SFC desconocidos, controles.
- `resumen.json` — bitácora y resumen del corte.

## Arquitectura del pipeline (`bmk/`)

| Etapa | Módulo | Función |
|---|---|---|
| Ingesta | `ingest/sfc.py` | Descarga y lee el Portafolio SFC (351/415/CSA). |
| Normalización | `prepare/normalize.py` | Extrae activos, normaliza AFP y portafolio. |
| Clasificación | `classify/classifier.py` | Resuelve clase/moneda/ubicación/riesgo (3 niveles). |
| Detector | `classify/detector.py` | Alerta instrumentos nuevos y códigos SFC nuevos. |
| Derivados | `derivatives/{router,forwards,swaps}.py` | Enruta 415, forwards, tipo de swap. |
| Valoración | `pricing/` | Motor de swaps v6 (validado <2% vs Precia). |
| Consolidación | `consolidate/{benchmark,reconcile}.py` | Posiciones, pesos, controles. |
| Salidas | `output/excel.py` | Consolidado Excel. |
| Alertas | `alerts/registry.py` | Registro transversal -> Excel/JSON. |
| Orquestación | `pipeline.py`, `run_monthly.py` | Encadena todo; entrada mensual. |

> El proceso manual completo, el diagnóstico y las fórmulas están documentados abajo.

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
