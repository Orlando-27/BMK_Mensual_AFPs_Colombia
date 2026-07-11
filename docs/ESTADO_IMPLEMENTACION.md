# Estado de implementación

Mapa honesto de lo construido y lo pendiente, para revisión. Verificado contra el corte real de
**junio 2026** descargado de la SFC.

## Implementado y funcionando (end-to-end)

| Fase | Módulo | Verificación en junio |
|---|---|---|
| F1 Ingesta SFC | `bmk/ingest/sfc.py` | Descarga junio de la web SFC (idFile) + lee 351/415/CSA. 24.694 activos, 4.532 derivados. |
| F2 Clasificación | `bmk/classify/classifier.py` | 99.9% auto-clasificado (3 niveles: ISIN/Nemo, fallback concat, leyenda). |
| F2 Detector + alertas | `bmk/classify/detector.py`, `bmk/alerts/registry.py` | 19 instrumentos nuevos + 6 códigos SFC nuevos → alertas Excel/JSON. |
| F3 Derivados | `bmk/derivatives/{router,forwards,swaps}.py` | 2615 forwards (compra/venta+paridad), 1518 swaps (por tipo), 72 futuros. |
| F3 Valoración swaps | `bmk/pricing/` (motor v6) | 1492 swaps valorados, \|error\|<2% vs Precia. |
| F4 Consolidación | `bmk/consolidate/{benchmark,reconcile}.py` | 594 billones COP; valor/fondo, pesos, por clase; controles. |
| F5 Salidas | `bmk/output/excel.py` | `consolidado_benchmark_*.xlsx` (6 hojas) + `alertas_*.xlsx/json` + `resumen.json`. |
| F6 Orquestación | `bmk/pipeline.py`, `bmk/run_monthly.py` | Corre todo en un comando; entrada mensual para cron. |
| Pruebas | `tests/test_logica.py` | 8 pruebas de la lógica determinista (pasan). |

## Pendiente (mapeado, no codificado) — para las próximas iteraciones

1. **Valoración de swaps de junio con curvas de junio.** Hoy se valora con las curvas del corte que
   está en los archivos (31-mar, según lo indicado). Para valorar junio se requiere la Calculadora
   Swap / curvas INFOVALMER de junio (extraíbles con `extraer_insumos.py`).
2. **Exposición detallada por región y moneda (CPRVI).** Las fórmulas del Detallado (`AD:AZ`) están
   decodificadas en `analisis/formulas/`; falta codificar el producto matricial peso-CPRVI × valor.
3. **Notas estructuradas (NE).** Insumo disponible (`condiciones_faciales_notas.csv`, payoffs y vols);
   falta el módulo de valoración y la conversión BRL→COP.
4. **Sensibilidad (KRD) y EWMA.** Modelos documentados (`analisis/logica/Sensibilidad_Fwds.md`, manual
   paso 8k); falta codificarlos en `consolidate/`.
5. **Plazos CORTO/MEDIANO/LARGO** (umbrales CP/LP) y **duración** por activo — fórmulas mapeadas.
6. **Salidas PDF + correos.** Diseñadas (WeasyPrint/SMTP); no codificadas.
7. **Cesantías Corto Plazo (CS-CP)** — paso 10 del manual; análogo al principal.
8. **Reconciliación exacta vs total oficial SFC** — hoy se calcula el total desde el propio insumo;
   falta cotejar contra la cifra oficial reportada.

## Notas de datos
- `bmk/config/reference/clasificacion.csv` es un **snapshot**; en producción crece a medida que se
  clasifican instrumentos nuevos (los ND del detector alimentan ese mantenimiento).
- Insumos externos (SFC, curvas, NE) se leen de almacenamiento configurable (reemplazo de `M:\`);
  ver `bmk/config/settings.py`.
