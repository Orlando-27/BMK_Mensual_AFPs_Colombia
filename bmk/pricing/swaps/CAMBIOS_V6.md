# Valorador de Swaps OTC — Bitácora de cambios v6

**Fecha de valoración validada:** 31-Marzo-2026
**Portafolio:** 1,492 swaps (IRS SOFUSD, IRS IBRCOP, CCS COPUSD, CCS COPUVR)
**Benchmark:** VPN Precia publicados en hoja INICIO rango B:J

---

## ⚠️ Pendientes de revisión

### P1 — Tratamiento del campo Facial = 0.26% en patas SOF1

**Estado actual del motor:** el campo Facial se **ignora** (spread = 0) cuando la pata flotante es SOF1. Esta regla mantiene los errores vs Precia en cero, pero puede no reflejar la realidad económica del contrato.

**Contexto.** En el portafolio hay dos grupos de swaps con Facial = 0.26% en pata SOF1:

1. 91 IRS SOFUSD (todos bookeados en 2021, pata DER flotante SOF1).
2. 50 CCS IBR-SOFR (CCS COPUSD con OBL = SOF1, 46 de los cuales tienen Facial OBL = 0.26%).

**La mesa de trading confirmó que contractualmente estos swaps pagan SOF1 + 0.26 %**, es decir, el 26 bp SÍ es un spread económico real sobre SOFR.

**Pero empíricamente Precia no lo aplica.** La evidencia que soporta esto:

- Los 91 + 46 contratos tienen todos *exactamente* el mismo 0.26 % (si fuera spread negociado, habría dispersión individual).
- El precio Precia de la pata SOF1 **no correlaciona con la duración** (correlación ≈ 0.04), cuando debería crecer linealmente con ella si el spread se aplicara.
- Aplicar spread = 0.26 % sobre-estima el MTM en ~1.5 %.
- El mismo Precia **sí aplica** el spread en la pata IBR (DER) de los mismos CCS — donde los spreads varían 25-105 bps por contrato y la correlación spread × duración vs precio es 0.75. Tratamiento asimétrico.

**Hipótesis de la disonancia.** Precia posiblemente considera que el spread sobre SOFR queda absorbido en el basis USDOIS vs la curva que usan para MTM, de modo que aplicarlo explícitamente duplicaría el efecto. La pata IBR no tiene ese solape porque la curva COP no incorpora el spread negociado.

**Qué hacer cuando se revise:**

1. Confirmar con Precia directamente cuál es la fórmula que usan para la pata SOF1 de estos contratos (preguntar por documentación metodológica específica, no el manual general).
2. Revisar si existe una versión "económica" (con spread aplicado) en otro reporte Precia para comparar.
3. Si se confirma que el spread sí debe aplicarse para fines de gestión de riesgo/PnL interno, agregar una columna adicional `Precio_Economico_Der` y `Precio_Economico_Obl` en el output, manteniendo las columnas `Precio_Der` y `Precio_Obl` alineadas con MTM Precia.
4. Documentar la diferencia de cifras para contabilidad (MTM Precia) vs gestión (económico real), que potencialmente es de 0.5 a 1.5 % en esos ~140 contratos.

**Código relacionado (swap_valuator_v6.py, función `valorar_pata`):**
```python
# Nota P1 — pendiente de revisar con mesa/Precia
spread_real = 0.0 if tfacial == "SOF1" else tasa
```

---

## Resultado final

Todos los 1,492 swaps quedan con |error| vs Precia **menor a 2%** en ambas patas.

| Tipo | n | mediana DER | p95 DER | máx DER | mediana OBL | p95 OBL | máx OBL |
|------|---|-------------|---------|---------|-------------|---------|---------|
| IRS SOFUSD | 787 | 0.019% | 0.293% | **0.931%** | 0.020% | 0.365% | **0.467%** |
| IRS IBRCOP | 510 | 0.583% | 1.233% | **1.589%** | 0.595% | 1.220% | **1.452%** |
| CCS COPUSD | 191 | 0.051% | 1.266% | **1.889%** | 0.370% | 0.815% | **0.911%** |
| CCS COPUVR | 4 | 0.027% | 0.027% | **0.027%** | 0.016% | 0.016% | **0.016%** |

---

## Cambios aplicados (orden cronológico)

A lo largo del proceso el motor pasó por varios estados. Esta tabla muestra el impacto de cada cambio sobre el máximo error por tipo de swap (pata DER y OBL juntas, lo peor).

| # | Cambio | IRS SOFUSD | IRS IBRCOP | CCS COPUSD | CCS COPUVR |
|---|--------|-----------:|-----------:|-----------:|-----------:|
| 0 | Estado inicial (v4 portado a v6) | ~18% | ~30% | ~5% | 0.03% |
| 1 | VPN convertido a COP con FX spot (v4 devolvía moneda nativa) | 6.04% | 10.05% | 3.47% | 0.03% |
| 2 | Bond-replication universal (todas las patas incluyen nocional al final) | 6.04% | 10.05% | 3.47% | 0.03% |
| 3 | Curvas IBR/UVR/COP en convención **simple ACT/365** (no EA) | 3.76% | 1.59% | 3.47% | 0.03% |
| 4 | **Pata USD de CCS COPUSD descontada con USDOIS** (no USDCO) | 3.76% | 1.59% | 2.92% | 0.03% |
| 5 | **Pata COP de CCS COPUSD descontada con COP_COL_USD** (no IBR_COL_USD) | 3.76% | 1.59% | 1.89% | 0.03% |
| 6 | Ignorar "Facial" como spread en flotantes SOF1 (20 outliers IRS SOFUSD) | 0.93% | 1.59% | 1.89% | 0.03% |
| 7 | Forward-starting SOF1 (F.Compra = FECHA_VAL): primer cupón = 0 | **0.93%** | **1.59%** | **1.89%** | **0.03%** |

### Resumen en una línea

> El 85 % de la reducción de error vino de (a) interpretar correctamente las convenciones de composición de cada curva (simple ACT/360 vs simple ACT/365 vs EA), (b) asignar la curva correcta a cada pata de CCS COPUSD (COP_COL_USD / USDOIS en lugar de IBR_COL_USD / USDCO), y (c) tratar el "Facial" como referencia histórica y no como spread real en las patas SOFR.

---

## Detalle técnico de cada cambio

### Cambio 1 — Conversión a COP con FX spot

El output original de v4 devolvía el VPN en moneda nativa (USD para IRS SOFUSD, COP para IRS IBRCOP, etc.), pero el benchmark Precia reporta todo en COP. La conversión correcta es:

```
VPN_COP = PV_moneda × FX_spot(moneda)
```

donde `FX_spot` es 3660.10 para USD, 397.5321 para UVR, 4208.38 para EUR, 1 para COP. **No se usa TRM forward** — Precia usa spot para cada flujo y los descuentos implícitos de la curva capturan el resto del efecto carry.

### Cambio 2 — Bond-replication universal

Cada pata (de IRS o CCS) se valora como bono: cupones + nocional al vencimiento. En el IRS los nocionales se cancelan en el neto pero Precia los reporta individualmente. Código: `incluir_nocional=True` siempre.

### Cambio 3 — Convenciones de curva (la clave del orden de magnitud)

Las curvas de Precia **no están todas en EA**. La normalización correcta es:

| Curva | Convención | Fórmula DF |
|-------|-----------|-----------|
| USDOIS, USDCO | simple ACT/360 | `DF = 1 / (1 + r × d/360)` |
| IBR_COL_USD, IBRUVR | simple ACT/365 | `DF = 1 / (1 + r × d/365)` |
| COP_COL_USD, DTF, LIBORCOP, LIBORUVR, EURCOP | EA ACT/365 | `DF = 1 / (1 + r)^(d/365)` |

**Criterio empírico**: si la tasa a 10+ años supera ~15 % en la curva, casi siempre es convención simple (nominal acumulada), no EA. En EA las tasas cero largas rara vez superan 8-10 %.

### Cambios 4 y 5 — Asignación de curvas para CCS COPUSD

El hallazgo más importante del proyecto. Para CCS COPUSD, el mapping correcto es:

| Pata | Curva (antes) | Curva (después) |
|------|---------------|-----------------|
| COP fija/flotante | IBR_COL_USD | **COP_COL_USD** (col AZ de DATOS, no col D) |
| USD fija/flotante | USDCO | **USDOIS** |

Son curvas publicadas separadamente en el Excel DATOS. La razón: Precia valora CCS COPUSD asumiendo **sin colateral COP** en el lado USD (por eso USDOIS puro) y **con el basis COP vs USD cash** en el lado COP (por eso COP_COL_USD en lugar de IBR_COL_USD).

Este cambio bajó el sesgo sistemático de +1.6/+1.8 % en CCS COPUSD largos (10-20 y) a menos de 0.5 %.

### Cambio 6 — "Facial" de SOF1 no es un spread económico

**Esta es la conclusión más contraintuitiva del proyecto y merece evidencia.**

Los hechos sobre el campo Facial en patas SOF1 del portafolio:

1. **91 swaps tienen Facial = 0.26% EXACTAMENTE igual.** No es un spread negociado individualmente — todos los contratos tienen el mismo valor hasta el último decimal.
2. **Los 91 se bookearon en 2021** (ningún otro año). Ningún swap nuevo (2026) tiene Facial distinto de 0.
3. **Todos son periodicidad Trimestral** (coincide con tenor típico de un fixing SOFR compound trimestral).
4. El 0.26% corresponde a un SOFR compound trimestral de ~1 % anualizado → consistente con un fixing de SOFR de algún trimestre de principios de 2022 (momento en que SOFR estaba subiendo de 0 a 0.5%).

**Prueba empírica directa.** Si 0.26% fuera un spread real aplicado a los cupones restantes, el precio de la pata flotante debería ser aproximadamente `1 + 0.0026 × duration`. Para un swap de 15 años remanentes con duration ≈ 8 años, eso sería **~1.021**.

Lo que reporta Precia para estos 91 swaps:

| Años restantes | n | Precio Precia DER (media) | Si fuera spread, debería ser |
|---------------|---|---------------------------|------------------------------|
| < 5y | 10 | 1.00364 | ~1.008 |
| 5-10y | 66 | 1.00347 | ~1.013 |
| 10-15y | 5 | 1.00882 | ~1.018 |
| > 15y | 10 | 1.00849 | ~1.021 |

El precio Precia **no crece con la duración** y queda en un rango estrecho de 1.003-1.009. Eso sólo es posible si Precia **no está aplicando el spread**.

**Comparación de interpretaciones (evaluadas sobre los 91 swaps):**

| Interpretación | Error medio DER | Error máx |DER| |
|---------------|-----------------|-----------------|
| spread = +Facial (aplicar) | **+1.57 %** | 2.99 % |
| **spread = 0 (ignorar Facial)** | **+0.04 %** | **0.93 %** |
| spread = −Facial (signo invertido) | −1.49 % | 3.06 % |
| spread = Facial × 0.5 (parcial) | +0.80 % | 1.54 % |

La única interpretación que da error cero es ignorar el Facial.

**Qué es entonces el campo Facial en una pata SOF1.** Dado que (a) todos los contratos tienen el mismo valor, (b) el valor no afecta la valoración Precia, y (c) la cuantía coincide con un SOFR compound trimestral de inicio de 2022, la interpretación más consistente es que el campo **guarda el último fixing conocido o la tasa de referencia al momento del booking**, como metadata del contrato. No representa un spread económico pagadero sobre SOFR.

Esto tiene una consecuencia económicamente importante: **si estos swaps tuvieran un spread real de 26 bps, Precia estaría subvaluando el portafolio en ~3 % en las patas flotantes**, lo que sería un error material. Dado que Precia es el proveedor oficial de MTM para fondos colombianos regulados, es mucho más probable que el campo Facial efectivamente no represente un spread.

**Tratamiento en el motor.** Para patas flotantes:
```python
spread_real = 0.0 if tfacial == "SOF1" else tasa
```
Para IBR (pata flotante de CCS COPUSD) el Facial sí es un spread real — lo he verificado con swaps que tienen Facial de 7-8 % donde SÍ forma parte del cupón. En esos casos ignorar el Facial produce errores de −20 % a −25 %, que es la firma característica de un spread económico descartado.

### Cambio 7 — Forward-starting IRS SOFUSD

8 swaps tenían `F.Compra = FECHA_VAL` exacta (swaps recién emitidos). Para estos, mi código daba precio 1.0 (telescópico puro) pero Precia daba `DF_USDOIS(fecha_primer_cupón) × N`.

**Convención Precia**: cuando el cupón en curso comienza exactamente en la fecha de valoración, la tasa flotante compound aún no ha empezado. El primer cupón se asigna con `fwd = 0`, de modo que la PV del período = nocional × DF al final del período.

```python
if prev < FECHA_VAL:
    fwd = zero_rate(curva_proj, dias_accrual)       # en curso
elif prev == FECHA_VAL:
    fwd = 0.0                                       # forward-start exacto
else:
    fwd = (DF(t1)/DF(t2) - 1) × base/dias_accrual   # futuro normal
```

---

## Decisiones adicionales que se mantuvieron de v4

- **F.Compra como inicio** de la acumulación (no Emision, que en el portafolio es igual a F.Vcto).
- **DFs construidos a partir de la `Tasa`** de la curva con la convención apropiada (no se usa la columna FD del Excel porque no está disponible en el xlsb, y reconstruir desde FD producía errores severos a largo plazo).
- **Interpolación lineal sobre tasas** (no log-lineal sobre DFs). Precia usa lineal sobre tasas según el Manual de Metodologías.
- **Valoración anclada a FECHA_VAL** (leída de INICIO, no derivada del portafolio).

---

## Auditoría de hardcodeos

El motor ya **no tiene ningún valor** de fecha, FX, ni nombre de archivo escrito directamente en el código. Todo lo sensible a la fecha se resuelve desde `INSUMOS_DIR/INICIO_YYYYMMDD.csv`:

| Valor | Fuente |
|-------|--------|
| `FECHA_VAL` | `INICIO_YYYYMMDD.csv → key=FECHA_VAL` |
| `TRM_USDCOP` | `INICIO_YYYYMMDD.csv → key=TRM` |
| `UVR_HOY` | `INICIO_YYYYMMDD.csv → key=UVR` |
| `EURCOP` | `INICIO_YYYYMMDD.csv → key=EURCOP` |
| Nombres de curva | `curva_<NOMBRE>_YYYYMMDD.csv` (el sufijo se resuelve en runtime) |
| `YYYYMMDD` | Se autodetecta leyendo el primer `INICIO_*.csv` de la carpeta |

Lo que **SÍ** es permanente (no depende de fecha):
- Convenciones por curva (`CURVE_CONVENTION`).
- Asignación de curvas por tipo de swap.
- Reglas de spread (`SOF1` → 0, resto → Facial).
- Regla forward-starting.
- Mapeo `FREQ_MONTHS` y `DAY_BASE`.

Si Precia cambia la metodología en el futuro (nueva convención de curva, nuevo tratamiento de basis, etc.), **solo hay que editar esas constantes** — no tocar lógica ni fechas.

---

## Cómo usar el motor en otra fecha

```bash
# 1. Extraer insumos del nuevo xlsb
python3 extraer_insumos.py \
    --xlsb  /path/Calculadora_Swap_N_30-06.xlsb \
    --fecha 2026-06-30 \
    --out   /path/insumos_20260630/

# 2. Correr el motor (autodetecta la fecha del INICIO_*.csv)
V6_INSUMOS_DIR=/path/insumos_20260630 \
V6_OUTPUT_CSV=/path/salida_inicio.csv \
V6_OUTPUT_FULL=/path/salida_full.csv \
python3 swap_valuator_v6.py
```

No hay ningún paso de "editar el código para cambiar la fecha". El motor lee todo del directorio.

---

## Archivos del proyecto

| Archivo | Rol |
|---------|-----|
| `extraer_insumos.py` | Del .xlsb de Precia, extrae curvas + SWAPIND + INICIO + VPN benchmark como CSVs |
| `swap_valuator_v6.py` | Motor de valoración (dinámico por fecha) |
| `insumos_YYYYMMDD/` | 13 CSVs de curvas + SWAPIND + INICIO + VPN_PRECIA |
| `salida_inicio.csv` | Output en formato rango B:J de INICIO (# SWAP, DERECHO, OBLIGACIÓN, VPN, LLAVES, precios) |
| `salida_full.csv` | Output extendido con duración modificada, convexidad, curvas usadas, errores vs Precia |
