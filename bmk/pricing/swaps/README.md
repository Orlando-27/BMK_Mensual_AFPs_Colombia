# Valorador de Swaps OTC v6 — Motor Precia

Motor de valoración que replica la metodología Precia PPV para mark-to-market de swaps OTC colombianos. Validado contra 1,492 contratos a 31-Marzo-2026 con error máximo < 2 % vs VPN Precia oficial.

## Contenido del paquete

```
v6_entrega/
├── README.md                    ← este archivo
├── CAMBIOS_V6.md                ← bitácora de cambios + pendientes de revisión
├── extraer_insumos.py           ← extractor desde .xlsb de Precia
├── swap_valuator_v6.py          ← motor de valoración
├── insumos_20260331/            ← insumos para la fecha validada
│   ├── INICIO_20260331.csv         (FECHA_VAL, TRM, UVR, EURCOP)
│   ├── SWAPIND_20260331.csv        (portafolio 1,492 swaps)
│   ├── VPN_PRECIA_20260331.csv     (benchmark oficial de Precia)
│   └── curva_*_20260331.csv        (12 curvas cero cupón)
├── salida_inicio.csv            ← output formato rango B:J de INICIO
└── salida_full.csv              ← output extendido (precios, DM, convexidad, errores)
```

## Requisitos

```bash
pip install pandas numpy scipy python-dateutil pyxlsb
```

## Uso para una fecha nueva

**Paso 1.** Extraer los insumos desde el .xlsb de Precia de esa fecha:

```bash
python3 extraer_insumos.py \
    --xlsb  /path/al/Calculadora_Swap_N_DD-MM.xlsb \
    --fecha AAAA-MM-DD \
    --out   insumos_AAAAMMDD/
```

**Paso 2.** Correr el motor. Autodetecta la fecha desde los archivos `INICIO_*.csv`:

```bash
V6_INSUMOS_DIR=insumos_AAAAMMDD \
V6_OUTPUT_CSV=salida_inicio.csv \
V6_OUTPUT_FULL=salida_full.csv \
python3 swap_valuator_v6.py
```

No hay que editar el código. Todo lo sensible a la fecha (fecha de valoración, TRM, UVR, EURCOP, nombres de archivos de curvas) se resuelve dinámicamente desde los CSVs de insumos.

## Formatos de salida

**`salida_inicio.csv`** replica el rango B:J de la hoja INICIO del Excel de Precia:

| # SWAP | DERECHO | OBLIGACIÓN | VPN | LLAVE_DER | DER | LLAVE_OBLIG | OBLIG |
|--------|---------|------------|-----|-----------|-----|-------------|-------|
| Col913412 | 7529167228.64 | 7553617673.32 | -24450444.68 | SWAPCol913412D | 1.0285 | SWAPCol913412O | 1.0319 |

**`salida_full.csv`** añade duración modificada, convexidad, errores vs Precia y curvas usadas por pata.

## Resultado de validación (31-Marzo-2026)

Todos los 1,492 swaps dentro de 2 % de error vs VPN Precia en ambas patas:

| Tipo | n | máx DER | máx OBL |
|------|---|---------|---------|
| IRS SOFUSD | 787 | 0.93 % | 0.47 % |
| IRS IBRCOP | 510 | 1.59 % | 1.45 % |
| CCS COPUSD | 191 | 1.89 % | 0.91 % |
| CCS COPUVR | 4 | 0.03 % | 0.02 % |

## Pendientes de revisión

Ver sección **⚠️ Pendientes de revisión** al inicio de `CAMBIOS_V6.md`. El más importante: el tratamiento del spread SOF1 = 0.26 % (~140 contratos) donde hay disonancia entre el contrato económico (mesa) y el MTM Precia.
