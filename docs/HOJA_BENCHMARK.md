# Output: hoja "Benchmark" para la herramienta diaria

La herramienta de actualización diaria es el archivo **Benchmark (Detallado)**. Su hoja `Benchmark`
recibe los activos consolidados de la industria; el resto de columnas son fórmulas que ya existen en
el archivo y se **arrastran**. Este output genera esa hoja con **solo los datos**, dejando las columnas
formuladas vacías.

Se genera en cada corrida: `salidas/AAAA-MM-DD/hoja_Benchmark_AAAA-MM-DD.xlsx`.

## Diagnóstico de la hoja (encabezado en fila 14)

| Col | Encabezado | Tipo | Origen en el pipeline |
|---|---|---|---|
| A | Portafolio | dato | código PO/PC/PM/PR/CS |
| B | AFP | dato | AFP normalizada |
| C | IDENTIFICADOR | dato | ISIN › Nemo › Emisor |
| D | Nemo | dato | Nemotécnico (351) |
| E | ISIN | dato | ISIN (351) |
| F | Clas SFC | dato | Clase de Inversión SFC (código) |
| G | Emisor | dato | Razón Social Emisor |
| H | F.Compra | dato | Fecha Compra |
| I | F.Vcto | dato | Fecha Vencimiento |
| J | Moneda | dato | Código moneda (351) |
| K | Valor Nominal | dato | Valor Nominal |
| **L** | **Vr Mercado INICIAL** | **fórmula** | *(vacía — se arrastra)* |
| M | TASA FACIAL | dato | Tasa Facial (título) |
| N | CLASE DE INVERSION | dato | clasificación (clase) |
| O | MONEDA | dato | moneda |
| P | TASA FACIAL | (vacía) | — |
| Q | UBICACION | dato | NACIONAL / INTERNACIONAL |
| R | CLASIFICACIÓN | dato | R FIJA / R VARIABLE / CAJA / … |
| S | RIESGO | dato | riesgo |
| T | VALOR TASA FACIAL | dato | Tasa de negoc |
| U | Precio inicial | dato (proxy) | Vr mercado / Valor nominal |
| V | Periodicidad | (vacía) | — |
| W | Precio Actual | (vacía) | — (vector diario) |
| **X … BU** | valor mercado, rentabilidad, participación, duración, sensibilidad, plazos, **exposición por moneda (AD-AR) y región (AS-AZ)**, notas (BE-BL), DM, cupones, % participación | **fórmula** | *(vacías — se arrastran)* |

## Uso

1. Corre el pipeline (o `run_monthly`). Se genera `hoja_Benchmark_*.xlsx`.
2. Copia las **columnas de datos** (A-K, M-W) y pégalas en la hoja `Benchmark` del archivo diario,
   debajo de la última fila de activos.
3. Arrastra las fórmulas (L y X en adelante) en el archivo diario — ahí ya existen.

## Notas / a validar contigo
- **Precio inicial (U)** es un *proxy* = Vr mercado / Valor nominal. En el proceso real viene del
  **vector de precios**; si prefieres, se reemplaza cuando se integre el vector del corte. Poblado ~99%
  (falta en caja/depósitos sin nominal).
- **Precio Actual (W)** y **Periodicidad (V)** se dejan vacías (dependen del vector / clasificación de
  frecuencia); se pueden llenar si se requiere.
- Se incluyen **todos** los activos consolidados (industria completa, todas las AFP), en el orden del
  archivo SFC.
- Las columnas formuladas se dejan **vacías** para no sobrescribir las fórmulas del archivo destino.
