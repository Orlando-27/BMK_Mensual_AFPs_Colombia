# Valorador de Futuros de la Industria

Replica las hojas `FutLoc Industria` (futuros locales de TES) y `FutInt Industria`
(futuros internacionales) del Benchmark (Detallado). Output:
`valoracion_futuros_industria_<corte>.xlsx` (una hoja por tipo).

Fuente: Formato_415 del archivo SFC.
- `tipo_derivado = 5` -> futuros locales (FutLoc)
- `tipo_derivado = 6` -> futuros internacionales (FutInt)

## Futuros locales (FutLoc Industria) — validado al peso

Cada futuro referencia un subyacente (codigo de contrato TES, p. ej.
`COL17CT03342`) que tiene precio en el **VECTOR DE PRECIOS** del corte (forma
decimal: `0.92411` = 92,411%).

```
signo = +1 compra (posicion 1),  -1 venta (posicion 2)
Valor de Mercado = NOMINAL * precio_vector * signo
```

### Validacion (contra la hoja FutLoc Industria del propio archivo)

| Metrica | Resultado |
|---|---|
| Filas | 81 |
| Valor de Mercado dentro de 1 COP | **81 / 81** |
| Total Valor de Mercado | **+0,000000%** (0 COP de diferencia) |

## Futuros internacionales (FutInt Industria)

Segun las formulas de la hoja (indices y treasuries):

```
signo  = +1 compra, -1 venta
precio = nivel del indice o precio del treasury (VECTOR DE PRECIOS); treasury /100
Valor de Mercado (moneda) = NOMINAL * signo * precio
Valor de Mercado (COP)    = valor_moneda * FX(moneda del contrato)
```

> En el corte del archivo la hoja `FutInt Industria` esta **vacia** (sin futuros
> internacionales), por lo que este camino queda implementado segun formula pero
> **sin validar con datos**. En Formato_415 de abril hay 1 futuro internacional
> (indice `BRS2W3A52`) cuyo precio no esta en el VECTOR del placeholder; queda sin
> valorar hasta tener el precio del indice.

## Insumo

`VECTOR DE PRECIOS` extraido a `insumos/vector_precios.csv` (id, precio, moneda)
por `bmk.pricing.extraer_vector_precios`. Es **placeholder** (corte del .xlsb); se
reemplaza por el vector del corte real (mismo CSV). `insumos/` no se versiona.

## Pendiente

1. Reemplazar el VECTOR placeholder por el del corte real.
2. Validar FutInt con un corte que tenga futuros internacionales; incorporar el
   precio de indices al vector.
3. Los futuros de TRM (`tipo_derivado = 4`) se valoran con el motor de forwards
   (ver docs/VALORADOR_FWDS.md), no aqui.
