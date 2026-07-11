"""
Controles de reconciliacion -> alertas.

Reimplementa los controles de la hoja Parametros del BMO:
  - Valor de fondo por (AFP, portafolio) (base de conciliacion vs SFC)
  - Numero de titulos por fondo
  - Fondos con valor cero o negativo (senal de error)

La conciliacion "MACRO vs SFC" exacta requiere el total oficial del reporte SFC;
aqui se calcula el total desde el propio insumo y se marcan atipicos.
"""
from __future__ import annotations

import pandas as pd

from bmk.alerts.registry import RegistroAlertas
from bmk.consolidate.benchmark import valor_por_fondo


def controles(clasificados: pd.DataFrame, registro: RegistroAlertas) -> dict:
    fondos = valor_por_fondo(clasificados)

    # 1) Fondos con valor no positivo.
    malos = fondos[fondos["valor_fondo"] <= 0]
    n_malos = 0
    for _, r in malos.iterrows():
        registro.agregar("consolidacion", "FONDO_VALOR_NO_POSITIVO", "ERROR",
                         afp=r["afp"], portafolio=r["cod_portafolio"],
                         valor_mercado=r["valor_fondo"],
                         descripcion="Valor de fondo cero o negativo",
                         accion_sugerida="Revisar importacion de activos de ese portafolio")
        n_malos += 1

    # 2) Activos sin valor de mercado (posibles faltantes).
    sin_valor = clasificados[pd.to_numeric(clasificados["vr_mercado"], errors="coerce").fillna(0) == 0]
    n_sin = len(sin_valor)
    if n_sin:
        registro.agregar("consolidacion", "ACTIVOS_SIN_VALOR_MERCADO", "WARN",
                         valor_mercado=n_sin,
                         descripcion=f"{n_sin} activos con valor de mercado 0 (revisar vector/precio)",
                         accion_sugerida="Verificar precios en el vector para esos titulos")

    return {
        "fondos": fondos,
        "n_fondos": len(fondos),
        "fondos_valor_no_positivo": n_malos,
        "activos_sin_valor": n_sin,
        "valor_total_industria": float(fondos["valor_fondo"].sum()),
    }
