"""
Configuracion del pipeline (reemplaza la hoja Parametros y las rutas M:\\).

Los valores se pueden sobreescribir por variables de entorno BMK_*.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]


@dataclass
class Config:
    # Corte
    anio: int = int(os.environ.get("BMK_ANIO", "2026"))
    mes: str = os.environ.get("BMK_MES", "Junio")

    # Almacenamiento (reemplazo de M:\)
    dir_insumos: Path = field(default_factory=lambda: Path(os.environ.get("BMK_INSUMOS", RAIZ / "insumos")))
    dir_salidas: Path = field(default_factory=lambda: Path(os.environ.get("BMK_SALIDAS", RAIZ / "salidas")))
    dir_referencia: Path = field(default_factory=lambda: RAIZ / "bmk" / "config" / "reference")

    # Insumos de valoracion de swaps (Calculadora Swap / motor v6)
    dir_swaps_insumos: Path | None = None  # carpeta con INICIO_*, SWAPIND_*, curva_*

    # Umbrales de control
    umbral_diferencia_precio: float = 0.10  # 10% -> alerta de precio atipico

    @property
    def corte(self) -> str:
        from bmk.ingest.sfc import MES_NUM
        m = MES_NUM[self.mes.lower()]
        # Ultimo dia del mes de corte
        import calendar
        d = calendar.monthrange(self.anio, m)[1]
        return f"{self.anio}-{m:02d}-{d:02d}"

    def dir_corte(self) -> Path:
        return self.dir_salidas / self.corte
