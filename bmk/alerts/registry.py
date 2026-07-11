"""
Registro de alertas transversal del pipeline.

Acumula alertas de todas las etapas y las escribe a Excel + JSON. Cada alerta
NO detiene el proceso: se registra para revision humana (instrumentos nuevos,
descuadres, fallos de etapa, etc.).
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path

import pandas as pd

SEVERIDADES = ("INFO", "WARN", "ERROR")

CAMPOS = ["corte", "etapa", "severidad", "tipo", "portafolio", "afp", "isin",
          "nemo", "clas_sfc", "moneda", "valor_mercado", "descripcion", "accion_sugerida"]


@dataclass
class RegistroAlertas:
    corte: str = ""
    alertas: list[dict] = field(default_factory=list)

    def agregar(self, etapa: str, tipo: str, severidad: str = "WARN", *,
                portafolio="", afp="", isin="", nemo="", clas_sfc="", moneda="",
                valor_mercado=None, descripcion="", accion_sugerida="") -> None:
        if severidad not in SEVERIDADES:
            severidad = "WARN"
        self.alertas.append({
            "corte": self.corte, "etapa": etapa, "severidad": severidad, "tipo": tipo,
            "portafolio": portafolio, "afp": afp, "isin": isin, "nemo": nemo,
            "clas_sfc": clas_sfc, "moneda": moneda, "valor_mercado": valor_mercado,
            "descripcion": descripcion, "accion_sugerida": accion_sugerida,
        })

    def agregar_lote(self, df: pd.DataFrame, etapa: str, tipo: str, severidad: str,
                     descripcion: str, accion_sugerida: str, colmap: dict) -> int:
        """Agrega una alerta por fila de df, tomando campos de colmap {campo: col_df}."""
        n = 0
        for _, r in df.iterrows():
            kw = {k: (r[v] if v in df.columns else "") for k, v in colmap.items()}
            self.agregar(etapa, tipo, severidad, descripcion=descripcion,
                         accion_sugerida=accion_sugerida, **kw)
            n += 1
        return n

    def to_dataframe(self) -> pd.DataFrame:
        if not self.alertas:
            return pd.DataFrame(columns=CAMPOS)
        return pd.DataFrame(self.alertas)[CAMPOS]

    def resumen(self) -> dict:
        df = self.to_dataframe()
        if df.empty:
            return {"total": 0}
        out = {"total": len(df)}
        out["por_severidad"] = df["severidad"].value_counts().to_dict()
        out["por_tipo"] = df["tipo"].value_counts().to_dict()
        return out

    def escribir(self, out_dir: str | Path, nombre: str = "alertas") -> dict:
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        df = self.to_dataframe()
        xlsx = out_dir / f"{nombre}_{self.corte}.xlsx"
        js = out_dir / f"{nombre}_{self.corte}.json"
        df.to_excel(xlsx, index=False)
        js.write_text(json.dumps({"corte": self.corte, "resumen": self.resumen(),
                                  "alertas": self.alertas}, ensure_ascii=False, indent=2,
                                 default=str), encoding="utf-8")
        return {"xlsx": str(xlsx), "json": str(js), "n": len(df)}
