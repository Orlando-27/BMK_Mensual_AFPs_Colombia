"""
Extrae los insumos de la hoja `Curvas` del Benchmark (Detallado) a CSVs
swappables, para que el valorador de forwards no dependa del .xlsb.

La hoja `Curvas` apila en A:E varias curvas por `Nombre_Indicador`:
  - FWPCOP  : puntos forward USDCOP por plazo (dias)   -> puntos (COP)
  - FWTCOP  : tasa forward/descuento COP por plazo     -> tasa (%)
  - FWP<xxx>: puntos forward del par cruzado (BRL, MXN, JPY, CLP, CHF, CAD...)
  - FWPEUR / Fwd_EUR2COP / FWTEURCOP: curvas EUR
Y en el bloque derecho, la tabla de PARIDADES (spot por par: TRM, EUR, USDBRL...).

Salida: <dir>/fwd_curves/<INDICADOR>.csv (plazo_dias,valor) y paridades.csv (par,spot).
Uso: python3 -m bmk.pricing.extraer_curvas_fwd "insumos/Benchmark (Detallado).xlsb"
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# Curvas por plazo que alimentan la valoracion de forwards.
CURVAS = ("FWPCOP", "FWTCOP", "LIBBTS", "FWPEUR", "FWTEURCOP", "Fwd_EUR2COP",
          "FWPBRL", "FWPMXN", "FWPJPY", "FWPCLP", "FWPCHF", "FWPCAD",
          "Fwd_GBPUSD_Diaria")


def extraer(xlsb_path: str | Path, out_dir: str | Path = "insumos") -> dict:
    import pyxlsb
    out_dir = Path(out_dir) / "fwd_curves"
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = pyxlsb.open_workbook(str(xlsb_path))
    curvas: dict[str, list[tuple[int, float]]] = {c: [] for c in CURVAS}
    paridades: list[tuple[str, float]] = []
    fecha_val = None
    with wb.get_sheet("Curvas") as sh:
        for row in sh.rows():
            v = [c.v for c in row]
            # Bloque izquierdo A:E -> curvas por plazo (col2=nombre, col3=plazo, col4=valor).
            if len(v) > 4 and isinstance(v[2], str) and v[2] in curvas \
                    and isinstance(v[3], (int, float)) and isinstance(v[4], (int, float)):
                curvas[v[2]].append((int(v[3]), float(v[4])))
            # Bloque derecho: paridades (col34 rotulo 'FORWARD<xxx>', col38 spot del corte).
            if len(v) > 38 and isinstance(v[34], str) and str(v[34]).startswith("FORWARD") \
                    and isinstance(v[38], (int, float)):
                moneda = str(v[34]).replace("FORWARD", "").strip()
                paridades.append((moneda, float(v[38])))
    manifest = {}
    for c, pares in curvas.items():
        if pares:
            df = pd.DataFrame(sorted(set(pares)), columns=["plazo_dias", "valor"])
            df.to_csv(out_dir / f"{c}.csv", index=False)
            manifest[c] = len(df)
    if paridades:
        pd.DataFrame(sorted(set(paridades)), columns=["par", "spot"]).to_csv(
            out_dir / "paridades.csv", index=False)
        manifest["paridades"] = len(set(p[0] for p in paridades))
    return manifest


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "insumos/Benchmark (Detallado).xlsb"
    m = extraer(src)
    print("Curvas forward extraidas a insumos/fwd_curves/:")
    for k, v in m.items():
        print(f"  {k}: {v}")
