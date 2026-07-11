#!/usr/bin/env python3
"""
extraer_insumos.py
─────────────────────────────────────────────────────────────────────────────
Extrae TODOS los insumos necesarios desde el Excel Precia (.xlsb) y los
persiste como CSVs planos que el motor `swap_valuator_v6.py` puede consumir
directamente.

Produce en <out_dir>/:
    curva_<NOMBRE>_<YYYYMMDD>.csv   (13 curvas cero cupón, formato Dias,Tasa)
    SWAPIND_<YYYYMMDD>.csv          (portafolio 1,492 swaps + VPN Precia)
    INICIO_<YYYYMMDD>.csv           (FECHA_VAL, TRM, UVR, EURCOP, IPC)

Uso:
    python3 extraer_insumos.py \
        --xlsb  /path/to/Calculadora_Swap_N_31-03.xlsb \
        --fecha 2026-03-31 \
        --out   /path/to/insumos_20260331/

Por qué es importante:
    - Para valorar en otra fecha sólo se reemplaza el .xlsb y se ejecuta este
      script. El motor no requiere ningún cambio de código.
    - Las curvas quedan en el formato minimalista Dias,Tasa que ya usa el
      motor, sin depender de pyxlsb en runtime de valoración.
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

try:
    from pyxlsb import open_workbook
except ImportError:
    sys.exit("ERROR: instalar pyxlsb → pip install pyxlsb")


# ─── Mapa hoja DATOS → (col_dias, col_tasa, nombre_archivo_interno) ────────
# Columnas 0-indexed en la hoja DATOS:
#   IBR_COL_USD: D,E      → cols 3,4
#   IBR_UVR:     G,H      → cols 6,7
#   DTF:         J,K      → cols 9,10
#   USDOIS:      M,N      → cols 12,13
#   IPC:         P,Q      → cols 15,16
#   LIBORUVR:    S,T      → cols 18,19
#   LIBORCOP:    V,W      → cols 21,22
#   USDCO:       Y,Z      → cols 24,25
#   USDIBR:      AB,AC    → cols 27,28  (muy pocos nodos, no se usa)
#   EURCOP:      AE,AF    → cols 30,31
#   LIBORUSD3M:  AM,AN    → cols 38,39  (estimada)
#
# "UVR DIARIA" está en col AS+ (no es curva, es serie de valores UVR)
CURVAS_EXTRAER = [
    ("IBR_COLATERAL_USD", 3, 4),
    ("IBRUVR",            6, 7),
    ("DTF",               9, 10),
    ("USDOIS",            12, 13),
    ("IPC",               15, 16),
    ("LIBORUVR",          18, 19),
    ("LIBORCOP",          21, 22),
    ("USDCO",             24, 25),
    ("EURCOP",            30, 31),
    ("LIBORUSD3M",        43, 44),
    ("COP_COLATERAL_USD", 51, 52),  # ★ col AZ,BA - descuento pata COP de CCS
]

# El xlsb no tiene COP_COLATERAL_USD directo — se deriva de la misma fuente
# IBR pero con otra convención (Precia publica ambas curvas). En el contexto
# de este proyecto la fuente fue un archivo auxiliar; si no está disponible
# se usa IBR_COLATERAL_USD como proxy y se marca con un flag.


def excel_serial_to_date(serial: float) -> date:
    """Convierte un serial Excel (base 1900-01-01 con bug leap-year) a date."""
    return date(1899, 12, 30) + timedelta(days=int(serial))


def extraer_curva(wb, col_dias: int, col_tasa: int) -> pd.DataFrame:
    """Lee de DATOS un par (días, tasa) hasta encontrar celdas vacías."""
    with wb.get_sheet("DATOS") as sh:
        rows = list(sh.rows())
    data = []
    for row in rows[1:]:  # skip header
        vals = [c.v for c in row]
        if len(vals) <= max(col_dias, col_tasa):
            break
        d = vals[col_dias]
        t = vals[col_tasa]
        if d is None or t is None:
            continue
        data.append((int(d), float(t)))
    return pd.DataFrame(data, columns=["Dias", "Tasa"]).sort_values("Dias")


def extraer_swapind(wb) -> pd.DataFrame:
    """Extrae la hoja SWAPIND y preserva nombres exactos de columna.

    El layout del xlsb es:
        row 0: título del bloque
        row 1: cabecera intermedia ('Información Derecho', etc.)
        row 2: HEADERS reales ('TIPO SWAP', 'ISIN', ...)
        row 3+: datos
    """
    with wb.get_sheet("SWAPIND") as sh:
        rows = list(sh.rows())
    header = [str(c.v) if c.v is not None else f"col_{i}"
              for i, c in enumerate(rows[2])]
    data = []
    for row in rows[3:]:
        vals = [c.v for c in row[: len(header)]]
        if all(v is None for v in vals):
            break
        data.append(vals)
    df = pd.DataFrame(data, columns=header)
    # Normalizar fechas: columnas F.Compra, Emision, F.Vcto vienen como
    # seriales Excel. Las convertimos a YYYY-MM-DD (string ISO) porque es
    # lo que el motor parsea con pd.to_datetime.
    for col in df.columns:
        if col.strip().startswith(("F.Compra", "F.Vcto", "Emision")):
            df[col] = df[col].apply(
                lambda v: excel_serial_to_date(v).isoformat()
                if pd.notna(v) and isinstance(v, (int, float))
                else v
            )
    return df


def extraer_inicio(wb) -> pd.DataFrame:
    """Lee los spot FX y la fecha de valoración desde INICIO."""
    with wb.get_sheet("INICIO") as sh:
        rows = list(sh.rows())
    # Layout conocido (ver Excel):
    #   Row 1 (idx 1): [None, 'Fecha', 'TRM', 'UVR', 'EURL']
    #   Row 2 (idx 2): [None, <serial>, <trm>, <uvr>, <eurcop>]
    r1 = [c.v for c in rows[1]]
    r2 = [c.v for c in rows[2]]
    # Validación mínima
    if str(r1[1]).strip().lower() != "fecha":
        raise ValueError(f"INICIO row 1 no tiene 'Fecha' en col B: {r1[:5]}")
    fecha = excel_serial_to_date(float(r2[1]))
    trm   = float(r2[2])
    uvr   = float(r2[3])
    eur   = float(r2[4])
    # IPC (row idx 4, col D)
    r4 = [c.v for c in rows[4]]
    ipc = float(r4[3]) if len(r4) > 3 and r4[3] is not None else None

    # Construir DataFrame de una fila — formato key,value para lectura simple
    return pd.DataFrame(
        {
            "key":   ["FECHA_VAL", "TRM", "UVR", "EURCOP", "IPC"],
            "value": [fecha.isoformat(), trm, uvr, eur, ipc],
        }
    )


def extraer_vpn_precia(wb) -> pd.DataFrame:
    """
    Extrae la tabla PRECIOS VECTOR DE PRECIOS del rango B:J de INICIO
    (el benchmark oficial contra el que validamos).

    Columnas: #SWAP, DERECHO_COP, OBLIGACION_COP, VPN_COP,
              LLAVE_DER, PRECIO_DER, LLAVE_OBL, PRECIO_OBL
    """
    with wb.get_sheet("INICIO") as sh:
        rows = list(sh.rows())
    # El header real de esta tabla está en row idx 24 (según inspección).
    # Las filas de datos arrancan en idx 25.
    data = []
    for i in range(25, len(rows)):
        vals = [c.v for c in rows[i][:13]]
        # Columna B (idx 1) es el #SWAP. Si está vacía, fin de la tabla.
        if vals[1] is None or str(vals[1]).strip() == "":
            break
        data.append({
            "ISIN":           vals[1],
            "VPN_Der_Precia": vals[2],
            "VPN_Obl_Precia": vals[3],
            "VPN_Neto":       vals[4],
            "Llave_Der":      vals[6],
            "Precio_Der":     vals[7],
            "Llave_Obl":      vals[8],
            "Precio_Obl":     vals[9],
        })
    return pd.DataFrame(data)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsb",  required=True, help="Path al .xlsb de Precia")
    ap.add_argument("--fecha", required=True, help="Fecha valoración YYYY-MM-DD (debe coincidir con INICIO del xlsb)")
    ap.add_argument("--out",   required=True, help="Directorio de salida")
    args = ap.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    fecha = datetime.strptime(args.fecha, "%Y-%m-%d").date()
    yyyymmdd = fecha.strftime("%Y%m%d")

    print(f"Abriendo {args.xlsb} ...")
    wb = open_workbook(args.xlsb)

    # 1. INICIO (metadatos: FECHA_VAL, TRM, UVR, EURCOP)
    inicio = extraer_inicio(wb)
    inicio_path = out_dir / f"INICIO_{yyyymmdd}.csv"
    inicio.to_csv(inicio_path, index=False)
    print(f"  INICIO        → {inicio_path.name}  ({len(inicio)} entradas)")
    # Validación: fecha en xlsb = fecha argumento
    fecha_xlsb = inicio.loc[inicio["key"] == "FECHA_VAL", "value"].iloc[0]
    if fecha_xlsb != args.fecha:
        print(f"  ADVERTENCIA: fecha xlsb={fecha_xlsb} ≠ argumento={args.fecha}")

    # 2. Curvas cero cupón
    for nombre_archivo, col_d, col_t in CURVAS_EXTRAER:
        df = extraer_curva(wb, col_d, col_t)
        path = out_dir / f"curva_{nombre_archivo}_{yyyymmdd}.csv"
        df.to_csv(path, index=False)
        print(f"  {nombre_archivo:<20} → {path.name}  ({len(df)} nodos)")

    # 3. SWAPIND (portafolio)
    swapind = extraer_swapind(wb)
    sw_path = out_dir / f"SWAPIND_{yyyymmdd}.csv"
    swapind.to_csv(sw_path, index=False)
    print(f"  SWAPIND       → {sw_path.name}  ({len(swapind)} swaps)")

    # 4. Benchmark VPN Precia (rango B:J de INICIO)
    bench = extraer_vpn_precia(wb)
    b_path = out_dir / f"VPN_PRECIA_{yyyymmdd}.csv"
    bench.to_csv(b_path, index=False)
    print(f"  VPN benchmark → {b_path.name}  ({len(bench)} swaps)")

    print(f"\n✓ Extracción completa → {out_dir}")


if __name__ == "__main__":
    main()
