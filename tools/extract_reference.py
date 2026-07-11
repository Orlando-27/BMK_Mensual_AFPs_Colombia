#!/usr/bin/env python3
"""
Extrae las tablas de diccionario/mapeo de los archivos Excel del proceso y las
materializa como datos de referencia versionados en config/reference/*.csv.

Fase F0 del roadmap (contrato de datos). Fuente: los .xlsb/.xlsm en insumos/.
Lee valores (data_only) y corta al final real de datos (no al máximo de Excel).

Uso: python3 tools/extract_reference.py
"""
from __future__ import annotations
import csv
from pathlib import Path

import openpyxl
from pyxlsb import open_workbook

INS = Path("insumos")
OUT = Path("config/reference")
OUT.mkdir(parents=True, exist_ok=True)

BMO = INS / "Benchmark Mensual Optimizado.xlsb"
INICIO = INS / "INICIO MACRO MULTIFONDOS.xlsm"


def write_csv(name: str, header: list[str], rows: list[list]) -> None:
    dest = OUT / name
    with dest.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"  {name}: {len(rows)} filas")


def s(v) -> str:
    return "" if v is None else str(v).strip()


def extract_xlsx_table(path, sheet, cols, header, *, start=1, stop_empty=25):
    """Lee openpyxl (xlsm/xlsx) cols (0-idx) desde fila start; corta tras
    stop_empty filas consecutivas con la 1ª columna vacía."""
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb[sheet]
    rows, empties = [], 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < start:
            continue
        key = s(row[cols[0]]) if cols[0] < len(row) else ""
        if not key:
            empties += 1
            if empties >= stop_empty:
                break
            continue
        empties = 0
        rows.append([s(row[c]) if c < len(row) else "" for c in cols])
    wb.close()
    write_csv_dedup(sheet, header, rows)
    return rows


def write_csv_dedup(name_sheet, header, rows):
    pass  # placeholder (no usar); se usa write_csv directamente abajo


def extract_diccionario_sfc():
    """Bloques de la hoja Diccionario SFC del BMO."""
    data = []
    with open_workbook(str(BMO)) as wb:
        with wb.get_sheet("Diccionario SFC") as sh:
            for i, row in enumerate(sh.rows()):
                if i > 80:
                    break
                d = {c.c: c.v for c in row}
                data.append(d)

    def col(idx):
        return [s(d.get(idx)) for d in data]

    # Clas SFC (K=10) -> Clase de Inversión (L=11)  [leyenda; puede ser 1:N]
    seen, clas = set(), []
    for d in data:
        k, v = s(d.get(10)), s(d.get(11))
        if k and v and k != "SFC" and (k, v) not in seen:
            seen.add((k, v)); clas.append([k, v])
    write_csv("clas_sfc_legend.csv", ["clas_sfc", "clase_inversion"], clas)

    # Códigos de derivados (W=22 -> X=23)
    seen, der = set(), []
    for d in data:
        code, tipo = s(d.get(22)), s(d.get(23))
        if code and tipo and code not in ("Código",) and code not in seen:
            seen.add(code); der.append([code, tipo])
    write_csv("derivados_sfc.csv", ["codigo", "tipo_derivado"], der)

    # Normalización AFP (A=0 buscar -> C=2 reemplazo); filas con ambos
    seen, afp = set(), []
    for d in data:
        orig, std = s(d.get(0)), s(d.get(2))
        if orig and std and orig not in ("DEBE ESTAR", "ESTA") and std not in ("DEBE ESTAR", "ESTA") \
           and orig != std and orig not in seen:
            seen.add(orig); afp.append([orig, std])
    write_csv("afp_normalizacion.csv", ["nombre_sfc", "nombre_estandar"], afp)


def extract_clasificacion():
    """Tabla maestra ISIN/NEMO -> atributos (INICIO!CLASIFICACION cols A..H)."""
    cols = list(range(8))  # A..H
    header = ["nemo_o_isin", "clase_inversion", "moneda", "tasa_facial",
              "ubicacion", "clasificacion", "riesgo", "clase"]
    wb = openpyxl.load_workbook(str(INICIO), read_only=True, data_only=True)
    ws = wb["CLASIFICACION"]
    rows, empties = [], 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        key = s(row[0]) if len(row) else ""
        if not key:
            empties += 1
            if empties >= 50:
                break
            continue
        empties = 0
        rows.append([s(row[c]) if c < len(row) else "" for c in cols])
    wb.close()
    write_csv("clasificacion.csv", header, rows)


def extract_cprvi():
    """Mapeo moneda/región de fondos RVI y notas (INICIO!CPRVI)."""
    wb = openpyxl.load_workbook(str(INICIO), read_only=True, data_only=True)
    ws = wb["CPRVI"]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = [s(v) for v in all_rows[1]]  # fila 1 = encabezados reales
    rows, empties = [], 0
    for row in all_rows[2:]:
        key = s(row[0]) if len(row) else ""
        if not key:
            empties += 1
            if empties >= 20:
                break
            continue
        empties = 0
        rows.append([s(v) for v in row[:len(header)]])
    # limpiar columnas de fórmula CONTROL/TOTAL si quedaron como texto vacío
    write_csv("cprvi.csv", header, rows)


def extract_titulos_nuevos():
    wb = openpyxl.load_workbook(str(INICIO), read_only=True, data_only=True)
    ws = wb["TITULOS NUEVOS"]
    header = ["isin", "nemo", "frecuencia", "moneda"]
    rows, empties = [], 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        key = s(row[0]) if len(row) else ""
        if not key:
            empties += 1
            if empties >= 20:
                break
            continue
        empties = 0
        rows.append([s(row[c]) if c < len(row) else "" for c in range(4)])
    wb.close()
    write_csv("titulos_nuevos.csv", header, rows)


def main():
    print("Extrayendo datos de referencia a config/reference/ ...")
    extract_diccionario_sfc()
    extract_titulos_nuevos()
    extract_cprvi()
    extract_clasificacion()  # el más grande, al final
    print("Listo.")


if __name__ == "__main__":
    main()
