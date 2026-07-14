"""
Consolida los PRECIOS reportados por la SFC (Formato_351) de los ISIN unicos,
en una matriz: una fila por ISIN, una columna por mes.

El Formato_351 (formato publicado de la SFC) trae, por instrumento:
  - col 81 'No. Iden Asignado Custodio' = ISIN
  - col 63 'Precio'                      = precio reportado
  - col 27 'Nemotecnico', col 25/26 clase de inversion, col 34/35 moneda,
    col 100 pais emisor  (metadatos)
Como un mismo ISIN lo tienen varias AFP/portafolios (mismo precio de mercado), se
consolida a un precio por ISIN y mes (mediana, robusta a diferencias de vendor).

Uso:
    python3 tools/consolidar_precios.py --sfc-dir insumos/sfc \
        --out salidas/precios_consolidados_2026.xlsx
    # o meses explicitos:
    python3 tools/consolidar_precios.py \
        --sfc insumos/sfc/2026_01portainvdeta.xls insumos/sfc/2026_02portainvdeta.xls ... \
        --out salidas/precios_consolidados_2026.xlsx
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd

from bmk.ingest import sfc as sfc_mod

# Columnas del Formato_351 (0-based).
C_ISIN, C_PRECIO, C_NEMO, C_CLASE, C_DCLASE = 81, 63, 27, 25, 26
C_MONEDA, C_DMONEDA, C_PAIS, C_VRMDO = 34, 35, 100, 53
_MES = re.compile(r"(\d{4})[_-]?(\d{2})portainvdeta", re.I)


def _periodo(path: str) -> str:
    """AAAA-MM desde el nombre del archivo SFC (AAAA_MMportainvdeta)."""
    m = _MES.search(Path(path).name)
    return f"{m.group(1)}-{m.group(2)}" if m else Path(path).stem


def _leer_precios(path: str) -> pd.DataFrame:
    """ISIN, precio y metadatos de un archivo SFC (Formato_351)."""
    arch = sfc_mod.leer(path)
    f = arch.formato_351
    isin = f.iloc[:, C_ISIN].astype(str).str.strip()
    precio = pd.to_numeric(f.iloc[:, C_PRECIO], errors="coerce")
    d = pd.DataFrame({
        "isin": isin,
        "precio": precio,
        "nemo": f.iloc[:, C_NEMO].astype(str).str.strip(),
        "clase": f.iloc[:, C_CLASE].astype(str).str.strip(),
        "desc_clase": f.iloc[:, C_DCLASE].astype(str).str.strip(),
        "moneda": f.iloc[:, C_DMONEDA].astype(str).str.strip(),
        "pais": f.iloc[:, C_PAIS].astype(str).str.strip(),
    })
    d = d[(d["isin"].str.upper().isin(["", "NAN", "NONE"]) == False) & d["precio"].notna()]
    return d


def consolidar(paths: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Matriz ISIN x mes de precios + tabla de metadatos por ISIN."""
    frames = {}
    metas = []
    for p in sorted(paths):
        per = _periodo(p)
        d = _leer_precios(p)
        # Un precio por ISIN y mes (mediana; robusta a diferencias entre holders).
        precio = d.groupby("isin")["precio"].median()
        frames[per] = precio
        def _primero(s):
            for x in s:
                x = str(x).strip()
                if x and x.lower() != "nan":
                    return x
            return ""
        meta = d.sort_values("isin").groupby("isin").agg(
            nemo=("nemo", _primero), clase=("clase", _primero),
            desc_clase=("desc_clase", _primero), moneda=("moneda", _primero),
            pais=("pais", _primero))
        meta["periodo"] = per
        metas.append(meta.reset_index())
        print(f"  {per}: {len(precio):>4} ISIN con precio  ({Path(p).name})")

    matriz = pd.DataFrame(frames).sort_index()          # index = ISIN, cols = meses
    matriz.index.name = "ISIN"
    # metadatos: se toma el registro mas reciente por ISIN.
    meta_all = pd.concat(metas, ignore_index=True).sort_values("periodo")
    meta_all = meta_all.groupby("isin").last().drop(columns="periodo")
    meses = list(matriz.columns)
    matriz["# meses"] = matriz[meses].notna().sum(axis=1)
    out = meta_all.join(matriz, how="right").reset_index().rename(columns={"index": "ISIN"})
    # ordenar columnas: metadatos, luego meses, luego # meses
    cols = ["ISIN", "nemo", "clase", "desc_clase", "moneda", "pais"] + meses + ["# meses"]
    out = out[[c for c in cols if c in out.columns]]
    return out, matriz


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--sfc-dir", help="Carpeta con archivos AAAA_MMportainvdeta")
    ap.add_argument("--sfc", nargs="*", help="Archivos SFC explicitos")
    ap.add_argument("--out", default="salidas/precios_consolidados.xlsx")
    a = ap.parse_args()

    paths = list(a.sfc or [])
    if a.sfc_dir:
        paths += [str(p) for p in Path(a.sfc_dir).glob("*portainvdeta*")]
    paths = sorted(set(paths))
    if not paths:
        raise SystemExit("No hay archivos SFC. Use --sfc-dir o --sfc.")
    print(f"Consolidando precios de {len(paths)} archivos SFC:")
    out, matriz = consolidar(paths)

    meses = [c for c in matriz.columns if re.match(r"\d{4}-\d{2}", str(c))]
    print(f"\nISIN unicos con al menos un precio: {len(out)}")
    print(f"Meses: {meses}")
    print(f"ISIN presentes en TODOS los meses: {int((out['# meses'] == len(meses)).sum())}")

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(a.out, engine="xlsxwriter") as xw:
        out.to_excel(xw, sheet_name="Precios ISIN x mes", index=False)
    print(f"\n-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
