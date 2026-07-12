"""
Ingesta del Portafolio de Inversion Detallado de la SFC (paso 1-2 del manual).

Descarga el archivo mensual publicado por la Superintendencia Financiera de
Colombia (dia 10 de cada mes, corte del mes anterior) y lo lee en DataFrames.

El archivo trae las hojas:
    Formato_351   -> activos y notas estructuradas de la industria
    Formato_415   -> derivados (col J = "Tipo de derivado": 1/4/5/6/16/17)
    Cuentas CSA   -> cuentas internacionales (caja)

Descubrimiento del enlace: la pagina de publicaciones lista los meses del anio
como anclas con un parametro idFile; se mapea nombre de mes -> idFile.

Uso CLI:
    python3 -m bmk.ingest.sfc --mes Junio --anio 2026 --out insumos/sfc
    python3 -m bmk.ingest.sfc --archivo insumos/sfc/2026_06portainvdeta.xls --solo-leer
"""
from __future__ import annotations

import argparse
import html
import os
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

# Pagina de publicaciones del Portafolio de Inversion Detallado (pension obligatoria).
PUBLICACION_PO = (
    "https://www.superfinanciera.gov.co/publicaciones/10097246/"
    "informes-y-cifraspensiones-cesantias-y-fiduciariasinformacion-por-sector-"
    "pensiones-y-cesantiasregimen-de-ahorro-individual-con-solidaridad-fondos-de-"
    "pensiones-obligatoriasportafolio-de-inversionportafolio-de-inversion-detallado-10097246/"
)
DESCARGA_BASE = "https://www.superfinanciera.gov.co/loader.php"

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MES_NUM = {m.lower(): i + 1 for i, m in enumerate(MESES)}

# Esquema de columnas de Formato_415 (0-indexed) — util para el enrutamiento.
COL_415 = {
    "entidad": 0, "fecha_corte": 1, "tipo_patrimonio": 2, "codigo_patrimonio": 4,
    "nombre": 5, "tipo_derivado": 9, "tipo_opcion": 10, "numero_contrato": 11,
    "moneda_liquidacion": 15, "fecha_celebracion": 22, "fecha_vencimiento": 23,
    "fecha_liquidacion": 24, "posicion_compra_venta": 25, "moneda_derecho": 26,
    "nominal_derecho": 27, "moneda_obligacion": 28, "nominal_obligacion": 29,
    "strike": 30, "tipo_subyacente": 31, "id_subyacente": 32,
}

# Codigos de "Tipo de derivado" -> destino en el proceso (ver Diccionario SFC).
TIPO_DERIVADO = {
    1: "FORWARD", 4: "FUTURO_TRM", 5: "FUTURO_LOCAL", 6: "FUTURO_INTERNACIONAL",
    16: "SWAP", 17: "SWAP",
}


@dataclass
class ArchivoSFC:
    """Contenido leido del archivo SFC de un corte."""
    formato_351: pd.DataFrame
    formato_415: pd.DataFrame
    cuentas_csa: pd.DataFrame
    ruta: Path


def _get(url: str, **kw):
    import requests  # solo se necesita para descargar de la SFC, no para leer local
    r = requests.get(url, timeout=kw.pop("timeout", 180), allow_redirects=True, **kw)
    r.raise_for_status()
    return r


def descubrir_meses(publicacion_url: str = PUBLICACION_PO) -> dict[str, str]:
    """Devuelve {nombre_mes: url_descarga} para el anio publicado en la pagina."""
    h = _get(publicacion_url, timeout=90).text
    anchors = re.findall(r'<a\s+[^>]*href="([^"]*idFile=\d+[^"]*)"[^>]*>(.*?)</a>',
                         h, re.I | re.S)
    out: dict[str, str] = {}
    for href, text in anchors:
        texto = html.unescape(re.sub(r"<[^>]+>", "", text)).strip()
        if texto in MESES:
            href = html.unescape(href)
            url = href if href.startswith("http") else \
                "https://www.superfinanciera.gov.co" + (href if href.startswith("/") else "/" + href)
            out[texto] = url
    return out


def nombre_archivo(anio: int, mes: int) -> str:
    """Convencion interna: AAAA_MMportainvdeta (mes con cero a la izquierda)."""
    return f"{anio}_{mes:02d}portainvdeta"


def descargar(mes: str, anio: int, out_dir: str | Path) -> Path:
    """Descarga el archivo SFC del mes indicado. Devuelve la ruta local."""
    mes = mes.capitalize()
    if mes not in MESES:
        raise ValueError(f"Mes invalido: {mes}. Use uno de {MESES}")
    urls = descubrir_meses()
    if mes not in urls:
        raise RuntimeError(f"El mes {mes} no aparece publicado. Disponibles: {sorted(urls)}")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    resp = _get(urls[mes])
    # Detectar extension por firma (OLE2 = .xls; ZIP 'PK' = .xlsx)
    firma = resp.content[:4]
    ext = ".xls" if firma[:2] == b"\xd0\xcf" else (".xlsx" if firma[:2] == b"PK" else ".bin")
    dest = out_dir / f"{nombre_archivo(anio, MES_NUM[mes.lower()])}{ext}"
    dest.write_bytes(resp.content)
    return dest


def _read_sheet(ruta: Path, hoja: str, header_row: int = 0) -> pd.DataFrame:
    ext = ruta.suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    return pd.read_excel(ruta, sheet_name=hoja, header=header_row, engine=engine)


def leer(ruta: str | Path) -> ArchivoSFC:
    """Lee las hojas relevantes del archivo SFC en DataFrames."""
    ruta = Path(ruta)
    ext = ruta.suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xls = pd.ExcelFile(ruta, engine=engine)
    # Los nombres de hoja pueden venir como 'Formato_351' o 'Formato_415'.
    def hoja(candidatos):
        for c in candidatos:
            if c in xls.sheet_names:
                return c
        raise KeyError(f"No se encontro ninguna de {candidatos} en {xls.sheet_names}")
    f351 = pd.read_excel(xls, sheet_name=hoja(["Formato_351", "Fmto-351"]))
    f415 = pd.read_excel(xls, sheet_name=hoja(["Formato_415", "Fmto-415"]))
    try:
        csa = pd.read_excel(xls, sheet_name=hoja(["Cuentas CSA", "CSA"]))
    except KeyError:
        csa = pd.DataFrame()
    return ArchivoSFC(formato_351=f351, formato_415=f415, cuentas_csa=csa, ruta=ruta)


def resumen(arch: ArchivoSFC) -> dict:
    """Resumen rapido de controles de ingesta (para log/alertas)."""
    col_tipo = arch.formato_415.columns[COL_415["tipo_derivado"]]
    tipos = arch.formato_415[col_tipo].value_counts(dropna=False).to_dict()
    tipos_norm = {}
    sin_tipo = 0
    for k, v in tipos.items():
        try:
            code = int(float(k))
            tipos_norm[TIPO_DERIVADO.get(code, f"OTRO_{code}")] = tipos_norm.get(
                TIPO_DERIVADO.get(code, f"OTRO_{code}"), 0) + v
        except (ValueError, TypeError):
            sin_tipo += v
    return {
        "activos_351": len(arch.formato_351),
        "derivados_415": len(arch.formato_415),
        "cuentas_csa": len(arch.cuentas_csa),
        "derivados_por_tipo": tipos_norm,
        "derivados_sin_tipo": sin_tipo,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mes", help="Nombre del mes a descargar (ej. Junio)")
    ap.add_argument("--anio", type=int, help="Anio (ej. 2026)")
    ap.add_argument("--out", default="insumos/sfc", help="Carpeta destino")
    ap.add_argument("--archivo", help="Leer un archivo ya descargado en vez de bajar")
    ap.add_argument("--solo-leer", action="store_true", help="Solo leer y mostrar resumen")
    args = ap.parse_args()

    if args.archivo:
        ruta = Path(args.archivo)
    elif args.mes and args.anio:
        print(f"Descargando {args.mes} {args.anio} desde la SFC ...")
        ruta = descargar(args.mes, args.anio, args.out)
        print(f"  guardado: {ruta} ({ruta.stat().st_size/1e6:.1f} MB)")
    else:
        ap.error("Indique --mes y --anio, o --archivo")

    arch = leer(ruta)
    r = resumen(arch)
    print("\nResumen de ingesta:")
    for k, v in r.items():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
