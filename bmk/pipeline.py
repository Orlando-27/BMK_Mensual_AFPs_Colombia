"""
Orquestador del pipeline Benchmark Mensual (end-to-end).

Encadena las etapas en orden, acumulando alertas (que NO detienen el proceso) y
generando las salidas:
  1. Ingesta SFC (descarga o lee local)
  2. Normalizacion + clasificacion + deteccion de instrumentos nuevos
  3. Derivados: enrutamiento, forwards, swaps (tipo) + valoracion opcional (v6)
  4. Consolidacion + controles de reconciliacion
  5. Salidas: consolidado.xlsx + alertas.xlsx/json + resumen

Uso:
    python3 -m bmk.pipeline --mes Junio --anio 2026
    python3 -m bmk.pipeline --archivo insumos/sfc/2026_06portainvdeta.xls
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from bmk.config.settings import Config
from bmk.ingest import sfc
from bmk.prepare import normalize, csa as csamod, fx
from bmk.classify import classifier, detector
from bmk.derivatives import router, forwards, swaps as swp, benchmark_rows as der_rows
from bmk.consolidate import benchmark, reconcile
from bmk.output import excel, hoja_benchmark, controles
from bmk.alerts.registry import RegistroAlertas


def correr(cfg: Config, archivo: str | None = None) -> dict:
    reg = RegistroAlertas(corte=cfg.corte)
    log: list[str] = []

    def paso(msg):
        log.append(msg)
        print(f"[pipeline] {msg}")

    # 1. Ingesta
    if archivo:
        ruta = Path(archivo)
        paso(f"Leyendo insumo local: {ruta.name}")
    else:
        paso(f"Descargando corte {cfg.mes} {cfg.anio} de la SFC ...")
        ruta = sfc.descargar(cfg.mes, cfg.anio, cfg.dir_insumos / "sfc")
    arch = sfc.leer(ruta)
    ing = sfc.resumen(arch)
    paso(f"Ingesta: {ing['activos_351']} activos, {ing['derivados_415']} derivados, "
         f"{ing['cuentas_csa']} CSA")
    if ing["derivados_sin_tipo"]:
        reg.agregar("ingesta", "DERIVADOS_SIN_TIPO", "INFO",
                    valor_mercado=ing["derivados_sin_tipo"],
                    descripcion=f"{ing['derivados_sin_tipo']} filas de derivados sin tipo (vacias)")

    # 2. Normalizacion + clasificacion + deteccion
    activos = normalize.extraer_activos(arch.formato_351, cfg.dir_referencia)
    clas = classifier.clasificar(activos, cfg.dir_referencia)
    n_ok = int(clas["is_clasificado"].sum())
    paso(f"Clasificacion: {n_ok}/{len(clas)} ({100*n_ok/len(clas):.1f}%) auto-clasificados")
    det = detector.detectar(clas, reg, cfg.dir_referencia)
    paso(f"Deteccion: {det['instrumentos_nuevos_ND']} instrumentos nuevos, "
         f"{det['clas_sfc_desconocidos']} codigos SFC desconocidos")

    # 3. Derivados
    rutas = router.enrutar(arch.formato_415)
    fwd = forwards.procesar(rutas["forwards"])
    sw = swp.procesar(rutas["swaps"])
    n_revisar = int((fwd["compra_venta"] == "REVISAR").sum()) if len(fwd) else 0
    if n_revisar:
        reg.agregar("derivados", "FORWARD_COMPRAVENTA_REVISAR", "WARN",
                    valor_mercado=n_revisar,
                    descripcion=f"{n_revisar} forwards sin poder determinar compra/venta")
    paso(f"Derivados: {len(fwd)} forwards, {len(rutas['futuros_local'])+len(rutas['futuros_int'])} "
         f"futuros, {len(sw)} swaps")

    # 3b. Valoracion de swaps (opcional, si hay insumos de curvas del corte)
    valoracion_swaps = None
    if cfg.dir_swaps_insumos and Path(cfg.dir_swaps_insumos).exists():
        try:
            from bmk.pricing import valuator
            valoracion_swaps = valuator.valorar_swaps(cfg.dir_swaps_insumos, cfg.dir_corte() / "swaps")
            paso(f"Valoracion swaps (motor v6): {len(valoracion_swaps)} swaps")
        except Exception as e:  # noqa: BLE001
            reg.agregar("valoracion", "FALLO_VALORACION_SWAPS", "ERROR",
                        descripcion=str(e)[:200],
                        accion_sugerida="Verificar insumos de curvas del corte")
            paso(f"Valoracion swaps: fallo ({str(e)[:80]})")
    else:
        paso("Valoracion swaps: omitida (sin insumos de curvas del corte)")

    # 4. Consolidacion + controles
    clas_pesos = benchmark.con_pesos(clas)
    ctrl = reconcile.controles(clas, reg)
    paso(f"Consolidacion: total industria {ctrl['valor_total_industria']/1e12:.1f} billones COP, "
         f"{ctrl['n_fondos']} fondos")

    # 5. Salidas
    hojas = {
        "Valor por fondo": ctrl["fondos"],
        "Por clase inversion": benchmark.por_clase(clas, "clase_inversion"),
        "Por clasificacion": benchmark.por_clase(clas, "clasificacion"),
        "Swaps por tipo": swp.resumen_por_tipo(sw),
        "Forwards por par": (fwd.groupby(["par", "compra_venta"]).size().reset_index(name="n")
                             if len(fwd) else None),
    }
    if valoracion_swaps is not None:
        hojas["Valoracion swaps"] = valoracion_swaps
    ruta_xlsx = excel.escribir_consolidado(cfg.dir_corte(), cfg.corte, hojas)
    paso(f"Consolidado: {ruta_xlsx}")

    # Cuentas CSA (caja internacional): TRM derivada del propio archivo SFC.
    trm = fx.trm(arch.formato_351)
    csa_rows = csamod.extraer_csa(arch.cuentas_csa, trm)
    paso(f"CSA: {len(csa_rows)} cuentas internacionales (TRM {trm:,.2f} del archivo)")

    # Filas de derivados (forwards/opciones/swaps) para la hoja Benchmark.
    der = der_rows.generar(arch.formato_415)
    paso(f"Derivados hoja: {len(der)} filas ({(der['clasificacion']=='FORWARD').sum()} fwd, "
         f"{(der['clasificacion']=='OPCIONES').sum()} opt, {(der['clasificacion']=='SWAP').sum()} swap)")

    # Hoja Benchmark para importar en la herramienta diaria (solo datos; formulas vacias)
    ruta_hoja = hoja_benchmark.escribir(clas, cfg.dir_corte(), cfg.corte, csa=csa_rows, derivados=der)
    paso(f"Hoja Benchmark (importar): {ruta_hoja}")

    # Archivo de controles formulado (verificacion trazable en Excel)
    ruta_ctrl = controles.generar(clas, cfg.dir_corte(), cfg.corte)
    paso(f"Controles (formulado): {ruta_ctrl}")

    alert_out = reg.escribir(cfg.dir_corte())
    paso(f"Alertas: {alert_out['n']} -> {alert_out['xlsx']}")

    resumen = {
        "corte": cfg.corte, "insumo": str(ruta),
        "ingesta": ing, "clasificacion_pct": round(100 * n_ok / len(clas), 1),
        "deteccion": det, "derivados": {"forwards": len(fwd), "swaps": len(sw)},
        "consolidacion": {k: v for k, v in ctrl.items() if k != "fondos"},
        "alertas": reg.resumen(),
        "salidas": {"consolidado": ruta_xlsx, "hoja_benchmark": ruta_hoja,
                    "controles": ruta_ctrl, **alert_out},
        "log": log,
    }
    (cfg.dir_corte() / "resumen.json").write_text(
        json.dumps(resumen, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return resumen


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mes", default="Junio")
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--archivo", help="Leer insumo SFC local en vez de descargar")
    ap.add_argument("--swaps-insumos", help="Carpeta con insumos de curvas para valorar swaps (motor v6)")
    args = ap.parse_args()
    cfg = Config(anio=args.anio, mes=args.mes)
    if args.swaps_insumos:
        cfg.dir_swaps_insumos = Path(args.swaps_insumos)
    resumen = correr(cfg, archivo=args.archivo)
    print("\n=== RESUMEN ===")
    print(json.dumps({k: v for k, v in resumen.items() if k != "log"},
                     ensure_ascii=False, indent=2, default=str)[:1500])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
