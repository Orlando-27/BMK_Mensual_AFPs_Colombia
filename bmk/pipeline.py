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

import pandas as pd

from bmk.config.settings import Config
from bmk.ingest import sfc
from bmk.prepare import normalize, csa as csamod, fx
from bmk.classify import classifier, detector
from bmk.derivatives import router, forwards, swaps as swp, benchmark_rows as der_rows
from bmk.consolidate import benchmark, reconcile
from bmk.output import excel, hoja_benchmark, controles, valoracion_fwds, valoracion_futuros
from bmk.alerts.registry import RegistroAlertas


def _insumo(nombre: str) -> Path:
    """Resuelve un insumo: usa insumos/ si existe, si no cae a insumos_demo/
    (placeholder versionado en el repo, para correr sin abrir los .xlsb)."""
    p = Path("insumos") / nombre
    return p if p.exists() else Path("insumos_demo") / nombre


def correr(cfg: Config, archivo: str | None = None, vpn_swaps=None,
           fecha_fwd: str | None = None) -> dict:
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

    # Valoracion de forwards de la industria (replica hoja Fwd Industria).
    # Se calcula ANTES de la hoja Benchmark para poder inyectar el valor revaluado
    # (curva de puntos) en las filas de forward en vez del proxy SFC col 53/54.
    ruta_fwd = None
    fwd_val = trm_val = local_fut = intl_fut = swap_sab = None
    dir_curvas = _insumo("fwd_curves")
    if (dir_curvas / "paridades.csv").exists():
        try:
            from bmk.pricing.fwd_valuator import Curvas, valorar
            curvas = Curvas(dir_curvas)
            # El macro revalua los forwards a la fecha de PUBLICACION (~10 dias
            # despues del corte), no al corte. Si se pasa fecha_fwd (con curvas de
            # esa fecha en fwd_curves), se valora a esa fecha.
            fecha_v = valoracion_fwds.corte_a_serial(fecha_fwd or cfg.corte)
            fwd415 = valoracion_fwds.normalizar_415(arch.formato_415)
            fwd_val = valorar(fwd415, curvas, fecha_v)  # df crudo (para controles e inyeccion)
            ruta_fwd = valoracion_fwds.escribir(fwd415, cfg.dir_corte(), cfg.corte, dir_curvas=dir_curvas)
            # Futuros de TRM (tipo 4) valorados con el mismo motor.
            trm415 = valoracion_fwds.normalizar_415(arch.formato_415, tipo_derivado=4)
            if len(trm415):
                trm_val = valorar(trm415, curvas, fecha_v)
            paso(f"Valoracion Fwds Industria: {len(fwd415)} forwards, {len(trm415)} fut TRM -> {ruta_fwd}")
        except Exception as e:  # noqa: BLE001
            reg.agregar("valoracion", "FALLO_VALORACION_FWDS", "ERROR", descripcion=str(e)[:200])
            paso(f"Valoracion Fwds: fallo ({str(e)[:80]})")
    else:
        paso("Valoracion Fwds: omitida (sin curvas en insumos/fwd_curves)")

    # Filas de derivados (forwards/opciones/swaps) para la hoja Benchmark.
    der = der_rows.generar(arch.formato_415)
    # Inyecta el valor revaluado de forwards (motor fwd_valuator, metodologia
    # bi-moneda del macro) en vez del proxy SFC col 53/54. El MtM (pyg) se atribuye
    # a la divisa del par. Requiere haber valorado a la fecha de publicacion.
    if fwd_val is not None and len(fwd_val):
        der = der_rows.inyectar_valor_forwards(der, fwd_val)
        paso("Forwards revaluados (bi-moneda) inyectados en la hoja")
    if vpn_swaps is not None and len(vpn_swaps):
        der = der_rows.inyectar_vpn_swaps(der, vpn_swaps)
        n_val = int(((der["clasificacion"] == "SWAP") & der["vr_mercado"].notna()).sum())
        paso(f"Swaps valorados (VPN v6) inyectados: {n_val} patas")
    paso(f"Derivados hoja: {len(der)} filas ({(der['clasificacion']=='FORWARD').sum()} fwd, "
         f"{(der['clasificacion']=='OPCIONES').sum()} opt, {(der['clasificacion']=='SWAP').sum()} swap)")

    # Hoja Benchmark para importar en la herramienta diaria (solo datos; formulas vacias)
    ruta_hoja = hoja_benchmark.escribir(clas, cfg.dir_corte(), cfg.corte, csa=csa_rows, derivados=der)
    paso(f"Hoja Benchmark (importar): {ruta_hoja}")

    # Valoracion de futuros (locales TES + internacionales), replica FutLoc/FutInt.
    ruta_fut = None
    vector_path = _insumo("vector_precios.csv")
    if vector_path.exists():
        try:
            trm = fx.trm(arch.formato_351)
            fx_map = {"USD": trm, "COP": 1.0}
            local_fut, intl_fut = valoracion_futuros.valorar(arch.formato_415, vector_path=vector_path, fx=fx_map)
            ruta_fut = valoracion_futuros.escribir(arch.formato_415, cfg.dir_corte(), cfg.corte,
                                                   vector_path=vector_path, fx=fx_map)
            paso(f"Valoracion Futuros Industria: {ruta_fut}")
            # Alertar futuros internacionales sin precio en el vector (instrumento nuevo).
            nuevos = intl_fut[intl_fut["precio_vector"].isna()] if len(intl_fut) else intl_fut
            for _, r in nuevos.iterrows():
                reg.agregar("valoracion", "FUTURO_INT_NO_VALORADO", "WARN",
                            afp=r.get("afp", ""), portafolio=r.get("portafolio", ""),
                            nemo=str(r.get("subyacente", "")), moneda=str(r.get("moneda", "")),
                            descripcion=f"Futuro internacional con subyacente '{r.get('subyacente','')}' "
                                        f"(tipo {r.get('tipo','')}) no esta en el diccionario de indices "
                                        f"ni en el VECTOR DE PRECIOS; no se pudo valorar.",
                            accion_sugerida="Mapear el subyacente a su indice y agregar su precio al vector.")
        except Exception as e:  # noqa: BLE001
            reg.agregar("valoracion", "FALLO_VALORACION_FUTUROS", "ERROR", descripcion=str(e)[:200])
            paso(f"Valoracion Futuros: fallo ({str(e)[:80]})")
    else:
        paso("Valoracion Futuros: omitida (sin insumos/vector_precios.csv)")

    # Sabana de condiciones faciales de swaps (insumo del motor de valoracion v6).
    try:
        from bmk.pricing.swaps import sabana as swap_sabana
        swap_sab = swap_sabana.construir(arch.formato_415)
        ruta_sab = swap_sabana.escribir(arch.formato_415, cfg.dir_corte(), cfg.corte)
        paso(f"Sabana condiciones faciales swaps: {len(swap_sab)} swaps -> {ruta_sab}")
    except Exception as e:  # noqa: BLE001
        reg.agregar("swaps", "FALLO_SABANA_SWAPS", "ERROR", descripcion=str(e)[:200])
        paso(f"Sabana swaps: fallo ({str(e)[:80]})")

    # Tabla unificada de derivados para los controles.
    from bmk.consolidate import derivados as der_control
    tabla_deriv = der_control.construir(fwd_val=fwd_val, trm_val=trm_val,
                                        fut_local=local_fut, fut_int=intl_fut,
                                        swaps_sabana=swap_sab, vpn_swaps=vpn_swaps)

    # Archivo de controles formulado (verificacion trazable en Excel), con derivados.
    ruta_ctrl = controles.generar(clas, cfg.dir_corte(), cfg.corte, derivados=tabla_deriv)
    paso(f"Controles (formulado): {ruta_ctrl} ({len(tabla_deriv)} derivados)")

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
