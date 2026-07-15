#!/usr/bin/env bash
# ============================================================================
#  Benchmark Mensual AFP — corre el proceso COMPLETO de un corte y empaqueta
#  todas las salidas en un zip que se descarga al final (Cloud Shell).
#
#  Uso (por defecto corre el corte de JUNIO 2026):
#      bash run_corte.sh
#
#  Para otro corte, sobre-escribe las variables al invocar. Ejemplo julio:
#      CORTE=2026-07-31 MES=Julio ANIO=2026 \
#      SFC=insumos/sfc/2026_07portainvdeta.xls \
#      SFC_PREV=insumos/sfc/2026_06portainvdeta.xls \
#      CURVAS=insumos_diarios/2026-07-31 \
#      CURVAS_PUB=insumos_diarios/2026-08-XX  FECHA_PUB=2026-08-XX \
#      CURVAS_PREV=insumos_diarios/2026-07-30 FECHA_PREV=2026-07-30 \
#      bash run_corte.sh
# ============================================================================
set -euo pipefail

# ── Parametros del corte (defaults = junio 2026) ────────────────────────────
CORTE="${CORTE:-2026-06-30}"          # fecha de corte SFC (M-1)
MES="${MES:-Junio}"
ANIO="${ANIO:-2026}"
SFC="${SFC:-insumos/sfc/2026_06portainvdeta.xls}"          # SFC del corte
SFC_PREV="${SFC_PREV:-insumos/sfc/2026_05portainvdeta.xls}" # SFC mes anterior (alertas)
CURVAS="${CURVAS:-insumos_diarios/2026-06-30}"             # curvas del corte (swaps)
CURVAS_PUB="${CURVAS_PUB:-insumos_diarios/2026-07-10}"     # curvas de publicacion (forwards)
FECHA_PUB="${FECHA_PUB:-2026-07-10}"
CURVAS_PREV="${CURVAS_PREV:-insumos_diarios/2026-06-29}"   # dia habil anterior (camara SOFR)
FECHA_PREV="${FECHA_PREV:-2026-06-29}"
MACRO="${MACRO:-insumos/Benchmark_corte_junio.xlsx}"        # Benchmark del macro (comparacion)

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO"
OUT="salidas/${CORTE}"

echo "############################################################"
echo "#  Corte ${CORTE} (${MES} ${ANIO})"
echo "############################################################"

# ── 1) Entorno Python ───────────────────────────────────────────────────────
if [ ! -d .venv ]; then python3 -m venv .venv; fi
# shellcheck disable=SC1091
source .venv/bin/activate
hash -r
pip install -q --upgrade pip
pip install -q -r requirements.txt

# ── 2) Limpia salidas viejas de este corte (zip limpio) ─────────────────────
rm -rf "$OUT" "salidas/swaps/${CORTE}" "salidas/swaps/${CORTE}_prev"
rm -f  salidas_"$(echo "$MES" | tr '[:upper:]' '[:lower:]')".zip

# ── 3) Proceso completo del corte ───────────────────────────────────────────
#     Forwards revaluados a publicacion; swaps al corte + variacion de camara.
python3 tools/correr_corte.py \
  --curvas      "$CURVAS"      --corte "$CORTE" --mes "$MES" --anio "$ANIO" \
  --sfc         "$SFC" \
  --curvas-pub  "$CURVAS_PUB"  --fecha-pub  "$FECHA_PUB" \
  --curvas-prev "$CURVAS_PREV" --fecha-prev "$FECHA_PREV"

# ── 4) Alertas mensuales (nuevos con region/moneda + cambios significativos) ─
if [ -f "$SFC_PREV" ]; then
  python3 tools/alertas_mensuales.py \
    --actual "$SFC" --anterior "$SFC_PREV" \
    --out "$OUT/alertas_mensuales_${CORTE}.xlsx"
else
  echo "(aviso) sin SFC del mes anterior ($SFC_PREV): se omiten las alertas mensuales"
fi

# ── 5) Validacion del motor de swaps vs precios reportados por la SFC ────────
python3 tools/comparar_swaps_sfc.py \
  --sfc "$SFC" \
  --valoracion "$OUT/valoracion_swaps_industria_${CORTE}.xlsx" \
  --out "$OUT/comparacion_swaps_vs_sfc_${CORTE}.xlsx"

# ── 6) Comparacion completa vs el macro (si esta disponible) ────────────────
if [ -f "$MACRO" ]; then
  python3 tools/comparar_todo.py \
    --macro "$MACRO" --nuestro "$OUT/hoja_Benchmark_${CORTE}.xlsx" \
    --out "$OUT/comparacion_vs_macro_${CORTE}.xlsx"
else
  echo "(aviso) sin macro ($MACRO): se omite la comparacion vs macro"
fi

# ── 7) Consolidado de precios por ISIN x mes (con los SFC que haya) ──────────
python3 tools/consolidar_precios.py --sfc-dir insumos/sfc \
  --out "$OUT/precios_consolidados_${ANIO}.xlsx" || \
  echo "(aviso) no se pudo generar el consolidado de precios"

# ── 8) Empaquetar TODAS las salidas del corte en un zip ─────────────────────
ZIP="salidas_${CORTE}.zip"
rm -f "$ZIP"
zip -r "$ZIP" "$OUT" >/dev/null
echo
echo "== ZIP listo: $ZIP ($(du -h "$ZIP" | cut -f1)) =="
unzip -l "$ZIP" | tail -n +2 | head -40

# ── 9) Descargar el zip ─────────────────────────────────────────────────────
#     En Cloud Shell 'cloudshell download' abre la descarga en el navegador.
if command -v cloudshell >/dev/null 2>&1; then
  echo "Descargando $ZIP ..."
  cloudshell download "$REPO/$ZIP" || \
    echo "(si no abrio la descarga, corre: cloudshell download $REPO/$ZIP)"
else
  echo "Para descargar: cloudshell download $REPO/$ZIP"
  echo "O subir a GCP:   gsutil cp $REPO/$ZIP gs://pensiones-analytics-data/carga_libre/$ZIP"
fi
