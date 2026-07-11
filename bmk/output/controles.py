"""
Archivo de CONTROLES formulado (Excel con formulas vivas).

Genera un .xlsx con dos hojas:
  - "Datos": los activos consolidados (para que las formulas calculen sobre ellos).
  - "Controles": chequeos con formulas de Excel (COUNTA, COUNTIF, SUMIF, ...) que
    comparan un valor Esperado (calculado por el pipeline) contra un valor
    Calculado en vivo sobre la hoja Datos, con Diferencia y Estado (OK/REVISAR)
    con formato condicional. Al abrirlo en Excel recalcula; si se edita/pega
    data, los controles se actualizan y resaltan el error -> facilmente rastreable.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import xlsxwriter

# Columnas de la hoja Datos (orden fijo -> las formulas dependen de estas letras).
DATOS_COLS = ["portafolio", "afp", "clas_sfc", "clase_inversion", "clasificacion",
              "ubicacion", "riesgo", "is_clasificado", "vr_mercado", "valor_nominal",
              "nemo", "isin"]
COL = {c: chr(ord("A") + i) for i, c in enumerate(DATOS_COLS)}  # 'vr_mercado'->'I'


def _datos_df(clas: pd.DataFrame) -> pd.DataFrame:
    num = lambda c: pd.to_numeric(clas[c], errors="coerce").fillna(0.0)
    return pd.DataFrame({
        "portafolio": clas["cod_portafolio"], "afp": clas["afp"],
        "clas_sfc": clas["clas_sfc"], "clase_inversion": clas["clase_inversion"],
        "clasificacion": clas["clasificacion"], "ubicacion": clas["ubicacion"],
        "riesgo": clas["riesgo"], "is_clasificado": clas["is_clasificado"].astype(bool),
        "vr_mercado": num("vr_mercado"), "valor_nominal": num("valor_nominal"),
        "nemo": clas["nemo"].astype(str), "isin": clas["isin"].astype(str),
    })


def generar(clas: pd.DataFrame, out_dir: str | Path, corte: str) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = str(out_dir / f"controles_{corte}.xlsx")
    df = _datos_df(clas)
    N = len(df)
    r0, r1 = 2, N + 1  # rango de datos en Excel (fila 2 a N+1)

    wb = xlsxwriter.Workbook(dest, {"nan_inf_to_errors": True})
    f_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
    f_title = wb.add_format({"bold": True, "font_size": 13, "font_color": "#1F4E78"})
    f_sub = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    f_num = wb.add_format({"num_format": "#,##0"})
    f_num2 = wb.add_format({"num_format": "#,##0", "border": 1})
    f_txt = wb.add_format({"border": 1})
    f_ok = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1, "align": "center"})
    f_bad = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1, "align": "center", "bold": True})

    # ---------- Hoja Datos ----------
    ws = wb.add_worksheet("Datos")
    for j, c in enumerate(DATOS_COLS):
        ws.write(0, j, c, f_hdr)
    for j, c in enumerate(DATOS_COLS):
        colvals = df[c].tolist()
        for i, v in enumerate(colvals, start=1):
            ws.write(i, j, v)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, N, len(DATOS_COLS) - 1)

    # ---------- Hoja Controles ----------
    cs = wb.add_worksheet("Controles")
    cs.set_column("A:A", 2)
    cs.set_column("B:B", 42)
    cs.set_column("C:E", 20)
    cs.set_column("F:F", 12)
    cs.write("B1", f"Controles Benchmark Mensual - corte {corte}", f_title)
    cs.write_row("B3", ["Concepto", "Esperado", "Calculado (formula)", "Diferencia", "Estado"], f_hdr)

    fila = [3]  # puntero de fila (0-based); empieza escribiendo en fila 4 (idx 3 ya usado por hdr)
    fila[0] = 3

    def seccion(titulo):
        fila[0] += 1
        cs.write(fila[0], 1, titulo, f_sub)
        cs.write_blank(fila[0], 2, None, f_sub)
        cs.write_blank(fila[0], 3, None, f_sub)
        cs.write_blank(fila[0], 4, None, f_sub)
        cs.write_blank(fila[0], 5, None, f_sub)
        fila[0] += 1

    def control(concepto, esperado, formula_calc, tol=0, es_num=True):
        r = fila[0]
        cs.write(r, 1, concepto, f_txt)
        cs.write(r, 2, esperado, f_num2 if es_num else f_txt)
        cs.write_formula(r, 3, formula_calc, f_num2 if es_num else f_txt)
        # Diferencia = Calculado - Esperado
        cs.write_formula(r, 4, f"=D{r+1}-C{r+1}", f_num2)
        # Estado
        cs.write_formula(r, 5, f'=IF(ABS(E{r+1})<={tol},"OK","REVISAR")')
        fila[0] += 1
        return r

    rng = lambda col: f"Datos!${col}${r0}:${col}${r1}"

    # --- Seccion 1: Conteos ---
    seccion("1. Conteos de activos")
    total = N
    n_ok = int(df["is_clasificado"].sum())
    control("Total de activos consolidados", total, f"=COUNTA({rng('A')})")
    control("Activos clasificados (is_clasificado=VERDADERO)", n_ok,
            f"=COUNTIF({rng('H')},TRUE)")
    control("Activos sin clasificar (ND)", total - n_ok, f"=COUNTIF({rng('H')},FALSE)")

    # --- Seccion 2: Valor de mercado por AFP (cuadre) ---
    seccion("2. Cuadre de valor de mercado por AFP")
    total_mkt = float(df["vr_mercado"].sum())
    filas_afp = []
    for afp in sorted(df["afp"].unique()):
        esp = float(df.loc[df["afp"] == afp, "vr_mercado"].sum())
        r = control(f"Valor mercado - {afp}", round(esp, 2),
                    f'=SUMIF({rng("B")},"{afp}",{rng("I")})', tol=1)
        filas_afp.append(r + 1)
    suma_cells = "+".join(f"D{r}" for r in filas_afp)
    r = fila[0]
    cs.write(r, 1, "Suma subtotales AFP = Total industria", f_txt)
    cs.write(r, 2, round(total_mkt, 2), f_num2)
    cs.write_formula(r, 3, f"={suma_cells}", f_num2)
    cs.write_formula(r, 4, f"=D{r+1}-C{r+1}", f_num2)
    cs.write_formula(r, 5, '=IF(ABS(E{0})<=1,"OK","REVISAR")'.format(r + 1))
    fila[0] += 1

    # --- Seccion 3: Valor de mercado por clasificacion (cuadre) ---
    seccion("3. Cuadre de valor de mercado por clasificacion")
    filas_cl = []
    for cl in sorted([x for x in df["clasificacion"].unique() if str(x).strip()]):
        esp = float(df.loc[df["clasificacion"] == cl, "vr_mercado"].sum())
        r = control(f"Valor mercado - {cl or '(vacio)'}", round(esp, 2),
                    f'=SUMIF({rng("E")},"{cl}",{rng("I")})', tol=1)
        filas_cl.append(r + 1)
    suma_cells = "+".join(f"D{r}" for r in filas_cl)
    r = fila[0]
    cs.write(r, 1, "Suma subtotales clasificacion = Total industria", f_txt)
    cs.write(r, 2, round(total_mkt, 2), f_num2)
    cs.write_formula(r, 3, f"={suma_cells}", f_num2)
    cs.write_formula(r, 4, f"=D{r+1}-C{r+1}", f_num2)
    cs.write_formula(r, 5, '=IF(ABS(E{0})<=1,"OK","REVISAR")'.format(r + 1))
    fila[0] += 1

    # --- Seccion 4: Total y integridad ---
    seccion("4. Total e integridad")
    control("Valor de mercado total industria", round(total_mkt, 2),
            f"=SUM({rng('I')})", tol=1)
    control("Activos con valor de mercado = 0", int((df["vr_mercado"] == 0).sum()),
            f"=COUNTIF({rng('I')},0)")
    sin_id = int(((df["nemo"].str.strip() == "") & (df["isin"].str.strip() == "")).sum())
    control("Activos sin ISIN ni Nemo", sin_id,
            f'=COUNTIFS({rng("K")},"",{rng("L")},"")')

    # Formato condicional sobre la columna Estado (F4:F...)
    cs.conditional_format(3, 5, fila[0], 5, {"type": "cell", "criteria": "==",
                          "value": '"OK"', "format": f_ok})
    cs.conditional_format(3, 5, fila[0], 5, {"type": "cell", "criteria": "==",
                          "value": '"REVISAR"', "format": f_bad})
    cs.freeze_panes(3, 0)

    wb.close()
    return dest
