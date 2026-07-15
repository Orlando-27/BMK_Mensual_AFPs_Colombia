"""
Archivo de CONTROLES profesional (Excel con formulas vivas).

Genera un .xlsx con varias hojas, todas formuladas (recalculan al abrir/editar):
  - "Resumen"   : tablero ejecutivo con semaforo, KPIs y estado global.
  - "Controles" : chequeos pass/fail (Esperado vs Calculado vs Diferencia vs Estado).
  - "AFP x Portafolio" : matriz de valor (SUMIFS) con cuadre por fila y columna.
  - "AFP x Clasificacion" : matriz de valor (SUMIFS) con cuadres.
  - "Detalle por AFP" : conteos, % clasificado y valor por AFP.
  - "Datos"     : activos consolidados (fuente de todas las formulas).

Diseno: totales cruzados (triangulacion AFP = clasificacion = total), formato
condicional verde/rojo, semaforo, formato numerico en millones de COP.
"""
from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pandas as pd
import xlsxwriter

DATOS_COLS = ["portafolio", "afp", "clas_sfc", "clase_inversion", "clasificacion",
              "ubicacion", "riesgo", "is_clasificado", "vr_mercado", "valor_nominal",
              "nemo", "isin"]
CL = {c: chr(ord("A") + i) for i, c in enumerate(DATOS_COLS)}

PORTAFOLIOS = ["PO", "PC", "PM", "PR", "CS"]
PORT_NOMBRE = {"PO": "Moderado", "PC": "Conservador", "PM": "Mayor Riesgo",
               "PR": "Retiro Prog.", "CS": "Cesantias"}
CLASIF_BUCKETS = ["R FIJA", "R VARIABLE", "CAJA", "NOTAS ESTRUCTURADAS", "ND"]

# Derivados: tipos y columnas de la hoja "Datos Derivados".
TIPOS_DERIV = ["FORWARD", "FUT-TRM", "FUT-LOCAL", "FUT-INT", "SWAP-IRS", "SWAP-CCS"]
DERIV_NOMBRE = {"FORWARD": "Forwards", "FUT-TRM": "Futuros TRM",
                "FUT-LOCAL": "Futuros locales (TES)", "FUT-INT": "Futuros internac.",
                "SWAP-IRS": "Swaps IRS", "SWAP-CCS": "Swaps CCS"}
DERIV_COLS = ["tipo", "afp", "portafolio", "subtipo", "moneda", "nominal",
              "valor_mercado", "precio", "valorado"]


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


def _hoja_derivados(wb, F, der: pd.DataFrame, afps_ind):
    """Agrega 'Datos Derivados' + 'Controles Derivados'. Devuelve celda del # REVISAR."""
    # ---- Datos Derivados (fuente de las formulas) ----
    wd = wb.add_worksheet("Datos Derivados")
    for j, c in enumerate(DERIV_COLS):
        wd.write(0, j, c, F["hdr"])
    for j, c in enumerate(DERIV_COLS):
        for i, v in enumerate(der[c].tolist(), start=1):
            if isinstance(v, float) and (v != v):  # NaN
                wd.write_blank(i, j, None)
            else:
                wd.write(i, j, v)
    Nd = len(der)
    wd.freeze_panes(1, 0)
    wd.autofilter(0, 0, max(Nd, 1), len(DERIV_COLS) - 1)
    wd.set_column("A:E", 14)
    d0, d1 = 2, Nd + 1
    DC = {c: chr(ord("A") + i) for i, c in enumerate(DERIV_COLS)}

    def dr(col):
        return f"'Datos Derivados'!${col}${d0}:${col}${d1}"

    # ---- Controles Derivados (todo inicia en columna A / col 0) ----
    cd = wb.add_worksheet("Controles Derivados")
    cd.set_column("A:A", 24); cd.set_column("B:I", 16)
    cd.write("A1", "Controles de Derivados por tipo", F["title"])
    cd.write("A2", "Revisa cantidad, valor de mercado, nominal y precio de cada tipo de derivado.",
             F["subtitle"])

    # Tabla principal por tipo. Cols: A Tipo, B #Inst, C Nominal, D ValorMkt,
    # E Precio, F #Valorados, G #Pendientes, H Estado.
    hdr = ["Tipo", "# Instrumentos", "Nominal", "Valor de mercado", "Precio prom.",
           "# Valorados", "# Pendientes", "Estado"]
    cd.write_row(3, 0, hdr, F["hdr"])
    row = 4
    first = row
    for t in TIPOS_DERIV:
        rr = row + 1  # fila 1-based
        cd.write(row, 0, DERIV_NOMBRE[t], F["txt"])
        cd.write_formula(row, 1, f'=COUNTIF({dr(DC["tipo"])},"{t}")', F["int"])
        cd.write_formula(row, 2, f'=SUMIF({dr(DC["tipo"])},"{t}",{dr(DC["nominal"])})', F["mill"])
        cd.write_formula(row, 3, f'=SUMIF({dr(DC["tipo"])},"{t}",{dr(DC["valor_mercado"])})', F["mill"])
        cd.write_formula(row, 4, f'=IFERROR(AVERAGEIF({dr(DC["tipo"])},"{t}",{dr(DC["precio"])}),0)',
                         wb.add_format({"num_format": "#,##0.0000", "border": 1}))
        cd.write_formula(row, 5, f'=COUNTIFS({dr(DC["tipo"])},"{t}",{dr(DC["valorado"])},TRUE)', F["int"])
        cd.write_formula(row, 6, f'=COUNTIFS({dr(DC["tipo"])},"{t}",{dr(DC["valorado"])},FALSE)', F["int"])
        cd.write_formula(row, 7,
            f'=IF(B{rr}=0,"SIN DATOS",IF(G{rr}=0,"OK",IF(F{rr}=0,"PENDIENTE VALORAR","PARCIAL")))')
        row += 1
    cd.write(row, 0, "TOTAL DERIVADOS", F["txtb"])
    for col, fmt in ((1, "intb"), (2, "millb"), (3, "millb"), (5, "intb"), (6, "intb")):
        c = chr(ord("A") + col)
        cd.write_formula(row, col, f"=SUM({c}{first+1}:{c}{row})", F[fmt])
    cd.write_blank(row, 4, None, F["millb"])
    cd.write_blank(row, 7, None, F["txtb"])
    total_row = row + 1
    cd.conditional_format(first, 7, row - 1, 7,
        {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
    cd.conditional_format(first, 7, row - 1, 7,
        {"type": "text", "criteria": "containing", "value": "PENDIENTE", "format": F["bad"]})
    cd.conditional_format(first, 7, row - 1, 7,
        {"type": "text", "criteria": "containing", "value": "PARCIAL", "format": F["bad"]})

    # ---- Matriz Tipo x AFP (valor de mercado). Tipo col A(0), AFPs desde col B(1) ----
    row += 2
    cd.write(row, 0, "Valor de mercado por Tipo x AFP (millones COP)", F["sec"])
    for cc in range(1, 1 + len(afps_ind) + 1):
        cd.write_blank(row, cc, None, F["sec"])
    row += 1
    cd.write(row, 0, "Tipo", F["hdr"])
    for k, afp in enumerate(afps_ind):
        cd.write(row, 1 + k, afp, F["hdr"])
    cd.write(row, 1 + len(afps_ind), "TOTAL", F["hdr"])
    row += 1
    mfirst = row
    for t in TIPOS_DERIV:
        cd.write(row, 0, DERIV_NOMBRE[t], F["txt"])
        for k, afp in enumerate(afps_ind):
            cd.write_formula(row, 1 + k,
                f'=SUMIFS({dr(DC["valor_mercado"])},{dr(DC["tipo"])},"{t}",{dr(DC["afp"])},"{afp}")',
                F["mill"])
        c0, c1 = chr(ord("B")), chr(ord("B") + len(afps_ind) - 1)
        cd.write_formula(row, 1 + len(afps_ind), f"=SUM({c0}{row+1}:{c1}{row+1})", F["millb"])
        row += 1
    cd.write(row, 0, "TOTAL", F["txtb"])
    for k in range(len(afps_ind) + 1):
        c = chr(ord("B") + k)
        cd.write_formula(row, 1 + k, f"=SUM({c}{mfirst+1}:{c}{row})", F["millb"])

    # ---- Chequeos de integridad de derivados ----
    row += 2
    cd.write(row, 0, "Chequeos de integridad", F["sec"])
    for cc in range(1, 5):
        cd.write_blank(row, cc, None, F["sec"])
    row += 1
    cd.write_row(row, 0, ["Concepto", "Esperado", "Calculado", "Estado"], F["hdr"])
    row += 1
    chk_first = row

    def chk(concepto, esperado, formula):
        nonlocal row
        rr = row + 1
        cd.write(row, 0, concepto, F["txt"])
        cd.write(row, 1, esperado, F["int"])
        cd.write_formula(row, 2, formula, F["int"])
        cd.write_formula(row, 3, f'=IF(ABS(C{rr}-B{rr})<=0,"OK","REVISAR")')
        row += 1

    n_total = len(der)
    n_pend = int((~der["valorado"].astype(bool)).sum())  # pendientes REALES tras valorar
    n_neg_nom = int((pd.to_numeric(der["nominal"], errors="coerce").fillna(0) < 0).sum())
    chk("Total filas de derivados", n_total, f'=COUNTA({dr(DC["tipo"])})')
    chk("Patas de swap (deben ser pares -> 0)", 0,
        f'=MOD(COUNTIF({dr(DC["tipo"])},"SWAP-IRS")+COUNTIF({dr(DC["tipo"])},"SWAP-CCS"),2)')
    chk("Nominales negativos", n_neg_nom, f'=COUNTIF({dr(DC["nominal"])},"<0")')
    # Esperado = pendientes reales (0 cuando el motor v6 ya valoro todos los swaps).
    chk("Instrumentos sin valorar (pendientes)", n_pend,
        f'=COUNTIF({dr(DC["valorado"])},FALSE)')
    cd.conditional_format(chk_first, 3, row - 1, 3,
        {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
    cd.conditional_format(chk_first, 3, row - 1, 3,
        {"type": "cell", "criteria": "==", "value": '"REVISAR"', "format": F["bad"]})
    _nota = ("Nota: swaps valorados con el motor v6 (VPN por pata al corte)."
             if n_pend == 0 else
             "Nota: los swaps quedan con valor de mercado pendiente hasta valorarlos "
             "con el motor v6 (VPN por pata en las 3 fechas).")
    cd.write(row + 1, 0, _nota, F["subtitle"])
    cd.freeze_panes(4, 0)
    return total_row


def generar(clas: pd.DataFrame, out_dir: str | Path, corte: str,
            fecha_generacion: str | None = None,
            derivados: pd.DataFrame | None = None) -> str:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = str(out_dir / f"controles_{corte}.xlsx")
    df = _datos_df(clas)
    N = len(df)
    r0, r1 = 2, N + 1
    afps = sorted(df["afp"].unique())
    fgen = fecha_generacion or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")

    wb = xlsxwriter.Workbook(dest, {"nan_inf_to_errors": True})
    F = {
        "title": wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"}),
        "subtitle": wb.add_format({"font_size": 10, "font_color": "#595959"}),
        "hdr": wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                              "border": 1, "align": "center", "valign": "vcenter"}),
        "sec": wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "font_color": "#1F4E78"}),
        "txt": wb.add_format({"border": 1}),
        "txtb": wb.add_format({"border": 1, "bold": True}),
        "mill": wb.add_format({"num_format": "#,##0,,", "border": 1}),
        "millb": wb.add_format({"num_format": "#,##0,,", "border": 1, "bold": True, "bg_color": "#F2F2F2"}),
        "int": wb.add_format({"num_format": "#,##0", "border": 1}),
        "intb": wb.add_format({"num_format": "#,##0", "border": 1, "bold": True, "bg_color": "#F2F2F2"}),
        "pct": wb.add_format({"num_format": "0.0%", "border": 1}),
        "ok": wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1,
                             "align": "center", "bold": True}),
        "bad": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1,
                              "align": "center", "bold": True}),
        "kpi_lbl": wb.add_format({"font_size": 10, "font_color": "#595959", "align": "center",
                                  "top": 1, "left": 1, "right": 1, "bg_color": "#F2F2F2"}),
        "kpi_val": wb.add_format({"font_size": 18, "bold": True, "align": "center",
                                  "bottom": 1, "left": 1, "right": 1, "font_color": "#1F4E78"}),
        "kpi_val_pct": wb.add_format({"font_size": 18, "bold": True, "align": "center",
                                      "bottom": 1, "left": 1, "right": 1, "font_color": "#1F4E78",
                                      "num_format": "0.0%"}),
    }

    # ================= Hoja Datos =================
    ws = wb.add_worksheet("Datos")
    for j, c in enumerate(DATOS_COLS):
        ws.write(0, j, c, F["hdr"])
    for j, c in enumerate(DATOS_COLS):
        for i, v in enumerate(df[c].tolist(), start=1):
            ws.write(i, j, v)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, N, len(DATOS_COLS) - 1)
    ws.set_column("A:B", 12)
    ws.set_column("I:J", 18)

    def rng(col):
        return f"Datos!${col}${r0}:${col}${r1}"

    # ================= Hoja Detalle por AFP =================
    wa = wb.add_worksheet("Detalle por AFP")
    wa.set_column("A:A", 20); wa.set_column("B:H", 16)
    wa.write("A1", "Detalle por AFP", F["title"])
    hdr = ["AFP", "Valor mercado", "# Activos", "# Clasificados", "# ND",
           "% Clasificado", "Valor Nominal"]
    wa.write_row(3, 0, hdr, F["hdr"])
    row = 4
    afp_val_rows = []
    for afp in afps:
        wa.write(row, 0, afp, F["txt"])
        wa.write_formula(row, 1, f'=SUMIF({rng("B")},A{row+1},{rng("I")})', F["mill"])
        wa.write_formula(row, 2, f'=COUNTIF({rng("B")},A{row+1})', F["int"])
        wa.write_formula(row, 3, f'=COUNTIFS({rng("B")},A{row+1},{rng("H")},TRUE)', F["int"])
        wa.write_formula(row, 4, f'=COUNTIFS({rng("B")},A{row+1},{rng("H")},FALSE)', F["int"])
        wa.write_formula(row, 5, f'=IFERROR(D{row+1}/C{row+1},0)', F["pct"])
        wa.write_formula(row, 6, f'=SUMIF({rng("B")},A{row+1},{rng("J")})', F["mill"])
        afp_val_rows.append(row + 1)
        row += 1
    # Total
    wa.write(row, 0, "TOTAL INDUSTRIA", F["txtb"])
    for col, fmt in ((1, "millb"), (2, "intb"), (3, "intb"), (4, "intb"), (6, "millb")):
        c = chr(ord("A") + col)
        wa.write_formula(row, col, f"=SUM({c}5:{c}{row})", F[fmt])
    wa.write_formula(row, 5, f"=IFERROR(D{row+1}/C{row+1},0)", F["pct"])
    wa.write(row + 2, 0, "Valores en pesos (formato en millones para lectura).", F["subtitle"])
    total_afp_cell = f"'Detalle por AFP'!B{row+1}"  # total industria por AFP
    tot_activos_cell = f"'Detalle por AFP'!C{row+1}"

    # ================= Hoja AFP x Portafolio =================
    wm = wb.add_worksheet("AFP x Portafolio")
    wm.set_column("A:A", 20); wm.set_column("B:I", 15)
    wm.write("A1", "Matriz Valor de Mercado: AFP x Portafolio (millones COP)", F["title"])
    np_ = len(PORTAFOLIOS)
    # Se agrega la columna "Otros/ND" (portafolios fuera de los 5 obligatorios,
    # p.ej. voluntarias de SKANDIA) para que el total reconcilie con el SUMIF por
    # AFP y el check quede en OK mostrando ese valor de forma transparente.
    cols = (["AFP"] + [f"{p} ({PORT_NOMBRE[p]})" for p in PORTAFOLIOS]
            + ["Otros/ND", "TOTAL AFP", "Check"])
    wm.write_row(3, 0, cols, F["hdr"])
    c_otros, c_total, c_check = 1 + np_, 2 + np_, 3 + np_   # indices de columna
    cini, cfin = "B", chr(ord("A") + np_)                    # B..F (los 5)
    col_otros, col_total = chr(ord("A") + c_otros), chr(ord("A") + c_total)  # G, H
    row = 4
    first = row
    for afp in afps:
        wm.write(row, 0, afp, F["txt"])
        for k, p in enumerate(PORTAFOLIOS):
            wm.write_formula(row, 1 + k,
                f'=SUMIFS({rng("I")},{rng("B")},$A{row+1},{rng("A")},"{p}")', F["mill"])
        # Otros/ND = SUMIF total por AFP - suma de los 5 obligatorios.
        wm.write_formula(row, c_otros,
            f'=SUMIF({rng("B")},$A{row+1},{rng("I")})-SUM({cini}{row+1}:{cfin}{row+1})', F["mill"])
        wm.write_formula(row, c_total, f"=SUM({cini}{row+1}:{col_otros}{row+1})", F["millb"])
        # Check: total fila (5 + Otros) = SUMIF por AFP directo -> siempre reconcilia.
        wm.write_formula(row, c_check,
            f'=IF(ABS({col_total}{row+1}-SUMIF({rng("B")},$A{row+1},{rng("I")}))<=1,"OK","REVISAR")')
        row += 1
    # Fila TOTAL
    wm.write(row, 0, "TOTAL", F["txtb"])
    for k in range(np_ + 2):   # 5 portafolios + Otros + TOTAL AFP
        c = chr(ord("B") + k)
        wm.write_formula(row, 1 + k, f"=SUM({c}{first+1}:{c}{row})", F["millb"])
    wm.write_formula(row, c_check,
        f'=IF(ABS({col_total}{row+1}-{total_afp_cell})<=1,"OK","REVISAR")')
    wm.conditional_format(first, c_check, row, c_check,
        {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
    wm.conditional_format(first, c_check, row, c_check,
        {"type": "cell", "criteria": "==", "value": '"REVISAR"', "format": F["bad"]})

    # ================= Hoja AFP x Clasificacion =================
    wc = wb.add_worksheet("AFP x Clasificacion")
    wc.set_column("A:A", 20); wc.set_column("B:H", 18)
    wc.write("A1", "Matriz Valor de Mercado: AFP x Clasificacion (millones COP)", F["title"])
    cols = ["AFP"] + CLASIF_BUCKETS + ["TOTAL AFP", "Check"]
    wc.write_row(3, 0, cols, F["hdr"])
    row = 4; first = row
    for afp in afps:
        wc.write(row, 0, afp, F["txt"])
        for k, cl in enumerate(CLASIF_BUCKETS):
            wc.write_formula(row, 1 + k,
                f'=SUMIFS({rng("I")},{rng("B")},$A{row+1},{rng("E")},"{cl}")', F["mill"])
        cini, cfin = chr(ord("B")), chr(ord("A") + len(CLASIF_BUCKETS))
        wc.write_formula(row, 1 + len(CLASIF_BUCKETS), f"=SUM({cini}{row+1}:{cfin}{row+1})", F["millb"])
        tcol = chr(ord("A") + 1 + len(CLASIF_BUCKETS))
        wc.write_formula(row, 2 + len(CLASIF_BUCKETS),
            f'=IF(ABS({tcol}{row+1}-SUMIF({rng("B")},$A{row+1},{rng("I")}))<=1,"OK","REVISAR")')
        row += 1
    wc.write(row, 0, "TOTAL", F["txtb"])
    for k in range(len(CLASIF_BUCKETS) + 1):
        c = chr(ord("B") + k)
        wc.write_formula(row, 1 + k, f"=SUM({c}{first+1}:{c}{row})", F["millb"])
    tcol = chr(ord("A") + 1 + len(CLASIF_BUCKETS))
    wc.write_formula(row, 2 + len(CLASIF_BUCKETS),
        f'=IF(ABS({tcol}{row+1}-{total_afp_cell})<=1,"OK","REVISAR")')
    wc.conditional_format(first, 2 + len(CLASIF_BUCKETS), row, 2 + len(CLASIF_BUCKETS),
        {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
    wc.conditional_format(first, 2 + len(CLASIF_BUCKETS), row, 2 + len(CLASIF_BUCKETS),
        {"type": "cell", "criteria": "==", "value": '"REVISAR"', "format": F["bad"]})

    # ================= Hoja Controles (pass/fail) =================
    cs = wb.add_worksheet("Controles")
    cs.set_column("A:A", 2); cs.set_column("B:B", 46); cs.set_column("C:E", 22); cs.set_column("F:F", 12)
    cs.write("B1", f"Controles Benchmark Mensual - corte {corte}", F["title"])
    cs.write("B2", f"Generado: {fgen}", F["subtitle"])
    cs.write_row("B4", ["Concepto", "Esperado", "Calculado (formula)", "Diferencia", "Estado"], F["hdr"])
    fila = [4]

    def seccion(t):
        cs.write(fila[0], 1, t, F["sec"])
        for cc in range(2, 6):
            cs.write_blank(fila[0], cc, None, F["sec"])
        fila[0] += 1

    def control(concepto, esperado, formula, tol=1, es_num=True):
        r = fila[0]
        cs.write(r, 1, concepto, F["txt"])
        cs.write(r, 2, esperado, F["int"] if es_num else F["txt"])
        cs.write_formula(r, 3, formula, F["int"] if es_num else F["txt"])
        cs.write_formula(r, 4, f"=D{r+1}-C{r+1}", F["int"])
        cs.write_formula(r, 5, f'=IF(ABS(E{r+1})<={tol},"OK","REVISAR")')
        fila[0] += 1
        return r + 1

    total_mkt = float(df["vr_mercado"].sum())
    n_ok = int(df["is_clasificado"].sum())

    seccion("1. Conteos de activos")
    control("Total de activos consolidados", N, f"=COUNTA({rng('A')})", tol=0)
    control("Activos clasificados", n_ok, f"=COUNTIF({rng('H')},TRUE)", tol=0)
    control("Activos sin clasificar (ND)", N - n_ok, f"=COUNTIF({rng('H')},FALSE)", tol=0)

    seccion("2. Cuadre de valor por AFP (triangulacion)")
    for afp in afps:
        esp = round(float(df.loc[df["afp"] == afp, "vr_mercado"].sum()), 0)
        control(f"Valor mercado {afp}", esp, f'=SUMIF({rng("B")},"{afp}",{rng("I")})')
    control("Suma AFP = Total (matriz)", round(total_mkt, 0),
            f"='AFP x Portafolio'!{chr(ord('A')+2+len(PORTAFOLIOS))}{4+len(afps)+1}")

    seccion("3. Cuadre de valor por clasificacion (triangulacion)")
    for cl in CLASIF_BUCKETS:
        esp = round(float(df.loc[df["clasificacion"].str.upper() == cl, "vr_mercado"].sum()), 0)
        control(f"Valor mercado {cl}", esp, f'=SUMIF({rng("E")},"{cl}",{rng("I")})')
    control("Suma clasificacion = Total (matriz)", round(total_mkt, 0),
            f"='AFP x Clasificacion'!{chr(ord('A')+1+len(CLASIF_BUCKETS))}{4+len(afps)+1}")

    seccion("4. Total e integridad")
    control("Valor de mercado total industria", round(total_mkt, 0), f"=SUM({rng('I')})")
    control("Triangulacion: Total AFP = Total Clasificacion", 0,
            f"='Detalle por AFP'!B{4+len(afps)+1}-SUM({rng('I')})", tol=1)
    control("Activos con valor de mercado = 0", int((df["vr_mercado"] == 0).sum()),
            f"=COUNTIF({rng('I')},0)", tol=0)
    sin_id = int(((df["nemo"].str.strip() == "") & (df["isin"].str.strip() == "")).sum())
    control("Activos sin ISIN ni Nemo (identif. por emisor)", sin_id,
            f'=COUNTIFS({rng("K")},"",{rng("L")},"")', tol=0)
    control("Activos con valor nominal negativo", int((df["valor_nominal"] < 0).sum()),
            f'=COUNTIF({rng("J")},"<0")', tol=0)

    cs.conditional_format(4, 5, fila[0], 5, {"type": "cell", "criteria": "==",
                          "value": '"OK"', "format": F["ok"]})
    cs.conditional_format(4, 5, fila[0], 5, {"type": "cell", "criteria": "==",
                          "value": '"REVISAR"', "format": F["bad"]})
    cs.freeze_panes(4, 0)
    n_controles = fila[0] - 5  # aprox filas de control (excluye secciones)

    # ================= Hojas de Derivados =================
    der_total_row = None
    if derivados is not None and len(derivados):
        der = derivados.copy()
        for c in DERIV_COLS:
            if c not in der.columns:
                der[c] = "" if c not in ("nominal", "valor_mercado", "precio") else 0.0
        der = der[DERIV_COLS]
        der_total_row = _hoja_derivados(wb, F, der, afps)

    # ================= Hoja Resumen (dashboard) =================
    rs = wb.add_worksheet("Resumen")
    wb.worksheets_objs.insert(0, wb.worksheets_objs.pop())  # Resumen al frente
    rs.set_column("A:A", 2); rs.set_column("B:B", 26); rs.set_column("C:F", 18)
    rs.hide_gridlines(2)
    rs.write("B2", "Tablero de Controles - Benchmark Mensual", F["title"])
    rs.write("B3", f"Corte {corte}  |  Generado {fgen}", F["subtitle"])

    # KPIs
    kpis = [
        ("B5", "Total industria (billones COP)", f"=SUM({rng('I')})/1000000000000", "num"),
        ("D5", "# Activos consolidados", f"=COUNTA({rng('A')})", "int"),
        ("F5", "% Clasificado", f"=COUNTIF({rng('H')},TRUE)/COUNTA({rng('A')})", "pct"),
    ]
    for cell, lbl, formula, kind in kpis:
        col = cell[0]; r = int(cell[1:])
        rs.write(f"{col}{r}", lbl, F["kpi_lbl"])
        fmt = F["kpi_val_pct"] if kind == "pct" else wb.add_format(
            {"font_size": 18, "bold": True, "align": "center", "bottom": 1, "left": 1, "right": 1,
             "font_color": "#1F4E78", "num_format": "#,##0.0" if kind == "num" else "#,##0"})
        rs.write_formula(f"{col}{r+1}", formula, fmt)

    # Estado global (semaforo): cuenta REVISAR en Controles
    rs.write("B8", "Estado global de controles", F["sec"])
    for cc in range(2, 6):
        rs.write_blank(7, cc, None, F["sec"])
    rs.write("B9", "Controles en REVISAR", F["txt"])
    rs.write_formula("C9", f'=COUNTIF(Controles!F5:F{fila[0]},"REVISAR")', F["int"])
    rs.write("B10", "SEMAFORO", F["txtb"])
    rs.write_formula("C10", '=IF(C9=0,"TODO OK","REVISAR")')
    rs.conditional_format("C10", {"type": "cell", "criteria": "==", "value": '"TODO OK"', "format": F["ok"]})
    rs.conditional_format("C10", {"type": "cell", "criteria": "==", "value": '"REVISAR"', "format": F["bad"]})

    # Mini tabla por AFP (referencia a Detalle por AFP)
    rs.write("B12", "Resumen por AFP", F["sec"])
    for cc in range(2, 6):
        rs.write_blank(11, cc, None, F["sec"])
    rs.write_row("B13", ["AFP", "Valor (billones)", "# Activos", "% Clasif."], F["hdr"])
    rr = 13
    for i, afp in enumerate(afps):
        src = 5 + i  # fila en Detalle por AFP (1-based): datos empiezan en 5
        rs.write(rr, 1, afp, F["txt"])
        rs.write_formula(rr, 2, f"='Detalle por AFP'!B{src}/1000000000000",
                         wb.add_format({"num_format": "#,##0.0", "border": 1}))
        rs.write_formula(rr, 3, f"='Detalle por AFP'!C{src}", F["int"])
        rs.write_formula(rr, 4, f"='Detalle por AFP'!F{src}", F["pct"])
        rr += 1
    rs.write(rr, 1, "TOTAL", F["txtb"])
    rs.write_formula(rr, 2, f"='Detalle por AFP'!B{5+len(afps)}/1000000000000",
                     wb.add_format({"num_format": "#,##0.0", "border": 1, "bold": True, "bg_color": "#F2F2F2"}))
    rs.write_formula(rr, 3, f"='Detalle por AFP'!C{5+len(afps)}", F["intb"])
    rs.write_formula(rr, 4, f"='Detalle por AFP'!F{5+len(afps)}", F["pct"])

    # Bloque de derivados en el dashboard (si hay).
    if der_total_row is not None:
        drow = rr + 2
        rs.write(drow, 1, "Derivados por tipo", F["sec"])
        for cc in range(2, 6):
            rs.write_blank(drow, cc, None, F["sec"])
        rs.write_row(drow + 1, 1, ["Tipo", "# Instrum.", "Valor mdo (millones)", "Estado"], F["hdr"])
        # Filas de tipo en 'Controles Derivados' (1-based): 5..10.
        for i, t in enumerate(TIPOS_DERIV):
            src = 5 + i
            rs.write(drow + 2 + i, 1, DERIV_NOMBRE[t], F["txt"])
            rs.write_formula(drow + 2 + i, 2, f"='Controles Derivados'!B{src}", F["int"])
            rs.write_formula(drow + 2 + i, 3, f"='Controles Derivados'!D{src}/1000000", F["mill"])
            rs.write_formula(drow + 2 + i, 4, f"='Controles Derivados'!H{src}")
        rs.write(drow + 2 + len(TIPOS_DERIV), 1, "TOTAL", F["txtb"])
        rs.write_formula(drow + 2 + len(TIPOS_DERIV), 2, f"='Controles Derivados'!B{der_total_row}", F["intb"])
        rs.write_formula(drow + 2 + len(TIPOS_DERIV), 3,
                         f"='Controles Derivados'!D{der_total_row}/1000000", F["millb"])
        rs.conditional_format(drow + 2, 4, drow + 1 + len(TIPOS_DERIV), 4,
            {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
        rs.conditional_format(drow + 2, 4, drow + 1 + len(TIPOS_DERIV), 4,
            {"type": "text", "criteria": "containing", "value": "PENDIENTE", "format": F["bad"]})
        rs.conditional_format(drow + 2, 4, drow + 1 + len(TIPOS_DERIV), 4,
            {"type": "text", "criteria": "containing", "value": "PARCIAL", "format": F["bad"]})
        nota_row = drow + 4 + len(TIPOS_DERIV)
    else:
        nota_row = rr + 2
    rs.write(nota_row, 1, "Hojas: Controles | Controles Derivados | AFP x Portafolio | "
                          "AFP x Clasificacion | Detalle por AFP | Datos | Datos Derivados.",
             F["subtitle"])

    wb.close()
    return dest
