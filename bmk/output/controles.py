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


def generar(clas: pd.DataFrame, out_dir: str | Path, corte: str,
            fecha_generacion: str | None = None) -> str:
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
    cols = ["AFP"] + [f"{p} ({PORT_NOMBRE[p]})" for p in PORTAFOLIOS] + ["TOTAL AFP", "Check"]
    wm.write_row(3, 0, cols, F["hdr"])
    row = 4
    first = row
    for afp in afps:
        wm.write(row, 0, afp, F["txt"])
        for k, p in enumerate(PORTAFOLIOS):
            wm.write_formula(row, 1 + k,
                f'=SUMIFS({rng("I")},{rng("B")},$A{row+1},{rng("A")},"{p}")', F["mill"])
        cini, cfin = chr(ord("B")), chr(ord("A") + len(PORTAFOLIOS))
        wm.write_formula(row, 1 + len(PORTAFOLIOS), f"=SUM({cini}{row+1}:{cfin}{row+1})", F["millb"])
        # Check: total fila = SUMIF por AFP directo
        tcol = chr(ord("A") + 1 + len(PORTAFOLIOS))
        wm.write_formula(row, 2 + len(PORTAFOLIOS),
            f'=IF(ABS({tcol}{row+1}-SUMIF({rng("B")},$A{row+1},{rng("I")}))<=1,"OK","REVISAR")')
        row += 1
    # Fila TOTAL
    wm.write(row, 0, "TOTAL", F["txtb"])
    for k in range(len(PORTAFOLIOS) + 1):
        c = chr(ord("B") + k)
        wm.write_formula(row, 1 + k, f"=SUM({c}{first+1}:{c}{row})", F["millb"])
    tcol = chr(ord("A") + 1 + len(PORTAFOLIOS))
    wm.write_formula(row, 2 + len(PORTAFOLIOS),
        f'=IF(ABS({tcol}{row+1}-{total_afp_cell})<=1,"OK","REVISAR")')
    wm.conditional_format(first, 2 + len(PORTAFOLIOS), row, 2 + len(PORTAFOLIOS),
        {"type": "cell", "criteria": "==", "value": '"OK"', "format": F["ok"]})
    wm.conditional_format(first, 2 + len(PORTAFOLIOS), row, 2 + len(PORTAFOLIOS),
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
            f"='AFP x Portafolio'!{chr(ord('A')+1+len(PORTAFOLIOS))}{4+len(afps)+1}")

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

    rs.write(rr + 2, 1, "Hojas: Controles (pass/fail) | AFP x Portafolio | AFP x Clasificacion | "
                        "Detalle por AFP | Datos.", F["subtitle"])

    wb.close()
    return dest
