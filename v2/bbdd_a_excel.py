# -*- coding: utf-8 -*-
"""
bbdd_a_excel.py — Convierte la BBDD (v2/bbdd/*.parquet) en UN Excel para compartir.
====================================================================================
Pensado para usuarios NO técnicos: un solo archivo .xlsx con hojas legibles, donde las
columnas ya traen los NOMBRES pegados (VP, Gerencia, Desc. CECO, Ítem, Desc. CLACO…), sin
tener que cruzar tablas. Formato amigable (encabezado, filtros, separador de miles) + hoja
"Léame".

Requiere haber corrido antes `python construir_v2.py` (que genera v2/bbdd/).
Uso:
  python bbdd_a_excel.py               → liviano: agregados + diccionarios (recomendado para compartir)
  python bbdd_a_excel.py --completo    → además el DETALLE línea-a-línea (~238k filas, archivo pesado)

Escribe: v2/bbdd/BBDD Dashboard Corporativo.xlsx
"""
import os
import sys
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
BBDD = os.path.join(HERE, "bbdd")
OUT_BASE = os.path.join(BBDD, "BBDD Dashboard Corporativo.xlsx")
OUT_FULL = os.path.join(BBDD, "BBDD Dashboard Corporativo (con detalle).xlsx")

TEAL = "14515A"          # encabezado (marca AMSA)
BUCKET_LBL = {"2022": "2022", "2023": "2023", "2024": "2024", "2025": "2025",
              "2026ytd": "2026 YTD (ene–may)", "2026fy": "2026 FY (Ppto anual)"}
MEDIDA_LBL = {"forecast_2026": "Forecast 5+7 2026", "ppto_2027": "Ppto 2027"}
SRC_LBL = {"propios": "Propios", "contratista": "Contratista"}


def log(m):
    try:
        print(m)
    except Exception:
        print(m.encode("ascii", "replace").decode())


def _load(name):
    p = os.path.join(BBDD, name + ".parquet")
    if not os.path.isfile(p):
        raise SystemExit(f"No encontré {p}. Corré primero: python construir_v2.py")
    return pd.read_parquet(p)


def construir_hojas(completo):
    """Devuelve una lista de (nombre_hoja, DataFrame, [columnas de valor a formatear con miles])."""
    reg = _load("fact_registros")
    anu = _load("fact_anual")
    dot = _load("fact_dotaciones")
    cec = _load("dim_cecos")
    itm = _load("dim_items")
    cla = _load("dim_clacos")
    comp = _load("dim_companias")

    # Diccionarios recortados para pegar nombres (CECO por estructura NUEVA).
    cec_new = (cec[cec.estructura == "new"][["ceco", "vp", "gerencia", "desc_ceco"]]
               .drop_duplicates("ceco"))
    itm2 = itm[["item", "nombre", "tipo_costo", "item_relevante_nombre", "clasif_cuenta"]] \
        .rename(columns={"nombre": "item_nombre", "item_relevante_nombre": "itemrel_nombre"})
    cla2 = cla[["claco", "desc_claco"]]
    comp2 = comp[["codigo", "nombre"]].rename(columns={"nombre": "compania_nombre"})

    def pegar(df):
        df = df.copy()
        df["codigo"] = df["ceco"].str[:4]
        df = df.merge(cec_new, on="ceco", how="left").merge(itm2, on="item", how="left")
        if "claco" in df.columns:   # el detalle Real no tiene CLACO en su grano
            df = df.merge(cla2, on="claco", how="left")
        return df.merge(comp2, on="codigo", how="left")

    hojas = []

    # 1) Registros (Real + Ppto histórico) — enriquecido y con encabezados legibles.
    r = pegar(reg)
    r["Período"] = r["bucket"].map(BUCKET_LBL).fillna(r["bucket"])
    registros = pd.DataFrame({
        "VP": r["vp"], "Gerencia": r["gerencia"], "Desc. CECO": r["desc_ceco"], "CECO": r["ceco"],
        "Compañía": r["compania_nombre"], "Ítem Relevante": r["itemrel_nombre"], "Ítem": r["item_nombre"],
        "Cód. Ítem": r["item"], "Tipo Costo": r["tipo_costo"], "Clasif. Cuenta": r["clasif_cuenta"],
        "Desc. CLACO": r["desc_claco"], "CLACO": r["claco"], "Contrapartida": r["contra"],
        "Período": r["Período"],
        "Real (USD)": r["real_n"], "Real (Aj. 2027)": r["real_a"],
        "Ppto (USD)": r["plan_n"], "Ppto (Aj. 2027)": r["plan_a"],
    })
    hojas.append(("Registros", registros,
                  ["Real (USD)", "Real (Aj. 2027)", "Ppto (USD)", "Ppto (Aj. 2027)"]))

    # 2) Forecast y Ppto 2027 (anual).
    a = pegar(anu)
    anual = pd.DataFrame({
        "VP": a["vp"], "Gerencia": a["gerencia"], "Desc. CECO": a["desc_ceco"], "CECO": a["ceco"],
        "Compañía": a["compania_nombre"], "Ítem Relevante": a["itemrel_nombre"], "Ítem": a["item_nombre"],
        "Cód. Ítem": a["item"], "Tipo Costo": a["tipo_costo"], "Desc. CLACO": a["desc_claco"], "CLACO": a["claco"],
        "Medida": a["medida"].map(MEDIDA_LBL).fillna(a["medida"]),
        "Valor (USD)": a["valor_n"], "Valor (Aj. 2027)": a["valor_a"],
    })
    hojas.append(("Forecast y Ppto 2027", anual, ["Valor (USD)", "Valor (Aj. 2027)"]))

    # 3) Dotaciones (FTE).
    dotaciones = pd.DataFrame({
        "Tipo": dot["src"].map(SRC_LBL).fillna(dot["src"]), "VP": dot["vp"], "Gerencia": dot["ger"],
        "Año": dot["year"], "Período": dot["period"], "Real (FTE)": dot["real"], "Ppto (FTE)": dot["plan"],
    })
    hojas.append(("Dotaciones (FTE)", dotaciones, ["Real (FTE)", "Ppto (FTE)"]))

    # 4-7) Diccionarios (referencia).
    hojas.append(("Dicc. CECOs", cec.rename(columns={
        "ceco": "CECO", "estructura": "Estructura", "vp": "VP", "gerencia": "Gerencia",
        "desc_ceco": "Desc. CECO", "tipo_costo": "Tipo Costo", "aplica": "¿Aplica?",
        "clasificacion": "Clasificación", "compania": "Cód. Compañía"}), []))
    hojas.append(("Dicc. Ítems", itm.rename(columns={
        "item": "Cód. Ítem", "nombre": "Ítem", "tipo_costo": "Tipo Costo",
        "item_relevante": "Cód. Ítem Relevante", "item_relevante_nombre": "Ítem Relevante",
        "clasif_cuenta": "Clasif. Cuenta"}), []))
    hojas.append(("Dicc. CLACOs", cla.rename(columns={"claco": "CLACO", "desc_claco": "Desc. CLACO"}), []))
    hojas.append(("Dicc. Compañías", comp.rename(columns={
        "codigo": "Código", "nombre": "Nombre", "abrev": "Abrev.", "clasificacion": "Clasificación"}), []))

    # 8-9) Detalle línea-a-línea (solo con --completo; pesado).
    if completo:
        dr = pegar(_load("fact_detalle_real"))
        dr["Período"] = dr["bucket"].map(BUCKET_LBL).fillna(dr["bucket"])
        det_real = pd.DataFrame({
            "VP": dr["vp"], "Gerencia": dr["gerencia"], "Desc. CECO": dr["desc_ceco"], "CECO": dr["ceco"],
            "Ítem": dr["item_nombre"], "Contrapartida": dr["contra"], "Texto pedido": dr["texto_pedido"],
            "Denominación": dr["denominacion"], "Documento": dr["documento"], "Período": dr["Período"],
            "Real (USD)": dr["valor_n"], "Real (Aj. 2027)": dr["valor_a"],
        })
        hojas.append(("Detalle Real", det_real, ["Real (USD)", "Real (Aj. 2027)"]))

        dp = _load("fact_detalle_ppto").merge(
            itm2[["item", "item_nombre"]], on="item", how="left")
        det_ppto = pd.DataFrame({
            "CECO": dp["ceco"], "Ítem": dp["item_nombre"], "Concepto Gasto": dp["concepto_gasto"],
            "Actividad": dp["actividad"], "Medida": dp["medida"].map(MEDIDA_LBL).fillna(dp["medida"]),
            "Valor (USD)": dp["valor_n"], "Valor (Aj. 2027)": dp["valor_a"],
        })
        hojas.append(("Detalle Ppto", det_ppto, ["Valor (USD)", "Valor (Aj. 2027)"]))

    return hojas


def _formatear(ws, valcols, headers):
    """Encabezado con estilo, panel congelado, autofiltro, anchos y formato de miles."""
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor=TEAL)
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = min(max(len(str(h)) + 2, 11), 46)
    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    # Formato de miles solo en columnas de valor (evita recorrer todo en hojas enormes).
    if valcols and ws.max_row > 1:
        idx = {h: j + 1 for j, h in enumerate(headers)}
        es_fte = any("FTE" in v for v in valcols)
        fmt = "#,##0.0" if es_fte else "#,##0"
        for v in valcols:
            j = idx.get(v)
            if not j:
                continue
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j, max_row=ws.max_row):
                row[0].number_format = fmt


def _hoja_leeme(ws, hojas, completo):
    ws.sheet_view.showGridLines = False
    ws["A1"] = "BBDD del Dashboard Corporativo · Ppto 2027"
    ws["A1"].font = Font(bold=True, size=15, color=TEAL)
    intro = [
        "",
        "Este archivo es un extracto de la base de datos del tablero, en un solo Excel.",
        "Cada hoja ya trae los NOMBRES (VP, Gerencia, Ítem, Desc. CLACO…): no hace falta cruzar nada.",
        "Valores en USD. «Aj. 2027» = misma cifra reexpresada a moneda equivalente 2027.",
        "Generado con bbdd_a_excel.py a partir de v2/bbdd/ (regenerable).",
        "",
    ]
    r = 2
    for t in intro:
        ws.cell(row=r, column=1, value=t); r += 1
    ws.cell(row=r, column=1, value="Hoja").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=TEAL)
    ws.cell(row=r, column=2, value="Qué contiene").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=TEAL)
    r += 1
    desc = {
        "Registros": "Real y Presupuesto histórico (2022–2025, 2026 YTD y 2026 FY) por CECO / Ítem / CLACO / Contrapartida.",
        "Forecast y Ppto 2027": "Forecast 5+7 2026 y Presupuesto 2027 (anual) por CECO / Ítem / CLACO.",
        "Dotaciones (FTE)": "Dotación (personas) Propios y Contratista por VP / Gerencia, año y período.",
        "Dicc. CECOs": "Catálogo de CECO: VP, Gerencia, Desc. CECO (estructuras Nueva y Antigua).",
        "Dicc. Ítems": "Catálogo de Ítem: nombre, Tipo Costo, Ítem Relevante, Clasificación Cuenta.",
        "Dicc. CLACOs": "Catálogo de CLACO (Clase de Costo) → Desc. CLACO.",
        "Dicc. Compañías": "Catálogo de compañías.",
        "Detalle Real": "Gasto Real línea a línea (Contrapartida › Texto pedido › Denominación › Documento).",
        "Detalle Ppto": "Presupuesto/Forecast a nivel Concepto Gasto › Actividad.",
    }
    for name, _df, _v in hojas:
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=desc.get(name, ""))
        r += 1
    if not completo:
        ws.cell(row=r + 1, column=1,
                value="Nota: el detalle línea-a-línea (Detalle Real / Detalle Ppto) no se incluyó para mantener "
                      "el archivo liviano. Para agregarlo: python bbdd_a_excel.py --completo")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 95


def main():
    completo = "--completo" in sys.argv
    out = OUT_FULL if completo else OUT_BASE
    log("Leyendo BBDD y armando hojas…")
    hojas = construir_hojas(completo)
    log(f"  {len(hojas)} hojas" + (" (incluye detalle línea-a-línea)" if completo else " (liviano, sin detalle)"))
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df, valcols in hojas:
            df.to_excel(xw, sheet_name=name[:31], index=False)
            _formatear(xw.book[name[:31]], valcols, list(df.columns))
            log(f"    {name}: {len(df):,} filas")
        # Léame al frente (index 0) y limpieza de la hoja default de openpyxl si quedó.
        _hoja_leeme(xw.book.create_sheet("Léame", 0), hojas, completo)
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    mb = os.path.getsize(out) / 1e6
    log(f"\n✓ LISTO: {out} ({mb:.1f} MB)")


if __name__ == "__main__":
    main()
