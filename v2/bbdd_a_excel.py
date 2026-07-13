# -*- coding: utf-8 -*-
"""
bbdd_a_excel.py — Convierte la BBDD (v2/bbdd/*.parquet) en Excel para compartir.
================================================================================
Pensado para usuarios NO técnicos: hojas legibles con los NOMBRES ya pegados (VP, Gerencia,
Desc. CECO, Ítem, Desc. CLACO…), sin cruzar tablas. Formato amigable + hoja "Léame".

Requiere haber corrido antes `python construir_v2.py` (genera v2/bbdd/).
Uso:
  python bbdd_a_excel.py             → 1 Excel global liviano (agregados + diccionarios)
  python bbdd_a_excel.py --completo  → 1 Excel global + DETALLE línea-a-línea (~238k filas)
  python bbdd_a_excel.py --por-vp    → 1 Excel POR VP (data completa de esa VP) en cada
                                        v3/por_vp/<VP>/BBDD <VP>.xlsx  (apoyo del dashboard por VP)

El Excel usa SOLO la estructura NUEVA (la vigente): cada CECO se asigna a su VP de la
estructura nueva. (Los pocos CECOs que solo existían en la estructura antigua no se incluyen.)
"""
import os
import re
import sys
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
BBDD = os.path.join(HERE, "bbdd")
V3_POR_VP = os.path.join(os.path.dirname(HERE), "v3", "por_vp")
OUT_BASE = os.path.join(BBDD, "BBDD Dashboard Corporativo.xlsx")
OUT_FULL = os.path.join(BBDD, "BBDD Dashboard Corporativo (con detalle).xlsx")

TEAL = "14515A"
BUCKET_LBL = {"2022": "2022", "2023": "2023", "2024": "2024", "2025": "2025",
              "2026ytd": "2026 YTD (ene–may)", "2026fy": "2026 FY (Ppto anual)"}
MEDIDA_LBL = {"forecast_2026": "Forecast 5+7 2026", "ppto_2027": "Ppto 2027"}
SRC_LBL = {"propios": "Propios", "contratista": "Contratista"}


def log(m):
    try:
        print(m)
    except Exception:
        print(m.encode("ascii", "replace").decode())


def _safe(name):   # nombre de carpeta/archivo (igual que dividir_por_vp)
    return re.sub(r'[\\/:*?"<>|]+', " ", str(name)).strip() or "SIN VP"


def _load(name):
    p = os.path.join(BBDD, name + ".parquet")
    if not os.path.isfile(p):
        raise SystemExit(f"No encontré {p}. Corré primero: python construir_v2.py")
    return pd.read_parquet(p)


def cargar(con_detalle):
    """Carga las tablas parquet UNA vez (para filtrar luego por VP sin releer)."""
    dfs = {n: _load(n) for n in ("fact_registros", "fact_anual", "fact_dotaciones",
                                 "dim_cecos", "dim_items", "dim_clacos", "dim_companias")}
    if con_detalle:
        dfs["fact_detalle_real"] = _load("fact_detalle_real")
        dfs["fact_detalle_ppto"] = _load("fact_detalle_ppto")
    return dfs


def preparar(dfs):
    """Diccionarios para pegar nombres. cec_enrich: ceco → VP/Gerencia/Desc. CECO según la
    estructura NUEVA (la vigente); es la única que se usa en el Excel."""
    cec = dfs["dim_cecos"]
    new = cec[cec.estructura == "new"].drop_duplicates("ceco")
    cec_enrich = new[["ceco", "vp", "gerencia", "desc_ceco"]].reset_index(drop=True)
    itm2 = dfs["dim_items"][["item", "nombre", "tipo_costo", "item_relevante_nombre", "clasif_cuenta"]] \
        .rename(columns={"nombre": "item_nombre", "item_relevante_nombre": "itemrel_nombre"})
    cla2 = dfs["dim_clacos"][["claco", "desc_claco"]]
    comp2 = dfs["dim_companias"][["codigo", "nombre"]].rename(columns={"nombre": "compania_nombre"})
    return {"cec_enrich": cec_enrich, "itm2": itm2, "cla2": cla2, "comp2": comp2}


def _pegar(df, P):
    df = df.copy()
    df["codigo"] = df["ceco"].str[:4]
    df = df.merge(P["cec_enrich"], on="ceco", how="left").merge(P["itm2"], on="item", how="left")
    if "claco" in df.columns:   # el detalle Real no tiene CLACO en su grano
        df = df.merge(P["cla2"], on="claco", how="left")
    return df.merge(P["comp2"], on="codigo", how="left")


def construir_hojas(dfs, P, completo, cecos=None):
    """Arma (nombre_hoja, DataFrame, [cols de valor]). Si cecos ≠ None, filtra a esos CECOs
    (modo por VP) y omite Dotaciones (usan otra convención de VP)."""
    def filt(df):
        return df if cecos is None else df[df["ceco"].isin(cecos)]

    hojas = []

    # 1) Registros (Real + Ppto histórico).
    r = _pegar(filt(dfs["fact_registros"]), P)
    r["Período"] = r["bucket"].map(BUCKET_LBL).fillna(r["bucket"])
    hojas.append(("Registros", pd.DataFrame({
        "VP": r["vp"], "Gerencia": r["gerencia"], "Desc. CECO": r["desc_ceco"], "CECO": r["ceco"],
        "Compañía": r["compania_nombre"], "Ítem Relevante": r["itemrel_nombre"], "Ítem": r["item_nombre"],
        "Cód. Ítem": r["item"], "Tipo Costo": r["tipo_costo"], "Clasif. Cuenta": r["clasif_cuenta"],
        "Desc. CLACO": r["desc_claco"], "CLACO": r["claco"], "Contrapartida": r["contra"], "Período": r["Período"],
        "Real (USD)": r["real_n"], "Real (Aj. 2027)": r["real_a"], "Ppto (USD)": r["plan_n"], "Ppto (Aj. 2027)": r["plan_a"],
    }), ["Real (USD)", "Real (Aj. 2027)", "Ppto (USD)", "Ppto (Aj. 2027)"]))

    # 2) Forecast y Ppto 2027 (anual).
    a = _pegar(filt(dfs["fact_anual"]), P)
    hojas.append(("Forecast y Ppto 2027", pd.DataFrame({
        "VP": a["vp"], "Gerencia": a["gerencia"], "Desc. CECO": a["desc_ceco"], "CECO": a["ceco"],
        "Compañía": a["compania_nombre"], "Ítem Relevante": a["itemrel_nombre"], "Ítem": a["item_nombre"],
        "Cód. Ítem": a["item"], "Tipo Costo": a["tipo_costo"], "Desc. CLACO": a["desc_claco"], "CLACO": a["claco"],
        "Medida": a["medida"].map(MEDIDA_LBL).fillna(a["medida"]),
        "Valor (USD)": a["valor_n"], "Valor (Aj. 2027)": a["valor_a"],
    }), ["Valor (USD)", "Valor (Aj. 2027)"]))

    # 3) Dotaciones (FTE) — solo en el Excel global (usan convención de VP distinta).
    if cecos is None:
        dot = dfs["fact_dotaciones"]
        hojas.append(("Dotaciones (FTE)", pd.DataFrame({
            "Tipo": dot["src"].map(SRC_LBL).fillna(dot["src"]), "VP": dot["vp"], "Gerencia": dot["ger"],
            "Año": dot["year"], "Período": dot["period"], "Real (FTE)": dot["real"], "Ppto (FTE)": dot["plan"],
        }), ["Real (FTE)", "Ppto (FTE)"]))

    # 4-7) Diccionarios (dim_cecos: SOLO estructura nueva/vigente; acotado a los CECOs del archivo).
    cec = dfs["dim_cecos"]
    cec = cec[cec.estructura == "new"].drop(columns=["estructura"])
    if cecos is not None:
        cec = cec[cec.ceco.isin(cecos)]
    hojas.append(("Dicc. CECOs", cec.rename(columns={
        "ceco": "CECO", "vp": "VP", "gerencia": "Gerencia", "desc_ceco": "Desc. CECO",
        "tipo_costo": "Tipo Costo", "aplica": "¿Aplica?", "clasificacion": "Clasificación", "compania": "Cód. Compañía"}), []))
    hojas.append(("Dicc. Ítems", dfs["dim_items"].rename(columns={
        "item": "Cód. Ítem", "nombre": "Ítem", "tipo_costo": "Tipo Costo", "item_relevante": "Cód. Ítem Relevante",
        "item_relevante_nombre": "Ítem Relevante", "clasif_cuenta": "Clasif. Cuenta"}), []))
    hojas.append(("Dicc. CLACOs", dfs["dim_clacos"].rename(columns={"claco": "CLACO", "desc_claco": "Desc. CLACO"}), []))
    hojas.append(("Dicc. Compañías", dfs["dim_companias"].rename(columns={
        "codigo": "Código", "nombre": "Nombre", "abrev": "Abrev.", "clasificacion": "Clasificación"}), []))

    # 8-9) Detalle línea-a-línea (con --completo o por VP).
    if completo:
        dr = _pegar(filt(dfs["fact_detalle_real"]), P)
        dr["Período"] = dr["bucket"].map(BUCKET_LBL).fillna(dr["bucket"])
        hojas.append(("Detalle Real", pd.DataFrame({
            "VP": dr["vp"], "Gerencia": dr["gerencia"], "Desc. CECO": dr["desc_ceco"], "CECO": dr["ceco"],
            "Ítem": dr["item_nombre"], "Contrapartida": dr["contra"], "Texto pedido": dr["texto_pedido"],
            "Denominación": dr["denominacion"], "Documento": dr["documento"], "Período": dr["Período"],
            "Real (USD)": dr["valor_n"], "Real (Aj. 2027)": dr["valor_a"],
        }), ["Real (USD)", "Real (Aj. 2027)"]))
        dp = filt(dfs["fact_detalle_ppto"]).merge(P["itm2"][["item", "item_nombre"]], on="item", how="left")
        hojas.append(("Detalle Ppto", pd.DataFrame({
            "CECO": dp["ceco"], "Ítem": dp["item_nombre"], "Concepto Gasto": dp["concepto_gasto"],
            "Actividad": dp["actividad"], "Medida": dp["medida"].map(MEDIDA_LBL).fillna(dp["medida"]),
            "Valor (USD)": dp["valor_n"], "Valor (Aj. 2027)": dp["valor_a"],
        }), ["Valor (USD)", "Valor (Aj. 2027)"]))

    return hojas


def _formatear(ws, valcols, headers):
    head_font, head_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=TEAL)
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j)
        c.font, c.fill, c.alignment = head_font, head_fill, Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = min(max(len(str(h)) + 2, 11), 46)
    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    if valcols and ws.max_row > 1:
        idx = {h: j + 1 for j, h in enumerate(headers)}
        fmt = "#,##0.0" if any("FTE" in v for v in valcols) else "#,##0"
        for v in valcols:
            j = idx.get(v)
            if j:
                for row in ws.iter_rows(min_row=2, min_col=j, max_col=j, max_row=ws.max_row):
                    row[0].number_format = fmt


def _hoja_leeme(ws, hojas, completo, titulo):
    ws.sheet_view.showGridLines = False
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=15, color=TEAL)
    intro = ["",
             "Extracto de la base de datos del tablero, en un solo Excel.",
             "Cada hoja ya trae los NOMBRES (VP, Gerencia, Ítem, Desc. CLACO…): no hace falta cruzar nada.",
             "Valores en USD. «Aj. 2027» = misma cifra reexpresada a moneda equivalente 2027.",
             "Generado con bbdd_a_excel.py a partir de v2/bbdd/ (regenerable).", ""]
    r = 2
    for t in intro:
        ws.cell(row=r, column=1, value=t); r += 1
    for j, h in enumerate(("Hoja", "Qué contiene"), start=1):
        c = ws.cell(row=r, column=j, value=h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=TEAL)
    r += 1
    desc = {
        "Registros": "Real y Presupuesto histórico (2022–2025, 2026 YTD y 2026 FY) por CECO / Ítem / CLACO / Contrapartida.",
        "Forecast y Ppto 2027": "Forecast 5+7 2026 y Presupuesto 2027 (anual) por CECO / Ítem / CLACO.",
        "Dotaciones (FTE)": "Dotación (personas) Propios y Contratista por VP / Gerencia, año y período.",
        "Dicc. CECOs": "Catálogo de CECO: VP, Gerencia, Desc. CECO (estructura nueva/vigente).",
        "Dicc. Ítems": "Catálogo de Ítem: nombre, Tipo Costo, Ítem Relevante, Clasificación Cuenta.",
        "Dicc. CLACOs": "Catálogo de CLACO (Clase de Costo) → Desc. CLACO.",
        "Dicc. Compañías": "Catálogo de compañías.",
        "Detalle Real": "Gasto Real línea a línea (Contrapartida › Texto pedido › Denominación › Documento).",
        "Detalle Ppto": "Presupuesto/Forecast a nivel Concepto Gasto › Actividad.",
    }
    for name, _df, _v in hojas:
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=desc.get(name, "")); r += 1
    if not completo:
        ws.cell(row=r + 1, column=1,
                value="Nota: el detalle línea-a-línea no se incluyó (archivo liviano). Para agregarlo: "
                      "python bbdd_a_excel.py --completo")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 95


def escribir_excel(out, hojas, completo, titulo):
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df, valcols in hojas:
            df.to_excel(xw, sheet_name=name[:31], index=False)
            _formatear(xw.book[name[:31]], valcols, list(df.columns))
        _hoja_leeme(xw.book.create_sheet("Léame", 0), hojas, completo, titulo)
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    return os.path.getsize(out) / 1e6


def main():
    por_vp = "--por-vp" in sys.argv
    completo = "--completo" in sys.argv

    log("Leyendo BBDD…")
    dfs = cargar(con_detalle=(completo or por_vp))
    P = preparar(dfs)

    if not por_vp:
        out = OUT_FULL if completo else OUT_BASE
        hojas = construir_hojas(dfs, P, completo)
        mb = escribir_excel(out, hojas, completo, "BBDD del Dashboard Corporativo · Ppto 2027")
        log(f"\n✓ LISTO: {out} ({mb:.1f} MB · {len(hojas)} hojas)")
        return

    # --por-vp: cada CECO → su VP de la estructura nueva; un Excel COMPLETO por VP.
    cec_vp = P["cec_enrich"][["ceco", "vp"]].dropna()
    grupos = cec_vp.groupby("vp")["ceco"].apply(set)
    log(f"Generando Excel por VP (data completa) — {len(grupos)} VPs…")
    total = 0.0
    n = 0
    for vp, cecos in grupos.items():
        hojas = construir_hojas(dfs, P, completo=True, cecos=cecos)
        nreg = len(hojas[0][1])
        if nreg == 0:   # VP sin datos (p.ej. CECO nuevo sin movimientos) → no genera Excel vacío
            log(f"    {vp:44} (0 registros — se omite)")
            continue
        out = os.path.join(V3_POR_VP, _safe(vp), "BBDD " + _safe(vp) + ".xlsx")
        mb = escribir_excel(out, hojas, True, f"BBDD · {vp} · Ppto 2027")
        total += mb; n += 1
        log(f"    {vp:44} {len(cecos):3} CECOs · {nreg:6,} regs · {mb:4.1f} MB")
    log(f"\n✓ LISTO. {n} Excel por VP en {V3_POR_VP} ({total:.0f} MB en total)")


if __name__ == "__main__":
    main()
