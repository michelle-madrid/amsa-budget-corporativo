# -*- coding: utf-8 -*-
"""
bbdd_a_excel.py — Convierte la BBDD (bbdd/*.parquet) en Excel para compartir.
============================================================================
Pensado para usuarios NO técnicos: DOS hojas, línea a línea, con los NOMBRES ya pegados
(VP, Gerencia, Desc. CECO, Ítem…), sin cruzar tablas:
  · «Reales»                  → gasto Real (2022–2025 y 2026 YTD).
  · «Presupuestos y Forecast» → Ppto 2025/2026, Forecast 5+7 2026 y Ppto 2027.
Más una hoja "Léame" de ayuda. Formato amigable (filtros, encabezados, freeze).

Requiere haber corrido antes `python construir.py` (genera bbdd/).
Uso:
  python bbdd_a_excel.py --gerencia "Data y Analitica"  → Excel de ESA Gerencia en salida/
  python bbdd_a_excel.py --vp "Comercializacion"        → Excel de ESA VP en salida/
  python bbdd_a_excel.py --por-vp                        → 1 Excel por CADA VP en
                                                           salida/por_vp/<VP>/BBDD <VP>.xlsx
  python bbdd_a_excel.py                                 → 1 Excel global (todas juntas)

El nombre de la Gerencia/VP se busca sin distinguir acentos ni mayúsculas y admite texto parcial
(p.ej. "analitica" encuentra "Data y Analítica Avanzada"). El archivo se llama «BBDD <Nombre>.xlsx».

El Excel usa SOLO la estructura NUEVA (la vigente): cada CECO se asigna a su VP de la
estructura nueva. (Los pocos CECOs que solo existían en la estructura antigua no se incluyen.)
"""
import os
import re
import sys
import unicodedata
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
BBDD = os.path.join(HERE, "bbdd")
SALIDA_DIR = os.path.join(os.path.dirname(HERE), "salida")
SALIDA_POR_VP = os.path.join(SALIDA_DIR, "por_vp")
OUT_BASE = os.path.join(BBDD, "BBDD Dashboard Corporativo.xlsx")

TEAL = "14515A"
BUCKET_LBL = {"2022": "2022", "2023": "2023", "2024": "2024", "2025": "2025",
              "2026ytd": "2026 YTD (ene–may)", "2026fy": "2026 FY (Ppto anual)"}
MEDIDA_LBL = {"ppto_2025": "Ppto 2025", "ppto_2026ytd": "Ppto 2026 YTD (ene–may)",
              "ppto_2026fy": "Ppto 2026 FY (anual)", "forecast_2026": "Forecast 5+7 2026",
              "ppto_2027": "Ppto 2027", "out66": "Outlook 6+6 2026"}


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
        raise SystemExit(f"No encontré {p}. Corré primero: python construir.py")
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


def construir_hojas(dfs, P, cecos=None):
    """Dos hojas, línea a línea, con los nombres ya pegados:
      · «Reales»                  → gasto Real (2022–2025 y 2026 YTD), línea a línea.
      · «Presupuestos y Forecast» → Ppto 2025/2026, Forecast 5+7 2026 y Ppto 2027, línea a línea.
    Si cecos ≠ None, filtra a esos CECOs (modo por VP / por Gerencia)."""
    def filt(df):
        return df if cecos is None else df[df["ceco"].isin(cecos)]

    # 1) Reales — gasto Real desglosado (Contrapartida › Texto pedido › Denominación › Documento).
    dr = _pegar(filt(dfs["fact_detalle_real"]), P)
    dr["Período"] = dr["bucket"].map(BUCKET_LBL).fillna(dr["bucket"])
    reales = pd.DataFrame({
        "VP": dr["vp"], "Gerencia": dr["gerencia"], "Desc. CECO": dr["desc_ceco"], "CECO": dr["ceco"],
        "Ítem Relevante": dr["itemrel_nombre"], "Ítem": dr["item_nombre"], "Tipo Costo": dr["tipo_costo"],
        "Contrapartida": dr["contra"], "Texto pedido": dr["texto_pedido"], "Denominación": dr["denominacion"],
        "Documento": dr["documento"], "Período": dr["Período"],
        "Real (USD)": dr["valor_n"], "Real (Aj. 2027)": dr["valor_a"],
    })

    # 2) Presupuestos y Forecast — Ppto/Forecast desglosado (Concepto Gasto › Actividad).
    dp = _pegar(filt(dfs["fact_detalle_ppto"]), P)
    dp["Medida"] = dp["medida"].map(MEDIDA_LBL).fillna(dp["medida"])
    ppto = pd.DataFrame({
        "VP": dp["vp"], "Gerencia": dp["gerencia"], "Desc. CECO": dp["desc_ceco"], "CECO": dp["ceco"],
        "Ítem Relevante": dp["itemrel_nombre"], "Ítem": dp["item_nombre"], "Tipo Costo": dp["tipo_costo"],
        "Concepto Gasto": dp["concepto_gasto"], "Actividad": dp["actividad"], "Medida": dp["Medida"],
        "Valor (USD)": dp["valor_n"], "Valor (Aj. 2027)": dp["valor_a"],
    })

    return [
        ("Reales", reales, ["Real (USD)", "Real (Aj. 2027)"]),
        ("Presupuestos y Forecast", ppto, ["Valor (USD)", "Valor (Aj. 2027)"]),
    ]


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


def escribir_excel(out, hojas):
    """Escribe las hojas (Reales + Presupuestos y Forecast). Sin hoja Léame."""
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df, valcols in hojas:
            df.to_excel(xw, sheet_name=name[:31], index=False)
            _formatear(xw.book[name[:31]], valcols, list(df.columns))
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    return os.path.getsize(out) / 1e6


# ── Generar el Excel de UNA Gerencia o VP puntual ──────────────────────────────────────────
def _norm(s):
    """Normaliza para comparar: sin acentos, minúsculas, sin espacios sobrantes."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def _resolver(P, texto, por):
    """Encuentra la Gerencia/VP que coincide con `texto` (sin distinguir acentos/mayúsculas).
    `por` ∈ {'gerencia', 'vp'}. Coincidencia exacta primero; si no, por substring."""
    col = "gerencia" if por == "gerencia" else "vp"
    opciones = sorted(P["cec_enrich"][col].dropna().astype(str).unique())
    q = _norm(texto)
    cand = [o for o in opciones if _norm(o) == q] or [o for o in opciones if q in _norm(o)]
    if not cand:
        raise SystemExit(f'No encontré ninguna {por} que coincida con "{texto}".\n'
                         f'  Opciones: ' + " | ".join(opciones))
    if len(cand) > 1:
        raise SystemExit(f'"{texto}" coincide con varias: ' + " | ".join(cand) + "\n  Afiná el texto.")
    return cand[0]


def generar(texto, por="gerencia", out_dir=None, dfs=None, P=None):
    """Genera el Excel (Reales + Presupuestos y Forecast) de una Gerencia o VP y lo guarda en salida/.
    Uso desde código:  bbdd_a_excel.generar("Data y Analitica")            → por Gerencia
                       bbdd_a_excel.generar("Comercializacion", por="vp")  → por VP
    Devuelve la ruta del archivo generado."""
    if dfs is None:
        log("Leyendo BBDD…")
        dfs = cargar(con_detalle=True)
    if P is None:
        P = preparar(dfs)
    nombre = _resolver(P, texto, por)
    col = "gerencia" if por == "gerencia" else "vp"
    cecos = set(P["cec_enrich"].loc[P["cec_enrich"][col] == nombre, "ceco"])
    hojas = construir_hojas(dfs, P, cecos=cecos)
    out = os.path.join(out_dir or SALIDA_DIR, "BBDD " + _safe(nombre) + ".xlsx")
    mb = escribir_excel(out, hojas)
    log(f"\nLISTO ({por}): {nombre}")
    log(f"  {out}")
    log(f"  {mb:.2f} MB · {len(cecos)} CECO(s) · " +
        " · ".join(f"{name} {len(df):,}" for name, df, _v in hojas))
    return out


def main():
    args = sys.argv[1:]

    # Excel de UNA Gerencia o VP:  python bbdd_a_excel.py --gerencia "Data y Analitica"
    for flag, por in (("--gerencia", "gerencia"), ("--vp", "vp")):
        if flag in args:
            i = args.index(flag)
            texto = args[i + 1] if i + 1 < len(args) else None
            if not texto:
                raise SystemExit(f'Uso: python bbdd_a_excel.py {flag} "Nombre"')
            generar(texto, por=por)
            return

    por_vp = "--por-vp" in args
    log("Leyendo BBDD…")
    dfs = cargar(con_detalle=True)   # las 2 hojas son siempre línea a línea
    P = preparar(dfs)

    if not por_vp:
        hojas = construir_hojas(dfs, P)
        mb = escribir_excel(OUT_BASE, hojas)
        log(f"\nLISTO: {OUT_BASE} ({mb:.1f} MB · {len(hojas)} hojas)")
        return

    # --por-vp: cada CECO → su VP de la estructura nueva; un Excel por VP (Reales + Ppto/Forecast).
    cec_vp = P["cec_enrich"][["ceco", "vp"]].dropna()
    grupos = cec_vp.groupby("vp")["ceco"].apply(set)
    log(f"Generando Excel por VP — {len(grupos)} VPs…")
    total = 0.0
    n = 0
    for vp, cecos in grupos.items():
        hojas = construir_hojas(dfs, P, cecos=cecos)
        nfilas = sum(len(df) for _n, df, _v in hojas)
        if nfilas == 0:   # VP sin datos (p.ej. CECO nuevo sin movimientos) → no genera Excel vacío
            log(f"    {vp:44} (0 filas — se omite)")
            continue
        out = os.path.join(SALIDA_POR_VP, _safe(vp), "BBDD " + _safe(vp) + ".xlsx")
        mb = escribir_excel(out, hojas)
        total += mb; n += 1
        log(f"    {vp:44} {len(cecos):3} CECOs · {nfilas:7,} filas · {mb:4.1f} MB")
    log(f"\nLISTO. {n} Excel por VP en {SALIDA_POR_VP} ({total:.0f} MB en total)")


if __name__ == "__main__":
    main()
