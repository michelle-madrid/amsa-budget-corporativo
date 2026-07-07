# -*- coding: utf-8 -*-
"""
construir_v2.py — Pipeline ÚNICO del Dashboard Corporativo v2
=============================================================
Excel (v2/uploads) → base de datos parquet (regenerable) → data.js → embebido en
el HTML autocontenido. Un solo script, reproducible.

FUENTES (v2/uploads):
  · Reales Históricos.xlsx  → hoja "Data Consolidada"        (solo compañía 1000 = Corporativo)
  · Planes Históricos.xlsx  → hoja "Budget Consolidado Unpivot" (corp + distribuibles)
  · diccionario/CECOS.xlsx  → hojas "CECOS Corporativo Nueva" y "CECOS Corporativo"
        (CECO → VP / Gerencia / Tipo Costo / ¿Aplica? / Clasificación del Gasto)
        Se embeben AMBAS versiones (antigua y nueva) para poder elegir en el dashboard.
  · diccionario/CLACOS.xlsx → GP_CO2P / GP_CO1 (CLACO → Cód_Agrupación2 = Ítem Relevante)
  · diccionario/COMPAÑÍAS.xlsx → código compañía → nombre / abreviado / clasificación
  · Dotaciones Histórico AMSA.xlsx → igual que v1 (FTE Propios/Contratista)

VALORES (dos modos elegibles en el dashboard):
  · Real  — Normal = 'Val/Mon.so.CO' (USD, col C) · Ajustada dic-2027 = 'Valor dic-2027 (USD)' (col T)
  · Ppto  — Normal = 'Valor'                       · Ajustada dic-2027 = 'Valor USD dic-2027'

MAPEO A ÍTEM RELEVANTE:
  · Real  — por NOMBRE de clase de costo ('Descrip.clases coste', col D) → CLACO → Agrupación2 (cruza 100% del valor)
  · Ppto  — por CÓDIGO CLACO ('Clase de Costo') → Agrupación2

Uso:  python construir_v2.py
"""
import os
import json
import gzip
import base64
import re
import datetime
import collections
import openpyxl
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
UP = os.path.join(HERE, "uploads")
DIC = os.path.join(UP, "diccionario")
REAL_FILE = os.path.join(UP, "Reales Históricos v2.xlsx")
DIST_REAL_FILE = os.path.join(UP, "Reales Distribuibles Histórico v2.xlsx")
BUD_FILE = os.path.join(UP, "Planes Históricos v2.xlsx")
DOT_FILE = os.path.join(UP, "Dotaciones Histórico AMSA.xlsx")
CECOS_FILE = os.path.join(DIC, "CECOS.xlsx")
CLACOS_FILE = os.path.join(DIC, "CLACOS.xlsx")
COMP_FILE = os.path.join(DIC, "COMPAÑÍAS.xlsx")

PARQUET = os.path.join(HERE, "bbdd_v2.parquet")
DATA_JS = os.path.join(HERE, "data_v2.js")
HTML_PATH = os.path.join(HERE, "Dashboard Corporativo v2.html")

REAL_SHEET = "Data Consolidada"
BUD_SHEET = "Budget Consolidado Unpivot"
YEARS_HIST = [2022, 2023, 2024, 2025]
YTD_2026 = {"01", "02", "03", "04", "05"}   # 2026 YTD = ene–may (Real y Ppto, mismo período)

# Columnas (0-based) — cada fuente de Reales trae su propio mapeo (ceco,valn,desc,contra,anio,mes,vala)
COLS_CORP = {"ceco": 1, "valn": 2, "desc": 4, "contra": 6, "anio": 14, "mes": 15, "vala": 19}  # Reales Corp (23 cols)
COLS_DIST = {"ceco": 1, "valn": 2, "desc": 4, "contra": 6, "anio": 12, "mes": 13, "vala": 17}  # Reales Distribuibles (21 cols)
CONTRA_PPTO = "(Presupuesto)"        # los registros de Ppto no tienen contrapartida
B_ANIO, B_CECO, B_CLACO, B_MESANIO, B_VALN, B_VALA = 0, 1, 3, 6, 7, 8  # Ppto / Budget Unpivot


def log(m):
    try:
        print(m)
    except Exception:
        print(m.encode("ascii", "replace").decode())


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return 0.0


def _txt(v):
    return "" if v is None else str(v).strip()


def _round(v):
    f = round(float(v), 2)
    return int(f) if f == int(f) else f


# ===========================================================================
#  1) DICCIONARIOS
# ===========================================================================
def leer_clacos():
    """Devuelve (code2item, name2item, item2name, item2tc).
    Ítem Relevante = nivel FINO (REMUNERACI.EE…): en GP_CO1 es Cód_Agrupación4 (col7),
    en GP_CO2P es Cód_Agrupación2 (col2) — mismo esquema. El Tipo Costo (C1/C2/C3/
    TRASPASOS) = Cód_Agrupación2 de GP_CO1 (col3), que es el tipo de costo del CLACO."""
    wb = openpyxl.load_workbook(CLACOS_FILE, data_only=True, read_only=True)
    code2item, name2item, item2name, item2tc = {}, {}, {}, {}
    # GP_CO1: item = Cód_Ag4(7)/Ag4(8) · tc = Cód_Ag2(3) · CLACO=12 · Desc.CLACO=13
    it = wb["GP_CO1"].iter_rows(values_only=True); next(it)
    for r in it:
        item = _txt(r[7]) if 7 < len(r) else ""
        if not item:
            continue
        item2name.setdefault(item, (_txt(r[8]) if 8 < len(r) else "") or item)
        item2tc.setdefault(item, (_txt(r[3]) if 3 < len(r) else "") or None)   # C1/C2/C3/TRASPASOS
        if 12 < len(r) and r[12] is not None:
            code2item.setdefault(_txt(r[12]), item)
        if 13 < len(r) and r[13] is not None:
            name2item.setdefault(_txt(r[13]).lower(), item)
    # GP_CO2P: item = Cód_Ag2(2)/Ag2(3) · CLACO=6 · Desc.CLACO=7 (complementa nombres/códigos)
    it = wb["GP_CO2P"].iter_rows(values_only=True); next(it)
    for r in it:
        item = _txt(r[2]) if 2 < len(r) else ""
        if not item:
            continue
        item2name.setdefault(item, (_txt(r[3]) if 3 < len(r) else "") or item)
        if 6 < len(r) and r[6] is not None:
            code2item.setdefault(_txt(r[6]), item)
        if 7 < len(r) and r[7] is not None:
            name2item.setdefault(_txt(r[7]).lower(), item)
    wb.close()
    return code2item, name2item, item2name, item2tc


def leer_cecos(sheet):
    """{ceco: {vp, ger, tc, ap, clasif, comp}} desde una hoja de CECOS.xlsx."""
    wb = openpyxl.load_workbook(CECOS_FILE, data_only=True, read_only=True)
    ws = wb[sheet]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    h = grid[0]
    idx = {_txt(c): i for i, c in enumerate(h)}
    ci = idx.get("CECO")
    vpC = next((i for i, c in enumerate(h) if "Nodo 2" in _txt(c) and "Desc" in _txt(c)), None)
    gerC = next((i for i, c in enumerate(h) if "Nodo 3" in _txt(c) and "Desc" in _txt(c)), None)
    dcecoC = next((i for i, c in enumerate(h) if _txt(c) == "Desc. CECO"), None)
    clsC = next((i for i, c in enumerate(h) if "Clasificaci" in _txt(c) and "Gasto" in _txt(c)), None)
    tcC = next((i for i, c in enumerate(h) if "Tipo Costo" in _txt(c)), None)
    apC = next((i for i, c in enumerate(h) if "Aplica" in _txt(c)), None)
    out = {}
    for r in grid[1:]:
        code = _txt(r[ci]) if ci is not None and ci < len(r) else ""
        if not code:
            continue
        g = lambda i: (_txt(r[i]) if i is not None and i < len(r) else "")
        ap = "No" if g(apC) == "No" else ("Sí" if g(apC) in ("Sí", "Si", "") else g(apC))
        vp = g(vpC)
        nodo3 = g(gerC)
        # Gerencia: usa Nodo 3 cuando es significativo; si viene igual a la VP (caso
        # corporativo, Nodo 3 degenerado) usa la Desc. CECO (gerencia real del CECO).
        ger = nodo3 if (nodo3 and nodo3 != vp) else (g(dcecoC) or nodo3 or vp)
        out[code] = {"vp": vp or "(sin VP)", "ger": ger or "(sin Gerencia)",
                     "tc": g(tcC) or None, "ap": ap, "cl": g(clsC) or "(sin clasificación)",
                     "comp": code[:4]}
    return out


def leer_companias():
    """{código(str): {nombre, abrev, clasif}} desde COMPAÑÍAS.xlsx."""
    wb = openpyxl.load_workbook(COMP_FILE, data_only=True, read_only=True)
    it = wb["Compañías"].iter_rows(values_only=True)
    next(it)
    out = {}
    for r in it:
        if r[0] is None:
            continue
        out[_txt(r[0])] = {"nombre": _txt(r[2]), "abrev": _txt(r[3]), "clasif": _txt(r[1])}
    wb.close()
    return out


# ===========================================================================
#  2) LECTURA DE REALES Y PLANES → agregación por (ceco, item, bucket)
# ===========================================================================
# key = (ceco, item, contra, bucket); bucket ∈ {'2022'..'2025','2026ytd','2026fy'}; valores [rn, ra, pn, pa]
def _agg_new():
    return collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])


def _leer_real_fuente(path, cols, name2ag, agg, realcecos, sin_item):
    """Lee una hoja 'Data Consolidada' de Reales (corp o distribuibles) y acumula en agg (in-place)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    it = wb[REAL_SHEET].iter_rows(values_only=True)
    next(it)
    n = 0
    for r in it:
        n += 1
        ceco = _txt(r[cols["ceco"]])
        if not ceco:
            continue
        realcecos.add(ceco)
        anio = r[cols["anio"]]
        nm = _txt(r[cols["desc"]]).lower()
        item = name2ag.get(nm) or "(sin ítem)"
        if item == "(sin ítem)":
            sin_item[nm] += 1
        contra = _txt(r[cols["contra"]]) or "(sin contrapartida)"   # Denom.cuenta contrapartida (solo Real)
        vn = _num(r[cols["valn"]]); va = _num(r[cols["vala"]])
        if anio in YEARS_HIST:
            bucket = str(anio)
        elif anio == 2026:
            mes = _txt(r[cols["mes"]]).split(".")[-1]
            if mes not in YTD_2026:
                continue
            bucket = "2026ytd"
        else:
            continue
        cell = agg[(ceco, item, contra, bucket)]
        cell[0] += vn; cell[1] += va
    wb.close()
    return n


def leer_reales(name2ag):
    agg = _agg_new()
    realcecos = set()   # CECOs con Real = universo en alcance (Corp + Distribuibles)
    sin_item = collections.Counter()
    nc = _leer_real_fuente(REAL_FILE, COLS_CORP, name2ag, agg, realcecos, sin_item)
    cecos_corp = len(realcecos)
    nd = _leer_real_fuente(DIST_REAL_FILE, COLS_DIST, name2ag, agg, realcecos, sin_item)
    log(f"  Reales: Corp {nc} filas ({cecos_corp} CECOs) + Distribuibles {nd} filas ({len(realcecos) - cecos_corp} CECOs)")
    log(f"          {len(realcecos)} CECOs con Real · sin ítem: {sum(sin_item.values())} filas / {len(sin_item)} nombres")
    return agg, realcecos


def leer_planes(code2ag, agg):
    sin_item = collections.Counter()
    wb = openpyxl.load_workbook(BUD_FILE, data_only=True, read_only=True)
    it = wb[BUD_SHEET].iter_rows(values_only=True)
    next(it)
    n = 0
    for r in it:
        n += 1
        ceco = _txt(r[B_CECO])
        if not ceco:
            continue
        try:
            anio = int(_txt(r[B_ANIO]))
        except ValueError:
            continue
        claco = _txt(r[B_CLACO])
        item = code2ag.get(claco) or "(sin ítem)"
        if item == "(sin ítem)":
            sin_item[claco] += 1
        vn = _num(r[B_VALN]); va = _num(r[B_VALA])
        if anio in YEARS_HIST:
            buckets = [str(anio)]
        elif anio == 2026:
            mes = _txt(r[B_MESANIO]).split("-")[-1]
            buckets = ["2026fy"] + (["2026ytd"] if mes in YTD_2026 else [])
        else:
            continue
        for b in buckets:
            cell = agg[(ceco, item, CONTRA_PPTO, b)]   # el Ppto no trae contrapartida
            cell[2] += vn; cell[3] += va
    wb.close()
    log(f"  Planes: {n} filas · sin ítem (CLACO no cruza): {sum(sin_item.values())} filas / {len(sin_item)} CLACOs")
    return agg


# ===========================================================================
#  3) DOTACIONES (igual que v1)
# ===========================================================================
def leer_dotaciones():
    if not os.path.isfile(DOT_FILE):
        log("  AVISO: no encontré el Excel de Dotaciones; se omite.")
        return [], []
    ANUAL = {2022: 1, 2023: 3, 2024: 5, 2025: 7}
    SHEETS = {"Consolidado Propios": "propios", "Consolidado Contratista": "contratista"}
    wb = openpyxl.load_workbook(DOT_FILE, data_only=True, read_only=True)
    records, tidy = [], []
    for sheet, src in SHEETS.items():
        if sheet not in wb.sheetnames:
            continue
        grid = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
        vp = None
        for r in grid[5:]:
            c0 = r[0] if r else None
            if c0 in (None, ""):
                continue
            s = str(c0)
            if not (s.startswith(" ") or s.startswith("\t")):
                vp = s.strip(); continue
            ger = s.strip()
            g = lambda i: _num(r[i]) if i < len(r) else None
            rec = {"src": src, "vp": vp, "ger": ger}
            for y, col in ANUAL.items():
                rec[f"y{y}"] = {"real": g(col), "plan": g(col + 1)}
                tidy.append(dict(src=src, vp=vp, ger=ger, year=y, period="TOTAL", real=g(col), plan=g(col + 1)))
            rec["y2026"] = {"real": g(9), "plan": g(10)}
            rec["y2026fy"] = {"plan": g(11)}
            tidy.append(dict(src=src, vp=vp, ger=ger, year=2026, period="YTD", real=g(9), plan=g(10)))
            tidy.append(dict(src=src, vp=vp, ger=ger, year=2026, period="FY", real=None, plan=g(11)))
            records.append(rec)
    wb.close()
    log(f"  Dotaciones: {len(records)} registros.")
    return records, tidy


# ===========================================================================
#  4) ARMADO: parquet + registros + data.js
# ===========================================================================
def construir_registros(agg):
    """agg[(ceco,item,contra,bucket)] = [rn,ra,pn,pa]  →  registros por (ceco,item,contra)."""
    por_key = collections.defaultdict(dict)   # (ceco,item,contra) -> {bucket: [rn,ra,pn,pa]}
    for (ceco, item, contra, bucket), vals in agg.items():
        por_key[(ceco, item, contra)][bucket] = vals
    records = []
    for (ceco, item, contra), bmap in por_key.items():
        rec = {"ceco": ceco, "item": item, "contra": contra}
        def cell(b):
            return bmap.get(b, [0.0, 0.0, 0.0, 0.0])
        for y in YEARS_HIST:
            v = cell(str(y))
            rec[f"y{y}"] = {"real": {"n": _round(v[0]), "a": _round(v[1])},
                            "plan": {"n": _round(v[2]), "a": _round(v[3])}}
        v = cell("2026ytd")
        rec["y2026"] = {"real": {"n": _round(v[0]), "a": _round(v[1])},
                        "plan": {"n": _round(v[2]), "a": _round(v[3])}}
        vf = cell("2026fy")
        rec["y2026fy"] = {"plan": {"n": _round(vf[2]), "a": _round(vf[3])}}
        records.append(rec)
    return records


def escribir_parquet(agg):
    rows = []
    for (ceco, item, contra, bucket), v in agg.items():
        rows.append(dict(ceco=ceco, item=item, contra=contra, bucket=bucket,
                         real_n=v[0], real_a=v[1], plan_n=v[2], plan_a=v[3]))
    df = pd.DataFrame(rows, columns=["ceco", "item", "contra", "bucket", "real_n", "real_a", "plan_n", "plan_a"])
    df.to_parquet(PARQUET, index=False)
    log(f"  Parquet: {PARQUET} ({len(df)} filas)")
    return df


def construir_data_js(records, itemNames, itemTc, cecoNew, cecoOld, comps, dot_records):
    j = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    items = sorted({r["item"] for r in records})
    v2 = {"records": records, "items": items, "itemNames": itemNames, "itemTc": itemTc,
          "cecoNew": cecoNew, "cecoOld": cecoOld, "comps": comps,
          "years": YEARS_HIST + [2026]}
    return ("window.V2_DATA = " + j(v2) + ";\n" +
            "window.DOT_DATA = " + j({"records": dot_records}) + ";")


# ===========================================================================
#  5) EMBEBER EN EL HTML AUTOCONTENIDO (código src/*.jsx + model.js + data.js)
# ===========================================================================
# UUIDs de los assets dentro del HTML (heredados de v1; estables). El código vive
# en v2/src/. Se valida el JSX con el Babel embebido antes de reemplazar.
DATA_UUID = "4220c217-9166-414f-8f44-a89ae6c16113"
BABEL_UUID = "bc7af47c-01a9-47f6-b802-09d216d56f10"
SRC = os.path.join(HERE, "src")
ASSETS = {   # uuid → (archivo en src/, es_jsx)
    "b10e2126-5e07-4358-9daa-f1805ba1bc96": ("app.jsx", True),
    "1a5bd7f6-3963-4301-955f-e692cf651ee7": ("charts.jsx", True),
    "0290c010-dd92-483c-ac2e-8b4fe5d69eab": ("filters.jsx", True),
    "a1f4c70d-e147-4337-a85e-7a5b341a43c9": ("matrix.jsx", True),
    "35b49f4c-a3fb-4834-8681-376cfb28922f": ("tweaks-panel.jsx", True),
    "5c895c10-43c7-4f6c-aa15-ea39fb941007": ("model.js", False),
}


def _asset_rx(uuid):
    return re.compile(r'("' + re.escape(uuid) + r'":\{"mime":"[^"]*","compressed":true,"data":")([^"]*)(")')


def _get_asset(html, uuid):
    m = _asset_rx(uuid).search(html)
    return gzip.decompress(base64.b64decode(m.group(2))).decode("utf-8") if m else None


def _set_asset(html, uuid, texto):
    b64 = base64.b64encode(gzip.compress(texto.encode("utf-8"), mtime=0)).decode("ascii")
    nuevo, n = _asset_rx(uuid).subn(lambda m: m.group(1) + b64 + m.group(3), html)
    if n != 1:
        raise SystemExit(f"Esperaba 1 asset {uuid}, encontré {n}.")
    return nuevo


def _validar_jsx(html, archivos):
    """Transpila cada JSX con el Babel embebido (Node); aborta si hay error."""
    import subprocess
    import tempfile
    babel = _get_asset(html, BABEL_UUID)
    if not babel:
        log("  AVISO: no encontré Babel embebido; se omite validación JSX.")
        return
    with tempfile.TemporaryDirectory() as td:
        bpath = os.path.join(td, "babel.js")
        open(bpath, "w", encoding="utf-8").write(babel)
        for name in archivos:
            src = open(os.path.join(SRC, name), encoding="utf-8").read()
            spath = os.path.join(td, "src.txt")
            open(spath, "w", encoding="utf-8").write(src)
            chk = ("const B=require(%r);const fs=require('fs');const c=fs.readFileSync(%r,'utf8');"
                   "try{B.transform(c,{presets:['react']});console.log('OK');}"
                   "catch(e){console.error(String(e));process.exit(1);}" % (bpath, spath))
            r = subprocess.run(["node", "-e", chk], capture_output=True, text=True)
            if r.returncode != 0:
                raise SystemExit(f"JSX inválido en {name}:\n{r.stderr or r.stdout}")
            log(f"    Babel OK: {name}")


def embeber(data_js):
    """Re-embebe código (src/*) + data.js en el HTML autocontenido, con respaldo + zip."""
    if not os.path.isfile(HTML_PATH):
        log("  AVISO: no existe el HTML v2; se omite embebido (solo data_v2.js).")
        return
    import shutil
    import zipfile
    html = open(HTML_PATH, "r", encoding="utf-8", newline="").read()
    # La barra de filtros debe ENVOLVER (no recortarse). El CSS base venía nowrap.
    html = html.replace("gap:10px;flex-wrap:nowrap;align-items:flex-end;",
                        "gap:10px;flex-wrap:wrap;align-items:flex-end;")
    # Indentación de los niveles 4 (CECO) y 5 (Contrapartida) — detalle máximo.
    if ".ind-4{" not in html:
        html = html.replace(".ind-3{padding-left:44px;}",
                            ".ind-3{padding-left:44px;}.ind-4{padding-left:66px;}.ind-5{padding-left:88px;}")
    elif ".ind-5{" not in html:
        html = html.replace(".ind-4{padding-left:66px;}",
                            ".ind-4{padding-left:66px;}.ind-5{padding-left:88px;}")
    log("  Validando JSX con Babel…")
    _validar_jsx(html, [n for (n, isjsx) in ASSETS.values() if isjsx])
    for uuid, (name, _isjsx) in ASSETS.items():
        html = _set_asset(html, uuid, open(os.path.join(SRC, name), encoding="utf-8").read())
    html = _set_asset(html, DATA_UUID, data_js)
    shutil.copyfile(HTML_PATH, HTML_PATH + ".bak")
    open(HTML_PATH, "w", encoding="utf-8", newline="").write(html)
    zip_path = os.path.join(HERE, "Dashboard Corporativo v2.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(HTML_PATH, arcname=os.path.basename(HTML_PATH))
        z.writestr("LEER - Como abrir.txt", "Doble clic en el HTML (Chrome/Edge/Firefox).")
    log("  HTML v2 actualizado (código + data embebidos) · ZIP generado.")


# ===========================================================================
def main():
    log("=== Construir Dashboard Corporativo v2 ===")
    log("Leyendo diccionarios…")
    code2item, name2item, item2name, item2tc = leer_clacos()
    cecoNew = leer_cecos("CECOS Corporativo Nueva")
    cecoOld = leer_cecos("CECOS Corporativo")
    comps = leer_companias()
    log(f"  CLACO: {len(code2item)} códigos · {len(name2item)} nombres · {len(item2name)} ítems (Agrupación4)")
    log(f"  CECOS: {len(cecoNew)} (nueva) · {len(cecoOld)} (antigua) · Compañías: {len(comps)}")

    log("Leyendo Reales…")
    agg, realcecos = leer_reales(name2item)
    log("Leyendo Planes…")
    agg = leer_planes(code2item, agg)

    # ALCANCE Actividad Corporativa: en el lado corporativo (compañías "Actividad
    # Corporativa": 1000/1200/0000) se dejan SOLO las CECO que tienen Real; las demás
    # (Capex, Exploraciones, Oficina Londres/China, Proyectos…) traen Ppto sin Real y
    # se excluyen. Los distribuibles (otras compañías) se mantienen (Ppto).
    def _en_alcance(ceco):
        clasif = (comps.get(ceco[:4]) or {}).get("clasif", "")
        if "Distribuib" in clasif:
            return True
        return ceco in realcecos
    antes = len(agg)
    agg = {k: v for k, v in agg.items() if _en_alcance(k[0])}
    log(f"  Alcance corporativo = CECOs con Real: {antes - len(agg)} combos fuera de alcance excluidos.")

    df = escribir_parquet(agg)
    records = construir_registros(agg)

    # Nombres y Tipo Costo (C1/C2/C3) de los ítems presentes.
    itemNames, itemTc = {}, {}
    for r in records:
        it = r["item"]
        itemNames[it] = item2name.get(it, it) if it != "(sin ítem)" else "(sin ítem)"
        tc = item2tc.get(it)
        if tc:
            itemTc[it] = tc

    log("Leyendo Dotaciones…")
    dot_records, dot_tidy = leer_dotaciones()
    if dot_tidy:
        pd.DataFrame(dot_tidy, columns=["src", "vp", "ger", "year", "period", "real", "plan"]) \
            .to_parquet(os.path.join(HERE, "dotaciones_v2.parquet"), index=False)

    data_js = construir_data_js(records, itemNames, itemTc, cecoNew, cecoOld, comps, dot_records)
    open(DATA_JS, "w", encoding="utf-8", newline="").write(data_js)
    log(f"  data.js: {DATA_JS} ({len(records)} registros ceco×ítem)")

    # Validación rápida (totales en MM USD)
    _validar(df, cecoNew, comps)

    embeber(data_js)
    log("\n✓ LISTO.")


def _validar(df, cecomap, comps):
    log("  --- Totales (MM USD) ---")
    for b in ["2022", "2023", "2024", "2025", "2026ytd", "2026fy"]:
        sub = df[df.bucket == b]
        log(f"    {b:8}: Real N {sub.real_n.sum()/1e6:8.2f} / A {sub.real_a.sum()/1e6:8.2f}"
            f"  ·  Ppto N {sub.plan_n.sum()/1e6:8.2f} / A {sub.plan_a.sum()/1e6:8.2f}")


if __name__ == "__main__":
    main()
