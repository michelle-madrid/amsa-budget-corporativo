# -*- coding: utf-8 -*-
"""
construir.py — Genera el Dashboard Corporativo de punta a punta.
===============================================================
Un solo comando lee los Excel, arma la base de datos y produce el HTML final,
autocontenido (un único archivo que abre offline en Chrome/Edge, sin instalar nada).

    python construir.py            → regenera todo (Excel → datos → dashboard)
    python construir.py --rapido   → solo rearma el dashboard, sin reescribir la
                                     plantilla ni los intermedios (para previsualizar rápido)

QUÉ PRODUCE:
  · salida/Dashboard Corporativo.html (+ .zip) — el dashboard que se comparte (~7 MB). Bajo
        cada Contrapartida cuelga el detalle del gasto: Texto pedido › Denominación (+ Documento).
  · bbdd/*.parquet — la misma data en tablas (esquema estrella) para Power BI / DuckDB / pandas.
  · plantilla.html, data_v2.js, bbdd_v2.parquet, dotaciones_v2.parquet — intermedios regenerables.

DE DÓNDE SALEN LOS DATOS (carpeta uploads/ y uploads/diccionario/):
  · Reales Históricos v2.xlsx             → "Data Consolidada" (corporativo, compañía 1000)
  · Reales Distribuibles Histórico v2.xlsx→ "Data Consolidada" (distribuibles 1001/1002/1003/1005)
        OJO: las columnas van en otro orden que el corporativo (ver COLS_CORP / COLS_DIST).
  · Planes Históricos v2.xlsx             → "Budget Consolidado Unpivot" (Presupuesto histórico)
  · EXPORT_FORECAST_5+7 2026.xlsx         → "Forecast 5+7 Unpivot" (forecast anual 2026)
  · diccionario/CECOS.xlsx    → "CECOS Corporativo Nueva" y "CECOS Corporativo" (se embeben las
        dos; el usuario elige con el toggle "Estructura CECOS"). CECO → VP / Gerencia / Tipo Costo.
  · diccionario/CLACOS.xlsx   → jerarquía del CLACO (ver leer_clacos()).
  · diccionario/COMPAÑÍAS.xlsx→ código de compañía → nombre / abreviado / clasificación.
  · Dotaciones Histórico AMSA.xlsx → dotación (FTE) Propios / Contratista.
  · assets/logo_amsa.png      → se embebe como data: URI; el HTML no necesita la carpeta assets/.

LA PLANTILLA: plantilla.html es el molde del que se parte (trae el bootstrap que hace que el
  HTML abra offline). No se regenera desde cero: embeber() la actualiza en sitio dejando un .bak.
  Si se pierde, se restaura desde git o desde el .bak.

VALORES (toggle "Base Moneda"): Normal (USD) vs Ajustada 2027. Cada celda guarda ambos {n, a}.
  · Real     — Normal 'Val/Mon.so.CO' · Ajustada 'Valor mes 2027 (USD)'
  · Ppto     — Normal 'Valor'         · Ajustada 'Valor USD 2027 (mes equivalente)'
  · Forecast — Normal 'Valor'         · Ajustada 'Valor USD 2027 (mes equivalente)'

CÓMO SE ARMA LA JERARQUÍA (a partir del CLACO, ver leer_clacos):
  Tipo Costo (Agrupación2) › Ítem Relevante (Agrupación3) › Ítem (Agrupación4).
  · Real          — cruza por NOMBRE ('Descrip.clases coste') → CLACO → Ítem.
  · Ppto/Forecast — cruzan por CÓDIGO ('Clase de Costo' / 'Clase Costo') → Ítem.

ALCANCE: aparecen TODOS los CECOs, con o sin Real (ver _en_alcance() en main()).

────────────────────────────────────────────────────────────────────────────
DÓNDE TOCAR PARA LOS CAMBIOS MÁS FRECUENTES (todo centralizado arriba del archivo):
  · Cambió el nombre del Excel de Ppto 2027 ...... PPTO27_FILE
  · Cambió el nombre del Forecast / Outlook ...... FCST_FILE / OUT_FILE
  · Se corrieron las columnas de los Reales ...... COLS_CORP / COLS_DIST
  · Columnas del detalle del gasto ............... DET_COLS_CORP / DET_COLS_DIST (en leer_detalle)
  · Columnas del Forecast ........................ FC_CECO / FC_CLACO / FC_VALN / FC_VALA
  · Años y períodos .............................. YEARS_HIST, YTD_HASTA_MES, BUCKET_IDX
  · Renombrar un CLACO dentro de un CECO ......... CLACO_REMAP
  · Forzar todo un CECO a un único CLACO ......... CLACO_FORCE_CECO
  · Comercialización a mostrar (f. «Otros Fletes») SHOW_COMERCIAL ← {CECO: [CLACO, ...]}
  · Reclasificar Clasificación Cuenta ............ CLAS_CUENTA_OVERRIDE
────────────────────────────────────────────────────────────────────────────
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
FCST_FILE = os.path.join(UP, "ejercicios_2026", "Forecast 5+7 2026.XLSX")
DOT_FILE = os.path.join(UP, "Dotaciones Histórico AMSA.xlsx")
CECOS_FILE = os.path.join(DIC, "CECOS.xlsx")
CLACOS_FILE = os.path.join(DIC, "CLACOS.xlsx")
COMP_FILE = os.path.join(DIC, "COMPAÑÍAS.xlsx")

PARQUET = os.path.join(HERE, "bbdd_v2.parquet")     # intermedio regenerable (BBDD agregada)
DATA_JS = os.path.join(HERE, "data_v2.js")          # intermedio regenerable (datos serializados a JS)
# plantilla.html: base HTML del bundler. embeber() la parchea en sitio (deja .bak). NO se comparte.
HTML_PATH = os.path.join(HERE, "plantilla.html")
# salida/: el dashboard terminado que se comparte. Es lo único que la usuaria abre/entrega.
SALIDA_DIR = os.path.join(os.path.dirname(HERE), "salida")
SALIDA_HTML = os.path.join(SALIDA_DIR, "Dashboard Corporativo.html")
BBDD_DIR = os.path.join(HERE, "bbdd")   # BBDD completa (esquema estrella, parquet) — espejo regenerable de la data

REAL_SHEET = "Data Consolidada"
BUD_SHEET = "Budget Consolidado Unpivot"
FCST_SHEET = "Forecast 5+7 Unpivot"
# Forecast 5+7 2026 (hoja Unpivot): CECO=2 · Clase Costo(CLACO)=4 · Valor=15 · Valor USD 2027=16
FC_CECO, FC_CLACO, FC_VALN, FC_VALA = 2, 4, 15, 16
FC_CG, FC_ACT = 8, 9   # Forecast: Concepto Gasto (col8) · Actividad (col9)
# Outlook 6+6 2026: otro ejercicio del mismo tipo. Su hoja Unpivot (generada con
# unpivot_ejercicio.py) tiene EXACTAMENTE el mismo layout que la del Forecast,
# así que reutiliza los índices FC_*. Se muestra como una columna más de la
# categoría "Forecast" (no reemplaza al 5+7).
OUT_FILE = os.path.join(UP, "ejercicios_2026", "Outlook 6+6 2026.XLSX")
OUT_SHEET = "Outlook 6+6 Unpivot"
# Reversas (ajustes al Forecast 5+7): filas extra que SOLO se aplican con la Estructura CECOS
# «Nueva con ajustes» (cecoMode='ajustes'). Mismo layout Unpivot que el Forecast → reusan FC_*.
# El resto del tiempo el modelo las oculta. Suelen ser redistribuciones (neto ≈ 0).
# AUTO-DESCUBRIMIENTO: cada .xlsx dejado en esta carpeta que tenga una hoja «… Unpivot» se suma
# automáticamente (no hay que editar código). Genera la hoja con: python unpivot_ejercicio.py.
REV_DIR = os.path.join(UP, "ajustes_forecast")
PPTO27_FILE = os.path.join(UP, "presupuesto_2027", "PPTO27_03_08_2026_09_15.XLSX")
PPTO27_SHEET = "Sheet1"
# CAPEX 2027 (hoja "BD AMSA"): base para la pestaña CAPEX. Header en 2 filas, datos desde la 3.
CAPEX_FILE = os.path.join(UP, "capex", "Presupuesto CAPEX AMSA 2027 v2.xlsx")
CAPEX_SHEET = "BD AMSA"
# Ppto 2027 (hoja ancha): CECO=2 · Clase Costo(CLACO)=4 · Total-2027=26 (USD, ya en moneda 2027).
P27_CECO, P27_CLACO, P27_TOTAL = 2, 4, 26
P27_CG, P27_ACT = 8, 9   # Ppto 2027: Concepto Gasto (col8) · Actividad (col9)


def _ppto_ts():
    """Fecha/hora de actualización del Ppto, sacada del NOMBRE del archivo
    (PPTO27_DD_MM_YYYY_HH_MM). Devuelve 'DD/MM/YYYY HH:MM' o '' si no matchea."""
    m = re.search(r"(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})", os.path.basename(PPTO27_FILE))
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)} {m.group(4)}:{m.group(5)}" if m else ""
YEARS_HIST = [2022, 2023, 2024, 2025]
# Último mes incluido en el YTD 2026 (1=ene … 12=dic). Corta el Real y el Ppto YTD en el MISMO mes.
# Editable desde el panel (server.py → «hasta qué mes»). Cambia esto y reconstruí para mover el corte.
YTD_HASTA_MES = 5
YTD_2026 = {f"{m:02d}" for m in range(1, YTD_HASTA_MES + 1)}   # 2026 YTD = ene..YTD_HASTA_MES (Real y Ppto, mismo período)
# Un mes de Real 2026 se considera CERRADO (y por tanto seleccionable como tope del YTD en el
# dashboard) solo si su gasto neto llega al menos a esta fracción del mes más alto cargado. Los
# meses "abiertos" traen únicamente el reverso de las provisiones del mes anterior (Provisión
# Cuentas por Pagar EE RR/PAs) → neto ~0 o negativo, y quedan fuera hasta que entre el gasto real.
REAL26_MES_FRAC = 0.30

# Columnas (0-based) — cada fuente de Reales trae su propio mapeo (ceco,valn,desc,contra,anio,mes,vala)
COLS_CORP = {"ceco": 1, "valn": 2, "desc": 4, "contra": 6, "anio": 14, "mes": 15, "vala": 19}  # Reales Corp (23 cols)
COLS_DIST = {"ceco": 1, "valn": 2, "desc": 4, "contra": 6, "anio": 12, "mes": 13, "vala": 17}  # Reales Distribuibles (21 cols)
CONTRA_PPTO = "(Presupuesto)"        # los registros de Ppto no tienen contrapartida
# Ppto / Budget Unpivot (el archivo ahora trae Concepto Gasto=5 y Actividad=6 → Mes/Valor corridos)
B_ANIO, B_CECO, B_CLACO, B_CG, B_ACT, B_MESANIO, B_VALN, B_VALA = 0, 1, 3, 5, 6, 8, 9, 10

# --- Services & Tech --------------------------------------------------------
# CLACOs 6125020 (Technical) y 6125021 (Service). Caen en el mismo Ítem SERV_CORP que el
# 6125015, así que NO se aíslan por Ítem. Se identifican:
#   · Ppto / Forecast / Ppto 2027 → por CÓDIGO (Clase de Costo = CLACO).
#   · Real (cruza por nombre)      → por el nombre de AGRUPACIÓN5 (col10 del CLACO), que es
#     el que trae "Descrip.clases coste": 6125020='Servicios Corporativos HH',
#     6125021='Servicios Corp. de comercialización' (≠ 6125015='Servicios Corporativos').
# Cada registro afectado se marca con el flag st=1 (parte de la clave de agregación) para
# poder filtrarlo sin cambiar la jerarquía de Ítem. ST_NAMES se arma en leer_clacos().
ST_CODES = {"6125020", "6125021"}

# --- Renombrado fijo de CLACO dentro de un CECO (regla de negocio) -----------
# {(ceco, claco_origen): claco_destino}. Se aplica SIEMPRE, en TODAS las fuentes (Real, Ppto,
# Forecast, Outlook, Reversas, Ppto2027 y su detalle), tras leer (ceco, claco) y ANTES de cruzar
# a Ítem y de guardar el CLACO. El ítem sigue al CLACO destino (en el Real se recalcula por código).
# Ej.: en 1002AD4503 la Liquidación 8200001 (Ítem Secundarias) pertenece a 6124407 (Servicios);
# al renombrarla, ambos quedan bajo 6124407 y se netean.
CLACO_REMAP = {
    ("1002AD4503", "8200001"): "6124407",
}

# --- Forzar el CLACO de un CECO entero (regla de negocio) --------------------
# {CECO: CLACO_destino}. Para ese CECO, TODO su gasto pasa a ese único CLACO (sin importar el CLACO
# de origen), en TODAS las fuentes. Se aplica ANTES de cruzar a Ítem, así que el Ítem también sigue
# al CLACO destino. Tiene prioridad sobre CLACO_REMAP.
# Ej.: 0000CHINA (Oficina China) consolida todo bajo 6000000 ("Oficina China", Ítem Otros Gastos);
# así deja de caer en 6125020 (Services & Tech), que el filtro "Sin S&T" ocultaba.
CLACO_FORCE_CECO = {
    "0000CHINA": "6000000",
}

# --- Clasificación Cuenta: reasignar por Ítem (regla de negocio) --------------
# {Ítem (Cód_Agrupación4): "Clasificación Cuenta"}. Pisa la Clasificación Cuenta que trae
# CLACOS.xlsx (col 16) para ESE ítem, es decir para TODOS los CLACOs de ese código. Sirve para
# crear/mover categorías del filtro "Clasificación Cuenta" (la categoría nueva aparece sola en el
# filtro). Ej.: todos los CLACOs del Ítem SEGUROS pasan a la categoría "Seguros".
CLAS_CUENTA_OVERRIDE = {
    "SEGUROS": "Seguros",
}

# --- Comercialización: combinaciones (CECO, CLACO) que SÍ se muestran (regla de negocio) ---
# Whitelist del filtro «Otros Fletes»: dentro de la VP Comercialización, SOLO estas combinaciones
# (CECO → [CLACO, ...]) se muestran; el resto de la VP Comercialización se OCULTA cuando el filtro
# está activo. Para agregar/quitar: editá este diccionario (CECO: lista de CLACOs a mostrar).
SHOW_COMERCIAL = {
    "1000AC6001": ["6124001", "6124004", "6124103", "6124484", "6124601", "6124605", "6124607",
                   "6124610", "6124711", "6125002", "6125003", "6125011", "6125007", "6124602",
                   "6125006", "6125009"],
    "1001AD6000": ["6124202", "6125003", "6124605", "6124706", "6125002", "6125005", "6124710",
                   "6124704", "6124702", "6124712", "6124709", "6124703", "6124701"],
    "1002AD6000": ["6124202", "6125003", "6124605", "6124706", "6125002", "6125005", "6124710",
                   "6124704", "6124702", "6124712", "6124709", "6124703", "6124701"],
    "1003AD6000": ["6124202", "6125003", "6124706", "6125002", "6125005", "6124710", "6124704",
                   "6124702", "6124712", "6124709", "6124703", "6124701"],
    "1005AD6001": ["6124202", "6125003", "6124706", "6125002", "6125005", "6124710", "6124704",
                   "6124702", "6124712", "6124709", "6124703", "6124701"],
}


def _remap_claco(ceco, claco):
    if ceco in CLACO_FORCE_CECO:          # el CECO entero se consolida en un único CLACO
        return CLACO_FORCE_CECO[ceco]
    return CLACO_REMAP.get((ceco, claco), claco)


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
    """Devuelve (code2item, name2item, item2name, item2tc, item2rel, rel2name, item2clas, st_names).
    Jerarquía CLACO (GP_CO1): Cód_Ag2 (C1/C2/C3 = Tipo Costo) › Cód_Ag3 (Ítem Relevante,
    ej. REMUNERACI/'Remuneraciones') › Cód_Ag4 (Ítem/detalle, ej. REMUNERACI.EE).
    · Ítem (detalle) = Cód_Agrupación4 (col7) / nombre col8.
    · Ítem Relevante = Cód_Agrupación3 (col5) / nombre col6  → item2rel (Ag4→Ag3), rel2name.
    · Tipo Costo (C1/C2/C3/TRASPASOS) = Cód_Agrupación2 de GP_CO1 (col3).
    · Clasificación Cuenta (col16, ej. 'Total Opex' / 'Mano de Obra') → item2clas (1:1 por ítem)."""
    wb = openpyxl.load_workbook(CLACOS_FILE, data_only=True, read_only=True)
    code2item, name2item, item2name, item2tc = {}, {}, {}, {}
    item2rel, rel2name, item2clas = {}, {}, {}
    st_names = set()   # nombres (Agrupación5, col10) de los CLACOs Services & Tech → para cruzar el Real
    name2claco = {}    # nombre (Desc.CLACO / Agrupación5) → código CLACO, para cruzar el Real por código
    claco2name = {}    # código CLACO → Desc. CLACO (col13), para mostrar la descripción en el front
    # GP_CO1: item=Ag4(7)/nombre(8) · rel=Ag3(5)/nombre(6) · tc=Ag2(3) · Ag5=9/10 · CLACO=12 · Desc.CLACO=13 · clasCuenta=16
    it = wb["GP_CO1"].iter_rows(values_only=True); next(it)
    for r in it:
        claco = _txt(r[12]) if 12 < len(r) else ""
        if claco in ST_CODES and 10 < len(r) and _txt(r[10]):
            st_names.add(_txt(r[10]).lower())   # p.ej. 'servicios corporativos hh'
        if claco:
            if 13 < len(r) and _txt(r[13]):
                name2claco.setdefault(_txt(r[13]).lower(), claco)   # Desc.CLACO → código
                claco2name.setdefault(claco, _txt(r[13]))           # código → Desc.CLACO
            if 10 < len(r) and _txt(r[10]):
                name2claco.setdefault(_txt(r[10]).lower(), claco)   # Agrupación5 → código (nombre que usa el Real)
        item = _txt(r[7]) if 7 < len(r) else ""
        if not item:
            continue
        item2name.setdefault(item, (_txt(r[8]) if 8 < len(r) else "") or item)
        item2tc.setdefault(item, (_txt(r[3]) if 3 < len(r) else "") or None)   # C1/C2/C3/TRASPASOS
        item2clas.setdefault(item, (_txt(r[16]) if 16 < len(r) else "") or None)  # Clasificación Cuenta
        rel = _txt(r[5]) if 5 < len(r) else ""                                  # Cód_Agrupación3
        if rel:
            item2rel.setdefault(item, rel)
            rel2name.setdefault(rel, (_txt(r[6]) if 6 < len(r) else "") or rel)
        if claco:
            code2item.setdefault(claco, item)
        if 13 < len(r) and r[13] is not None:
            name2item.setdefault(_txt(r[13]).lower(), item)
    # GP_CO2P: item = Cód_Ag2(2)/Ag2(3) · CLACO=6 · Desc.CLACO=7 (complementa nombres/códigos)
    it = wb["GP_CO2P"].iter_rows(values_only=True); next(it)
    for r in it:
        item = _txt(r[2]) if 2 < len(r) else ""
        if not item:
            continue
        item2name.setdefault(item, (_txt(r[3]) if 3 < len(r) else "") or item)
        claco = _txt(r[6]) if 6 < len(r) else ""
        if claco:
            code2item.setdefault(claco, item)
        if 7 < len(r) and r[7] is not None:
            name2item.setdefault(_txt(r[7]).lower(), item)
            if claco:
                name2claco.setdefault(_txt(r[7]).lower(), claco)   # Desc (GP_CO2P) → código
                claco2name.setdefault(claco, _txt(r[7]))           # código → Desc (complementa GP_CO1)
    wb.close()
    # Reasignaciones manuales de Clasificación Cuenta (ver CLAS_CUENTA_OVERRIDE arriba).
    for it, cl in CLAS_CUENTA_OVERRIDE.items():
        item2clas[it] = cl
    return code2item, name2item, item2name, item2tc, item2rel, rel2name, item2clas, st_names, name2claco, claco2name


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
                     "dceco": g(dcecoC) or code,   # Desc. CECO (nombre del CECO) → nivel bajo Gerencia
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


def _leer_real_fuente(path, cols, name2ag, agg, realcecos, sin_item, name2claco, code2item):
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
        claco = name2claco.get(nm) or "(sin claco)"   # Clase de Costo por NOMBRE (Real cruza por nombre)
        rc = _remap_claco(ceco, claco)                # regla fija de renombrado de CLACO por CECO
        if rc != claco:
            claco = rc
            item = code2item.get(claco, item)         # el ítem sigue al CLACO destino (por código)
        contra = _txt(r[cols["contra"]]) or "(sin contrapartida)"   # Denom.cuenta contrapartida (solo Real)
        vn = _num(r[cols["valn"]]); va = _num(r[cols["vala"]])
        if anio in YEARS_HIST:
            cell = agg[(ceco, item, contra, str(anio), claco)]
            cell[0] += vn; cell[1] += va
        elif anio == 2026:
            mes = _txt(r[cols["mes"]]).split(".")[-1].zfill(2)
            if not (mes.isdigit() and 1 <= int(mes) <= 12):
                continue
            # Bucket MENSUAL (todos los meses de Real que existan) → corte YTD dinámico en el front.
            cm = agg[(ceco, item, contra, "2026m" + mes, claco)]
            cm[0] += vn; cm[1] += va
            # Bucket YTD del build (corte por defecto YTD_HASTA_MES) — intacto para parquet/back-compat.
            if mes in YTD_2026:
                cy = agg[(ceco, item, contra, "2026ytd", claco)]
                cy[0] += vn; cy[1] += va
    wb.close()
    return n


def leer_reales(name2ag, name2claco, code2item):
    agg = _agg_new()
    realcecos = set()   # CECOs con Real = universo en alcance (Corp + Distribuibles)
    sin_item = collections.Counter()
    nc = _leer_real_fuente(REAL_FILE, COLS_CORP, name2ag, agg, realcecos, sin_item, name2claco, code2item)
    cecos_corp = len(realcecos)
    nd = _leer_real_fuente(DIST_REAL_FILE, COLS_DIST, name2ag, agg, realcecos, sin_item, name2claco, code2item)
    log(f"  Reales: Corp {nc} filas ({cecos_corp} CECOs) + Distribuibles {nd} filas ({len(realcecos) - cecos_corp} CECOs)")
    log(f"          {len(realcecos)} CECOs con Real · sin ítem: {sum(sin_item.values())} filas / {len(sin_item)} nombres")
    return agg, realcecos


# ===========================================================================
#  DETALLE DEL GASTO: niveles línea-a-línea BAJO Contrapartida.
#  Jerarquía: … › Ítem › Contrapartida › Texto pedido › Denominación (+ Documento).
#  Para EDITAR qué columnas forman el detalle en el futuro, cambiar aquí los índices
#  (0-based) por fuente. 'texto'/'denom'/'doc' son los campos del detalle; el resto
#  ubica cada línea (ceco, ítem vía 'desc', contrapartida) y su valor (n/a) y bucket.
# ===========================================================================
DET_COLS_CORP = {"ceco": 1, "desc": 4, "contra": 6, "texto": 3, "denom": 8, "doc": 9, "valn": 2, "vala": 19, "anio": 14, "mes": 15}
DET_COLS_DIST = {"ceco": 1, "desc": 4, "contra": 6, "texto": 3, "denom": 8, "doc": 9, "valn": 2, "vala": 17, "anio": 12, "mes": 13}
BUCKET_IDX = {y: i for i, y in enumerate(YEARS_HIST)}   # 2022→0 … 2025→3 · 2026ytd→4


def leer_detalle(name2ag, en_alcance, st_names, name2claco, code2item):
    """Detalle línea-a-línea de los Reales (corp+dist). Devuelve un blob compacto
    {s:[strings únicos], k:{'item\\x01contra': [[cecoIdx,textoIdx,denomIdx,docIdx,bk,n,a,st], …]}}.
    Agrega por (item,contra,ceco,texto,denom,doc,bucket,st) e interna strings para pesar poco.
    st=1 = Services & Tech (nombre de Agrupación5 en st_names) — para filtrarlo en el detalle."""
    agg = {}

    def proc(path, C):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        it = wb[REAL_SHEET].iter_rows(values_only=True); next(it)
        for r in it:
            ceco = _txt(r[C["ceco"]])
            if not ceco or not en_alcance(ceco):
                continue
            anio = r[C["anio"]]
            if anio in BUCKET_IDX:
                bk = BUCKET_IDX[anio]
            elif anio == 2026:
                mes = _txt(r[C["mes"]]).split(".")[-1]
                if mes not in YTD_2026:
                    continue
                bk = len(YEARS_HIST)   # 2026ytd
            else:
                continue
            nm = _txt(r[C["desc"]]).lower()
            item = name2ag.get(nm) or "(sin ítem)"
            _cl = name2claco.get(nm)                       # regla fija de renombrado de CLACO por CECO
            if _cl and _remap_claco(ceco, _cl) != _cl:     # si el CLACO se renombra, el ítem lo sigue
                item = code2item.get(_remap_claco(ceco, _cl), item)
            st = 1 if nm in st_names else 0
            contra = _txt(r[C["contra"]]) or "(sin contrapartida)"
            texto = _txt(r[C["texto"]]) or "(sin texto de pedido)"
            denom = _txt(r[C["denom"]]) or "(sin denominación)"
            doc = _txt(r[C["doc"]]) or "—"
            vn = _num(r[C["valn"]]); va = _num(r[C["vala"]])
            k = (item, contra, ceco, texto, denom, doc, bk, st)
            cell = agg.get(k)
            if cell:
                cell[0] += vn; cell[1] += va
            else:
                agg[k] = [vn, va]
        wb.close()

    proc(REAL_FILE, DET_COLS_CORP)
    proc(DIST_REAL_FILE, DET_COLS_DIST)
    S, sidx = [], {}
    def I(x):
        j = sidx.get(x)
        if j is None:
            j = len(S); sidx[x] = j; S.append(x)
        return j
    K = {}
    for (item, contra, ceco, texto, denom, doc, bk, st), (vn, va) in agg.items():
        K.setdefault(item + "\x01" + contra, []).append(
            [I(ceco), I(texto), I(denom), I(doc), bk, _round(vn), _round(va), st])
    log(f"  Detalle Real: {len(agg)} líneas · {len(S)} strings · {len(K)} claves ítem×contra")
    return {"s": S, "k": K}


# Medidas del detalle Ppto/Forecast (Concepto Gasto › Actividad). Cada línea lleva su medida:
#   0 Ppto 2025 · 1 Ppto 2026 YTD · 2 Ppto 2026 FY · 3 Forecast 5+7 2026 · 4 Ppto 2027
#   5 Outlook 6+6 2026
def leer_detalle_pf(code2ag, en_alcance):
    """Detalle Concepto Gasto › Actividad para Ppto (Budget 2025-2026), Forecast y Ppto 2027.
    (El Ppto histórico 2022-2024 no trae Concepto/Actividad, así que no entra.) Cruza por CÓDIGO.
    Devuelve blob {s:[strings], k:{'item': [[cecoIdx, cgIdx, actIdx, medida, n, a], …]}}."""
    agg = {}   # (item, ceco, cg, act, medida) -> [n, a]

    def add(item, ceco, cg, act, meas, n, a):
        if n == 0 and a == 0:
            return
        k = (item, ceco, cg, act, meas); c = agg.get(k)
        if c:
            c[0] += n; c[1] += a
        else:
            agg[k] = [n, a]

    # Ppto (Budget) 2025 y 2026 (los únicos años con Concepto/Actividad)
    wb = openpyxl.load_workbook(BUD_FILE, data_only=True, read_only=True)
    it = wb[BUD_SHEET].iter_rows(values_only=True); next(it)
    for r in it:
        ceco = _txt(r[B_CECO])
        if not ceco or not en_alcance(ceco):
            continue
        try:
            anio = int(_txt(r[B_ANIO]))
        except ValueError:
            continue
        if anio not in (2025, 2026):
            continue
        item = code2ag.get(_remap_claco(ceco, _txt(r[B_CLACO]))) or "(sin ítem)"
        cg = _txt(r[B_CG]) or "(sin concepto)"; act = _txt(r[B_ACT]) or "(sin actividad)"
        vn = _num(r[B_VALN]); va = _num(r[B_VALA])
        if anio == 2025:
            add(item, ceco, cg, act, 0, vn, va)
        else:
            add(item, ceco, cg, act, 2, vn, va)                      # FY
            if _txt(r[B_MESANIO]).split("-")[-1] in YTD_2026:
                add(item, ceco, cg, act, 1, vn, va)                  # YTD
    wb.close()
    # Forecast 5+7 2026
    if os.path.isfile(FCST_FILE):
        wb = openpyxl.load_workbook(FCST_FILE, data_only=True, read_only=True)
        if FCST_SHEET in wb.sheetnames:                    # sin la hoja Unpivot: se omite (no revienta el build)
            it = wb[FCST_SHEET].iter_rows(values_only=True); next(it)
            for r in it:
                ceco = _txt(r[FC_CECO])
                if not ceco or not en_alcance(ceco):
                    continue
                item = code2ag.get(_remap_claco(ceco, _txt(r[FC_CLACO]))) or "(sin ítem)"
                cg = _txt(r[FC_CG]) or "(sin concepto)"; act = _txt(r[FC_ACT]) or "(sin actividad)"
                add(item, ceco, cg, act, 3, _num(r[FC_VALN]), _num(r[FC_VALA]))
        wb.close()
    # Outlook 6+6 2026 (medida 5) — mismo layout de hoja que el Forecast
    if os.path.isfile(OUT_FILE):
        wb = openpyxl.load_workbook(OUT_FILE, data_only=True, read_only=True)
        if OUT_SHEET in wb.sheetnames:
            it = wb[OUT_SHEET].iter_rows(values_only=True); next(it)
            for r in it:
                ceco = _txt(r[FC_CECO])
                if not ceco or not en_alcance(ceco):
                    continue
                item = code2ag.get(_remap_claco(ceco, _txt(r[FC_CLACO]))) or "(sin ítem)"
                cg = _txt(r[FC_CG]) or "(sin concepto)"; act = _txt(r[FC_ACT]) or "(sin actividad)"
                add(item, ceco, cg, act, 5, _num(r[FC_VALN]), _num(r[FC_VALA]))
        wb.close()
    # Ppto 2027 (valor tal cual; original = ajustada)
    if os.path.isfile(PPTO27_FILE):
        wb = openpyxl.load_workbook(PPTO27_FILE, data_only=True, read_only=True)
        it = wb[PPTO27_SHEET].iter_rows(values_only=True); next(it)
        for r in it:
            ceco = _txt(r[P27_CECO])
            if not ceco or not en_alcance(ceco):
                continue
            item = code2ag.get(_remap_claco(ceco, _txt(r[P27_CLACO]))) or "(sin ítem)"
            cg = _txt(r[P27_CG]) or "(sin concepto)"; act = _txt(r[P27_ACT]) or "(sin actividad)"
            v = _num(r[P27_TOTAL]); add(item, ceco, cg, act, 4, v, v)
        wb.close()
    S, sidx = [], {}
    def I(x):
        j = sidx.get(x)
        if j is None:
            j = len(S); sidx[x] = j; S.append(x)
        return j
    K = {}
    for (item, ceco, cg, act, meas), (n, a) in agg.items():
        K.setdefault(item, []).append([I(ceco), I(cg), I(act), meas, _round(n), _round(a)])
    log(f"  Detalle Ppto/Fcst: {len(agg)} líneas · {len(S)} strings · {len(K)} ítems")
    return {"s": S, "k": K}


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
        claco = _remap_claco(ceco, _txt(r[B_CLACO]))   # regla fija de renombrado de CLACO por CECO
        item = code2ag.get(claco) or "(sin ítem)"
        if item == "(sin ítem)":
            sin_item[claco] += 1
        clacok = claco or "(sin claco)"
        vn = _num(r[B_VALN]); va = _num(r[B_VALA])
        if anio in YEARS_HIST:
            buckets = [str(anio)]
        elif anio == 2026:
            mes = _txt(r[B_MESANIO]).split("-")[-1].zfill(2)
            buckets = ["2026fy"]                                # Ppto FY = todos los meses
            if mes.isdigit() and 1 <= int(mes) <= 12:
                buckets.append("2026m" + mes)                  # bucket mensual → corte YTD dinámico
            if mes in YTD_2026:
                buckets.append("2026ytd")                      # YTD del build (corte por defecto)
        else:
            continue
        for b in buckets:
            cell = agg[(ceco, item, CONTRA_PPTO, b, clacok)]   # el Ppto no trae contrapartida; claco = Clase de Costo
            cell[2] += vn; cell[3] += va
    wb.close()
    log(f"  Planes: {n} filas · sin ítem (CLACO no cruza): {sum(sin_item.values())} filas / {len(sin_item)} CLACOs")
    return agg


def leer_forecast(code2ag, en_alcance, path=None, sheet=None, etiqueta="Forecast 5+7"):
    """Ejercicio anual (Forecast 5+7 / Outlook 6+6) agregado por (ceco, item, claco) con valor
    Normal y Ajustado 2027. Ambos usan la misma hoja Unpivot y los mismos índices FC_*.
    Cruza CLACO (Clase Costo) → Ítem por CÓDIGO. Devuelve {(ceco,item,claco): [n, a]}."""
    path = path or FCST_FILE
    sheet = sheet or FCST_SHEET
    if not os.path.isfile(path):
        log(f"  AVISO: no está el archivo de {etiqueta}; se omite.")
        return {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        log(f"  AVISO: {os.path.basename(path)} no tiene la hoja «{sheet}»; se omite {etiqueta}.")
        wb.close()
        return {}
    it = wb[sheet].iter_rows(values_only=True); next(it)
    agg = {}; n = 0; sin_item = collections.Counter()
    for r in it:
        ceco = _txt(r[FC_CECO])
        if not ceco or not en_alcance(ceco):
            continue
        claco = _remap_claco(ceco, _txt(r[FC_CLACO]))   # regla fija de renombrado de CLACO por CECO
        item = code2ag.get(claco) or "(sin ítem)"
        if item == "(sin ítem)":
            sin_item[claco] += 1
        vn = _num(r[FC_VALN]); va = _num(r[FC_VALA])
        if vn == 0 and va == 0:
            continue
        k = (ceco, item, claco or "(sin claco)"); c = agg.get(k)
        if c:
            c[0] += vn; c[1] += va
        else:
            agg[k] = [vn, va]
        n += 1
    wb.close()
    log(f"  {etiqueta}: {n} filas → {len(agg)} (ceco,item) · sin ítem (CLACO no cruza): {sum(sin_item.values())} filas / {len(sin_item)} CLACOs")
    return agg


def leer_ppto27(code2ag, en_alcance):
    """Ppto 2027 (anual) por (ceco, item, claco), tomando el Total-2027 tal cual (USD ya en moneda
    2027; original = ajustada). Cruza CLACO → Ítem por CÓDIGO. Devuelve {(ceco,item,claco): valor}."""
    if not os.path.isfile(PPTO27_FILE):
        log("  AVISO: no está el archivo de Ppto 2027; se omite.")
        return {}
    wb = openpyxl.load_workbook(PPTO27_FILE, data_only=True, read_only=True)
    it = wb[PPTO27_SHEET].iter_rows(values_only=True); next(it)
    agg = {}; n = 0; sin_item = collections.Counter()
    for r in it:
        ceco = _txt(r[P27_CECO])
        if not ceco or not en_alcance(ceco):
            continue
        claco = _remap_claco(ceco, _txt(r[P27_CLACO]))   # regla fija de renombrado de CLACO por CECO
        item = code2ag.get(claco) or "(sin ítem)"
        if item == "(sin ítem)":
            sin_item[claco] += 1
        v = _num(r[P27_TOTAL])
        if v == 0:
            continue
        k = (ceco, item, claco or "(sin claco)")
        agg[k] = agg.get(k, 0.0) + v
        n += 1
    wb.close()
    log(f"  Ppto 2027: {n} filas → {len(agg)} (ceco,item) · sin ítem (CLACO no cruza): {sum(sin_item.values())} filas / {len(sin_item)} CLACOs")
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
    """agg[(ceco,item,contra,bucket,claco)] = [rn,ra,pn,pa]  →  registros por (ceco,item,contra,claco)."""
    por_key = collections.defaultdict(dict)   # (ceco,item,contra,claco) -> {bucket: [rn,ra,pn,pa]}
    for (ceco, item, contra, bucket, claco), vals in agg.items():
        por_key[(ceco, item, contra, claco)][bucket] = vals
    records = []
    for (ceco, item, contra, claco), bmap in por_key.items():
        rec = {"ceco": ceco, "item": item, "contra": contra}
        if claco and claco != "(sin claco)":
            rec["claco"] = claco               # Clase de Costo (para el filtro por código)
        if claco in ST_CODES:
            rec["st"] = 1                      # Services & Tech (derivado del CLACO)
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
        # Meses 2026 (sparse) para el corte YTD dinámico en el front:
        #   m26 = {"r": {"MM": [n,a] Real}, "p": {"MM": [n,a] Ppto}}  (solo meses con dato ≠ 0)
        mr, mp = {}, {}
        for _mm in range(1, 13):
            mv = bmap.get("2026m%02d" % _mm)
            if not mv:
                continue
            k = "%02d" % _mm
            if mv[0] or mv[1]:
                mr[k] = [_round(mv[0]), _round(mv[1])]
            if mv[2] or mv[3]:
                mp[k] = [_round(mv[2]), _round(mv[3])]
        if mr or mp:
            rec["m26"] = {"r": mr, "p": mp}
        records.append(rec)
    return records


def escribir_parquet(agg, write=True):
    rows = []
    for (ceco, item, contra, bucket, claco), v in agg.items():
        rows.append(dict(ceco=ceco, item=item, contra=contra, bucket=bucket, claco=claco,
                         real_n=v[0], real_a=v[1], plan_n=v[2], plan_a=v[3]))
    df = pd.DataFrame(rows, columns=["ceco", "item", "contra", "bucket", "claco", "real_n", "real_a", "plan_n", "plan_a"])
    if write:   # en modo rápido no se toca el parquet intermedio
        df.to_parquet(PARQUET, index=False)
        log(f"  Parquet: {PARQUET} ({len(df)} filas)")
    return df


def _logo_data_uri():
    """Logo AMSA como data: URI (para que el HTML sea autocontenido, sin assets/ externo)."""
    logo = os.path.join(HERE, "assets", "logo_amsa.png")
    if not os.path.isfile(logo):
        return None
    b64 = base64.b64encode(open(logo, "rb").read()).decode("ascii")
    return "data:image/png;base64," + b64


def _ultimo_mes_cerrado(records):
    """Último mes de Real 2026 CERRADO = tope del selector YTD del dashboard.

    Suma el Real neto (Normal) por mes sobre todos los registros y devuelve el mes más alto cuyo
    total llega al menos a REAL26_MES_FRAC del mes más alto cargado. Los meses "abiertos" (solo el
    reverso de provisiones, sin gasto real todavía) tienen neto ~0 o negativo y quedan excluidos,
    para que el YTD no baje al acumular un mes incompleto. Sin datos → 0.
    """
    tot = collections.defaultdict(float)
    for r in records:
        for m, v in (r.get("m26", {}).get("r") or {}).items():
            tot[int(m)] += (v[0] or 0.0)
    if not tot:
        return 0
    umbral = max(tot.values()) * REAL26_MES_FRAC
    cerrados = [m for m, s in tot.items() if s >= umbral]
    return max(cerrados) if cerrados else 0


def construir_data_js(records, itemNames, itemTc, cecoNew, cecoOld, comps, dot_records, itemRel=None, relNames=None, itemClas=None, clacoNames=None):
    j = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    items = sorted({r["item"] for r in records})
    # Solo los nombres de CLACO presentes en los registros (evita cargar el diccionario completo).
    clacosUsados = {r["claco"] for r in records if r.get("claco")}
    clacoNames = {c: n for c, n in (clacoNames or {}).items() if c in clacosUsados}
    v2 = {"records": records, "items": items, "itemNames": itemNames, "itemTc": itemTc,
          "itemRel": itemRel or {}, "relNames": relNames or {}, "itemClas": itemClas or {},
          "clacoNames": clacoNames,
          "cecoNew": cecoNew, "cecoOld": cecoOld, "comps": comps,
          # Whitelist del filtro «Otros Fletes»: combinaciones (CECO, CLACO) de la VP Comercialización
          # que SÍ se muestran (el resto de la VP se oculta cuando el filtro está activo).
          "showCombos": [[c, k] for c, ks in SHOW_COMERCIAL.items() for k in ks],
          "comercialVp": "VP Comercialización",
          # Corte YTD 2026: mes por defecto (YTD_HASTA_MES) y máximo mes con Real cargado (para el
          # selector en vivo del dashboard). El front recalcula el YTD Real/Ppto con rec.m26.
          "ytdHastaMes": YTD_HASTA_MES,
          "real26MesMax": _ultimo_mes_cerrado(records),
          "pptoTs": _ppto_ts(),   # fecha/hora de actualización del Ppto (del nombre del archivo)
          "years": YEARS_HIST + [2026]}
    out = ("window.V2_DATA = " + j(v2) + ";\n" +
           "window.DOT_DATA = " + j({"records": dot_records}) + ";")
    logo = _logo_data_uri()
    if logo:
        out += "\nwindow.CORP_LOGO = " + j(logo) + ";"
    return out


# ===========================================================================
#  5) EMBEBER EN EL HTML AUTOCONTENIDO (código src/*.jsx + model.js + data.js)
# ===========================================================================
# UUIDs de los assets dentro del HTML (identifican cada bloque de código; son estables).
# El código vive en src/. Se valida el JSX con el Babel embebido antes de reemplazar.
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


# Bootstrap "todo incrustado": incrusta los scripts INLINE (sin blob: ni fetch) y las
# fuentes/imágenes como data: URI. Así el HTML abre desde file:// aunque esté descargado
# (los contextos que bloquean blob: —vista previa, archivo con "marca de Internet"— ya
# no rompen la carga). Reemplaza el bootstrap original (idempotente por el marcador).
NEW_BOOTSTRAP = r"""
document.addEventListener('DOMContentLoaded', async function() {
  /*INLINE-BOOTSTRAP-v3*/
  const loading = document.getElementById('__bundler_loading');
  function setStatus(msg) { if (loading) loading.textContent = msg; }
  window.addEventListener('error', function(e) {
    var t = e && e.target;
    // Solo un SCRIPT crítico faltante muestra el aviso; una imagen/fuente que falle
    // NUNCA debe tapar el panel (así el archivo funciona 100% autocontenido).
    var isResource = t && t !== window && t.tagName === 'SCRIPT';
    if (!isResource) { try { console.error('[bundle]', e.message || e.type); } catch (x) {} return; }
    if (window.__bundler_hint_shown) return;
    window.__bundler_hint_shown = true;
    var p = document.body || document.documentElement;
    var d = p.appendChild(document.createElement('div'));
    d.id = '__bundler_hint';
    d.style.cssText = 'position:fixed;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;text-align:center;padding:24px;background:#14515a;color:#fff;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;z-index:999999';
    d.innerHTML =
      '<div style="font-size:24px;font-weight:800;letter-spacing:.2px">Actividad Corporativa + Distribuibles Ppto 2027</div>' +
      '<div style="font-size:15px;line-height:1.6;max-width:560px;opacity:.95">No se pudo mostrar este panel interactivo.<br>Ábrelo con <b>Chrome</b> o <b>Edge</b>.</div>';
  }, true);
  try {
    window.__resources = {};
    const manifestEl = document.querySelector('script[type="__bundler/manifest"]');
    const templateEl = document.querySelector('script[type="__bundler/template"]');
    if (!manifestEl || !templateEl) { setStatus('Error: missing bundle data'); return; }
    const manifest = JSON.parse(manifestEl.textContent);
    let template = JSON.parse(templateEl.textContent);
    const uuids = Object.keys(manifest);
    setStatus('Unpacking ' + uuids.length + ' assets...');
    const scriptText = {};   // uuid -> codigo JS/JSX (se incrusta inline)
    const dataUri = {};      // uuid -> data: URI (fuentes / imagenes)
    function toB64(bytes) { var s = ''; for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]); return btoa(s); }
    await Promise.all(uuids.map(async (uuid) => {
      const entry = manifest[uuid];
      try {
        const binaryStr = atob(entry.data);
        const bytes = new Uint8Array(binaryStr.length);
        for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
        let finalBytes = bytes;
        if (entry.compressed && typeof DecompressionStream !== 'undefined') {
          const ds = new DecompressionStream('gzip');
          const writer = ds.writable.getWriter();
          const reader = ds.readable.getReader();
          writer.write(bytes); writer.close();
          const chunks = []; let totalLen = 0;
          while (true) { const rr = await reader.read(); if (rr.done) break; chunks.push(rr.value); totalLen += rr.value.length; }
          finalBytes = new Uint8Array(totalLen); let off = 0;
          for (const c of chunks) { finalBytes.set(c, off); off += c.length; }
        }
        if (/(javascript|ecmascript|jsx)/i.test(entry.mime)) {
          scriptText[uuid] = new TextDecoder('utf-8').decode(finalBytes);
        } else {
          const b64 = entry.compressed ? toB64(finalBytes) : entry.data;
          dataUri[uuid] = 'data:' + entry.mime + ';base64,' + b64;
        }
      } catch (err) { console.error('Failed to decode asset ' + uuid + ':', err); }
    }));
    // Recursos (fuentes/imagenes) -> data: URI directamente en el template.
    for (const uuid in dataUri) template = template.split(uuid).join(dataUri[uuid]);
    template = template.replace(/\s+integrity="[^"]*"/gi, '').replace(/\s+crossorigin="[^"]*"/gi, '');
    setStatus('Rendering...');
    const doc = new DOMParser().parseFromString(template, 'text/html');
    document.documentElement.replaceWith(doc.documentElement);
    // Scripts: los que apuntan a un asset se INCRUSTAN inline (sin blob, sin fetch),
    // ejecutando en orden (React -> ReactDOM -> Babel -> text/babel).
    const dead = Array.from(document.scripts);
    for (const old of dead) {
      const s = document.createElement('script');
      for (const a of old.attributes) s.setAttribute(a.name, a.value);
      s.textContent = old.textContent;
      const ref = s.getAttribute('src');
      if (ref && scriptText[ref]) { s.textContent = scriptText[ref]; s.removeAttribute('src'); }
      old.replaceWith(s);
    }
    if (window.Babel && typeof window.Babel.transformScriptTags === 'function') window.Babel.transformScriptTags();
  } catch (err) { setStatus('Error unpacking: ' + err.message); console.error('Bundle unpack error:', err); }
});
"""


# Pantalla estática (sin JS) que se ve por defecto y el bootstrap borra al montar.
# En navegador normal es la pantalla de carga (flash breve); en la vista previa de
# SharePoint/Teams —que bloquea el JS— QUEDA visible con el aviso, en vez de una caja
# en blanco. Estilos INLINE (los sanitizadores de preview suelen borrar <style>).
NEW_THUMB = (
    '<div id="__bundler_thumbnail" data-fallback-msg="2" '
    'style="position:fixed;inset:0;display:flex;flex-direction:column;align-items:center;'
    'justify-content:center;gap:18px;text-align:center;padding:32px;background:#14515a;'
    'color:#fff;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;z-index:9999">'
    '<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" style="width:84px;height:84px">'
    '<rect width="100" height="100" rx="14" fill="#0e3f47"></rect>'
    '<g transform="translate(50 46)" fill="none" stroke="#ffffff" stroke-width="5" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M-22 14 L-7 -14 L7 8 L22 -18"></path></g>'
    '<rect x="28" y="70" width="44" height="6" rx="3" fill="#f0a929"></rect>'
    '<rect x="28" y="80" width="28" height="6" rx="3" fill="#e63b2e"></rect></svg>'
    '<div style="font-size:24px;font-weight:800;letter-spacing:.3px">Actividad Corporativa + Distribuibles Ppto 2027</div>'
    '<div style="font-size:15px;line-height:1.65;max-width:620px;opacity:.95">'
    'Cargando panel interactivo…<br>'
    'Si no aparece (por ejemplo, en la vista previa de SharePoint), <b>descárgalo</b> y ábrelo con un navegador (<b>Chrome</b> o <b>Edge</b>).'
    '</div></div>'
)


def _parchar_thumbnail(html):
    """Reemplaza el placeholder de carga por la pantalla estática con aviso (idempotente)."""
    if 'data-fallback-msg="2"' in html:
        return html
    i = html.find('<div id="__bundler_thumbnail"')
    if i < 0:
        raise SystemExit("No encontré #__bundler_thumbnail para parchar.")
    # Reemplaza TODO el bloque del thumbnail (aunque tenga <div> internos) hasta el
    # toast de carga que siempre le sigue — robusto ante re-parcheos.
    j = html.find('<div id="__bundler_loading"', i)
    if j < 0:
        raise SystemExit("No encontré #__bundler_loading tras el thumbnail.")
    return html[:i] + NEW_THUMB + "\n  " + html[j:]


def _parchar_bootstrap(html):
    """Reemplaza el bootstrap del HTML por la versión 'todo incrustado' (idempotente)."""
    if "/*INLINE-BOOTSTRAP-v3*/" in html:
        return html
    anchor = "document.addEventListener('DOMContentLoaded'"
    i = html.find(anchor)
    if i < 0:
        raise SystemExit("No encontré el bootstrap (DOMContentLoaded) para parchar.")
    start_tag = html.rfind("<script", 0, i)
    content_start = html.find(">", start_tag) + 1
    content_end = html.find("</script>", content_start)
    if start_tag < 0 or content_end < 0:
        raise SystemExit("No pude delimitar el bootstrap.")
    return html[:content_start] + "\n    " + NEW_BOOTSTRAP + "  " + html[content_end:]


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


def embeber(data_js, write=True):
    """Toma la plantilla, le incrusta el código (src/*) + los datos, y DEVUELVE el HTML resultante.
    Con write=True además guarda la plantilla actualizada en disco (deja respaldo .bak). Con
    write=False (modo rápido) no reescribe la plantilla: el HTML recién armado —ya con el código
    nuevo de src/— se usa solo en memoria como base del dashboard, así un cambio de CÓDIGO llega
    al dashboard sin tener que reescribir la plantilla."""
    if not os.path.isfile(HTML_PATH):
        log("  AVISO: no existe plantilla.html; se omite el embebido (solo data_v2.js).")
        return None
    import shutil
    html = open(HTML_PATH, "r", encoding="utf-8", newline="").read()
    # La barra de filtros debe ENVOLVER (no recortarse). El CSS base venía nowrap.
    html = html.replace("gap:10px;flex-wrap:nowrap;align-items:flex-end;",
                        "gap:10px;flex-wrap:wrap;align-items:flex-end;")
    # Indentación de los niveles 4 (CECO) y 5 (Contrapartida) — detalle máximo.
    if ".ind-4{" not in html:
        html = html.replace(".ind-3{padding-left:44px;}",
                            ".ind-3{padding-left:44px;}.ind-4{padding-left:66px;}.ind-5{padding-left:88px;}.ind-6{padding-left:110px;}")
    elif ".ind-5{" not in html:
        html = html.replace(".ind-4{padding-left:66px;}",
                            ".ind-4{padding-left:66px;}.ind-5{padding-left:88px;}.ind-6{padding-left:110px;}")
    if ".ind-6{" not in html:
        html = html.replace(".ind-5{padding-left:88px;}",
                            ".ind-5{padding-left:88px;}.ind-6{padding-left:110px;}")
    if ".ind-7{" not in html:   # niveles del detalle de línea (Texto pedido / Denominación)
        html = html.replace(".ind-6{padding-left:110px;}",
                            ".ind-6{padding-left:110px;}.ind-7{padding-left:132px;}.ind-8{padding-left:154px;}")
    if ".ind-9{" not in html:   # nivel extra al insertar Desc. CECO bajo Gerencia
        html = html.replace(".ind-8{padding-left:154px;}",
                            ".ind-8{padding-left:154px;}.ind-9{padding-left:176px;}")
    # Tabla Resumen: tabla al 100% (llena la pantalla), columna de nombres REDIMENSIBLE (arrastrando
    # el borde del encabezado) con un spacer que absorbe el sobrante, breadcrumb que baja de línea y
    # barrita de scroll horizontal fina. Bloque reemplazable (marcadores) para poder actualizarlo.
    _OLD_RESUMEN = (".mtable.resumen{width:auto;}"
        ".mtable.resumen td.name,.mtable.resumen thead .cols th.left{max-width:340px;}"
        ".mtable.resumen thead .cols th.left{white-space:normal;line-height:1.7;}"
        ".mtable.resumen td.name .twig{max-width:100%;min-width:0;}"
        ".mtable.resumen td.name .twig>span:not(.tog){overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}"
        ".hscroll-top{scrollbar-width:thin;}.hscroll-top::-webkit-scrollbar{height:10px;}")
    html = html.replace(_OLD_RESUMEN, "", 1)                       # limpia la 1ª versión (sin marcadores)
    html = re.sub(r"/\*RESUMEN-CSS\*/.*?/\*/RESUMEN-CSS\*/", "", html, flags=re.S)  # limpia versiones marcadas
    RESUMEN_CSS = ("/*RESUMEN-CSS*/"
        ".mtable.resumen{width:100%;}"
        ".mtable.resumen thead .cols th.left{white-space:normal;line-height:1.7;position:relative;}"
        ".mtable.resumen td.name{overflow:hidden;}"
        ".mtable.resumen td.name .twig{max-width:100%;min-width:0;}"
        ".mtable.resumen td.name .twig>span:not(.tog){overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}"
        ".mtable.resumen .col-resize{position:absolute;top:0;right:-4px;width:9px;height:100%;cursor:col-resize;z-index:3;}"
        ".mtable.resumen .col-resize:hover{background:color-mix(in srgb,var(--amsa-teal) 35%,transparent);}"
        ".mtable.resumen td.spacer-fill,.mtable.resumen th.spacer-fill{padding:0;border:none;background:transparent;}"
        # ⊘ ocultar fila: tabla Resumen y tabla de Presentación (Formato 2) → selector sin `.resumen`.
        ".mtable tbody .rowhide{opacity:0;border:none;background:transparent;color:var(--fg-soft);cursor:pointer;font-size:13px;line-height:1;margin-left:6px;padding:0 3px;transition:opacity .1s;}"
        ".mtable tbody tr:hover .rowhide{opacity:.5;}"
        ".mtable tbody .rowhide:hover{opacity:1;color:var(--amsa-red);}"
        ".hscroll-top{scrollbar-width:thin;}.hscroll-top::-webkit-scrollbar{height:10px;}"
        "/*/RESUMEN-CSS*/")
    html = html.replace(
        ".mtable{width:100%;border-collapse:separate;border-spacing:0;font-size:12.5px;}",
        ".mtable{width:100%;border-collapse:separate;border-spacing:0;font-size:12.5px;}" + RESUMEN_CSS, 1)
    # Bootstrap "todo incrustado": scripts inline + fuentes data: URI (abre desde file://
    # aunque esté descargado). Idempotente.
    antes = "/*INLINE-BOOTSTRAP-v3*/" in html
    html = _parchar_bootstrap(html)
    if not antes:
        log("  Bootstrap 'todo incrustado' aplicado (sin blob:).")
    antes_th = 'data-fallback-msg="2"' in html
    html = _parchar_thumbnail(html)
    if not antes_th:
        log("  Pantalla estática con aviso aplicada (visible sin JS).")
    log("  Validando JSX con Babel…")
    _validar_jsx(html, [n for (n, isjsx) in ASSETS.values() if isjsx])
    for uuid, (name, _isjsx) in ASSETS.items():
        html = _set_asset(html, uuid, open(os.path.join(SRC, name), encoding="utf-8").read())
    html = _set_asset(html, DATA_UUID, data_js)
    if write:
        shutil.copyfile(HTML_PATH, HTML_PATH + ".bak")
        open(HTML_PATH, "w", encoding="utf-8", newline="").write(html)
        log("  plantilla.html actualizada (código + datos embebidos).")
    else:
        log("  [modo rápido] no se reescribe plantilla.html; su código+datos nuevos quedan en memoria como base del dashboard.")
    return html


def leer_capex():
    """Lee la hoja 'BD AMSA' del Excel de CAPEX 2027 → estructura para la pestaña CAPEX.
    Columnas (0-idx, header en fila 2): 1 Compañía · 2 Código PEP · 3 Nombre Proyecto ·
    4 Nuevo/Remanente · 5 Sustaining/Development · 6 Estatus · 9 Gerencia Ejecutora ·
    11 Vicepresidencia · 27..38 $ Ene-27..$ Dic-27 · 39 Total 2027 (montos en USD)."""
    if not os.path.isfile(CAPEX_FILE):
        log(f"  AVISO: no existe {CAPEX_FILE}; se omite CAPEX.")
        return None
    log("Leyendo CAPEX…")
    wb = openpyxl.load_workbook(CAPEX_FILE, data_only=True, read_only=True)
    if CAPEX_SHEET not in wb.sheetnames:
        log(f"  AVISO: la hoja '{CAPEX_SHEET}' no está en {CAPEX_FILE}; se omite CAPEX.")
        return None
    ws = wb[CAPEX_SHEET]
    MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    def _num(v):
        try:
            return float(v) if v is not None else 0.0
        except (TypeError, ValueError):
            return 0.0
    def _txt(v, default="(vacío)"):
        s = str(v).strip() if v is not None else ""
        return s if s else default
    def _nr(v):
        s = (str(v).strip() if v is not None else "").lower()
        if s.startswith("nuevo"):     return "Nuevo"
        if s.startswith("remanente"): return "Remanente"
        return "(sin clasificar)"
    def _est(v):
        s = (str(v).strip() if v is not None else "").lower()
        if s == "aprobado":    return "Aprobado"
        if s == "no aprobado": return "No aprobado"
        return "Sin estado"
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r is None or len(r) < 40:
            continue
        proj = r[3]
        if proj is None or str(proj).strip() == "":
            continue
        months = [round(_num(r[27 + i]), 2) for i in range(12)]
        tot = round(_num(r[39]), 2)
        if tot == 0:
            tot = round(sum(months), 2)
        fcst = round(_num(r[26]), 2)   # Forecast 5+7 2026 = Acumulado Dic 2026 (F5+7), en USD
        # AVANCE FÍSICO (% acumulado): col 45 = al cierre Dic-26 · cols 46..57 = F Ene-27..F Dic-27.
        # Es un % ACUMULADO por hito (meses sin dato = mismo valor del hito anterior → arrastre).
        def _pct(v):
            try:
                return float(v) if v is not None else None
            except (TypeError, ValueError):
                return None
        f0 = _pct(r[45]) if len(r) > 45 else None
        _fm = [(_pct(r[46 + i]) if len(r) > 46 + i else None) for i in range(12)]
        any_f = (f0 is not None) or any(v is not None for v in _fm)
        # Curva COMPLETA (12 meses, sin huecos) para que la agregación ponderada tenga denominador
        # fijo → % monótono. Meses previos al primer hito = baseline Dic-26 (o 0 si no hay baseline).
        # Máximo corrido: el % ACUMULADO no puede bajar (si el origen reporta un valor menor
        # —re-scope o error de carga— se mantiene el hito más alto ya alcanzado).
        f_curve, last = [], (f0 if f0 is not None else 0.0)
        for i in range(12):
            if _fm[i] is not None and _fm[i] > last:
                last = _fm[i]
            f_curve.append(round(last, 4))
        rows.append({
            "sd":   _txt(r[5], "(sin categoría)"),   # Sostenimiento / Desarrollo (Sustaining/Development/TMM)
            "vp":   _txt(r[11]),                       # Vicepresidencia
            "ger":  _txt(r[9]),                        # Gerencia Ejecutora
            "nr":   _nr(r[4]),                         # Nuevo / Remanente
            "proj": str(proj).strip(),                 # Nombre de Proyecto
            "pep":  _txt(r[2], ""),                    # Código PEP
            "comp": _txt(r[1], "AMSA"),                # Compañía
            "est":  _est(r[6]),                        # Estatus
            "m":    months,                            # $ Ene-27 .. $ Dic-27 (USD)
            "tot":  tot,                               # Total 2027 (USD)
            "fcst": fcst,                              # Forecast 5+7 2026 = Acumulado Dic 2026 (USD)
            "f":    (f_curve if any_f else None),      # Avance físico % acumulado por mes (o null si el proyecto no reporta)
            "f0":   (round(f0, 4) if f0 is not None else None),   # Avance físico % al cierre Dic-26 (punto de partida)
        })
    log(f"  CAPEX: {len(rows)} proyectos · Total 2027 {sum(x['tot'] for x in rows)/1e6:.1f} MM USD")
    return {"months": MESES, "rows": rows}


def construir_dashboard(base_html, data_js, det, detp=None, capex=None):
    """Arma el dashboard final (salida/): parte del HTML ya embebido (base_html, con el código+datos
    nuevos) y le agrega el DETALLE del gasto:
    · window.DET  = detalle Real (Contrapartida › Texto pedido › Denominación).
    · window.DETP = detalle Ppto/Forecast (Concepto Gasto › Actividad).
    El mismo código JSX los activa solo si existen. Es el HTML que se comparte (~7 MB, abre offline)."""
    if not base_html:
        log("  AVISO: no hay HTML base embebido; se omite el dashboard.")
        return
    import zipfile
    os.makedirs(SALIDA_DIR, exist_ok=True)
    html = base_html
    det_js = "window.DET=" + json.dumps(det, ensure_ascii=False, separators=(",", ":")) + ";"
    if detp is not None:
        det_js += "\nwindow.DETP=" + json.dumps(detp, ensure_ascii=False, separators=(",", ":")) + ";"
    if capex is not None:
        det_js += "\nwindow.CAPEX_DATA=" + json.dumps(capex, ensure_ascii=False, separators=(",", ":")) + ";"
    html = _set_asset(html, DATA_UUID, data_js + "\n" + det_js)
    html = html.replace("<title>Actividad Corporativa + Distribuibles · AMSA</title>",
                        "<title>Actividad Corporativa + Distribuibles (detalle) · AMSA</title>")
    open(SALIDA_HTML, "w", encoding="utf-8", newline="").write(html)
    zip_path = os.path.join(SALIDA_DIR, "Dashboard Corporativo.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(SALIDA_HTML, arcname=os.path.basename(SALIDA_HTML))
        z.writestr("LEER - Como abrir.txt", "Doble clic en el HTML (Chrome/Edge/Firefox). Incluye detalle del gasto.")
    mb = os.path.getsize(SALIDA_HTML) / 1e6
    log(f"  Dashboard generado: {SALIDA_HTML} ({mb:.1f} MB) · ZIP generado.")


# ===========================================================================
#  BBDD COMPLETA (esquema estrella, parquet) — espejo REGENERABLE de la MISMA data que se
#  embebe en el HTML. Varias tablas en bbdd/: hechos (registros/detalle/dotaciones) +
#  diccionarios (cecos/items/clacos/compañías). Se unen por ceco / item / claco. NO es un
#  deliverable (no es el dashboard); se reescribe en cada corrida (completa o --rapido).
# ===========================================================================
_BBDD_TABLAS = {
    "fact_registros": "Real + Ppto histórico (long por bucket). Grano: ceco×item×contra×claco×bucket. "
                      "buckets: 2022-2025, 2026ytd, 2026fy. Valores {n=Normal USD, a=Ajustada 2027}.",
    "fact_anual": "Forecast 5+7 2026 y Ppto 2027 (anual). Grano: ceco×item×claco×medida.",
    "fact_detalle_real": "Detalle del Real línea a línea (Contrapartida › Texto pedido › Denominación › Documento). "
                         "Grano: ceco×item×contra×texto×denom×doc×bucket. st=1 → Services & Tech.",
    "fact_detalle_ppto": "Detalle de Ppto/Forecast (Concepto Gasto › Actividad). Grano: ceco×item×concepto×actividad×medida.",
    "dim_cecos": "Diccionario de CECO (ambas estructuras). Une por ceco. estructura ∈ {new, old}.",
    "dim_items": "Diccionario de Ítem (Agrupación4): nombre, tipo costo, ítem relevante, clasificación cuenta. Une por item.",
    "dim_clacos": "Diccionario de CLACO (Clase de Costo) → Desc. CLACO. Une por claco.",
    "dim_companias": "Diccionario de compañías. Une por codigo (= prefijo del ceco).",
    "fact_dotaciones": "Dotaciones (FTE) Propios/Contratista por VP/Gerencia, año y período.",
}


def escribir_bbdd(agg, fcst, ppto27, det, detp, cecoNew, cecoOld, comps,
                  itemNames, itemTc, itemRel, relNames, itemClas, claco2name, dot_tidy):
    """Vuelca TODA la info del dashboard como tablas parquet (esquema estrella) en bbdd/."""
    os.makedirs(BBDD_DIR, exist_ok=True)
    log("Escribiendo BBDD (parquet, esquema estrella)…")

    def w(name, rows, cols):
        pd.DataFrame(rows, columns=cols).to_parquet(os.path.join(BBDD_DIR, name + ".parquet"), index=False)
        log(f"    bbdd/{name}.parquet · {len(rows)} filas")

    _cl = lambda cl: (cl if cl and cl != "(sin claco)" else None)

    # 1) fact_registros — Real + Ppto histórico (long por bucket), tal cual el agg.
    w("fact_registros",
      [dict(ceco=c, item=i, contra=co, claco=_cl(cl), bucket=b,
            real_n=_round(v[0]), real_a=_round(v[1]), plan_n=_round(v[2]), plan_a=_round(v[3]))
       for (c, i, co, b, cl), v in agg.items()],
      ["ceco", "item", "contra", "claco", "bucket", "real_n", "real_a", "plan_n", "plan_a"])

    # 2) fact_anual — Forecast 5+7 2026 y Ppto 2027 (anual, por ceco×item×claco).
    anual = [dict(ceco=c, item=i, claco=_cl(cl), medida="forecast_2026", valor_n=_round(n), valor_a=_round(a))
             for (c, i, cl), (n, a) in fcst.items()]
    anual += [dict(ceco=c, item=i, claco=_cl(cl), medida="ppto_2027", valor_n=_round(v), valor_a=_round(v))
              for (c, i, cl), v in ppto27.items()]
    w("fact_anual", anual, ["ceco", "item", "claco", "medida", "valor_n", "valor_a"])

    # 3) fact_detalle_real — de-interna el blob DET (índices → strings).
    S = det["s"]; IDXBK = ["2022", "2023", "2024", "2025", "2026ytd"]
    drows = []
    for key, rows in det["k"].items():
        item, contra = key.split("\x01", 1)
        for ci, ti, di, doi, bk, n, a, st in rows:
            drows.append(dict(ceco=S[ci], item=item, contra=contra, texto_pedido=S[ti], denominacion=S[di],
                              documento=S[doi], bucket=IDXBK[bk] if bk < len(IDXBK) else str(bk),
                              valor_n=n, valor_a=a, st=st))
    w("fact_detalle_real", drows,
      ["ceco", "item", "contra", "texto_pedido", "denominacion", "documento", "bucket", "valor_n", "valor_a", "st"])

    # 4) fact_detalle_ppto — de-interna el blob DETP.
    Sp = detp["s"]; MED = {0: "ppto_2025", 1: "ppto_2026ytd", 2: "ppto_2026fy", 3: "forecast_2026", 4: "ppto_2027", 5: "out66"}
    prows = []
    for item, rows in detp["k"].items():
        for ci, cgi, ai, med, n, a in rows:
            prows.append(dict(ceco=Sp[ci], item=item, concepto_gasto=Sp[cgi], actividad=Sp[ai],
                              medida=MED.get(med, str(med)), valor_n=n, valor_a=a))
    w("fact_detalle_ppto", prows,
      ["ceco", "item", "concepto_gasto", "actividad", "medida", "valor_n", "valor_a"])

    # 5) dim_cecos — ambas estructuras (long por 'estructura').
    crows = []
    for estr, mapa in (("new", cecoNew), ("old", cecoOld)):
        for c, d in mapa.items():
            crows.append(dict(ceco=c, estructura=estr, vp=d.get("vp"), gerencia=d.get("ger"),
                              desc_ceco=d.get("dceco"), tipo_costo=d.get("tc"), aplica=d.get("ap"),
                              clasificacion=d.get("cl"), compania=d.get("comp")))
    w("dim_cecos", crows,
      ["ceco", "estructura", "vp", "gerencia", "desc_ceco", "tipo_costo", "aplica", "clasificacion", "compania"])

    # 6) dim_items — nombre + tipo costo + ítem relevante + clasificación cuenta.
    allit = set(itemNames) | set(itemTc) | set(itemRel) | set(itemClas)
    irows = []
    for it in sorted(allit):
        rel = itemRel.get(it)
        irows.append(dict(item=it, nombre=itemNames.get(it, it), tipo_costo=itemTc.get(it),
                          item_relevante=rel, item_relevante_nombre=(relNames.get(rel) if rel else None),
                          clasif_cuenta=itemClas.get(it)))
    w("dim_items", irows,
      ["item", "nombre", "tipo_costo", "item_relevante", "item_relevante_nombre", "clasif_cuenta"])

    # 7) dim_clacos — código → Desc. CLACO (diccionario completo).
    w("dim_clacos", [dict(claco=c, desc_claco=n) for c, n in sorted(claco2name.items())],
      ["claco", "desc_claco"])

    # 8) dim_companias.
    w("dim_companias",
      [dict(codigo=k, nombre=v.get("nombre"), abrev=v.get("abrev"), clasificacion=v.get("clasif"))
       for k, v in comps.items()],
      ["codigo", "nombre", "abrev", "clasificacion"])

    # 9) fact_dotaciones — ya viene tidy.
    w("fact_dotaciones", dot_tidy or [], ["src", "vp", "ger", "year", "period", "real", "plan"])

    # LEEME con el esquema y las claves de unión.
    lines = ["BBDD del Dashboard Corporativo — tablas parquet (esquema estrella)",
             "=" * 66,
             "Espejo REGENERABLE de la MISMA data embebida en el HTML. Se reescribe en cada",
             "corrida de construir.py. Valores {n}=Normal (USD) · {a}=Ajustada 2027.",
             "Se leen con pandas (pd.read_parquet), Power BI/Power Query o DuckDB.", "",
             "Uniones: fact.* ⋈ dim_cecos por 'ceco' · ⋈ dim_items por 'item' · ⋈ dim_clacos",
             "por 'claco' · ⋈ dim_companias por codigo = ceco[:4]. En dim_cecos filtrá",
             "estructura='new' (o 'old') según el toggle 'Estructura CECOS' del tablero.", ""]
    for name, desc in _BBDD_TABLAS.items():
        lines.append(f"· {name}.parquet")
        lines.append(f"    {desc}")
    open(os.path.join(BBDD_DIR, "LEEME.txt"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    log(f"  BBDD (esquema estrella) escrita en {BBDD_DIR} · {len(_BBDD_TABLAS)} tablas + LEEME.txt")


def _reversas_files():
    """Descubre los ajustes de reversas: cada .xlsx en REV_DIR que tenga una hoja cuyo nombre
    termina en «Unpivot». Devuelve [(path, hoja, etiqueta), …] ordenado por nombre de archivo.
    Así, agregar un ajuste al Forecast = dejar el Excel ahí con su hoja «… Unpivot» (sin tocar código)."""
    out = []
    if not os.path.isdir(REV_DIR):
        return out
    for fn in sorted(os.listdir(REV_DIR)):
        if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xlsm")):
            continue
        path = os.path.join(REV_DIR, fn)
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            hoja = next((s for s in wb.sheetnames if s.strip().lower().endswith("unpivot")), None)
            wb.close()
        except Exception as e:
            log(f"  AVISO: no pude abrir {fn} ({type(e).__name__}); se omite.")
            hoja = None
        if hoja:
            out.append((path, hoja, os.path.splitext(fn)[0]))
        else:
            log(f"  AVISO: {fn} no tiene hoja «… Unpivot»; se omite (¿corriste unpivot_ejercicio.py?).")
    return out


# ===========================================================================
def main(rapido=False):
    # rapido=True: regenera SOLO el dashboard de salida/ (re-embebe en memoria el código de src/ +
    # los datos y los vuelca al dashboard). NO reescribe la plantilla ni los intermedios (parquet,
    # data.js). Es la opción para previsualizar rápido cambios de código o de Ppto.
    log("=== Construir Dashboard Corporativo " + ("(modo rápido)" if rapido else "(completo)") + " ===")
    log("Leyendo diccionarios…")
    code2item, name2item, item2name, item2tc, item2rel, rel2name, item2clas, st_names, name2claco, claco2name = leer_clacos()
    cecoNew = leer_cecos("CECOS Corporativo Nueva")
    cecoOld = leer_cecos("CECOS Corporativo")
    comps = leer_companias()
    log(f"  CLACO: {len(code2item)} códigos · {len(name2item)} nombres · {len(item2name)} ítems (Agrupación4)")
    log(f"  CECOS: {len(cecoNew)} (nueva) · {len(cecoOld)} (antigua) · Compañías: {len(comps)}")
    log(f"  Services & Tech: códigos {sorted(ST_CODES)} · nombres Real (Agrupación5): {sorted(st_names)}")

    log("Leyendo Reales…")
    agg, realcecos = leer_reales(name2item, name2claco, code2item)
    log("Leyendo Planes…")
    agg = leer_planes(code2item, agg)

    # ALCANCE: se muestran TODOS los CECOs (con o sin Real). Antes se excluían los CECOs
    # corporativos sin Real (Capex, Exploraciones, Proyectos, arriendos…); esa regla se quitó
    # a pedido — ahora si un CECO trae Ppto/Forecast/Ppto2027 sin Real, igual aparece.
    def _en_alcance(ceco):
        return True
    log(f"  Alcance: TODOS los CECOs (sin la regla 'solo corporativos con Real'). realcecos={len(realcecos)}")

    df = escribir_parquet(agg, write=not rapido)
    records = construir_registros(agg)

    # Forecast 5+7 2026 (anual): pseudo-registros con contra "(Forecast)" (se oculta en la
    # tabla, como "(Presupuesto)") que solo aportan el valor 'fcst' {n,a}. Sirve como columna
    # de comparación (categoría Forecast) sin afectar Real/Ppto.
    log("Leyendo Forecast 5+7…")
    fcst = leer_forecast(code2item, _en_alcance)
    ZERO = lambda: {"real": {"n": 0, "a": 0}, "plan": {"n": 0, "a": 0}}
    for (ceco, item, claco), (vn, va) in fcst.items():
        rec = {"ceco": ceco, "item": item, "contra": "(Forecast)"}
        if claco and claco != "(sin claco)":
            rec["claco"] = claco
        if claco in ST_CODES:
            rec["st"] = 1
        for y in YEARS_HIST:
            rec[f"y{y}"] = ZERO()
        rec["y2026"] = ZERO()
        rec["y2026fy"] = {"plan": {"n": 0, "a": 0}}
        rec["fcst"] = {"n": _round(vn), "a": _round(va)}
        records.append(rec)

    # Outlook 6+6 2026 (anual): idéntico al Forecast 5+7 pero en su propia medida 'out66',
    # con contra "(Outlook)". Es una columna MÁS de la categoría Forecast (no lo reemplaza).
    log("Leyendo Outlook 6+6…")
    out66 = leer_forecast(code2item, _en_alcance, OUT_FILE, OUT_SHEET, "Outlook 6+6")
    for (ceco, item, claco), (vn, va) in out66.items():
        rec = {"ceco": ceco, "item": item, "contra": "(Outlook)"}
        if claco and claco != "(sin claco)":
            rec["claco"] = claco
        if claco in ST_CODES:
            rec["st"] = 1
        for y in YEARS_HIST:
            rec[f"y{y}"] = ZERO()
        rec["y2026"] = ZERO()
        rec["y2026fy"] = {"plan": {"n": 0, "a": 0}}
        rec["out66"] = {"n": _round(vn), "a": _round(va)}
        records.append(rec)

    # Reversas (ajustes al Forecast 5+7): pseudo-registros con contra "(Forecast)" y flag adj=1.
    # Se SUMAN al Forecast 5+7 SOLO cuando la Estructura CECOS activa es «Nueva con ajustes»
    # (el modelo los oculta en los demás modos). Suelen ser redistribuciones (neto ≈ 0) que
    # cambian el reparto por VP/Gerencia/Ítem sin mover el gran total. AUTO-DESCUBIERTAS: cada
    # .xlsx en ajustes_forecast con una hoja «… Unpivot» se suma (no requiere editar código).
    log("Leyendo Reversas (ajustes al Forecast)…")
    for _rev_path, _rev_sheet, _rev_lbl in _reversas_files():
        rev = leer_forecast(code2item, _en_alcance, _rev_path, _rev_sheet, "Reversas·" + _rev_lbl[:26])
        for (ceco, item, claco), (vn, va) in rev.items():
            rec = {"ceco": ceco, "item": item, "contra": "(Forecast)", "adj": 1}
            if claco and claco != "(sin claco)":
                rec["claco"] = claco
            if claco in ST_CODES:
                rec["st"] = 1
            for y in YEARS_HIST:
                rec[f"y{y}"] = ZERO()
            rec["y2026"] = ZERO()
            rec["y2026fy"] = {"plan": {"n": 0, "a": 0}}
            rec["fcst"] = {"n": _round(vn), "a": _round(va)}
            records.append(rec)

    # Ppto 2027 (anual, valor tal cual): pseudo-registros con contra "(Ppto2027)" que solo
    # aportan 'prop27'. Alimenta la columna/comparación "Ppto 2027" (antes Propuesta, base 0).
    # Es un único valor por (ceco,item): original = ajustada 2027 (no depende de Base Moneda).
    log("Leyendo Ppto 2027…")
    ppto27 = leer_ppto27(code2item, _en_alcance)
    for (ceco, item, claco), v in ppto27.items():
        rec = {"ceco": ceco, "item": item, "contra": "(Ppto2027)"}
        if claco and claco != "(sin claco)":
            rec["claco"] = claco
        if claco in ST_CODES:
            rec["st"] = 1
        for y in YEARS_HIST:
            rec[f"y{y}"] = ZERO()
        rec["y2026"] = ZERO()
        rec["y2026fy"] = {"plan": {"n": 0, "a": 0}}
        rec["prop27"] = _round(v)
        records.append(rec)

    # Nombres, Tipo Costo (C1/C2/C3) y Clasificación Cuenta de los ítems presentes.
    itemNames, itemTc, itemClas = {}, {}, {}
    itemRel, relNames = {}, {}   # Ítem(Ag4) → Ítem Relevante(Ag3) y sus nombres
    for r in records:
        it = r["item"]
        itemNames[it] = item2name.get(it, it) if it != "(sin ítem)" else "(sin ítem)"
        tc = item2tc.get(it)
        if tc:
            itemTc[it] = tc
        clas = item2clas.get(it)
        if clas:
            itemClas[it] = clas
        rel = item2rel.get(it)
        if rel:
            itemRel[it] = rel
            relNames[rel] = rel2name.get(rel, rel)
    log(f"  Ítem Relevante (Agrupación3): {len(relNames)} grupos · {len(itemRel)} ítems mapeados · Clasif Cuenta: {len(set(itemClas.values()))} valores")

    log("Leyendo Dotaciones…")
    dot_records, dot_tidy = leer_dotaciones()
    if dot_tidy and not rapido:
        pd.DataFrame(dot_tidy, columns=["src", "vp", "ger", "year", "period", "real", "plan"]) \
            .to_parquet(os.path.join(HERE, "dotaciones_v2.parquet"), index=False)

    data_js = construir_data_js(records, itemNames, itemTc, cecoNew, cecoOld, comps, dot_records, itemRel, relNames, itemClas, claco2name)
    if not rapido:
        open(DATA_JS, "w", encoding="utf-8", newline="").write(data_js)
        log(f"  data.js: {DATA_JS} ({len(records)} registros ceco×ítem)")

    # Validación rápida (totales en MM USD)
    _validar(df, cecoNew, comps)

    # Se re-embebe SIEMPRE el código actual de src/ + los datos (así el dashboard toma los cambios
    # de código). En modo rápido (write=False) NO se reescribe la plantilla en disco: el HTML recién
    # armado se usa solo en memoria como base del dashboard.
    base_html = embeber(data_js, write=not rapido)

    # Dashboard final = misma app + DETALLE del gasto embebido: Real (Contrapartida › Texto pedido ›
    # Denominación) y Ppto/Forecast (Concepto Gasto › Actividad).
    log("Construyendo el detalle del gasto…")
    det = leer_detalle(name2item, _en_alcance, st_names, name2claco, code2item)
    detp = leer_detalle_pf(code2item, _en_alcance)
    capex = leer_capex()
    construir_dashboard(base_html, data_js, det, detp, capex)

    # BBDD completa (parquet, esquema estrella): toda la data del dashboard en tablas unibles.
    # Se escribe SIEMPRE (también en modo rápido): no es un deliverable, refleja la data vigente.
    escribir_bbdd(agg, fcst, ppto27, det, detp, cecoNew, cecoOld, comps,
                  itemNames, itemTc, itemRel, relNames, itemClas, claco2name, dot_tidy)
    log("\n✓ LISTO.")


def _validar(df, cecomap, comps):
    log("  --- Totales (MM USD) ---")
    for b in ["2022", "2023", "2024", "2025", "2026ytd", "2026fy"]:
        sub = df[df.bucket == b]
        log(f"    {b:8}: Real N {sub.real_n.sum()/1e6:8.2f} / A {sub.real_a.sum()/1e6:8.2f}"
            f"  ·  Ppto N {sub.plan_n.sum()/1e6:8.2f} / A {sub.plan_a.sum()/1e6:8.2f}")


if __name__ == "__main__":
    import sys
    main(rapido="--rapido" in sys.argv)
