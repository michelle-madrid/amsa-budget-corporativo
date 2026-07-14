# -*- coding: utf-8 -*-
"""
dividir_por_vp.py — Divide el Dashboard v3 en UN HTML por Vicepresidencia (solo su data).
=========================================================================================
Requiere haber corrido antes `python construir_v2.py` (usa el v3 ya generado).
Uso:  python dividir_por_vp.py

Lee  : v3/Dashboard Corporativo v3.html   (autocontenido, con detalle)
Escribe: v3/por_vp/<VP>/Dashboard <VP>.html    (uno por Vicepresidencia)

Cada archivo:
  · Los registros cuyo CECO pertenece a esa VP en la estructura ANTIGUA **o** en la NUEVA
    (un CECO que cambió de VP entre estructuras aparece en ambos archivos; en cada uno el
    modelo lo oculta en la estructura donde no corresponde). Así no se pierde plata al togglear.
  · CECOS (cada mapa filtrado por su propia VP), DETALLE (re-interned) y Forecast/Ppto2027 acotados.
  · SIN Dotaciones (se vacía DOT_DATA → la pestaña "Dotaciones AMSA (FTE)" se oculta sola).
  · Título (H1 y pestaña del navegador) con el nombre de la VP.

SPLIT POR GERENCIA: además, una VP puede dividirse por Gerencia en SUBCARPETAS dentro de su
  carpeta (ver SPLIT_POR_GERENCIA). Ej.: VP Finanzas → una subcarpeta por Gerencia, con combos
  (TICA = TICA + TICA Corporativo, etc.). Cada subcarpeta lleva su dashboard filtrado.
Reutiliza el embebido (gzip+base64) y la plantilla del v3 vía construir_v2.
"""
import os
import re
import json
import construir_v2 as B   # reutiliza V3_HTML/V3_DIR, DATA_UUID, ASSETS, _get_asset, _set_asset

OUT_DIR = os.path.join(B.V3_DIR, "por_vp")
MODEL_UUID = next(u for u, (n, _) in B.ASSETS.items() if n == "model.js")
_j = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))

# División interna de una VP por GERENCIA: dentro de su carpeta, una subcarpeta por Gerencia
# (o combo) con su propio dashboard filtrado (estructura nueva). Las gerencias listadas en un
# combo van JUNTAS bajo el nombre del grupo; el resto de las gerencias de la VP van cada una
# en su propia subcarpeta. La gerencia que se llama igual que la VP se omite.
SPLIT_POR_GERENCIA = {
    "VP Finanzas": [
        ("TICA", {"TICA", "TICA Corporativo"}),
        ("Abastecimiento", {"Abastecimiento", "Gerencia de Abastecimiento Corporativo", "Gestión Servicios y Control"}),
        ("Gcia de Planificación y Gestión", {"Gcia de Planificación y Gestión", "Gerencia Competitividad", "Programa competitividad de costos"}),
        ("Grcia. Riesgos y Control Interno", {"Grcia. Riesgos y Control Interno", "Gerencia de Riesgos"}),
    ],
}


def _extract(js, name):
    """Devuelve el valor JSON de `window.<name> = <json>;` dentro del texto del asset."""
    m = re.search(r"window\." + name + r"\s*=\s*", js)
    if not m:
        return None
    return json.JSONDecoder().raw_decode(js, m.end())[0]


def _safe(name):
    return re.sub(r'[\\/:*?"<>|]+', " ", str(name)).strip() or "SIN VP"


def _filtrar_det(DET, cecos):
    """Recorta el detalle Real a los CECOs dados, re-internando strings (archivo liviano)."""
    if not DET:
        return None
    S = DET["s"]; newS, idx = [], {}
    def I(x):
        k = idx.get(x)
        if k is None:
            k = len(newS); idx[x] = k; newS.append(x)
        return k
    K = {}
    for key, rows in DET["k"].items():
        fr = [[I(S[r[0]]), I(S[r[1]]), I(S[r[2]]), I(S[r[3]]), r[4], r[5], r[6], r[7]]
              for r in rows if S[r[0]] in cecos]
        if fr:
            K[key] = fr
    return {"s": newS, "k": K}


def _filtrar_detp(DETP, cecos):
    """Recorta el detalle Ppto/Forecast (Concepto Gasto › Actividad) a los CECOs dados."""
    if not DETP:
        return None
    Sp = DETP["s"]; newSp, idxp = [], {}
    def Ip(x):
        k = idxp.get(x)
        if k is None:
            k = len(newSp); idxp[x] = k; newSp.append(x)
        return k
    Kp = {}
    for key, rows in DETP["k"].items():
        fr = [[Ip(Sp[r[0]]), Ip(Sp[r[1]]), Ip(Sp[r[2]]), r[3], r[4], r[5]]
              for r in rows if Sp[r[0]] in cecos]
        if fr:
            Kp[key] = fr
    return {"s": newSp, "k": Kp}


def _escribir(html, model, LOGO, DET, DETP, V, recs, cecoOldSub, cecoNewSub, titulo, fn):
    """Escribe UN HTML autocontenido con la data recortada (recs + mapas CECO + detalle) y el
    título dado. Devuelve el tamaño en MB."""
    cecos = {r["ceco"] for r in recs}
    Vf = dict(V)
    Vf["records"] = recs
    Vf["cecoOld"] = cecoOldSub
    Vf["cecoNew"] = cecoNewSub
    parts = ["window.V2_DATA = " + _j(Vf) + ";", 'window.DOT_DATA = {"records":[]};']
    if LOGO:
        parts.append("window.CORP_LOGO = " + _j(LOGO) + ";")
    detf = _filtrar_det(DET, cecos)
    if detf is not None:
        parts.append("window.DET = " + _j(detf) + ";")
    detpf = _filtrar_detp(DETP, cecos)
    if detpf is not None:
        parts.append("window.DETP = " + _j(detpf) + ";")
    out = B._set_asset(html, B.DATA_UUID, "\n".join(parts))
    # Título en el H1 (TEXTOS de model.js) y en la pestaña del navegador.
    model_vp = (model
                .replace('"titulo": "Actividad Corporativa"', '"titulo": ' + _j(str(titulo)))
                .replace('"tituloPlus": "+ Distribuibles Ppto 2027"', '"tituloPlus": ""'))
    out = B._set_asset(out, MODEL_UUID, model_vp)
    out = re.sub(r"<title>.*?</title>", "<title>" + _safe(titulo) + " · AMSA</title>", out, count=1)
    os.makedirs(os.path.dirname(fn), exist_ok=True)
    open(fn, "w", encoding="utf-8", newline="").write(out)
    return os.path.getsize(fn) / 1e6


def main():
    if not os.path.isfile(B.V3_HTML):
        raise SystemExit("No existe el v3; corré primero: python construir_v2.py")
    html = open(B.V3_HTML, "r", encoding="utf-8", newline="").read()
    model = B._get_asset(html, MODEL_UUID)
    data = B._get_asset(html, B.DATA_UUID)
    V = _extract(data, "V2_DATA")
    LOGO = _extract(data, "CORP_LOGO")
    DET = _extract(data, "DET")
    DETP = _extract(data, "DETP")   # detalle Ppto/Forecast (Concepto Gasto › Actividad)
    if V is None:
        raise SystemExit("No pude leer window.V2_DATA del v3.")

    cecoOld = V.get("cecoOld", {}) or {}
    cecoNew = V.get("cecoNew", {}) or {}

    # Un CECO puede pertenecer a DISTINTA VP en la estructura antigua vs. la nueva. Para no
    # perder plata, un registro va al archivo de una VP si su CECO pertenece a esa VP en la
    # estructura ANTIGUA **o** en la NUEVA. Dentro del archivo, el modelo lo oculta donde no aplica.
    def vpsOf(ceco):
        s = set()
        mo, mn = cecoOld.get(ceco), cecoNew.get(ceco)
        if mo and mo.get("vp"):
            s.add(mo["vp"])
        if mn and mn.get("vp"):
            s.add(mn["vp"])
        return s or {"(sin VP)"}

    por_vp = {}
    for r in V["records"]:
        for vp in vpsOf(r["ceco"]):
            por_vp.setdefault(vp, []).append(r)

    # Solo se emiten carpetas de VPs de la estructura NUEVA (vigente).
    new_vps = {m.get("vp") for m in cecoNew.values() if m.get("vp")}
    por_vp = {vp: recs for vp, recs in por_vp.items() if vp in new_vps}

    os.makedirs(OUT_DIR, exist_ok=True)
    # Limpieza de corridas anteriores (archivos planos y carpetas por VP).
    for nombre in os.listdir(OUT_DIR):
        ruta = os.path.join(OUT_DIR, nombre)
        if os.path.isfile(ruta) and nombre.lower().endswith(".html"):
            os.remove(ruta)
        elif os.path.isdir(ruta):
            import shutil
            shutil.rmtree(ruta)
    B.log(f"Dividiendo el v3 en {len(por_vp)} Vicepresidencias…")
    generados = []
    for vp, recs in sorted(por_vp.items()):
        # Cada mapa de estructura se filtra por SU PROPIA VP (independiente del otro).
        cecosOld_vp = {c for c, m in cecoOld.items() if m.get("vp") == vp}
        cecosNew_vp = {c for c, m in cecoNew.items() if m.get("vp") == vp}
        fn = os.path.join(OUT_DIR, _safe(vp), "Dashboard " + _safe(vp) + ".html")
        mb = _escribir(html, model, LOGO, DET, DETP, V, recs,
                       {c: m for c, m in cecoOld.items() if c in cecosOld_vp},
                       {c: m for c, m in cecoNew.items() if c in cecosNew_vp}, vp, fn)
        generados.append((vp, len(recs), len(cecosOld_vp | cecosNew_vp), mb))

    B.log("  --- Dashboards por VP generados ---")
    for vp, nr, nc, mb in generados:
        B.log(f"    {vp:42} {nr:5} regs · {nc:3} CECOs · {mb:4.1f} MB")

    # División interna por GERENCIA: dentro de la carpeta de la VP, una subcarpeta por Gerencia
    # (o combo) con su dashboard filtrado. Combos = gerencias juntas; el resto, cada una sola.
    n_ger = 0
    for vp_split, combos in SPLIT_POR_GERENCIA.items():
        if vp_split not in por_vp:
            B.log(f"  (split por Gerencia de '{vp_split}': la VP no tiene dashboard, se omite)")
            continue
        gers_vp = sorted({m["ger"] for m in cecoNew.values()
                          if m.get("vp") == vp_split and m.get("ger") and m["ger"] != vp_split})
        en_combo = set().union(*[g for _, g in combos]) if combos else set()
        grupos = list(combos) + [(g, {g}) for g in gers_vp if g not in en_combo]
        for nombre, gset in grupos:
            cnew = {c: m for c, m in cecoNew.items() if m.get("vp") == vp_split and m.get("ger") in gset}
            cecos_g = set(cnew)
            recs = [r for r in V["records"] if r["ceco"] in cecos_g]
            if not recs:
                continue
            fn = os.path.join(OUT_DIR, _safe(vp_split), _safe(nombre), "Dashboard " + _safe(nombre) + ".html")
            _escribir(html, model, LOGO, DET, DETP, V, recs,
                      {c: m for c, m in cecoOld.items() if c in cecos_g}, cnew, nombre, fn)
            n_ger += 1
            B.log(f"    {vp_split} › {nombre:44} {len(recs):5} regs · {len(cecos_g):2} CECOs")

    B.log(f"\n✓ LISTO. {len(generados)} dashboards por VP + {n_ger} por Gerencia en {OUT_DIR}")


if __name__ == "__main__":
    main()
