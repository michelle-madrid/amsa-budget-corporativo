# -*- coding: utf-8 -*-
"""
dividir_por_vp.py — Divide el Dashboard v3 en UN HTML por Vicepresidencia (solo su data).
=========================================================================================
Requiere haber corrido antes `python construir_v2.py` (usa el v3 ya generado).
Uso:  python dividir_por_vp.py

Lee  : v3/Dashboard Corporativo v3.html   (autocontenido, con detalle)
Escribe: v3/por_vp/Dashboard <VP>.html    (uno por Vicepresidencia)

Cada archivo:
  · Los registros cuyo CECO pertenece a esa VP en la estructura ANTIGUA **o** en la NUEVA
    (un CECO que cambió de VP entre estructuras aparece en ambos archivos; en cada uno el
    modelo lo oculta en la estructura donde no corresponde). Así no se pierde plata al togglear.
  · CECOS (cada mapa filtrado por su propia VP), DETALLE (re-interned) y Forecast/Ppto2027 acotados.
  · SIN Dotaciones (se vacía DOT_DATA → la pestaña "Dotaciones AMSA (FTE)" se oculta sola).
  · Título (H1 y pestaña del navegador) con el nombre de la VP.
Reutiliza el embebido (gzip+base64) y la plantilla del v3 vía construir_v2.
"""
import os
import re
import json
import construir_v2 as B   # reutiliza V3_HTML/V3_DIR, DATA_UUID, ASSETS, _get_asset, _set_asset

OUT_DIR = os.path.join(B.V3_DIR, "por_vp")
MODEL_UUID = next(u for u, (n, _) in B.ASSETS.items() if n == "model.js")
_j = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))


def _extract(js, name):
    """Devuelve el valor JSON de `window.<name> = <json>;` dentro del texto del asset."""
    m = re.search(r"window\." + name + r"\s*=\s*", js)
    if not m:
        return None
    return json.JSONDecoder().raw_decode(js, m.end())[0]


def _safe(name):
    return re.sub(r'[\\/:*?"<>|]+', " ", str(name)).strip() or "SIN VP"


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

    # Un CECO puede pertenecer a DISTINTA VP en la estructura antigua vs. la nueva.
    # Para no perder plata, un registro va al archivo de una VP si su CECO pertenece a
    # esa VP en la estructura ANTIGUA **o** en la NUEVA (puede quedar en 2 archivos).
    # Dentro de cada archivo, el modelo oculta el CECO en la estructura donde no aplica.
    def vpsOf(ceco):
        s = set()
        mo, mn = cecoOld.get(ceco), cecoNew.get(ceco)
        if mo and mo.get("vp"):
            s.add(mo["vp"])
        if mn and mn.get("vp"):
            s.add(mn["vp"])
        return s or {"(sin VP)"}

    # Registros agrupados por VP (incluye pseudo-registros Forecast/Ppto2027: llevan CECO).
    por_vp = {}
    for r in V["records"]:
        for vp in vpsOf(r["ceco"]):
            por_vp.setdefault(vp, []).append(r)

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
        cecos = {r["ceco"] for r in recs}
        # Cada mapa de estructura se filtra por SU PROPIA VP (independiente del otro),
        # así el toggle Antigua/Nueva muestra en cada caso justo los CECOs de esta VP.
        cecosOld_vp = {c for c, m in cecoOld.items() if m.get("vp") == vp}
        cecosNew_vp = {c for c, m in cecoNew.items() if m.get("vp") == vp}

        Vf = dict(V)
        Vf["records"] = recs
        Vf["cecoOld"] = {c: m for c, m in cecoOld.items() if c in cecosOld_vp}
        Vf["cecoNew"] = {c: m for c, m in cecoNew.items() if c in cecosNew_vp}

        # DETALLE: solo líneas de los CECOs de esta VP, re-internando strings (archivo liviano).
        detf = None
        if DET:
            S = DET["s"]
            newS, idx = [], {}
            def I(x):
                k = idx.get(x)
                if k is None:
                    k = len(newS); idx[x] = k; newS.append(x)
                return k
            K = {}
            for key, rows in DET["k"].items():
                fr = [[I(S[row[0]]), I(S[row[1]]), I(S[row[2]]), I(S[row[3]]), row[4], row[5], row[6], row[7]]
                      for row in rows if S[row[0]] in cecos]
                if fr:
                    K[key] = fr
            detf = {"s": newS, "k": K}

        # DETALLE Ppto/Forecast (Concepto Gasto › Actividad): filas [cecoIdx,cgIdx,actIdx,meas,n,a].
        detpf = None
        if DETP:
            Sp = DETP["s"]
            newSp, idxp = [], {}
            def Ip(x):
                k = idxp.get(x)
                if k is None:
                    k = len(newSp); idxp[x] = k; newSp.append(x)
                return k
            Kp = {}
            for key, rows in DETP["k"].items():
                fr = [[Ip(Sp[row[0]]), Ip(Sp[row[1]]), Ip(Sp[row[2]]), row[3], row[4], row[5]]
                      for row in rows if Sp[row[0]] in cecos]
                if fr:
                    Kp[key] = fr
            detpf = {"s": newSp, "k": Kp}

        parts = ["window.V2_DATA = " + _j(Vf) + ";", 'window.DOT_DATA = {"records":[]};']
        if LOGO:
            parts.append("window.CORP_LOGO = " + _j(LOGO) + ";")
        if detf is not None:
            parts.append("window.DET = " + _j(detf) + ";")
        if detpf is not None:
            parts.append("window.DETP = " + _j(detpf) + ";")

        out = B._set_asset(html, B.DATA_UUID, "\n".join(parts))
        # Título de la VP en el H1 (TEXTOS de model.js) y en la pestaña del navegador.
        model_vp = (model
                    .replace('"titulo": "Actividad Corporativa"', '"titulo": ' + _j(str(vp)))
                    .replace('"tituloPlus": "+ Distribuibles Ppto 2027"', '"tituloPlus": ""'))
        out = B._set_asset(out, MODEL_UUID, model_vp)
        out = re.sub(r"<title>.*?</title>", "<title>" + _safe(vp) + " · AMSA</title>", out, count=1)

        vpdir = os.path.join(OUT_DIR, _safe(vp))          # una carpeta por VP
        os.makedirs(vpdir, exist_ok=True)
        fn = os.path.join(vpdir, "Dashboard " + _safe(vp) + ".html")
        open(fn, "w", encoding="utf-8", newline="").write(out)
        generados.append((vp, len(recs), len(cecosOld_vp | cecosNew_vp), os.path.getsize(fn) / 1e6))

    B.log("  --- Dashboards por VP generados ---")
    for vp, nr, nc, mb in generados:
        B.log(f"    {vp:42} {nr:5} regs · {nc:3} CECOs · {mb:4.1f} MB")
    B.log(f"\n✓ LISTO. {len(generados)} archivos en {OUT_DIR}")


if __name__ == "__main__":
    main()
