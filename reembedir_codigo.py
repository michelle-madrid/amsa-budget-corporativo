# -*- coding: utf-8 -*-
"""
Reembebe el código fuente (.js/.jsx) dentro del HTML autocontenido
==================================================================
El dashboard embebe app.jsx/charts.jsx/model.js… como assets gzip+base64 en un
manifest. `construir_bbdd.py`/`actualizar_dashboard.py` SOLO cambian data.js; este
script vuelve a empaquetar el código cuando se edita un .jsx/.js.

Valida primero el JSX con el Babel embebido (vía Node) para no romper el bundle.
Uso:  python reembedir_codigo.py
"""
import os, re, gzip, base64, subprocess, tempfile, sys, zipfile, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "Dashboard Actividad Corporativa.html")

# Asset (UUID completo) → archivo fuente · jsx=True requiere transpilar con Babel.
# OJO: el mapa UUID→archivo correcto (verificado por los exports de cada asset):
#   0290c010 = filters.jsx   (MultiSelect/FilterBar/ViewTabs/KpiCards)
#   35b49f4c = tweaks-panel.jsx (useTweaks/TweaksPanel/…)
# (estaban cruzados en una nota vieja). El orden de carga en el HTML es:
#   data.js, model.js, react, react-dom, babel, 35b49f4c, charts, 0290c010, matrix, app.
ASSETS = {
    "b10e2126-5e07-4358-9daa-f1805ba1bc96": ("app.jsx",         True),
    "1a5bd7f6-3963-4301-955f-e692cf651ee7": ("charts.jsx",      True),
    "0290c010-dd92-483c-ac2e-8b4fe5d69eab": ("filters.jsx",     True),
    "35b49f4c-a3fb-4834-8681-376cfb28922f": ("tweaks-panel.jsx", True),
    "a1f4c70d-e147-4337-a85e-7a5b341a43c9": ("matrix.jsx",      True),
    "5c895c10-43c7-4f6c-aa15-ea39fb941007": ("model.js",        False),
}
BABEL_UUID = "bc7af47c-01a9-47f6-b802-09d216d56f10"


def _rx(uuid):
    return re.compile(r'("' + re.escape(uuid) +
                      r'":\{"mime":"[^"]*","compressed":true,"data":")([^"]*)(")')


def _get_asset(html, uuid):
    m = _rx(uuid).search(html)
    if not m:
        raise ValueError(f"No encontré el asset {uuid} en el HTML.")
    return gzip.decompress(base64.b64decode(m.group(2))).decode("utf-8")


def _set_asset(html, uuid, source):
    b64 = base64.b64encode(gzip.compress(source.encode("utf-8"), mtime=0)).decode("ascii")
    new, n = _rx(uuid).subn(lambda m: m.group(1) + b64 + m.group(3), html)
    if n != 1:
        raise ValueError(f"Esperaba reemplazar 1 asset {uuid}, encontré {n}.")
    return new


def _validar_jsx(html, fuentes):
    """Transpila cada JSX con el Babel embebido en Node; aborta si hay error."""
    if not fuentes:
        return
    babel = _get_asset(html, BABEL_UUID)
    with tempfile.TemporaryDirectory() as td:
        bpath = os.path.join(td, "babel.js")
        open(bpath, "w", encoding="utf-8").write(babel)
        for name in fuentes:
            src = open(os.path.join(HERE, name), encoding="utf-8").read()
            spath = os.path.join(td, "src.txt")
            open(spath, "w", encoding="utf-8").write(src)
            check = (
                "const Babel=require(%r);"
                "const fs=require('fs');"
                "const code=fs.readFileSync(%r,'utf8');"
                "try{Babel.transform(code,{presets:['react']});"
                "console.log('OK');}"
                "catch(e){console.error(String(e));process.exit(1);}"
                % (bpath, spath)
            )
            r = subprocess.run(["node", "-e", check], capture_output=True, text=True)
            if r.returncode != 0:
                raise ValueError(f"JSX inválido en {name}:\n{r.stderr or r.stdout}")
            print(f"  Babel OK: {name}")


def main():
    html = open(HTML, "r", encoding="utf-8", newline="").read()
    jsx = [n for u, (n, isjsx) in ASSETS.items() if isjsx]
    print("Validando JSX con Babel embebido…")
    _validar_jsx(html, jsx)
    for uuid, (name, _isjsx) in ASSETS.items():
        src = open(os.path.join(HERE, name), encoding="utf-8").read()
        html = _set_asset(html, uuid, src)
        print(f"  Reembebido: {name}  ({len(src)} chars)")
    open(HTML, "w", encoding="utf-8", newline="").write(html)
    print("HTML actualizado con el código reembebido.")
    # Rehacer el zip para que el entregable quede consistente con el código nuevo.
    zip_path = os.path.join(HERE, "Dashboard Actividad Corporativa.zip")
    readme = ("DASHBOARD - Actividad Corporativa + Distribuibles (AMSA)\r\n\r\n"
              "Doble clic en 'Dashboard Actividad Corporativa.html' (Chrome/Edge/Firefox).\r\n"
              f"\r\nActualizado: {datetime.date.today().strftime('%d-%m-%Y')}\r\n")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(HTML, arcname=os.path.basename(HTML))
        z.writestr("LEER - Como abrir.txt", readme)
    print("ZIP regenerado: " + os.path.basename(zip_path))


if __name__ == "__main__":
    main()
