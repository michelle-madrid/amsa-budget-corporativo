# -*- coding: utf-8 -*-
"""
server.py — Panel de Control Web para el Dashboard Corporativo (Presupuesto 2027).
==================================================================================
Un front en el navegador (como el del Reporte Semanal) para, sin recordar comandos:
  · Actualizar Presupuesto 2027 / Forecast 5+7 / Outlook 6+6
  · Preparar un ejercicio nuevo (hoja «… Unpivot»)
  · Agregar ajustes/reversas al Forecast 5+7
  · Publicar: Local · OneDrive (solo completo) · OneDrive (completo + por VP)

No requiere Flask: usa solo la librería estándar de Python.

Uso:  python v2/server.py      →  abre http://localhost:5000
"""
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
CV   = os.path.join(HERE, "construir_v2.py")
UP   = os.path.join(HERE, "uploads")
PY   = sys.executable
PORT = 5050   # 5000 suele estar ocupado por el panel del Reporte Semanal

# tipo → carpeta de uploads y (si aplica) constante + hoja Unpivot
TIPOS = {
    "ppto":       {"folder": "presupuesto_2027", "var": "PPTO27_FILE"},
    "forecast":   {"folder": "ejercicios_2026",  "var": "FCST_FILE", "salida": "Forecast 5+7 Unpivot"},
    "outlook":    {"folder": "ejercicios_2026",  "var": "OUT_FILE",  "salida": "Outlook 6+6 Unpivot"},
    "ejercicios": {"folder": "ejercicios_2026"},
    "reversas":   {"folder": "ajustes_forecast"},
    "real":       {"folder": "real_total"},
}

# ── Estado / logs ────────────────────────────────────────────────────────────
LOGS = []
LOCK = threading.Lock()
RUNNING = False


def log(s):
    with LOCK:
        LOGS.append(s)


# ── Referencias en construir_v2.py ───────────────────────────────────────────
def leer_constante(var):
    txt = open(CV, encoding="utf-8").read()
    m = re.search(var + r'\s*=\s*os\.path\.join\(UP,\s*"([^"]+)",\s*"([^"]+)"\)', txt)
    return (m.group(1), m.group(2)) if m else (None, None)


def set_constante(var, filename):
    txt = open(CV, encoding="utf-8").read()
    pat = re.compile(r'(' + var + r'\s*=\s*os\.path\.join\(UP,\s*"[^"]+",\s*")([^"]+)(")')
    new, n = pat.subn(lambda mm: mm.group(1) + filename + mm.group(3), txt)
    if n != 1:
        raise RuntimeError("esperaba 1 reemplazo de %s, hubo %d" % (var, n))
    open(CV, "w", encoding="utf-8").write(new)


def listar_archivos(folder):
    d = os.path.join(UP, folder)
    if not os.path.isdir(d):
        return []
    fs = [f for f in os.listdir(d)
          if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    fs.sort(key=lambda f: os.path.getmtime(os.path.join(d, f)), reverse=True)
    return fs


# ── Ejecución de comandos (hilo, con captura de salida) ──────────────────────
def run_cmds(cmds, on_success=None, fin_msg=""):
    """cmds = [(titulo, [comando]), …]. Corre en un hilo y vuelca la salida a LOGS."""
    global RUNNING
    if RUNNING:
        return False
    RUNNING = True

    def worker():
        global RUNNING
        ok = True
        try:
            # Hijo sin buffer (salida en vivo) y en UTF-8 (para que tildes/caracteres
            # especiales no salgan como «�» al leerlos aquí como UTF-8).
            env = {**os.environ, "PYTHONUNBUFFERED": "1", "PYTHONIOENCODING": "utf-8"}
            for titulo, cmd in cmds:
                log("\n▶ " + titulo + "\n")
                p = subprocess.Popen(cmd, cwd=HERE, stdout=subprocess.PIPE,
                                     stderr=subprocess.STDOUT, text=True,
                                     encoding="utf-8", errors="replace", bufsize=1, env=env)
                for line in p.stdout:
                    log(line)
                p.wait()
                if p.returncode != 0:
                    log("[ERROR] codigo %d\n" % p.returncode)
                    ok = False
                    break
                log("[OK]\n")
            if ok and on_success:
                on_success()
            if ok and fin_msg:
                log("\n✔ " + fin_msg + "\n")
        except Exception as e:
            log("[ERROR] %s\n" % e)
        finally:
            RUNNING = False

    threading.Thread(target=worker, daemon=True).start()
    return True


def browse_nativo(folder):
    """Abre un diálogo nativo (subproceso tkinter) y devuelve la ruta elegida."""
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "r=tk.Tk(); r.withdraw(); r.wm_attributes('-topmost',1); "
        "print(filedialog.askopenfilename(title='Elegir Excel', initialdir=%r, "
        "filetypes=[('Excel','*.xlsx *.xlsm *.XLSX'),('Todos','*.*')]), end='')"
        % os.path.join(UP, folder)
    )
    r = subprocess.run([PY, "-c", script], capture_output=True, text=True)
    return r.stdout.strip()


# ── Acciones ─────────────────────────────────────────────────────────────────
def accion_actualizar_ref(tipo, filename):
    cfg = TIPOS[tipo]
    var = cfg["var"]
    if "salida" not in cfg:                      # Ppto: solo cambia la referencia
        set_constante(var, filename)
        log("\n✔ Presupuesto 2027  ->  %s\n" % filename)
        return {"running": False}
    # Forecast / Outlook: genera la hoja Unpivot y luego apunta la referencia
    ruta = os.path.join(UP, cfg["folder"], filename)
    started = run_cmds(
        [("Ajustando a 2027 y generando la hoja «%s»" % cfg["salida"],
          [PY, "unpivot_ejercicio.py", ruta, "--salida", cfg["salida"], "--anio-destino", "2027"])],
        on_success=lambda: set_constante(var, filename),
        fin_msg="%s  ->  %s" % (tipo.capitalize(), filename))
    return {"running": started}


def accion_ejercicio(filename, anio, salida):
    ruta = os.path.join(UP, "ejercicios_2026", filename)
    cmd = [PY, "unpivot_ejercicio.py", ruta, "--anio-destino", anio or "2027"]
    if salida:
        cmd += ["--salida", salida]
    started = run_cmds([("Generando la hoja Unpivot del ejercicio", cmd)],
                       fin_msg="Ejercicio preparado. Usalo como Forecast/Outlook o como reversa.")
    return {"running": started}


def accion_reversa(filename, anio, ya_tiene):
    if ya_tiene:
        log("\n✔ El ajuste ya tiene su hoja Unpivot; el build lo tomara.\n")
        return {"running": False}
    ruta = os.path.join(UP, "ajustes_forecast", filename)
    started = run_cmds([("Generando la hoja Unpivot del ajuste",
                        [PY, "unpivot_ejercicio.py", ruta, "--anio-destino", anio or "2027"])],
                       fin_msg="Ajuste listo; el build lo tomara automaticamente.")
    return {"running": started}


def accion_agregar_real(filename):
    ruta = os.path.join(UP, "real_total", filename)
    started = run_cmds([("Ajustando a 2027 y agregando al maestro de Reales",
                        [PY, "agregar_real.py", ruta])],
                       fin_msg="Real agregado. Reconstruí/publicá para verlo en el dashboard.")
    return {"running": started}


def accion_publicar(modo):
    if modo == "1":
        cmds = [("Reconstruyendo v3 local (previsualizacion)", [PY, "construir_v2.py", "--solo-v3"])]
        fin = "Previsualizacion lista en v3/."
    elif modo == "2":
        cmds = [("Reconstruyendo el dashboard completo", [PY, "construir_v2.py", "--solo-v3"]),
                ("Copiando el completo a OneDrive",
                 [PY, "-c", "import actualizar_dashboards as a; a._copiar_v3_extra()"])]
        fin = "Dashboard completo publicado en OneDrive."
    else:
        cmds = [("Reconstruyendo completo + dashboards por VP", [PY, "actualizar_dashboards.py"]),
                ("Publicando por VP a OneDrive", [PY, "publicar.py"])]
        fin = "Completo + por VP publicados en OneDrive."
    started = run_cmds(cmds, fin_msg=fin)
    return {"running": started}


def estado():
    refs = {}
    for tipo in ("ppto", "forecast", "outlook"):
        _, fn = leer_constante(TIPOS[tipo]["var"])
        refs[tipo] = fn
    archivos = {t: listar_archivos(TIPOS[t]["folder"]) for t in TIPOS}
    return {"refs": refs, "archivos": archivos, "running": RUNNING}


# ── HTTP ─────────────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def _send(self, obj, ctype="application/json"):
        body = obj if isinstance(obj, bytes) else json.dumps(obj).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if "json" in ctype or "html" in ctype else ""))
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(n) or b"{}") if n else {}

    def log_message(self, *a):
        pass  # silencio

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/index"):
            with open(os.path.join(HERE, "panel_dashboard.html"), "rb") as f:
                self._send(f.read(), "text/html")
        elif self.path == "/api/estado":
            self._send(estado())
        elif self.path.startswith("/api/logs"):
            from urllib.parse import urlparse, parse_qs
            off = int(parse_qs(urlparse(self.path).query).get("offset", [0])[0])
            with LOCK:
                chunk = LOGS[off:]
            self._send({"lines": chunk, "next_offset": off + len(chunk), "running": RUNNING})
        else:
            self.send_error(404)

    def do_POST(self):
        try:
            d = self._body()
            if self.path == "/api/browse":
                ruta = browse_nativo(TIPOS[d["tipo"]]["folder"])
                if not ruta:
                    return self._send({"filename": ""})
                folder = os.path.join(UP, TIPOS[d["tipo"]]["folder"])
                if os.path.dirname(os.path.abspath(ruta)).lower() != os.path.abspath(folder).lower():
                    shutil.copy2(ruta, os.path.join(folder, os.path.basename(ruta)))
                    log("  copiado a %s\n" % TIPOS[d["tipo"]]["folder"])
                return self._send({"filename": os.path.basename(ruta)})
            if self.path == "/api/actualizar-ref":
                return self._send(accion_actualizar_ref(d["tipo"], d["filename"]))
            if self.path == "/api/ejercicio":
                return self._send(accion_ejercicio(d["filename"], d.get("anio"), d.get("salida")))
            if self.path == "/api/reversa":
                return self._send(accion_reversa(d["filename"], d.get("anio"), d.get("ya_tiene")))
            if self.path == "/api/agregar-real":
                return self._send(accion_agregar_real(d["filename"]))
            if self.path == "/api/publicar":
                return self._send(accion_publicar(d["modo"]))
            if self.path == "/api/logs/clear":
                with LOCK:
                    LOGS.clear()
                return self._send({"ok": True})
            self.send_error(404)
        except Exception as e:
            self._send({"error": str(e)})


def main():
    # Busca un puerto libre desde PORT (por si 5050 también está ocupado).
    srv = None
    for p in range(PORT, PORT + 10):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            puerto = p
            break
        except OSError:
            continue
    if srv is None:
        print("No encontré un puerto libre entre %d y %d." % (PORT, PORT + 9))
        return
    url = "http://localhost:%d" % puerto
    print("=" * 60)
    print("  Panel Dashboard Corporativo  ->  " + url)
    print("  (Ctrl+C para cerrar)")
    print("=" * 60)
    try:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    except Exception:
        pass
    srv.serve_forever()


if __name__ == "__main__":
    main()
