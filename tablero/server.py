# -*- coding: utf-8 -*-
"""
server.py — Panel de Control Web para el Dashboard Corporativo (Presupuesto 2027).
==================================================================================
Un front en el navegador (como el del Reporte Semanal) para, sin recordar comandos:
  · Actualizar Presupuesto 2027 / Forecast 5+7 / Outlook 6+6
  · Preparar un ejercicio nuevo (hoja «… Unpivot»)
  · Agregar ajustes/reversas al Forecast 5+7
  · Publicar: Local · Local (completo + por VP) · OneDrive (solo completo) · OneDrive (completo + por VP)

No requiere Flask: usa solo la librería estándar de Python.

Uso:  python server.py      →  abre http://localhost:5050
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
CV   = os.path.join(HERE, "construir.py")
UP   = os.path.join(HERE, "uploads")
REV_DIR   = os.path.join(UP, "ajustes_forecast")            # ajustes (reversas) al Forecast 5+7
REV_OFF    = os.path.join(REV_DIR, "_descartados")          # subcarpeta: ajustes desactivados (no se leen)
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


# ── Referencias a las constantes de construir.py ─────────────────────────────
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


# ── Corte YTD 2026 (último mes incluido en Real y Ppto YTD) ──────────────────
def leer_ytd_mes():
    txt = open(CV, encoding="utf-8").read()
    m = re.search(r'^YTD_HASTA_MES\s*=\s*(\d{1,2})', txt, re.M)
    return int(m.group(1)) if m else None


def set_ytd_mes(mes):
    mes = int(mes)
    if not (1 <= mes <= 12):
        raise RuntimeError("mes fuera de rango (1..12): %r" % mes)
    txt = open(CV, encoding="utf-8").read()
    new, n = re.subn(r'(^YTD_HASTA_MES\s*=\s*)\d{1,2}', lambda mm: mm.group(1) + str(mes), txt, flags=re.M)
    if n != 1:
        raise RuntimeError("esperaba 1 reemplazo de YTD_HASTA_MES, hubo %d" % n)
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


def accion_agregar_real(filename, anio=None):
    ruta = os.path.join(UP, "real_total", filename)
    cmd = [PY, "agregar_real.py", ruta]
    if anio:
        cmd += ["--anio", str(anio)]
    titulo = "Ajustando a 2027 y agregando al maestro de Reales" + \
             (" (año %s)" % anio if anio else " (año del archivo)")
    started = run_cmds([(titulo, cmd)],
                       fin_msg="Real agregado. Reconstruí/publicá para verlo en el dashboard.")
    return {"running": started}


def accion_publicar(modo):
    if modo == "1":
        cmds = [("Reconstruyendo el dashboard local (previsualizacion)", [PY, "construir.py", "--rapido"])]
        fin = "Previsualizacion lista en salida/."
    elif modo == "1b":
        cmds = [("Reconstruyendo el dashboard local", [PY, "construir.py", "--rapido"]),
                ("Dividiendo el dashboard por VP (local)", [PY, "dividir_por_vp.py"])]
        fin = "Dashboard local + por VP generados en salida/ (sin OneDrive)."
    elif modo == "2":
        cmds = [("Reconstruyendo el dashboard completo", [PY, "construir.py", "--rapido"]),
                ("Copiando el completo a OneDrive",
                 [PY, "-c", "import actualizar as a; a._copiar_dashboard_extra()"])]
        fin = "Dashboard completo publicado en OneDrive."
    else:
        cmds = [("Reconstruyendo completo + dashboards por VP", [PY, "actualizar.py"]),
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
    return {"refs": refs, "archivos": archivos, "ytdMes": leer_ytd_mes(),
            "ajustes": listar_ajustes(), "running": RUNNING}


def accion_ytd_mes(mes):
    set_ytd_mes(mes)
    nombres = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio",
               "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    log("\n✔ Corte YTD 2026 → hasta %s (mes %d). Reconstruye/publica para aplicarlo.\n"
        % (nombres[int(mes)], int(mes)))
    return {"running": False}


# ── Ajustes (reversas) al Forecast 5+7: activar / desactivar / eliminar ──────
def _xlsx_en(d):
    if not os.path.isdir(d):
        return []
    fs = [f for f in os.listdir(d)
          if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    fs.sort(key=lambda f: f.lower())
    return fs


def listar_ajustes():
    # Activos = los que están directamente en ajustes_forecast/ (el build los lee).
    # Descartados = los movidos a la subcarpeta _descartados/ (el build los ignora).
    return {"activos": _xlsx_en(REV_DIR), "descartados": _xlsx_en(REV_OFF)}


def accion_ajuste(nombre, accion):
    nombre = os.path.basename(nombre or "")   # sin rutas: solo el nombre del archivo
    if not nombre:
        return {"error": "falta el nombre del archivo", **listar_ajustes()}
    if accion in ("desactivar", "reactivar"):
        src_dir, dst_dir = (REV_DIR, REV_OFF) if accion == "desactivar" else (REV_OFF, REV_DIR)
        src, dst = os.path.join(src_dir, nombre), os.path.join(dst_dir, nombre)
        if not os.path.isfile(src):
            return {"error": "no existe %s" % nombre, **listar_ajustes()}
        os.makedirs(dst_dir, exist_ok=True)
        shutil.move(src, dst)
        verbo = "desactivado (movido a _descartados)" if accion == "desactivar" else "reactivado"
        log("\n✔ Ajuste %s: %s. Reconstruye/publica para aplicarlo.\n" % (verbo, nombre))
    elif accion == "eliminar":
        ruta = os.path.join(REV_OFF, nombre)   # solo se elimina definitivamente desde _descartados
        if not os.path.isfile(ruta):
            return {"error": "solo se elimina un ajuste ya desactivado", **listar_ajustes()}
        os.remove(ruta)
        log("\n✔ Ajuste eliminado definitivamente: %s\n" % nombre)
    else:
        return {"error": "acción desconocida: %s" % accion, **listar_ajustes()}
    return {"running": False, **listar_ajustes()}


def accion_consolidar_ajustes():
    started = run_cmds([("Generando el Excel consolidado de ajustes",
                         [PY, "revisar_ajustes.py", "--excel"])],
                       fin_msg="Consolidado escrito en salida/ajustes_forecast/.")
    return {"running": started}


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
                return self._send(accion_agregar_real(d["filename"], d.get("anio")))
            if self.path == "/api/ytd-mes":
                return self._send(accion_ytd_mes(d["mes"]))
            if self.path == "/api/ajuste":
                return self._send(accion_ajuste(d.get("nombre"), d.get("accion")))
            if self.path == "/api/consolidar-ajustes":
                return self._send(accion_consolidar_ajustes())
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
