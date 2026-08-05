# -*- coding: utf-8 -*-
"""
publicar.py — Publica los dashboards por VP a la carpeta compartida (OneDrive/SharePoint).
============================================================================================
Corré esto DESPUÉS de dividir_por_vp.py. Copia SOLO los HTML (nunca los Excel).

Reglas (pedidas):
  · Solo actualiza las carpetas VP que YA EXISTEN en el destino (no crea VP nuevas, no las borra).
  · REEMPLAZA el/los Dashboard*.html dentro; no toca otros archivos del destino.
  · VP con subcarpetas por Gerencia (ej. VP Finanzas): también publica esas subcarpetas (crea las
    que falten). Y quita las subcarpetas de Gerencia OBSOLETAS (combos renombradas/fusionadas):
    las que ya no existen en el repo y son carpetas de dashboard nuestras (tienen Dashboard*.html).
  · Nunca copia los Excel (BBDD *.xlsx).

Uso:  python publicar.py
"""
import os
import stat
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "salida", "por_vp")
DEST = (r"C:\Users\amsmmadrida\Grupo Minero Antofagasta Minerals"
        r"\Gcia. Planificación y Control de Gestión - Disco Gestión"
        r"\02 Planificación Estratégica\01 Ciclo Planificación\Ciclo Planificación 2026"
        r"\04 Presupuesto 2027\06 Dashboard Histórico")


def _long(p):
    """Prefijo \\?\ para superar el límite de 260 chars de Windows (el destino es muy largo)."""
    if p.startswith("\\\\?\\"):
        return p
    return "\\\\?\\" + os.path.abspath(p)


def _rmtree_resiliente(path):
    """rmtree que primero limpia el bit read-only (Windows). Lanza si no lo logra."""
    def onerror(func, p, exc):
        os.chmod(p, stat.S_IWRITE)
        func(p)
    shutil.rmtree(path, onerror=onerror)


def _copiar_dashboards(src_dir, dst_dir):
    """Copia los Dashboard*.html de src_dir a dst_dir (reemplaza) y QUITA del destino los
    Dashboard*.html VIEJOS cuyo nombre ya no existe en el repo (renombres de Gerencia), para que
    cada carpeta quede con UN solo dashboard, el vigente. NUNCA toca carpetas ni otros archivos
    (Excel, etc.): solo elimina archivos 'Dashboard*.html' que quedaron obsoletos por un renombre."""
    src_dash = [f for f in os.listdir(src_dir) if f.startswith("Dashboard") and f.endswith(".html")]
    for f in src_dash:
        shutil.copy2(os.path.join(src_dir, f), _long(os.path.join(dst_dir, f)))
    src_set = set(src_dash)
    for f in os.listdir(_long(dst_dir)):
        if f.startswith("Dashboard") and f.endswith(".html") and f not in src_set:
            try:
                os.remove(_long(os.path.join(dst_dir, f)))
                print(f"      (quitado dashboard viejo por renombre: {f})")
            except Exception as e:
                print(f"      OJO: no pude quitar el dashboard viejo '{f}': {e}")
    return len(src_dash)


def main():
    if not os.path.isdir(SRC):
        raise SystemExit(f"No existe {SRC}. Corré antes: python dividir_por_vp.py")
    if not os.path.isdir(DEST):
        raise SystemExit(f"No existe el destino:\n  {DEST}")
    vps = sorted(d for d in os.listdir(_long(DEST)) if os.path.isdir(_long(os.path.join(DEST, d))))
    print(f"Publicando en:\n  {DEST}\n  ({len(vps)} carpetas VP en el destino)\n")
    total = 0
    for vp in vps:
        src_vp = os.path.join(SRC, vp)
        if not os.path.isdir(src_vp):
            print(f"  · {vp}: no se genera en el repo → se deja como está.")
            continue
        n = _copiar_dashboards(src_vp, os.path.join(DEST, vp))   # dashboard principal de la VP
        ng = 0                                                    # subcarpetas por Gerencia
        src_subs = sorted(s for s in os.listdir(src_vp) if os.path.isdir(os.path.join(src_vp, s)))
        for sub in src_subs:
            dst_sub = os.path.join(DEST, vp, sub)
            os.makedirs(_long(dst_sub), exist_ok=True)             # crea si falta (no borra)
            ng += _copiar_dashboards(os.path.join(src_vp, sub), dst_sub)
        # BORRADO DESACTIVADO (pedido de la usuaria, 2026-07-15): NUNCA se eliminan subcarpetas del
        # destino. Aunque una Gerencia se renombre en CECOS y su carpeta vieja quede huérfana, se
        # CONSERVA. publicar.py solo AGREGA/REEMPLAZA dashboards; jamás borra carpetas de OneDrive.
        rm = 0
        total += n + ng
        print(f"  - {vp}: {n} dashboard" + (f" + {ng} por Gerencia" if ng else "") + (f" · {rm} obsoleta(s) quitada(s)" if rm else ""))
    print(f"\nLISTO. {total} archivos publicados (solo HTML; sin borrar nada del destino).")


if __name__ == "__main__":
    main()
