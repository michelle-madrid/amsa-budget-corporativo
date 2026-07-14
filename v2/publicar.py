# -*- coding: utf-8 -*-
"""
publicar.py — Publica los dashboards por VP a la carpeta compartida (OneDrive/SharePoint).
============================================================================================
Corré esto DESPUÉS de dividir_por_vp.py. Copia SOLO los HTML (nunca los Excel).

Reglas (pedidas):
  · Solo actualiza las carpetas VP que YA EXISTEN en el destino (no crea VP nuevas).
  · REEMPLAZA el/los Dashboard*.html dentro; NO borra carpetas ni otros archivos del destino.
  · VP Finanzas (o cualquier VP con subcarpetas por Gerencia): también publica esas subcarpetas
    (crea las que falten, reemplaza el HTML; no borra).
  · Nunca copia los Excel (BBDD *.xlsx).

Uso:  python publicar.py
"""
import os
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "v3", "por_vp")
DEST = (r"C:\Users\amsmmadrida\Grupo Minero Antofagasta Minerals"
        r"\Gcia. Planificación y Control de Gestión - Disco Gestión"
        r"\02 Planificación Estratégica\01 Ciclo Planificación\Ciclo Planificación 2026"
        r"\04 Presupuesto 2027\06 Dashboard Histórico")


def _long(p):
    """Prefijo \\?\ para superar el límite de 260 chars de Windows (el destino es muy largo)."""
    if p.startswith("\\\\?\\"):
        return p
    return "\\\\?\\" + os.path.abspath(p)


def _copiar_dashboards(src_dir, dst_dir):
    """Copia solo Dashboard*.html de src_dir a dst_dir (reemplaza). No borra nada."""
    n = 0
    for f in os.listdir(src_dir):
        if f.startswith("Dashboard") and f.endswith(".html"):
            shutil.copy2(os.path.join(src_dir, f), _long(os.path.join(dst_dir, f)))
            n += 1
    return n


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
        for sub in sorted(os.listdir(src_vp)):
            src_sub = os.path.join(src_vp, sub)
            if os.path.isdir(src_sub):
                dst_sub = os.path.join(DEST, vp, sub)
                os.makedirs(_long(dst_sub), exist_ok=True)         # crea si falta (no borra)
                ng += _copiar_dashboards(src_sub, dst_sub)
        total += n + ng
        print(f"  - {vp}: {n} dashboard" + (f" + {ng} por Gerencia" if ng else ""))
    print(f"\nLISTO. {total} archivos publicados (solo HTML; sin borrar nada del destino).")


if __name__ == "__main__":
    main()
