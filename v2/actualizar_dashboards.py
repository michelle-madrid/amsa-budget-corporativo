# -*- coding: utf-8 -*-
"""
actualizar_dashboards.py — Actualiza en UNA sola corrida el Dashboard v3 y los dashboards por VP.
=================================================================================================
Un único proceso Python que hace, en orden:
  1) Reconstruye SOLO v3 con el Ppto/CECOS/Reales vigentes (v2 queda CONGELADO).
  2) Regenera los dashboards por VP (y sus subcarpetas por Gerencia) a partir del v3 recién hecho.

NO toca los Excel de apoyo (bbdd_a_excel) ni publica a OneDrive (para eso: python publicar.py).
Para cambiar el Ppto, editá PPTO27_FILE en construir_v2.py y volvé a correr esto.

Uso:  python actualizar_dashboards.py
"""
import hashlib
import os
import shutil

import construir_v2
import dividir_por_vp

# Baseline conocido de v2 (debe quedar intacto tras la corrida).
_V2_BASE = {
    "Dashboard Corporativo v2.html": "0c2bf7beff38e5eb13650720ceba8d14",
    "data_v2.js": "b47f9e3a307ab406043e595f21ac02c7",
}

# Copias EXTRA del Dashboard v3 (además de v3/Dashboard Corporativo v3.html en el repo).
# Se copia el HTML autocontenido a cada carpeta de esta lista (agregá más si hace falta).
EXTRA_DESTS = [
    (r"C:\Users\amsmmadrida\Grupo Minero Antofagasta Minerals"
     r"\Gcia. Planificación y Control de Gestión - Disco Gestión"
     r"\01 Control de Gestión\20 Mejoras Reportes\Propuesta Dashboard Ppto 27"),
]


def _md5(p):
    return hashlib.md5(open(p, "rb").read()).hexdigest() if os.path.isfile(p) else "(no existe)"


def _long(p):
    """Prefijo \\?\ para rutas largas de Windows (>260 chars)."""
    return p if p.startswith("\\\\?\\") else "\\\\?\\" + os.path.abspath(p)


def _copiar_v3_extra():
    """Copia el Dashboard v3 recién generado a cada carpeta de EXTRA_DESTS (si existe)."""
    src = construir_v2.V3_HTML
    for d in EXTRA_DESTS:
        if not os.path.isdir(_long(d)):
            print(f"  OJO: no existe la carpeta, NO copié: {d}")
            continue
        shutil.copy2(src, _long(os.path.join(d, os.path.basename(src))))
        print(f"  v3 copiado a: {d}")


def main():
    here = os.path.dirname(os.path.abspath(__file__))

    print("=" * 72)
    print("PASO 1/3  ·  Reconstruyendo Dashboard v3 (solo v3; v2 se mantiene congelado)")
    print(f"           Ppto 2027: {os.path.basename(construir_v2.PPTO27_FILE)}")
    print("=" * 72)
    construir_v2.main(solo_v3=True)

    print("\n" + "=" * 72)
    print("PASO 2/3  ·  Copiando Dashboard v3 a las carpetas extra")
    print("=" * 72)
    _copiar_v3_extra()

    print("\n" + "=" * 72)
    print("PASO 3/3  ·  Regenerando dashboards por VP (y subcarpetas por Gerencia)")
    print("=" * 72)
    dividir_por_vp.main()

    # Chequeo de que v2 no se movió.
    print("\n" + "-" * 72)
    ok = True
    for f, base in _V2_BASE.items():
        h = _md5(os.path.join(here, f))
        estado = "OK (congelado)" if h == base else "CAMBIO!"
        if h != base:
            ok = False
        print(f"  v2 · {f:32} {estado}")
    print("-" * 72)
    print("\nTODO LISTO: Dashboard v3 y dashboards por VP actualizados con el Ppto vigente."
          + ("" if ok else "  *** ATENCION: v2 cambio, revisar. ***"))
    print("(Excel de apoyo intactos. Para publicar a OneDrive: python publicar.py)")


if __name__ == "__main__":
    main()
