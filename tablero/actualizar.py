# -*- coding: utf-8 -*-
"""
actualizar.py — Actualiza en UNA sola corrida el dashboard y los dashboards por VP.
==================================================================================
Un único proceso Python que hace, en orden:
  1) Reconstruye el dashboard con el Ppto / CECOS / Reales vigentes.
  2) Copia el dashboard a las carpetas extra (EXTRA_DESTS).
  3) Regenera los dashboards por VP (y sus subcarpetas por Gerencia) a partir del recién hecho.

NO toca los Excel de apoyo (bbdd_a_excel) ni publica a OneDrive (para eso: python publicar.py).
Para cambiar el Ppto, editá PPTO27_FILE en construir.py y volvé a correr esto.

Uso:  python actualizar.py
"""
import os
import shutil

import construir
import dividir_por_vp

# Copias EXTRA del dashboard (además del que queda en salida/Dashboard Corporativo.html).
# Se copia el HTML autocontenido a cada carpeta de esta lista. Para sumar un destino, agregá su ruta.
EXTRA_DESTS = [
    (r"C:\Users\amsmmadrida\Grupo Minero Antofagasta Minerals"
     r"\Gcia. Planificación y Control de Gestión - Disco Gestión"
     r"\01 Control de Gestión\20 Mejoras Reportes\Propuesta Dashboard Ppto 27"),
]


def _long(p):
    """Prefijo \\?\ para rutas largas de Windows (>260 chars)."""
    return p if p.startswith("\\\\?\\") else "\\\\?\\" + os.path.abspath(p)


def _copiar_dashboard_extra():
    """Copia el dashboard recién generado a cada carpeta de EXTRA_DESTS (si existe)."""
    src = construir.SALIDA_HTML
    for d in EXTRA_DESTS:
        if not os.path.isdir(_long(d)):
            print(f"  OJO: no existe la carpeta, NO copié: {d}")
            continue
        shutil.copy2(src, _long(os.path.join(d, os.path.basename(src))))
        print(f"  Dashboard copiado a: {d}")


def main():
    print("=" * 72)
    print("PASO 1/3  ·  Reconstruyendo el dashboard")
    print(f"           Ppto 2027: {os.path.basename(construir.PPTO27_FILE)}")
    print("=" * 72)
    construir.main(rapido=True)

    print("\n" + "=" * 72)
    print("PASO 2/3  ·  Copiando el dashboard a las carpetas extra")
    print("=" * 72)
    _copiar_dashboard_extra()

    print("\n" + "=" * 72)
    print("PASO 3/3  ·  Regenerando dashboards por VP (y subcarpetas por Gerencia)")
    print("=" * 72)
    dividir_por_vp.main()

    print("\n" + "-" * 72)
    print("TODO LISTO: dashboard y dashboards por VP actualizados con el Ppto vigente.")
    print("(Excel de apoyo intactos. Para publicar a OneDrive: python publicar.py)")


if __name__ == "__main__":
    main()
