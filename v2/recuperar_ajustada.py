# -*- coding: utf-8 -*-
"""
recuperar_ajustada.py — Repara el maestro Corp cuya columna Ajustada quedó vacía.
=================================================================================
Causa: la columna «Valor mes 2027 (USD)» del maestro eran FÓRMULAS con caché.
Al agregar filas con openpyxl (load+save) se perdió el caché → quedó vacía.

Este script:
  1) Restaura el maestro desde el respaldo (Reales Históricos v2.BAK.xlsx).
  2) Lo carga con data_only=True → APLANA las fórmulas a sus valores (los conserva).
  3) Re-agrega los 3 datasets de real_total (como valores, sin fórmulas).
  4) Guarda una sola vez.

Uso:  python v2/recuperar_ajustada.py
"""
import os
import sys
import shutil
import warnings

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import agregar_real as A  # noqa: E402

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BAK = os.path.join(A.UP, "Reales Históricos v2.BAK.xlsx")
FILES = [os.path.join(A.UP, "real_total", f) for f in
         ["Real_1000AC3510.xlsx", "Oficina Londres Real.xlsx", "Exploraciones Real.xlsx"]]


def main():
    if not os.path.isfile(BAK):
        raise SystemExit("No existe el respaldo: %s" % BAK)

    print("1) Restaurando maestro desde el respaldo…")
    shutil.copy2(BAK, A.CORP_MASTER)

    print("2) Cargando maestro (aplana fórmulas → valores)…")
    wb = openpyxl.load_workbook(A.CORP_MASTER, data_only=True)
    ws = wb[A.MASTER_SHEET]
    existentes = set(str(c[0]) for c in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True) if c[0] is not None)
    print("   %d CECOs en el maestro restaurado" % len(existentes))

    print("3) Recalculando y re-agregando los 3 datasets…")
    vec = A.cargar_vector(A.VECTOR_XLSX, A.VECTOR_SHEET)
    total = 0
    for f in FILES:
        print("  ·", os.path.basename(f))
        df = A.leer_y_calcular(f, vec)
        pref = df[A.COL_CECO].astype(str).str[:4]
        corp = df[~pref.isin(A.DIST_PREFIXES)]                 # los 3 son corporativos
        nuevos = corp[~corp[A.COL_CECO].astype(str).isin(existentes)]
        for _, row in nuevos.iterrows():
            ws.append([A._val(k, key, row) for k, key in A.MAP_CORP])
        existentes |= set(df[A.COL_CECO].astype(str))
        total += len(nuevos)
        print("    + %d filas" % len(nuevos))

    print("4) Guardando maestro… (+%d filas)" % total)
    wb.save(A.CORP_MASTER)
    wb.close()
    print("LISTO.")


if __name__ == "__main__":
    main()
