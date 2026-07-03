# -*- coding: utf-8 -*-
"""
agregar_cecos_c3.py — Agrega los CECOs de 'CECOS C3.xlsx' al maestro CECOS.xlsx
==============================================================================
'uploads/CECOS C3.xlsx' (hoja 'C3 MLP') trae CECOs No Operacionales (Tipo Costo C3)
de las compañías distribuibles con su jerarquía Nodo1 / Nodo2 (VP) / Nodo3 (Gerencia).
Este script agrega a la hoja 'CECOS Corporativo' de 'uploads/CECOS.xlsx' los CECOs
que aún NO están (idempotente: no duplica). Deja respaldo .bak.

Mapeo de columnas  C3  →  CECOS Corporativo:
  CECO(6)→CECO · DescCECO(7)→Desc.CECO y Desc.CECO(Original)
  Nodo1(0/1)→Nodo1/Desc.Nodo1 · Nodo2(2/3)→Nodo2/Desc.Nodo2(VP) · Nodo3(4/5)→Nodo3/Desc.Nodo3(Gerencia)
Campos que C3 no trae (defaults, editables en el Excel):
  Resumen='No Operacional (C3)' · Nodo Compañías='No Operacional'
  Clasificación del Gasto='Gastos Distribuibles' · Tipo Costo='C3' · VP Abreviada/¿Aplica?/N_CECO vacíos

Uso:  python agregar_cecos_c3.py
"""
import os
import shutil
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
UP = os.path.join(HERE, "uploads")
MAESTRO = os.path.join(UP, "CECOS.xlsx")
C3_FILE = os.path.join(UP, "CECOS C3.xlsx")

RESUMEN_C3 = "No Operacional (C3)"
NODO_COMP_C3 = "No Operacional"
CLASIF_C3 = "Gastos Distribuibles"
TIPO_COSTO_C3 = "C3"


def _norm(v):
    return str(v).strip().lower() if v is not None else ""


def leer_c3(path=C3_FILE):
    """Devuelve [(nodo1,dn1,nodo2,dn2vp,nodo3,dn3ger,ceco,descceco)] desde C3 MLP."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):   # fila 1 = encabezado (1..9)
        ceco = r[6] if len(r) > 6 else None
        if ceco is None:
            continue
        g = lambda i: (str(r[i]).strip() if i < len(r) and r[i] is not None else "")
        out.append((g(0), g(1), g(2), g(3), g(4), g(5), str(ceco).strip(), g(7)))
    wb.close()
    return out


def main():
    if not os.path.exists(MAESTRO):
        raise SystemExit("No existe el maestro: " + MAESTRO)
    if not os.path.exists(C3_FILE):
        raise SystemExit("No existe: " + C3_FILE)

    c3 = leer_c3()
    print(f"CECOS C3: {len(c3)} filas.")

    wb = openpyxl.load_workbook(MAESTRO)   # conserva las demás hojas/estilos
    sheet = next((s for s in wb.sheetnames if "orporativo" in s.lower()), None)
    if sheet is None:
        raise SystemExit(f"No encontré la hoja 'CECOS Corporativo'. Hojas: {wb.sheetnames}")
    ws = wb[sheet]

    hdr = [c.value for c in ws[1]]
    ncols = len(hdr)
    idx = {_norm(c): i for i, c in enumerate(hdr)}
    ci_ceco = idx.get("ceco")
    if ci_ceco is None:
        raise SystemExit(f"No encontré la columna CECO. Encabezado: {hdr}")

    existentes = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if ci_ceco < len(r) and r[ci_ceco] is not None:
            existentes.add(str(r[ci_ceco]).strip())
    print(f"Maestro '{sheet}': {len(existentes)} CECOs existentes.")

    # Construye una fila del maestro (por nombre de columna) para un CECO de C3.
    def fila(n1, dn1, n2, dn2, n3, dn3, ceco, descceco):
        row = [None] * ncols
        def put(col, val):
            i = idx.get(_norm(col))
            if i is not None:
                row[i] = val
        put("resumen", RESUMEN_C3)
        put("nodo compañías", NODO_COMP_C3)
        # División = Nodo 1 del C3 (ej. 'Los Pelambres No Operacional', 'Comercialización').
        put("clasificación del gasto", dn1 or CLASIF_C3)
        put("ceco", ceco)
        put("desc. ceco", descceco)
        put("desc. ceco (original)", descceco)
        put("nodo 1", n1)
        put("desc. nodo 1", dn1)
        put("nodo 2", n2)
        put("desc. nodo 2 (vp)", dn2)
        put("nodo 3", n3)
        put("desc. nodo 3 (gerencia)", dn3)
        put("tipo costo", TIPO_COSTO_C3)
        return row

    nuevos, saltados = [], 0
    for (n1, dn1, n2, dn2, n3, dn3, ceco, descceco) in c3:
        if ceco in existentes:
            saltados += 1
            continue
        nuevos.append(fila(n1, dn1, n2, dn2, n3, dn3, ceco, descceco))
        existentes.add(ceco)

    if not nuevos:
        print("Nada que agregar: todos los CECOs de C3 ya están en el maestro.")
        return

    shutil.copyfile(MAESTRO, MAESTRO + ".bak")
    print("Respaldo: " + os.path.basename(MAESTRO) + ".bak")
    for row in nuevos:
        ws.append(row)
    wb.save(MAESTRO)
    print(f"Agregados {len(nuevos)} CECOs nuevos a '{sheet}' (saltados {saltados} ya existentes).")
    print("Muestra:", [r[ci_ceco] for r in nuevos[:8]])


if __name__ == "__main__":
    main()
