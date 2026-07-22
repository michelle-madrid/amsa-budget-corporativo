# -*- coding: utf-8 -*-
"""
Consolida TODA la información de Reales del dashboard (Corp + Distribuibles) en
una única BBDD maestra en formato parquet (tabla de hechos, fila-a-fila, con las
columnas de ajuste a mes 2027 ya calculadas).

Es un consolidado/respaldo PARALELO: el dashboard sigue leyendo los Excel; este
parquet es la foto única de todas las transacciones reales.

Uso:  python v2/consolidar_reales_parquet.py
"""
import os
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

# ─────────────────────────── CONFIGURACIÓN ───────────────────────────
FUENTES = [
    ("Corp",          "v2/uploads/Reales Históricos v2.xlsx"),
    ("Distribuibles", "v2/uploads/Reales Distribuibles Histórico v2.xlsx"),
]
SHEET = "Data Consolidada"

OUT_DIR     = "v2/bbdd_maestra"
OUT_PARQUET = os.path.join(OUT_DIR, "reales_maestro.parquet")
# ──────────────────────────────────────────────────────────────────────


def main():
    partes = []
    for fuente, path in FUENTES:
        print(f"Leyendo {fuente}… {path} / {SHEET}")
        df = pd.read_excel(path, sheet_name=SHEET)
        df.insert(0, "Fuente", fuente)
        print(f"  {len(df):,} filas · {df.shape[1]} columnas")
        partes.append(df)

    print("Consolidando…")
    maestro = pd.concat(partes, ignore_index=True, sort=False)

    # Tipos seguros para parquet: columnas de texto/mixtas → string; numéricas se conservan.
    for c in maestro.columns:
        if maestro[c].dtype == "object":
            maestro[c] = maestro[c].apply(lambda v: "" if pd.isna(v) else str(v))

    os.makedirs(OUT_DIR, exist_ok=True)
    maestro.to_parquet(OUT_PARQUET, engine="pyarrow", index=False)

    size_mb = os.path.getsize(OUT_PARQUET) / 1e6
    print(f"\n[OK] BBDD maestra escrita: {OUT_PARQUET}  ({size_mb:.1f} MB)")
    print(f"  Total: {len(maestro):,} filas · {maestro.shape[1]} columnas")
    print("  Filas por fuente:")
    print(maestro["Fuente"].value_counts().to_string())
    print("\n  Columnas:", list(maestro.columns))
    print("\n  head():")
    with pd.option_context("display.max_columns", None, "display.width", 240):
        print(maestro.head(4).to_string(index=False))


if __name__ == "__main__":
    main()
