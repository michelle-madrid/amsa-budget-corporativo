# -*- coding: utf-8 -*-
"""
distribuibles.py — Un dashboard + una BBDD Excel por compañía, solo con SUS Distribuibles.
==========================================================================================
Requiere haber corrido antes `python construir.py` (usa el dashboard ya generado y bbdd/).
Uso:  python distribuibles.py

Escribe en  salida/Distribuibles/<ABREV>/ :
  · Dashboard Distribuibles <Compañía>.html  → solo las pestañas «Gastos Corporativos» y
      «Tabla Resumen Gastos» (vía V2_DATA.soloTabs; CAPEX y Dotaciones se caen solas al
      reescribir el bloque de datos, igual que en los dashboards por VP). Además, con
      V2_DATA.soloDist se ocultan el selector Alcance y el filtro Compañía: no aplican
      cuando todo el archivo es distribuible de una sola compañía.
  · BBDD Distribuibles <Compañía>.xlsx       → UNA hoja «Base de Datos», línea a línea,
      Real y Ppto/Forecast juntos en el mismo formato (ver COLUMNAS más abajo).

ALCANCE: los CECOs cuya compañía (los 4 primeros dígitos del CECO) es la de esa compañía.
Coincide exactamente con la clasificación «Gastos Distribuibles» de CECOS.xlsx.
"""
import os
import re
import pandas as pd

import construir as B
import dividir_por_vp as D
import bbdd_a_excel as X
import detalle_mensual

OUT_DIR = os.path.join(B.SALIDA_DIR, "Distribuibles")
SOLO_TABS = ["dashboard", "resumen"]     # pestañas que quedan visibles

# Compañías a emitir: código de compañía (= CECO[:4]) → (abreviatura, nombre).
# Son las 4 distribuibles pedidas; Michilla (1057) queda fuera a propósito.
COMPANIAS = [("1001", "MLP", "Los Pelambres"),
             ("1002", "CEN", "Centinela"),
             ("1003", "ANT", "Antucoya"),
             ("1005", "CMZ", "Zaldívar")]

# Ejercicio (antes «Medida») de las líneas de Real: una por bucket del dashboard.
EJERCICIO_REAL = {"2022": "Real 2022", "2023": "Real 2023", "2024": "Real 2024",
                  "2025": "Real 2025", "2026ytd": "Real 2026 YTD"}


def _periodo(anio, mes):
    """Mes-Año en el mismo formato que traen las fuentes ('2026-01'): ordena cronológicamente."""
    m = str(mes)
    return f"{anio}-{m}" if m.isdigit() else f"{anio} {m}"


# ── 1) Dashboard recortado ────────────────────────────────────────────────────────────────
def _dashboard(html, model, LOGO, DET, DETP, V, cecos, titulo, fn):
    """Igual que dividir_por_vp._escribir pero fijando además las pestañas visibles."""
    recs = [r for r in V["records"] if r["ceco"] in cecos]
    Vf = dict(V)
    Vf["records"] = recs
    Vf["cecoOld"] = {c: m for c, m in (V.get("cecoOld") or {}).items() if c in cecos}
    Vf["cecoNew"] = {c: m for c, m in (V.get("cecoNew") or {}).items() if c in cecos}
    Vf["soloTabs"] = SOLO_TABS
    # Toda la data es distribuible de UNA compañía: el selector Alcance (Corporativo/Distribuible)
    # y el filtro Compañía no aplican y se ocultan (src/app.jsx y src/filters.jsx → soloDist).
    Vf["soloDist"] = True
    parts = ["window.V2_DATA = " + D._j(Vf) + ";", 'window.DOT_DATA = {"records":[]};']
    if LOGO:
        parts.append("window.CORP_LOGO = " + D._j(LOGO) + ";")
    detf = D._filtrar_det(DET, cecos)
    if detf is not None:
        parts.append("window.DET = " + D._j(detf) + ";")
    detpf = D._filtrar_detp(DETP, cecos)
    if detpf is not None:
        parts.append("window.DETP = " + D._j(detpf) + ";")
    out = B._set_asset(html, B.DATA_UUID, "\n".join(parts))
    model_c = (model
               .replace('"titulo": "Actividad Corporativa"', '"titulo": ' + D._j(str(titulo)))
               .replace('"tituloPlus": "+ Distribuibles Ppto 2027"', '"tituloPlus": ""'))
    out = B._set_asset(out, D.MODEL_UUID, model_c)
    out = re.sub(r"<title>.*?</title>", "<title>" + D._safe(titulo) + " · AMSA</title>", out, count=1)
    os.makedirs(os.path.dirname(fn), exist_ok=True)
    open(fn, "w", encoding="utf-8", newline="").write(out)
    return len(recs), os.path.getsize(fn) / 1e6


# ── 2) BBDD Excel ─────────────────────────────────────────────────────────────────────────
# Formato único (Real y Ppto/Forecast en la misma hoja). En las líneas de Real, «Concepto de
# Gasto» lleva la Contrapartida y «Actividad» la Denominación; el Documento solo aplica al Real.
COLUMNAS = ["CECO", "CLACO", "Concepto de Gasto", "Actividad", "Documento", "VP",
            "Ejercicio", "Periodo", "Valor (USD)", "Valor ajustado (USD)"]


def _base_datos(dr, dp, ceco2vp, cecos):
    """Arma la hoja única (Real + Ppto/Forecast) para los CECOs dados."""
    r = dr[dr["ceco"].isin(cecos)]
    real = pd.DataFrame({
        "CECO": r["ceco"], "CLACO": r["claco"],
        "Concepto de Gasto": r["contra"],          # en el Real, la Contrapartida
        "Actividad": r["denominacion"],            # en el Real, la Denominación
        # '—' es el placeholder del detalle para las líneas sin documento de compra → va vacío.
        "Documento": r["documento"].replace("—", ""),
        "VP": r["ceco"].map(ceco2vp),
        "Ejercicio": r["bucket"].map(EJERCICIO_REAL).fillna(r["bucket"]),
        "Periodo": [_periodo(a, m) for a, m in zip(r["anio"], r["mes"])],
        "Valor (USD)": r["valor_n"], "Valor ajustado (USD)": r["valor_a"],
    })

    p = dp[dp["ceco"].isin(cecos)]
    ejercicio = p["medida"].map(X.MEDIDA_LBL).fillna(p["medida"])
    # Las reversas llevan el nombre del archivo de ajuste para poder distinguirlas entre sí.
    ejercicio = ejercicio.where(p["ajuste"].eq(""), ejercicio + " · " + p["ajuste"])
    ppto = pd.DataFrame({
        "CECO": p["ceco"], "CLACO": p["claco"],
        "Concepto de Gasto": p["concepto_gasto"], "Actividad": p["actividad"],
        "Documento": "",                           # no aplica fuera del Real
        "VP": p["ceco"].map(ceco2vp),
        "Ejercicio": ejercicio,
        "Periodo": [_periodo(a, m) for a, m in zip(p["anio"], p["mes"])],
        "Valor (USD)": p["valor_n"], "Valor ajustado (USD)": p["valor_a"],
    })

    df = pd.concat([real, ppto], ignore_index=True)[COLUMNAS]
    return df.sort_values(["Ejercicio", "Periodo", "CECO", "CLACO"], kind="stable").reset_index(drop=True)


def main():
    if not os.path.isfile(B.SALIDA_HTML):
        raise SystemExit("No existe el dashboard; corré primero: python construir.py")
    B.log("Leyendo el dashboard base…")
    html = open(B.SALIDA_HTML, "r", encoding="utf-8", newline="").read()
    model = B._get_asset(html, D.MODEL_UUID)
    data = B._get_asset(html, B.DATA_UUID)
    V = D._extract(data, "V2_DATA")
    if V is None:
        raise SystemExit("No pude leer window.V2_DATA del dashboard.")
    LOGO, DET, DETP = (D._extract(data, n) for n in ("CORP_LOGO", "DET", "DETP"))

    B.log("Leyendo la BBDD (detalle mensual)…")
    dr, dp = detalle_mensual.construir()
    cec = pd.read_parquet(os.path.join(B.BBDD_DIR, "dim_cecos.parquet"))
    new = cec[cec.estructura == "new"].drop_duplicates("ceco")
    ceco2vp = dict(zip(new["ceco"], new["vp"]))

    os.makedirs(OUT_DIR, exist_ok=True)
    for cod, abrev, nombre in COMPANIAS:
        # Universo de la compañía: todo CECO que empiece con su código, venga de donde venga
        # (records del dashboard, mapas de CECOS o detalle) — así no se pierde ninguna línea.
        cecos = {c for c in (set(V.get("cecoNew") or {}) | set(V.get("cecoOld") or {})
                             | {r["ceco"] for r in V["records"]}
                             | set(dr["ceco"].unique()) | set(dp["ceco"].unique()))
                 if str(c).startswith(cod)}
        carpeta = os.path.join(OUT_DIR, abrev)
        titulo = f"Distribuibles {nombre}"

        nregs, mb = _dashboard(html, model, LOGO, DET, DETP, V, cecos, titulo,
                               os.path.join(carpeta, f"Dashboard {titulo}.html"))
        df = _base_datos(dr, dp, ceco2vp, cecos)
        xls = os.path.join(carpeta, f"BBDD {titulo}.xlsx")
        mbx = X.escribir_excel(xls, [("Base de Datos", df, ["Valor (USD)", "Valor ajustado (USD)"])])
        B.log(f"  {abrev} · {nombre:16} {len(cecos):3} CECOs · dashboard {nregs:5} regs / {mb:4.1f} MB"
              f" · Excel {len(df):7,} filas / {mbx:4.1f} MB")

    B.log(f"\n✓ LISTO. {len(COMPANIAS)} compañías en {OUT_DIR}")


if __name__ == "__main__":
    main()
