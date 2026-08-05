# Cómo actualizar reversas y ejercicios del Dashboard (comandos Python)

Todos los comandos se corren desde la carpeta raíz del repo
(`...\amsa-budget-corporativo`), en una terminal.

---

## 1) Reversas al Forecast 5+7  (carpeta `v2/uploads/ajustes_forecast/`)

Cada Excel de reversas necesita una hoja **«… Unpivot»** (con la columna
`Valor USD 2027 (mes equivalente)`). El dashboard las **descubre solas**: toma
TODO `.xlsx` de esta carpeta que tenga una hoja terminada en «Unpivot» y las
suma al Forecast 5+7 **solo** en la Estructura CECOS **«Nueva con ajustes»**.
Suelen ser redistribuciones (neto ≈ 0) que mueven plata entre CECOs/Ítems.

### a) Generar / actualizar la hoja Unpivot de un archivo
```
python v2/unpivot_ejercicio.py "v2/uploads/ajustes_forecast/<ARCHIVO>.xlsx" --hoja "<hoja origen>" --salida "<hoja origen> Unpivot"
```
- `--hoja`   = nombre de la hoja cruda (Reversa, Hoja1, etc.)
- `--salida` = nombre de la hoja a generar; **debe terminar en «Unpivot»**.

Ejemplos reales:
```
python v2/unpivot_ejercicio.py "v2/uploads/ajustes_forecast/Distribución Compliance-Riesgo 2026 F5+7.xlsx" --hoja "Reversa" --salida "Reversa Unpivot"
python v2/unpivot_ejercicio.py "v2/uploads/ajustes_forecast/Dist Team building VPEI 5+7.xlsx" --hoja "Hoja1" --salida "Reversa Unpivot"
```

### b) Agregar un Excel de reversas NUEVO (sin tocar código)
1. Deja el `.xlsx` en `v2/uploads/ajustes_forecast/`.
2. Genera su hoja Unpivot con el comando de arriba.
3. Reconstruye el dashboard (sección 4). Listo: se descubre automáticamente.

### ⚠️ Cuidado con fórmulas
Si las celdas de mes del Excel son **fórmulas**: ábrelo y **guárdalo en Excel
(Ctrl+S)** antes de correr el unpivot (para que Excel calcule los valores).
**No** re-corras el unpivot sobre el mismo archivo ya procesado: las fórmulas
quedan sin valor cacheado y daría 0. Regenera siempre desde una copia fresca de Excel.

---

## 2) Ejercicios: Forecast 5+7 y Outlook 6+6  (carpeta `v2/uploads/ejercicios_2026/`)

Mismo script; la hoja cruda suele llamarse `Sheet1`:
```
python v2/unpivot_ejercicio.py "v2/uploads/ejercicios_2026/Outlook 6+6 2026.XLSX" --hoja "Sheet1" --salida "Outlook 6+6 Unpivot"
python v2/unpivot_ejercicio.py "v2/uploads/ejercicios_2026/Forecast 5+7 2026.XLSX" --hoja "Sheet1" --salida "Forecast 5+7 Unpivot"
```

---

## 3) Nuevo Presupuesto 2027

1. Deja el archivo en `v2/uploads/presupuesto_2027/`.
2. En `v2/construir_v2.py` edita la línea:
   `PPTO27_FILE = os.path.join(UP, "presupuesto_2027", "PPTO27_DD_MM_AAAA_HH_MM.XLSX")`
   (la fecha del subtítulo del dashboard sale del nombre del archivo).

---

## 4) Reconstruir y publicar el Dashboard

- **Solo previsualizar** (v3 local, no publica):
  ```
  python v2/construir_v2.py --solo-v3
  ```
- **Completo + copias extra + por-VP** (todo en uno):
  ```
  python v2/actualizar_dashboards.py
  ```
- **Publicar por-VP/Gerencia a OneDrive** (después de generar):
  ```
  python v2/dividir_por_vp.py
  python v2/publicar.py
  ```
