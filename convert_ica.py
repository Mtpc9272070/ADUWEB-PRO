import pandas as pd
import json
import re
import os

# ---------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------
EXCEL_PATH = "Ch_2_Annex2B_COL_s.xlsx"
OUTPUT_JSON = "aranceles_desde_excel.json"

# ---------------------------------------------------
# DETECTA CONTROLES ICA EN LA DESCRIPCIÓN
# ---------------------------------------------------
def extraer_controles(texto):
    if not isinstance(texto, str):
        return []
    controles = re.findall(
        r"\b(DZI|DRFI|LV|CI|CA|P\.I\.|P\.I)\b",
        texto,
        flags=re.IGNORECASE
    )
    controles = [c.upper().replace("P.I.", "P.I") for c in controles]
    return list(dict.fromkeys(controles))

# ---------------------------------------------------
# GENERA ESTRUCTURA ARANCELARIA
# ---------------------------------------------------
def estructura_arancelaria(codigo):
    codigo = str(codigo).strip()
    if not codigo.isdigit():
        return None
    if len(codigo) < 6:
        return None

    return {
        "capitulo": codigo[:2],
        "partida": codigo[:4],
        "subpartida": codigo[:6],
        "nivel_1": codigo[:8] if len(codigo) >= 8 else None,
        "nivel_2": codigo[:10] if len(codigo) >= 10 else None
    }

# ---------------------------------------------------
# LEER EXCEL
# ---------------------------------------------------
print("📄 Leyendo archivo Excel...")
df = pd.read_excel(EXCEL_PATH, header=1)


print("Columnas detectadas:", df.columns.tolist())

# ---------------------------------------------------
# SELECCIÓN AUTOMÁTICA DE COLUMNAS
# ---------------------------------------------------
col_codigo = None
col_desc = None

for col in df.columns:
    nombre = col.lower().strip()

    # detectar código
    if "código arancelario" in nombre or "codigo arancelario" in nombre:
        col_codigo = col
    
    # detectar descripción
    if "descripción" in nombre or "descripcion" in nombre:
        col_desc = col

if not col_codigo:
    raise Exception("❌ No encontré la columna 'CÓDIGO ARANCELARIO' en el Excel.")

if not col_desc:
    raise Exception("❌ No encontré la columna 'DESCRIPCIÓN' en el Excel.")

print(f"➡ Columna de código: {col_codigo}")
print(f"➡ Columna de descripción: {col_desc}")

# ---------------------------------------------------
# PROCESAMIENTO DE FILAS
# ---------------------------------------------------
registros = []

for _, row in df.iterrows():
    codigo = str(row[col_codigo]).strip()

    if codigo.lower() == "nan" or codigo == "":
        continue

    estructura = estructura_arancelaria(codigo)
    if not estructura:
        continue

    descripcion = str(row[col_desc]).strip()
    controles = extraer_controles(descripcion)

    registros.append({
        "codigo": codigo,
        "estructura": estructura,
        "producto": {
            "descripcion": descripcion
        },
        "control_ica": {
            "documentos_requeridos": controles,
            "detalles": ""
        },
        "notas_marginales": [],
        "fuente": EXCEL_PATH
    })

# ---------------------------------------------------
# GUARDAR JSON
# ---------------------------------------------------
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(registros, f, ensure_ascii=False, indent=2)

print(f"✅ JSON generado: {OUTPUT_JSON}")
print(f"📦 Registros procesados: {len(registros)}")
