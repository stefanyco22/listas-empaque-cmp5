import pandas as pd
import os
import re

# Carpeta donde están tus listas de empaque
carpeta = r"RUTA\DE\TU\CARPETA"

def limpiar_texto(t):
    if not isinstance(t, str):
        return ""
    return re.sub(r'\s+', ' ', t.strip().upper().replace('.', '')).replace('Ó', 'O').replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ú', 'U')

def procesar_lista(ruta_archivo):
    try:
        # Leer más filas para incluir encabezado combinado
        df = pd.read_excel(ruta_archivo, header=None)

        # Tomar las filas 14 y 15 (índices 13 y 14)
        encabezado_fila1 = df.iloc[13]
        encabezado_fila2 = df.iloc[14]

        # Fusionar encabezados (por si hay combinaciones)
        encabezado = []
        for i in range(len(encabezado_fila1)):
            val = str(encabezado_fila1[i]) if pd.notna(encabezado_fila1[i]) else ''
            val2 = str(encabezado_fila2[i]) if pd.notna(encabezado_fila2[i]) else ''
            encabezado.append(limpiar_texto(f"{val} {val2}".strip()))

        # Crear nuevo DataFrame con encabezado unificado
        df = pd.read_excel(ruta_archivo, header=None, skiprows=15)
        df.columns = encabezado

        # Detectar columnas relevantes
        col_caja = next((c for c in df.columns if "CAJA" in c), None)
        col_parte = next((c for c in df.columns if "PARTE" in c), None)
        col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)

        if not all([col_caja, col_parte, col_cant]):
            print(f"⚠️ Columnas esperadas no encontradas en {os.path.basename(ruta_archivo)}")
            return None

        # Filtrar filas válidas
        df = df[[col_caja, col_parte, col_cant]]
        df = df.dropna(subset=[col_parte])
        df = df[~df[col_parte].astype(str).str.contains("PALLET", case=False, na=False)]

        df["Archivo"] = os.path.basename(ruta_archivo)
        return df

    except Exception as e:
        print(f"❌ Error procesando {ruta_archivo}: {e}")
        return None

# Procesar todos los archivos .xlsx en la carpeta
listas = []
for archivo in os.listdir(carpeta):
    if archivo.endswith(".xlsx"):
        ruta = os.path.join(carpeta, archivo)
        df = procesar_lista(ruta)
        if df is not None:
            listas.append(df)

# Unir todas las listas
if listas:
    consolidado = pd.concat(listas, ignore_index=True)
    consolidado.to_excel(os.path.join(carpeta, "LISTAS_PROCESADAS.xlsx"), index=False)
    print("✅ Listas procesadas correctamente. Archivo generado: LISTAS_PROCESADAS.xlsx")
else:
    print("⚠️ No se procesó ninguna lista.")


