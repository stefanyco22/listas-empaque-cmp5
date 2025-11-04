import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.title("Procesador de Listas de Empaque")

# Carpeta donde están tus archivos (relativa al proyecto)
carpeta = Path("archivos")  # asegúrate de subir la carpeta 'archivos' a tu proyecto

def limpiar_texto(t):
    if not isinstance(t, str):
        return ""
    return re.sub(r'\s+', ' ', t.strip().upper().replace('.', '')).replace('Ó', 'O').replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ú', 'U')

def procesar_lista(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo, header=None)

        # Tomar filas 14 y 15 (índices 13 y 14)
        encabezado_fila1 = df.iloc[13]
        encabezado_fila2 = df.iloc[14]

        encabezado = []
        for i in range(len(encabezado_fila1)):
            val = str(encabezado_fila1[i]) if pd.notna(encabezado_fila1[i]) else ''
            val2 = str(encabezado_fila2[i]) if pd.notna(encabezado_fila2[i]) else ''
            encabezado.append(limpiar_texto(f"{val} {val2}".strip()))

        df = pd.read_excel(ruta_archivo, header=None, skiprows=15)
        df.columns = encabezado

        col_caja = next((c for c in df.columns if "CAJA" in c), None)
        col_parte = next((c for c in df.columns if "PARTE" in c), None)
        col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)

        if not all([col_caja, col_parte, col_cant]):
            st.warning(f"⚠️ Columnas esperadas no encontradas en {ruta_archivo.name}")
            return None

        df = df[[col_caja, col_parte, col_cant]]
        df = df.dropna(subset=[col_parte])
        df = df[~df[col_parte].astype(str).str.contains("PALLET", case=False, na=False)]

        df["Archivo"] = ruta_archivo.name
        return df

    except Exception as e:
        st.error(f"❌ Error procesando {ruta_archivo.name}: {e}")
        return None

listas = []

if carpeta.exists():
    archivos_xlsx = list(carpeta.glob("*.xlsx"))
    if archivos_xlsx:
        for archivo in archivos_xlsx:
            df = procesar_lista(archivo)
            if df is not None:
                listas.append(df)

        if listas:
            consolidado = pd.concat(listas, ignore_index=True)
            salida = carpeta / "LISTAS_PROCESADAS.xlsx"
            consolidado.to_excel(salida, index=False)
            st.success(f"✅ Listas procesadas correctamente. Archivo generado: {salida.name}")
            st.download_button(
                label="Descargar consolidado",
                data=consolidado.to_excel(index=False),
                file_name="LISTAS_PROCESADAS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No se procesó ninguna lista.")
    else:
        st.warning(f"⚠️ No se encontraron archivos .xlsx en la carpeta '{carpeta}'")
else:
    st.error(f"❌ La carpeta '{carpeta}' no existe. Por favor súbela al proyecto.")
