
import streamlit as st
import pandas as pd
import io
from datetime import date

st.set_page_config(page_title="📦 Sistema de Consolidación de Listas de Empaque CMP", page_icon="📦", layout="wide")

st.markdown("<h2 style='text-align:center;color:#004AAD;'>📦 Sistema de Consolidación de Listas de Empaque CMP</h2>", unsafe_allow_html=True)
st.write("### 1️⃣ Cargar Consolidado")
consolidado_file = st.file_uploader("Sube el archivo CONSOLIDADO (.xlsx)", type=["xlsx"], key="consolidado")

st.write("### 2️⃣ Cargar Listas de Empaque")
lista_files = st.file_uploader("Sube una o varias Listas de Empaque (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="listas")

def encontrar_encabezados(df):
    for i in range(min(30, len(df))):
        fila = " ".join(str(x) for x in df.iloc[i].fillna(""))
        if ("caja" in fila.lower() and "parte" in fila.lower()) or ("cantidad" in fila.lower() and "empacada" in fila.lower()):
            return i
    return None

def procesar_lista(df, df_cons):
    fila_headers = encontrar_encabezados(df)
    if fila_headers is None:
        return None, f"Encabezados no encontrados en las primeras filas."

    df.columns = df.iloc[fila_headers]
    df = df.iloc[fila_headers + 1:].reset_index(drop=True)
    df = df.rename(columns=lambda x: str(x).strip())

    col_caja = next((c for c in df.columns if "caja" in c.lower()), None)
    col_parte = next((c for c in df.columns if "parte" in c.lower()), None)
    col_cant = next((c for c in df.columns if "empac" in c.lower()), None)

    if not all([col_caja, col_parte, col_cant]):
        return None, f"Columnas esperadas no encontradas (No. de Caja, Número de Parte, Cantidad Empacada)."

    df = df[[col_caja, col_parte, col_cant]]
    df.columns = ["No. de Caja", "Número de Parte", "Cantidad Empacada"]
    df["No. de Caja"] = df["No. de Caja"].ffill()

    df_final = pd.merge(df, df_cons, how="left", left_on="Número de Parte", right_on="COD.")
    df_final = df_final[["No. de Caja", "Número de Parte", "Cantidad Empacada", "DESCRIPCION"]]
    for col in ["CANTIDAD FISICA", "U/M", "U/M POR CADA", "ORDEN DE PRODUCCION", "LOTE", "OBSERVACION"]:
        df_final[col] = ""

    return df_final, None

if st.button("🚀 Procesar Listas de Empaque"):
    if consolidado_file is None or not lista_files:
        st.error("Por favor sube el consolidado y al menos una lista de empaque.")
    else:
        df_cons = pd.read_excel(consolidado_file, dtype=str)
        if not all(col in df_cons.columns for col in ["DESPACHO", "COD.", "DESCRIPCION"]):
            st.error("El archivo CONSOLIDADO debe tener las columnas: DESPACHO, COD. y DESCRIPCION.")
        else:
            with pd.ExcelWriter(io.BytesIO(), engine="xlsxwriter") as writer:
                errores = []
                for file in lista_files:
                    df = pd.read_excel(file, header=None)
                    df_proc, err = procesar_lista(df, df_cons)
                    if err:
                        errores.append(f"{file.name}: {err}")
                        continue
                    nombre_hoja = file.name.replace(".xlsx", "")[:31]
                    df_proc.to_excel(writer, sheet_name=nombre_hoja, index=False)

                writer.book.formats[0].set_font_color("#004AAD")
                st.success(f"✅ Procesamiento completado ({len(lista_files) - len(errores)} exitosas).")
                if errores:
                    st.warning("\n".join(errores))

                writer.save()
                output = writer.book.filename
                st.download_button("📥 Descargar Excel Consolidado",
                                   data=writer.book.filename.getvalue(),
                                   file_name=f"LISTAS_DE_EMPAQUE_CMP_{date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
