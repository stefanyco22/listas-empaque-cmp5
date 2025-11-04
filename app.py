
import streamlit as st
import pandas as pd
import io
import unicodedata
from datetime import datetime

st.set_page_config(page_title="📦 Sistema de Consolidación de Listas de Empaque CMP", layout="centered")

def normalize(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove accents
    return s.upper()

st.markdown('<div style="background:#0b5ed7;padding:14px;border-radius:8px;color:white"><h2>📦 Sistema de Consolidación de Listas de Empaque CMP</h2></div>', unsafe_allow_html=True)
st.write("Sube primero el archivo **CONSOLIDADO.xlsx** y luego una o varias **Listas de Empaque**. Cada lista generará su propia hoja en el Excel resultante.")

st.subheader("1) Subir CONSOLIDADO (.xlsx)")
cons_file = st.file_uploader("Selecciona CONSOLIDADO (DESPACHO | COD. | DESCRIPCION)", type=["xlsx"], key="cons")

if cons_file is not None:
    try:
        df_cons = pd.read_excel(io.BytesIO(cons_file.getvalue()), dtype=str)
    except Exception as e:
        st.error(f"Error leyendo CONSOLIDADO: {e}")
        st.stop()

    # Normalize consolidated headers and find COD and DESCRIPCION columns
    cons_cols = {normalize(c): c for c in df_cons.columns}
    cod_key = None
    desc_key = None
    for k_norm, orig in cons_cols.items():
        if k_norm.replace(".", "").replace(" ", "").startswith("COD"):
            cod_key = orig
        if "DESCRIPCION" in k_norm or "DESCRIP" in k_norm:
            desc_key = orig

    if cod_key is None or desc_key is None:
        st.error("No se encontraron las columnas 'COD.' y/o 'DESCRIPCION' en el CONSOLIDADO. Revise el archivo.")
        st.stop()

    # Build mapping from COD to DESCRIPCION
    df_cons[cod_key] = df_cons[cod_key].astype(str).str.strip()
    df_cons[desc_key] = df_cons[desc_key].astype(str).fillna("")
    mapping = dict(zip(df_cons[cod_key], df_cons[desc_key]))

    st.success("✅ Consolidado cargado correctamente.")
    if st.checkbox("Ver primeras filas del consolidado"):
        st.dataframe(df_cons.head())

    st.subheader("2) Subir Listas de Empaque (.xlsx) - varias")
    files = st.file_uploader("Selecciona uno o varios archivos de listas de empaque", type=["xlsx"], accept_multiple_files=True, key="lists")

    if files:
        st.info(f"Procesando {len(files)} archivo(s)...")
        resultados = {}
        problemas = []

        for f in files:
            try:
                # Read sheet starting from row 17 (skip first 16 rows), header is the next row
                # Use BytesIO to ensure Streamlit file object works reliably
                raw = pd.read_excel(io.BytesIO(f.getvalue()), header=None, dtype=str)
                # The user confirmed headers are in row 17 (1-based) -> index 16 (0-based)
                header_row_index = 16
                if header_row_index >= len(raw):
                    problemas.append((f.name, "Archivo muy corto — no se encontró fila 17."))
                    continue

                # Read again using the header row as header
                df = pd.read_excel(io.BytesIO(f.getvalue()), header=header_row_index, dtype=str)
                # Normalize column names
                df.columns = [normalize(c) for c in df.columns]

                # Find required columns
                # Expected exact names (normalized): "NO. DE CAJA", "NÚMERO DE PARTE" -> normalized will remove accents
                # We'll look for contains to be robust
                col_caja = next((c for c in df.columns if "CAJA" in c), None)
                col_parte = next((c for c in df.columns if "PARTE" in c), None)
                col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)

                if not col_caja or not col_parte or not col_cant:
                    problemas.append((f.name, "Columnas esperadas no encontradas después de la fila 17."))
                    continue

                df_extract = df[[col_caja, col_parte, col_cant]].copy()
                # rename to final friendly names
                df_extract.columns = ["No. de Caja", "Número de Parte", "Cantidad Empacada"]

                # forward fill No. de Caja (first valid value will fill subsequent blanks)
                df_extract["No. de Caja"] = df_extract["No. de Caja"].ffill()

                # Clean Numero de Parte and map description
                df_extract["Número de Parte"] = df_extract["Número de Parte"].astype(str).str.strip()
                df_extract["Descripción"] = df_extract["Número de Parte"].map(mapping).fillna("NO ENCONTRADO")

                # Extra blank columns
                for c in ["CANTIDAD FISICA", "U/M", "U/M POR CADA", "ORDEN DE PRODUCCION", "LOTE", "OBSERVACION"]:
                    df_extract[c] = ""

                # Save into resultados, sheet name from filename (without extension), truncate to 31 chars
                sheet_name = f.name.rsplit(".", 1)[0][:31]
                resultados[sheet_name] = df_extract

            except Exception as e:
                problemas.append((f.name, f"Error al procesar: {e}"))

        if not resultados:
            st.warning("No se procesó ninguna lista. Revisa los problemas listados abajo.")
            for p in problemas:
                st.write("•", p[0], "-", p[1])
        else:
            # Build Excel in memory with one sheet per list
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet, df_out in resultados.items():
                    # Ensure sheet name is valid
                    safe_name = sheet[:31]
                    df_out.to_excel(writer, sheet_name=safe_name, index=False)
            output.seek(0)

            fecha = datetime.now().strftime("%Y-%m-%d")
            filename = f"Listas_Empaque_CMP_{fecha}.xlsx"

            st.success("✅ Consolidación completada.")
            st.download_button("⬇️ Descargar archivo consolidado", data=output, file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            if problemas:
                st.warning("Algunos archivos tuvieron problemas:")
                for p in problemas:
                    st.write("•", p[0], "-", p[1])
