import streamlit as st
import pandas as pd
import io
import unicodedata
from datetime import date

st.set_page_config(page_title="📦 Sistema de Consolidación de Listas de Empaque CMP", page_icon="📦", layout="wide")

st.markdown("<div style='background:#0b5ed7;padding:14px;border-radius:8px;color:white'><h2>📦 Sistema de Consolidación de Listas de Empaque CMP</h2></div>", unsafe_allow_html=True)
st.write("Sube primero el archivo **CONSOLIDADO.xlsx** y luego una o varias **Listas de Empaque**.")

def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.upper()

def find_header_row(df_no_header, max_rows=30):
    for i in range(min(max_rows, len(df_no_header))):
        row_text = " ".join(str(x) for x in df_no_header.iloc[i].fillna(""))
        row_up = row_text.upper()
        if ("NO" in row_up and "CAJA" in row_up and "PARTE" in row_up) or ("CANTIDAD" in row_up and "EMPAC" in row_up):
            return i
    return None

def normalize_consolidado(df_cons):
    cols_map = {normalize_text(c).replace(" ", "").replace(".", ""): c for c in df_cons.columns}
    despacho = next((cols_map[k] for k in cols_map if "DESPACHO" in k), None)
    cod = next((cols_map[k] for k in cols_map if k.startswith("COD") or "CODIGO" in k), None)
    desc = next((cols_map[k] for k in cols_map if "DESCRIP" in k), None)
    if not despacho or not cod or not desc:
        raise ValueError("No se encontraron columnas similares a DESPACHO, COD. y DESCRIPCION en el consolidado.")
    df_cons = df_cons.rename(columns={despacho: "DESPACHO", cod: "COD.", desc: "DESCRIPCION"})
    df_cons["COD."] = df_cons["COD."].astype(str).str.strip()
    df_cons["DESCRIPCION"] = df_cons["DESCRIPCION"].astype(str).fillna("")
    return df_cons[["COD.", "DESCRIPCION"]]

st.subheader("1) Subir archivo CONSOLIDADO (.xlsx)")
cons_file = st.file_uploader("Selecciona CONSOLIDADO (DESPACHO | COD. | DESCRIPCION)", type=["xlsx"], key="cons")

st.subheader("2) Subir Listas de Empaque (.xlsx) - puedes seleccionar varias")
list_files = st.file_uploader("Selecciona listas de empaque", type=["xlsx"], accept_multiple_files=True, key="lists")

if cons_file and list_files:
    try:
        df_cons_raw = pd.read_excel(io.BytesIO(cons_file.getvalue()), dtype=str)
        df_cons = normalize_consolidado(df_cons_raw)
    except Exception as e:
        st.error(f"Error leyendo consolidado: {e}")
    else:
        st.success("Consolidado cargado correctamente.")
        if st.checkbox("Ver primeras filas del consolidado"):
            st.dataframe(df_cons_raw.head())

        resultados = {}
        problemas = []

        for f in list_files:
            try:
                raw = pd.read_excel(io.BytesIO(f.getvalue()), header=None, dtype=str)
            except Exception as e:
                problemas.append((f.name, f"Error lectura: {e}"))
                continue

            header_row = find_header_row(raw, max_rows=30)
            if header_row is None:
                problemas.append((f.name, "No se encontró fila de encabezado en las primeras 30 filas."))
                continue

            try:
                df = pd.read_excel(io.BytesIO(f.getvalue()), header=header_row, dtype=str)
            except Exception as e:
                problemas.append((f.name, f"Error leyendo con header: {e}"))
                continue

            df.columns = [normalize_text(c) for c in df.columns]
            col_caja = next((c for c in df.columns if "CAJA" in c), None)
            col_parte = next((c for c in df.columns if "PARTE" in c), None)
            col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)

            if not col_caja or not col_parte or not col_cant:
                problemas.append((f.name, "Columnas esperadas No. de Caja / Número de Parte / Cantidad Empacada no encontradas después del encabezado."))
                continue

            df_extract = df[[col_caja, col_parte, col_cant]].copy()
            df_extract.columns = ["No. de Caja", "Número de Parte", "Cantidad Empacada"]
            df_extract["No. de Caja"] = df_extract["No. de Caja"].ffill()
            df_extract["Número de Parte"] = df_extract["Número de Parte"].astype(str).str.strip()

            df_extract = df_extract.merge(df_cons, how="left", left_on="Número de Parte", right_on="COD.")
            df_extract["DESCRIPCION"] = df_extract["DESCRIPCION"].fillna("NO ENCONTRADO")

            for c in ["CANTIDAD FISICA", "U/M", "U/M POR CADA", "ORDEN DE PRODUCCION", "LOTE", "OBSERVACION"]:
                df_extract[c] = ""

            sheet_name = f.name.rsplit(".",1)[0][:31]
            resultados[sheet_name] = df_extract

        if not resultados:
            st.warning("No se procesó ninguna lista. Revisa los problemas listados abajo.")
            for p in problemas:
                st.write("•", p[0], "-", p[1])
        else:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet, df_out in resultados.items():
                    df_out.to_excel(writer, sheet_name=sheet, index=False)
            output.seek(0)
            filename = f"Listas_Empaque_CMP_{date.today()}.xlsx"
            st.success("Consolidación completada ✅")
            st.download_button("⬇️ Descargar Excel consolidado", data=output.getvalue(), file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if problemas:
                st.warning("Algunos archivos tuvieron problemas:")
                for p in problemas:
                    st.write("•", p[0], "-", p[1])
