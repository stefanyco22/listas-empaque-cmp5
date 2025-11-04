import streamlit as st
import pandas as pd
import io, re, unicodedata
from datetime import date

# --- Configuración de página ---
st.set_page_config(page_title="📦 Sistema de Consolidación de Listas de Empaque CMP",
                   page_icon="📦", layout="wide")
st.markdown(
    '<div style="background:#0b5ed7;padding:12px;border-radius:8px;color:white">'
    '<h2>📦 Sistema de Consolidación de Listas de Empaque CMP</h2></div>', unsafe_allow_html=True
)
st.write("Sube primero el archivo **CONSOLIDADO.xlsx** y luego una o varias **Listas de Empaque**. "
         "Marca 'Vista previa' para revisar los encabezados antes de consolidar.")

# --- Funciones auxiliares ---
def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.upper()

def detect_header_row(df):
    """
    Detecta encabezados combinados automáticamente:
    - Preferencia: filas 14–15 (índices 13–14)
    - fallback: buscar coincidencias de palabras clave en primeras 30 filas
    Retorna:
        header_row_index (int), preview_rows (DataFrame de filas 14–17)
    """
    max_rows = min(30, len(df))
    preview_rows = df.iloc[13:17] if len(df) >= 17 else df.iloc[:max_rows]

    # Intento filas 14–15 combinadas
    if len(df) >= 15:
        combined = " ".join(str(x) for x in df.iloc[13].fillna("")) + " " + " ".join(str(x) for x in df.iloc[14].fillna(""))
        combined = combined.upper()
        if any(word in combined for word in ["CAJA", "PARTE", "CANTIDAD", "EMPAC"]):
            return 13, preview_rows

    # Fallback: buscar en primeras 30 filas
    for i in range(max_rows):
        row_text = " ".join(str(x) for x in df.iloc[i].fillna("")).upper()
        if all(word in row_text for word in ["CAJA", "PARTE"]):
            return i, preview_rows

    return None, preview_rows

def normalize_consolidado(df_cons):
    cols_map = {re.sub(r'[\s\.]','', normalize_text(c)): c for c in df_cons.columns}
    despacho = next((cols_map[k] for k in cols_map if "DESPACHO" in k), None)
    cod = next((cols_map[k] for k in cols_map if k.startswith("COD") or "CODIGO" in k), None)
    desc = next((cols_map[k] for k in cols_map if "DESCRIP" in k), None)
    if not despacho or not cod or not desc:
        raise ValueError("No se encontraron columnas similares a DESPACHO, COD. y DESCRIPCION en el consolidado.")
    df_cons = df_cons.rename(columns={despacho: "DESPACHO", cod: "COD.", desc: "DESCRIPCION"})
    df_cons["COD."] = df_cons["COD."].astype(str).str.strip()
    df_cons["DESCRIPCION"] = df_cons["DESCRIPCION"].astype(str).fillna("")
    return df_cons[["COD.", "DESCRIPCION"]]

def extract_short_sheet_name(filename):
    name = filename.rsplit(".",1)[0]
    m = re.search(r'(DC[\s\-]*\d{1,3}[\-\s]*\d{1,3})', name, re.IGNORECASE)
    if m:
        s = m.group(1).upper().replace("_"," ").strip()
        s = re.sub(r'\s+', ' ', s)
        return s[:31]
    parts = re.split(r'[\s_]+', name)
    if len(parts) >= 2:
        return " ".join(parts[-2:])[:31]
    return name[:31]

# --- Controles de subida ---
st.subheader("1) Subir archivo CONSOLIDADO (.xlsx)")
cons_file = st.file_uploader("Selecciona CONSOLIDADO (DESPACHO | COD. | DESCRIPCION)", type=["xlsx"], key="cons")

st.subheader("2) Subir Listas de Empaque (.xlsx) - puedes seleccionar varias")
list_files = st.file_uploader("Selecciona listas de empaque", type=["xlsx"], accept_multiple_files=True, key="lists")

preview = st.checkbox("👁️ Ver vista previa de encabezados detectados y primeras filas")

process_button = st.button("🚀 Procesar y Consolidar")

# --- Procesamiento ---
if cons_file and list_files and process_button:
    try:
        df_cons_raw = pd.read_excel(io.BytesIO(cons_file.getvalue()), dtype=str)
        df_cons = normalize_consolidado(df_cons_raw)
    except Exception as e:
        st.error(f"Error leyendo consolidado: {e}")
    else:
        st.success("Consolidado cargado correctamente.")
        if st.checkbox("Mostrar primeras filas del consolidado"):
            st.dataframe(df_cons_raw.head())

        resultados = {}
        problemas = {}
        previews = {}

        prog = st.progress(0)
        total = len(list_files)
        idx = 0

        for f in list_files:
            idx += 1
            try:
                raw = pd.read_excel(io.BytesIO(f.getvalue()), header=None, dtype=str)
            except Exception as e:
                problemas[f.name] = f"Error lectura: {e}"
                prog.progress(int(idx/total*100))
                continue

            header_row, preview_rows = detect_header_row(raw)
            previews[f.name] = preview_rows

            if header_row is None:
                problemas[f.name] = "No se encontró fila de encabezado."
                prog.progress(int(idx/total*100))
                continue

            # Ignorar fila 16 (índice 15), iniciar desde fila 17
            start_data_row = header_row + 3
            df = pd.read_excel(io.BytesIO(f.getvalue()), header=header_row, skiprows=[header_row+2], dtype=str)

            # Normalizar columnas
            df.columns = [normalize_text(c) for c in df.columns]
            col_caja = next((c for c in df.columns if "CAJA" in c), None)
            col_parte = next((c for c in df.columns if "PARTE" in c), None)
            col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)
            if not col_caja or not col_parte or not col_cant:
                problemas[f.name] = "Columnas esperadas No. de Caja / Número de Parte / Cantidad Empacada no encontradas."
                prog.progress(int(idx/total*100))
                continue

            df_extract = df[[col_caja, col_parte, col_cant]].copy()
            df_extract.columns = ["No. de Caja", "Número de Parte", "Cantidad Empacada"]
            df_extract["No. de Caja"] = df_extract["No. de Caja"].ffill()
            df_extract["Número de Parte"] = df_extract["Número de Parte"].astype(str).str.strip()

            # Merge con consolidado
            df_extract = df_extract.merge(df_cons, how="left", left_on="Número de Parte", right_on="COD.")
            df_extract["DESCRIPCION"] = df_extract["DESCRIPCION"].fillna("NO ENCONTRADO")

            # Columnas adicionales
            for c in ["CANTIDAD FISICA", "U/M", "U/M POR CADA", "ORDEN DE PRODUCCION", "LOTE", "OBSERVACION"]:
                df_extract[c] = ""
            df_extract["Inicio de datos"] = start_data_row + 1  # fila en Excel (1-based)

            short_name = extract_short_sheet_name(f.name)
            resultados[short_name] = df_extract

            prog.progress(int(idx/total*100))

        # --- Mostrar vista previa de encabezados y primeras filas ---
        if preview:
            st.subheader("Vista previa de encabezados detectados y primeras filas")
            for name, dfp in previews.items():
                st.markdown(f"**{name}**")
                st.dataframe(dfp)

        if not resultados:
            st.warning("No se procesó ninguna lista. Revisa los problemas:")
            for name, prob in problemas.items():
                st.write("•", name, "-", prob)
        else:
            # Generar Excel consolidado
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet, df_out in resultados.items():
                    df_out.to_excel(writer, sheet_name=sheet[:31], index=False)
                merged = pd.concat(resultados.values(), ignore_index=True)
                merged.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
            output.seek(0)

            filename = f"CONSOLIDADO_CMP_{date.today()}.xlsx"
            st.success("Consolidación lista ✅")
            st.download_button("⬇️ Descargar Excel consolidado",
                               data=output.getvalue(),
                               file_name=filename,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            if problemas:
                st.warning("Algunos archivos tuvieron problemas:")
                for name, prob in problemas.items():
                    st.write("•", name, "-", prob)

