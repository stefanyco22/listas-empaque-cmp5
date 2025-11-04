import streamlit as st
import pandas as pd
import io, re, unicodedata
from datetime import date

# --- Page config & simple styling ---
st.set_page_config(page_title="📦 Sistema de Consolidación de Listas de Empaque CMP", page_icon="📦", layout="wide")
st.markdown('<div style="background:#0b5ed7;padding:12px;border-radius:8px;color:white"><h2>📦 Sistema de Consolidación de Listas de Empaque CMP</h2></div>', unsafe_allow_html=True)
st.write("Sube primero el archivo **CONSOLIDADO.xlsx** y luego una o varias **Listas de Empaque**. Marca 'Vista previa' para revisar las primeras 10 filas de cada lista antes de consolidar.")

def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.upper()

def find_header_row_allow_combined(df_no_header, max_rows=30):
    # Search first max_rows rows for header keywords, including combined-row headers (row + next row)
    rows = min(max_rows, len(df_no_header))
    for i in range(rows):
        row_text = " ".join(str(x) for x in df_no_header.iloc[i].fillna("")).upper()
        if (("CAJA" in row_text and "PARTE" in row_text and ("NO" in row_text or "NRO" in row_text or "NÚM" in row_text or "NUMERO" in row_text))
            or ("CANTIDAD" in row_text and "EMPAC" in row_text)):
            return i
        if i+1 < len(df_no_header):
            combined = (row_text + " " + " ".join(str(x) for x in df_no_header.iloc[i+1].fillna(""))).upper()
            if (("CAJA" in combined and "PARTE" in combined and ("NO" in combined or "NRO" in combined or "NUMERO" in combined))
                or ("CANTIDAD" in combined and "EMPAC" in combined)):
                return i
    return None

def normalize_consolidado(df_cons):
    # map columns like DESPACHO, COD., DESCRIPCION ignoring dots/spaces/accents
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
    # find pattern like DC 083-25 or DC083-25 etc.
    name = filename.rsplit(".",1)[0]
    m = re.search(r'(DC[\s\-]*\d{1,3}[\-\s]*\d{1,3})', name, re.IGNORECASE)
    if m:
        s = m.group(1).upper().replace("_"," ").strip()
        # normalize spacing around DC and dash
        s = re.sub(r'\s+', ' ', s)
        # ensure up to 31 chars
        return s[:31]
    # fallback: shorter file-stem
    parts = re.split(r'[\s_]+', name)
    if len(parts) >= 2:
        candidate = " ".join(parts[-2:])[:31]
        return candidate
    return name[:31]

# Upload controls
st.subheader("1) Subir archivo CONSOLIDADO (.xlsx)")
cons_file = st.file_uploader("Selecciona CONSOLIDADO (DESPACHO | COD. | DESCRIPCION)", type=["xlsx"], key="cons")

st.subheader("2) Subir Listas de Empaque (.xlsx) - puedes seleccionar varias")
list_files = st.file_uploader("Selecciona listas de empaque", type=["xlsx"], accept_multiple_files=True, key="lists")

preview = st.checkbox("👁️ Ver vista previa (primeras 10 filas) de cada lista antes de consolidar")

process_button = st.button("🚀 Procesar y Consolidar")

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
        problemas = []
        previews = {}

        prog = st.progress(0)
        total = len(list_files)
        idx = 0

        for f in list_files:
            idx += 1
            try:
                file_bytes = f.getvalue()
                raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
            except Exception as e:
                problemas.append((f.name, f"Error lectura: {e}"))
                prog.progress(int(idx/total*100))
                continue

            header_row = find_header_row_allow_combined(raw, max_rows=30)
            if header_row is None:
                # maybe headers are combined across 2 rows with merged cells; still try joining two rows into a synthetic header
                problemas.append((f.name, "No se encontró fila de encabezado en las primeras 30 filas."))
                prog.progress(int(idx/total*100))
                continue

            try:
                df = pd.read_excel(io.BytesIO(file_bytes), header=header_row, dtype=str)
            except Exception as e:
                problemas.append((f.name, f"Error leyendo con header: {e}"))
                prog.progress(int(idx/total*100))
                continue

            # Normalize column names for detection (remove accents/spaces)
            df.columns = [normalize_text(c) for c in df.columns]

            col_caja = next((c for c in df.columns if "CAJA" in c), None)
            col_parte = next((c for c in df.columns if "PARTE" in c), None)
            col_cant = next((c for c in df.columns if "EMPAC" in c or "CANTIDAD" in c), None)

            if not col_caja or not col_parte or not col_cant:
                problemas.append((f.name, "Columnas esperadas No. de Caja / Número de Parte / Cantidad Empacada no encontradas después del encabezado."))
                prog.progress(int(idx/total*100))
                continue

            df_extract = df[[col_caja, col_parte, col_cant]].copy()
            df_extract.columns = ["No. de Caja", "Número de Parte", "Cantidad Empacada"]
            # Forward fill No. de Caja (fill down)
            df_extract["No. de Caja"] = df_extract["No. de Caja"].ffill()
            df_extract["Número de Parte"] = df_extract["Número de Parte"].astype(str).str.strip()

            # Merge with consolidated descriptions
            df_extract = df_extract.merge(df_cons, how="left", left_on="Número de Parte", right_on="COD.")
            df_extract["DESCRIPCION"] = df_extract["DESCRIPCION"].fillna("NO ENCONTRADO")

            # Add extra blank cols
            for c in ["CANTIDAD FISICA", "U/M", "U/M POR CADA", "ORDEN DE PRODUCCION", "LOTE", "OBSERVACION"]:
                df_extract[c] = ""

            short_name = extract_short_sheet_name(f.name)
            resultados[short_name] = df_extract
            previews[short_name] = df_extract.head(10)

            prog.progress(int(idx/total*100))

        if not resultados:
            st.warning("No se procesó ninguna lista. Revisa los problemas listados abajo.")
            for p in problemas:
                st.write("•", p[0], "-", p[1])
        else:
            if preview:
                st.subheader("Vista previa (primeras 10 filas) de cada lista")
                for name, dfp in previews.items():
                    st.markdown(f"**{name}**")
                    st.dataframe(dfp)

            # Build Excel in memory (one sheet per list) + optional consolidated sheet
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # write each sheet
                for sheet, df_out in resultados.items():
                    df_out.to_excel(writer, sheet_name=sheet[:31], index=False)
                # also write a merged "CONSOLIDADO" sheet if desired - here we include merged rows
                try:
                    merged = pd.concat(resultados.values(), ignore_index=True)
                    # drop duplicates maybe? keep as is
                    merged.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
                except Exception:
                    pass
            output.seek(0)
            filename = f"CONSOLIDADO_CMP_{date.today()}.xlsx"
            st.success("Consolidación lista ✅")
            st.download_button("⬇️ Descargar Excel consolidado", data=output.getvalue(), file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if problemas:
                st.warning("Algunos archivos tuvieron problemas:")
                for p in problemas:
                    st.write("•", p[0], "-", p[1])
