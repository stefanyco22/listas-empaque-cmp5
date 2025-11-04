import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime
import unicodedata

def normalizar_texto(texto):
    """Normaliza texto: convierte a mayúsculas, elimina acentos y puntos"""
    if pd.isna(texto):
        return ""
    texto = str(texto)
    # Eliminar acentos
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    # Convertir a mayúsculas y eliminar puntos
    texto = texto.upper().replace('.', '')
    return texto

def procesar_archivo_consolidado(uploaded_file):
    """Procesa el archivo CONSOLIDADO.xlsx"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # Buscar columnas por nombres esperados
        columnas = df.columns.tolist()
        columnas_normalizadas = [normalizar_texto(col) for col in columnas]
        
        # Buscar DESPACHO
        despacho_idx = None
        for i, col in enumerate(columnas_normalizadas):
            if 'DESPACHO' in col:
                despacho_idx = i
                break
        
        # Buscar COD o CODIGO
        cod_idx = None
        for i, col in enumerate(columnas_normalizadas):
            if 'COD' in col and 'DESCRIPCION' not in col:
                cod_idx = i
                break
        
        # Buscar DESCRIPCION
        descripcion_idx = None
        for i, col in enumerate(columnas_normalizadas):
            if 'DESCRIPCION' in col:
                descripcion_idx = i
                break
        
        if despacho_idx is None or cod_idx is None or descripcion_idx is None:
            st.error("No se pudieron detectar las columnas necesarias en el archivo consolidado")
            st.write("Columnas encontradas:", df.columns.tolist())
            return None
        
        # Renombrar columnas
        df_consolidado = df.iloc[:, [despacho_idx, cod_idx, descripcion_idx]].copy()
        df_consolidado.columns = ['DESPACHO', 'COD', 'DESCRIPCION']
        
        # Normalizar datos
        df_consolidado['DESPACHO'] = df_consolidado['DESPACHO'].apply(normalizar_texto)
        df_consolidado['COD'] = df_consolidado['COD'].apply(normalizar_texto)
        df_consolidado['DESCRIPCION'] = df_consolidado['DESCRIPCION'].apply(normalizar_texto)
        
        # Eliminar filas vacías
        df_consolidado = df_consolidado.dropna(subset=['COD', 'DESCRIPCION'])
        
        st.success(f"Archivo consolidado procesado: {len(df_consolidado)} registros")
        return df_consolidado
    
    except Exception as e:
        st.error(f"Error al procesar archivo consolidado: {str(e)}")
        return None

def procesar_lista_empaque_simple(df, nombre_archivo, df_consolidado):
    """Procesa lista de empaque con estructura fija conocida - VERSIÓN SIMPLIFICADA"""
    try:
        st.write(f"🔍 Procesando: {nombre_archivo}")
        
        # ESTRUCTURA FIJA CONOCIDA:
        # - Datos empiezan en fila 17 (índice 16 en base 0)
        # - Columna A (0): No. de Caja/PALLET
        # - Columna B (1): Número de Parte  
        # - Columna D (3): Cantidad Empacada
        
        if len(df) < 17:
            return None, f"El archivo {nombre_archivo} no tiene suficientes filas"
        
        # Empezar directamente desde la fila 17 (índice 16)
        datos_inicio = 16
        df_datos = df.iloc[datos_inicio:].copy()
        
        st.write(f"📊 Leyendo datos desde fila {datos_inicio + 1}")
        
        if len(df_datos) == 0:
            return None, f"No hay datos en {nombre_archivo}"
        
        # Tomar solo las columnas que necesitamos
        df_procesado = df_datos.iloc[:, [0, 1, 3]].copy()  # Columnas A, B, D
        df_procesado.columns = ['NO_DE_CAJA', 'NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA']
        
        # Limpiar datos
        df_procesado = df_procesado.dropna(how='all')  # Eliminar filas completamente vacías
        df_procesado = df_procesado.dropna(subset=['NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA'])  # Eliminar filas sin datos esenciales
        
        # Rellenar valores de PALLET hacia abajo
        df_procesado['NO_DE_CAJA'] = df_procesado['NO_DE_CAJA'].ffill()
        
        # Normalizar números de parte
        df_procesado['NUMERO_DE_PARTE'] = df_procesado['NUMERO_DE_PARTE'].apply(normalizar_texto)
        
        # Convertir cantidad a numérico
        df_procesado['CANTIDAD_EMPACADA'] = pd.to_numeric(df_procesado['CANTIDAD_EMPACADA'], errors='coerce')
        df_procesado = df_procesado.dropna(subset=['CANTIDAD_EMPACADA'])
        
        st.write(f"✅ {len(df_procesado)} registros procesados")
        
        # Unir con consolidado
        df_final = pd.merge(
            df_procesado, 
            df_consolidado[['COD', 'DESCRIPCION']], 
            left_on='NUMERO_DE_PARTE', 
            right_on='COD', 
            how='left'
        )
        
        # Manejar no encontrados
        df_final['DESCRIPCION'] = df_final['DESCRIPCION'].fillna('NO ENCONTRADO')
        
        # Eliminar columna COD duplicada
        df_final = df_final.drop('COD', axis=1)
        
        # Agregar columnas adicionales vacías
        columnas_adicionales = ['CANTIDAD_FISICA', 'U/M', 'U/M_POR_CADA', 'ORDEN_DE_PRODUCCION', 'LOTE', 'OBSERVACION']
        for col in columnas_adicionales:
            df_final[col] = ""
        
        # Reordenar columnas
        columnas_ordenadas = ['NO_DE_CAJA', 'NUMERO_DE_PARTE', 'DESCRIPCION', 'CANTIDAD_EMPACADA'] + columnas_adicionales
        df_final = df_final[columnas_ordenadas]
        
        return df_final, None
        
    except Exception as e:
        return None, f"Error al procesar {nombre_archivo}: {str(e)}"

def main():
    st.set_page_config(page_title="Consolidador de Listas de Empaque", layout="wide")
    
    st.title("📦 Consolidador de Listas de Empaque")
    st.markdown("### Estructura Fija: Datos desde fila 17, Columnas A/B/D")
    
    # Información de estructura
    with st.expander("ℹ️ Estructura Esperada", expanded=True):
        st.markdown("""
        **Formato de Listas de Empaque:**
        - 📍 **Encabezados:** Filas 14-15
        - 📍 **Fila vacía:** 16
        - 📍 **Datos:** Desde fila 17
        - 📍 **Columnas:**
          - A: No. de Caja/PALLET
          - B: Número de Parte  
          - D: Cantidad Empacada
        """)
    
    # Sección 1: Archivo CONSOLIDADO
    st.header("1. Archivo CONSOLIDADO.xlsx")
    archivo_consolidado = st.file_uploader(
        "Sube el archivo CONSOLIDADO.xlsx", 
        type=['xlsx'],
        key='consolidado'
    )
    
    df_consolidado = None
    if archivo_consolidado is not None:
        df_consolidado = procesar_archivo_consolidado(archivo_consolidado)
        
        if df_consolidado is not None:
            st.subheader("Vista previa del consolidado")
            st.dataframe(df_consolidado.head(10))
    
    # Sección 2: Listas de Empaque
    st.header("2. Listas de Empaque")
    archivos_listas = st.file_uploader(
        "Sube las listas de empaque (.xlsx)", 
        type=['xlsx'], 
        accept_multiple_files=True,
        key='listas'
    )
    
    if df_consolidado is not None and archivos_listas:
        st.success(f"✅ {len(archivos_listas)} archivo(s) cargado(s)")
        
        # Procesar listas
        resultados = {}
        errores = []
        
        for archivo in archivos_listas:
            with st.spinner(f"Procesando {archivo.name}..."):
                try:
                    df_lista = pd.read_excel(archivo)
                    df_procesado, error = procesar_lista_empaque_simple(df_lista, archivo.name, df_consolidado)
                    
                    if error:
                        errores.append(error)
                    else:
                        resultados[archivo.name] = df_procesado
                        st.success(f"✅ {archivo.name}: {len(df_procesado)} registros")
                
                except Exception as e:
                    errores.append(f"Error leyendo {archivo.name}: {str(e)}")
        
        # Mostrar resultados
        if resultados:
            st.header("3. Resultados")
            
            total_registros = sum(len(df) for df in resultados.values())
            st.success(f"🎉 Procesamiento completado: {len(resultados)} archivos, {total_registros} registros")
            
            # Mostrar preview del primer archivo
            primer_archivo = list(resultados.keys())[0]
            st.subheader(f"Vista previa: {primer_archivo}")
            st.dataframe(resultados[primer_archivo].head(10))
            
            # Generar archivo final
            st.header("4. Descargar Archivo Consolidado")
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Hojas individuales
                for nombre, df in resultados.items():
                    nombre_hoja = re.sub(r'[\\/*?:\[\]]', '', nombre)[:31]
                    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
                
                # Hoja consolidada
                df_consolidado_final = pd.concat(resultados.values(), ignore_index=True)
                df_consolidado_final.to_excel(writer, sheet_name='CONSOLIDADO', index=False)
            
            output.seek(0)
            
            # Botón de descarga
            fecha_actual = datetime.now().strftime("%Y-%m-%d")
            nombre_descarga = f"CONSOLIDADO_CMP_{fecha_actual}.xlsx"
            
            st.download_button(
                label="📥 Descargar Archivo Consolidado",
                data=output,
                file_name=nombre_descarga,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        if errores:
            st.error("Errores encontrados:")
            for error in errores:
                st.write(f"• {error}")

if __name__ == "__main__":
    main()
