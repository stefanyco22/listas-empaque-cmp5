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

def detectar_columnas_consolidado(df):
    """Detecta automáticamente las columnas DESPACHO, COD. y DESCRIPCION"""
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
    
    return despacho_idx, cod_idx, descripcion_idx

def procesar_archivo_consolidado(uploaded_file):
    """Procesa el archivo CONSOLIDADO.xlsx"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # Detectar columnas automáticamente
        despacho_idx, cod_idx, descripcion_idx = detectar_columnas_consolidado(df)
        
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

def encontrar_fila_encabezado(df):
    """Encuentra la fila de encabezado, manejando encabezados de 2 filas"""
    for i in range(min(5, len(df))):  # Revisar primeras 5 filas
        fila_actual = df.iloc[i].astype(str).apply(normalizar_texto).tolist()
        fila_siguiente = df.iloc[i+1].astype(str).apply(normalizar_texto).tolist() if i+1 < len(df) else []
        
        # Buscar patrones de columnas esperadas
        patrones = ['CAJA', 'PARTE', 'CANTIDAD', 'EMPACADA']
        
        # Verificar si la fila actual contiene patrones
        encontrados_actual = sum(1 for patron in patrones if any(patron in celda for celda in fila_actual))
        
        # Verificar combinación con fila siguiente
        combinacion = fila_actual + fila_siguiente
        encontrados_combinacion = sum(1 for patron in patrones if any(patron in celda for celda in combinacion))
        
        if encontrados_actual >= 2:
            return i, 1  # Encabezado en una fila
        elif encontrados_combinacion >= 3:
            return i, 2  # Encabezado en dos filas
    
    return None, None

def detectar_columnas_lista(df, fila_inicio, num_filas_encabezado):
    """Detecta las columnas No. de Caja, Número de Parte y Cantidad Empacada"""
    if num_filas_encabezado == 1:
        encabezados = df.iloc[fila_inicio].astype(str).apply(normalizar_texto).tolist()
    else:  # 2 filas
        encabezados = []
        for j in range(len(df.columns)):
            celda1 = normalizar_texto(df.iloc[fila_inicio, j])
            celda2 = normalizar_texto(df.iloc[fila_inicio + 1, j])
            encabezado_combinado = f"{celda1} {celda2}".strip()
            encabezados.append(encabezado_combinado)
    
    # Buscar columnas
    caja_idx = None
    parte_idx = None
    cantidad_idx = None
    
    for i, encabezado in enumerate(encabezados):
        encabezado_clean = encabezado.upper().replace(' ', '')
        
        if any(palabra in encabezado for palabra in ['CAJA', 'NO', 'NUMERO', 'NRO']):
            caja_idx = i
        elif any(palabra in encabezado for palabra in ['PARTE', 'CODIGO', 'COD', 'ARTICULO']):
            parte_idx = i
        elif any(palabra in encabezado for palabra in ['CANTIDAD', 'EMPACADA', 'CANT', 'QTY']):
            cantidad_idx = i
    
    return caja_idx, parte_idx, cantidad_idx

def procesar_lista_empaque(df, nombre_archivo, df_consolidado):
    """Procesa una lista de empaque individual"""
    try:
        # Encontrar fila de encabezado
        fila_encabezado, num_filas = encontrar_fila_encabezado(df)
        
        if fila_encabezado is None:
            return None, f"No se pudo encontrar la fila de encabezado en {nombre_archivo}"
        
        # Detectar columnas
        caja_idx, parte_idx, cantidad_idx = detectar_columnas_lista(df, fila_encabezado, num_filas)
        
        if None in [caja_idx, parte_idx, cantidad_idx]:
            return None, f"No se pudieron detectar todas las columnas necesarias en {nombre_archivo}"
        
        # Leer datos (saltando encabezados)
        datos_inicio = fila_encabezado + num_filas
        df_datos = df.iloc[datos_inicio:].copy()
        
        if len(df_datos) == 0:
            return None, f"No hay datos después del encabezado en {nombre_archivo}"
        
        # Seleccionar y renombrar columnas
        df_procesado = df_datos.iloc[:, [caja_idx, parte_idx, cantidad_idx]].copy()
        df_procesado.columns = ['NO_DE_CAJA', 'NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA']
        
        # Limpiar datos
        df_procesado = df_procesado.dropna(subset=['NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA'])
        df_procesado['NO_DE_CAJA'] = df_procesado['NO_DE_CAJA'].ffill()  # Forward fill para cajas
        df_procesado['NUMERO_DE_PARTE'] = df_procesado['NUMERO_DE_PARTE'].apply(normalizar_texto)
        
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
        
        # Agregar columnas adicionales
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
    st.markdown("Consolida múltiples listas de empaque con un archivo principal CONSOLIDADO.xlsx")
    
    # Sidebar para configuraciones
    with st.sidebar:
        st.header("Configuración")
        mostrar_preview = st.checkbox("Mostrar vista previa", value=True)
        mostrar_consolidado_preview = st.checkbox("Mostrar preview del consolidado", value=True)
    
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
        
        if df_consolidado is not None and mostrar_consolidado_preview:
            st.subheader("Vista previa del archivo consolidado")
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
        st.success(f"{len(archivos_listas)} archivo(s) de lista de empaque cargado(s)")
        
        # Procesar listas
        resultados = {}
        errores = []
        
        for archivo in archivos_listas:
            with st.spinner(f"Procesando {archivo.name}..."):
                try:
                    df_lista = pd.read_excel(archivo)
                    df_procesado, error = procesar_lista_empaque(df_lista, archivo.name, df_consolidado)
                    
                    if error:
                        errores.append(error)
                    else:
                        resultados[archivo.name] = df_procesado
                        
                        if mostrar_preview:
                            with st.expander(f"Vista previa: {archivo.name}"):
                                st.dataframe(df_procesado.head(10))
                
                except Exception as e:
                    errores.append(f"Error leyendo {archivo.name}: {str(e)}")
        
        # Mostrar errores
        if errores:
            st.error("Errores encontrados:")
            for error in errores:
                st.write(f"• {error}")
        
        # Generar archivo final si hay resultados
        if resultados:
            st.header("3. Archivo Final Consolidado")
            
            # Crear Excel en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Hojas individuales
                for nombre, df in resultados.items():
                    # Limpiar nombre de hoja (máx 31 caracteres, sin caracteres inválidos)
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
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success(f"Archivo generado: {nombre_descarga}")
            st.info(f"Total de listas procesadas: {len(resultados)}")
            st.info(f"Total de registros consolidados: {len(df_consolidado_final)}")

if __name__ == "__main__":
    main()

