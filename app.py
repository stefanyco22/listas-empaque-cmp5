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

def procesar_lista_empaque_estructura_fija(df, nombre_archivo, df_consolidado):
    """Procesa lista de empaque con estructura fija conocida"""
    try:
        st.write(f"🔍 Procesando: {nombre_archivo} (estructura fija)")
        
        # Estructura fija conocida:
        # - Encabezados en filas 14 y 15 (índices 13 y 14 en base 0)
        # - Fila 16 vacía (índice 15)
        # - Datos desde fila 17 (índice 16)
        
        if len(df) < 17:
            return None, f"El archivo {nombre_archivo} no tiene suficientes filas (mínimo 17 requeridas)"
        
        # Verificar que tenemos los encabezados correctos en las filas 14 y 15
        fila_14 = df.iloc[13].astype(str).apply(normalizar_texto).tolist()
        fila_15 = df.iloc[14].astype(str).apply(normalizar_texto).tolist()
        
        st.write("Fila 14 (encabezado):", fila_14)
        st.write("Fila 15 (encabezado):", fila_15)
        
        # Asignar columnas fijas según la estructura conocida
        # Columna A (0): No. de Caja / PALLET
        # Columna B (1): Número de Parte
        # Columna D (3): Cantidad Empacada
        
        caja_idx = 0      # Columna A
        parte_idx = 1     # Columna B  
        cantidad_idx = 3  # Columna D
        
        st.write(f"Columnas fijas - Caja: {caja_idx} (A), Parte: {parte_idx} (B), Cantidad: {cantidad_idx} (D)")
        
        # Leer datos desde fila 17 (índice 16)
        datos_inicio = 16
        df_datos = df.iloc[datos_inicio:].copy()
        
        st.write(f"📊 Datos desde fila {datos_inicio + 1}, total filas: {len(df_datos)}")
        
        if len(df_datos) == 0:
            return None, f"No hay datos después de la fila 17 en {nombre_archivo}"
        
        # Seleccionar y renombrar columnas
        df_procesado = df_datos.iloc[:, [caja_idx, parte_idx, cantidad_idx]].copy()
        df_procesado.columns = ['NO_DE_CAJA', 'NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA']
        
        # Limpiar datos
        # Eliminar filas completamente vacías
        df_procesado = df_procesado.dropna(how='all')
        
        # Eliminar filas donde falta número de parte o cantidad
        df_procesado = df_procesado.dropna(subset=['NUMERO_DE_PARTE', 'CANTIDAD_EMPACADA'])
        
        # Rellenar valores de PALLET hacia abajo (forward fill)
        df_procesado['NO_DE_CAJA'] = df_procesado['NO_DE_CAJA'].ffill()
        
        # Normalizar números de parte
        df_procesado['NUMERO_DE_PARTE'] = df_procesado['NUMERO_DE_PARTE'].apply(normalizar_texto)
        
        # Convertir cantidad a numérico y limpiar
        df_procesado['CANTIDAD_EMPACADA'] = pd.to_numeric(df_procesado['CANTIDAD_EMPACADA'], errors='coerce')
        df_procesado = df_procesado.dropna(subset=['CANTIDAD_EMPACADA'])
        
        st.write(f"✅ Datos procesados: {len(df_procesado)} registros")
        
        # Mostrar preview de datos procesados
        st.write("Preview de datos procesados:")
        st.dataframe(df_procesado.head(10))
        
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
        
        st.write("✅ Lista procesada exitosamente con estructura fija")
        return df_final, None
        
    except Exception as e:
        st.error(f"❌ Error procesando {nombre_archivo} con estructura fija: {str(e)}")
        import traceback
        st.write("Detalles del error:", traceback.format_exc())
        return None, f"Error al procesar {nombre_archivo}: {str(e)}"

def main():
    st.set_page_config(page_title="Consolidador de Listas de Empaque", layout="wide")
    
    st.title("📦 Consolidador de Listas de Empaque - ESTRUCTURA FIJA")
    st.markdown("**Estructura esperada:** Encabezados en filas 14-15, datos desde fila 17, columnas A/B/D")
    
    # Sidebar para configuraciones
    with st.sidebar:
        st.header("Configuración")
        mostrar_preview = st.checkbox("Mostrar vista previa", value=True)
        mostrar_consolidado_preview = st.checkbox("Mostrar preview del consolidado", value=True)
        modo_debug = st.checkbox("Modo Debug (muestra detalles)", value=True)
    
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
    st.info("📋 **Formato esperado:** Encabezados en filas 14-15, fila 16 vacía, datos desde fila 17")
    st.info("📍 **Columnas:** A=No. de Caja/PALLET, B=Número de Parte, D=Cantidad Empacada")
    
    archivos_listas = st.file_uploader(
        "Sube las listas de empaque (.xlsx) - Mismo formato", 
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
                    
                    # Usar el procesador de estructura fija
                    df_procesado, error = procesar_lista_empaque_estructura_fija(df_lista, archivo.name, df_consolidado)
                    
                    if error:
                        errores.append(error)
                    else:
                        resultados[archivo.name] = df_procesado
                        
                        if mostrar_preview:
                            with st.expander(f"✅ Vista previa: {archivo.name}"):
                                st.dataframe(df_procesado.head(15))
                                st.write(f"Total registros: {len(df_procesado)}")
                
                except Exception as e:
                    errores.append(f"Error leyendo {archivo.name}: {str(e)}")
        
        # Mostrar resumen
        st.header("3. Resumen del Procesamiento")
        
        if resultados:
            st.success(f"✅ {len(resultados)} lista(s) procesada(s) exitosamente")
            total_registros = sum(len(df) for df in resultados.values())
            st.info(f"📊 Total de registros consolidados: {total_registros}")
            
            # Mostrar estadísticas por PALLET
            st.subheader("📦 Distribución por PALLET")
            todos_los_datos = pd.concat(resultados.values(), ignore_index=True)
            conteo_pallet = todos_los_datos['NO_DE_CAJA'].value_counts()
            st.dataframe(conteo_pallet)
        
        if errores:
            st.error(f"❌ {len(errores)} error(es) encontrado(s):")
            for error in errores:
                st.write(f"• {error}")
        
        # Generar archivo final si hay resultados
        if resultados:
            st.header("4. Archivo Final Consolidado")
            
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
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.success(f"🎉 Archivo generado: {nombre_descarga}")

if __name__ == "__main__":
    main()
