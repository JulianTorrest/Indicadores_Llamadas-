
import pandas as pd
import streamlit as st
import os

@st.cache_data(show_spinner="Analizando estructura y categorías...")
def analizar_estructura_completa(ruta_archivo):
    """
    Analiza el Excel para obtener columnas, tipos de datos y opciones de respuesta
    estandarizadas (con pocos valores únicos).
    """
    try:
        xls = pd.ExcelFile(ruta_archivo)
        estructura = {}
        UMBRAL_ESTANDARIZACION = 15  # Si hay más de 15 valores únicos, no se considera "estandarizado"

        for nombre_hoja in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=nombre_hoja)
            detalles_columnas = {}
            
            for col in df.columns:
                # Obtener valores únicos eliminando nulos
                valores_unicos = df[col].dropna().unique()
                tipo_dato = str(df[col].dtype)
                
                info = {
                    "tipo": tipo_dato,
                    "opciones": None
                }
                
                # Solo incluimos opciones si el campo parece estar estandarizado (pocos valores únicos)
                if 0 < len(valores_unicos) <= UMBRAL_ESTANDARIZACION:
                    info["opciones"] = list(valores_unicos)
                
                detalles_columnas[col] = info
                
            estructura[nombre_hoja] = detalles_columnas

        return estructura

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return None

st.set_page_config(page_title="Diccionario de Datos Avanzado", layout="wide")

st.title("📊 Diccionario de Campos y Categorías")
st.info("Leyendo datos directamente desde el repositorio de GitHub.")

NOMBRE_ARCHIVO = "Seguimiento encuestas consolidado.xlsx"

if os.path.exists(NOMBRE_ARCHIVO):
    estructura = analizar_estructura_completa(NOMBRE_ARCHIVO)

    if estructura:
        st.subheader("Exploración por Hoja")
        
        for hoja, columnas in estructura.items():
            with st.expander(f"📂 Hoja: {hoja}"):
                st.write(f"Esta hoja contiene **{len(columnas)}** campos.")
                
                # Crear una tabla para mostrar los tipos de datos de forma limpia
                datos_tabla = []
                for col, info in columnas.items():
                    opciones_str = ", ".join(map(str, info["opciones"])) if info["opciones"] else "N/A (Campo no estandarizado o abierto)"
                    datos_tabla.append({
                        "Campo": col,
                        "Tipo de Dato": info["tipo"],
                        "Opciones de Respuesta": opciones_str
                    })
                
                st.table(pd.DataFrame(datos_tabla))
                
        st.subheader("Estructura técnica (JSON)")
        st.json(estructura)
else:
    st.error(f"No se encontró el archivo '{NOMBRE_ARCHIVO}' en el repositorio.")
