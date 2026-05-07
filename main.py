import pandas as pd
import streamlit as st
import os

@st.cache_data
def obtener_diccionario_campos(ruta_archivo):
    """
    Lee un archivo Excel y devuelve un diccionario donde las llaves son 
    los nombres de las hojas y los valores son las listas de columnas.
    """
    try:
        # Cargamos el archivo Excel
        # Usamos pd.ExcelFile para evitar leer todos los datos de golpe y ahorrar memoria
        xls = pd.ExcelFile(ruta_archivo)
        
        diccionario_campos = {}
        
        for nombre_hoja in xls.sheet_names:
            # Leemos solo la primera fila (encabezados) para mayor eficiencia
            df = pd.read_excel(xls, sheet_name=nombre_hoja, nrows=0)
            diccionario_campos[nombre_hoja] = df.columns.tolist()
            
        return diccionario_campos

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return None

st.set_page_config(page_title="Analizador de Campañas", layout="wide")

st.title("📊 Estructura de Campañas de Llamadas")
st.info("Leyendo datos directamente desde el repositorio de GitHub.")

# Al estar en Streamlit Cloud, usamos el nombre del archivo directamente si está en la raíz
NOMBRE_ARCHIVO = "Seguimiento encuestas consolidado.xlsx"

if os.path.exists(NOMBRE_ARCHIVO):
    campos_hojas = obtener_diccionario_campos(NOMBRE_ARCHIVO)
    
    if campos_hojas:
        st.subheader("Estructura detectada (Formato JSON)")
        st.json(campos_hojas)
        
        st.subheader("Desglose por Hoja")
        for hoja, columnas in campos_hojas.items():
            with st.expander(f"Hoja: {hoja}"):
                st.write(f"**Total de campos:** {len(columnas)}")
                st.write(columnas)
else:
    st.error(f"No se encontró el archivo '{NOMBRE_ARCHIVO}' en el repositorio.")
