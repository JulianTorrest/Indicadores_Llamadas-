import pandas as pd
import streamlit as st
import plotly.express as px
import os

st.set_page_config(page_title="Dashboard Campañas", layout="wide")

NOMBRE_ARCHIVO = "Seguimiento encuestas consolidado.xlsx"

@st.cache_data(show_spinner="Cargando datos...")
def cargar_datos(hoja):
    try:
        df = pd.read_excel(NOMBRE_ARCHIVO, sheet_name=hoja)
        # Limpieza básica: convertir fechas si existen
        if 'Marca temporal' in df.columns:
            df['Marca temporal'] = pd.to_datetime(df['Marca temporal'])
        return df
    except Exception as e:
        st.error(f"Error al cargar la hoja {hoja}: {e}")
        return None

def main():
    st.sidebar.title("Navegación")
    opcion = st.sidebar.selectbox("Seleccione Vista", 
        ["Resumen Ejecutivo", "Análisis por Encuestador", "Auditoría Técnica (Twilio/Cámara)"])

    if not os.path.exists(NOMBRE_ARCHIVO):
        st.error(f"No se encontró el archivo '{NOMBRE_ARCHIVO}'")
        return

    if opcion == "Resumen Ejecutivo":
        st.title("📊 Resumen Ejecutivo de Campañas")
        df_gestiones = cargar_datos("Base_gestiones realizadas")
        
        if df_gestiones is not None:
            # KPIs superiores
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Gestiones", len(df_gestiones))
            
            # Resultado de Gestión (Gráfico de Barras)
            if "Resultado de la gestión" in df_gestiones.columns:
                st.subheader("Distribución de Resultados")
                fig = px.bar(df_gestiones["Resultado de la gestión"].value_counts(), 
                             labels={'value': 'Cantidad', 'index': 'Resultado'})
                st.plotly_chart(fig, use_container_width=True)
            
            # Molestia del cliente (Gráfico de Pastel)
            campo_molestia = "¿Cree que la persona se sintió molesta por la llamada? (si la persona no contestó marcar 1)"
            if campo_molestia in df_gestiones.columns:
                st.subheader("Nivel de Molestia Percibida")
                fig_pie = px.pie(df_gestiones, names=campo_molestia)
                st.plotly_chart(fig_pie)

    elif opcion == "Auditoría Técnica (Twilio/Cámara)":
        st.title("🛡️ Auditoría de Llamadas")
        df_twilio = cargar_datos("Twilio")
        if df_twilio is not None:
            st.write("Últimos registros de Twilio", df_twilio.head())
            if "Status" in df_twilio.columns:
                st.write("**Estado de llamadas (Twilio):**")
                st.bar_chart(df_twilio["Status"].value_counts())

if __name__ == "__main__":
    main()
