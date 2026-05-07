
import pandas as pd
import streamlit as st
import os
import unicodedata
import plotly.express as px

def limpiar_texto(texto):
    if not isinstance(texto, str): return texto
    texto = texto.strip()
    # Eliminar tildes
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto.title()

def agrupar_resultado_gestion(valor):
    valor = str(valor).lower()
    if "responde la encuesta" in valor or "encuesta incompleta" in valor:
        return "Éxito / Parcial"
    elif "no contesta" in valor or "sin respuesta" in valor or "no entra" in valor or "invalido" in valor:
        return "No Contactado"
    elif "rechaza" in valor or "rechazo" in valor:
        return "Rechazo"
    elif "pide que" in valor or "llamen despues" in valor or "formulario online" in valor:
        return "Seguimiento"
    elif "no hace parte de universo" in valor:
        return "Fuera de Perfil"
    elif "no aplica" in valor: # Para manejar posibles N/A o vacíos en la agrupación
        return "No Aplica"
    return "Otros"

@st.cache_data(show_spinner="Cargando y limpiando datos...")
def cargar_y_limpiar_datos(ruta_archivo):
    try:
        xls = pd.ExcelFile(ruta_archivo)
        datos_limpios = {}
        
        for nombre_hoja in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=nombre_hoja)
            
            # Limpieza específica para hojas de gestión y contacto
            if nombre_hoja in ["Base_gestiones realizadas", "Contactados"]:
                # 1. Estandarizar nombres de encuestadores
                col_nombre = "Nombre del o de la encuestadora"
                if col_nombre in df.columns:
                    df[col_nombre] = df[col_nombre].apply(limpiar_texto)
                
                # 2. Agrupar Resultados de la gestión
                col_resultado = "Resultado de la gestión"
                if col_resultado in df.columns:
                    df[f"{col_resultado} (Agrupado)"] = df[col_resultado].apply(agrupar_resultado_gestion)
            
            # Limpieza específica para hojas de Entregas y Correo Masivo
            elif nombre_hoja in ["Entregados", "Correo Masivo"]:
                # Estandarizar nombres de encuestadores
                col_encuestador = "Encuestador"
                if col_encuestador in df.columns:
                    df[col_encuestador] = df[col_encuestador].apply(limpiar_texto)
                
                # Estandarizar Constructora
                col_constructora = "Constructora"
                if col_constructora in df.columns:
                    df[col_constructora] = df[col_constructora].apply(limpiar_texto)
                
                # Estandarizar Ciudad
                col_ciudad = "Ciudad"
                if col_ciudad in df.columns:
                    df[col_ciudad] = df[col_ciudad].apply(limpiar_texto)
                    df[col_ciudad] = df[col_ciudad].replace({"Bogota D.c.": "Bogota"}) # Unificar "Bogota D.C."
                
                # Estandarizar Ciudad2
                col_ciudad2 = "Ciudad2"
                if col_ciudad2 in df.columns:
                    df[col_ciudad2] = df[col_ciudad2].apply(limpiar_texto)
            
            # Limpieza específica para la hoja 'Camara_llamadas_salientes'
            elif nombre_hoja == "Camara_llamadas_salientes":
                # Estandarizar nombres de encuestadores
                if "encuestador" in df.columns:
                    df["encuestador"] = df["encuestador"].apply(limpiar_texto)
                
                # Estandarizar fecha de llamada
                if "fecha_llamada" in df.columns:
                    df["fecha_llamada"] = pd.to_datetime(df["fecha_llamada"], errors='coerce')

            # Limpieza específica para la hoja 'Twilio'
            elif nombre_hoja == "Twilio":
                # Estandarizar fechas
                for col_fecha in ["Start Time", "End Time", "Date Created"]:
                    if col_fecha in df.columns:
                        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')

            datos_limpios[nombre_hoja] = df
        return datos_limpios
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return None

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

st.set_page_config(page_title="Dashboard de Campañas y Entregas", layout="wide")

st.title("📊 Dashboard de Campañas y Entregas")
st.info("Analizando datos directamente desde el repositorio de GitHub.")

NOMBRE_ARCHIVO = "Seguimiento encuestas consolidado.xlsx"

if os.path.exists(NOMBRE_ARCHIVO):
    # Proceso de carga y limpieza
    datos = cargar_y_limpiar_datos(NOMBRE_ARCHIVO)
    estructura = analizar_estructura_completa(NOMBRE_ARCHIVO)

    if datos:
        st.sidebar.header("Filtros Globales")
        vistas = ["Resumen de KPIs Críticos", "Base_gestiones realizadas", "Contactados", "Entregados", "Correo Masivo", "Camara_llamadas_salientes", "Twilio"]
        hoja_seleccionada = st.sidebar.selectbox("Seleccione la fuente de datos:", vistas)

        if hoja_seleccionada == "Resumen de KPIs Críticos":
            st.header("🚀 Indicadores Críticos de Campaña")
            
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            
            if not df_g.empty:
                # 1. Tasa de Contactabilidad
                total = len(df_g)
                no_contactados = len(df_g[df_g["Resultado de la gestión (Agrupado)"] == "No Contactado"])
                contactabilidad = ((total - no_contactados) / total) * 100 if total > 0 else 0
                kpi1.metric("Contactabilidad", f"{contactabilidad:.1f}%", help="Porcentaje de gestiones que resultaron en un contacto real")

                # 2. Tasa de Conversión
                exitos = len(df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito / Parcial"])
                conversion = (exitos / (total - no_contactados)) * 100 if (total - no_contactados) > 0 else 0
                kpi2.metric("Conversión", f"{conversion:.1f}%", help="Encuestas exitosas sobre el total de personas contactadas")

            if not df_t.empty:
                # 3. AHT (Tiempo Promedio)
                if "Duration" in df_t.columns:
                    aht = df_t[df_t["Duration"] > 0]["Duration"].mean()
                    kpi3.metric("AHT (Segundos)", f"{aht:.1f}s", help="Tiempo promedio de conversación")

                # 4. Eficiencia de Costos
                if "Price" in df_t.columns and not df_g.empty and exitos > 0:
                    costo_total = df_t["Price"].abs().sum()
                    cpx = costo_total / exitos
                    kpi4.metric("Costo por Éxito", f"USD {cpx:.2f}")

            st.divider()
            
            col_l, col_r = st.columns(2)
            with col_l:
                st.subheader("Análisis de Discrepancia")
                # Comparar reporte manual vs técnico
                manual_count = len(df_g)
                tech_count = len(df_t)
                st.write(f"Gestiones manuales: **{manual_count}**")
                st.write(f"Llamadas técnicas (Twilio): **{tech_count}**")
                diff = manual_count - tech_count
                if diff > 0:
                    st.warning(f"Hay {diff} gestiones manuales sin respaldo en Twilio.")
                else:
                    st.success("El reporte manual coincide o es menor al registro técnico.")

            with col_r:
                st.subheader("Calidad y Sentimiento")
                campo_molestia = "¿Cree que la persona se sintió molesta por la llamada? (si la persona no contestó marcar 1)"
                if campo_molestia in df_g.columns:
                    avg_molestia = df_g[campo_molestia].mean()
                    st.progress(avg_molestia / 7.0)
                    st.write(f"Nivel de molestia promedio: **{avg_molestia:.2f} / 7**")
            
            # Gráfico comparativo de productividad por encuestador cruzado
            st.subheader("Productividad: Gestiones vs Duración Real")
            if not df_g.empty and not df_t.empty:
                # Nota: Este gráfico asume que podemos cruzar por nombre, 
                # para una versión pro, necesitaríamos una tabla de mapeo entre Nombres y Números de Twilio.
                st.info("Para un cruce exacto por encuestador, se recomienda vincular los números de 'From' en Twilio con los nombres.")

            st.stop() # Finaliza la vista de KPIs para no mostrar el resto

        df_base = datos[hoja_seleccionada]

        st.header(f"📈 Dashboard: {hoja_seleccionada}")
        
        # Metricas Principales
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Gestiones", len(df_base))

        # Métrica de Encuestadores Activos
        if hoja_seleccionada in ["Base_gestiones realizadas", "Contactados"]:
            col_encuestador_name = "Nombre del o de la encuestadora"
            if col_encuestador_name in df_base.columns:
                col2.metric("Encuestadores Activos", df_base[col_encuestador_name].nunique())
            else:
                col2.metric("Encuestadores Activos", "N/A")

            col_agrupado = "Resultado de la gestión (Agrupado)"
            if col_agrupado in df_base.columns:
                exito_parcial_count = (df_base[col_agrupado] == "Éxito / Parcial").sum()
                total_registros = len(df_base)
                if total_registros > 0:
                    porcentaje_exito = (exito_parcial_count / total_registros) * 100
                    col3.metric("Éxito / Parcial", f"{porcentaje_exito:.1f}%")
                else:
                    col3.metric("Éxito / Parcial", "0.0%")
            else:
                col3.metric("Éxito / Parcial", "N/A")

        elif hoja_seleccionada in ["Entregados", "Correo Masivo"]:
            col_encuestador_name = "Encuestador"
            if col_encuestador_name in df_base.columns:
                col2.metric("Encuestadores Activos", df_base[col_encuestador_name].nunique())
            else:
                col2.metric("Encuestadores Activos", "N/A")

            if "Proyecto" in df_base.columns:
                col3.metric("Proyectos Únicos", df_base["Proyecto"].nunique())
            else:
                col3.metric("Proyectos Únicos", "N/A")
        
        elif hoja_seleccionada == "Camara_llamadas_salientes":
            if "encuestador" in df_base.columns:
                col2.metric("Encuestadores Activos", df_base["encuestador"].nunique())
            else:
                col2.metric("Encuestadores Activos", "N/A")

            if "minutos_redondeados" in df_base.columns:
                total_minutos = df_base["minutos_redondeados"].sum()
                col3.metric("Total Minutos Hablados", f"{total_minutos:,}")
            else:
                col3.metric("Total Minutos", "N/A")

        elif hoja_seleccionada == "Twilio":
            col1.metric("Total Llamadas Twilio", len(df_base))
            
            if "Price" in df_base.columns:
                total_cost = df_base["Price"].abs().sum()
                col2.metric("Costo Total (USD)", f"${total_cost:.2f}")
            
            if "Duration" in df_base.columns:
                avg_dur = df_base["Duration"].mean()
                col3.metric("Duración Promedio", f"{avg_dur:.1f}s")

        # Visualizaciones de limpieza solicitadas
        if hoja_seleccionada in ["Base_gestiones realizadas", "Contactados"]:
            tab1, tab2 = st.tabs(["Análisis de Resultados", "Productividad Encuestadores"])
            
            with tab1:
                st.subheader("Resultados Estandarizados")
                col_agrupado = "Resultado de la gestión (Agrupado)"
                if col_agrupado in df_base.columns:
                    fig_res = px.pie(df_base, names=col_agrupado, 
                                     title="Distribución de Gestiones (Máximo 5 Categorías)")
                    st.plotly_chart(fig_res, use_container_width=True)
                    
                    with st.expander("Ver detalle de agrupación"):
                        st.dataframe(df_base[["Resultado de la gestión", col_agrupado]].drop_duplicates())
                else:
                    st.info(f"La columna '{col_agrupado}' no está disponible en esta hoja.")

            with tab2:
                st.subheader("Gestiones por Encuestador (Nombres Limpios)")
                col_encuestador = "Nombre del o de la encuestadora"
                if col_encuestador in df_base.columns:
                    encuestadores = df_base[col_encuestador].value_counts().reset_index()
                    fig_enc = px.bar(encuestadores, x=col_encuestador, y="count",
                                     labels={'count': 'Número de Llamadas'},
                                     color=col_encuestador)
                    st.plotly_chart(fig_enc, use_container_width=True)
                else:
                    st.info(f"La columna '{col_encuestador}' no está disponible en esta hoja.")

        elif hoja_seleccionada in ["Entregados", "Correo Masivo"]:
            tab1, tab2, tab3 = st.tabs(["Análisis General", "Productividad Encuestadores", "Distribución Geográfica"])

            with tab1:
                st.subheader(f"Análisis General: {hoja_seleccionada}")
                col_constructora = "Constructora"
                if col_constructora in df_base.columns:
                    constructora_counts = df_base[col_constructora].value_counts().reset_index()
                    fig_const = px.bar(constructora_counts, x=col_constructora, y="count",
                                       title="Entregados por Constructora",
                                       labels={'count': 'Número de Entregados'})
                    st.plotly_chart(fig_const, use_container_width=True)
                else:
                    st.info(f"La columna '{col_constructora}' no está disponible en esta hoja.")

            with tab2:
                st.subheader("Entregados por Encuestador (Nombres Limpios)")
                col_encuestador = "Encuestador"
                if col_encuestador in df_base.columns:
                    encuestadores = df_base[col_encuestador].value_counts().reset_index()
                    fig_enc = px.bar(encuestadores, x=col_encuestador, y="count",
                                     labels={'count': 'Número de Entregados'},
                                     color=col_encuestador)
                    st.plotly_chart(fig_enc, use_container_width=True)
                else:
                    st.info(f"La columna '{col_encuestador}' no está disponible en esta hoja.")

            with tab3:
                st.subheader("Distribución Geográfica de Entregados")
                col_ciudad = "Ciudad"
                col_ciudad2 = "Ciudad2"
                if col_ciudad in df_base.columns:
                    ciudad_counts = df_base[col_ciudad].value_counts().reset_index()
                    fig_ciudad = px.pie(ciudad_counts, names=col_ciudad,
                                        title="Entregados por Ciudad")
                    st.plotly_chart(fig_ciudad, use_container_width=True)
                elif col_ciudad2 in df_base.columns:
                    ciudad2_counts = df_base[col_ciudad2].value_counts().reset_index()
                    fig_ciudad2 = px.pie(ciudad2_counts, names=col_ciudad2,
                                         title="Entregados por Región (Ciudad2)")
                    st.plotly_chart(fig_ciudad2, use_container_width=True)
                else:
                    st.info("No se encontraron columnas de ciudad para analizar.")

        elif hoja_seleccionada == "Camara_llamadas_salientes":
            tab1, tab2 = st.tabs(["Estadísticas de Llamadas", "Productividad por Encuestador"])

            with tab1:
                st.subheader("Análisis de Tráfico Telefónico")
                col_duracion = "segundos"
                if col_duracion in df_base.columns:
                    avg_duration = df_base[col_duracion].mean()
                    st.write(f"**Duración promedio por llamada:** {avg_duration:.2f} segundos")
                    
                    fig_dur = px.histogram(df_base, x=col_duracion, nbins=30,
                                           title="Distribución de la Duración de Llamadas (Segundos)",
                                           labels={col_duracion: 'Segundos'})
                    st.plotly_chart(fig_dur, use_container_width=True)

            with tab2:
                st.subheader("Minutos Hablados por Encuestador")
                col_enc = "encuestador"
                col_min = "minutos_redondeados"
                if col_enc in df_base.columns and col_min in df_base.columns:
                    prod_enc = df_base.groupby(col_enc)[col_min].sum().sort_values(ascending=False).reset_index()
                    fig_prod = px.bar(prod_enc, x=col_enc, y=col_min,
                                      title="Total Minutos por Encuestador",
                                      labels={col_min: 'Minutos Totales', col_enc: 'Encuestador'},
                                      color=col_min)
                    st.plotly_chart(fig_prod, use_container_width=True)
                else:
                    st.info("Faltan columnas de encuestador o minutos para este análisis.")

        elif hoja_seleccionada == "Twilio":
            tab1, tab2, tab3 = st.tabs(["Estado y Calidad", "Análisis de Costos", "Distribución Horaria"])
            
            with tab1:
                st.subheader("Estado Técnico de Llamadas")
                if "Status" in df_base.columns:
                    fig_status = px.pie(df_base, names="Status", title="Efectividad de Conexión (Status)")
                    st.plotly_chart(fig_status, use_container_width=True)
                
                if "Direction" in df_base.columns:
                    fig_dir = px.bar(df_base["Direction"].value_counts(), title="Dirección de las Llamadas")
                    st.plotly_chart(fig_dir, use_container_width=True)

            with tab2:
                st.subheader("Análisis de Inversión")
                if "Price" in df_base.columns and "Type" in df_base.columns:
                    df_cost = df_base.copy()
                    df_cost["Costo Absoluto"] = df_cost["Price"].abs()
                    fig_cost = px.box(df_cost, x="Type", y="Costo Absoluto", color="Type", title="Distribución de Costos por Tipo")
                    st.plotly_chart(fig_cost, use_container_width=True)

            with tab3:
                st.subheader("Actividad por Hora (Picos de Llamadas)")
                if "Start Time" in df_base.columns:
                    df_base["Hora"] = df_base["Start Time"].dt.hour
                    hourly_calls = df_base.groupby("Hora").size().reset_index(name="Cantidad")
                    fig_hour = px.line(hourly_calls, x="Hora", y="Cantidad", title="Volumen de Llamadas por Hora del Día", markers=True)
                    st.plotly_chart(fig_hour, use_container_width=True)

        # Diccionario original (como referencia técnica)
        with st.expander("Ver Diccionario de Datos Técnico"):
            st.json(estructura)
else:
    st.error(f"No se encontró el archivo '{NOMBRE_ARCHIVO}' en el repositorio.")
