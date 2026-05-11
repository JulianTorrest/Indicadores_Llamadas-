
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

def limpiar_telefono(tel):
    if pd.isna(tel): return None
    s = str(tel).split('.')[0]
    s = ''.join(filter(str.isdigit, s))
    if s.startswith('57') and len(s) > 10:
        s = s[2:]
    if len(s) > 10:
        s = s[-10:]
    if len(s) == 10 and s.startswith('3'):
        return s
    return None

def limpiar_marca_temporal(valor):
    if pd.isna(valor): return pd.NaT
    s = str(valor).strip()
    fecha_texto = pd.to_datetime(s, errors='coerce', format='%m/%d/%Y %H:%M')
    if pd.notna(fecha_texto):
        return fecha_texto
    fecha_iso = pd.to_datetime(s, errors='coerce', format='%Y-%m-%d')
    if pd.notna(fecha_iso) and fecha_iso.month != 4 and fecha_iso.day == 4:
        return pd.Timestamp(year=fecha_iso.year, month=fecha_iso.day, day=fecha_iso.month)
    return fecha_iso

def agrupar_resultado_gestion(valor):
    if pd.isna(valor): return "Otros"
    # Normalizar: minúsculas, sin espacios laterales y sin tildes
    v = "".join(c for c in unicodedata.normalize('NFD', str(valor).lower().strip()) if unicodedata.category(c) != 'Mn')

    # 1. Éxito (Basado en el nuevo requerimiento: solo contesta y responde)
    if v in ["contesta y responde la encuesta", "contesta y responde la encuesta por forms"]:
        return "Éxito Total"
    
    if "encuesta incompleta" in v:
        return "Parcial / Incompleta"

    # 3. No Contactado -> Mapeo forzado a "No Contestaron"
    if any(x in v for x in ["no contesta", "sin respuesta", "no entra", "invalido", "no contest", "no contactado", "no contestaron"]):
        return "No Contestaron"

    # 4. Rechazo
    if "rechaza" in v or "rechazo" in v:
        return "Rechazo"

    # 5. Seguimiento
    if any(x in v for x in ["pide que", "llamen despues", "online", "chat", "mensaje"]):
        return "Seguimiento"

    # 6. Perfil
    if "universo" in v:
        return "Fuera de Perfil"

    return "Otros"

@st.cache_data(show_spinner="Cargando y limpiando datos...")
def cargar_y_limpiar_datos(ruta_archivo):
    try:
        xls = pd.ExcelFile(ruta_archivo)
        datos_limpios = {}
        
        for nombre_hoja in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=nombre_hoja)
            # Limpiar nombres de columnas para evitar espacios ocultos
            df.columns = df.columns.astype(str).str.strip()
            
            # Limpieza específica para hojas de gestión y contacto
            if nombre_hoja in ["Base_gestiones realizadas", "Contactados"]:
                # 1. Estandarizar nombres de encuestadores
                col_nombre = "Nombre del o de la encuestadora"
                if col_nombre in df.columns:
                    df[col_nombre] = df[col_nombre].apply(limpiar_texto)
                
                # Estandarizar Ciudad y Constructora para cruces precisos
                for col_to_clean in ["Ciudad", "Constructora"]:
                    if col_to_clean in df.columns:
                        df[col_to_clean] = df[col_to_clean].apply(limpiar_texto)
                        if col_to_clean == "Ciudad":
                            df[col_to_clean] = df[col_to_clean].replace({
                                "Bogota D.C.": "Bogota", "Bogota D.C": "Bogota"
                            })
                
                # Estandarizar fecha, día y teléfono
                if "Marca temporal" in df.columns:
                    df["Marca temporal"] = df["Marca temporal"].apply(limpiar_marca_temporal)
                    df["Dia_Semana"] = df["Marca temporal"].dt.day_name().map({
                        'Monday': '1. Lunes', 'Tuesday': '2. Martes', 'Wednesday': '3. Miércoles',
                        'Thursday': '4. Jueves', 'Friday': '5. Viernes', 'Saturday': '6. Sábado', 'Sunday': '7. Domingo'
                    })
                    df["Hora"] = df["Marca temporal"].dt.hour
                    # Crear franjas horarias de 3 horas
                    bins = [0, 3, 6, 9, 12, 15, 18, 21, 24]
                    labels = ["00:00 a 03:00", "03:00 a 06:00", "06:00 a 09:00", "09:00 a 12:00", 
                              "12:00 a 15:00", "15:00 a 18:00", "18:00 a 21:00", "21:00 a 00:00"]
                    df["Franja Horaria"] = pd.cut(df["Hora"], bins=bins, labels=labels, right=False)
                
                col_tel = "Número de teléfono sobre el que se realizó la gestión"
                if col_tel in df.columns:
                    df["tel_link"] = df[col_tel].apply(limpiar_telefono)
                
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
                
                col_tel = "Teléfono"
                if col_tel in df.columns:
                    df["tel_link"] = df[col_tel].apply(limpiar_telefono)

                # Estandarizar Constructora
                col_constructora = "Constructora"
                if col_constructora in df.columns:
                    df[col_constructora] = df[col_constructora].apply(limpiar_texto)
                
                # Estandarizar Ciudad
                col_ciudad = "Ciudad"
                if col_ciudad in df.columns:
                    df[col_ciudad] = df[col_ciudad].apply(limpiar_texto)
                    # Unificación robusta de Bogota
                    df[col_ciudad] = df[col_ciudad].replace({
                        "Bogota D.C.": "Bogota",
                        "Bogota D.C": "Bogota"
                    })
                
                # Estandarizar Ciudad2
                col_ciudad2 = "Ciudad2"
                if col_ciudad2 in df.columns:
                    df[col_ciudad2] = df[col_ciudad2].apply(limpiar_texto)
            
            # Limpieza específica para la hoja 'Camara_llamadas_salientes'
            elif nombre_hoja == "Camara_llamadas_salientes":
                # Estandarizar nombres de encuestadores
                if "encuestador" in df.columns:
                    df["encuestador"] = df["encuestador"].apply(limpiar_texto)
                
                if "numero_marcado" in df.columns:
                    df["tel_link"] = df["numero_marcado"].apply(limpiar_telefono)
                if "telefono_origen" in df.columns:
                    df["tel_origen"] = df["telefono_origen"].apply(limpiar_telefono)

                # Estandarizar fecha de llamada
                if "fecha_llamada" in df.columns:
                    df["fecha_llamada"] = pd.to_datetime(df["fecha_llamada"], errors='coerce', dayfirst=True)

            # Limpieza específica para la hoja 'Twilio'
            elif nombre_hoja == "Twilio":
                # Estandarizar fechas
                for col_fecha in ["Start Time", "End Time", "Date Created"]:
                    if col_fecha in df.columns:
                        # Manejo robusto de formato: 12:51:15 PDT 2026-04-13
                        # Reordenamos de "Hora TZ Fecha" a "Fecha Hora" ignorando la zona horaria
                        df[col_fecha] = df[col_fecha].astype(str).str.strip().str.replace(
                            r'(\d{1,2}:\d{1,2}:\d{1,2}).*?(\d{4}[-/]\d{1,2}[-/]\d{1,2})', 
                            r'\2 \1', 
                            regex=True
                        )
                        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce', dayfirst=True)
                
                if "To" in df.columns:
                    df["tel_link"] = df["To"].apply(limpiar_telefono)
                if "From" in df.columns:
                    df["tel_from"] = df["From"].apply(limpiar_telefono)

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

st.title("Dashboard de Campañas y Entregas")
st.info("Analizando datos directamente desde el repositorio de GitHub.")

NOMBRE_ARCHIVO = "Seguimiento encuestas consolidado.xlsx"

if os.path.exists(NOMBRE_ARCHIVO):
    # Proceso de carga y limpieza
    datos = cargar_y_limpiar_datos(NOMBRE_ARCHIVO)
    estructura = analizar_estructura_completa(NOMBRE_ARCHIVO)

    if datos:
        vistas = ["Funnel de Conversión", "Resumen de KPIs Críticos", "Visión Estratégica (Optimización)", "Análisis de Persistencia y Éxito", "Comportamiento 24h y Efectividad", "Análisis Cruzado (Auditoría)", "Comparativa vs. Asignación", "Base_gestiones realizadas", "Contactados", "Entregados (Base de Origen)", "Correo Masivo", "Camara_llamadas_salientes", "Twilio"]
        
        # Navegación Centralizada (Centro de la página)
        col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
        with col_nav2:
            hoja_seleccionada = st.selectbox("Navegación Principal - Seleccione la vista o fuente de datos:", vistas)
        st.divider()

        if hoja_seleccionada == "Funnel de Conversión":
            st.header("Embudo de Conversión de la Campaña")
            st.write("Representación del ciclo de vida: desde el intento inicial hasta la entrega efectiva.")

            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_c = datos.get("Contactados", pd.DataFrame())

            if not df_g.empty and not df_c.empty:
                # --- Filtros del Funnel ---
                st.subheader("Filtros del Embudo")
                col_f1, col_f2, col_f3 = st.columns(3)
                
                with col_f1:
                    enc_list = ["Todos"] + sorted(df_g["Nombre del o de la encuestadora"].dropna().unique().tolist())
                    sel_enc = st.selectbox("Encuestador", enc_list)
                with col_f2:
                    con_list = ["Todas"] + sorted(df_g["Constructora"].dropna().unique().tolist())
                    sel_con = st.selectbox("Constructora", con_list)
                with col_f3:
                    fran_list = ["Todas"] + ["00:00 a 03:00", "03:00 a 06:00", "06:00 a 09:00", "09:00 a 12:00", "12:00 a 15:00", "15:00 a 18:00", "18:00 a 21:00", "21:00 a 00:00"]
                    sel_fran = st.selectbox("Franja Horaria", fran_list)

                # Aplicar filtros a ambas dataframes
                df_g_f = df_g.copy()
                df_c_f = df_c.copy()

                if sel_enc != "Todos":
                    df_g_f = df_g_f[df_g_f["Nombre del o de la encuestadora"] == sel_enc]
                    df_c_f = df_c_f[df_c_f["Nombre del o de la encuestadora"] == sel_enc]
                if sel_con != "Todas":
                    df_g_f = df_g_f[df_g_f["Constructora"] == sel_con]
                    df_c_f = df_c_f[df_c_f["Constructora"] == sel_con]
                if sel_fran != "Todas":
                    df_g_f = df_g_f[df_g_f["Franja Horaria"] == sel_fran]
                    df_c_f = df_c_f[df_c_f["Franja Horaria"] == sel_fran]

                # Calcular Encuestas Exitosas desde df_g
                exitos_g = len(df_g_f[df_g_f["Resultado de la gestión (Agrupado)"] == "Éxito Total"])

                # Preparar datos para el Funnel
                etapas = ["Total Gestiones", "Contactos Efectivos", "Encuestas Exitosas"]
                valores = [len(df_g_f), len(df_c_f), exitos_g]
                
                fig_funnel = px.funnel(
                    data_frame=pd.DataFrame({"Etapa": etapas, "Cantidad": valores}),
                    x='Cantidad',
                    y='Etapa',
                    title="Ciclo de Vida del Proceso",
                    color_discrete_sequence=px.colors.qualitative.Prism
                )
                fig_funnel.update_traces(textinfo="value+percent initial")
                st.plotly_chart(fig_funnel, use_container_width=True)

                # Métricas de tasa de caída
                c1, c2 = st.columns(2)
                with c1:
                    tasa_contacto = (len(df_c_f) / len(df_g_f)) * 100 if len(df_g_f) > 0 else 0
                    st.metric("Tasa de Contactabilidad (Base -> Contactos)", f"{tasa_contacto:.1f}%")
                with c2:
                    tasa_conversion = (exitos_g / len(df_c_f)) * 100 if len(df_c_f) > 0 else 0
                    st.metric("Tasa de Efectividad (Contactos -> Éxito)", f"{tasa_conversion:.1f}%")
                
                st.divider()
                st.subheader("Análisis por Segmento")
                # Podemos ver el funnel por Ciudad si existe en todas las hojas
                if "Ciudad" in df_g_f.columns and "Ciudad" in df_c_f.columns: # Asegurar que Ciudad esté en ambas para filtrar
                    ciudad_sel = st.selectbox("Ver Funnel por Ciudad", ["Todas"] + sorted(df_g_f["Ciudad"].dropna().unique().tolist()))
                    if ciudad_sel != "Todas":
                        v_g = len(df_g_f[df_g_f["Ciudad"] == ciudad_sel])
                        v_c = len(df_c_f[df_c_f["Ciudad"] == ciudad_sel])
                        # Recalcular exitos para el segmento
                        v_e_segment = len(df_g_f[(df_g_f["Ciudad"] == ciudad_sel) & (df_g_f["Resultado de la gestión (Agrupado)"] == "Éxito Total")])
                        
                        fig_funnel_sub = px.funnel( # Usar v_e_segment
                            data_frame=pd.DataFrame({"Etapa": etapas, "Cantidad": [v_g, v_c, v_e_segment]}),
                            x='Cantidad', y='Etapa',
                            title=f"Embudo en {ciudad_sel}"
                        )
                        fig_funnel_sub.update_traces(textinfo="value+percent initial")
                        st.plotly_chart(fig_funnel_sub, use_container_width=True)
                
                with st.expander("Verificación Técnica de Categorías"):
                    st.write("Valores detectados en 'Resultado de la gestión (Agrupado)':")
                    st.write(df_g_f["Resultado de la gestión (Agrupado)"].value_counts())
            else:
                st.warning("Se requieren datos en las hojas 'Base_gestiones realizadas' y 'Contactados' para generar el funnel.")
            
            st.stop()

        elif hoja_seleccionada == "Resumen de KPIs Críticos":
            st.header("Indicadores Críticos de Campaña")
            
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            df_c = datos.get("Camara_llamadas_salientes", pd.DataFrame())
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            
            if not df_g.empty:
                # 1. Tasa de Contactabilidad
                total = len(df_g)
                no_contactados = len(df_g[df_g["Resultado de la gestión (Agrupado)"] == "No Contestaron"])
                contactabilidad = ((total - no_contactados) / total) * 100 if total > 0 else 0
                kpi1.metric("Contactabilidad", f"{contactabilidad:.1f}%", help="Porcentaje de gestiones que resultaron en un contacto real")

                # 2. Tasa de Conversión
                exitos = len(df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito Total"])
                conversion = (exitos / (total - no_contactados)) * 100 if (total - no_contactados) > 0 else 0
                kpi2.metric("Conversión", f"{conversion:.1f}%", help="Encuestas exitosas sobre el total de personas contactadas")

            # 3. AHT (Tiempo Promedio Consolidado de Twilio y Cámara)
            list_dur_kpi = []
            if not df_t.empty and "Duration" in df_t.columns:
                list_dur_kpi.append(df_t[df_t["Duration"] > 0]["Duration"])
            if not df_c.empty and "segundos" in df_c.columns:
                list_dur_kpi.append(df_c[df_c["segundos"] > 0]["segundos"])
            
            if list_dur_kpi:
                all_durations = pd.concat(list_dur_kpi)
                aht = all_durations.mean()
                kpi3.metric("AHT Consolidado", f"{aht:.1f}s", help="Promedio de duración de llamadas (Manuales + Twilio)")
            else:
                kpi3.metric("AHT Consolidado", "N/A")

            if not df_t.empty:
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

            st.stop() # Finaliza la vista de KPIs para no mostrar el resto

        elif hoja_seleccionada == "Visión Estratégica (Optimización)":
            st.header("Visión Estratégica: Optimización de Recursos")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            
            if not df_g.empty:
                tab_strat1, tab_strat2 = st.tabs(["Saturación de Leads (Burnout)", "Cuadrante de Eficiencia de Agentes"])
                
                with tab_strat1:
                    st.subheader("Análisis de Saturación de Contactos")
                    st.write("Identifica números con alta cantidad de intentos sin éxito para evitar el desgaste de la base.")
                    
                    intentos = df_g.groupby('tel_link').size().reset_index(name='Num_Intentos')
                    tels_exito = set(df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito Total"]["tel_link"].unique())
                    intentos['Estado'] = intentos['tel_link'].apply(lambda x: 'Exitoso' if x in tels_exito else 'Pendiente/Fallido')
                    
                    saturados = intentos[(intentos['Estado'] == 'Pendiente/Fallido') & (intentos['Num_Intentos'] > 5)]
                    
                    col_s1, col_s2 = st.columns(2)
                    col_s1.metric("Leads Saturados (>5 intentos)", len(saturados))
                    col_s2.metric("% de la Base en Saturación", f"{(len(saturados)/len(intentos)*100):.1f}%" if len(intentos)>0 else "0%")
                    
                    fig_sat = px.histogram(intentos[intentos['Estado'] == 'Pendiente/Fallido'], x='Num_Intentos', 
                                           title="Distribución de Intentos en Leads No Exitosos",
                                           labels={'Num_Intentos': 'Número de Intentos'}, text_auto=True)
                    st.plotly_chart(fig_sat, use_container_width=True)
                    
                with tab_strat2:
                    st.subheader("Cuadrante de Desempeño de Encuestadores")
                    st.write("Relación entre el volumen de trabajo y la efectividad real.")
                    
                    col_agent = "Nombre del o de la encuestadora"
                    if col_agent in df_g.columns:
                        agentes = df_g.groupby(col_agent).agg(
                            Total_Gestiones=('tel_link', 'count'),
                            Exitos=('Resultado de la gestión (Agrupado)', lambda x: (x == "Éxito Total").sum())
                        ).reset_index()
                        agentes['Tasa_Conversion'] = (agentes['Exitos'] / agentes['Total_Gestiones']) * 100
                        
                        fig_quad = px.scatter(agentes, x='Total_Gestiones', y='Tasa_Conversion', 
                                              text=col_agent, size='Exitos', color='Tasa_Conversion',
                                              title="Efectividad vs Volumen por Encuestador",
                                              labels={'Total_Gestiones': 'Volumen de Gestiones', 'Tasa_Conversion': '% Conversión (Éxito)'},
                                              color_continuous_scale='RdYlGn')
                        
                        # Añadir líneas de promedio
                        fig_quad.add_hline(y=agentes['Tasa_Conversion'].mean(), line_dash="dash", annotation_text="Media Eficiencia")
                        fig_quad.add_vline(x=agentes['Total_Gestiones'].mean(), line_dash="dash", annotation_text="Media Volumen")
                        
                        st.plotly_chart(fig_quad, use_container_width=True)
                        st.info("- **Superior Derecha:** Alto volumen y alta eficiencia (Top Performers).\n"
                                "- **Superior Izquierda:** Bajo volumen pero alta eficiencia (Calidad sobre cantidad).\n"
                                "- **Inferior Derecha:** Alto volumen pero baja eficiencia (Revisar discurso).\n"
                                "- **Inferior Izquierda:** Bajo volumen y baja eficiencia.")
            else:
                st.warning("No hay datos suficientes para el análisis estratégico.")
            st.stop()

        elif hoja_seleccionada == "Análisis de Persistencia y Éxito":
            st.header("Análisis de Persistencia y Factores de Éxito")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            df_c = datos.get("Camara_llamadas_salientes", pd.DataFrame())

            if not df_g.empty:
                # 1. Análisis de Intentos
                intentos_por_tel = df_g.groupby('tel_link').size().reset_index(name='Intentos')
                
                # Identificar si el número terminó en éxito (basado en df_g)
                if "Resultado de la gestión (Agrupado)" in df_g.columns:
                    tels_exito = set(df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito Total"]["tel_link"].dropna().unique())
                    intentos_por_tel['Resultado Final'] = intentos_por_tel['tel_link'].apply(
                        lambda x: 'Exitoso (Entrega)' if x in tels_exito else 'No Efectivo'
                    )
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.subheader("¿Cuántos intentos toma tener éxito?")
                        avg_intentos = intentos_por_tel.groupby('Resultado Final')['Intentos'].mean().reset_index()
                        fig_int = px.bar(avg_intentos, x='Resultado Final', y='Intentos', 
                                         color='Resultado Final', text_auto='.1f',
                                         title="Promedio de Intentos Realizados")
                        st.plotly_chart(fig_int, use_container_width=True)
                    
                    with col_b:
                        st.subheader("Distribución de Intentos")
                        fig_dist = px.histogram(intentos_por_tel, x='Intentos', color='Resultado Final',
                                                marginal="box", barmode="overlay",
                                                title="Frecuencia de intentos por contacto")
                        st.plotly_chart(fig_dist, use_container_width=True)

                # 2. Análisis de Duración (Cruce Consolidado Twilio + Cámara)
                st.divider()
                st.subheader("Duración de Llamadas (Técnico) vs. Resultado de Gestión")
                
                # Consolidar registros de duración de ambas fuentes
                list_tech_logs = []
                if not df_t.empty:
                    list_tech_logs.append(df_t[['tel_link', 'Duration']])
                if not df_c.empty:
                    list_tech_logs.append(df_c[['tel_link', 'segundos']].rename(columns={'segundos': 'Duration'}))
                
                if list_tech_logs:
                    df_all_tech = pd.concat(list_tech_logs)
                    # Agrupamos por teléfono para obtener duración promedio técnica por contacto
                    tech_dur_mean = df_all_tech.groupby('tel_link')['Duration'].mean().reset_index()
                    
                    df_dur = pd.merge(df_g, tech_dur_mean, on='tel_link', how='inner')
                    
                    # Filtrar duraciones > 0 para evitar ruido
                    df_dur = df_dur[df_dur['Duration'] > 0]
                    
                    res_dur = df_dur.groupby('Resultado de la gestión (Agrupado)')['Duration'].mean().reset_index()
                    fig_res_dur = px.bar(res_dur, x='Resultado de la gestión (Agrupado)', y='Duration',
                                         text_auto='.1s', color='Resultado de la gestión (Agrupado)',
                                         title="Duración Promedio Real (Segundos) por Categoría de Gestión",
                                         labels={'Duration': 'Segundos (Promedio)'})
                    st.plotly_chart(fig_res_dur, use_container_width=True)
                    st.info("Nota: Este gráfico cruza el 'Resultado de la gestión' manual con la duración técnica consolidada de Twilio y Cámara.")

                # 3. Efectividad por Día y Hora
                st.divider()
                st.subheader("¿Cuándo es mejor llamar?")
                if 'Dia_Semana' in df_g.columns and 'Franja Horaria' in df_g.columns and "Resultado de la gestión (Agrupado)" in df_g.columns:
                    st.write("Conversión a Encuesta Exitosa con filtros por encuestador y constructora.")
                    col_temp_f1, col_temp_f2 = st.columns(2)
                    df_g_temp = df_g.copy()

                    with col_temp_f1:
                        if "Nombre del o de la encuestadora" in df_g_temp.columns:
                            enc_temp_list = ["Todos"] + sorted(df_g_temp["Nombre del o de la encuestadora"].dropna().unique().tolist())
                            sel_enc_temp = st.selectbox("Encuestador - Conversión a Encuesta Exitosa", enc_temp_list)
                            if sel_enc_temp != "Todos":
                                df_g_temp = df_g_temp[df_g_temp["Nombre del o de la encuestadora"] == sel_enc_temp]
                        else:
                            st.warning("La columna 'Nombre del o de la encuestadora' no está disponible para filtrar.")

                    with col_temp_f2:
                        if "Constructora" in df_g_temp.columns:
                            con_temp_list = ["Todas"] + sorted(df_g_temp["Constructora"].dropna().unique().tolist())
                            sel_con_temp = st.selectbox("Constructora - Conversión a Encuesta Exitosa", con_temp_list)
                            if sel_con_temp != "Todas":
                                df_g_temp = df_g_temp[df_g_temp["Constructora"] == sel_con_temp]
                        else:
                            st.warning("La columna 'Constructora' no está disponible para filtrar.")

                    df_g_temp['Es_Exito'] = df_g_temp["Resultado de la gestión (Agrupado)"] == "Éxito Total"

                    dia_order = ["1. Lunes", "2. Martes", "3. Miércoles", "4. Jueves", "5. Viernes", "6. Sábado", "7. Domingo"]
                    franja_horaria_order = ["00:00 a 03:00", "03:00 a 06:00", "06:00 a 09:00", "09:00 a 12:00", 
                    "12:00 a 15:00", "15:00 a 18:00", "18:00 a 21:00", "21:00 a 00:00"]

                    if not df_g_temp.empty:
                        efectividad_dia = df_g_temp.groupby('Dia_Semana')['Es_Exito'].mean().reset_index()
                        efectividad_dia['Es_Exito'] *= 100
                        efectividad_hora = df_g_temp.groupby('Franja Horaria')['Es_Exito'].mean().reset_index()
                        efectividad_hora['Es_Exito'] *= 100

                        col_dia, col_hora = st.columns(2)
                        with col_dia:
                            fig_dia = px.bar(efectividad_dia, x='Dia_Semana', y='Es_Exito',
                                             text_auto='.1f', color='Es_Exito',
                                             color_continuous_scale='Viridis',
                                             title="Mapa de Calor: % de Efectividad por Día",
                                             labels={'Es_Exito': '% Efectividad', 'Dia_Semana': 'Día'},
                                             category_orders={"Dia_Semana": dia_order})
                            fig_dia.update_traces(texttemplate='%{y:.1f}%')
                            fig_dia.update_layout(yaxis_ticksuffix="%")
                            st.plotly_chart(fig_dia, use_container_width=True)
                            st.caption("% de Efectividad por Día = (gestiones con Éxito Total en ese día ÷ total de gestiones realizadas en ese día) × 100. Mide qué proporción de las gestiones de cada día terminó en Encuesta Exitosa.")

                        with col_hora:
                            fig_hora = px.bar(efectividad_hora, x='Franja Horaria', y='Es_Exito',
                                              text_auto='.1f', color='Es_Exito',
                                              color_continuous_scale='Viridis',
                                              title="Mapa de Calor: % de Efectividad por Hora",
                                              labels={'Es_Exito': '% Efectividad', 'Franja Horaria': 'Franja Horaria del Día'},
                                              category_orders={"Franja Horaria": franja_horaria_order})
                            fig_hora.update_traces(texttemplate='%{y:.1f}%')
                            fig_hora.update_layout(yaxis_ticksuffix="%")
                            st.plotly_chart(fig_hora, use_container_width=True)
                            st.caption("% de Efectividad por Hora = (gestiones con Éxito Total en esa franja horaria ÷ total de gestiones realizadas en esa franja horaria) × 100. Mide qué proporción de las gestiones de cada horario terminó en Encuesta Exitosa.")

                        st.caption("Los valores más altos indican días u horas con mayor probabilidad de que la gestión termine en una Encuesta Exitosa. La efectividad se calcula como el promedio de éxitos (0% a 100%).")
                    else:
                        st.warning("No hay datos disponibles con los filtros seleccionados para calcular la efectividad.")
                else:
                    st.warning("Las columnas 'Dia_Semana', 'Franja Horaria' o 'Resultado de la gestión (Agrupado)' no están disponibles para el análisis del mapa de calor.")

                st.divider()
                st.subheader("Desempeño Diario por Encuestador desde su Fecha de Inicio")
                col_encuestador_perf = "Nombre del o de la encuestadora"
                cols_perf = [col_encuestador_perf, "Marca temporal", "Franja Horaria", "Resultado de la gestión (Agrupado)"]

                if all(col in df_g.columns for col in cols_perf):
                    df_perf = df_g.dropna(subset=[col_encuestador_perf, "Marca temporal"]).copy()
                    df_perf["Fecha"] = df_perf["Marca temporal"].dt.date
                    df_perf["Es_Exito"] = df_perf["Resultado de la gestión (Agrupado)"] == "Éxito Total"

                    perf_f1, perf_f2 = st.columns(2)
                    with perf_f1:
                        enc_perf_list = ["Todos"] + sorted(df_perf[col_encuestador_perf].dropna().unique().tolist())
                        sel_enc_perf = st.selectbox("Encuestador - Desempeño Diario", enc_perf_list)
                        if sel_enc_perf != "Todos":
                            df_perf = df_perf[df_perf[col_encuestador_perf] == sel_enc_perf]

                    with perf_f2:
                        franja_perf_order = ["Todas", "00:00 a 03:00", "03:00 a 06:00", "06:00 a 09:00", "09:00 a 12:00", 
                                             "12:00 a 15:00", "15:00 a 18:00", "18:00 a 21:00", "21:00 a 00:00"]
                        franjas_disponibles = [f for f in franja_perf_order if f == "Todas" or f in df_perf["Franja Horaria"].dropna().astype(str).unique().tolist()]
                        sel_franja_perf = st.selectbox("Franja Horaria - Desempeño Diario", franjas_disponibles)
                        if sel_franja_perf != "Todas":
                            df_perf = df_perf[df_perf["Franja Horaria"].astype(str) == sel_franja_perf]

                    if not df_perf.empty:
                        fecha_inicio = df_perf.groupby(col_encuestador_perf)["Marca temporal"].min().dt.date.reset_index(name="Fecha_Inicio")
                        desempeño_diario = df_perf.groupby([col_encuestador_perf, "Fecha"]).agg(
                            Llamadas_Realizadas=("tel_link", "count"),
                            Llamadas_Efectivas=("Es_Exito", "sum")
                        ).reset_index()
                        desempeño_diario = pd.merge(desempeño_diario, fecha_inicio, on=col_encuestador_perf, how="left")
                        desempeño_diario["Día desde inicio"] = (
                            pd.to_datetime(desempeño_diario["Fecha"]) - pd.to_datetime(desempeño_diario["Fecha_Inicio"])
                        ).dt.days + 1
                        desempeño_diario["% Efectividad"] = desempeño_diario["Llamadas_Efectivas"] / desempeño_diario["Llamadas_Realizadas"] * 100

                        fig_vol_perf = px.line(
                            desempeño_diario,
                            x="Día desde inicio",
                            y=["Llamadas_Realizadas", "Llamadas_Efectivas"],
                            color=col_encuestador_perf,
                            markers=True,
                            title="Llamadas Realizadas vs Llamadas Efectivas por Día desde Inicio",
                            labels={"value": "Cantidad", "variable": "Métrica", col_encuestador_perf: "Encuestador"}
                        )
                        st.plotly_chart(fig_vol_perf, use_container_width=True)

                        fig_efec_perf = px.line(
                            desempeño_diario,
                            x="Día desde inicio",
                            y="% Efectividad",
                            color=col_encuestador_perf,
                            markers=True,
                            title="% de Efectividad Diario por Encuestador desde Inicio",
                            labels={"% Efectividad": "% Efectividad", col_encuestador_perf: "Encuestador"}
                        )
                        fig_efec_perf.update_layout(yaxis_ticksuffix="%")
                        st.plotly_chart(fig_efec_perf, use_container_width=True)

                        st.dataframe(desempeño_diario.sort_values([col_encuestador_perf, "Día desde inicio"]))
                    else:
                        st.warning("No hay datos disponibles con los filtros seleccionados para analizar el desempeño diario.")
                else:
                    st.warning("No están disponibles las columnas necesarias para calcular el desempeño diario por encuestador.")

                
                # 4. Análisis específico de "Llamar después" (Seguimiento)
                st.subheader("Evolución de los casos 'Llamar Después'")
                
                # Identificar el primer 'Seguimiento' por teléfono
                df_seg = df_g[df_g['Resultado de la gestión (Agrupado)'] == 'Seguimiento'].sort_values('Marca temporal')
                primer_seguimiento = df_seg.groupby('tel_link')['Marca temporal'].min().reset_index()
                primer_seguimiento.columns = ['tel_link', 'Fecha_Primer_Seguimiento']
                
                # Unir con la base general para ver qué pasó después
                df_post = pd.merge(df_g, primer_seguimiento, on='tel_link')
                df_post = df_post[df_post['Marca temporal'] > df_post['Fecha_Primer_Seguimiento']]
                
                num_contactos_seg = primer_seguimiento['tel_link'].nunique()
                st.write(f"Número de contactos únicos que pidieron ser llamados después: **{num_contactos_seg}**")
                
                if not df_post.empty:
                    # Obtener el último estado de esos contactos
                    ultimo_estado = df_post.sort_values('Marca temporal').groupby('tel_link').last().reset_index()
                    conversion_seg = ultimo_estado['Resultado de la gestión (Agrupado)'].value_counts().reset_index()
                    conversion_seg.columns = ['Resultado Posterior', 'Cantidad']
                    
                    c_seg1, c_seg2 = st.columns([1, 2])
                    with c_seg1:
                        exitos_pos = conversion_seg[conversion_seg['Resultado Posterior'].isin(["Éxito Total", "Parcial / Incompleta"])]['Cantidad'].sum()
                        tasa_recuperacion = (exitos_pos / num_contactos_seg) * 100
                        st.metric("Tasa de Recuperación (Cualquier éxito)", f"{tasa_recuperacion:.1f}%", 
                                  help="Porcentaje de casos que pidieron 'Llamar después' y terminaron en Éxito")
                        st.write(f"De los {num_contactos_seg} casos, **{exitos_pos}** se convirtieron en éxito finalmente.")
                        st.caption(f"Este indicador mide la recuperación sobre todos los contactos únicos que pidieron seguimiento. Fórmula: ({exitos_pos} casos con Éxito Total o Parcial / Incompleta ÷ {num_contactos_seg} contactos que pidieron seguimiento) × 100 = {tasa_recuperacion:.1f}%.")

                    with c_seg2:
                        fig_post = px.pie(conversion_seg, names='Resultado Posterior', values='Cantidad',
                                          title="Resultado Final de los contactos que pidieron Seguimiento",
                                          color_discrete_sequence=px.colors.qualitative.Pastel)
                        fig_post.update_traces(textinfo='percent+label')
                        st.plotly_chart(fig_post, use_container_width=True)
                        total_post = len(ultimo_estado)
                        st.caption(f"Este gráfico muestra la distribución del último resultado posterior solo para los contactos que pidieron seguimiento y tuvieron al menos una gestión después. Fórmula de cada porcentaje: (cantidad de contactos en cada resultado ÷ {total_post} contactos con gestión posterior) × 100.")

                    # Nuevo gráfico: ¿En qué intento se logra el éxito?
                    st.write("### Esfuerzo de Conversión Post-Seguimiento")
                    
                    # --- CAMBIO CRÍTICO AQUÍ ---
                    # 1. Identificar los tel_link que realmente terminaron en éxito (según la métrica de arriba)
                    successful_tel_links_from_metric = ultimo_estado[ultimo_estado['Resultado de la gestión (Agrupado)'] == "Éxito Total"]['tel_link']
                    
                    # 2. Filtrar df_post para incluir solo las actividades de ESTOS contactos exitosos
                    df_post_successful_only = df_post[df_post['tel_link'].isin(successful_tel_links_from_metric)]
                    
                    # 3. Filtrar las gestiones de éxito dentro de este subconjunto
                    df_exitos_post_filtered = df_post_successful_only[df_post_successful_only['Resultado de la gestión (Agrupado)'] == "Éxito Total"]
                    
                    if not df_exitos_post_filtered.empty:
                        # Encontrar la fecha del primer éxito post-seguimiento
                        # Usamos df_exitos_post_filtered para asegurar que solo consideramos los éxitos finales
                        primer_exito_post_filtered = df_exitos_post_filtered.sort_values('Marca temporal').groupby('tel_link')['Marca temporal'].min().reset_index()
                        primer_exito_post_filtered.columns = ['tel_link', 'Fecha_Primer_Exito']
                        
                        # Contar intentos ocurridos entre el primer seguimiento y el primer éxito
                        # Merge con df_post_successful_only para contar intentos para solo estos contactos
                        df_camino_exito_filtered = pd.merge(df_post_successful_only, primer_exito_post_filtered, on='tel_link')
                        df_camino_exito_filtered = df_camino_exito_filtered[df_camino_exito_filtered['Marca temporal'] <= df_camino_exito_filtered['Fecha_Primer_Exito']]
                        intentos_hasta_exito_filtered = df_camino_exito_filtered.groupby('tel_link').size().reset_index(name='Intentos_Adicionales')
                        
                        fig_momento = px.bar(intentos_hasta_exito_filtered['Intentos_Adicionales'].value_counts().sort_index(),
                                            title="¿Cuántas llamadas adicionales toma convertir un 'Llamar Después' en Éxito?",
                                            labels={'index': 'Nro de Intentos Adicionales', 'value': 'Cantidad de Contactos'},
                                            text_auto=True)
                        st.plotly_chart(fig_momento, use_container_width=True)
                else:
                    st.warning("No se detectan intentos posteriores a la primera solicitud de seguimiento todavía.")
            
            st.stop()

        elif hoja_seleccionada == "Comportamiento 24h y Efectividad":
            st.header("Comportamiento Temporal y Efectividad")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())

            if not df_g.empty and "Marca temporal" in df_g.columns:
                # 1. Filtros de Análisis
                st.subheader("Filtros de Análisis Temporal")
                f1, f2, f3 = st.columns(3)
                
                with f1:
                    list_enc = ["Todos"] + sorted(df_g["Nombre del o de la encuestadora"].dropna().unique().tolist())
                    sel_enc = st.selectbox("Filtrar por Encuestador", list_enc)
                with f2:
                    list_ciu = ["Todas"] + sorted(df_g["Ciudad"].dropna().unique().tolist())
                    sel_ciu = st.selectbox("Filtrar por Ciudad", list_ciu)
                with f3:
                    list_con = ["Todas"] + sorted(df_g["Constructora"].dropna().unique().tolist())
                    sel_con = st.selectbox("Filtrar por Constructora", list_con)

                # 2. Aplicar Filtros a la base de gestiones
                df_g_f = df_g.copy()
                if sel_enc != "Todos": df_g_f = df_g_f[df_g_f["Nombre del o de la encuestadora"] == sel_enc]
                if sel_ciu != "Todas": df_g_f = df_g_f[df_g_f["Ciudad"] == sel_ciu]
                if sel_con != "Todas": df_g_f = df_g_f[df_g_f["Constructora"] == sel_con]

                df_g_f['Hora'] = df_g_f['Marca temporal'].dt.hour
                
                # Cruzar con entregados para medir efectividad real
                if "Resultado de la gestión (Agrupado)" in df_g_f.columns:
                    # Determinar 'Efectivo' basado en df_g_f's success criteria
                    df_g_f['Efectivo'] = df_g_f["Resultado de la gestión (Agrupado)"] == "Éxito Total"
                    
                    # Agrupar por hora
                    hourly_stats = df_g_f.groupby('Hora').agg(
                        Gestiones=('Marca temporal', 'count'),
                        Entregas=('Efectivo', 'sum')
                    ).reset_index()
                    
                    # Asegurar que las 24 horas estén presentes para que las líneas sean continuas
                    horas_full = pd.DataFrame({'Hora': range(24)})
                    hourly_stats = pd.merge(horas_full, hourly_stats, on='Hora', how='left').fillna(0)
                    
                    fig_24h = px.line(hourly_stats, x='Hora', y=['Gestiones', 'Entregas'], 
                                      title=f"Efectividad Horaria - {sel_enc} | {sel_ciu} | {sel_con}",
                                      markers=True, 
                                      labels={'value': 'Cantidad', 'variable': 'Métrica', 'Hora': 'Hora del Día'},
                                      color_discrete_map={'Gestiones': '#1f77b4', 'Entregas': '#ff7f0e'})
                    
                    fig_24h.update_layout(hovermode="x unified", xaxis=dict(tickmode='linear', tick0=0, dtick=1))
                    fig_24h.update_traces(textposition="top center")
                    st.plotly_chart(fig_24h, use_container_width=True)
                    
                    st.info("Consejo: Las líneas muestran la actividad de 0 a 23h. Si una línea está en 0, significa que no hubo actividad en esa franja para los filtros seleccionados.")
                else:
                    st.warning("Se requieren datos de 'Entregados' para medir efectividad.")
            st.stop()

        elif hoja_seleccionada == "Análisis Cruzado (Auditoría)":
            st.header("Auditoría: Gestión vs Telefonía")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            df_c = datos.get("Camara_llamadas_salientes", pd.DataFrame())

            if not df_g.empty:
                # Consolidar telefonía técnica
                tels_tecnicos = set()
                if not df_t.empty: tels_tecnicos.update(df_t['tel_link'].dropna().unique())
                if not df_c.empty: tels_tecnicos.update(df_c['tel_link'].dropna().unique())
                
                df_g['Validado_Tecnico'] = df_g['tel_link'].isin(tels_tecnicos)
                
                audit_stats = df_g['Validado_Tecnico'].value_counts().reset_index()
                audit_stats.columns = ['Estado', 'Cantidad']
                audit_stats['Estado'] = audit_stats['Estado'].map({True: 'Con Respaldo Técnico', False: 'Sin Respaldo (Solo Manual)'})
                
                fig_audit = px.pie(audit_stats, names='Estado', values='Cantidad', 
                                   title="Validación de Gestiones Manuales vs Logs de Telefonía")
                fig_audit.update_traces(textinfo='percent+label')
                st.plotly_chart(fig_audit, use_container_width=True)

                st.divider()
                st.subheader("Comparativo de Efectividad por Estrategia de Contacto")

                tels_twilio = set()
                tels_camara = set()
                if not df_t.empty and 'tel_link' in df_t.columns:
                    tels_twilio = set(df_t['tel_link'].dropna().unique())
                if not df_c.empty and 'tel_link' in df_c.columns:
                    tels_camara = set(df_c['tel_link'].dropna().unique())

                def clasificar_estrategia_declarada(medio):
                    if medio == "Plataforma web":
                        return "Twilio"
                    if medio == "Celular/Tablet":
                        return "Cámara"
                    return "Sin clasificar"

                df_estrategia = df_g.copy()
                medio_col = "¿Por qué medio se realizó la llamada?"
                df_estrategia['Estrategia'] = df_estrategia[medio_col].apply(clasificar_estrategia_declarada)
                df_estrategia['Respaldo_Twilio'] = df_estrategia['tel_link'].isin(tels_twilio)
                df_estrategia['Respaldo_Camara'] = df_estrategia['tel_link'].isin(tels_camara)
                df_estrategia['Respaldo_Esperado'] = (
                    ((df_estrategia['Estrategia'] == "Twilio") & df_estrategia['Respaldo_Twilio']) |
                    ((df_estrategia['Estrategia'] == "Cámara") & df_estrategia['Respaldo_Camara'])
                )
                df_estrategia['Es_Exito'] = df_estrategia["Resultado de la gestión (Agrupado)"] == "Éxito Total"
                df_estrategia_valida = df_estrategia[df_estrategia['Estrategia'].isin(["Twilio", "Cámara"])]

                if not df_estrategia_valida.empty:
                    resumen_estrategia = df_estrategia_valida.groupby('Estrategia').agg(
                        Total_Gestiones=('tel_link', 'count'),
                        Encuestas_Exitosas=('Es_Exito', 'sum')
                    ).reset_index()
                    resumen_estrategia['No_Exitosas'] = resumen_estrategia['Total_Gestiones'] - resumen_estrategia['Encuestas_Exitosas']
                    resumen_estrategia['Tasa_Efectividad'] = (resumen_estrategia['Encuestas_Exitosas'] / resumen_estrategia['Total_Gestiones']) * 100

                    col_est1, col_est2 = st.columns(2)
                    with col_est1:
                        fig_efectividad_estrategia = px.bar(
                            resumen_estrategia,
                            x='Estrategia',
                            y='Tasa_Efectividad',
                            color='Estrategia',
                            text='Tasa_Efectividad',
                            title="Barras Comparativas de Efectividad",
                            labels={'Tasa_Efectividad': '% de Efectividad', 'Estrategia': 'Estrategia de Contacto'}
                        )
                        fig_efectividad_estrategia.update_traces(texttemplate='%{text:.1f}%')
                        fig_efectividad_estrategia.update_layout(yaxis_ticksuffix="%")
                        st.plotly_chart(fig_efectividad_estrategia, use_container_width=True)
                        st.caption("% de Efectividad = (Encuestas Exitosas de la estrategia declarada ÷ Total de gestiones de la estrategia declarada) × 100. Plataforma web se clasifica como Twilio y Celular/Tablet como Cámara.")

                    with col_est2:
                        volumen_estrategia = resumen_estrategia.melt(
                            id_vars='Estrategia',
                            value_vars=['Total_Gestiones', 'Encuestas_Exitosas'],
                            var_name='Métrica',
                            value_name='Cantidad'
                        )
                        volumen_estrategia['Métrica'] = volumen_estrategia['Métrica'].replace({
                            'Total_Gestiones': 'Total Gestiones',
                            'Encuestas_Exitosas': 'Encuestas Exitosas'
                        })
                        fig_volumen_estrategia = px.bar(
                            volumen_estrategia,
                            x='Estrategia',
                            y='Cantidad',
                            color='Métrica',
                            barmode='group',
                            text_auto=True,
                            title="Barras Comparativas de Volumen",
                            labels={'Cantidad': 'Cantidad', 'Estrategia': 'Estrategia de Contacto'}
                        )
                        st.plotly_chart(fig_volumen_estrategia, use_container_width=True)
                        st.caption("El volumen compara el total de gestiones registradas contra la cantidad de gestiones que terminaron en Éxito Total para Twilio y Cámara según el medio declarado.")

                    st.dataframe(resumen_estrategia, use_container_width=True)
                else:
                    st.warning("No se encontraron gestiones clasificadas como Twilio o Cámara para comparar estrategias.")

                st.subheader("Control de Calidad del Cruce Técnico Esperado")
                resumen_respaldo = df_estrategia_valida.groupby('Estrategia').agg(
                    Total_Gestiones=('tel_link', 'count'),
                    Gestiones_Con_Respaldo_Esperado=('Respaldo_Esperado', 'sum')
                ).reset_index()
                if not resumen_respaldo.empty:
                    resumen_respaldo['Gestiones_Sin_Respaldo_Esperado'] = resumen_respaldo['Total_Gestiones'] - resumen_respaldo['Gestiones_Con_Respaldo_Esperado']
                    resumen_respaldo['% Sin Respaldo Esperado'] = (resumen_respaldo['Gestiones_Sin_Respaldo_Esperado'] / resumen_respaldo['Total_Gestiones']) * 100

                    cols_respaldo = st.columns(len(resumen_respaldo))
                    for idx, row in resumen_respaldo.iterrows():
                        with cols_respaldo[idx]:
                            st.metric(
                                f"{row['Estrategia']} sin respaldo esperado",
                                f"{row['% Sin Respaldo Esperado']:.1f}%",
                                f"{int(row['Gestiones_Sin_Respaldo_Esperado'])} de {int(row['Total_Gestiones'])}"
                            )

                    fig_respaldo = px.bar(
                        resumen_respaldo,
                        x='Estrategia',
                        y='% Sin Respaldo Esperado',
                        color='Estrategia',
                        text='% Sin Respaldo Esperado',
                        title="Gestiones sin Respaldo Técnico Esperado",
                        labels={'% Sin Respaldo Esperado': '% sin respaldo esperado', 'Estrategia': 'Estrategia Declarada'}
                    )
                    fig_respaldo.update_traces(texttemplate='%{text:.1f}%')
                    fig_respaldo.update_layout(yaxis_ticksuffix="%")
                    st.plotly_chart(fig_respaldo, use_container_width=True)
                    st.caption("Validación esperada: las gestiones declaradas como Plataforma web deben aparecer en Twilio, y las declaradas como Celular/Tablet deben aparecer en Cámara.")

                    if (resumen_respaldo['% Sin Respaldo Esperado'] > 2).any():
                        st.error("Una o más estrategias superan el margen esperado del 2% sin respaldo técnico. Se recomienda revisar completitud de logs, fechas de corte y consistencia de los teléfonos.")
                    else:
                        st.success("Todas las estrategias están dentro del margen esperado menor o igual al 2% sin respaldo técnico.")
                
            st.stop()

        elif hoja_seleccionada == "Comparativa vs. Asignación":
            st.header("Análisis de Cobertura: Gestión vs. Asignación")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_e = datos.get("Entregados", pd.DataFrame())

            if not df_g.empty and not df_e.empty:
                # 1. Éxitos por Constructora vs Asignación
                st.subheader("1. Efectividad Real vs. Universo Asignado (Por Constructora)")
                
                exitos_const = df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito Total"].groupby("Constructora").size().reset_index(name="Encuestas Efectivas")
                asig_const = df_e.groupby("Constructora").size().reset_index(name="Asignación Total")
                
                comp_const = pd.merge(asig_const, exitos_const, on="Constructora", how="left").fillna(0)
                comp_const_melt = comp_const.melt(id_vars="Constructora", var_name="Métrica", value_name="Cantidad")
                
                fig_const = px.bar(comp_const_melt, x="Constructora", y="Cantidad", color="Métrica", barmode="group",
                                  title="Cumplimiento de Meta por Constructora",
                                  color_discrete_map={"Asignación Total": "#CBD5E0", "Encuestas Efectivas": "#2F855A"},
                                  text_auto=True)
                st.plotly_chart(fig_const, use_container_width=True)

                # 2. Gestiones por Región vs Asignación
                st.subheader("2. Esfuerzo de Gestión vs. Distribución de Base (Por Ciudad)")
                
                gest_ciu = df_g.groupby("Ciudad").size().reset_index(name="Gestiones Realizadas (Intentos)")
                asig_ciu = df_e.groupby("Ciudad").size().reset_index(name="Asignación Original")
                
                comp_ciu = pd.merge(asig_ciu, gest_ciu, on="Ciudad", how="left").fillna(0)
                comp_ciu_melt = comp_ciu.melt(id_vars="Ciudad", var_name="Métrica", value_name="Cantidad")
                
                fig_ciu = px.bar(comp_ciu_melt, x="Ciudad", y="Cantidad", color="Métrica", barmode="group",
                                title="Intensidad de Trabajo vs. Tamaño de Base por Región",
                                color_discrete_map={"Asignación Original": "#A0AEC0", "Gestiones Realizadas (Intentos)": "#3182CE"},
                                text_auto=True)
                st.plotly_chart(fig_ciu, use_container_width=True)
            else:
                st.warning("Se requieren datos en las hojas 'Base_gestiones realizadas' y 'Entregados' para este análisis.")
            st.stop()

        df_base = datos[hoja_seleccionada]

        st.header(f"Dashboard: {hoja_seleccionada}")
        
        # Metricas Principales
        col1, col2, col3 = st.columns(3)
        if hoja_seleccionada != "Twilio":
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
                exito_count = (df_base[col_agrupado] == "Éxito Total").sum()
                total_registros = len(df_base)
                if total_registros > 0:
                    porcentaje_exito = (exito_count / total_registros) * 100
                    col3.metric("Tasa de Éxito", f"{porcentaje_exito:.1f}%")
                else:
                    col3.metric("Tasa de Éxito", "0.0%")
            else:
                col3.metric("Tasa de Éxito", "N/A")

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

            if "segundos" in df_base.columns:
                duraciones_validas = df_base["segundos"].dropna()
                if not duraciones_validas.empty:
                    col4, col5, col6, col7 = st.columns(4)
                    col4.metric("Duración Promedio", f"{duraciones_validas.mean():.1f}s")
                    col5.metric("Llamada Más Larga", f"{duraciones_validas.max():.0f}s")
                    col6.metric("Llamada Más Corta", f"{duraciones_validas.min():.0f}s")
                    col7.metric("Mediana Duración", f"{duraciones_validas.median():.1f}s")

        elif hoja_seleccionada == "Twilio":
            col1.metric("Total Llamadas Twilio", len(df_base))
            
            if "Duration" in df_base.columns:
                duraciones_validas = df_base["Duration"].dropna()
                if not duraciones_validas.empty:
                    col4, col5, col6, col7 = st.columns(4)
                    col4.metric("Duración Promedio", f"{duraciones_validas.mean():.1f}s")
                    col5.metric("Llamada Más Larga", f"{duraciones_validas.max():.0f}s")
                    col6.metric("Llamada Más Corta", f"{duraciones_validas.min():.0f}s")
                    col7.metric("Mediana Duración", f"{duraciones_validas.median():.1f}s")

        # Visualizaciones de limpieza solicitadas
        if hoja_seleccionada in ["Base_gestiones realizadas", "Contactados"]:
            tab1, tab2 = st.tabs(["Análisis de Resultados", "Productividad Encuestadores"])
            
            with tab1:
                st.subheader("Resultados Estandarizados")
                col_agrupado = "Resultado de la gestión (Agrupado)"
                if col_agrupado in df_base.columns:
                    fig_res = px.pie(df_base, names=col_agrupado, 
                                     title="Distribución de Gestiones (Máximo 5 Categorías)")
                    fig_res.update_traces(textinfo='percent+label')
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
                                     color=col_encuestador, text_auto=True)
                    st.plotly_chart(fig_enc, use_container_width=True)
                else:
                    st.info(f"La columna '{col_encuestador}' no está disponible en esta hoja.")

        elif hoja_seleccionada in ["Entregados", "Correo Masivo"]:
            tab_labels = ["Análisis General", "Productividad Encuestadores", "Distribución Geográfica"]
            if hoja_seleccionada == "Correo Masivo":
                tab_labels[1] = "Análisis por Proyecto/Ciudad"
            tab1, tab2, tab3 = st.tabs(tab_labels)

            with tab1:
                st.subheader(f"Análisis General: {hoja_seleccionada}")
                col_constructora = "Constructora"
                if col_constructora in df_base.columns:
                    counts = df_base[col_constructora].value_counts()
                    # Agrupación si supera 5 categorías
                    if len(counts) > 5:
                        top_5 = counts.head(5)
                        otros = pd.Series({'Otros': counts.iloc[5:].sum()})
                        counts = pd.concat([top_5, otros])
                    
                    constructora_counts = counts.reset_index()
                    constructora_counts.columns = [col_constructora, 'count']
                    
                    fig_const = px.bar(constructora_counts, x=col_constructora, y="count",
                                       title=f"{hoja_seleccionada} por Constructora (Top 5)",
                                       labels={'count': 'Número de Entregados'}, text_auto=True)
                    st.plotly_chart(fig_const, use_container_width=True)
                else:
                    st.info(f"La columna '{col_constructora}' no está disponible en esta hoja.")

            with tab2:
                if hoja_seleccionada == "Correo Masivo":
                    st.subheader("Correo Masivo por Proyecto")
                    col_proyecto = "Proyecto"
                    if col_proyecto in df_base.columns:
                        proyectos = df_base[col_proyecto].value_counts().reset_index()
                        fig_proj = px.bar(proyectos, x=col_proyecto, y="count",
                                         labels={'count': 'Número de Registros'},
                                         color=col_proyecto, text_auto=True)
                        st.plotly_chart(fig_proj, use_container_width=True)
                    else:
                        st.info(f"La columna '{col_proyecto}' no está disponible en esta hoja.")

                    st.subheader("Correo Masivo por Ciudad")
                    col_ciudad_bar = "Ciudad"
                    if col_ciudad_bar in df_base.columns:
                        ciudades = df_base[col_ciudad_bar].value_counts().reset_index()
                        fig_city = px.bar(ciudades, x=col_ciudad_bar, y="count",
                                         labels={'count': 'Número de Registros'},
                                         color=col_ciudad_bar, text_auto=True)
                        st.plotly_chart(fig_city, use_container_width=True)
                    else:
                        st.info(f"La columna '{col_ciudad_bar}' no está disponible en esta hoja.")
                else:
                    st.subheader("Entregados por Encuestador (Nombres Limpios)")
                    col_encuestador = "Encuestador"
                    if col_encuestador in df_base.columns:
                        encuestadores = df_base[col_encuestador].value_counts().reset_index()
                        fig_enc = px.bar(encuestadores, x=col_encuestador, y="count",
                                         labels={'count': 'Número de Entregados'},
                                         color=col_encuestador, text_auto=True)
                        st.plotly_chart(fig_enc, use_container_width=True)
                    else:
                        st.info(f"La columna '{col_encuestador}' no está disponible en esta hoja.")

            with tab3:
                st.subheader("Distribución Geográfica de Entregados")
                col_ciudad = "Ciudad"
                col_ciudad2 = "Ciudad2"
                if col_ciudad in df_base.columns:
                    counts = df_base[col_ciudad].value_counts()
                    # Agrupación si supera 5 ciudades
                    if len(counts) > 5:
                        top_5 = counts.head(5)
                        otros = pd.Series({'Otros': counts.iloc[5:].sum()})
                        counts = pd.concat([top_5, otros])
                    
                    ciudad_counts = counts.reset_index()
                    ciudad_counts.columns = [col_ciudad, 'count']
                    fig_ciudad = px.pie(ciudad_counts, names=col_ciudad, values='count',
                                        title=f"Distribución Geográfica (Top 5)")
                    fig_ciudad.update_traces(textinfo='percent+label')
                    st.plotly_chart(fig_ciudad, use_container_width=True)
                elif col_ciudad2 in df_base.columns:
                    ciudad2_counts = df_base[col_ciudad2].value_counts().reset_index()
                    fig_ciudad2 = px.pie(ciudad2_counts, names=col_ciudad2,
                                         title="Entregados por Región (Ciudad2)")
                    fig_ciudad2.update_traces(textinfo='percent+label')
                    st.plotly_chart(fig_ciudad2, use_container_width=True)
                else:
                    st.info("No se encontraron columnas de ciudad para analizar.")

        elif hoja_seleccionada == "Camara_llamadas_salientes":
            tab1, tab2, tab3 = st.tabs(["Estadísticas de Llamadas", "Productividad por Encuestador", "Distribución Horaria"])

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
                    df_prod = df_base.copy()
                    df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
                    df_entregados = datos.get("Entregados", pd.DataFrame())

                    tels_exito = set()
                    if not df_g.empty and "Resultado de la gestión (Agrupado)" in df_g.columns and "tel_link" in df_g.columns:
                        tels_exito = set(df_g[df_g["Resultado de la gestión (Agrupado)"] == "Éxito Total"]["tel_link"].dropna().unique())

                    tels_entregados = set()
                    if not df_entregados.empty and "tel_link" in df_entregados.columns:
                        tels_entregados = set(df_entregados["tel_link"].dropna().unique())

                    df_prod["Llamada_Efectiva"] = df_prod["tel_link"].isin(tels_exito).astype(int) if "tel_link" in df_prod.columns else 0
                    df_prod["Numero_Entregado"] = df_prod["tel_link"].isin(tels_entregados).astype(int) if "tel_link" in df_prod.columns else 0

                    prod_enc = df_prod.groupby(col_enc).agg(
                        Minutos_Totales=(col_min, "sum"),
                        Llamadas_Realizadas=(col_enc, "count"),
                        Llamadas_Efectivas=("Llamada_Efectiva", "sum"),
                        Numeros_Entregados=("Numero_Entregado", "sum")
                    ).sort_values("Minutos_Totales", ascending=False).reset_index()

                    prod_enc_largo = prod_enc.melt(
                        id_vars=col_enc,
                        value_vars=["Minutos_Totales", "Llamadas_Realizadas", "Llamadas_Efectivas", "Numeros_Entregados"],
                        var_name="Métrica",
                        value_name="Total"
                    )

                    fig_prod = px.bar(prod_enc_largo, x=col_enc, y="Total", color="Métrica",
                                      title="Total Minutos, Llamadas, Efectivas y Entregados por Encuestador",
                                      labels={"Total": "Total", col_enc: "Encuestador"},
                                      barmode="group", text_auto=True)
                    st.plotly_chart(fig_prod, use_container_width=True)
                else:
                    st.info("Faltan columnas de encuestador o minutos para este análisis.")

            with tab3:
                st.subheader("Actividad por Hora (Picos de Llamadas)")
                if "fecha_llamada" in df_base.columns:
                    # Filtramos filas sin fecha válida para evitar que el gráfico salga vacío
                    df_hora = df_base.dropna(subset=["fecha_llamada"]).copy()
                    if not df_hora.empty:
                        df_hora["Hora"] = df_hora["fecha_llamada"].dt.hour
                        
                        df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
                        df_g_filtrado = df_g.copy()
                        col_encuestador_g = "Nombre del o de la encuestadora"
                        col_constructora_g = "Constructora"

                        if not df_g_filtrado.empty:
                            f1, f2 = st.columns(2)
                            if col_encuestador_g in df_g_filtrado.columns:
                                opciones_enc = ["Todos"] + sorted(df_g_filtrado[col_encuestador_g].dropna().unique().tolist())
                                sel_enc = f1.selectbox("Filtrar por Encuestador", opciones_enc, key="camara_hora_enc")
                                if sel_enc != "Todos":
                                    df_g_filtrado = df_g_filtrado[df_g_filtrado[col_encuestador_g] == sel_enc]
                            if col_constructora_g in df_g_filtrado.columns:
                                opciones_const = ["Todas"] + sorted(df_g_filtrado[col_constructora_g].dropna().unique().tolist())
                                sel_const = f2.selectbox("Filtrar por Constructora", opciones_const, key="camara_hora_const")
                                if sel_const != "Todas":
                                    df_g_filtrado = df_g_filtrado[df_g_filtrado[col_constructora_g] == sel_const]

                            if "tel_link" in df_g_filtrado.columns and "tel_link" in df_hora.columns:
                                tels_filtro = set(df_g_filtrado["tel_link"].dropna().unique())
                                df_hora = df_hora[df_hora["tel_link"].isin(tels_filtro)]

                        tels_exito = set()
                        if not df_g_filtrado.empty and "Resultado de la gestión (Agrupado)" in df_g_filtrado.columns and "tel_link" in df_g_filtrado.columns:
                            tels_exito = set(df_g_filtrado[df_g_filtrado["Resultado de la gestión (Agrupado)"] == "Éxito Total"]["tel_link"].dropna().unique())
                        
                        df_hora["Es_Exito"] = df_hora["tel_link"].isin(tels_exito).astype(int)
                        
                        hourly_stats = df_hora.groupby("Hora").agg(
                            Total_Llamadas=("Hora", "count"),
                            Llamadas_Exitosas=("Es_Exito", "sum")
                        ).reset_index()
                        
                        # Asegurar que las 24 horas estén presentes para que las líneas sean continuas
                        horas_full = pd.DataFrame({'Hora': range(24)})
                        hourly_calls = pd.merge(horas_full, hourly_stats, on='Hora', how='left').fillna(0)
                        hourly_calls = hourly_calls.sort_values("Hora")

                        # Transformar a formato largo para mostrar etiquetas en ambas líneas
                        hourly_calls = hourly_calls.melt(id_vars='Hora', var_name='Métrica', value_name='Cantidad')
                        
                        fig_hour = px.line(hourly_calls, x="Hora", y="Cantidad", color="Métrica",
                                         title="Volumen de Llamadas vs Éxitos por Hora del Día (Cámara)", 
                                         markers=True, text="Cantidad",
                                         color_discrete_map={'Total_Llamadas': '#1f77b4', 'Llamadas_Exitosas': '#ff7f0e'})
                        
                        fig_hour.update_traces(textposition="top center")
                        st.plotly_chart(fig_hour, use_container_width=True)
                    else:
                        st.warning("No se encontraron datos válidos en 'fecha_llamada' para generar el gráfico horario.")

        elif hoja_seleccionada == "Twilio":
            tab1, tab2 = st.tabs(["Estado y Calidad", "Distribución Horaria"])
            
            with tab1:
                st.subheader("Estado Técnico de Llamadas")
                if "Status" in df_base.columns:
                    fig_status = px.pie(df_base, names="Status", title="Efectividad de Conexión (Status)")
                    fig_status.update_traces(textinfo='percent+label')
                    st.plotly_chart(fig_status, use_container_width=True)
                
                if "Direction" in df_base.columns:
                    fig_dir = px.bar(df_base["Direction"].value_counts(), title="Dirección de las Llamadas", 
                                     text_auto=True)
                    st.plotly_chart(fig_dir, use_container_width=True)

            with tab2:
                st.subheader("Actividad por Hora (Picos de Llamadas)")
                if "Start Time" in df_base.columns:
                    # Filtramos filas sin fecha válida para evitar que el gráfico salga vacío
                    df_hora = df_base.dropna(subset=["Start Time"]).copy()
                    if not df_hora.empty:
                        df_hora["Hora"] = df_hora["Start Time"].dt.hour
                        
                        df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
                        df_g_filtrado = df_g.copy()
                        col_encuestador_g = "Nombre del o de la encuestadora"
                        col_constructora_g = "Constructora"

                        if not df_g_filtrado.empty:
                            f1, f2 = st.columns(2)
                            if col_encuestador_g in df_g_filtrado.columns:
                                opciones_enc = ["Todos"] + sorted(df_g_filtrado[col_encuestador_g].dropna().unique().tolist())
                                sel_enc = f1.selectbox("Filtrar por Encuestador", opciones_enc, key="twilio_hora_enc")
                                if sel_enc != "Todos":
                                    df_g_filtrado = df_g_filtrado[df_g_filtrado[col_encuestador_g] == sel_enc]
                            if col_constructora_g in df_g_filtrado.columns:
                                opciones_const = ["Todas"] + sorted(df_g_filtrado[col_constructora_g].dropna().unique().tolist())
                                sel_const = f2.selectbox("Filtrar por Constructora", opciones_const, key="twilio_hora_const")
                                if sel_const != "Todas":
                                    df_g_filtrado = df_g_filtrado[df_g_filtrado[col_constructora_g] == sel_const]

                            if "tel_link" in df_g_filtrado.columns and "tel_link" in df_hora.columns:
                                tels_filtro = set(df_g_filtrado["tel_link"].dropna().unique())
                                df_hora = df_hora[df_hora["tel_link"].isin(tels_filtro)]

                        tels_exito = set()
                        if not df_g_filtrado.empty and "Resultado de la gestión (Agrupado)" in df_g_filtrado.columns and "tel_link" in df_g_filtrado.columns:
                            tels_exito = set(df_g_filtrado[df_g_filtrado["Resultado de la gestión (Agrupado)"] == "Éxito Total"]["tel_link"].dropna().unique())
                        
                        df_hora["Es_Exito"] = df_hora["tel_link"].isin(tels_exito).astype(int)
                        
                        hourly_stats = df_hora.groupby("Hora").agg(
                            Total_Llamadas=("Hora", "count"),
                            Llamadas_Exitosas=("Es_Exito", "sum")
                        ).reset_index()
                        
                        # Asegurar que las 24 horas estén presentes para que las líneas sean continuas
                        horas_full = pd.DataFrame({'Hora': range(24)})
                        hourly_calls = pd.merge(horas_full, hourly_stats, on='Hora', how='left').fillna(0)
                        hourly_calls = hourly_calls.sort_values("Hora")

                        # Transformar a formato largo para mostrar etiquetas en ambas líneas
                        hourly_calls = hourly_calls.melt(id_vars='Hora', var_name='Métrica', value_name='Cantidad')
                        
                        fig_hour = px.line(hourly_calls, x="Hora", y="Cantidad", color="Métrica",
                                         title="Volumen de Llamadas vs Éxitos por Hora del Día (Twilio)", 
                                         markers=True, text="Cantidad",
                                         color_discrete_map={'Total_Llamadas': '#1f77b4', 'Llamadas_Exitosas': '#ff7f0e'})
                        
                        fig_hour.update_traces(textposition="top center")
                        st.plotly_chart(fig_hour, use_container_width=True)
                    else:
                        st.warning("No se encontraron datos válidos en 'Start Time' para generar el gráfico horario.")

        # Diccionario original (como referencia técnica)
        with st.expander("Ver Diccionario de Datos Técnico"):
            st.json(estructura)
else:
    st.error(f"No se encontró el archivo '{NOMBRE_ARCHIVO}' en el repositorio.")
