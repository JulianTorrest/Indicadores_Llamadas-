
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
    return s

def agrupar_resultado_gestion(valor):
    if pd.isna(valor): return "Otros"
    # Normalizar: minúsculas, sin espacios laterales y sin tildes
    v = "".join(c for c in unicodedata.normalize('NFD', str(valor).lower().strip()) if unicodedata.category(c) != 'Mn')

    # 1. Coincidencias Exactas (Basado en el requerimiento del usuario)
    if v in ["contesta y responde la encuesta", "contesta y responde la encuesta por forms"]:
        return "Éxito Total"
    if "encuesta incompleta" in v:
        return "Parcial / Incompleta"

    # 2. Éxito por palabras clave (Respaldo)
    if "encuesta" in v:
        if any(x in v for x in ["respond", "complet", "forms", "exito"]):
            return "Éxito Total"

    # 3. No Contactado
    if any(x in v for x in ["no contesta", "sin respuesta", "no entra", "invalido", "no contest"]):
        return "No Contactado"

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
                
                # Estandarizar fecha, día y teléfono
                if "Marca temporal" in df.columns:
                    df["Marca temporal"] = pd.to_datetime(df["Marca temporal"], errors='coerce', dayfirst=True)
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
                
                if "telefono_destino" in df.columns:
                    df["tel_link"] = df["telefono_destino"].apply(limpiar_telefono)

                # Estandarizar fecha de llamada
                if "fecha_llamada" in df.columns:
                    df["fecha_llamada"] = pd.to_datetime(df["fecha_llamada"], errors='coerce')

            # Limpieza específica para la hoja 'Twilio'
            elif nombre_hoja == "Twilio":
                # Estandarizar fechas
                for col_fecha in ["Start Time", "End Time", "Date Created"]:
                    if col_fecha in df.columns:
                        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
                
                if "To" in df.columns:
                    df["tel_link"] = df["To"].apply(limpiar_telefono)

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
        vistas = ["Funnel de Conversión", "Resumen de KPIs Críticos", "Análisis de Persistencia y Éxito", "Comportamiento 24h y Efectividad", "Análisis Cruzado (Auditoría)", "Base_gestiones realizadas", "Contactados", "Entregados (Base de Origen)", "Correo Masivo", "Camara_llamadas_salientes", "Twilio"]
        hoja_seleccionada = st.sidebar.selectbox("Seleccione la fuente de datos:", vistas)

        if hoja_seleccionada == "Funnel de Conversión":
            st.header("🏆 Embudo de Conversión de la Campaña")
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
                exitos_g = len(df_g_f[df_g_f["Resultado de la gestión (Agrupado)"].isin(["Éxito Total", "Parcial / Incompleta"])])

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
                        v_e_segment = len(df_g_f[(df_g_f["Ciudad"] == ciudad_sel) & (df_g_f["Resultado de la gestión (Agrupado)"].isin(["Éxito Total", "Parcial / Incompleta"]))])
                        
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
            st.header("🚀 Indicadores Críticos de Campaña")
            
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            df_c = datos.get("Camara_llamadas_salientes", pd.DataFrame())
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            
            if not df_g.empty:
                # 1. Tasa de Contactabilidad
                total = len(df_g)
                no_contactados = len(df_g[df_g["Resultado de la gestión (Agrupado)"] == "No Contactado"])
                contactabilidad = ((total - no_contactados) / total) * 100 if total > 0 else 0
                kpi1.metric("Contactabilidad", f"{contactabilidad:.1f}%", help="Porcentaje de gestiones que resultaron en un contacto real")

                # 2. Tasa de Conversión
                exitos = len(df_g[df_g["Resultado de la gestión (Agrupado)"].isin(["Éxito Total", "Parcial / Incompleta"])])
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

        elif hoja_seleccionada == "Análisis de Persistencia y Éxito":
            st.header("🎯 Análisis de Persistencia y Factores de Éxito")
            df_g = datos.get("Base_gestiones realizadas", pd.DataFrame())
            df_t = datos.get("Twilio", pd.DataFrame())
            df_c = datos.get("Camara_llamadas_salientes", pd.DataFrame())

            if not df_g.empty:
                # 1. Análisis de Intentos
                intentos_por_tel = df_g.groupby('tel_link').size().reset_index(name='Intentos')
                
                # Identificar si el número terminó en éxito (basado en df_g)
                if "Resultado de la gestión (Agrupado)" in df_g.columns:
                    tels_exito = set(df_g[df_g["Resultado de la gestión (Agrupado)"].isin(["Éxito Total", "Parcial / Incompleta"])]["tel_link"].dropna().unique())
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
                st.subheader("⏱️ Duración de Llamadas (Técnico) vs. Resultado de Gestión")
                
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
                st.subheader("📅 ¿Cuándo es mejor llamar?")
                if 'Dia_Semana' in df_g.columns and 'Hora' in df_g.columns:
                    # Marcamos cuales fueron exitosas (basado en df_g)
                    if "Resultado de la gestión (Agrupado)" in df_g.columns:
                        df_g['Es_Exito'] = df_g['tel_link'].isin(tels_exito)

                    # Pivot table para Heatmap
                    heatmap_data = df_g.groupby(['Dia_Semana', 'Hora'])['Es_Exito'].mean().reset_index()
                    heatmap_data['Es_Exito'] *= 100 # Convertir a porcentaje
                    
                    # Ordenar días correctamente
                    fig_heat = px.density_heatmap(heatmap_data, x='Hora', y='Dia_Semana', z='Es_Exito',
                                                  color_continuous_scale='Viridis',
                                                  title="Mapa de Calor: % de Efectividad (Conversión a Encuesta Exitosa)",
                                                  labels={'Es_Exito': '% Efectividad', 'Hora': 'Hora del Día', 'Dia_Semana': 'Día'},
                                                  category_orders={"Dia_Semana": ["1. Lunes", "2. Martes", "3. Miércoles", "4. Jueves", "5. Viernes", "6. Sábado", "7. Domingo"]})
                    st.plotly_chart(fig_heat, use_container_width=True)
                    st.caption("Los colores más claros (amarillo) indican horas y días con mayor probabilidad de que la gestión termine en una Entrega.")
                
                # 4. Análisis específico de "Llamar después" (Seguimiento)
                st.subheader("🔄 Evolución de los casos 'Llamar Después'")
                
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
                        st.metric("Tasa de Recuperación", f"{tasa_recuperacion:.1f}%", 
                                  help="Porcentaje de casos que pidieron 'Llamar después' y terminaron en Éxito")
                        st.write(f"De los {num_contactos_seg} casos, **{exitos_pos}** se convirtieron en éxito finalmente.")

                    with c_seg2:
                        fig_post = px.pie(conversion_seg, names='Resultado Posterior', values='Cantidad',
                                          title="Resultado Final de los contactos que pidieron Seguimiento",
                                          color_discrete_sequence=px.colors.qualitative.Pastel)
                        fig_post.update_traces(textinfo='percent+label')
                        st.plotly_chart(fig_post, use_container_width=True)

                    # Nuevo gráfico: ¿En qué intento se logra el éxito?
                    st.write("### 📈 Esfuerzo de Conversión Post-Seguimiento")
                    
                    # --- CAMBIO CRÍTICO AQUÍ ---
                    # 1. Identificar los tel_link que realmente terminaron en éxito (según la métrica de arriba)
                    successful_tel_links_from_metric = ultimo_estado[ultimo_estado['Resultado Posterior'].isin(["Éxito Total", "Parcial / Incompleta"])]['tel_link']
                    
                    # 2. Filtrar df_post para incluir solo las actividades de ESTOS contactos exitosos
                    df_post_successful_only = df_post[df_post['tel_link'].isin(successful_tel_links_from_metric)]
                    
                    # 3. Filtrar las gestiones de éxito dentro de este subconjunto
                    df_exitos_post_filtered = df_post_successful_only[df_post_successful_only['Resultado de la gestión (Agrupado)'].isin(["Éxito Total", "Parcial / Incompleta"])]
                    
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
            st.header("🕒 Comportamiento Temporal y Efectividad")
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
                    df_g_f['Efectivo'] = df_g_f["Resultado de la gestión (Agrupado)"].isin(["Éxito Total", "Parcial / Incompleta"])
                    
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
                    
                    st.info("💡 Consejo: Las líneas muestran la actividad de 0 a 23h. Si una línea está en 0, significa que no hubo actividad en esa franja para los filtros seleccionados.")
                else:
                    st.warning("Se requieren datos de 'Entregados' para medir efectividad.")
            st.stop()

        elif hoja_seleccionada == "Análisis Cruzado (Auditoría)":
            st.header("🔍 Auditoría: Gestión vs Telefonía")
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
                
                if False in df_g['Validado_Tecnico'].values:
                    st.subheader("Gestiones Manuales sin registro de llamada detectado")
                    st.dataframe(df_g[df_g['Validado_Tecnico'] == False][['Marca temporal', 'Nombre del o de la encuestadora', 'Número de teléfono sobre el que se realizó la gestión', 'Resultado de la gestión']].head(50))
            st.stop()

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
            tab1, tab2, tab3 = st.tabs(["Análisis General", "Productividad Encuestadores", "Distribución Geográfica"])

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
                                      color=col_min, text_auto=True)
                    st.plotly_chart(fig_prod, use_container_width=True)
                else:
                    st.info("Faltan columnas de encuestador o minutos para este análisis.")

        elif hoja_seleccionada == "Twilio":
            tab1, tab2, tab3 = st.tabs(["Estado y Calidad", "Análisis de Costos", "Distribución Horaria"])
            
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
