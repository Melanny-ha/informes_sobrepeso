import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import os
from io import BytesIO
import datetime

# Especificar la ruta del archivo Excel
file_path_Mol = r"G:\.shortcut-targets-by-id\1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN\SIRI_2024_INFORMES_SOBREPESO\INFORME_EMPAQUES_2\Consolidado.xlsx"

# ===========================================================================
# Funciones
# ===========================================================================
#Funcion para convertir DataFrame a Excel(Descargar tabla)
def convertir_a_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    return output.getvalue()

#Funcion para reducir decimales visualmente
def formatear_visual(df, columnas_0f=[], columnas_2f=[]):
    df_mostrar = df.copy()
    for col in columnas_0f:
        df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{x:.0f}")
    for col in columnas_2f:
        df_mostrar[col] = df_mostrar[col].apply(lambda x: f"{x:.2f}")
    return df_mostrar
# ===========================================================================

#Verificar si el archivo existe
if not os.path.exists(file_path_Mol):
    st.info("No se encontró el archivo de datos en la ruta especificada.")
else:
    #Leer el archivo Excel con Pandas
    df = pd.read_excel(file_path_Mol, sheet_name="Hoja1")

    #Convertir los nombres de las columnas a str y eliminar espacios adicionales
    df.columns = df.columns.astype(str).str.strip()

    #Renombrar las columnas utilizadas
    df.rename(columns={'Día:':'Dia', 'Mes:':'Mes', 'Año:':'Año', 'Máquina / Equipo:':'Máquina / Equipo', 'Turno:':'Turno', 'Unidades Producidas (Conformes) :':'Unidades Producidas',
                       'Real_empaque_calculado':'Cantidad Real Empacada [kg]', 'Debe_ser_empaque_calculado':'Cantidad Teórica a Empacar [kg]',
                       'Diferencia_calculado':'Diferencia [kg]', 'Sobrepeso_calculado':'Sobrepeso [%]'}, inplace=True)
    
    #Convertir los tipos de datos
    if 'Codigo' in df.columns:
        df['Codigo'] = df['Codigo'].astype(str)
    if 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'])
    if 'Año' in df.columns:
        df['Año'] = df['Año'].astype(int)

    ## ==============================================================================
    ## Barra literal (navegación y filtros)
    ## ==============================================================================
    #Subtitulo
    st.sidebar.markdown('<h2 style="font-size: 23px; color: #eb4f37; font-weight: bold;">Filtros de datos</h2>', unsafe_allow_html=True)

    #Rangos seleccion de fechas
    min_fecha = df['Fecha'].min()
    max_fecha = df['Fecha'].max()
    fecha_por_defecto = datetime.date(2025, 1, 1)
    #Se crean los input para seleccionar fecha
    fecha_inicio = st.sidebar.date_input("Fecha inicio:", min_value=min_fecha, max_value=max_fecha, value=fecha_por_defecto)
    fecha_fin = st.sidebar.date_input("Fecha fin:", min_value=fecha_inicio, max_value=max_fecha) #Fecha fin no puede ser menor a fecha inicio pero si igual

    #Se crea una copia del dataframe para comenzar a filtrar los datos
    df_filtrado = df.copy()

    #Se filtran los posibles datos en base a la fecha
    if 'Fecha' in df_filtrado.columns:
        df_filtrado = df_filtrado[(df_filtrado['Fecha'] >= pd.to_datetime(fecha_inicio)) & (df_filtrado['Fecha'] <= pd.to_datetime(fecha_fin))] #Rango de fechas permitido

    #Filtrar las maquinas disponibles en base a ese rango de fechas filtrado
    maquinas_disponibles = df_filtrado['Máquina / Equipo'].unique()

    #Verificar si hay maquinas disponibles luego del filtro
    if len(maquinas_disponibles) > 0:
        maquina_seleccionada = st.sidebar.multiselect("Seleccione Máquina/Equipo:", maquinas_disponibles)
    else:
        #Mostrar mensaje
        st.info("No hay máquinas disponibles en este rango de fechas.")

    #Filtrar lso datos en base a la maquina seleccionada
    if maquina_seleccionada:
        df_filtrado = df_filtrado[df_filtrado['Máquina / Equipo'].isin(maquina_seleccionada)]

    #Filtrar solo los turnos disponible spara ese rango de fechas y maquina seleccionados
    turnos_disponibles = df_filtrado['Turno'].unique()

    #Verificar si hay turnos disponibles despues de los filtros
    if len(turnos_disponibles) > 0:
        turno_seleccionado = st.sidebar.multiselect("Seleccione Turno:", turnos_disponibles) #¡¡¡¡CORREGIR!!!! variable en Envase Soluble
    else:
        st.info("No hay turnos disponibles en este rango de fechas y/o para esta(s) máquina(s).")

    #Filtrar en base al turno seleccionado
    if turno_seleccionado:
        df_filtrado = df_filtrado[df_filtrado['Turno'].isin(turno_seleccionado)]

    #Filtrar por codigo en base a la maquina y turno seleccionado
    codigo_disponible = df_filtrado['Codigo'].unique()

    #Verificar si hay codigos disponibles para dichos filtros
    if len(codigo_disponible) > 0:
        codigo_seleccionado = st.sidebar.multiselect("Seleccione Código:", codigo_disponible)
    else:
        st.info("No hay codigos disponibles enlazados a estos filtros") #¡¡¡¡CORREGIR!!!! mensajes por: no hay codigos disponibles para estos filtros(filtros en general)

    #Filtrar los datos en base al codigo seleccionado
    if codigo_seleccionado:
        df_filtrado = df_filtrado[df_filtrado['Codigo'].isin(codigo_seleccionado)]


    #Calcular tabla pivote incluyendo codigo, maquina/equipo(ambos para agrupar por codigo y maquina por mes) y mes_n(para ordenar)
    df_Emp2_MES = df_filtrado.pivot_table(
        index=['Año', 'Mes', 'IdMes', 'Codigo', 'Máquina / Equipo'],
        values=['Unidades Producidas', 'Gramaje (K):', 'Peso Promedio de la unidad (K):'],
        aggfunc={'Unidades Producidas': 'sum', 'Gramaje (K):': 'mean',
                 'Peso Promedio de la unidad (K):': 'mean'}
    ).reset_index()

    # Calcular nuevas columnas por cada dato
    df_Emp2_MES['Cantidad Real Empacada [kg]'] = df_Emp2_MES['Unidades Producidas'] * df_Emp2_MES['Peso Promedio de la unidad (K):']
    df_Emp2_MES['Cantidad Teórica a Empacar [kg]'] = df_Emp2_MES['Gramaje (K):'] * df_Emp2_MES['Unidades Producidas']
    df_Emp2_MES['Diferencia [kg]'] = df_Emp2_MES['Cantidad Real Empacada [kg]'] - df_Emp2_MES['Cantidad Teórica a Empacar [kg]']
    df_Emp2_MES['Sobrepeso [%]'] = (df_Emp2_MES['Diferencia [kg]'] / df_Emp2_MES['Cantidad Teórica a Empacar [kg]']) * 100

    #Crear tabla pivote agrupando por Año y mes, inclutendo Mes_N para ordenar(sumar todas las columnas menos las del index)
    df_Emp2_MES_visual = df_Emp2_MES.pivot_table(
        index=['Año', 'Mes', 'IdMes'],
        values=['Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]'],
        aggfunc={'Unidades Producidas':'sum', 
                 'Cantidad Real Empacada [kg]':'sum', 
                 'Cantidad Teórica a Empacar [kg]':'sum', 
                 'Diferencia [kg]':'sum'}
    ).reset_index()

    #Se calcula nuevamente el sobrepeso siendo este ahora el sobrepeso total por mes
    df_Emp2_MES_visual['Sobrepeso [%]'] = (df_Emp2_MES_visual['Diferencia [kg]']/df_Emp2_MES_visual['Cantidad Teórica a Empacar [kg]']) * 100
    
    # Ordenar por Año y por Mes los dataframe
    df_Emp2_MES_visual = df_Emp2_MES_visual.sort_values(by=['Año', 'IdMes'], ascending=[True, True])

    #Eliminar las columnas que ya no se necesitan o listar solo las necesarias
    #df_Env_MES = df_Env_MES.drop(columns=['Mes_N', 'Codigo', 'Máquina / Equipo:', 'Gramaje (K):', 'Peso Promedio de la unidad (K):'])
    #df_Env_MES_visual = df_Env_MES_visual[['Año', 'Mes', 'Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]', 'Sobrepeso [%]']]

    ## ==============================================================================
    ## Vista principal según la opción seleccionada
    ## ==============================================================================
    #Si el dataframe esta vacio muestra un mensaje, sino continua
    if df_Emp2_MES.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        col_izq, col_der = st.columns([1, 3])
        with col_izq:
            #Logo
            st.image("images/logo.webp", width=120)
            #Metricas con valores totales
            st.markdown(f'<div class="metric-card">Total Teórico [kg]<br><b>{df_Emp2_MES_visual['Cantidad Real Empacada [kg]'].sum():,.0f}</b></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-card">Total Real [kg]<br><b>{df_Emp2_MES_visual['Cantidad Teórica a Empacar [kg]'].sum():,.0f}</b></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-card">Diferencia Total [kg]<br><b>{df_Emp2_MES_visual['Diferencia [kg]'].sum():,.0f}</b></div>', unsafe_allow_html=True)
            sobrepeso_ratio = df_Emp2_MES_visual['Diferencia [kg]'].sum() / df_Emp2_MES_visual['Cantidad Teórica a Empacar [kg]'].sum()
            st.markdown(f'<div class="metric-card">Sobrepeso Total [%]<br><b>{sobrepeso_ratio:.2%}</b></div>', unsafe_allow_html=True)

        with col_der:
            ## ==============================================================================
            ## Tabla principal informe Empaques 2
            ## ==============================================================================
            #Titulo tabla1
            st.markdown(
                """
                <div style="display: flex; justify-content: flex-end; padding: 0;">
                    <h1 style="font-size: 15px; margin: 0; color: black;">Informe Mensual de Sobrepeso - Empaques 2</h1>
                </div>
                """, unsafe_allow_html=True)
            
            #Columnas seleccionadas para el dataframe a mostrar
            columnas=['Año', 'Mes', 'Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]', 'Sobrepeso [%]']
            df_Emp2_MES_visual=df_Emp2_MES_visual[columnas]

            #Quitar decimales a dichas columnas
            columnas_0f = ['Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]']
            df_Emp2_MES_visual[columnas_0f] = df_Emp2_MES_visual[columnas_0f].apply(lambda x: x.map("{:.0f}".format))
            
            #Limitar la columna Sobrepeso a 2 decimales
            df_Emp2_MES_visual['Sobrepeso [%]'] = df_Emp2_MES_visual['Sobrepeso [%]'].apply(lambda x: f"{x:.2f}")

            #Mostrar tabla con funcion formatear
            st.markdown(f"""
                <div class="custom-table">
                    {df_Emp2_MES_visual.to_html(index=False)}
                </div>
            """, unsafe_allow_html=True)

            #Generar el Excel en base a la tabla visual
            excel_data = convertir_a_excel(df_Emp2_MES_visual)

            #Botón de descarga
            st.download_button(
                label="Descargar tabla como Excel",
                data=excel_data,
                file_name="informe_mensual_sobrepeso.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        ## ==============================================================================
        ## Heatmap de sobrepeso por máquina y código
        ## ==============================================================================
        #Crear tabla pivote de df_Env_MEs para el histograma
        pivot_heatmap = df_Emp2_MES.pivot_table(
                index='Máquina / Equipo', 
                columns='Codigo', 
                values='Sobrepeso [%]', 
                aggfunc='mean')
        
        #Convertir a unidades porcentuales df_filtrado['Sobrepeso [%]']= df_filtrado['Sobrepeso [%]' *100
        df_filtrado['Sobrepeso [%]'] *= 100
        col_izq, col_der = st.columns([2, 2])
        with col_izq:
            #Gráfico de dispersión en base a la tabla con línea de sobrepeso ideal en 1.5%
            if 'Fecha' in df_filtrado.columns and 'Diferencia [kg]' in df_filtrado.columns and 'Cantidad Teórica a Empacar [kg]' in df_filtrado.columns:
                #Grafico
                grafico = px.scatter(
                    df_filtrado, 
                    x='Fecha', 
                    y='Sobrepeso [%]', 
                    title='Sobrepeso real vs Meta sobrepeso')
                #Linea meta sobrepeso
                grafico.add_hline(
                    y=1.5, 
                    line_dash="solid", 
                    line_color="red", 
                    annotation_text="Meta 1.5%")
                #Centrar titulo grafico
                grafico.update_layout(
                    title={
                        'text': 'Distribución Sobrepeso por Máquina y Código',
                        'x': 0.5,  #0.5 lo lleva al centro horizontal
                        'xanchor': 'center'  #Ancla el título en el centro
                    }
                )
                #Muestra el grafico
                st.plotly_chart(grafico)
            else:
                st.info("No se encontraron datos suficientes para generar el gráfico de sobrepeso por fecha.")
        with col_der:
            if 'Sobrepeso [%]' in df_Emp2_MES.columns:
                ## ==============================================================================
                ## Agregar histograma según código
                ## ==============================================================================
                #Eje x
                histograma = px.histogram(df_Emp2_MES, x='Sobrepeso [%]', nbins=25, title='Histograma de Sobrepeso')
                #Eje y
                histograma.update_yaxes(title_text="Frecuencia")
                histograma.update_traces(marker_color="#6374CC")
                #Linea Meta sobrepeso
                histograma.add_vline(x=1.5, line_dash="solid", line_color="red", annotation_text="Meta 1.5%")
                #Mostrar el histograma
                st.plotly_chart(histograma)
            else:
                st.info("No se encontró la columna 'Sobrepeso' en los datos.")

        ## ==============================================================================
        ## Grafico de cajas y bigotes
        ## ==============================================================================
        st.markdown(
                """
                <div style="display: flex; justify-content: center; padding: 10px;">
                    <h1 style="font-size: 15px; margin: 0; color: black;">Distribución Sobrepeso por Máquina y Código</h1>
                </div>
                """, unsafe_allow_html=True)
        #Si la columna sobrepeso existe en el dataframe
        if 'Sobrepeso [%]' in df_Emp2_MES.columns:
            #Crear grafico
            fig, ax = plt.subplots(figsize=(8, 12))
            #Define x y y
            fig_box = px.box(
                df_Emp2_MES, 
                x='Máquina / Equipo', 
                y='Sobrepeso [%]', 
                color='Codigo')
            #Añadir linea de meta sobrepeso
            fig_box.add_hline(
                y=1.5, 
                line_dash="solid", 
                line_color="red", 
                annotation_text="Meta 1.5%")
            #Mostrar el grafico
            st.plotly_chart(fig_box)
        else:
            st.info("No se encontró la columna 'Sobrepeso' en los datos.")
    
        ## ==============================================================================
        ## Graficos y dataframes de Pareto
        ## ==============================================================================
        # ===========================================
        # DataFrame de pareto Cantidad Real a Empacar
        # ===========================================
        #Titulo
        st.markdown(
                """
                <div style="display: flex; justify-content: center; padding: 10px;">
                    <h1 style="font-size: 15px; margin: 0; color: black;">Código con mayor Producción</h1>
                </div>
                """, unsafe_allow_html=True)

        #Definir columnas a mostrar
        columnas_pareto = ['Codigo', 'Máquina / Equipo', 'Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]']

        #Slider de porcentaje acumulado 
        Acumulado_Real_Empaque = st.slider("Selecciona el porcentaje acumulado a visualizar:", min_value=0.0, max_value=100.0, value=80.0, key="Cantidad real empacada")

        #Copia en base al dataframe df_Env_mes
        df_ranking_prod = df_Emp2_MES[columnas_pareto].copy()

        # Agrupar por Código y Máquina / Equipo y sumar valores
        df_ranking_prod = df_ranking_prod.groupby(['Codigo', 'Máquina / Equipo'], as_index=False).sum()

        # Ordenar por Real empacado en orden descendente
        df_ranking_prod = df_ranking_prod.sort_values(by='Cantidad Real Empacada [kg]', ascending=False)

        # Agregar la columna 'Ranking' y '[%] Acumulado'
        df_ranking_prod["Ranking"] = range(1, len(df_ranking_prod) + 1)
        df_ranking_prod["[%] Acumulado"] = df_ranking_prod["Cantidad Real Empacada [kg]"].cumsum() / df_ranking_prod["Cantidad Real Empacada [kg]"].sum() * 100

        #Si se ecoge un acumulado menor al menor en el dataframe genera un aviso, sino continua
        if Acumulado_Real_Empaque < df_ranking_prod['[%] Acumulado'].min():
            st.info("No hay datos con dicho acumulado")
        else:
            #Filtra los acumulados menores o iguales al acumulado escogido
            df_ranking_prod = df_ranking_prod.loc[df_ranking_prod['[%] Acumulado'] <= Acumulado_Real_Empaque]

            #Reduce visualmente los decimales
            columnas_0f = ['Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]']
            columnas_2f = ['[%] Acumulado']
            df_ranking_prod_visual = formatear_visual(df_ranking_prod, columnas_0f, columnas_2f)

            #Muestra la tabla visual, no con los tipos datos internos
            st.markdown(f"""
                <div class="custom-table">
                    {df_ranking_prod_visual.to_html(index=False)}
                </div>
            """, unsafe_allow_html=True)

            # ===========================================
            # Boton de descarga DataFrame a Excel
            # ===========================================
            #Generar el Excel en base a la tabla visual
            excel_data = convertir_a_excel(df_ranking_prod_visual)

            #Botón de descarga
            st.download_button(
                label="Descargar tabla como Excel",
                data=excel_data,
                file_name="tabla_produccion_por_codigo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ===========================================
        # Grafico de Pareto de Cantidad real empacada
        # ===========================================

        #Se genera una copia para evitar errores
        df_Pareto_prod = df_ranking_prod.copy()

        #Crear la figura del pareto
        fig, ax1 = plt.subplots(figsize=(10, 6))

        #Barras verticales con mayor ancho
        ax1.bar(df_Pareto_prod["Codigo"], df_Pareto_prod["Cantidad Real Empacada [kg]"], color="C0", alpha=0.7, width=0.6)
        ax1.set_ylabel("Cantidad Real Empacada [kg]", color="C7")
        ax1.tick_params(axis="y", labelcolor="C7")

        #Línea de % Acumulado
        ax2 = ax1.twinx()
        ax2.plot(df_Pareto_prod["Codigo"], df_Pareto_prod["[%] Acumulado"], color="r", marker="o", linestyle="-")
        ax2.set_ylabel("[%] Acumulado", color="C7")
        ax2.tick_params(axis="y", labelcolor="C7")

        #Barra horizontal en el 80% acumulado
        ax2.axhline(80, color="r", linestyle="--", alpha=0.6)
        ax2.text(df_Pareto_prod["Codigo"].iloc[-1], 81, "80% Acumulado", color="r", fontsize=9)

        # Rotar etiquetas de código en el eje X
        #plt.xticks(rotation=90, ha='left')
        ax1.set_xticklabels(df_Pareto_prod["Codigo"], rotation=90, ha='left')

        # Título y layout
        plt.title("Cantidad Real Empacada")
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        st.pyplot(fig)

        # ===========================================
        # DataFrame de pareto Sobrepeso
        # ===========================================
        #Titulo
        st.markdown(
                """
                <div style="display: flex; justify-content: center; padding: 10px;">
                    <h1 style="font-size: 15px; margin: 0; color: black;">Código con mayor Sobrepeso</h1>
                </div>
                """, unsafe_allow_html=True)

        #linea desliz para elegirla visualizacion en cuanto a porcentaje acumulado del dataframe
        Acumulado_Sobrepeso = st.slider("Selecciona el porcentaje acumulado a visualizar:", min_value=0.0, max_value=100.0, value=80.0,key="Cantidad sobrepeso")

        df_ranking_sobre = df_Emp2_MES[columnas_pareto].copy()

        # Agrupar por Código y Máquina / Equipo y sumar valores
        df_ranking_sobre = df_ranking_sobre.groupby(['Codigo', 'Máquina / Equipo'], as_index=False).sum()

        df_ranking_sobre["Sobrepeso [%]"] = (df_ranking_sobre["Diferencia [kg]"] / df_ranking_sobre['Cantidad Teórica a Empacar [kg]']) * 100

        # Ordenar por sobrepeso en orden descendente
        df_ranking_sobre = df_ranking_sobre.sort_values(by='Sobrepeso [%]', ascending=False)

        # Agregar la columna 'Ranking'
        df_ranking_sobre["Ranking"] = range(1, len(df_ranking_sobre) + 1)
        df_ranking_sobre["[%] Acumulado"] = df_ranking_sobre["Sobrepeso [%]"].cumsum() / df_ranking_sobre["Sobrepeso [%]"].sum() * 100

        #Si el aumulado elegido es menor al acumulado minimo del dataframe genera un aviso, sino continua
        if Acumulado_Sobrepeso < df_ranking_sobre['[%] Acumulado'].min():
            st.info("No hay datos con dicho acumulado")
        else:
            #Mostrar los acumulados del dataframe menores o iguales al acumulado escogido
            df_ranking_sobre=df_ranking_sobre.loc[df_ranking_sobre['[%] Acumulado'] <= Acumulado_Sobrepeso]

            #Quitar decimales a dichas columnas
            columnas_0f = ['Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]']
            columnas_2f=['[%] Acumulado', 'Sobrepeso [%]']
            #Formatear la tabla
            df_ranking_sobre_visual = formatear_visual(df_ranking_sobre,columnas_0f,columnas_2f)

            #Mostrar tabla con funcion formatear
            st.markdown(f"""
                <div class="custom-table">
                    {df_ranking_sobre_visual.to_html(index=False)}
                </div>
            """, unsafe_allow_html=True)

            # ===========================================
            # Boton de descarga DataFrame a Excel
            # ===========================================
            #Generar el Excel en base a la tabla visual
            excel_data = convertir_a_excel(df_ranking_sobre_visual)

            #Botón de descarga
            st.download_button(
                label="Descargar tabla como Excel",
                data=excel_data,
                file_name="tabla_sobrepeso_por_codigo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ===========================================
        # Grafico de Pareto de Sobrepeso
        # ===========================================
        df_Pareto_sobre = df_ranking_sobre

        # Crear la figura del pareto
        fig1, ax1_1 = plt.subplots(figsize=(10, 6))

        # Barras verticales con mayor ancho
        ax1_1.bar(df_Pareto_sobre["Codigo"], df_Pareto_sobre["Sobrepeso [%]"], color="C0", alpha=0.7, width=0.6)
        ax1_1.set_ylabel("Sobrepeso [%]", color="C7")
        ax1_1.tick_params(axis="y", labelcolor="C7")

        # Línea de % Acumulado
        ax2_2 = ax1_1.twinx()
        ax2_2.plot(df_Pareto_sobre["Codigo"], df_Pareto_sobre["[%] Acumulado"], color="red", marker="o", linestyle="-")
        ax2_2.set_ylabel("[%] Acumulado", color="C7")
        ax2_2.tick_params(axis="y", labelcolor="C7")

        # Barra horizontal en el 80% acumulado
        ax2_2.axhline(80, color="r", linestyle="--", alpha=0.6)
        ax2_2.text(df_Pareto_sobre["Codigo"].iloc[-1], 81, "80% Acumulado", color="r", fontsize=9)

        # Rotar etiquetas de código en el eje X
        #plt.xticks(rotation=-90, ha='right')
        ax1_1.set_xticklabels(df_Pareto_sobre["Codigo"], rotation=90, ha='left')

        # Título y layout
        plt.title("Sobrepeso")
        plt.grid(axis="y", linestyle="--", alpha=0.7)

        # Mostrar en Streamlit
        st.pyplot(fig1)


        ## ==============================================================================
        ## Tabla novedades sobrepeso
        ## ==============================================================================
        st.markdown(
                """
                <div style="display: flex; justify-content: center; padding: 10px;">
                    <h1 style="font-size: 15px; margin: 0; color: black;">Novedades Sobrepeso</h1>
                </div>
                """, unsafe_allow_html=True)
        
        #Variable fija de sobrepeso maximo(Meta)
        Sobrepeso_Novedades = 1.5

        #Filtrar los sobrepesos mayores a la meta en base a los filtros ingresados
        df_novedades_sobrepeso = df_filtrado[df_filtrado['Sobrepeso [%]'] >= Sobrepeso_Novedades]
        
        #Linea desliz para definir rango de sobrepeso a visualizar
        rango_sobrepeso = st.slider(
            'Rango sobrepeso [%]:', 
            min_value=1.5, 
            max_value=6.0, 
            value=(1.5, 12.0)
        )

        #Filtrar los datos en base al sobrepeso seleccionado
        if 'Sobrepeso [%]' in df_novedades_sobrepeso.columns:
            df_novedades_sobrepeso = df_novedades_sobrepeso[
                (df_novedades_sobrepeso['Sobrepeso [%]'] >= rango_sobrepeso[0]) &
                (df_novedades_sobrepeso['Sobrepeso [%]'] <= rango_sobrepeso[1])
        ]
            
        #Listar las columnas a mostrar
        columnas=['Año', 'Mes', 'Dia', 'Máquina / Equipo', 'Codigo', 'Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]', 'Sobrepeso [%]']
        df_novedades_sobrepeso = df_novedades_sobrepeso[columnas]
        
        #Quitar decimales a dichas columnas
        columnas_0f = ['Unidades Producidas', 'Cantidad Real Empacada [kg]', 'Cantidad Teórica a Empacar [kg]', 'Diferencia [kg]']
        df_novedades_sobrepeso[columnas_0f] = df_novedades_sobrepeso[columnas_0f].apply(lambda x: x.map("{:.0f}".format))

        #Limitar la columna Sobrepeso a 2 decimales
        df_novedades_sobrepeso['Sobrepeso [%]'] = df_novedades_sobrepeso['Sobrepeso [%]'].apply(lambda x: f"{x:.2f}")

        #si existe la columna sobrepeso en el dataframe continua
        if 'Sobrepeso [%]' in df_novedades_sobrepeso:
            #Si no esta vacio el dataframe novedades
            if not df_novedades_sobrepeso['Sobrepeso [%]'].empty:
                #Mostrar tabla con funcion formatear
                st.markdown(f"""
                    <div class="custom-table">
                        {(df_novedades_sobrepeso).to_html(index=False)}
                    </div>
                """, unsafe_allow_html=True)

                # ===========================================
                # Boton de descarga DataFrame a Excel
                # ===========================================
                #Generar el Excel en base a la tabla visual
                excel_data = convertir_a_excel(df_novedades_sobrepeso)

                #Botón de descarga
                st.download_button(
                    label="Descargar tabla como Excel",
                    data=excel_data,
                    file_name="tabla_novedades_sobrepeso.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                #Dejar fijo encabezado
                st.markdown(f"""
                    <div class="custom-table">
                        {(df_novedades_sobrepeso).to_html(index=False)}
                    </div>
                """, unsafe_allow_html=True)
                st.info("No hay datos con sobrepeso para estos filtros.")
        else:
            st.info("No hay datos con sobrepeso para estos filtros.")

## ==============================================================================
# Estilos personalizados(CSS)
## ==============================================================================
st.markdown("""
    <style>
        /*widht->ancho  heigth->alto*/
        /*Estilos generales*/
        .main {background-color: #f0f0f0;}
        .sidebar header{color: red;}
        .image { padding: 200px;}
        
        /*Estilos para las tablas*/
        .custom-table {
            max-height: 300px;
            overflow-y: auto; /* Solo scroll vertical */
            display: block;
            width: 100%;
            margin: 10px 0 12px 0;
            border-radius: none;
        }
        .custom-table table {
            table-layout: auto; /* Permite que las columnas se ajusten automáticamente */
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            min-width: 20px;
        }
        .dataframe{font-size: 12px;}
        .dataframe tbody tr td {
            font-size: 10px;
            padding: 5px;
            white-space: nowrap;
        }
        .custom-table th, .custom-table td {
            border: 1px solid #ddd;  /*grosor bordes tabla*/
            padding: 8px;  /*ancho del cuadro del titulo*/  
            text-align: center;
            word-wrap: break-word;
            white-space: normal;
        }
        .custom-table th {  /*color fondo y letra encabezado*/
            background-color: #ff4630;  
            color: white;
        }
        .custom-table tr:nth-child(even){  /*filas gris oscuras en la tabla*/
            background-color: #dedede;
        }
        .custom-table tr:hover {  /*cambiar color de fila al pasar por ella*/
            background-color: #fc735d;
            color: white;
        }
        .dataframe tbody tr td {
            font-size: 14px !important;  /* Tamaño del texto */
            padding: 0px 4px !important;  /* Ancho dentro de la celda */
        }
        .dataframe thead tr th {
            font-size: 12px !important;  /* Tamaño del texto del encabezado */
            padding: 5px !important;  /* Espaciado dentro de las celdas del encabezado */
        }
            
        /*Estilos para el boton descargar*/
        .stDownloadButton {
            display: flex;
            justify-content: flex-end; /* Alinea el contenedor a la derecha */
            margin-top: 5px;
        }
        .stDownloadButton>button {
            background-color: #fe5237;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 10px;
            font-size: 10px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.2s ease;
        }
        .stDownloadButton>button:hover {
            background-color: #c3270f;
            color: white;
        }

        /*Estilos para las tarjetas*/
        .metric-card {
            background-color: white;
            margin: 6px;
            padding: 12px;
            border-radius: 10px;
            box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
            font-size: 13px;
            text-align: center;
        }
    </style>
""", unsafe_allow_html=True)