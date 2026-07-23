import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Generador de Reportes de Capacitación", layout="wide")

st.title("📊 Generador de Reportes de Capacitación")
st.markdown("Sube los archivos requeridos para procesar la asistencia y generar los formatos consolidados.")

# Sección de carga de archivos
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Reporte de Teams")
    file_teams = st.file_uploader("Sube el Excel de asistencia de Teams", type=["xlsx", "csv"])

with col2:
    st.subheader("2. Base de Datos Maestra")
    file_maestra = st.file_uploader("Sube el Excel con los datos de los empleados", type=["xlsx", "csv"])

if file_teams and file_maestra:
    try:
        # Lectura de datos
        if file_teams.name.endswith('csv'):
            df_teams = pd.read_csv(file_teams)
        else:
            df_teams = pd.read_excel(file_teams)
            
        if file_maestra.name.endswith('csv'):
            df_maestra = pd.read_csv(file_maestra)
        else:
            df_maestra = pd.read_excel(file_maestra)

        st.success("Archivos cargados correctamente. Procesando datos...")

        # ---------------------------------------------------------
        # LÓGICA DE PROCESAMIENTO
        # ---------------------------------------------------------
        
        # 1. Limpieza de nombres de columnas (para evitar errores por espacios)
        df_teams.columns = df_teams.columns.str.strip().str.lower()
        df_maestra.columns = df_maestra.columns.str.strip().str.lower()

        # Asumimos que las columnas clave se llaman así (puedes ajustarlas según tus archivos reales)
        col_correo_teams = 'correo electrónico' if 'correo electrónico' in df_teams.columns else df_teams.columns[1]
        col_correo_maestra = 'correo' if 'correo' in df_maestra.columns else df_maestra.columns[0]
        
        # 2. Cruce de bases de datos usando el correo electrónico
        df_cruce = pd.merge(
            df_teams, 
            df_maestra, 
            left_on=col_correo_teams, 
            right_on=col_correo_maestra, 
            how='inner'
        )

        # 3. Tratamiento de horas (Conversión de duración de Teams a horas decimales)
        # Asumiendo que Teams entrega "1h 30m" o minutos puros. 
        # Aquí simplificamos asignando un valor numérico a una columna 'horas_calculadas'
        # NOTA: Debes ajustar el parsing de tiempo según el formato exacto de tu Teams
        if 'duración' in df_cruce.columns:
            # Ejemplo simplificado: si viene en formato texto o numérico
            df_cruce['horas_calculadas'] = pd.to_numeric(df_cruce['duración'], errors='coerce').fillna(1.0)
        else:
            df_cruce['horas_calculadas'] = 1.0 # Valor por defecto por sesión

        # ---------------------------------------------------------
        # GENERACIÓN DEL FORMATO: Horas de formacion.xlsx
        # ---------------------------------------------------------
        
        # Estructura base del reporte matricial
        cargos = ['Asistencial', 'Asesores', 'Coordinadores', 'Gerente']
        matriz = pd.DataFrame({
            'Formación empleados': cargos,
            'Hombres_Alcanzados': [0]*4,
            'Hombres_Horas': [0.0]*4,
            'Mujeres_Alcanzadas': [0]*4,
            'Mujeres_Horas': [0.0]*4
        }).set_index('Formación empleados')

        # Llenado de la matriz con los datos cruzados
        for index, row in df_cruce.iterrows():
            cargo = str(row.get('cargo', '')).capitalize()
            sexo = str(row.get('sexo', '')).capitalize()
            horas = row['horas_calculadas']
            
            # Ajuste al nombre exacto del cargo para que coincida con la matriz
            if 'Asesor' in cargo: cargo = 'Asesores'
            
            if cargo in matriz.index:
                if sexo == 'Hombre':
                    matriz.loc[cargo, 'Hombres_Alcanzados'] += 1
                    matriz.loc[cargo, 'Hombres_Horas'] += horas
                elif sexo == 'Mujer':
                    matriz.loc[cargo, 'Mujeres_Alcanzadas'] += 1
                    matriz.loc[cargo, 'Mujeres_Horas'] += horas

        # Cálculos de totales
        matriz['Total empleados alcanzados'] = matriz['Hombres_Alcanzados'] + matriz['Mujeres_Alcanzadas']
        matriz['Total horas de formación'] = matriz['Hombres_Horas'] + matriz['Mujeres_Horas']

        # ---------------------------------------------------------
        # DESCARGA DE ARCHIVOS
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Descarga de Reportes")
        
        col_desc1, col_desc2 = st.columns(2)
        
        # 1. Reporte Consolidado (Matriz)
        with col_desc1:
            st.subheader("Formato Oficial (Matriz)")
            st.dataframe(matriz)
            
            buffer_matriz = io.BytesIO()
            with pd.ExcelWriter(buffer_matriz, engine='xlsxwriter') as writer:
                matriz.to_excel(writer, sheet_name='Matriz Formación')
            
            st.download_button(
                label="Descargar Formato Oficial en Excel",
                data=buffer_matriz.getvalue(),
                file_name="Horas_de_formacion_Calculado.xlsx",
                mime="application/vnd.ms-excel"
            )

        # 2. Reporte Detallado
        with col_desc2:
            st.subheader("Base de Datos Detallada")
            # Seleccionamos las columnas de interés para el reporte detallado
            cols_detalladas = [c for c in df_cruce.columns if c not in [col_correo_teams]]
            df_detallado = df_cruce[cols_detalladas]
            st.dataframe(df_detallado.head())
            
            buffer_detallado = io.BytesIO()
            with pd.ExcelWriter(buffer_detallado, engine='xlsxwriter') as writer:
                df_detallado.to_excel(writer, index=False, sheet_name='Detalle_Vitrina_Asesor')
            
            st.download_button(
                label="Descargar Detalle por Vitrina y Asesor",
                data=buffer_detallado.getvalue(),
                file_name="Detalle_Capacitacion.xlsx",
                mime="application/vnd.ms-excel"
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos. Verifica que los nombres de las columnas sean correctos. Detalle del error: {e}")
