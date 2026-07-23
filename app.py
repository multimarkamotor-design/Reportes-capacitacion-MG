import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Generador de Reportes de Capacitación", layout="wide")

st.title("📊 Generador de Reportes de Capacitación")
st.markdown("Sube los archivos requeridos para procesar la asistencia y generar los formatos consolidados.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Reporte de Teams")
    file_teams = st.file_uploader("Sube el Excel de asistencia de Teams", type=["xlsx", "csv"])

with col2:
    st.subheader("2. Base de Datos Maestra")
    file_maestra = st.file_uploader("Sube el Excel con los datos de los empleados", type=["xlsx", "csv"])

if file_teams and file_maestra:
    try:
        # Extraer nombres de hojas si son archivos Excel
        xls_teams = pd.ExcelFile(file_teams) if file_teams.name.endswith('xlsx') else None
        xls_maestra = pd.ExcelFile(file_maestra) if file_maestra.name.endswith('xlsx') else None

        st.divider()
        st.subheader("⚙️ Configuración de los archivos")
        st.markdown("Selecciona la hoja y en qué fila están los títulos (Ej: Nombre, Correo, Cargo) para que el sistema encuentre los datos.")
        
        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            sheet_teams = st.selectbox("Hoja de Teams:", xls_teams.sheet_names) if xls_teams else 0
            # Por defecto las descargas directas de Teams están en la fila 1, o en la 10 si tienen resumen
            fila_teams = st.number_input("Fila de títulos en Teams:", min_value=1, value=1)
            
        with col_opt2:
            sheet_maestra = st.selectbox("Hoja Maestra:", xls_maestra.sheet_names) if xls_maestra else 0
            # En tus bases maestras los títulos están en la fila 4
            fila_maestra = st.number_input("Fila de títulos en Maestra (Tus bases usan la 4):", min_value=1, value=4)

        if st.button("🚀 Procesar y Generar Reportes"):
            # Leer datos saltando las filas vacías de arriba (skiprows)
            if file_teams.name.endswith('csv'):
                df_teams = pd.read_csv(file_teams, skiprows=fila_teams-1)
            else:
                df_teams = pd.read_excel(file_teams, sheet_name=sheet_teams, skiprows=fila_teams-1)
                
            if file_maestra.name.endswith('csv'):
                df_maestra = pd.read_csv(file_maestra, skiprows=fila_maestra-1)
            else:
                df_maestra = pd.read_excel(file_maestra, sheet_name=sheet_maestra, skiprows=fila_maestra-1)

            # Estandarizar nombres de columnas para que el sistema no se confunda
            df_teams.columns = df_teams.columns.astype(str).str.strip().str.lower()
            df_maestra.columns = df_maestra.columns.astype(str).str.strip().str.lower()

            # Buscar inteligentemente la columna de correo
            col_correo_teams = next((col for col in df_teams.columns if 'correo' in col), None)
            col_correo_maestra = next((col for col in df_maestra.columns if 'correo' in col), None)
            
            if not col_correo_teams or not col_correo_maestra:
                st.error("❌ No se encontró la columna de correo en uno de los archivos. Asegúrate de haber seleccionado la hoja y la fila correcta.")
            else:
                # Cruce maestro de bases de datos
                df_cruce = pd.merge(df_teams, df_maestra, left_on=col_correo_teams, right_on=col_correo_maestra, how='inner')

                # Calcular horas asumiendo 1 hora por sesión de forma predeterminada
                df_cruce['horas_calculadas'] = 1.0 

                # Construir la estructura de la Matriz Oficial
                cargos = ['Asistencial', 'Asesores', 'Coordinadores', 'Gerente']
                matriz = pd.DataFrame({
                    'Formación empleados': cargos,
                    'Hombres_Alcanzados': [0]*4,
                    'Hombres_Horas': [0.0]*4,
                    'Mujeres_Alcanzadas': [0]*4,
                    'Mujeres_Horas': [0.0]*4
                }).set_index('Formación empleados')

                # Llenar la matriz
                for index, row in df_cruce.iterrows():
                    cargo = str(row.get('cargo', '')).capitalize()
                    # Si no hay columna de sexo, asume Hombre para no romper el cálculo
                    sexo = str(row.get('sexo', 'Hombre')).capitalize() 
                    horas = row['horas_calculadas']
                    
                    if 'Asesor' in cargo: cargo = 'Asesores'
                    
                    if cargo in matriz.index:
                        if sexo == 'Hombre':
                            matriz.loc[cargo, 'Hombres_Alcanzados'] += 1
                            matriz.loc[cargo, 'Hombres_Horas'] += horas
                        elif sexo == 'Mujer':
                            matriz.loc[cargo, 'Mujeres_Alcanzadas'] += 1
                            matriz.loc[cargo, 'Mujeres_Horas'] += horas

                # Sumatorias finales
                matriz['Total empleados alcanzados'] = matriz['Hombres_Alcanzados'] + matriz['Mujeres_Alcanzadas']
                matriz['Total horas de formación'] = matriz['Hombres_Horas'] + matriz['Mujeres_Horas']

                st.success("✅ ¡Reportes cruzados y generados con éxito!")
                
                # Descargas
                col_desc1, col_desc2 = st.columns(2)
                with col_desc1:
                    st.subheader("Formato Oficial (Matriz)")
                    st.dataframe(matriz)
                    buffer_matriz = io.BytesIO()
                    with pd.ExcelWriter(buffer_matriz, engine='xlsxwriter') as writer:
                        matriz.to_excel(writer, sheet_name='Matriz Formación')
                    st.download_button("Descargar Matriz Excel", data=buffer_matriz.getvalue(), file_name="Horas_de_formacion_Calculado.xlsx")

                with col_desc2:
                    st.subheader("Base de Datos Detallada")
                    cols_detalladas = [c for c in df_cruce.columns if c != col_correo_teams]
                    df_detallado = df_cruce[cols_detalladas]
                    st.dataframe(df_detallado.head())
                    buffer_detallado = io.BytesIO()
                    with pd.ExcelWriter(buffer_detallado, engine='xlsxwriter') as writer:
                        df_detallado.to_excel(writer, index=False, sheet_name='Detalle')
                    st.download_button("Descargar Detalle Excel", data=buffer_detallado.getvalue(), file_name="Detalle_Capacitacion.xlsx")

    except Exception as e:
        st.error(f"Ocurrió un error inesperado al cruzar los datos: {e}")
