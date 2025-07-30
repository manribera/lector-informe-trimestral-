import streamlit as st
import pandas as pd
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores y resultados por trimestre desde la hoja **'Informe de avance'**.")

# Subida de múltiples archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

@st.cache_data
def procesar_informes(lista_archivos):
    resultados = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")
            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'.")
                continue

            df = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")
            delegacion = str(df.iloc[2, 7]).strip() if pd.notna(df.iloc[2, 7]) else "Desconocida"

            # Encabezados en fila 9 (índice 9), datos reales desde fila 10
            columnas = df.iloc[9].tolist()
            df_datos = df.iloc[10:].copy()
            df_datos.columns = columnas

            # Buscar columnas de resultados repetidas
            columnas_resultado = [col for col in df_datos.columns if str(col).strip().lower().startswith("resultado")]

            # Renombrar a Resultado 1T, 2T, 3T, 4T
            for i, c

