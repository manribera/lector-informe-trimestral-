import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Resultados por Línea y Trimestre", layout="wide")
st.title("📋 Consolidado de Resultados por Línea y Trimestre")
st.write("Carga archivos Excel (.xlsm o .xlsx) para extraer resultados desde la hoja **'Informe de avance'**.")

# Subida de archivos
archivos = st.file_uploader("📁 Sube archivos Excel", type=["xlsm", "xlsx"], accept_multiple_files=True)

# 🔍 Vista previa de la hoja
if archivos:
    archivo = archivos[0]
    try:
        df_preview = pd.read_excel(archivo, sheet_name="Informe_
