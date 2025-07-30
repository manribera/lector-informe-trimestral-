import streamlit as st
import pandas as pd
import io

st.title("📋 Consolidado de Resultados por Línea y Trimestre")
st.write("Carga uno o varios archivos Excel para extraer resultados desde la hoja **'Informe de avance'**, organizados por línea (L1, L2, ...) y trimestre.")

# Subida de archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

@st.cache_data
def procesar_resultados_por_linea(lista_archivos):
    re
