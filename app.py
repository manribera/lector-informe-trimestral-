import streamlit as st
import pandas as pd
import unicodedata
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# Cargar múltiples archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

@st.cache_data
def procesar_informes(lista_archivos):
    def normalize(s: str) -> str:
        if s is None:
            return ""
        s = str(s).strip().lower()
        s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
        while "  " in s:
            s = s.replace("  ", " ")
        return s

    def first_notna(values):
        for v in values:
            if pd.notna(v) and str(v).strip() not in ("", "nan", "None"):
                return v
        return None

    resultados = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")

            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'. Se omite.")
                continue

            # Lee toda la hoja sin encabezados para poder tomar G3 y luego la fila de encabezados
            df_raw = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")

            # Delegación en G3 (fila 3 -> index 2, columna G -> index 6)
            delegacion = str(df_raw.iloc[2, 6]).strip() if df_raw.shape[1] > 6 else ""

            # Fila de encabezados (visible como 10), índice 9:
            headers_row = 9
            if df_raw.shape[0] <= headers_row:
                st.warning(f"⚠️ '{archivo.name}': No hay suficientes filas para encabezados (se esperaba fila 10).")
                continue

            encabezados = df_raw.iloc[headers_row].fillna("").astype(str)
            # Normalizamos un mapa nombre_normalizado -> índice
            norm_to_idx = {normalize(h): i for i, h in enumerate(encabezados)}

            # Columnas objetivo por nombre
            wanted = {
                "líder estratégico": ["líder estratégico", "lider estrategico", "jefe estrategico"],
                "línea de acción": ["línea de acción", "linea de accion", "lineas de accion"],
                "indicador": ["indicador", "nombre del indicador"],
                "descripción del indicador": ["descripción del indicador", "descripcion del indicador", "descripcion indicador", "detalle del indicador"],
                "meta": ["meta", "meta anual", "meta trimestral"],
            }

            def get_col_idx(possible_names):
                for name in pos




