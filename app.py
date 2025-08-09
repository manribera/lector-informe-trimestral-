import streamlit as st
import pandas as pd
import unicodedata
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# Cargar múltiples archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    while "  " in s:
        s = s.replace("  ", " ")
    return s

@st.cache_data
def procesar_informes(lista_archivos):
    resultados = []

    # Variantes toleradas por campo
    VARS = {
        "lider": ["líder estratégico", "lider estrategico", "jefe estrategico", "líder estrategico"],
        "linea": ["línea de acción", "linea de accion", "lineas de accion", "líneas de acción"],
        "indicador": ["indicador", "nombre del indicador", "indicador (texto)"],
        "descripcion": ["descripción del indicador", "descripcion del indicador", "descripcion indicador", "detalle del indicador", "definicion del indicador"],
        "meta": ["meta", "meta anual", "meta trimestral", "meta del indicador"],
    }

    def get_idx(norm_to_idx, posibles):
        for name in posibles:
            n = normalize(name)
            if n in norm_to_idx:
                return norm_to_idx[n]
        return None

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")
            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'. Se omite.")
                continue

            df = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")

            # Delegación en G3 (fila 3 -> idx 2, col G -> idx 6)
            delegacion = ""
            if df.shape[0] > 2 and df.shape[1] > 6:
                delegacion = str(df.iloc[2, 6]).strip()

            # Encabezados en fila 10 visible -> idx 9
            if df.shape[0] <= 9:
                st.warning(f"⚠️ '{archivo.name}': no hay suficientes filas para leer encabezados (se esperaba fila 10).")
                continue

            encabezados = df.iloc[9].fillna("").astype(str)
            norm_to_idx = {normalize(h): i for i, h in enumerate(encabezados)}

            # Ubicar índices por nombre (tolerante a tildes/variantes)
            idx_lider = get_idx(norm_to_idx, VARS["lider"])
            idx_linea = get_idx(norm_to_idx, VARS["linea"])
            idx_indicador = get_idx(norm_to_idx, VARS["indicador"])
            idx_desc = get_idx(norm_to_idx, VARS["descripcion"])
            idx_meta = get_idx(norm_to_idx, VARS["meta"])

            # Columnas de Resultado detectadas por encabezado
            columnas_resultado = {i for i, h in enumerate(encabezados) if normalize(h).startswith("resultado")}
            # Forzar N (14 -> idx 13) y T (20 -> idx 19) si existen
            if df.shape[1] > 13: columnas_resultado.add(13)
            if df.shape[1] > 19: columnas_resultado.add(19)

            # Datos desde la fila 11 -> idx 10
            for r in range(10, df.shape[0]):
                fila = df.iloc[r]

                # Leer campos requeridos si existen los índices
                lider = fila[idx_lider] if idx_lider is not None and idx_lider < len(fila) else None
                linea = fila[idx_linea] if idx_linea is not None and idx_linea < len(fila) else None
                indicador = fila[idx_indicador] if idx_indicador is not None and idx_indicador < len(fila) e_]()

