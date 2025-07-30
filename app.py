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
        df_preview = pd.read_excel(archivo, sheet_name="Informe de avance", header=None, engine="openpyxl")
        st.subheader("🔍 Vista previa de la hoja 'Informe de avance'")
        st.dataframe(df_preview.head(25))
    except Exception as e:
        st.error(f"❌ No se pudo leer el archivo: {e}")

# Función para procesar resultados por línea y trimestre
@st.cache_data
def procesar_resultados_por_linea(lista_archivos):
    resultado_final = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")
            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'.")
                continue

            df = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")
            delegacion = str(df.iloc[2, 7]).strip() if pd.notna(df.iloc[2, 7]) else "Desconocida"

            for _, fila in df.iterrows():
                linea = str(fila[1]).strip() if pd.notna(fila[1]) else None         # Columna B (índice 1)
                etiqueta_resultado = str(fila[4]).strip() if pd.notna(fila[4]) else None  # Columna E (índice 4)
                texto = str(fila[19]).strip() if pd.notna(fila[19]) else None      # Columna T (índice 19)

                if linea and etiqueta_resultado and texto:
                    if "resultado" in etiqueta_resultado.lower():
                        resultado_final.append({
                            "Delegación": delegacion,
                            "Línea": linea,
                            etiqueta_resultado: texto
                        })

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    if not resultado_final:
        return pd.DataFrame()

    # Convertir a DataFrame y agrupar por delegación y línea
    df_resultado = pd.DataFrame(resultado_final)
    df_resultado = df_resultado.groupby(["Delegación", "Línea"], as_index=False).first()
    return df_resultado

# Procesamiento principal
if archivos:
    with st.spinner("Procesando archivos..."):
        df_resultado = procesar_resultados_por_linea(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")
        st.subheader("📄 Resultados consolidados")
        st.dataframe(df_resultado)

        # Descargar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Resultados por Línea")

        st.download_button(
            label="📥 Descargar resultados en Excel",
            data=output.getvalue(),
            file_name="resultados_lineas_trimestres.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ No se encontraron resultados válidos para mostrar.")
