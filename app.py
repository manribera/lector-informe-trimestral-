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
            for i, col in enumerate(columnas_resultado[:4]):
                df_datos.rename(columns={col: f"Resultado {i+1}T"}, inplace=True)

            # Filtrar filas válidas
            df_filtrado = df_datos[df_datos["Indicadores"].notna() & df_datos["Meta"].notna()].copy()
            df_filtrado["Delegación"] = delegacion

            for _, fila in df_filtrado.iterrows():
                resultados.append({
                    "Delegación": delegacion,
                    "Líder Estratégico": fila.get("Lider"),
                    "Indicador": fila.get("Indicadores"),
                    "Meta": fila.get("Meta"),
                    "Cantidad": fila.get("Cantidad"),
                    "Avance General": fila.get("Avance General"),
                    "Resultado 1T": fila.get("Resultado 1T"),
                    "Resultado 2T": fila.get("Resultado 2T"),
                    "Resultado 3T": fila.get("Resultado 3T"),
                    "Resultado 4T": fila.get("Resultado 4T"),
                })

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    return pd.DataFrame(resultados)

# Procesamiento y exportación
if archivos:
    with st.spinner("Procesando archivos..."):
        df_resultado = procesar_informes(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")
        st.dataframe(df_resultado)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Resumen Indicadores")

        st.download_button(
            label="📥 Descargar resumen en Excel",
            data=output.getvalue(),
            file_name="resumen_indicadores_trimestral.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No se extrajeron datos válidos.")

