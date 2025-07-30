import streamlit as st
import pandas as pd
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# Subida de archivos
archivos = st.file_uploader("📁 Sube archivos .xlsm o .xlsx", type=["xlsm", "xlsx"], accept_multiple_files=True)

@st.cache_data
def procesar_informes(lista_archivos):
    resultados = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")
            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'. Se omite.")
                continue

            df = pd.read_excel(xls, sheet_name="Informe de avance", header=None, engine="openpyxl")
            delegacion = str(df.iloc[2, 7]).strip() if pd.notna(df.iloc[2, 7]) else "Desconocida"

            # Buscar fila de encabezado que contenga la palabra 'Indicadores'
            header_row_index = None
            for i in range(len(df)):
                if "Indicadores" in df.iloc[i].astype(str).values:
                    header_row_index = i
                    break

            if header_row_index is None:
                st.warning(f"⚠️ No se encontró fila de encabezado en '{archivo.name}'.")
                continue

            df_data = pd.read_excel(
                xls,
                sheet_name="Informe de avance",
                header=header_row_index,
                engine="openpyxl"
            )

            # Detectar columnas de resultados
            columnas_resultado = [col for col in df_data.columns if str(col).lower().startswith("resultado")]

            for _, fila in df_data.iterrows():
                if pd.notna(fila.get("Lider")) and pd.notna(fila.get("Linea de Accion")) and pd.notna(fila.get("Indicadores")) and pd.notna(fila.get("Meta")):
                    fila_resultado = {
                        "Delegación": delegacion,
                        "Líder Estratégico": fila.get("Lider"),
                        "Línea de Acción": fila.get("Linea de Accion"),
                        "Tipo de Indicador": fila.get("Indicadores"),
                        "Meta": fila.get("Meta")
                    }

                    # Añadir resultados por trimestre si existen
                    for col in columnas_resultado:
                        fila_resultado[col] = fila.get(col)

                    resultados.append(fila_resultado)

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    return pd.DataFrame(resultados)

# Procesamiento
if archivos:
    df_resultado = procesar_informes(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")
        st.dataframe(df_resultado)

        # Descargar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Resumen Indicadores")

        st.download_button(
            label="📥 Descargar resumen en Excel",
            data=output.getvalue(),
            file_name="resumen_informe_avance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No se encontraron datos válidos en los archivos cargados.")

