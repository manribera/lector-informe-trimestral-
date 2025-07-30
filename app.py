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

            # Buscar encabezados
            header_row_index = None
            for i in range(len(df)):
                if "Indicador" in df.iloc[i].astype(str).values:
                    header_row_index = i
                    break

            if header_row_index is None:
                st.warning(f"⚠️ No se encontró la fila de encabezado en '{archivo.name}'.")
                continue

            df_data = pd.read_excel(xls, sheet_name="Informe de avance", header=header_row_index, engine="openpyxl")

            # Mostrar columnas reales para confirmar
            st.write(f"🧩 Columnas detectadas en '{archivo.name}':")
            st.write(df_data.columns.tolist())

            # Buscar columnas de resultado
            columnas_resultado = [col for col in df_data.columns if str(col).lower().startswith("resultado")]

            for _, fila in df_data.iterrows():
                if (
                    pd.notna(fila.get("Líder Estratégico")) and
                    pd.notna(fila.get("Línea de Acción")) and
                    pd.notna(fila.get("Indicador")) and
                    pd.notna(fila.get("Meta"))
                ):
                    fila_resultado = {
                        "Delegación": delegacion,
                        "Líder Estratégico": fila.get("Líder Estratégico"),
                        "Línea de Acción": fila.get("Línea de Acción"),
                        "Indicador": fila.get("Indicador"),
                        "Meta": fila.get("Meta")
                    }

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
        st.info("ℹ️ No se encontraron datos válidos.")


