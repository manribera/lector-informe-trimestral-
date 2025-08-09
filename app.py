import streamlit as st
import pandas as pd
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

# Cargar múltiples archivos
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

            # Celda G3 -> fila índice 2, columna índice 6
            delegacion = ""
            if df.shape[0] > 2 and df.shape[1] > 6:
                delegacion = str(df.iloc[2, 6]).strip()

            # Encabezados visibles en fila 10 -> índice 9
            if df.shape[0] <= 9:
                st.warning(f"⚠️ '{archivo.name}': no hay suficientes filas para leer encabezados (se esperaba fila 10).")
                continue

            encabezados = df.iloc[9]
            columnas_resultado = [i for i, val in enumerate(encabezados) if str(val).strip().lower().startswith("resultado")]

            # Datos desde la fila 11 -> índice 10
            for i in range(10, len(df)):
                fila = df.iloc[i]

                # Verificación mínima de ancho (necesitamos al menos hasta la col 7 para 'meta')
                if len(fila) <= 7:
                    continue

                lider = fila[3]
                linea = fila[4]
                tipo_indicador = fila[5]
                meta = fila[7]

                if pd.notna(lider) and pd.notna(linea) and pd.notna(tipo_indicador) and pd.notna(meta):
                    fila_resultado = {
                        "Delegación": delegacion,
                        "Líder Estratégico": lider,
                        "Línea de Acción": linea,
                        "Tipo de Indicador": tipo_indicador,
                        "Meta": meta
                    }

                    # Extraer todas las columnas cuyo encabezado empiece con "Resultado"
                    for col_index in columnas_resultado:
                        if col_index < len(fila):
                            nombre_col = str(encabezados[col_index])
                            fila_resultado[nombre_col] = fila[col_index]

                    resultados.append(fila_resultado)

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    return pd.DataFrame(resultados)

# Procesamiento
if archivos:
    df_resultado = procesar_informes(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")
        st.dataframe(df_resultado, use_container_width=True)

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
        # <<--- ESTE BLOQUE FALTABA Y CAUSABA EL ERROR DEL 'else:' VACÍO
        st.warning("No se encontraron filas válidas para consolidar.")
else:
    st.info("Sube uno o varios archivos para comenzar.")

