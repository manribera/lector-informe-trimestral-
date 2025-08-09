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

            # Delegación en G3 (fila 3 -> index 2, col G -> index 6)
            delegacion = ""
            if df.shape[0] > 2 and df.shape[1] > 6:
                delegacion = str(df.iloc[2, 6]).strip()

            # Encabezados en fila 10 visible -> índice 9
            if df.shape[0] <= 9:
                st.warning(f"⚠️ '{archivo.name}': no hay suficientes filas para leer encabezados (se esperaba fila 10).")
                continue

            encabezados = df.iloc[9].fillna("").astype(str)

            # 1) Detectar todas las columnas cuyo encabezado empiece con "Resultado"
            columnas_resultado = {i for i, val in enumerate(encabezados) if str(val).strip().lower().startswith("resultado")}

            # 2) Forzar incluir N y T aunque no digan "Resultado"
            #    N = columna 14 (índice 13), T = columna 20 (índice 19)
            if df.shape[1] > 13:
                columnas_resultado.add(13)
            if df.shape[1] > 19:
                columnas_resultado.add(19)

            # Datos desde la fila 11 -> índice 10
            for i in range(10, len(df)):
                fila = df.iloc[i]

                # Seguridad: necesitamos al menos hasta H (índice 7) para 'meta'
                if len(fila) <= 7:
                    continue

                lider = fila[3]  # D
                linea = fila[4]  # E
                tipo_indicador = fila[5]  # F
                meta = fila[7]   # H

                if pd.notna(lider) and pd.notna(linea) and pd.notna(tipo_indicador) and pd.notna(meta):
                    fila_out = {
                        "Delegación": delegacion,
                        "Líder Estratégico": lider,
                        "Línea de Acción": linea,
                        "Tipo de Indicador": tipo_indicador,
                        "Meta": meta
                    }

                    # Guardar valores de N y T explícitamente (si existen)
                    valor_N = None
                    valor_T = None

                    # Copiar todas las columnas de resultado detectadas (incluidas N y T)
                    for col_index in sorted(columnas_resultado):
                        if col_index < len(fila):
                            nombre_col = str(encabezados[col_index]).strip()
                            # Si nombre vacío, etiquetar por letra
                            if not nombre_col:
                                nombre_col = "Resultado " + ("N" if col_index == 13 else "T" if col_index == 19 else f"Col{col_index+1}")
                            fila_out[nombre_col] = fila[col_index]

                            if col_index == 13:
                                valor_N = fila[col_index]
                            elif col_index == 19:
                                valor_T = fila[col_index]

                    # Campo unificado "Resultado": prioriza T sobre N; si ambos vacíos, pone el primero no nulo de cualquier "Resultado*"
                    unificado = None
                    if pd.notna(valor_T) and str(valor_T).strip() != "":
                        unificado = valor_T
                    elif pd.notna(valor_N) and str(valor_N).strip() != "":
                        unificado = valor_N
                    else:
                        # Busca algún otro "Resultado*" con contenido
                        for col_index in sorted(columnas_resultado):
                            if col_index < len(fila):
                                val = fila[col_index]
                                if pd.notna(val) and str(val).strip() != "":
                                    unificado = val
                                    break

                    fila_out["Resultado"] = unificado
                    resultados.append(fila_out)

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
        st.warning("No se encontraron filas válidas para consolidar.")
else:
    st.info("Sube uno o varios archivos para comenzar.")

