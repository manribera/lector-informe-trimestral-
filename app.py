import streamlit as st
import pandas as pd
import io

st.title("📋 Consolidado de Indicadores - Informe de Avance")
st.write("Carga uno o varios archivos Excel para extraer indicadores desde la hoja **'Informe de avance'**.")

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

            # Delegación en G3 (fila 3 -> idx 2, col G -> idx 6)
            delegacion = ""
            if df.shape[0] > 2 and df.shape[1] > 6:
                delegacion = str(df.iloc[2, 6]).strip()

            # Encabezados en fila 10 visible -> idx 9
            if df.shape[0] <= 9:
                st.warning(f"⚠️ '{archivo.name}': no hay suficientes filas para leer encabezados (se esperaba fila 10).")
                continue

            encabezados = df.iloc[9].fillna("").astype(str)
            # Detectar índices de columnas que empiezan con "Resultado"
            columnas_resultado = {i for i, val in enumerate(encabezados) if str(val).strip().lower().startswith("resultado")}
            # Incluir N (idx 13) y T (idx 19) si existen
            if df.shape[1] > 13: columnas_resultado.add(13)
            if df.shape[1] > 19: columnas_resultado.add(19)

            # Datos desde la fila 11 -> idx 10
            for r in range(10, len(df)):
                fila = df.iloc[r]
                if len(fila) <= 7:
                    continue

                lider = fila[3]    # D
                linea = fila[4]    # E
                indicador = fila[5]# F
                meta = fila[7]     # H

                # Si está vacía la fila clave, saltar
                if pd.isna(lider) and pd.isna(linea) and pd.isna(indicador) and pd.isna(meta):
                    continue

                # Unificar Resultado: prioriza T(19) -> N(13) -> cualquier otra columna resultado con dato
                valor_T = fila[19] if (19 in columnas_resultado and 19 < len(fila)) else None
                valor_N = fila[13] if (13 in columnas_resultado and 13 < len(fila)) else None

                unificado = None
                if pd.notna(valor_T) and str(valor_T).strip() != "":
                    unificado = valor_T
                elif pd.notna(valor_N) and str(valor_N).strip() != "":
                    unificado = valor_N
                else:
                    # cualquier otro "Resultado*" con contenido
                    for ci in sorted(columnas_resultado):
                        if ci < len(fila):
                            v = fila[ci]
                            if pd.notna(v) and str(v).strip() != "":
                                unificado = v
                                break

                fila_out = {
                    "Delegación": delegacion,
                    "Líder Estratégico": lider,
                    "Línea de Acción": linea,
                    "Indicador": indicador,
                    "Descripción del Indicador": None,  # si quieres, mapea por índice/nombre real
                    "Meta": meta,
                    "Resultado": unificado
                }

                resultados.append(fila_out)

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    df_final = pd.DataFrame(resultados)
    # Orden final fijo
    cols = ["Delegación", "Líder Estratégico", "Línea de Acción", "Indicador", "Descripción del Indicador", "Meta", "Resultado"]
    for c in cols:
        if c not in df_final.columns:
            df_final[c] = None
    return df_final[cols] if not df_final.empty else df_final

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

