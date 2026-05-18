import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Consolidador Líneas SIGESS", layout="wide")

st.title("📋 Consolidado de Líneas de Acción SIGESS 2026")
st.write("Carga los archivos Excel de las delegaciones para extraer la planificación de líneas de acción desde la hoja **Informe de avance**.")

archivos = st.file_uploader(
    "📁 Sube archivos .xlsm o .xlsx",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True
)

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def crear_id(delegacion, trimestre, consecutivo):
    delegacion_limpia = re.sub(r"\s+", "_", delegacion.strip())
    trimestre_limpio = re.sub(r"\s+", "_", trimestre.strip())
    return f"{delegacion_limpia}_{trimestre_limpio}_{consecutivo:03d}"

@st.cache_data
def procesar_informes(lista_archivos):
    resultados = []

    for archivo in lista_archivos:
        try:
            xls = pd.ExcelFile(archivo, engine="openpyxl")

            if "Informe de avance" not in xls.sheet_names:
                st.warning(f"⚠️ El archivo '{archivo.name}' no tiene hoja 'Informe de avance'. Se omite.")
                continue

            df = pd.read_excel(
                xls,
                sheet_name="Informe de avance",
                header=None,
                engine="openpyxl"
            )

            # Ajustar estas celdas si el formato cambia
            delegacion = limpiar_texto(df.iloc[2, 7]) if df.shape[0] > 2 and df.shape[1] > 7 else ""
            region = limpiar_texto(df.iloc[1, 7]) if df.shape[0] > 1 and df.shape[1] > 7 else ""
            trimestre = limpiar_texto(df.iloc[3, 7]) if df.shape[0] > 3 and df.shape[1] > 7 else ""

            encabezados = df.iloc[9]
            columnas_resultado = [
                i for i, val in enumerate(encabezados)
                if str(val).lower().strip().startswith("resultado")
            ]

            consecutivo = 1

            for i in range(10, len(df)):
                fila = df.iloc[i]

                lider = limpiar_texto(fila[3]) if len(fila) > 3 else ""
                linea = limpiar_texto(fila[4]) if len(fila) > 4 else ""
                indicador = limpiar_texto(fila[5]) if len(fila) > 5 else ""
                meta = limpiar_texto(fila[7]) if len(fila) > 7 else ""

                if lider and linea and indicador and meta:
                    id_linea = crear_id(delegacion, trimestre, consecutivo)

                    registro = {
                        "ID_LINEA": id_linea,
                        "Archivo Origen": archivo.name,
                        "Delegación Policial": delegacion,
                        "Delegación Regional": region,
                        "Trimestre": trimestre,
                        "Número de Línea": consecutivo,
                        "Línea de Acción": linea,
                        "Problemática": "",
                        "Líder Estratégico": lider,
                        "Indicador": indicador,
                        "Meta": meta,
                        "Responsable": "",
                        "Cogestor": "",
                        "Estado Base": "Activa"
                    }

                    for col_index in columnas_resultado:
                        nombre_col = limpiar_texto(encabezados[col_index])
                        registro[nombre_col] = fila[col_index]

                    resultados.append(registro)
                    consecutivo += 1

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    return pd.DataFrame(resultados)

if archivos:
    df_resultado = procesar_informes(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")

        st.subheader("Vista previa del consolidado")
        st.dataframe(df_resultado, use_container_width=True)

        st.metric("Total de líneas extraídas", len(df_resultado))
        st.metric("Total de delegaciones procesadas", df_resultado["Delegación Policial"].nunique())

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(
                writer,
                index=False,
                sheet_name="PLANIFICACION_LINEAS_BASE"
            )

        st.download_button(
            label="📥 Descargar PLANIFICACION_LINEAS_BASE",
            data=output.getvalue(),
            file_name="PLANIFICACION_LINEAS_BASE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.info("ℹ️ No se encontraron datos válidos.")


