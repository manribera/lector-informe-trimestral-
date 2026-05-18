import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Consolidador Líneas SIGESS", layout="wide")

st.title("📋 Consolidado de Líneas de Acción SIGESS 2026")
st.write("Carga archivos Excel para extraer líneas, indicadores y metas desde la hoja Informe de avance.")

archivos = st.file_uploader(
    "📁 Sube archivos .xlsm o .xlsx",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True
)

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def normalizar_lider(valor):
    valor = limpiar_texto(valor).lower()

    if "municipal" in valor or "gobierno local" in valor or valor == "gl":
        return "Gobierno Local"

    if "fuerza" in valor or valor == "fp":
        return "Fuerza Pública"

    return limpiar_texto(valor)

def extraer_numero_linea(texto):
    texto = limpiar_texto(texto)
    match = re.search(r"#\s*(\d+)", texto)
    if match:
        return int(match.group(1))
    return ""

def crear_id_registro(delegacion, trimestre, numero_linea, numero_indicador):
    delegacion = re.sub(r"\s+", "_", limpiar_texto(delegacion))
    trimestre = re.sub(r"\s+", "_", limpiar_texto(trimestre))
    return f"{delegacion}_{trimestre}_L{numero_linea}_I{numero_indicador}"

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

            delegacion = limpiar_texto(df.iloc[2, 7]) if df.shape[0] > 2 and df.shape[1] > 7 else ""
            region = ""

            # Buscar bloques de Línea de Acción
            for i in range(len(df)):
                fila = df.iloc[i]

                texto_linea = limpiar_texto(fila[3]) if len(fila) > 3 else ""

                if texto_linea.lower().startswith("linea de accion"):
                    numero_linea = extraer_numero_linea(texto_linea)

                    problematica = limpiar_texto(fila[5]) if len(fila) > 5 else ""
                    lider_bloque = normalizar_lider(fila[7]) if len(fila) > 7 else ""

                    # Buscar trimestres en el bloque
                    trimestre = limpiar_texto(fila[10]) if len(fila) > 10 else ""

                    # Los indicadores normalmente empiezan 4 filas después del título del bloque
                    fila_inicio_indicadores = i + 4
                    fila_fin_bloque = i + 20

                    numero_indicador_real = 1

                    for j in range(fila_inicio_indicadores, min(fila_fin_bloque, len(df))):
                        fila_ind = df.iloc[j]

                        sigla_lider = limpiar_texto(fila_ind[2]) if len(fila_ind) > 2 else ""
                        responsable = limpiar_texto(fila_ind[3]) if len(fila_ind) > 3 else ""
                        indicador_num = limpiar_texto(fila_ind[4]) if len(fila_ind) > 4 else ""
                        indicador = limpiar_texto(fila_ind[5]) if len(fila_ind) > 5 else ""
                        meta = limpiar_texto(fila_ind[7]) if len(fila_ind) > 7 else ""

                        # Evita filas vacías tipo "Indicador 4" sin descripción ni meta
                        if not indicador or not meta:
                            continue

                        lider_estrategico = lider_bloque

                        if not lider_estrategico:
                            lider_estrategico = normalizar_lider(responsable or sigla_lider)

                        id_registro = crear_id_registro(
                            delegacion,
                            trimestre,
                            numero_linea,
                            numero_indicador_real
                        )

                        resultados.append({
                            "ID_REGISTRO": id_registro,
                            "Archivo Origen": archivo.name,
                            "Delegación Policial": delegacion,
                            "Delegación Regional": region,
                            "Trimestre": trimestre,
                            "Número de Línea": numero_linea,
                            "Problemática": problematica,
                            "Líder Estratégico": lider_estrategico,
                            "Responsable": responsable,
                            "Indicador Número": indicador_num,
                            "Indicador": indicador,
                            "Meta": meta,
                            "Estado Base": "Activa"
                        })

                        numero_indicador_real += 1

        except Exception as e:
            st.error(f"❌ Error procesando '{archivo.name}': {e}")

    return pd.DataFrame(resultados)

if archivos:
    df_resultado = procesar_informes(archivos)

    if not df_resultado.empty:
        st.success("✅ Archivos procesados correctamente.")

        st.metric("Total de indicadores extraídos", len(df_resultado))
        st.metric("Total de delegaciones procesadas", df_resultado["Delegación Policial"].nunique())
        st.metric("Total de líneas reales", df_resultado[["Delegación Policial", "Número de Línea"]].drop_duplicates().shape[0])

        st.subheader("Vista previa del consolidado")
        st.dataframe(df_resultado, use_container_width=True)

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
