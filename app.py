# app.py
import io
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Lector Sembremos Seguridad", layout="wide")

st.title("Lector de Indicadores – Sembremos Seguridad")
st.caption("Sube el libro y extrae: Delegación, Líder Estratégico, Problemática Priorizada, Indicador, Descripción del Indicador, Meta, Resultados 1T, Resultados 2T.")

# === Utilidades ===
REQ_HEADERS = [
    "Delegación",
    "Líder Estratégico",
    "Problemática Priorizada",
    "Indicador",
    "Descripción del Indicador",
    "Meta",
    "Resultados 1T",
    "Resultados 2T",
]

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    # quitar acentos
    s = "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )
    # espacios múltiples -> uno
    while "  " in s:
        s = s.replace("  ", " ")
    return s

# mapear columnas por nombre normalizado
def build_header_map(cols):
    norm_map = {normalize(c): c for c in cols}
    return norm_map

def find_cols_for_required(df: pd.DataFrame):
    norm_map = build_header_map(df.columns)
    found = {}
    missing = []
    req_norm = [normalize(h) for h in REQ_HEADERS]
    for req, req_n in zip(REQ_HEADERS, req_norm):
        if req_n in norm_map:
            found[req] = norm_map[req_n]
        else:
            # tolerar variantes frecuentes
            variants = {
                "descripcion del indicador": ["descripcion del indicador", "descripción del indicador", "descripcion indicador"],
                "lider estrategico": ["lider estrategico", "líder estrategico", "lider estrategico(a)", "jefe estrategico"],
            }
            if req_n in variants:
                chosen = None
                for v in variants[req_n]:
                    if v in norm_map:
                        chosen = norm_map[v]
                        break
                if chosen:
                    found[req] = chosen
                else:
                    missing.append(req)
            else:
                missing.append(req)
    return found, missing

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracción")
    return buf.getvalue()

# === UI ===
file = st.file_uploader("Sube tu archivo Excel (.xlsx o .xlsm)", type=["xlsx", "xlsm"])
if file:
    # Descubrir hojas
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
        sheet = st.selectbox("Elige la hoja de trabajo", xls.sheet_names, index=0)
        header_row_1_based = st.number_input("Fila de encabezados (1 = primera fila)", min_value=1, value=1, step=1)
        header_row = header_row_1_based - 1

        # Leer con la fila de encabezados seleccionada
        df = pd.read_excel(xls, sheet_name=sheet, header=header_row, engine="openpyxl")

        st.write("**Vista previa (primeras 30 filas):**")
        st.dataframe(df.head(30), use_container_width=True)

        # Encontrar columnas requeridas
        col_map, missing = find_cols_for_required(df)

        # Mostrar mapeo
        with st.expander("Ver mapeo de encabezados detectados"):
            st.json(col_map)
            if missing:
                st.warning("No se encontraron estos encabezados: " + ", ".join(missing))

        # Permitir corrección manual si falta algo
        if missing:
            st.info("Puedes asignar manualmente columnas para los encabezados faltantes:")
            for req in missing:
                choice = st.selectbox(
                    f"Selecciona la columna para «{req}»",
                    options=["(no asignar)"] + list(df.columns),
                    key=f"fix_{req}"
                )
                if choice != "(no asignar)":
                    col_map[req] = choice
            # recalcular faltantes
            missing = [r for r in REQ_HEADERS if r not in col_map]

        # Si ya tenemos al menos una columna, permitir extraer
        if len(col_map) > 0:
            # Construir DataFrame de salida en el orden solicitado
            ordered_cols = []
            for req in REQ_HEADERS:
                if req in col_map:
                    ordered_cols.append(col_map[req])
                else:
                    # si falta, crear columna vacía
                    df[req] = ""
                    ordered_cols.append(req)

            out_df = df[ordered_cols].copy()
            # Renombrar a los nombres "oficiales"
            rename_map = {old: new for new, old in col_map.items()}
            out_df.rename(columns=rename_map, inplace=True)

            st.success("Extracción generada.")
            st.dataframe(out_df.head(50), use_container_width=True)

            # Descargas
            csv_bytes = out_df.to_csv(index=False).encode("utf-8-sig")
            xlsx_bytes = to_excel_bytes(out_df)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "⬇️ Descargar CSV",
                    data=csv_bytes,
                    file_name="extraccion_sembremos.csv",
                    mime="text/csv"
                )
            with col2:
                st.download_button(
                    "⬇️ Descargar Excel",
                    data=xlsx_bytes,
                    file_name="extraccion_sembremos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if missing:
                st.warning("Aún faltan columnas sin asignar: " + ", ".join(missing))
        else:
            st.error("No se detectó ningún encabezado requerido. Revisa la fila de encabezados o asigna manualmente.")
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
else:
    st.info("Sube tu archivo para comenzar.")
