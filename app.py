# app.py
import io
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extractor de Indicadores", layout="wide")
st.title("Extractor Sembremos – Columnas específicas")
st.caption("Delegación, Líder Estratégico, Línea de Acción, Indicador, Descripción del Indicador, Meta, Resultado")

# === Columnas requeridas (en ese orden) ===
REQ_HEADERS = [
    "Delegación",
    "Líder Estratégico",
    "Línea de Acción",
    "Indicador",
    "Descripción del Indicador",
    "Meta",
    "Resultado",
]

# Variantes aceptadas por cada encabezado
VARIANTS = {
    "delegación": ["delegación", "delegacion", "delegacion policial"],
    "líder estratégico": ["líder estratégico", "lider estrategico", "jefe estrategico", "líder estrategico"],
    "línea de acción": ["línea de acción", "linea de accion", "linea de acción", "líneas de acción", "lineas de accion"],
    "indicador": ["indicador", "nombre del indicador"],
    "descripción del indicador": ["descripción del indicador", "descripcion del indicador", "descripcion indicador", "detalle del indicador"],
    "meta": ["meta", "meta anual", "meta trimestral"],
    "resultado": ["resultado", "resultados", "resultado 1t", "resultado 2t", "resultados 1t", "resultados 2t"],
}

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    while "  " in s:
        s = s.replace("  ", " ")
    return s

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracción")
    return buf.getvalue()

def auto_map_columns(df: pd.DataFrame):
    # mapa normalizado -> real
    norm_map = {normalize(c): c for c in df.columns}
    found = {}
    missing = []
    for req in REQ_HEADERS:
        req_n = normalize(req)
        # Primero, coincidencia exacta
        if req_n in norm_map:
            found[req] = norm_map[req_n]
            continue
        # Luego, probar variantes
        variants = VARIANTS.get(req_n, [])
        chosen = None
        for v in variants:
            v_n = normalize(v)
            if v_n in norm_map:
                chosen = norm_map[v_n]
                break
        if chosen:
            found[req] = chosen
        else:
            missing.append(req)
    return found, missing

file = st.file_uploader("Sube tu Excel (.xlsx o .xlsm)", type=["xlsx", "xlsm"])
if not file:
    st.info("Sube el archivo para comenzar.")
    st.stop()

try:
    xls = pd.ExcelFile(file, engine="openpyxl")
except Exception as e:
    st.error(f"No se pudo abrir el archivo: {e}")
    st.stop()

sheet = st.selectbox("Elige la hoja", xls.sheet_names, index=0)
header_row_1 = st.number_input("Fila de encabezados (1 = primera fila)", min_value=1, value=1)
header_row = header_row_1 - 1

df = pd.read_excel(xls, sheet_name=sheet, header=header_row, engine="openpyxl")
st.write("**Vista previa (primeras 30 filas):**")
st.dataframe(df.head(30), use_container_width=True)

col_map, missing = auto_map_columns(df)

with st.expander("Mapeo detectado"):
    st.json(col_map)
    if missing:
        st.warning("Faltan: " + ", ".join(missing))

# Corrección manual de faltantes
if missing:
    st.info("Asigna manualmente columnas para los faltantes:")
    for req in missing:
        choice = st.selectbox(
            f"Columna para «{req}»",
            options=["(no asignar)"] + list(df.columns),
            key=f"fix_{req}"
        )
        if choice != "(no asignar)":
            col_map[req] = choice
    missing = [r for r in REQ_HEADERS if r not in col_map]

# Construir salida en el orden requerido
ordered_cols = []
for req in REQ_HEADERS:
    if req in col_map:
        ordered_cols.append(col_map[req])
    else:
        df[req] = ""
        ordered_cols.append(req)

out_df = df[ordered_cols].copy()

# Renombrar a los nombres "oficiales"
rename_map = {v: k for k, v in col_map.items()}
out_df.rename(columns=rename_map, inplace=True)

# Si “Resultado” vino de “Resultados 1T/2T/etc.” lo renombra igualmente a “Resultado”
for c in out_df.columns:
    if normalize(c) in [normalize(v) for v in VARIANTS["resultado"]]:
        out_df.rename(columns={c: "Resultado"}, inplace=True)

st.success("Extracción generada.")
st.dataframe(out_df.head(50), use_container_width=True)

# Descargas
csv_bytes = out_df.to_csv(index=False).encode("utf-8-sig")
xlsx_bytes = to_excel_bytes(out_df)

c1, c2 = st.columns(2)
with c1:
    st.download_button("⬇️ Descargar CSV", data=csv_bytes, file_name="extraccion_sembremos.csv", mime="text/csv")
with c2:
    st.download_button("⬇️ Descargar Excel", data=xlsx_bytes,
                       file_name="extraccion_sembremos.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
