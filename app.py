import io, os, sys
import streamlit as st
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

st.set_page_config(page_title="Proyecto Juliana — Informe Cliente", layout="wide")
st.title("Proyecto Juliana")
st.caption("Informe de resultados — actualizado periódicamente")

RUTA_ACUM = os.path.join(os.path.dirname(__file__), "INFORME CLIENTE", "ACUMULADO GENERAL.xlsx")

@st.cache_data
def leer_acumulado():
    if not os.path.exists(RUTA_ACUM):
        return None
    wb = openpyxl.load_workbook(RUTA_ACUM)
    dfs = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        headers = [str(c.value) if c.value else f"COL{i}" for i, c in enumerate(ws[1])]
        eval_cols = [i for i, h in enumerate(headers) if h.endswith("_EVAL")]
        meta_nombres = ["NOMBRE SEDE", "CÓDIGO DANE SEDE", "NOMBRES ESTUDIANTE",
                         "CÓD. EST.", "GRADO", "CURSO", "PRUEBA"]
        meta_idx = [(i, headers.index(c)) for i, c in enumerate(meta_nombres) if c in headers]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(c is None or str(c).strip() == "" for c in row):
                continue
            correctas = sum(1 for i in eval_cols if i < len(row) and str(row[i]).strip().upper() == "CORRECTO")
            total = len(eval_cols)
            meta = {}
            for mi, col_i in meta_idx:
                v = row[col_i] if col_i < len(row) and row[col_i] is not None else ""
                meta[meta_nombres[mi]] = str(v).strip()
            meta["_correctas"] = correctas
            meta["_total"] = total
            meta["_porcentaje"] = round(correctas / total * 100, 1) if total else 0
            rows.append(meta)
        if rows:
            dfs.append(pd.DataFrame(rows))
    wb.close()
    return pd.concat(dfs, ignore_index=True) if dfs else None

def build_dashboard(df):
    if df is None or df.empty:
        return None
    cols = ["NOMBRE SEDE", "CÓDIGO DANE SEDE", "CURSO", "GRADO", "PRUEBA", "NOMBRES ESTUDIANTE", "_correctas", "_total"]
    if any(c not in df.columns for c in cols):
        return None
    grupo = df.groupby(["NOMBRE SEDE", "CÓDIGO DANE SEDE", "CURSO", "GRADO", "PRUEBA"], dropna=False)
    resumen = grupo.agg(
        estudiantes=("NOMBRES ESTUDIANTE", "count"),
        pct_promedio=("_porcentaje", "mean"),
    ).reset_index()
    resumen["pct_promedio"] = resumen["pct_promedio"].round(1)
    return resumen

df = leer_acumulado()
dash = build_dashboard(df) if df is not None else None

t1, t2 = st.tabs(["📊 Dashboard", "📋 Detalle"])

with t1:
    if dash is not None and not dash.empty:
        total_global = int(dash["estudiantes"].sum())
        st.success(f"📌 {total_global} estudiantes en total")
        for colegio in dash["NOMBRE SEDE"].unique():
            dc = dash[dash["NOMBRE SEDE"] == colegio]
            dane = dc["CÓDIGO DANE SEDE"].iloc[0]
            total_est = int(dc["estudiantes"].sum())
            st.divider()
            cols = st.columns([2, 1, 1])
            cols[0].markdown(f"### 🏫 {colegio}")
            cols[1].metric("Código DANE", str(dane) if dane else "—")
            cols[2].metric("Estudiantes", f"{total_est}")
            tabla = dc[["CURSO", "GRADO", "PRUEBA", "estudiantes", "pct_promedio"]].copy()
            tabla.columns = ["Curso", "Grado", "Materia", "Estudiantes", "% Acierto Prom."]
            st.dataframe(tabla, use_container_width=True, hide_index=True)
            st.bar_chart(tabla.set_index("Curso")["Estudiantes"])
    else:
        st.info("No hay datos disponibles. Vuelve pronto.")

with t2:
    if df is not None and not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos disponibles.")

# Descargas
if os.path.exists(RUTA_ACUM):
    st.divider()
    with open(RUTA_ACUM, "rb") as f:
        st.download_button("⬇️ Descargar ACUMULADO GENERAL", f.read(), "ACUMULADO GENERAL.xlsx")
