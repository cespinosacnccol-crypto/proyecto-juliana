import io, os, sys
import streamlit as st
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

st.set_page_config(page_title="Proyecto Juliana", layout="wide")
st.title("Proyecto Juliana")
st.caption("Sistema de seguimiento académico")

RUTA_ACUM = os.path.join(os.path.dirname(__file__), "INFORME CLIENTE", "ACUMULADO GENERAL.xlsx")

# ─── LECTURA DE DATOS ────────────────────────────────────────────
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
            meta["CORRECTAS"] = correctas
            meta["TOTAL"] = total
            meta["%"] = round(correctas / total * 100, 1) if total else 0
            rows.append(meta)
        if rows:
            dfs.append(pd.DataFrame(rows))
    wb.close()
    return pd.concat(dfs, ignore_index=True) if dfs else None

# ─── PROCESAMIENTO ───────────────────────────────────────────────
def procesar_datos(df):
    if df is None or df.empty:
        return None, None
    # Agrupar por estudiante para detectar incompletos
    id_cols = ["NOMBRE SEDE", "CÓDIGO DANE SEDE", "NOMBRES ESTUDIANTE", "CÓD. EST.", "GRADO", "CURSO"]
    materias = df.groupby(id_cols)["PRUEBA"].apply(set).reset_index()
    materias["FALTA"] = materias["PRUEBA"].apply(
        lambda s: ", ".join(sorted({"MATEMÁTICAS", "LENGUAJE"} - s)) or ""
    )
    materias["COMPLETO"] = materias["FALTA"] == ""
    # Detalle por curso
    curso_cols = ["NOMBRE SEDE", "CÓDIGO DANE SEDE", "CURSO", "GRADO"]
    cursos = df.groupby(curso_cols).agg(
        estudiantes=("NOMBRES ESTUDIANTE", "nunique"),
        completo=("NOMBRES ESTUDIANTE", lambda x: 0),
    ).reset_index()
    # Completitud por curso
    compl = materias.groupby(curso_cols).agg(
        completos=("COMPLETO", "sum"),
        incompletos=("COMPLETO", lambda x: (~x).sum()),
    ).reset_index()
    cursos = cursos.merge(compl, on=curso_cols, how="left")
    cursos["completos"] = cursos["completos"].fillna(0).astype(int)
    cursos["incompletos"] = cursos["incompletos"].fillna(0).astype(int)
    # Resumen por colegio
    colegios = cursos.groupby(["NOMBRE SEDE", "CÓDIGO DANE SEDE"]).agg(
        cursos=("CURSO", "count"),
        estudiantes=("estudiantes", "sum"),
        completos=("completos", "sum"),
        incompletos=("incompletos", "sum"),
    ).reset_index()
    return {
        "df": df,
        "materias": materias,
        "cursos": cursos,
        "colegios": colegios,
    }

# ─── INICIO ──────────────────────────────────────────────────────
df = leer_acumulado()
datos = procesar_datos(df)

if datos is None:
    st.info("No hay datos disponibles. Vuelve pronto.")
    st.stop()

DF = datos["df"]
MAT = datos["materias"]
CURSOS = datos["cursos"]
COL = datos["colegios"]

# ─── NAVEGACIÓN ──────────────────────────────────────────────────
total_col = len(COL)
total_est = int(COL["estudiantes"].sum())
total_inc = int(COL["incompletos"].sum())
pct_completo = round((total_est - total_inc) / total_est * 100, 1) if total_est else 0

nivel = st.sidebar.radio("Navegación", ["Vista General", "Por Colegio", "Por Curso", "Detalle Completo"])
st.sidebar.markdown("---")
st.sidebar.metric("🏫 Colegios", f"{total_col}")
st.sidebar.metric("👨‍🎓 Estudiantes", f"{total_est}")
if total_inc > 0:
    st.sidebar.error(f"⚠️ {total_inc} estudiantes con pruebas incompletas")
st.sidebar.metric("✅ Completitud", f"{pct_completo}%")

# ══════════════════════════════════════════════════════════════════
# VISTA GENERAL
# ══════════════════════════════════════════════════════════════════
if nivel == "Vista General":
    st.subheader("📊 Vista General")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🏫 Colegios", total_col)
    m2.metric("👨‍🎓 Estudiantes", total_est)
    m3.metric("✅ Completos", total_est - total_inc)
    m4.metric("⚠️ Incompletos", total_inc)

    st.divider()
    st.subheader("Avance por colegio")
    tabla_col = COL.copy()
    tabla_col["% Completo"] = ((tabla_col["completos"] / tabla_col["estudiantes"]) * 100).round(1)
    tabla_col.columns = ["Colegio", "Código DANE", "Cursos", "Estudiantes", "Completos", "Incompletos", "% Completo"]
    st.dataframe(tabla_col, use_container_width=True, hide_index=True)

    st.bar_chart(tabla_col.set_index("Colegio")["Estudiantes"])

    if total_inc > 0:
        st.divider()
        st.error(f"⚠️ **{total_inc} estudiantes** con pruebas incompletas")
        inc = MAT[~MAT["COMPLETO"]][["NOMBRE SEDE", "NOMBRES ESTUDIANTE", "CURSO", "GRADO", "FALTA"]]
        inc.columns = ["Colegio", "Estudiante", "Curso", "Grado", "Materia(s) faltante(s)"]
        st.dataframe(inc, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# POR COLEGIO
# ══════════════════════════════════════════════════════════════════
elif nivel == "Por Colegio":
    st.subheader("🏫 Vista por Colegio")
    cole = st.selectbox("Selecciona un colegio", COL["NOMBRE SEDE"].unique())
    info = COL[COL["NOMBRE SEDE"] == cole].iloc[0]
    st.metric("Código DANE", info["CÓDIGO DANE SEDE"])

    cursos_col = CURSOS[CURSOS["NOMBRE SEDE"] == cole].copy()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cursos", len(cursos_col))
    m2.metric("Estudiantes", int(cursos_col["estudiantes"].sum()))
    m3.metric("✅ Completos", int(cursos_col["completos"].sum()))
    m4.metric("⚠️ Incompletos", int(cursos_col["incompletos"].sum()))

    st.divider()
    st.subheader("Cursos")
    cc = cursos_col[["CURSO", "GRADO", "estudiantes", "completos", "incompletos"]].copy()
    cc.columns = ["Curso", "Grado", "Estudiantes", "Completos", "Incompletos"]
    st.dataframe(cc, use_container_width=True, hide_index=True)

    inc_col = MAT[(MAT["NOMBRE SEDE"] == cole) & (~MAT["COMPLETO"])]
    if not inc_col.empty:
        st.divider()
        st.error(f"⚠️ Estudiantes con pruebas incompletas en {cole}")
        ic = inc_col[["NOMBRES ESTUDIANTE", "CURSO", "GRADO", "FALTA"]]
        ic.columns = ["Estudiante", "Curso", "Grado", "Materia(s) faltante(s)"]
        st.dataframe(ic, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# POR CURSO
# ══════════════════════════════════════════════════════════════════
elif nivel == "Por Curso":
    st.subheader("📚 Vista por Curso")
    cole = st.selectbox("Colegio", COL["NOMBRE SEDE"].unique(), key="sel_cole_curso")
    cursos_cole = CURSOS[CURSOS["NOMBRE SEDE"] == cole]
    curso_opts = cursos_cole["CURSO"].unique()
    curso_sel = st.selectbox("Curso", sorted(curso_opts))
    grado = cursos_cole[cursos_cole["CURSO"] == curso_sel]["GRADO"].iloc[0]

    filtro = DF[(DF["NOMBRE SEDE"] == cole) & (DF["CURSO"] == curso_sel)]

    m1, m2 = st.columns(2)
    m1.metric("Grado", grado)
    m2.metric("Estudiantes", filtro["NOMBRES ESTUDIANTE"].nunique())

    for materia in ["MATEMÁTICAS", "LENGUAJE"]:
        sub = filtro[filtro["PRUEBA"] == materia]
        st.divider()
        st.subheader(f"📖 {materia}")
        if sub.empty:
            st.warning(f"Ningún estudiante presentó {materia}")
            continue
        prom = sub["%"].mean()
        max_v = sub["%"].max()
        min_v = sub["%"].min()
        c1, c2, c3 = st.columns(3)
        c1.metric("Promedio", f"{prom:.1f}%")
        c2.metric("Máximo", f"{max_v:.1f}%")
        c3.metric("Mínimo", f"{min_v:.1f}%")

        col_bar, col_pie = st.columns([2, 1])
        with col_bar:
            sub_chart = sub[["NOMBRES ESTUDIANTE", "%"]].set_index("NOMBRES ESTUDIANTE")
            st.bar_chart(sub_chart)
        with col_pie:
            rangos = pd.cut(sub["%"], bins=[0, 25, 50, 75, 100], labels=["0-25%", "25-50%", "50-75%", "75-100%"])
            pie_data = rangos.value_counts().reset_index()
            pie_data.columns = ["Rango", "Cantidad"]
            st.dataframe(pie_data, use_container_width=True, hide_index=True)

    # Faltantes en este curso
    inc_curso = MAT[(MAT["NOMBRE SEDE"] == cole) & (MAT["CURSO"] == curso_sel) & (~MAT["COMPLETO"])]
    if not inc_curso.empty:
        st.divider()
        st.error(f"⚠️ Estudiantes con pruebas incompletas")
        ic = inc_curso[["NOMBRES ESTUDIANTE", "FALTA"]]
        ic.columns = ["Estudiante", "Materia(s) faltante(s)"]
        st.dataframe(ic, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# DETALLE COMPLETO
# ══════════════════════════════════════════════════════════════════
elif nivel == "Detalle Completo":
    st.subheader("📋 Detalle completo de datos")
    st.dataframe(DF, use_container_width=True, hide_index=True)
    if os.path.exists(RUTA_ACUM):
        with open(RUTA_ACUM, "rb") as f:
            st.download_button("⬇️ Descargar ACUMULADO GENERAL.xlsx", f.read(), "ACUMULADO GENERAL.xlsx")
