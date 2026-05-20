import io, os, sys
import streamlit as st
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from calificar_pruebas import (
    norm, sanitizar_sheet, estilo_header, ancho_columnas, colorear_eval,
    cargar_respuestas, estandarizar_columnas,
    COLUMNAS_ESTANDAR, MATERIA_NOMBRES, ST_CELL_FONT, ST_CELL_ALIGN, ST_BORDER,
    ACUMULADO,
)

st.set_page_config(page_title="Proyecto Juliana", layout="wide")

# ── Session state ──
if "acumulado_wb" not in st.session_state:
    st.session_state.acumulado_wb = None
if "resultados" not in st.session_state:
    st.session_state.resultados = []
if "procesado" not in st.session_state:
    st.session_state.procesado = False

# ── Cargar respuestas correctas ──
RUTA_RESPS = os.path.join(os.path.dirname(__file__), "RESPUESTAS CORRECTAS PROYECTO JULIANA.xlsx")
if not os.path.exists(RUTA_RESPS):
    st.error("No se encuentra el archivo de respuestas correctas en el servidor.")
    st.stop()
resp_correctas = cargar_respuestas()
from calificar_pruebas import crear_carpetas
crear_carpetas()


# ── Helper: leer acumulado a DataFrame unificado ──
@st.cache_data
def leer_acumulado():
    if not os.path.exists(ACUMULADO):
        return None
    wb = openpyxl.load_workbook(ACUMULADO)
    dfs = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        data = list(ws.iter_rows(values_only=True))
        if len(data) < 2:
            continue
        headers = [str(c) if c else f"COL{i}" for i, c in enumerate(data[0])]
        rows = [r for r in data[1:] if not all(c is None or str(c).strip() == "" for c in r)]
        if not rows:
            continue
        df = pd.DataFrame(rows, columns=headers)
        df["HOJA"] = sheet
        dfs.append(df)
    wb.close()
    return pd.concat(dfs, ignore_index=True) if dfs else None


# ── Helper: generar INFORME CLIENTE en memoria ──
def generar_informe_bytes(wb_acum_path):
    if not os.path.exists(wb_acum_path):
        return None
    wb = openpyxl.load_workbook(wb_acum_path)
    for ws in wb.worksheets:
        for col in range(8, ws.max_column + 1):
            h = str(ws.cell(1, col).value or "")
            if h.startswith("P") and h[1:].rstrip().isdigit():
                ws.column_dimensions[get_column_letter(col)].hidden = True
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wb.close()
    return buf


# ── Helper: construir dashboard ──
def build_dashboard(df):
    if df is None or df.empty:
        return None
    cols_needed = ["NOMBRE SEDE", "CÓDIGO DANE SEDE", "CURSO", "GRADO", "PRUEBA",
                    "NOMBRES ESTUDIANTE", "PORCENTAJE ACIERTO"]
    missing = [c for c in cols_needed if c not in df.columns]
    if missing:
        return None
    grupo = df.groupby(["NOMBRE SEDE", "CÓDIGO DANE SEDE", "CURSO", "GRADO", "PRUEBA"], dropna=False)
    resumen = grupo.agg(
        estudiantes=("NOMBRES ESTUDIANTE", "count"),
        pct_promedio=("PORCENTAJE ACIERTO", lambda x: round(x.astype(float).mean() * 100, 1)),
    ).reset_index()
    return resumen


# ── Title ──
st.title("Proyecto Juliana")
st.caption("Sistema de calificación y consulta de resultados")

# ── Leer acumulado del disco ──
acum_df = leer_acumulado()
dashboard = build_dashboard(acum_df) if acum_df is not None else None

# ── Tabs ──
t1, t2, t3 = st.tabs(["📊 Dashboard", "📋 Acumulado General", "📤 Cargar Pruebas"])

# ═══════════════════════════════════════════════════════════════════
# TAB 1: DASHBOARD
# ═══════════════════════════════════════════════════════════════════
with t1:
    if dashboard is not None and not dashboard.empty:
        colegios = dashboard["NOMBRE SEDE"].unique()
        for colegio in colegios:
            dc = dashboard[dashboard["NOMBRE SEDE"] == colegio]
            dane = dc["CÓDIGO DANE SEDE"].iloc[0]
            total_est = int(dc["estudiantes"].sum())

            st.divider()
            cols = st.columns([2, 1, 1, 1])
            cols[0].markdown(f"### 🏫 {colegio}")
            cols[1].metric("Código DANE", str(dane) if dane else "—")
            cols[2].metric("Total estudiantes", f"{total_est}")
            cols[3].metric("Cursos", f"{len(dc)}")

            tabla = dc[["CURSO", "GRADO", "PRUEBA", "estudiantes", "pct_promedio"]].copy()
            tabla.columns = ["Curso", "Grado", "Materia", "Estudiantes", "% Acierto Prom."]
            st.dataframe(tabla, use_container_width=True, hide_index=True)
            # Bar chart
            st.bar_chart(tabla.set_index("Curso")["Estudiantes"])
    else:
        st.info("No hay datos acumulados. Sube pruebas en la pestaña **Cargar Pruebas**.")

# ═══════════════════════════════════════════════════════════════════
# TAB 2: ACUMULADO GENERAL + INFORME CLIENTE
# ═══════════════════════════════════════════════════════════════════
with t2:
    col_acum, col_inf = st.columns(2)

    if acum_df is not None and not acum_df.empty:
        st.dataframe(acum_df, use_container_width=True, hide_index=True)

        with col_acum:
            if os.path.exists(ACUMULADO):
                with open(ACUMULADO, "rb") as f:
                    st.download_button(
                        "⬇️ Descargar ACUMULADO GENERAL",
                        f.read(),
                        "ACUMULADO GENERAL.xlsx",
                        use_container_width=True,
                    )

        with col_inf:
            buf = generar_informe_bytes(ACUMULADO)
            if buf:
                st.download_button(
                    "⬇️ Descargar INFORME CLIENTE",
                    buf,
                    "INFORME CLIENTE.xlsx",
                    use_container_width=True,
                )
    else:
        st.info("No hay datos acumulados.")

# ═══════════════════════════════════════════════════════════════════
# TAB 3: CARGAR PRUEBAS
# ═══════════════════════════════════════════════════════════════════
with t3:
    sube_acum = st.file_uploader(
        "Sube ACUMULADO GENERAL.xlsx existente (opcional — si no, empieza desde cero)",
        type=["xlsx"],
        key="acum_upload",
    )
    archivos = st.file_uploader(
        "Selecciona archivos .xlsx de pruebas a calificar",
        type=["xlsx"],
        accept_multiple_files=True,
        key="test_upload",
    )

    if archivos and st.button("🚀 Calificar", type="primary", use_container_width=True):
        st.session_state.resultados = []

        # Cargar acumulado base
        wb_acum = None
        if sube_acum:
            try:
                wb_acum = openpyxl.load_workbook(io.BytesIO(sube_acum.read()))
            except Exception:
                pass
        if wb_acum is None:
            wb_acum = openpyxl.Workbook()
            if "Sheet" in wb_acum.sheetnames:
                del wb_acum["Sheet"]

        bar = st.progress(0, "Procesando...")

        for idx, archivo in enumerate(archivos):
            bar.progress((idx) / len(archivos), f"{archivo.name}...")
            try:
                wb_in = openpyxl.load_workbook(io.BytesIO(archivo.read()))
                ws_in = wb_in.active
            except Exception as e:
                st.error(f"{archivo.name}: Error al abrir ({e})")
                continue

            wb_std = estandarizar_columnas(ws_in)
            ws = wb_std.active
            headers = [str(c.value) if c.value is not None else "" for c in ws[1]]
            n_meta = len(COLUMNAS_ESTANDAR)
            e_resp = list(range(n_meta, len(headers)))

            alumnos = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if all(c is None or str(c).strip() == "" for c in row):
                    continue
                cod_dane = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                sede = str(row[1]).strip() if len(row) > 1 and row[1] else archivo.name
                nombre = str(row[2]).strip() if len(row) > 2 and row[2] else "SIN NOMBRE"
                cod = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                grado = None
                if len(row) > 4:
                    try:
                        grado = int(float(str(row[4]).strip()))
                    except: pass
                grupo = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                materia_raw = ""
                if len(row) > 6 and row[6]:
                    materia_raw = norm(str(row[6]))
                    if materia_raw == "lectura":
                        materia_raw = "lenguaje"

                if grado is None or not materia_raw:
                    nb = os.path.splitext(archivo.name)[0].replace(" ", "_")
                    for p in nb.split("_"):
                        pu = norm(p).upper()
                        if grado is None and pu in ("3","5","7","9","11"):
                            grado = int(pu)
                        if not materia_raw:
                            if pu in ("MATEMATICAS","MATE","MAT"):
                                materia_raw = "matematicas"
                            elif pu in ("LENGUAJE","LENGUA","LEN"):
                                materia_raw = "lenguaje"

                if grado is None or not materia_raw:
                    continue
                key = (grado, materia_raw)
                if key not in resp_correctas:
                    continue

                cdict = resp_correctas[key]
                total = len(cdict)
                resp_est = []
                for c in e_resp:
                    v = row[c] if c < len(row) else None
                    resp_est.append(str(v).strip().upper() if v is not None else "")

                detalles = []
                ok = 0
                for i in range(total):
                    re = resp_est[i] if i < len(resp_est) else ""
                    rc = cdict.get(i + 1, "")
                    okk = (re == rc and re != "")
                    if okk:
                        ok += 1
                    detalles.append({"resp": re, "correcta": rc, "ok": okk})

                alumnos.append({
                    "cod_dane": cod_dane, "sede": sede, "estudiante": nombre,
                    "cod": cod, "grado": grado, "grupo": grupo,
                    "materia": materia_raw, "detalles": detalles,
                    "total": total, "ok": ok, "bad": total - ok,
                    "pct": round(ok / total, 4) if total else 0,
                })

            if not alumnos:
                st.warning(f"{archivo.name}: No se encontraron estudiantes válidos")
                continue

            total_preg = alumnos[0]["total"]

            # Generar workbook calificado
            wb_out = openpyxl.Workbook()
            ws_out = wb_out.active
            ws_out.title = "Evaluaciones"
            cols_meta = ["CÓDIGO DANE SEDE", "NOMBRE SEDE", "NOMBRES ESTUDIANTE",
                          "CÓD. EST.", "GRADO", "CURSO", "PRUEBA"]
            cols_preg = []
            for q in range(1, total_preg + 1):
                cols_preg.append(f"P{q:02d}")
                cols_preg.append(f"P{q:02d}_EVAL")
            cols_sum = ["CORRECTAS", "INCORRECTAS", "TOTAL_EVAL", "PORCENTAJE ACIERTO"]
            todos = cols_meta + cols_preg + cols_sum

            for c, h in enumerate(todos, 1):
                ws_out.cell(1, c, h)
            estilo_header(ws_out, len(todos))
            ancho_columnas(ws_out, todos)

            col_ini_eval = len(cols_meta) + 2
            col_fin_eval = len(cols_meta) + total_preg * 2
            letra_ini = get_column_letter(col_ini_eval)
            letra_fin = get_column_letter(col_fin_eval)
            col_ci = todos.index("CORRECTAS") + 1
            col_ii = todos.index("INCORRECTAS") + 1
            col_ti = todos.index("TOTAL_EVAL") + 1
            col_pi = todos.index("PORCENTAJE ACIERTO") + 1

            for i, al in enumerate(alumnos):
                f = 2 + i
                ws_out.cell(f, 1, al["cod_dane"]).font = ST_CELL_FONT
                ws_out.cell(f, 2, al["sede"].upper()).font = ST_CELL_FONT
                ws_out.cell(f, 3, al["estudiante"].upper()).font = ST_CELL_FONT
                ws_out.cell(f, 4, al["cod"].upper()).font = ST_CELL_FONT
                ws_out.cell(f, 5, al["grado"]).font = ST_CELL_FONT
                gpo = int(al["grupo"]) if al["grupo"].isdigit() else al["grupo"].upper()
                ws_out.cell(f, 6, gpo).font = ST_CELL_FONT
                ws_out.cell(f, 7, MATERIA_NOMBRES.get(al["materia"], al["materia"].upper())).font = ST_CELL_FONT
                for c in range(1, 8):
                    ws_out.cell(f, c).alignment = ST_CELL_ALIGN
                    ws_out.cell(f, c).border = ST_BORDER
                for q, det in enumerate(al["detalles"]):
                    cr = len(cols_meta) + 1 + q * 2
                    ce = cr + 1
                    ws_out.cell(f, cr, det["resp"]).font = ST_CELL_FONT
                    ws_out.cell(f, cr).alignment = ST_CELL_ALIGN
                    ws_out.cell(f, cr).border = ST_BORDER
                    colorear_eval(ws_out.cell(f, ce), "CORRECTO" if det["ok"] else "INCORRECTO")
                ws_out.cell(f, col_ci, f'=COUNTIF({letra_ini}{f}:{letra_fin}{f},"CORRECTO")')
                ws_out.cell(f, col_ci).font = ST_CELL_FONT
                ws_out.cell(f, col_ci).alignment = ST_CELL_ALIGN
                ws_out.cell(f, col_ci).border = ST_BORDER
                ws_out.cell(f, col_ii, f'=COUNTIF({letra_ini}{f}:{letra_fin}{f},"INCORRECTO")')
                ws_out.cell(f, col_ii).font = ST_CELL_FONT
                ws_out.cell(f, col_ii).alignment = ST_CELL_ALIGN
                ws_out.cell(f, col_ii).border = ST_BORDER
                ws_out.cell(f, col_ti, total_preg)
                ws_out.cell(f, col_ti).font = ST_CELL_FONT
                ws_out.cell(f, col_ti).alignment = ST_CELL_ALIGN
                ws_out.cell(f, col_ti).border = ST_BORDER
                ws_out.cell(f, col_pi, f'={get_column_letter(col_ci)}{f}/{get_column_letter(col_ti)}{f}')
                ws_out.cell(f, col_pi).font = ST_CELL_FONT
                ws_out.cell(f, col_pi).alignment = ST_CELL_ALIGN
                ws_out.cell(f, col_pi).border = ST_BORDER
                ws_out.cell(f, col_pi).number_format = "0.00%"
            ws_out.freeze_panes = "A2"

            buf = io.BytesIO()
            wb_out.save(buf)
            buf.seek(0)

            # Acumular en el workbook en sesión
            from calificar_pruebas import actualizar_acumulado_wb
            wb_acum = actualizar_acumulado_wb(alumnos, wb_acum)

            st.session_state.resultados.append({
                "archivo": archivo.name,
                "descarga": buf,
            })

        # Guardar acumulado actualizado a disco
        try:
            wb_acum.save(ACUMULADO)
        except Exception:
            st.warning("No se pudo guardar el acumulado en disco.")

        bar.progress(1.0, "¡Listo!")
        st.session_state.procesado = True
        st.cache_data.clear()
        st.success("Calificación completada. Ve al Dashboard o Acumulado General para ver los resultados.")

    if st.session_state.resultados:
        st.divider()
        st.subheader("Descargas")
        for res in st.session_state.resultados:
            nb = res["archivo"].replace(".xlsx", "")
            st.download_button(
                label=f"⬇️ {nb}_CALIFICADO.xlsx",
                data=res["descarga"],
                file_name=f"{nb}_CALIFICADO.xlsx",
                use_container_width=True,
            )

        if os.path.exists(ACUMULADO):
            with open(ACUMULADO, "rb") as f:
                st.download_button(
                    "⬇️ Descargar ACUMULADO GENERAL actualizado",
                    f.read(),
                    "ACUMULADO GENERAL.xlsx",
                    use_container_width=True,
                )
