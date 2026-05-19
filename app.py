import io, os, sys, tempfile
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd

# ── Asegurar que encuentra calificar_pruebas.py ──────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from calificar_pruebas import (
    norm, sanitizar_sheet, estilo_header, ancho_columnas, colorear_eval,
    cargar_respuestas, estandarizar_columnas, actualizar_acumulado,
    COLUMNAS_ESTANDAR, MATERIA_NOMBRES, ST_CELL_FONT, ST_CELL_ALIGN, ST_BORDER,
    ANCHOS_META,
)

st.set_page_config(page_title="Calificadora Proyecto Juliana", layout="wide")
st.title("📊 Sistema de Calificación — Proyecto Juliana")

# ── Cargar respuestas correctas ──────────────────────────────────────────────
RUTA_RESPS = os.path.join(os.path.dirname(__file__), "RESPUESTAS CORRECTAS PROYECTO JULIANA.xlsx")
if not os.path.exists(RUTA_RESPS):
    st.error("No se encuentra el archivo de respuestas correctas en el servidor.")
    st.stop()

resp_correctas = cargar_respuestas()
from calificar_pruebas import crear_carpetas
crear_carpetas()
st.success(f"Respuestas cargadas: {len(resp_correctas)} combinaciones")
for (g, m), v in sorted(resp_correctas.items()):
    st.caption(f"Grado {g} / {m.upper()}: {len(v)} preguntas")

# ── Estado de sesión ─────────────────────────────────────────────────────────
if "acumulado_wb" not in st.session_state:
    st.session_state.acumulado_wb = None
if "resultados" not in st.session_state:
    st.session_state.resultados = []
if "procesado" not in st.session_state:
    st.session_state.procesado = False

# ── Subir archivos ───────────────────────────────────────────────────────────
archivos = st.file_uploader(
    "Selecciona uno o más archivos .xlsx de pruebas",
    type=["xlsx"], accept_multiple_files=True,
)

if archivos and st.button("🚀 Calificar", type="primary"):
    st.session_state.resultados = []
    wb_acum = st.session_state.acumulado_wb
    if wb_acum is None:
        wb_acum = openpyxl.Workbook()
        if "Sheet" in wb_acum.sheetnames:
            del wb_acum["Sheet"]
    else:
        wb_acum = openpyxl.load_workbook(wb_acum)

    bar = st.progress(0, "Procesando...")
    for idx, archivo in enumerate(archivos):
        bar.progress((idx) / len(archivos), f"{archivo.name}...")
        try:
            wb_in = openpyxl.load_workbook(io.BytesIO(archivo.read()))
            ws_in = wb_in.active
        except Exception as e:
            st.error(f"{archivo.name}: Error al abrir ({e})")
            continue

        # Estandarizar
        wb_std = estandarizar_columnas(ws_in)
        ws = wb_std.active
        headers = [str(c.value) if c.value is not None else "" for c in ws[1]]
        n_meta = len(COLUMNAS_ESTANDAR)

        e_cod_dane, e_nom_sede, e_estudiante, e_id = 0, 1, 2, 3
        e_grado, e_grupo, e_materia = 4, 5, 6
        e_resp = list(range(n_meta, len(headers)))

        alumnos = []
        for fila_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            if all(c is None or str(c).strip() == "" for c in row):
                continue
            cod_dane = str(row[e_cod_dane]).strip() if len(row) > e_cod_dane and row[e_cod_dane] else ""
            nombre = str(row[e_estudiante]).strip() if len(row) > e_estudiante and row[e_estudiante] else "SIN NOMBRE"
            sede = str(row[e_nom_sede]).strip() if len(row) > e_nom_sede and row[e_nom_sede] else archivo.name
            cod = str(row[e_id]).strip() if len(row) > e_id and row[e_id] else ""
            grupo = str(row[e_grupo]).strip() if len(row) > e_grupo and row[e_grupo] else ""
            grado = None
            if len(row) > e_grado:
                try: grado = int(float(str(row[e_grado]).strip()))
                except: pass
            materia_raw = ""
            if len(row) > e_materia and row[e_materia]:
                materia_raw = norm(str(row[e_materia]))

            # fallback nombre archivo
            if grado is None or not materia_raw:
                nb = os.path.splitext(archivo.name)[0].replace(" ", "_")
                for p in nb.split("_"):
                    pu = norm(p).upper()
                    if grado is None and pu in ("3","5","7","9","11"): grado = int(pu)
                    if not materia_raw:
                        if pu in ("MATEMATICAS","MATE","MAT"): materia_raw = "matematicas"
                        elif pu in ("LENGUAJE","LENGUA","LEN"): materia_raw = "lenguaje"

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
                okk = (re == rc)
                if okk: ok += 1
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

        # ── Generar workbook calificado en memoria ──
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = "Evaluaciones"
        cols_meta = ["CÓDIGO DANE SEDE", "NOMBRE SEDE", "NOMBRES ESTUDIANTE",
                      "ID", "GRADO", "CURSO", "PRUEBA"]
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
            ws_out.cell(f, 4, al["cod"]).font = ST_CELL_FONT
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

        # Acumular
        actualizar_acumulado(alumnos)

        # Recargar acumulado desde disco (actualizar_acumulado guarda en disco)
        ruta_acum = os.path.join(os.path.dirname(__file__), "Pruebas acumuladas", "ACUMULADO GENERAL.xlsx")
        if os.path.exists(ruta_acum):
            wb_acum = openpyxl.load_workbook(ruta_acum)
            st.session_state.acumulado_wb = ruta_acum

        # Tabla resumen para mostrar
        rows = []
        for al in alumnos:
            rows.append({
                "Estudiante": al["estudiante"].upper(),
                "Grado": al["grado"],
                "Materia": MATERIA_NOMBRES.get(al["materia"], al["materia"].upper()),
                "Correctas": al["ok"],
                "Total": al["total"],
                "%": f"{al['pct']*100:.1f}%",
            })
        st.session_state.resultados.append({
            "archivo": archivo.name,
            "tabla": rows,
            "descarga": buf,
        })

    bar.progress(1.0, "¡Listo!")
    st.session_state.procesado = True

# ── Mostrar resultados ───────────────────────────────────────────────────────
if st.session_state.resultados:
    st.divider()
    st.subheader("📋 Resultados")

    for res in st.session_state.resultados:
        with st.expander(f"**{res['archivo']}**", expanded=True):
            df = pd.DataFrame(res["tabla"])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.download_button(
                label="⬇️ Descargar Excel calificado",
                data=res["descarga"],
                file_name=res["archivo"].replace(".xlsx", "_CALIFICADO.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # Descargar acumulado
    ruta_acum = st.session_state.acumulado_wb
    if ruta_acum and os.path.exists(ruta_acum):
        with open(ruta_acum, "rb") as f:
            st.download_button(
                label="⬇️ Descargar ACUMULADO GENERAL",
                data=f.read(),
                file_name="ACUMULADO GENERAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # Generar INFORME CLIENTE
    if st.button("📁 Generar INFORME CLIENTE"):
        from calificar_pruebas import generar_informe_cliente, DIR_INFORME
        generar_informe_cliente()
        ruta_informe = os.path.join(DIR_INFORME, "ACUMULADO GENERAL.xlsx")
        if os.path.exists(ruta_informe):
            with open(ruta_informe, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar INFORME CLIENTE",
                    data=f.read(),
                    file_name="INFORME CLIENTE - ACUMULADO GENERAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.success("INFORME CLIENTE generado")

elif not archivos:
    st.info("Sube archivos .xlsx y presiona **Calificar** para empezar.")
