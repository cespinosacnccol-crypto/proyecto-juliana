"""
Genera archivos de ejemplo para probar el sistema de calificación Proyecto Juliana.
Cada archivo simula una prueba de un colegio con múltiples estudiantes.
"""
import openpyxl, os

BASE = os.path.dirname(os.path.abspath(__file__))
DEST = os.path.join(BASE, "Pruebas a calificar")


def crear_archivo_colegio(nombre_archivo, grado, materia, estudiantes):
    """
    Crea un archivo .xlsx con el formato esperado por el sistema.
    estudiantes = lista de (nombre, [respuestas...])
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados"

    num_p = len(estudiantes[0][1]) if estudiantes else 0
    headers = ["ESTUDIANTE", "GRADO", "MATERIA"] + [f"P{i+1}" for i in range(num_p)]
    ws.append(headers)

    for est in estudiantes:
        nombre, respuestas = est
        ws.append([nombre, grado, materia] + list(respuestas))

    ruta = os.path.join(DEST, nombre_archivo)
    wb.save(ruta)
    print(f"  + Creado: {nombre_archivo}  ({len(estudiantes)} estudiantes, {num_p} preguntas)")


def main():
    os.makedirs(DEST, exist_ok=True)

    # ── Colegio A - Grado 3 MATEMATICAS (15 preguntas) ───────────────────────
    crear_archivo_colegio(
        "Colegio_San_Jose_3M.xlsx", 3, "MATEMATICAS", [
            ("Ana López",      ["D","B","C","C","B","C","B","D","C","C","D","D","C","D","A"]),
            ("Carlos Ruiz",    ["D","B","A","C","B","C","B","D","C","C","D","D","A","D","B"]),
            ("María Torres",   ["D","B","C","A","B","C","A","D","C","C","D","D","C","D","C"]),
        ]
    )

    # ── Colegio A - Grado 3 LENGUAJE (15 preguntas) ──────────────────────────
    crear_archivo_colegio(
        "Colegio_San_Jose_3L.xlsx", 3, "LENGUAJE", [
            ("Pedro Gómez",    ["B","B","B","C","C","B","C","C","B","C","C","B","B","B","B"]),
            ("Laura Díaz",     ["B","B","A","C","C","B","C","C","B","C","C","B","B","B","A"]),
            ("Sofía Martínez", ["B","B","B","A","C","B","C","C","B","C","C","B","B","B","C"]),
        ]
    )

    # ── Colegio B - Grado 5 MATEMATICAS (20 preguntas) ───────────────────────
    crear_archivo_colegio(
        "Institucion_La_Cima_5M.xlsx", 5, "MATEMATICAS", [
            ("Diego Ramírez",  ["C","D","C","B","C","A","B","C","B","D","B","C","B","A","D","C","D","C","B","C"]),
            ("Valentina Ortiz",["C","D","A","B","C","A","B","C","B","D","B","C","B","A","D","C","D","A","B","C"]),
        ]
    )

    # ── Colegio B - Grado 5 LENGUAJE (20 preguntas) ──────────────────────────
    crear_archivo_colegio(
        "Institucion_La_Cima_5L.xlsx", 5, "LENGUAJE", [
            ("Camila Rojas",   ["B","C","A","D","B","B","C","C","B","C","B","B","C","B","A","B","C","A","D","B"]),
            ("Andrés Medina",  ["B","C","A","D","B","B","A","C","B","C","B","B","C","B","A","B","C","B","D","B"]),
        ]
    )

    # ── Colegio C - Grado 9 MATEMATICAS (20 preguntas) ───────────────────────
    crear_archivo_colegio(
        "Liceo_Moderna_9M.xlsx", 9, "MATEMATICAS", [
            ("Gabriela Silva", ["C","D","C","B","C","A","B","C","B","D","B","C","B","A","D","C","D","C","B","C"]),
            ("Felipe Castro",  ["C","D","C","B","C","A","B","C","B","D","B","C","B","A","D","C","D","C","B","A"]),
            ("Daniela Paz",    ["C","D","C","B","C","A","B","C","B","D","B","C","B","A","D","C","D","C","B","D"]),
        ]
    )

    print(f"\nTodos los archivos de ejemplo creados en: {DEST}")
    print("Ejecuta 'calificar_pruebas.py' para procesarlos.")


if __name__ == "__main__":
    main()
