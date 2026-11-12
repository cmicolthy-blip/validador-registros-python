"""Genera un Excel de ejemplo con errores tipicos de matricula.

Uso:
    python crear_datos_ejemplo.py

El archivo generado es ficticio. No contiene datos reales de estudiantes.
"""

from pathlib import Path

import pandas as pd

ARCHIVO_SALIDA = Path(__file__).with_name("registros_ejemplo.xlsx")

DATOS = [
    {
        "dni": "12345678",
        "nombres": "Ana",
        "apellidos": "Perez",
        "correo": "ana.perez@example.com",
        "carrera": "Administracion",
        "ciclo": "3",
        "fecha_matricula": "2026-07-15",
        "estado": "Matriculado",
    },
    {
        "dni": "87654321",
        "nombres": "Luis",
        "apellidos": "Garcia",
        "correo": "luis.garcia@example.com",
        "carrera": "Contabilidad",
        "ciclo": "5",
        "fecha_matricula": "2026-07-16",
        "estado": "Matriculado",
    },
    {
        "dni": "1234567",  # DNI con 7 digitos
        "nombres": "Maria",
        "apellidos": "Lopez",
        "correo": "maria.lopez@example.com",
        "carrera": "Ingenieria de Sistemas",
        "ciclo": "2",
        "fecha_matricula": "2026-07-16",
        "estado": "Matriculado",
    },
    {
        "dni": "45678912",
        "nombres": "Carlos",
        "apellidos": "Rojas",
        "correo": "",  # correo vacio
        "carrera": "Administracion",
        "ciclo": "4",
        "fecha_matricula": "2026-07-17",
        "estado": "Matriculado",
    },
    {
        "dni": "45678912",  # DNI duplicado
        "nombres": "Carlos",
        "apellidos": "Rojas",
        "correo": "carlos.rojas@example.com",
        "carrera": "Administracion",
        "ciclo": "4",
        "fecha_matricula": "2026-07-17",
        "estado": "Matriculado",
    },
    {
        "dni": "11223344",
        "nombres": "Sofia",
        "apellidos": "Mendoza",
        "correo": "sofia.mendoza@example",  # correo invalido
        "carrera": "Contabilidad",
        "ciclo": "1",
        "fecha_matricula": "2026-07-18",
        "estado": "Matriculado",
    },
    {
        "dni": "22334455",
        "nombres": "Jorge",
        "apellidos": "Castillo",
        "correo": "jorge.castillo@example.com",
        "carrera": "Ingenieria de Sistemas",
        "ciclo": "0",  # ciclo fuera de rango
        "fecha_matricula": "2026-07-18",
        "estado": "Matriculado",
    },
    {
        "dni": "33445566",
        "nombres": "Valeria",
        "apellidos": "Torres",
        "correo": "valeria.torres@example.com",
        "carrera": "Administracion",
        "ciclo": "6",
        "fecha_matricula": "31/02/2026",  # fecha invalida
        "estado": "Matriculado",
    },
    {
        "dni": "44556677",
        "nombres": "",  # nombre vacio
        "apellidos": "Ramirez",
        "correo": "ramirez@example.com",
        "carrera": "Contabilidad",
        "ciclo": "2",
        "fecha_matricula": "2026-07-19",
        "estado": "Matriculado",
    },
    {
        "dni": "55667788",
        "nombres": "Diego",
        "apellidos": "Flores",
        "correo": "diego.flores@example.com",
        "carrera": "Administracion",
        "ciclo": "3",
        "fecha_matricula": "2026-07-19",
        "estado": "",  # estado vacio
    },
]


def main() -> None:
    df = pd.DataFrame(DATOS)
    df.columns = [
        "DNI",
        "Nombres",
        "Apellidos",
        "Correo",
        "Carrera",
        "Ciclo",
        "Fecha de matricula",
        "Estado",
    ]
    df.to_excel(ARCHIVO_SALIDA, index=False)
    print(f"Archivo creado: {ARCHIVO_SALIDA}")
    print(f"Filas: {len(df)}")


if __name__ == "__main__":
    main()
