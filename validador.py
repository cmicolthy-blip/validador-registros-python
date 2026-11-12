"""Validador de registros de matricula.

Lee un Excel, aplica reglas de validacion y genera un reporte de observaciones.

Uso:
    python validador.py
    python validador.py --archivo otro.xlsx
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd
from dateutil import parser as date_parser

ARCHIVO_ENTRADA = Path(__file__).with_name("registros_ejemplo.xlsx")
ARCHIVO_REPORTE = Path(__file__).with_name("reporte_observaciones.xlsx")

# Mapeo: nombre interno -> columna real del Excel.
# Si tu Excel usa otros encabezados, cambia solo este diccionario.
COLUMNAS = {
    "dni": "DNI",
    "nombres": "Nombres",
    "apellidos": "Apellidos",
    "correo": "Correo",
    "carrera": "Carrera",
    "ciclo": "Ciclo",
    "fecha_matricula": "Fecha de matricula",
    "estado": "Estado",
}

ESTADOS_VALIDOS = {"matriculado", "reservado", "retirado", "pendiente"}

REGEX_DNI = re.compile(r"^\d{8}$")
REGEX_CORREO = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def cargar_datos(ruta: Path) -> pd.DataFrame:
    """Carga el Excel como texto para no perder ceros ni formatos."""
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontro el archivo: {ruta}")

    df = pd.read_excel(ruta, dtype=str).fillna("")
    df.columns = [str(columna).strip() for columna in df.columns]

    faltantes = [col for col in COLUMNAS.values() if col not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas en el Excel: "
            + ", ".join(faltantes)
            + ". Revisa el diccionario COLUMNAS."
        )

    return df


def es_fecha_valida(valor: str) -> bool:
    """Acepta formatos como 15/07/2026, 2026-07-15 o 15-07-2026."""
    try:
        date_parser.parse(valor, dayfirst=True)
        return True
    except (ValueError, OverflowError):
        return False


def agregar_observacion(
    observaciones: list[dict],
    fila: int,
    columna: str,
    valor: str,
    mensaje: str,
) -> None:
    observaciones.append(
        {
            "Fila": fila,
            "Columna": columna,
            "Valor": valor,
            "Observacion": mensaje,
        }
    )


def validar(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica todas las reglas y devuelve una tabla de observaciones."""
    observaciones: list[dict] = []

    # 1. Campos obligatorios vacios
    for indice, fila in df.iterrows():
        numero_fila = indice + 2  # +2: fila 1 son los encabezados
        for columna in COLUMNAS.values():
            valor = str(fila[columna]).strip()
            if not valor:
                agregar_observacion(
                    observaciones,
                    numero_fila,
                    columna,
                    "(vacio)",
                    "Campo obligatorio vacio",
                )

    # 2. DNI: exactamente 8 digitos
    for indice, fila in df.iterrows():
        valor = str(fila[COLUMNAS["dni"]]).strip()
        if valor and not REGEX_DNI.match(valor):
            agregar_observacion(
                observaciones,
                indice + 2,
                COLUMNAS["dni"],
                valor,
                "El DNI debe tener exactamente 8 digitos",
            )

    # 3. Correo con formato valido
    for indice, fila in df.iterrows():
        valor = str(fila[COLUMNAS["correo"]]).strip()
        if valor and not REGEX_CORREO.match(valor):
            agregar_observacion(
                observaciones,
                indice + 2,
                COLUMNAS["correo"],
                valor,
                "Correo con formato invalido",
            )

    # 4. Ciclo entre 1 y 10
    for indice, fila in df.iterrows():
        valor = str(fila[COLUMNAS["ciclo"]]).strip()
        if valor:
            if not valor.isdigit() or not 1 <= int(valor) <= 10:
                agregar_observacion(
                    observaciones,
                    indice + 2,
                    COLUMNAS["ciclo"],
                    valor,
                    "El ciclo debe ser un numero entre 1 y 10",
                )

    # 5. Fecha de matricula valida
    for indice, fila in df.iterrows():
        valor = str(fila[COLUMNAS["fecha_matricula"]]).strip()
        if valor and not es_fecha_valida(valor):
            agregar_observacion(
                observaciones,
                indice + 2,
                COLUMNAS["fecha_matricula"],
                valor,
                "Fecha invalida o con formato incorrecto",
            )

    # 6. Estado permitido
    for indice, fila in df.iterrows():
        valor = str(fila[COLUMNAS["estado"]]).strip()
        if valor and valor.lower() not in ESTADOS_VALIDOS:
            agregar_observacion(
                observaciones,
                indice + 2,
                COLUMNAS["estado"],
                valor,
                f"Estado no reconocido. Validos: {', '.join(sorted(ESTADOS_VALIDOS))}",
            )

    # 7. DNI duplicado
    dnis = df[COLUMNAS["dni"]].astype(str).str.strip()
    duplicados = dnis[dnis.ne("") & dnis.duplicated(keep=False)]
    for indice in duplicados.index:
        agregar_observacion(
            observaciones,
            indice + 2,
            COLUMNAS["dni"],
            dnis.loc[indice],
            "DNI duplicado en el archivo",
        )

    return pd.DataFrame(
        observaciones,
        columns=["Fila", "Columna", "Valor", "Observacion"],
    )


def imprimir_resumen(df: pd.DataFrame, observaciones: pd.DataFrame) -> None:
    """Muestra por consola el resumen de la validacion."""
    filas_con_error = observaciones["Fila"].nunique() if not observaciones.empty else 0

    print("=" * 52)
    print("REPORTE DE VALIDACION DE MATRICULA")
    print("=" * 52)
    print(f"Registros revisados : {len(df)}")
    print(f"Registros con obs.  : {filas_con_error}")
    print(f"Observaciones totales: {len(observaciones)}")
    print("-" * 52)

    if observaciones.empty:
        print("Sin observaciones. Todos los registros son validos.")
        return

    resumen = (
        observaciones.groupby("Observacion")
        .size()
        .sort_values(ascending=False)
    )
    for mensaje, cantidad in resumen.items():
        print(f"{cantidad:>3}  {mensaje}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Valida registros de matricula.")
    parser.add_argument(
        "--archivo",
        type=Path,
        default=ARCHIVO_ENTRADA,
        help="Ruta del Excel de entrada",
    )
    parser.add_argument(
        "--reporte",
        type=Path,
        default=ARCHIVO_REPORTE,
        help="Ruta del Excel de salida",
    )
    args = parser.parse_args()

    try:
        df = cargar_datos(args.archivo)
    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}")
        return

    observaciones = validar(df)
    observaciones.to_excel(args.reporte, index=False)
    imprimir_resumen(df, observaciones)
    print("-" * 52)
    print(f"Reporte guardado en: {args.reporte}")


if __name__ == "__main__":
    main()
