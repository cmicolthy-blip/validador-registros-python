# Validador de registros de matrícula

Script que revisa un Excel de matrícula, detecta observaciones (datos vacíos,
DNI inválido, correos mal escritos, ciclos fuera de rango, fechas inválidas,
DNI duplicados) y genera un reporte en Excel.

## Problema que resuelve

En matrícula, revisar cientos de registros a mano hace perder tiempo y deja
pasar errores. Este script revisa todo el archivo en segundos y dice
exactamente **qué fila y qué columna** corregir.

## Reglas que valida

| Regla | Ejemplo inválido |
|---|---|
| Campos obligatorios | Nombre vacío |
| DNI de 8 dígitos | `1234567` |
| Correo con formato válido | `correo@example` |
| Ciclo entre 1 y 10 | `0`, `12` |
| Fecha real | `31/02/2026` |
| Estado permitido | `Matriculad` |
| DNI duplicado | Dos filas con el mismo DNI |

## Cómo ejecutarlo

```bash
python crear_datos_ejemplo.py   # crea registros_ejemplo.xlsx
python validador.py             # valida y genera el reporte
```

Con otro archivo:

```bash
python validador.py --archivo mi_matricula.xlsx
```

## Salida

- `reporte_observaciones.xlsx` con columnas: Fila, Columna, Valor, Observación.
- Resumen por consola con el total de registros y observaciones.

## Tecnologías

- Python 3
- pandas
- openpyxl

## Adaptarlo a tu Excel

Si tus columnas tienen otros nombres, edita el diccionario `COLUMNAS` al inicio
de `validador.py`:

```python
COLUMNAS = {
    "dni": "DNI",
    "nombres": "Nombres",
    ...
}
```

## Privacidad

El archivo de ejemplo es ficticio. **No subas datos reales de estudiantes a
GitHub.** Usa datos anonimizados o inventados para tu portafolio.

## Autor

**Orlando Ccahuana Vásquez**
GitHub: [cmicolthy-blip](https://github.com/cmicolthy-blip)
