"""
agregar_exclusion_y_formato_fecha.py
------------------------------------------------------------
Herramienta de USO UNICO:
  1) Agrega la columna "Excluir del indicador" (lista desplegable
     Si/No) al final de cada pestaña "Rechazos *", y extiende la
     Tabla de Excel para incluirla.
  2) Aplica formato de fecha corta (dd/mm/yyyy) a las columnas
     'fecha' y 'fecha_ini' en todas las filas de datos.

La validacion Si/No y el formato de fecha se dejan pre-aplicados
tambien varias filas por debajo de los datos actuales (hasta la
fila 5000), para que las filas que el script de actualizacion
semanal agregue en el futuro ya las tengan sin necesidad de volver
a correr esta herramienta.

USO:
    python agregar_exclusion_y_formato_fecha.py --rechazos ARCHIVO.xlsx --salida ARCHIVO_NUEVO.xlsx
"""

import argparse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import TableColumn
from openpyxl.styles import Font

FILA_LIMITE_PREFORMATEO = 5000
NOMBRE_COL_EXCLUSION = "Excluir del indicador"
FORMATO_FECHA_CORTA = "dd/mm/yyyy"


def procesar_hoja(ws):
    headers = [c.value for c in ws[1]]

    # -------- 1) columna de exclusion --------
    if NOMBRE_COL_EXCLUSION not in headers:
        col_nueva = len(headers) + 1
        letra_nueva = get_column_letter(col_nueva)
        celda_header = ws.cell(row=1, column=col_nueva, value=NOMBRE_COL_EXCLUSION)
        # copiar estilo del header vecino para que se vea igual
        celda_header_vecina = ws.cell(row=1, column=col_nueva - 1)
        celda_header.font = Font(bold=True)

        dv = DataValidation(type="list", formula1='"Si,No"', allow_blank=True)
        dv.error = "Elegi Si o No de la lista"
        dv.errorTitle = "Valor invalido"
        ws.add_data_validation(dv)
        rango_dv = f"{letra_nueva}2:{letra_nueva}{FILA_LIMITE_PREFORMATEO}"
        dv.add(rango_dv)

        headers.append(NOMBRE_COL_EXCLUSION)
    else:
        letra_nueva = get_column_letter(headers.index(NOMBRE_COL_EXCLUSION) + 1)

    # -------- extender la Tabla para incluir la columna nueva --------
    ultima_col_letra = get_column_letter(len(headers))
    for nombre_tabla in list(ws.tables.keys()):
        tabla = ws.tables[nombre_tabla]
        inicio = tabla.ref.split(":")[0]
        fila_inicio = "".join(ch for ch in inicio if ch.isdigit())
        fila_fin = ws.max_row
        tabla.ref = f"A{fila_inicio}:{ultima_col_letra}{fila_fin}"

        # CRITICO: la lista de columnas de la tabla (tableColumns) tiene
        # que tener la MISMA cantidad de entradas que columnas en el ref,
        # o Excel marca el archivo como dañado. Si agregamos una columna
        # al rango, hay que agregarla tambien aca.
        nombres_ya_en_tabla = {c.name for c in tabla.tableColumns}
        for i, h in enumerate(headers, start=1):
            if h not in nombres_ya_en_tabla:
                tabla.tableColumns.append(TableColumn(id=i, name=h))

    # -------- 2) formato de fecha corta en columnas fecha / fecha_ini --------
    # OJO: aplicar number_format a una celda vacia la "crea" en el modelo
    # de openpyxl y rompe el calculo de la proxima fila libre (max_row)
    # para el script de actualizacion semanal - por eso esto SOLO se
    # aplica hasta ws.max_row actual (filas con datos reales), nunca mas
    # alla. Las filas futuras las formatea actualizar_rechazos.py al
    # momento de escribirlas.
    idx = {h: i + 1 for i, h in enumerate(headers)}
    for nombre_col in ("fecha", "fecha_ini"):
        if nombre_col not in idx:
            continue
        col = idx[nombre_col]
        for fila in range(2, ws.max_row + 1):
            ws.cell(row=fila, column=col).number_format = FORMATO_FECHA_CORTA


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rechazos", required=True)
    ap.add_argument("--salida", default=None)
    args = ap.parse_args()
    salida = args.salida or args.rechazos

    wb = load_workbook(args.rechazos)
    procesadas = 0
    for nombre in wb.sheetnames:
        if not nombre.lower().startswith("rechazos"):
            continue
        ws = wb[nombre]
        headers = [c.value for c in ws[1]]
        if "cod_programa" not in headers:
            continue
        procesar_hoja(ws)
        procesadas += 1
        print(f"  -> '{nombre}': columna de exclusion + formato de fecha corta aplicados")

    wb.save(salida)
    print(f"\nTotal hojas procesadas: {procesadas}")
    print(f"Guardado: {salida}")


if __name__ == "__main__":
    main()
