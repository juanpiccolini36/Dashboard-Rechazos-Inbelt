"""
formato_encabezados.py
------------------------------------------------------------
Aplica formato uniforme a la fila de encabezado (fila 1) de todas
las pestañas "Rechazos *" y las hojas de configuracion:
  - Fuente Cambria, tamaño 12, color blanco, negrita
  - Relleno azul intermedio (#0070C0)

USO:
    python formato_encabezados.py --rechazos ARCHIVO.xlsx --salida ARCHIVO_NUEVO.xlsx
"""

import argparse
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

COLOR_FONDO = "1F4E78"
COLOR_LETRA = "FFFFFF"
FUENTE = "Arial"
TAMANIO = 12
ALTO_FILA_1 = 30

HOJAS_A_FORMATEAR_ADEMAS_DE_RECHAZOS = ["Config_Maestro_Pesos", "Config_Tipos_Falla"]


def formatear_encabezado(ws):
    fuente = Font(name=FUENTE, size=TAMANIO, bold=False, color=COLOR_LETRA)
    relleno = PatternFill(fill_type="solid", start_color=COLOR_FONDO, end_color=COLOR_FONDO)
    for celda in ws[1]:
        if celda.value is None:
            continue
        celda.font = fuente
        celda.fill = relleno
    ws.row_dimensions[1].height = ALTO_FILA_1

    # corregir posicion de scroll guardada (independiente del panel
    # inmovilizado) - si quedo grabada una fila distinta de la 1, Excel
    # abre la hoja "scrolleada" aunque el freeze este bien configurado
    ws.sheet_view.topLeftCell = None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rechazos", required=True)
    ap.add_argument("--salida", default=None)
    args = ap.parse_args()
    salida = args.salida or args.rechazos

    wb = load_workbook(args.rechazos)
    aplicadas = 0
    for nombre in wb.sheetnames:
        es_rechazos = nombre.lower().startswith("rechazos")
        es_config = nombre in HOJAS_A_FORMATEAR_ADEMAS_DE_RECHAZOS
        if not (es_rechazos or es_config):
            continue
        ws = wb[nombre]
        formatear_encabezado(ws)
        aplicadas += 1
        print(f"  -> encabezado formateado en '{nombre}'")

    wb.save(salida)
    print(f"\nTotal hojas formateadas: {aplicadas}")
    print(f"Guardado: {salida}")


if __name__ == "__main__":
    main()
