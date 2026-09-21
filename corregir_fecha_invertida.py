"""
corregir_fecha_invertida.py
------------------------------------------------------------
Herramienta de USO UNICO. Corrige filas donde los valores de
'fecha' y 'fecha_ini' quedaron cargados AL REVES (error de carga
manual anterior a este proyecto, no generado por ningun script).

Patron detectado: en una fila correcta, 'fecha' tiene fecha+hora
real y 'fecha_ini' es solo la fecha (hora 00:00:00). Cuando esta
invertido, 'fecha' aparece con hora 00:00:00 (parece un dia pelado)
y 'fecha_ini' tiene la hora real.

Solo actua sobre filas que:
  - NO tienen el resaltado amarillo (es decir, son filas originales,
    cargadas a mano antes de que este proyecto existiera - nunca
    toca una fila agregada por el script de actualizacion).
  - Cumplen el patron exacto de inversion.

No borra ni modifica ninguna otra columna. Es reversible: si algo
sale mal, el --salida deja el original intacto.

USO:
    python corregir_fecha_invertida.py --rechazos ARCHIVO.xlsx --salida ARCHIVO_CORREGIDO.xlsx
"""

import argparse
from openpyxl import load_workbook

FILL_NUEVO_RGB = "FFFFF2CC"


def es_medianoche(valor):
    return (hasattr(valor, "time") and valor.time().hour == 0
            and valor.time().minute == 0 and valor.time().second == 0)


def corregir_hoja(ws):
    headers = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}
    if "fecha" not in idx or "fecha_ini" not in idx:
        return 0

    corregidas = 0
    for fila in range(2, ws.max_row + 1):
        resaltada = ws.cell(row=fila, column=1).fill.start_color.rgb == FILL_NUEVO_RGB
        if resaltada:
            continue  # nunca tocar filas agregadas por el script

        c_fecha = ws.cell(row=fila, column=idx["fecha"])
        c_fecha_ini = ws.cell(row=fila, column=idx["fecha_ini"])
        fecha = c_fecha.value
        fecha_ini = c_fecha_ini.value

        if fecha is None or fecha_ini is None:
            continue

        # patron de inversion: fecha en medianoche, fecha_ini con hora real
        if es_medianoche(fecha) and not es_medianoche(fecha_ini):
            c_fecha.value, c_fecha_ini.value = fecha_ini, fecha
            corregidas += 1

    return corregidas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rechazos", required=True)
    ap.add_argument("--salida", default=None)
    args = ap.parse_args()
    salida = args.salida or args.rechazos

    wb = load_workbook(args.rechazos)
    total = 0
    for nombre in wb.sheetnames:
        if not nombre.lower().startswith("rechazos"):
            continue
        ws = wb[nombre]
        corregidas = corregir_hoja(ws)
        if corregidas:
            print(f"  -> '{nombre}': {corregidas} filas con fecha/fecha_ini invertida, corregidas.")
        total += corregidas

    wb.save(salida)
    print(f"\nTotal filas corregidas: {total}")
    print(f"Guardado: {salida}")


if __name__ == "__main__":
    main()
