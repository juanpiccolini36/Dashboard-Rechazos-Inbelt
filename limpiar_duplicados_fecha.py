"""
limpiar_duplicados_fecha.py
------------------------------------------------------------
Herramienta de USO UNICO para corregir archivos de Rechazos con
filas duplicadas por problemas de precision de fecha:

  PASADA 1 (precision de segundo): cubre el caso de milisegundos
  que no siempre se conservan igual entre distintos origenes del
  dato.

  PASADA 2 (precision de minuto): cubre el caso puntual de filas
  donde 'fecha' y 'fecha_ini' estaban invertidas (ver
  corregir_fecha_invertida.py) - una vez corregido el orden, el
  valor que habia quedado en fecha_ini nunca tuvo precision de
  segundos (fue tipeado a mano sin segundos), asi que sigue sin
  calzar exacto contra la fecha real del crudo. Se usa una
  tolerancia mayor SOLO en esta segunda pasada, y solo sobre lo que
  no se pudo resolver en la primera, para no perder precision en el
  resto del archivo (se probo: redondear todo el archivo al minuto
  de una eleva las colisiones de 6 a 39 sobre ~13.000 filas -
  demasiado riesgo para aplicarlo en general).

En ambas pasadas se usa la MISMA regla de seguridad: una fila solo
se borra si (a) tiene el resaltado amarillo de "fila nueva" Y (b)
todas sus columnas de analisis manual estan vacias Y (c) el grupo
tiene al menos una fila "buena" que sobrevive. Si no se cumple
limpiamente, se deja sin tocar y se reporta como ambiguo.

USO:
    python limpiar_duplicados_fecha.py --rechazos ARCHIVO.xlsx --salida ARCHIVO_LIMPIO.xlsx
"""

import argparse
from openpyxl import load_workbook

COLUMNAS_ANALISIS_MANUAL = [
    "Tipo de falla", "Tipo de Falla",
    "Descripcion de falla",
    "Falla analizada", "Falla Analizada",
    "Solucion implementada", "Solucion Implementada",
    "Descripcion de la solucion",
]

FILL_NUEVO_RGB = "FFFFF2CC"


def _fecha_normalizada(valor_fecha, precision="segundo"):
    if not hasattr(valor_fecha, "replace"):
        return valor_fecha
    try:
        if precision == "minuto":
            return valor_fecha.replace(second=0, microsecond=0)
        return valor_fecha.replace(microsecond=0)
    except (TypeError, ValueError):
        return valor_fecha


def _armar_grupos(ws, idx, col_proceso_nombre, filas_disponibles, precision):
    grupos = {}
    for fila in filas_disponibles:
        cod_programa = ws.cell(row=fila, column=idx["cod_programa"]).value
        proceso = ws.cell(row=fila, column=idx[col_proceso_nombre]).value
        fecha = ws.cell(row=fila, column=idx["fecha"]).value
        cod_art = ws.cell(row=fila, column=idx["cod_art"]).value
        legajo_op = ws.cell(row=fila, column=idx["legajo_op"]).value if "legajo_op" in idx else None
        k = (cod_programa, proceso, _fecha_normalizada(fecha, precision), cod_art, legajo_op)
        grupos.setdefault(k, []).append(fila)
    return grupos


def _procesar_grupos(ws, idx, cols_manual_presentes, grupos):
    """Devuelve (filas_a_borrar, grupos_ambiguos) aplicando la regla de seguridad."""
    filas_a_borrar = []
    grupos_ambiguos = []
    for k, filas in grupos.items():
        if len(filas) < 2:
            continue
        candidatas_borrado = []
        for fila in filas:
            resaltada = ws.cell(row=fila, column=1).fill.start_color.rgb == FILL_NUEVO_RGB
            vacia = all(
                ws.cell(row=fila, column=idx[c]).value in (None, "")
                for c in cols_manual_presentes
            )
            if resaltada and vacia:
                candidatas_borrado.append(fila)

        sobrevivientes = len(filas) - len(candidatas_borrado)
        if candidatas_borrado and sobrevivientes >= 1:
            filas_a_borrar.extend(candidatas_borrado)
        else:
            grupos_ambiguos.append((k, filas))
    return filas_a_borrar, grupos_ambiguos


def limpiar_hoja(ws, con_tolerancia_minuto=False):
    headers = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}
    col_proceso_nombre = "Proceso" if "Proceso" in idx else "proceso"
    cols_manual_presentes = [c for c in COLUMNAS_ANALISIS_MANUAL if c in idx]

    todas_las_filas = list(range(2, ws.max_row + 1))

    # --- Pasada 1: precision de segundo (siempre) ---
    grupos_seg = _armar_grupos(ws, idx, col_proceso_nombre, todas_las_filas, "segundo")
    a_borrar_1, ambiguos_1 = _procesar_grupos(ws, idx, cols_manual_presentes, grupos_seg)

    a_borrar_2, ambiguos_2 = [], []
    if con_tolerancia_minuto:
        # --- Pasada 2: precision de minuto, SOLO si esta hoja lo pidio ---
        # (se probo aplicar esto a todo el archivo por defecto: generaba
        # ambiguedad falsa en hojas que nunca tuvieron el problema de
        # fecha invertida - se limita a las hojas donde se sabe que hizo
        # falta, via --con-tolerancia-minuto)
        filas_restantes = sorted(set(todas_las_filas) - set(a_borrar_1))
        grupos_min = _armar_grupos(ws, idx, col_proceso_nombre, filas_restantes, "minuto")
        a_borrar_2, ambiguos_2 = _procesar_grupos(ws, idx, cols_manual_presentes, grupos_min)

    filas_a_borrar = set(a_borrar_1) | set(a_borrar_2)

    # borrar de abajo hacia arriba para no correr los indices de fila
    for fila in sorted(filas_a_borrar, reverse=True):
        ws.delete_rows(fila, 1)

    ambiguos_finales = ambiguos_2 if con_tolerancia_minuto else ambiguos_1
    return len(filas_a_borrar), ambiguos_finales


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rechazos", required=True)
    ap.add_argument("--salida", default=None)
    ap.add_argument("--con-tolerancia-minuto", default="",
                     help="Nombres de hoja separados por coma donde aplicar tambien la "
                          "pasada de tolerancia al minuto (usar solo en hojas donde se "
                          "corrigio fecha/fecha_ini invertida con corregir_fecha_invertida.py)")
    args = ap.parse_args()
    salida = args.salida or args.rechazos
    hojas_con_tolerancia = {h.strip() for h in args.con_tolerancia_minuto.split(",") if h.strip()}

    wb = load_workbook(args.rechazos)
    total_borradas = 0
    for nombre in wb.sheetnames:
        if not nombre.lower().startswith("rechazos"):
            continue
        ws = wb[nombre]
        headers = [c.value for c in ws[1]]
        if "cod_programa" not in headers or "fecha" not in headers:
            continue
        borradas, ambiguos = limpiar_hoja(ws, con_tolerancia_minuto=(nombre in hojas_con_tolerancia))
        if borradas:
            print(f"  -> '{nombre}': {borradas} filas duplicadas eliminadas.")
        if ambiguos:
            print(f"  -> '{nombre}': {len(ambiguos)} grupos AMBIGUOS sin tocar "
                  f"(revisar a mano, no se pudo decidir con seguridad):")
            for k, filas in ambiguos[:10]:
                print(f"       cod_programa={k[0]} proceso={k[1]} cod_art={k[3]} -> filas {filas}")
        total_borradas += borradas

    wb.save(salida)
    print(f"\nTotal filas eliminadas: {total_borradas}")
    print(f"Guardado: {salida}")


if __name__ == "__main__":
    main()
