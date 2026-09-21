"""
actualizar_rechazos.py
------------------------------------------------------------
Actualiza la planilla "Rechazos" (multi-sector) a partir del
export crudo de la Matriz General del sistema de gestion.

LOGICA:
  1. Lee la Matriz General (crudo, acumulativo cada 7 dias).
  2. El SECTOR de cada fila se lee directo de la columna M
     (descripcion) de la Matriz General - el sistema ya lo resuelve,
     no requiere clasificacion manual. Ver normalizar_nombre_sector()
     para las excepciones confirmadas con el supervisor (unificacion
     de Mecanizado/Centro de Mecanizado, nombre corto para el sector
     de Generacion de Programas por el limite de 31 caracteres de
     Excel, siglas en mayuscula).
  3. Lee (o crea si no existe) la hoja de configuracion DENTRO del
     archivo de Rechazos:
        - Config_Maestro_Pesos: peso unitario por cod_art, precargada
          con TODOS los articulos de la Matriz General (no solo los
          que ya tenian peso), marcando con "URGENTE" los que alguna
          vez tuvieron un rechazo real. Se va completando a mano y se
          reutiliza siempre que aparezca ese cod_art de nuevo.
  4. Para CADA sector encontrado en la columna M (no hay limite, se
     agregan solos si aparece un sector nuevo en el crudo):
        - Calcula una clave unica por fila: cod_programa+proceso+fecha+cod_art
        - Compara contra las claves que YA existen en la pestaña del
          sector (si la pestaña no existe, se crea).
        - Agrega SOLO las filas nuevas, con las columnas automaticas
          completas (A-J + observacion de referencia) y las columnas
          de analisis manual en blanco, listas para completar.
        - Las filas ya existentes y su analisis manual NO se tocan.
        - Peso Unitario queda como FORMULA que busca en vivo contra
          Config_Maestro_Pesos (si se corrige el peso despues, todas
          las filas de ese articulo recalculan solas).
        - Kg rechazados se escribe como FORMULA (Rechazadas x Peso
          Unitario).

USO:
    python actualizar_rechazos.py --crudo Matriz_Rechazos_General.xlsx \
                                   --rechazos Rechazos_1_semestre_2026_Rev02.xlsx

Se puede correr las veces que se quiera: es IDEMPOTENTE (si no hay
filas nuevas en el crudo, no agrega nada).
"""

import argparse
import datetime as dt
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

# ------------------------------------------------------------------
# Columnas automaticas que siempre vienen del crudo (mismo orden que
# se usan para poblar cualquier pestaña de sector).
# clave interna -> nombre de columna en el crudo (Matriz General)
# ------------------------------------------------------------------
COLS_AUTOMATICAS = {
    "fecha": "fecha",
    "fecha_ini": "fecha_ini",
    "cod_programa": "cod_programa",
    "proceso": "proceso",
    "cod_art": "cod_art",
    "nombre": "nombre",
    "Buenas": "cant_buenas",
    "Rechazadas": "cant_rechazado",
    "legajo_op": "legajo_op",
    "legajo_op2": "legajo_op2",
}
COL_OBS_REF = "OBS_REFERENCIA (Matriz General - NO usar en el indicador)"

FILL_NUEVO = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_HEADER_CONFIG = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")


def leer_crudo(path_crudo):
    """Lee la Matriz General y devuelve lista de dicts (una por fila)."""
    wb = load_workbook(path_crudo, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        filas.append({
            "fecha": row[idx["fecha"]],
            "fecha_ini": row[idx["fecha_ini"]],
            "cod_programa": row[idx["cod_programa"]],
            "proceso": row[idx["proceso"]],
            "cod_art": row[idx["cod_art"]],
            "nombre": row[idx["nombre"]],
            "cant_buenas": row[idx["cant_buenas"]],
            "cant_rechazado": row[idx["cant_rechazado"]],
            "legajo_op": row[idx["legajo_op"]],
            "legajo_op2": row[idx["legajo_op2"]],
            "obs": row[idx["obs"]],
            "sector_raw": row[idx["descripcion"]],  # columna M: sector, directo del sistema
        })
    return filas


# ------------------------------------------------------------------
# Normalizacion del nombre de sector (columna M) -> nombre de pestaña.
# Regla general: "Rechazos " + Título (con conectores en minuscula).
# Excepciones puntuales confirmadas con el supervisor:
#   - CENTRO DE MECANIZADO se unifica con MECANIZADO (mismo sector real)
#   - GENERACION DE PROGRAMAS DE PRODUCCION supera el limite de 31
#     caracteres de Excel para nombres de pestaña -> nombre corto
#   - Siglas (RRHH) se mantienen en mayusculas, no quedan como "Rrhh"
# ------------------------------------------------------------------
UNIFICAR_SECTOR = {
    "CENTRO DE MECANIZADO": "MECANIZADO",
}
NOMBRE_PESTAÑA_ESPECIAL = {
    "GENERACION DE PROGRAMAS DE PRODUCCION": "Rechazos Gen. Prog. Produccion",
}
SIGLAS = {"RRHH"}
CONECTORES = {"de", "del", "la", "las", "el", "los"}


def normalizar_nombre_sector(sector_raw):
    """Convierte el valor crudo de la columna M en el nombre de pestaña final."""
    if not sector_raw:
        return None
    s = sector_raw.strip().upper()
    s = UNIFICAR_SECTOR.get(s, s)
    if s in NOMBRE_PESTAÑA_ESPECIAL:
        return NOMBRE_PESTAÑA_ESPECIAL[s]
    if s in SIGLAS:
        return f"Rechazos {s}"
    palabras = []
    for i, palabra in enumerate(s.split(" ")):
        pl = palabra.lower()
        if i > 0 and pl in CONECTORES:
            palabras.append(pl)
        else:
            palabras.append(pl.capitalize())
    return "Rechazos " + " ".join(palabras)


def _fecha_normalizada(valor_fecha):
    """
    Redondea la fecha al segundo entero (descarta microsegundos).
    Excel no siempre conserva los milisegundos con la misma precision
    entre distintos origenes del dato (tipeado a mano, pegado, o
    generado por el sistema) - si se usa la fecha exacta hasta el
    microsegundo en la clave, dos registros del MISMO evento real
    pueden no coincidir por una diferencia de milisegundos, y el
    script terminaria agregando una fila "nueva" que en realidad ya
    existia (falso duplicado).
    """
    if hasattr(valor_fecha, "replace"):
        try:
            return valor_fecha.replace(microsecond=0)
        except (TypeError, ValueError):
            return valor_fecha
    return valor_fecha


def clave(fila_dict):
    """
    Clave unica de fila:
      cod_programa + proceso + fecha (redondeada al segundo) + cod_art + legajo_op

    Se agrego legajo_op porque se detectaron casos reales en el sistema
    origen donde DOS operarios distintos cargan la misma etapa del mismo
    pedido en paralelo (mismo cod_programa+proceso+fecha+cod_art, pero
    cada uno con su propia cantidad) - sin legajo_op estas dos filas
    reales y distintas se confundian como una sola.
    No se uso 'nro_orden' (mas preciso todavia) porque ese campo no
    existe en las pestañas armadas a mano antes de este script, y se
    necesita poder comparar filas viejas contra filas nuevas del crudo.
    """
    return (fila_dict["cod_programa"], fila_dict["proceso"],
            _fecha_normalizada(fila_dict["fecha"]), fila_dict["cod_art"],
            fila_dict.get("legajo_op"))


def asegurar_hoja_config(wb):
    """Crea la hoja de configuracion de pesos si todavia no existe.
    (Config_Mapeo_Sector quedo obsoleta: el sector ahora se lee
    directo de la columna M de la Matriz General, no requiere
    clasificacion manual. Si existe de una version anterior del
    archivo, se deja intacta sin usarla, por si tiene algo cargado
    que el supervisor quiera consultar.)"""
    if "Config_Maestro_Pesos" not in wb.sheetnames:
        ws = wb.create_sheet("Config_Maestro_Pesos")
        ws.append(["cod_art", "nombre (referencia)", "Peso Unitario (kg)", "Prioridad", "Observacion"])
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = FILL_HEADER_CONFIG
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 55
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 55
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:E1"


def leer_maestro_pesos(wb):
    ws = wb["Config_Maestro_Pesos"]
    pesos = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is not None and r[2] is not None:
            pesos[r[0]] = r[2]
    return pesos


def sincronizar_articulos_desde_matriz(wb, crudo):
    """
    Asegura que TODO articulo (cod_art) presente en la Matriz General
    tenga una fila en Config_Maestro_Pesos, con su nombre de referencia
    y marcado de prioridad segun si alguna vez tuvo un rechazo. Nunca
    pisa un peso ya cargado. Es la fuente completa (no solo lo que ya
    aparecia en las pestañas de Rechazos).
    """
    ws = wb["Config_Maestro_Pesos"]
    ya_existe = {r[0].value for r in ws.iter_rows(min_row=2) if r[0].value is not None}

    nombre_por_art = {}
    tuvo_rechazo = set()
    for f in crudo:
        art = f["cod_art"]
        if not art:
            continue
        nombre_por_art.setdefault(art, f["nombre"])
        if f["cant_rechazado"] and f["cant_rechazado"] > 0:
            tuvo_rechazo.add(art)

    fila = ws.max_row + 1
    agregados = 0
    for art in sorted(nombre_por_art):
        if art in ya_existe:
            continue
        prioridad = "URGENTE - tuvo rechazo historico" if art in tuvo_rechazo else "Sin rechazo historico"
        ws.cell(row=fila, column=1, value=art)
        ws.cell(row=fila, column=2, value=nombre_por_art[art])
        # columna 3 (Peso Unitario) se deja en blanco a proposito
        ws.cell(row=fila, column=4, value=prioridad)
        if prioridad.startswith("URGENTE"):
            ws.cell(row=fila, column=4).font = Font(color="C00000", bold=True)
        fila += 1
        agregados += 1
    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    return agregados


def rescatar_pesos_de_hojas_existentes(wb, hojas_a_ignorar=("Config_Mapeo_Sector",
                                                             "Config_Maestro_Pesos")):
    """
    Recorre TODAS las pestañas de sector que ya existen en el archivo
    y rescata cualquier 'Peso Unitario' ya cargado a mano en el pasado,
    para no perder ese trabajo. Completa la fila correspondiente en
    Config_Maestro_Pesos (que ya existe gracias a
    sincronizar_articulos_desde_matriz) solo si esa fila todavia no
    tiene peso cargado. Si el mismo cod_art tiene pesos distintos en
    distintas filas, NO elige uno solo: lo deja marcado como conflicto
    en la columna Observacion para que el supervisor lo resuelva.
    Nunca pisa un peso que ya este cargado en Config_Maestro_Pesos.
    """
    ws_cfg = wb["Config_Maestro_Pesos"]
    # cod_art -> numero de fila en Config_Maestro_Pesos
    fila_por_art = {}
    peso_ya_cargado = {}
    for r in ws_cfg.iter_rows(min_row=2):
        if r[0].value is not None:
            fila_por_art[r[0].value] = r[0].row
            peso_ya_cargado[r[0].value] = r[2].value  # columna C = Peso Unitario

    encontrados = {}   # cod_art -> set(pesos)
    for sheet_name in wb.sheetnames:
        if sheet_name in hojas_a_ignorar or not sheet_name.lower().startswith("rechazos"):
            continue
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        if "cod_art" not in headers or "Peso Unitario" not in headers:
            continue
        col_art = headers.index("cod_art") + 1
        col_peso = headers.index("Peso Unitario") + 1
        for r in range(2, ws.max_row + 1):
            art = ws.cell(row=r, column=col_art).value
            peso = ws.cell(row=r, column=col_peso).value
            # ignorar celdas que ya son formula (=IFERROR(INDEX...)) generadas
            # por este mismo script en corridas anteriores: no son un peso
            # "tipeado a mano", son un link al maestro, no aporta nada rescatar eso.
            if art is not None and isinstance(peso, (int, float)):
                encontrados.setdefault(art, set()).add(peso)

    agregados, conflictos = 0, 0
    fila_nueva = ws_cfg.max_row + 1
    for art, pesos_set in sorted(encontrados.items()):
        if art in peso_ya_cargado and peso_ya_cargado[art] is not None:
            continue  # ya tiene peso cargado en el maestro, no tocar

        if art in fila_por_art:
            fila = fila_por_art[art]
        else:
            # cod_art que no vino de la Matriz General (caso raro) -> se agrega
            fila = fila_nueva
            ws_cfg.cell(row=fila, column=1, value=art)
            fila_nueva += 1

        if len(pesos_set) == 1:
            ws_cfg.cell(row=fila, column=3, value=next(iter(pesos_set)))
            agregados += 1
        else:
            obs_cell = ws_cfg.cell(
                row=fila, column=5,
                value=f"CONFLICTO: valores distintos encontrados {sorted(pesos_set)} "
                      f"- revisar y cargar el correcto en la columna C")
            obs_cell.font = Font(color="FF0000", italic=True)
            conflictos += 1
    return agregados, conflictos


COL_EXCLUSION = "Excluir del indicador"
FILA_LIMITE_PREFORMATEO = 5000
FORMATO_FECHA_CORTA = "dd/mm/yyyy"


def obtener_o_crear_hoja_sector(wb, sector, hoja_moldeo_como_plantilla=True):
    """
    Si la pestaña del sector no existe, la crea con el layout estandar
    (igual al de 'Rechazos Moldeo': incluye Descripcion de la solucion,
    columna de exclusion del indicador, formato de fecha corta y
    validacion Si/No pre-aplicados hasta la fila 5000).
    """
    nombre_hoja = sector
    if nombre_hoja in wb.sheetnames:
        return wb[nombre_hoja]

    ws = wb.create_sheet(nombre_hoja)
    headers = ["fecha", "fecha_ini", "cod_programa", "Proceso", "cod_art", "nombre",
               "Buenas", "Rechazadas", "legajo_op", "legajo_op2",
               "Tipo de falla", "Descripcion de falla", "Falla analizada",
               "Solucion implementada", "Descripcion de la solucion",
               "Peso Unitario", "Kg rechazados", COL_OBS_REF, COL_EXCLUSION]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    idx = {h: i + 1 for i, h in enumerate(headers)}

    # validacion Si/No pre-aplicada en la columna de exclusion (esto NO
    # infla max_row porque no toca celdas individuales, solo agrega una
    # regla de rango - a diferencia de number_format, que si "crea" la
    # celda y rompe el calculo de la proxima fila libre)
    col_excl_letra = get_column_letter(idx[COL_EXCLUSION])
    dv = DataValidation(type="list", formula1='"Si,No"', allow_blank=True)
    dv.error = "Elegi Si o No de la lista"
    dv.errorTitle = "Valor invalido"
    ws.add_data_validation(dv)
    dv.add(f"{col_excl_letra}2:{col_excl_letra}{FILA_LIMITE_PREFORMATEO}")

    # el sector nuevo se registra en la lista fuente de Tipo de falla;
    # el desplegable en si se aplica centralizado, una sola vez por
    # corrida, en reaplicar_dropdown_tipo_falla() - ver esa funcion
    # para el motivo (evitar el bug de validacion contra otra hoja).
    ws_lista = asegurar_hoja_lista_tipos_falla(wb)
    ya_en_lista = {ws_lista.cell(row=r, column=1).value for r in range(2, ws_lista.max_row + 1)}
    if sector.replace("Rechazos ", "", 1) not in ya_en_lista and sector not in ya_en_lista:
        ws_lista.cell(row=ws_lista.max_row + 1, column=1,
                       value=sector.replace("Rechazos ", "", 1))

    return ws


NOMBRE_HOJA_LISTA_TIPOS_FALLA = "Config_Tipos_Falla"


def asegurar_hoja_lista_tipos_falla(wb):
    """Crea la hoja de referencia para el desplegable de Tipo de falla,
    si todavia no existe. Se puebla con los sectores existentes al
    crear cada pestaña nueva - no hace falta mantenerla a mano."""
    if NOMBRE_HOJA_LISTA_TIPOS_FALLA in wb.sheetnames:
        return wb[NOMBRE_HOJA_LISTA_TIPOS_FALLA]
    ws = wb.create_sheet(NOMBRE_HOJA_LISTA_TIPOS_FALLA)
    ws["A1"] = "Tipo de falla (lista fuente del desplegable - se completa sola)"
    ws["A1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 55
    return ws


def reaplicar_dropdown_tipo_falla(wb):
    """
    Aplica (o re-aplica) la validacion de "Tipo de falla" en TODAS las
    pestañas de sector, usando una LISTA FIJA EN LINEA (ej. "Moldeo,
    Poliuretano,...") en vez de una referencia a la hoja
    Config_Tipos_Falla.

    Por que: se detecto que Excel, al guardar una validacion tipo
    lista que apunta a un rango de OTRA hoja, a veces la convierte a
    un formato interno mas nuevo ("extension x14") que openpyxl no
    sabe leer - y la descarta en silencio la proxima vez que abre el
    archivo (aparece como "Data Validation extension is not
    supported and will be removed"). Una lista fija en linea (como
    ya se usaba para "Excluir del indicador": "Si,No") no tiene ese
    problema porque nunca dispara esa conversion.

    Se llama en CADA corrida, sobre TODAS las hojas (no solo las
    nuevas), asi si alguna vez Excel vuelve a romper la validacion,
    la proxima corrida la repara sola sin que el supervisor tenga que
    acordarse de nada.

    Limite: Excel permite hasta 255 caracteres en una lista en linea.
    Con la cantidad de sectores actual sobra margen; si en el futuro
    la lista de sectores crece mucho, esto lo va a avisar por consola.
    """
    ws_lista = asegurar_hoja_lista_tipos_falla(wb)
    sectores = [ws_lista.cell(row=r, column=1).value
                for r in range(2, ws_lista.max_row + 1)
                if ws_lista.cell(row=r, column=1).value]
    if not sectores:
        return

    lista_inline = ",".join(sectores)
    if len(lista_inline) > 255:
        print(f"  -> ADVERTENCIA: la lista de Tipo de falla tiene {len(lista_inline)} "
              f"caracteres, supera el limite de 255 de Excel para listas en linea. "
              f"El desplegable puede no funcionar bien - avisar para resolverlo.")

    for nombre in wb.sheetnames:
        if not nombre.lower().startswith("rechazos"):
            continue
        ws = wb[nombre]
        headers = [c.value for c in ws[1]]
        if "Tipo de falla" not in headers:
            continue
        col = headers.index("Tipo de falla") + 1
        letra = get_column_letter(col)

        # sacar cualquier validacion previa en esta columna (la vieja
        # rota contra Config_Tipos_Falla, o una version anterior de
        # esta misma) para no ir acumulando reglas duplicadas
        for dv_vieja in list(ws.data_validations.dataValidation):
            if str(dv_vieja.sqref).startswith(f"{letra}2:{letra}"):
                ws.data_validations.dataValidation.remove(dv_vieja)

        dv = DataValidation(type="list", formula1=f'"{lista_inline}"', allow_blank=True)
        dv.error = "Elegi un tipo de falla de la lista"
        dv.errorTitle = "Valor invalido"
        ws.add_data_validation(dv)
        dv.add(f"{letra}2:{letra}{FILA_LIMITE_PREFORMATEO}")


def actualizar_sector(wb, sector, crudo_sector):
    ws = obtener_o_crear_hoja_sector(wb, sector)
    headers = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}  # 1-based para openpyxl

    # nombre de columna de proceso puede ser 'Proceso' o 'proceso' segun la hoja
    col_proceso_nombre = "Proceso" if "Proceso" in idx else "proceso"

    # tiene o no la columna de referencia de obs
    tiene_obs_ref = COL_OBS_REF in idx

    # claves ya existentes en la hoja
    claves_existentes = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        k = (d.get("cod_programa"), d.get(col_proceso_nombre),
             _fecha_normalizada(d.get("fecha")), d.get("cod_art"), d.get("legajo_op"))
        claves_existentes.add(k)

    # fila que sirve de "plantilla" de formato (ultima fila con datos ya existente)
    fila_plantilla = ws.max_row if ws.max_row >= 2 else None

    nuevas = 0
    fila = ws.max_row + 1
    for f in crudo_sector:
        k = clave(f)
        if k in claves_existentes:
            continue
        # fila nueva -> completar columnas automaticas, resto en blanco
        # 1) copiar el formato (fecha, fuente, bordes, alineacion) de la ultima
        #    fila existente ANTES de escribir valores, para que la fila nueva
        #    se vea igual que el resto (numero_format de fecha, etc.)
        if fila_plantilla:
            for col in range(1, len(headers) + 1):
                origen = ws.cell(row=fila_plantilla, column=col)
                destino = ws.cell(row=fila, column=col)
                destino.number_format = origen.number_format
                destino.font = copy(origen.font)
                destino.border = copy(origen.border)
                destino.alignment = copy(origen.alignment)

        ws.cell(row=fila, column=idx["fecha"], value=f["fecha"]).number_format = FORMATO_FECHA_CORTA
        ws.cell(row=fila, column=idx["fecha_ini"], value=f["fecha_ini"]).number_format = FORMATO_FECHA_CORTA
        ws.cell(row=fila, column=idx["cod_programa"], value=f["cod_programa"])
        ws.cell(row=fila, column=idx[col_proceso_nombre], value=f["proceso"])
        ws.cell(row=fila, column=idx["cod_art"], value=f["cod_art"])
        ws.cell(row=fila, column=idx["nombre"], value=f["nombre"])
        ws.cell(row=fila, column=idx["Buenas"], value=f["cant_buenas"])
        ws.cell(row=fila, column=idx["Rechazadas"], value=f["cant_rechazado"])
        ws.cell(row=fila, column=idx["legajo_op"], value=f["legajo_op"])
        ws.cell(row=fila, column=idx["legajo_op2"], value=f["legajo_op2"])

        # Peso Unitario: formula que busca en Config_Maestro_Pesos (fuente
        # unica de verdad). Si el dia de mañana se corrige el peso en el
        # maestro, esta celda recalcula sola - no queda un valor "congelado".
        col_peso = idx.get("Peso Unitario")
        if col_peso:
            col_art_letra = get_column_letter(idx["cod_art"])
            formula_lookup = (f'INDEX(Config_Maestro_Pesos!$C:$C,'
                               f'MATCH({col_art_letra}{fila},Config_Maestro_Pesos!$A:$A,0))')
            ws.cell(row=fila, column=col_peso,
                    value=f'=IFERROR(IF({formula_lookup}="","",{formula_lookup}),"")')

        # Kg rechazados como formula = Rechazadas * Peso Unitario
        if "Kg rechazados" in idx or "Kg Rechazados" in idx:
            col_kg = idx.get("Kg rechazados") or idx.get("Kg Rechazados")
            col_rech_letra = get_column_letter(idx["Rechazadas"])
            col_peso_letra = get_column_letter(idx.get("Peso Unitario"))
            ws.cell(row=fila, column=col_kg,
                    value=f"=IF({col_peso_letra}{fila}=\"\",0,{col_rech_letra}{fila}*{col_peso_letra}{fila})")
        if tiene_obs_ref:
            ws.cell(row=fila, column=idx[COL_OBS_REF], value=f["obs"])

        # resaltar la fila como "nueva, pendiente de analisis"
        # (se aplica DESPUES de copiar formato, para que el amarillo no se pierda)
        for col in range(1, len(headers) + 1):
            ws.cell(row=fila, column=col).fill = FILL_NUEVO

        fila += 1
        nuevas += 1

    # extender cualquier Tabla de Excel (ListObject) definida en la hoja para
    # que incluya las filas nuevas; si no hay tabla definida, no hace nada.
    if nuevas > 0:
        for tabla in list(ws.tables.values()):
            col_inicio, fila_inicio, col_fin, fila_fin_actual = _rango_tabla(tabla.ref)
            nuevo_ref = f"{col_inicio}{fila_inicio}:{col_fin}{fila - 1}"
            tabla.ref = nuevo_ref

    return nuevas


def _rango_tabla(ref):
    """Descompone un ref tipo 'A1:Q393' en sus componentes."""
    import re
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2)), m.group(3), int(m.group(4))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--crudo", required=True)
    ap.add_argument("--rechazos", required=True)
    ap.add_argument("--salida", default=None, help="Por defecto sobreescribe --rechazos")
    ap.add_argument("--sin-backup", action="store_true",
                     help="No generar copia de respaldo antes de sobreescribir")
    args = ap.parse_args()

    salida = args.salida or args.rechazos
    sobreescribe_original = (salida == args.rechazos)

    # -------- chequeo previo: el archivo no debe estar abierto en Excel --------
    # En Windows, un archivo abierto en Excel no se puede abrir en modo
    # escritura desde otro programa -> intentamos abrir/cerrar en modo 'a'
    # (append) como forma rapida de detectar el bloqueo antes de procesar
    # todo (evita perder tiempo si total se va a cortar igual al final).
    import os
    for chequear in (args.crudo, args.rechazos):
        if os.path.exists(chequear):
            try:
                f = open(chequear, "a")
                f.close()
            except PermissionError:
                print(f"\n*** ERROR: '{chequear}' esta abierto en Excel (u otro programa). ***")
                print("*** Cerralo y volve a correr la actualizacion.                     ***\n")
                return

    # -------- backup del archivo de Rechazos antes de tocarlo --------
    if sobreescribe_original and not args.sin_backup:
        carpeta_backup = os.path.join(os.path.dirname(os.path.abspath(args.rechazos)), "Backups")
        os.makedirs(carpeta_backup, exist_ok=True)
        marca = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        nombre_base = os.path.basename(args.rechazos)
        destino_backup = os.path.join(carpeta_backup, f"{marca}__{nombre_base}")
        import shutil
        shutil.copy2(args.rechazos, destino_backup)
        print(f"Backup guardado en: {destino_backup}")

    print(f"Leyendo crudo: {args.crudo}")
    crudo = leer_crudo(args.crudo)
    print(f"  -> {len(crudo)} filas en Matriz General")

    wb = load_workbook(args.rechazos)
    asegurar_hoja_config(wb)

    agregados_art = sincronizar_articulos_desde_matriz(wb, crudo)
    if agregados_art:
        print(f"  -> Config_Maestro_Pesos: {agregados_art} articulos nuevos agregados "
              f"desde la Matriz General (peso en blanco, listos para completar).")

    agregados, conflictos = rescatar_pesos_de_hojas_existentes(wb)
    if agregados or conflictos:
        print(f"  -> Config_Maestro_Pesos: {agregados} pesos completados automaticamente "
              f"a partir de lo ya cargado en las hojas, {conflictos} con valores en "
              f"conflicto (revisar a mano).")

    # el sector de cada fila sale directo de la columna M (descripcion) de
    # la Matriz General - ya no requiere clasificacion manual
    for f in crudo:
        f["sector"] = normalizar_nombre_sector(f["sector_raw"])

    sectores = sorted(set(f["sector"] for f in crudo if f["sector"]))
    for sector in sectores:
        crudo_sector = [f for f in crudo if f["sector"] == sector]
        nuevas = actualizar_sector(wb, sector, crudo_sector)
        print(f"  -> Sector '{sector}': {len(crudo_sector)} filas totales en crudo, "
              f"{nuevas} filas nuevas agregadas.")

    # se reaplica el desplegable de Tipo de falla en TODAS las hojas en
    # cada corrida (auto-reparacion, ver docstring de la funcion)
    reaplicar_dropdown_tipo_falla(wb)
    print("  -> Desplegable de 'Tipo de falla' verificado/reparado en todas las hojas.")

    try:
        wb.save(salida)
    except PermissionError:
        print(f"\n*** ERROR: no se pudo guardar '{salida}'. ¿Esta abierto en Excel? ***")
        print("*** Cerralo y volve a correr la actualizacion.                      ***\n")
        return

    print(f"Guardado: {salida}")


if __name__ == "__main__":
    main()
