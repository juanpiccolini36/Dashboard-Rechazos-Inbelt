# Planilla de Rechazos — Automatización

Automatiza la actualización semanal de la planilla de rechazos (calidad,
producción) a partir del export crudo del sistema de gestión de la
empresa (`Matriz_Rechazos_General.xlsx`).

## Qué hace

Cada semana, `actualizar_rechazos.py`:

1. Lee el crudo (`Matriz_Rechazos_General.xlsx`), que es **acumulativo**
   (trae todo el historial, no solo lo nuevo).
2. Determina el **sector** de cada fila directo desde la columna M
   (`descripcion`) del crudo — no requiere clasificación manual.
3. Agrega **solo las filas nuevas** a la pestaña de cada sector,
   dejando intacto cualquier análisis manual ya cargado (Tipo de
   falla, Descripción, Solución, etc.).
4. Mantiene un maestro de pesos por artículo (`Config_Maestro_Pesos`)
   con fórmula en vivo — corregís un peso una vez y se actualiza en
   todas las filas históricas de ese artículo.
5. Reaplica el desplegable de "Tipo de falla" en todas las hojas en
   cada corrida (auto-reparación — ver sección de gotchas más abajo).
6. Genera un backup automático antes de tocar el archivo real.

Es **idempotente**: correrlo varias veces seguidas sin datos nuevos no
duplica nada.

## Requisitos

- Python 3.10+
- `openpyxl` (`pip install openpyxl --break-system-packages` en Windows
  si da error de permisos)

## Uso semanal

1. Bajar el archivo nuevo del sistema de gestión (mismo nombre siempre,
   pisa al anterior).
2. Doble clic en `Actualizar_Rechazos.bat`.
3. Leer el resumen que tira por consola (qué sectores tuvieron filas
   nuevas, si hay algo para revisar a mano).

La primera vez, editar las 3 líneas de configuración al principio del
`.bat` (carpeta de trabajo, nombre del archivo crudo, nombre del
archivo de Rechazos).

## Estructura del repo

```
actualizar_rechazos.py       -> script principal, corre cada semana
Actualizar_Rechazos.bat      -> lanzador de doble clic para el script principal
Reparar_Rechazos_UNA_VEZ.bat -> lanzador para las herramientas de reparación puntual
herramientas/                -> scripts de uso único (ya aplicados, quedan
                                 documentados por si hace falta repetir
                                 el proceso en otro archivo)
```

Los scripts de `herramientas/` se usaron para corregir problemas
puntuales de datos históricos (ver changelog). El script principal
(`actualizar_rechazos.py`) es autosuficiente para el uso semanal normal
y ya incorpora todas las correcciones — no hace falta volver a correr
las herramientas salvo que aparezca un problema similar en datos
nuevos.

## ⚠️ Los archivos `.xlsx` NO se suben a este repo

Son datos de producción reales, no código, y además son binarios
pesados que no tiene sentido versionar en git. El `.gitignore` ya los
excluye. El archivo de Rechazos real vive en la PC de trabajo, fuera
del repo.

## Decisiones de diseño y lecciones aprendidas

Esto se documenta para que quien toque el código después no tenga que
redescubrir estos problemas — todos costaron horas de diagnóstico real.

### Clave única de fila

```
cod_programa + proceso + fecha (redondeada al segundo) + cod_art + legajo_op
```

- Se probó con menos campos y con más precisión de fecha: generaba
  falsos duplicados o fallaba en distinguir eventos reales (dos
  operarios cargando la misma etapa en paralelo).
- **No se usó `nro_orden`** (más preciso todavía) porque ese campo no
  existe en pestañas armadas a mano antes de este proyecto — haría
  imposible comparar filas viejas contra filas nuevas del crudo.
- Quedan ~3-6 colisiones irreducibles sobre ~13.000 filas: ambigüedad
  real del sistema origen (dos sub-lotes del mismo operario en el
  mismo segundo), no resoluble con los datos disponibles.

### Nunca usar el recalculador de fórmulas (LibreOffice) sobre archivos con formato

Cualquier herramienta que recalcule fórmulas vía LibreOffice (incluido
el `recalc.py` usado durante el desarrollo) **destruye el formato de
Tabla de Excel** (estilo, bandas de color) al volver a guardar. Si hay
que validar fórmulas de un archivo ya formateado, hacerlo con
`openpyxl` directo, nunca pasándolo por un motor de recálculo externo.

### Al extender una Tabla de Excel con una columna nueva

El `ref` de la tabla (rango) y la lista `tableColumns` tienen que tener
la MISMA cantidad de columnas, o Excel marca el archivo como dañado al
abrirlo. Si se agrega una columna al rango, hay que agregar también un
`TableColumn` correspondiente.

### `number_format` en celdas vacías "infla" `max_row`

Pre-formatear una celda vacía (sin valor) para dejarla "lista" para
datos futuros hace que `openpyxl` la cuente como celda ocupada, lo que
corre mal el cálculo de "próxima fila libre" en el script principal.
`DataValidation` NO tiene este problema (es una regla de rango, no
toca celdas individuales) — es seguro pre-aplicarla mucho más allá de
los datos actuales.

### Validación de datos (desplegables) con lista en OTRA hoja

Excel a veces convierte una validación tipo lista que apunta a un
rango de otra hoja a un formato interno más nuevo ("extensión x14")
que `openpyxl` no sabe leer, y la descarta en silencio la próxima vez
que abre el archivo (mensaje: *"Data Validation extension is not
supported and will be removed"*). Por eso el desplegable de "Tipo de
falla" usa una **lista fija en línea** (`"Moldeo,Poliuretano,..."`,
límite de 255 caracteres) en vez de una referencia a
`Config_Tipos_Falla` — y se **reaplica en cada corrida** como
auto-reparación, por si Excel la vuelve a romper.

### Excel no permite nombres de pestaña de más de 31 caracteres

Ver `normalizar_nombre_sector()` para las excepciones de nombre corto
(ej. "Generación de Programas de Producción" → "Gen. Prog. Produccion").

## Changelog (resumen de incidentes resueltos)

- **Duplicación por precisión de milisegundos**: la fecha con
  microsegundos no siempre se conserva igual entre distintos orígenes
  del dato → clave redondeada al segundo.
- **`fecha`/`fecha_ini` invertidas** en datos cargados a mano antes de
  este proyecto → herramienta de corrección puntual, y clave más
  tolerante para ese caso específico (tolerancia al minuto, acotada
  solo a la hoja afectada).
- **Colisión de operarios en paralelo** → se agregó `legajo_op` a la
  clave.
- **Archivo marcado como dañado por Excel** al agregar una columna →
  desalineación entre `ref` de tabla y `tableColumns` (ver arriba).
- **Desplegable de Tipo de falla se "convertía en texto fijo"** → ver
  sección de validación de datos arriba.
