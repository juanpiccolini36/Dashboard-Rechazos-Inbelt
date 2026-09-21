@echo off
REM ============================================================
REM  Reparar Rechazos - USO UNICO (no es para correr cada semana)
REM ============================================================
REM  Corrige, en un archivo NUEVO (no toca el original):
REM    1) Filas con fecha/fecha_ini invertida
REM    2) Filas duplicadas por precision de fecha
REM
REM  Antes de usar este boton, edita las 2 lineas marcadas con
REM  "EDITAR" mas abajo con tus datos reales.
REM ============================================================

REM --- EDITAR: carpeta donde estan tus archivos ---
cd /d "C:\Rechazos"

REM --- EDITAR: nombre exacto de tu archivo de Rechazos REAL ---
set RECHAZOS=Rechazos_1_semestre_2026.xlsx

REM -------- no hace falta editar nada de aca para abajo --------
set PASO1=%RECHAZOS:.xlsx=%_paso1_temporal.xlsx
set REPARADO=%RECHAZOS:.xlsx=%_REPARADO.xlsx

echo ============================================================
echo  Paso 1 de 2: corrigiendo fecha/fecha_ini invertida
echo ============================================================
python corregir_fecha_invertida.py --rechazos "%RECHAZOS%" --salida "%PASO1%"

echo.
echo ============================================================
echo  Paso 2 de 2: limpiando filas duplicadas
echo ============================================================
python limpiar_duplicados_fecha.py --rechazos "%PASO1%" --salida "%REPARADO%" --con-tolerancia-minuto "Rechazos Moldeo"

del "%PASO1%" >nul 2>&1

echo.
echo ============================================================
echo  LISTO. Tu archivo original NO fue modificado.
echo.
echo  Se genero un archivo nuevo:
echo     %REPARADO%
echo.
echo  Revisalo con atencion. Si te cierra:
echo    1) Cerra tu archivo original (%RECHAZOS%) si lo tenes abierto
echo    2) Borralo o renombralo (ej: agregarle "_viejo")
echo    3) Renombra "%REPARADO%" a "%RECHAZOS%"
echo    4) De ahi en mas segui usando el boton de actualizacion semanal
echo       como siempre
echo ============================================================
pause
