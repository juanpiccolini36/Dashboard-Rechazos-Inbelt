@echo off
REM ============================================================
REM  Actualizar Rechazos - doble clic para correr
REM ============================================================
REM  Preconfigurado para el archivo generado desde cero.
REM  Solo hace falta editar la primera linea (la carpeta) si tu
REM  carpeta de trabajo no es C:\Rechazos
REM ============================================================

REM --- EDITAR SOLO SI TU CARPETA ES DISTINTA ---
cd /d "C:\Rechazos"

REM --- estos dos nombres ya estan seteados, no hace falta tocarlos ---
set CRUDO=Matriz_Rechazos_General.xlsx
set RECHAZOS=Rechazos_DESDE_CERO.xlsx

echo ============================================================
echo  Actualizando %RECHAZOS%
echo  usando datos de %CRUDO%
echo ============================================================
echo.

python actualizar_rechazos.py --crudo "%CRUDO%" --rechazos "%RECHAZOS%"

echo.
echo ============================================================
echo  Proceso terminado. Revisa los mensajes de arriba.
echo  (Si dice ERROR, leelo con atencion antes de cerrar esta ventana)
echo ============================================================
pause
