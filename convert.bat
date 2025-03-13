@echo off

REM Función principal
:main
echo IMPORTANTE: Los nombres de las columnas del excel y del xsd han de coincidir, aunque las columnas del excel pueden estar desordenadas
echo Crear 1 unico archivo xml o agrupar por columna en diferentes archivos xml?:
echo 1. Crear 1 unico archivo
echo ENTER (o cualquier otra tecla) agrupar por columna
set /p group="--> "

if "%group%"=="1" (
    python convert.py
    goto end
) else (
    call :group_by_column
    goto end
)

REM Función para agrupar por nombre de columna
:group_by_column
set /p column="NOMBRE de la columna para agrupar: "
set /p max_records="Limite de registros por xml (ENTER para no limitar registros): "
if "%max_records%"=="" (
    python convert.py %column%
) else (
    python convert.py %column% --max_records %max_records%
)
exit /b

:end
pause