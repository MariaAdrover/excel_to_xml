@echo off

REM Función principal
:main
echo IMPORTANTE: Los nombres de las columnas del excel y del xsd han de coincidir, aunque las columnas del excel pueden estar desordenadas
echo Agrupar archivos por:
echo 1. NOMBRE de la columna del excel
echo 2. NUMERO de la columna del excel (empieza en 1)
set /p column_type="1. nombre, 2. numero (ENTER crear 1 unico archivo): "

if "%column_type%"=="2" (
    set by_number=--by_number
    set criterio=NUMERO
    call :group_by_column
    goto end
) else if "%column_type%"=="1" (
    set by_number=
    set criterio=NOMBRE
    call :group_by_column
    goto end
) else (
    python convert.py
    goto end
)

REM Función para agrupar por nombre o número de columna
:group_by_column
set /p column="Indica %criterio% de columna para agrupar: "
set /p max_records="Limite de registros por xml (ENTER para no limitar registros): "
if "%max_records%"=="" (
    python convert.py %column% %by_number%
) else (
    python convert.py %column% %by_number% --max_records %max_records%
)
exit /b

:end
pause