@echo off
echo Agrupar archivos por:
echo 1. NOMBRE de la columna del excel
echo 2. NUMERO de la columna del excel (empieza en 1)
set /p column_type="(1 nombre, 2 numero): "

if "%column_type%"=="2" (
    set by_number=--by_number
    set criterio=NUMERO
) else (
    set by_number=
    set criterio=NOMBRE
)

set /p column="Indica %criterio% de columna para agrupar: "
set /p max_records="Enter the maximum number of records per XML file (leave blank if not needed): "

if "%max_records%"=="" (
    python convert.py %column% %by_number%
) else (
    python convert.py %column% %by_number% --max_records %max_records%
)
pause