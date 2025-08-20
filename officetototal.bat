@echo off
setlocal enabledelayedexpansion
title Activar Office por Volumen (KMS) y Convertir Retail a VL

echo.
echo ===========================================================
echo  SCRIPT PARA VALIDAR, CONVERTIR A VL Y ACTIVAR OFFICE POR VOLUMEN (KMS)
echo  Incluye Office 2010, 2013, 2016 y superiores (32 y 64 bits)
echo ===========================================================
echo.

:: Solicitar al usuario el servidor KMS
set /p KMS_Servidor=Ingrese la IP o nombre del servidor KMS: 

if "%KMS_Servidor%"=="" (
    echo No ingreso ningun servidor KMS. Saliendo...
    pause
    exit /b
)

:: Detectar ruta y versión de Office
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
    set "OfficePath=%ProgramFiles%\Microsoft Office\Office16"
    set "OfficeVersion=Office 16"
    set "GVLK=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
    set "OfficePath=%ProgramFiles(x86)%\Microsoft Office\Office16"
    set "OfficeVersion=Office 16"
    set "GVLK=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
) else if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
    set "OfficePath=%ProgramFiles%\Microsoft Office\Office15"
    set "OfficeVersion=Office 15"
    set "GVLK=KBHBX-GP9P3-KH4H4-HKJP4-9W72F"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
    set "OfficePath=%ProgramFiles(x86)%\Microsoft Office\Office15"
    set "OfficeVersion=Office 15"
    set "GVLK=KBHBX-GP9P3-KH4H4-HKJP4-9W72F"
) else if exist "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" (
    set "OfficePath=%ProgramFiles%\Microsoft Office\Office14"
    set "OfficeVersion=Office 14"
    set "GVLK=Y7T8G-PW63Y-W9HR3-6D4T2-WV7FG"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" (
    set "OfficePath=%ProgramFiles(x86)%\Microsoft Office\Office14"
    set "OfficeVersion=Office 14"
    set "GVLK=Y7T8G-PW63Y-W9HR3-6D4T2-WV7FG"
) else (
    echo No se encontro una instalacion compatible de Office.
    pause
    exit /b
)

echo Office detectado: %OfficeVersion%
echo Ruta: %OfficePath%
echo.

cd /d "%OfficePath%"
if errorlevel 1 (
    echo No se pudo cambiar al directorio de Office.
    pause
    exit /b
)

echo ===========================================================
echo Paso 1: Mostrar estado actual de la licencia
echo ===========================================================
set "licenseLine="
for /f "tokens=*" %%A in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /i "Tipo de licencia License Type"') do (
    set "licenseLine=%%A"
)
echo !licenseLine!
echo.

echo !licenseLine! | findstr /i "Retail" >nul
if !errorlevel! == 0 (
    echo Licencia RETAIL detectada. Se procedera a cambiar a licencia de volumen (VL).
    echo Ejecutando cambio de clave genérica KMS (GVLK)...
    cscript OSPP.VBS /inpkey:%GVLK%
    echo.
) else (
    echo La licencia NO es Retail, no se requiere cambio de clave.
    echo.
)

pause

echo ===========================================================
echo Paso 2: Configurar servidor KMS (%KMS_Servidor%)
echo ===========================================================
cscript OSPP.VBS /sethst:%KMS_Servidor%
cscript OSPP.VBS /setprt:1688
echo.
pause

echo ===========================================================
echo Paso 3: Activar Office via KMS
echo ===========================================================
cscript OSPP.VBS /act
echo.
pause

echo ===========================================================
echo Paso 4: Verificar estado despues de activacion
echo ===========================================================
cscript OSPP.VBS /dstatus
echo.

echo Proceso completado. Presione una tecla para salir.
pause
