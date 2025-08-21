@echo off
setlocal enabledelayedexpansion

rem --- Define tu clave VL aquí ---
set "VL_KEY=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"

rem --- Detectar ruta de Office desde registro para cualquier version ---
set "OFFICE_PATH="
for /f "tokens=3*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v InstallPath 2^>nul') do (
    set "OFFICE_PATH=%%B"
)

if not defined OFFICE_PATH (
    for /f "tokens=3*" %%A in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v InstallPath 2^>nul') do (
        set "OFFICE_PATH=%%B"
    )
)

if not defined OFFICE_PATH (
    rem Intentar ruta MSI tradicional
    for %%p in (
        "%ProgramFiles%\Microsoft Office\Office16",
        "%ProgramFiles(x86)%\Microsoft Office\Office16",
        "%ProgramFiles%\Microsoft Office\Office15",
        "%ProgramFiles(x86)%\Microsoft Office\Office15",
        "%ProgramFiles%\Microsoft Office\Office14",
        "%ProgramFiles(x86)%\Microsoft Office\Office14"
    ) do (
        if exist "%%~p\ospp.vbs" (
            set "OFFICE_PATH=%%~p"
            goto :found
        )
    )
)

:found
if not defined OFFICE_PATH (
    echo No se encontro instalacion compatible de Office.
    pause
    exit /b 1
)

echo Office instalado en: %OFFICE_PATH%
cd /d "%OFFICE_PATH%"

rem --- Buscar y eliminar claves retail ---
echo Buscando claves Retail instaladas...
set "KEYS_TO_REMOVE="
for /f "tokens=4" %%k in ('cscript ospp.vbs /dstatus ^| findstr /i "Ultimos 5 Caracteres"') do (
    set "KEY=%%k"
    set "KEYS_TO_REMOVE=!KEYS_TO_REMOVE! %%k"
)

if "%KEYS_TO_REMOVE%"=="" (
    echo No se encontraron claves Retail para eliminar.
) else (
    echo Eliminando claves Retail:
    for %%k in (%KEYS_TO_REMOVE%) do (
        echo - Eliminando clave terminada en %%k...
        cscript ospp.vbs /unpkey:%%k >nul
    )
)

rem --- Instalar nueva clave VL ---
echo Instalando clave VL...
cscript ospp.vbs /inpkey:%VL_KEY%

echo.
echo Proceso completado
pause
endlocal
