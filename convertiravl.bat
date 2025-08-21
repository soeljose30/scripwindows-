::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFC5HWQWQNWSGIroL5uT07u6UnkofR+s8d8+Tif3AKeMcig==
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+JeA==
::cxY6rQJ7JhzQF1fEqQJQ
::ZQ05rAF9IBncCkqN+0xwdVs0
::ZQ05rAF9IAHYFVzEqQJQ
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDGQ=
::cRolqwZ3JBvQF1fEqQJQ
::dhA7uBVwLU+EWDk=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCyDJGyX8VAjFC5HWQWQNWSGIrof/eX+4f6Unl4eRusvb4bV3ruZM60v713hSZUi2Gxfit8FHi1UMBeza28=
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
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
