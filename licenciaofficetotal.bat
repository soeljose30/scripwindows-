::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFC5HWQWQNWSGIroL5uT07u6UnkofR+s8d8+Tif3AKeMcig==
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSjk=
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
::Zh4grVQjdCyDJGyX8VAjFC5HWQWQNWSGIrof/eX+4f6Unl4eRusvb4bV3ruZM60v5kzncJgu33tVns0FDx5McQaqYkExsWsi
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
setlocal

echo Introduce la dirección IP o nombre del servidor KMS local:
set /p kmsserver=

echo Detectando la versión de Office instalada...

REM Buscar carpetas comunes de Office para diferentes versiones (16 = 2016/2019/O365, 15 = 2013, 14 = 2010)
set officepath=
if exist "%ProgramFiles%\Microsoft Office\Office16\" set officepath=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\" set officepath=%ProgramFiles(x86)%\Microsoft Office\Office16
if exist "%ProgramFiles%\Microsoft Office\Office15\" set officepath=%ProgramFiles%\Microsoft Office\Office15
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\" set officepath=%ProgramFiles(x86)%\Microsoft Office\Office15
if exist "%ProgramFiles%\Microsoft Office\Office14\" set officepath=%ProgramFiles%\Microsoft Office\Office14
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\" set officepath=%ProgramFiles(x86)%\Microsoft Office\Office14

if "%officepath%"=="" (
    echo No se encontro una version compatible de Office instalada.
    pause
    exit /b 1
)

cd /d "%officepath%"

echo Configurando servidor KMS a %kmsserver% y puerto 1688...

cscript ospp.vbs /sethst:%kmsserver%
cscript ospp.vbs /setprt:1688

echo Activando Office...
cscript ospp.vbs /act

echo Completado. Presiona una tecla para salir.
pause
endlocal
