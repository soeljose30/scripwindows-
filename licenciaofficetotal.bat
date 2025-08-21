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
