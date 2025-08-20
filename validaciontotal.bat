@echo off
setlocal enabledelayedexpansion
:: Verificar permisos de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Ejecutando como administrador...
    powershell start -verb runas '%0' %*
    exit /b
)

slmgr.vbs -upk
:: Menú de selección
echo Selecciona tu versión de Windows:
echo [1] Windows 7/8.1/10/11
echo  Windows Server 2008 R2/2012/2012 R2
echo  Windows Server 2016/2019
echo  Windows Server 2022/2025
set /p opcion="Opción (1-4): "
:: Asignar GVLK según versión
if "%opcion%"=="1" (
    set gvlk=W269N-WFGWX-YVC9B-4J6C9-T83GX
) else if "%opcion%"=="2" (
    set gvlk=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
) else if "%opcion%"=="3" (
    set gvlk=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
) else if "%opcion%"=="4" (
    set gvlk=WX4NM-KYWYW-QJJR4-XV3QB-6VM33
) else (
    echo Opción inválida
    pause
    exit /b
)

:: Solicitar IP o nombre del servidor KMS
set /p kmsServer="Introduce la IP o nombre del servidor KMS: "

:: Proceso de activación
echo Instalando clave KMS...
slmgr.vbs /ipk !gvlk! >nul
echo Configurando servidor KMS...
slmgr.vbs -skms !kmsServer! >nul
echo Activando Windows...
slmgr.vbs /ato >nul
:: Verificar estado
slmgr.vbs /xpr
pause
