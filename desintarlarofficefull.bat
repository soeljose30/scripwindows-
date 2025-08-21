@echo off
echo Cerrando cualquier proceso de Office...
taskkill /f /im winword.exe >nul 2>&1
taskkill /f /im excel.exe >nul 2>&1
taskkill /f /im outlook.exe >nul 2>&1
taskkill /f /im powerpnt.exe >nul 2>&1
taskkill /f /im msaccess.exe >nul 2>&1

echo Desinstalando Microsoft Office (si está instalado via MSI)...
REM Ejemplo de comando MSIExec para Office 2013 (ajustar segun versión)
MsiExec.exe /X{90150000-008C-0000-1000-0000000FF1CE} /qn REBOOT=ReallySuppress

echo Borrando carpetas residuales de Office...
rmdir /s /q "C:\Program Files\Microsoft Office"
rmdir /s /q "C:\Program Files (x86)\Microsoft Office"
rmdir /s /q "%APPDATA%\Microsoft\Office"
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office"
rmdir /s /q "%PROGRAMDATA%\Microsoft\Office"

echo Eliminando claves de registro de Office...
reg delete "HKCU\Software\Microsoft\Office" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office" /f
reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office" /f

echo Limpiando licencias de Office instaladas (si existen)...
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /unpkey:XXXXX >nul 2>&1
)
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /unpkey:XXXXX >nul 2>&1
)

echo Proceso completado.
pause
