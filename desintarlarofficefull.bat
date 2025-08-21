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
::Zh4grVQjdCyDJGyX8VAjFC5HWQWQNWSGIrof/eX+4f6Unl4eRusvb4bV3ruZM60v713hSZIoxXNUi98NABpKURStZwwx52taswQ=
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
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
