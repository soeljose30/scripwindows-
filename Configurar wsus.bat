@echo off
set /p WSUS_IP=Introduce la dirección IP del servidor WSUS: 

echo Configurando WSUS en %WSUS_IP%...

reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v WUServer /t REG_SZ /d "http://%WSUS_IP%:8530" /f
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v WUStatusServer /t REG_SZ /d "http://%WSUS_IP%:8530" /f

reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v AUOptions /t REG_DWORD /d 3 /f

echo Forzando detección de nuevas actualizaciones...
wuauclt.exe /reportnow /detectnow
wuauclt.exe /resetauthorization

echo Actualizando políticas de grupo...
gpupdate /force

echo Configuración completada.
pause
