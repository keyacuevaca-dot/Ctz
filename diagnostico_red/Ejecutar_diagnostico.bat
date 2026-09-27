@echo off
rem Doble clic para ejecutar el diagnostico con los valores predeterminados (60 minutos).
rem Tambien acepta parametros, por ejemplo:  Ejecutar_diagnostico.bat -Minutos 180 -PlanMbps 100
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnostico_red.ps1" %*
echo.
pause
