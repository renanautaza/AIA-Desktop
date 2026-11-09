@echo off
setlocal
title AUTAZA INSPECTION ANALYTICS V1.1.6
cd /d "%~dp0.."
echo ============================================================
echo AUTAZA INSPECTION ANALYTICS V1.1.6
echo INDUSTRIAL CONTROL ROOM - NOVO FRONT-END
echo ============================================================
echo Executando arquivo:
echo %CD%\src\Autaza_Inspection_Analytics.py
echo.
where py >nul 2>&1
if %errorlevel%==0 (
  py -3 "%CD%\src\Autaza_Inspection_Analytics.py"
) else (
  python "%CD%\src\Autaza_Inspection_Analytics.py"
)
if errorlevel 1 (
  echo.
  echo ERRO AO INICIAR O AIA V1.1.6
  pause
)
