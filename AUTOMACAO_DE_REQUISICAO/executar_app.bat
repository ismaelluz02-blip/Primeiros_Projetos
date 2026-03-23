@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
  echo Ambiente virtual nao encontrado em .venv\Scripts\pythonw.exe
  echo Use executar_app_debug.bat para diagnosticar.
  exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" -m src.main
exit /b 0
