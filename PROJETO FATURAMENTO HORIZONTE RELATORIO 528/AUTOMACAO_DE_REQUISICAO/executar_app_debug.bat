@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Ambiente virtual nao encontrado em .venv\Scripts\python.exe
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m src.main
if errorlevel 1 (
  echo.
  echo O app encontrou um erro ao iniciar.
  pause
  exit /b 1
)
