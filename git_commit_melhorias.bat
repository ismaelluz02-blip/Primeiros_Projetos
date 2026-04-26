@echo off
cd /d "D:\Projetos\FATURAMENTO HORIZONTE"
if exist ".git\index.lock" del /f /q ".git\index.lock"

REM Apaga todos os .bat de commit anteriores
del /f /q git_commit_refatoracao.bat 2>nul
del /f /q git_commit2.bat 2>nul
del /f /q git_commit_sync.bat 2>nul
del /f /q git_commit_importacao.bat 2>nul
del /f /q git_commit_relatorios.bat 2>nul
del /f /q git_commit_dashboard.bat 2>nul
del /f /q git_commit_fase7.bat 2>nul

git add -A
git commit -m "improve: src/cache.py, src/logger.py, setup.py src pkgs, filial config, testes, remove duplicatas"
echo.
echo === Commit concluido ===
git log --oneline -8
pause
