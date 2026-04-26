@echo off
cd /d "D:\Projetos\FATURAMENTO HORIZONTE"
if exist ".git\index.lock" del /f /q ".git\index.lock"

del /f /q git_commit_melhorias.bat 2>nul

git add -A
git commit -m "ui: donut chart, barra destacada, badge impostos, variacao periodo, monitor responsivo"
echo.
echo === Commit concluido ===
git log --oneline -6
pause
