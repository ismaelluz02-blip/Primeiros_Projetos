@echo off
REM Script para buildear o aplicativo Sistema de Faturamento
REM Requer PyInstaller instalado: pip install pyinstaller

echo.
echo ======================================
echo   BUILD - Sistema de Faturamento
echo ======================================
echo.

REM Verifica se PyInstaller estah instalado
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] PyInstaller nao esta instalado!
    echo Instale com: pip install pyinstaller
    pause
    exit /b 1
)

REM Gera o icone
echo [1/4] Gerando icone a partir da logo...
python gerar_icone.py
if not exist logo.ico (
    echo [AVISO] logo.ico nao foi criado. Continuando sem icone customizado...
)

REM Cria o executavel
echo [2/4] Gerando executavel com PyInstaller...
pyinstaller --onefile ^
    --windowed ^
    --name "Sistema de Faturamento" ^
    --icon "logo.ico" ^
    --add-data "logo.png;." ^
    --add-data "logo.ico;." ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=fitz ^
    --hidden-import=customtkinter ^
    --hidden-import=xlrd ^
    --hidden-import=matplotlib ^
    sistema_faturamento.py

if errorlevel 1 (
    echo [ERRO] Falha ao gerar executavel!
    pause
    exit /b 1
)

echo [3/4] Estrutura de arquivos pronta!
echo.
echo Arquivos gerados em:
echo   - dist\Sistema de Faturamento.exe (Executavel)
echo   - build\ (Arquivos de compilacao)
echo   - sistema_faturamento.spec (Especificacoes)
echo.
echo [4/4] Limpando arquivos temporarios...
rmdir /s /q build
del sistema_faturamento.spec

echo.
echo ======================================
echo   BUILD COMPLETO!
echo ======================================
echo.
echo O executavel esta em: dist\Sistema de Faturamento.exe
echo Voce pode compartilhar este arquivo!
echo.
pause
