@echo off
setlocal EnableExtensions

set "PY_CMD="
set "PY_EXE="
set "PY_DIR="
set "PY_SCRIPTS="
set "APP_VERSION="
set "MSI_FILE="
set "PATH_ORIG=%PATH%"

echo.
echo ======================================
echo   GERAR INSTALADOR MSI
echo ======================================
echo.

cd /d "%~dp0"

echo [1/8] Selecionando versao do Python...
py -3.11 -V >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3.11"
) else (
    py -3.12 -V >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py -3.12"
    ) else (
        py -3.10 -V >nul 2>&1
        if not errorlevel 1 (
            set "PY_CMD=py -3.10"
        )
    )
)

if "%PY_CMD%"=="" (
    python -V >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=python"
    )
)

if "%PY_CMD%"=="" (
    echo [ERRO] Nao encontrei Python compativel.
    echo Tente instalar Python 3.11 e reinicie o computador.
    pause
    exit /b 1
)

echo     - Usando: %PY_CMD%
for /f "usebackq delims=" %%I in (`%PY_CMD% -c "import sys; print(sys.executable)"`) do set "PY_EXE=%%I"
if "%PY_EXE%"=="" (
    echo [ERRO] Nao foi possivel localizar o executavel do Python.
    pause
    exit /b 1
)

for %%I in ("%PY_EXE%") do set "PY_DIR=%%~dpI"
set "PY_DIR=%PY_DIR:~0,-1%"
set "PY_SCRIPTS=%PY_DIR%\Scripts"
echo     - Executavel: %PY_EXE%

echo [2/8] Lendo APP_VERSION do setup.py...
for /f "tokens=2 delims==" %%I in ('findstr /R /C:"^APP_VERSION[ ]*=[ ]*\".*\"" setup.py') do set "APP_VERSION=%%~I"
set "APP_VERSION=%APP_VERSION: =%"
set "APP_VERSION=%APP_VERSION:"=%"
if "%APP_VERSION%"=="" (
    echo [ERRO] Nao consegui ler APP_VERSION em setup.py.
    pause
    exit /b 1
)
echo     - Versao alvo: %APP_VERSION%

echo [3/8] Atualizando dependencias de build...
echo     - Atualizando pip...
"%PY_EXE%" -m pip install --upgrade pip --disable-pip-version-check
if errorlevel 1 (
    echo [ERRO] Falha ao atualizar o pip.
    pause
    exit /b 1
)

echo     - Instalando requirements...
"%PY_EXE%" -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo [4/8] Gerando icone...
"%PY_EXE%" gerar_icone.py
if not exist logo.ico (
    echo [ERRO] logo.ico nao foi gerado.
    pause
    exit /b 1
)

echo [5/8] Validando scripts principais...
"%PY_EXE%" -m py_compile setup.py
if errorlevel 1 (
    echo [ERRO] setup.py possui erro de sintaxe.
    pause
    exit /b 1
)
"%PY_EXE%" -m py_compile sistema_faturamento.py
if errorlevel 1 (
    echo [ERRO] sistema_faturamento.py possui erro de sintaxe.
    pause
    exit /b 1
)

echo [6/8] Limpando artefatos antigos...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [7/8] Gerando MSI com cx_Freeze...
echo     - Limpando PATH para evitar pastas protegidas...
set "PATH=%PY_SCRIPTS%;%PY_DIR%;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0"
"%PY_EXE%" setup.py bdist_msi
set "PATH=%PATH_ORIG%"
if errorlevel 1 (
    echo [ERRO] Falha ao gerar o instalador MSI.
    pause
    exit /b 1
)

echo [8/8] Finalizando...
for %%F in ("dist\*-%APP_VERSION%-*.msi") do (
    if exist "%%~fF" set "MSI_FILE=%%~fF"
)
if "%MSI_FILE%"=="" (
    for /f "delims=" %%F in ('dir /b /o-d "dist\*.msi" 2^>nul') do (
        if not defined MSI_FILE set "MSI_FILE=%CD%\dist\%%F"
    )
)
if "%MSI_FILE%"=="" (
    echo [ERRO] Nenhum arquivo .msi foi encontrado em dist\.
    echo Verifique as mensagens acima: o build pode ter concluido com aviso/erro.
    pause
    exit /b 1
)

echo.
echo Instalador gerado com sucesso:
echo %MSI_FILE%
echo.
echo Proximo passo recomendado:
echo 1^) Feche o app aberto em outras maquinas.
echo 2^) Execute este MSI por cima da versao anterior.
echo.
echo [I] Instalar agora nesta maquina
echo [P] Abrir pasta dist
echo [S] Sair
choice /C IPS /N /M "Escolha (I/P/S): "
if errorlevel 3 goto :fim
if errorlevel 2 (
    start "" explorer "%CD%\dist"
    goto :fim
)
if errorlevel 1 (
    start "" msiexec.exe /i "%MSI_FILE%"
)

:fim
echo.
pause
exit /b 0
