@echo off
setlocal

echo.
echo ======================================
echo   DESINSTALAR SISTEMA DE FATURAMENTO
echo ======================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$keys=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*');" ^
  "$app=Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq 'Sistema de Faturamento' } | Select-Object -First 1;" ^
  "if(-not $app){ Write-Output 'APP_NAO_ENCONTRADO'; exit 2 }" ^
  "$u=$app.UninstallString;" ^
  "if(-not $u){ Write-Output 'SEM_UNINSTALLSTRING'; exit 3 }" ^
  "if($u -match 'MsiExec\.exe\s*/I\{([A-F0-9\-]+)\}'){ $cmd='msiexec.exe /x {' + $Matches[1] + '}'; } else { $cmd=$u };" ^
  "Write-Output ('Executando: ' + $cmd);" ^
  "Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', $cmd -Wait; exit $LASTEXITCODE"

if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel desinstalar automaticamente.
    echo Abra Configuracoes ^> Aplicativos ^> Aplicativos instalados
    echo e remova "Sistema de Faturamento" manualmente.
    pause
    exit /b 1
)

echo.
echo Sistema desinstalado com sucesso.
pause
exit /b 0
