# Gravador de Tela (Base)

Versao com foco em qualidade de video, som e controle por atalhos.

## Recursos
- Escolha do monitor: `monitor 1`, `monitor 2`, ..., ou `todos`.
- Opcao `Area personalizada` para selecionar exatamente a area de captura.
- Qualidade rapida: `Alta (60 FPS)`, `Boa (30 FPS)`, `Leve (20 FPS)`.
- Audio configuravel:
  - `Sem audio`
  - `Apenas microfone`
  - `Apenas som do sistema` (loopback real)
  - `Microfone + som do sistema` (mixagem automatica)
- Escolha do microfone e do dispositivo de audio desktop.
- Escolha da pasta de destino e nome do arquivo.
- Countdown de `3 segundos` antes de iniciar gravacao.
- App minimiza automaticamente ao iniciar gravacao.
- Borda vermelha no contorno da area/monitor em gravacao.
- Atalhos globais personalizados no formato `Ctrl+Shift+tecla`:
  - Iniciar
  - Pausar/Retomar
  - Encerrar
  - Mudo/Com som
- Botao `Tirar print` para salvar screenshot em PNG.
- Exportacao final em `MP4 (H.264 + AAC)`.

## Instalacao
```powershell
cd gravador
pip install -r requirements.txt
```

Se precisar corrigir ambiente ja instalado (ex.: conflito `numpy`/`pandas`), use:
```powershell
python -m pip install -r requirements.txt --upgrade --force-reinstall
```

## Execucao
```powershell
python Gravador_de_tela.py
```

## Observacoes
- Para capturar audio do desktop, o Windows precisa expor dispositivo loopback.
- Se o audio desktop nao aparecer, clique em `Atualizar dispositivos`.
- OneDrive bloqueado por regra: o app nao inicia gravacao se a pasta destino estiver no OneDrive.
- Pasta padrao de saida: `C:\Projetos\Gravacoes`.
- Esta e a base para a proxima etapa: registrar eventos de clique/teclado para gerar automacoes futuras.
