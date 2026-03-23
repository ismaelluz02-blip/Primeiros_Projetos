# NOVO_PROJETO

App desktop para automatizar o fluxo de requisicao no Visual Rodopar com 2 monitores.

## Objetivo

- Manter o app de controle no monitor 2.
- Executar cliques no Visual Rodopar no monitor 1.
- Fluxo inicial:
  - `Materiais` -> `Movimentacao` -> `Requisicao de Compras` -> `Manutencao`
- Botao principal: `CRIAR REQUISICAO`.

## Modo sem coordenadas fixas

O app agora trabalha em **modo por imagem**:

- Voce seleciona o passo.
- Clica em `Treinar Imagem (3s)`.
- Posiciona o mouse no item do sistema.
- O app salva um template visual do alvo e depois encontra esse alvo automaticamente para clicar.

Se a imagem nao for encontrada, ele pode usar fallback em XY do ultimo treino.

## Como executar

1. Ativar ambiente virtual:
   - PowerShell: `.\\.venv\\Scripts\\Activate.ps1`
2. Instalar dependencias:
   - `pip install -r requirements.txt`
3. Rodar o app:
   - `python -m src.main`

## Como treinar os passos

1. Abra o Visual Rodopar no monitor 1.
2. Abra este app no monitor 2.
3. Clique em `Mover App para Monitor` se necessario.
4. Selecione um passo na tabela.
5. Clique em `Treinar Imagem (3s)`.
6. Durante a contagem, deixe o mouse no alvo correto no monitor 1.
7. Repita para os 4 passos.
8. Clique em `CRIAR REQUISICAO`.

## Configuracao salva

- `config/requisicao_config.json`
- Templates visuais: pasta `templates/`

## Observacoes

- Use o botao `Parar` para interromper a execucao.
- Caso a UI do sistema mude muito (tema, tamanho, escala), refaca o treino dos templates.
