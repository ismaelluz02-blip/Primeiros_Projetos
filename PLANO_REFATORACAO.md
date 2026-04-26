# Plano de Refatoração — sistema_faturamento.py

## Situação atual

O sistema inteiro vive num único arquivo de 7.576 linhas. Isso funciona, mas traz problemas práticos:

- Difícil de navegar ("onde está a função que importa o Excel?")
- Difícil de passar para outros desenvolvedores
- Qualquer alteração mexe no mesmo arquivo, aumentando o risco de quebrar algo sem perceber
- Impossível testar partes isoladas do sistema

---

## O principal desafio: variáveis globais

O código usa ~30 variáveis globais que são lidas e modificadas por funções em partes diferentes do arquivo (ex: `dados_importados`, `relatorio_selecionado`, `UI_THEME`, `APP_DATA_DIR`, etc.). Antes de dividir o código em arquivos, é preciso organizar esse estado — senão os módulos ficariam dependendo uns dos outros de forma circular.

**Solução**: criar um módulo `state.py` que centraliza todo o estado mutável. Todos os outros módulos importam de lá. Simples e sem risco de circular import.

---

## Estrutura proposta

```
FATURAMENTO HORIZONTE/
├── sistema_faturamento.py      ← mantido como ponto de entrada (vira só 5 linhas)
└── src/
    ├── state.py                ← todas as variáveis globais mutáveis
    ├── config.py               ← constantes, caminhos, lock de instância
    ├── banco.py                ← SQLite: conexão, tabelas, configurações
    ├── utils.py                ← formatação de moeda, datas, normalização de texto
    ├── importacao.py           ← leitura de PDF e Excel do Visual Rodopar
    ├── documentos.py           ← salvar, cancelar, substituir, competência, frete
    ├── relatorios.py           ← filtros, geração de Excel, abrir relatório
    ├── sync.py                 ← exportar/importar backup JSON
    ├── dashboard.py            ← dados e gráficos do dashboard
    └── ui/
        ├── app.py              ← classe App e tela principal
        ├── tema.py             ← modo claro/escuro, cores, layout de botões
        ├── componentes.py      ← widgets reutilizáveis, animações, cards
        ├── dialogos.py         ← janelas modais (alterar competência, cancelar, etc.)
        └── telas.py            ← telas (dashboard, relatórios, faturamento, etc.)
```

O `sistema_faturamento.py` atual passaria a ter apenas:

```python
from src.ui.app import iniciar
if __name__ == "__main__":
    iniciar()
```

---

## Mapa de divisão por linha

| Módulo | Linhas atuais | Tamanho estimado | Responsabilidade |
|---|---|---|---|
| `config.py` | 1 – 251 | ~250 linhas | Caminhos, lock, inicialização de diretórios |
| `banco.py` | 252 – 512 | ~260 linhas | Conexão SQLite, criação de tabelas, configurações |
| `utils.py` | 513 – 747 | ~235 linhas | Moeda, datas, competência, normalização |
| `state.py` | (espalhado) | ~60 linhas | Todas as variáveis globais mutáveis |
| `importacao.py` | 748 – 2138 | ~1.390 linhas | PDF, Excel, parser de colunas |
| `sync.py` | 2139 – 2556 | ~418 linhas | Backup/restore JSON |
| `documentos.py` | 2557 – 3050 | ~494 linhas | CRUD de documentos |
| `relatorios.py` | 3051 – 3566 | ~516 linhas | Filtros, Excel, abrir relatório |
| `dashboard.py` | 4294 – 5366 | ~1.073 linhas | Dados filtrados, gráficos matplotlib |
| `ui/app.py` | 3567 – 3739 | ~173 linhas | Classe App, inicialização |
| `ui/dialogos.py` | 3757 – 4293 | ~536 linhas | Todas as janelas modais |
| `ui/tema.py` | 5367 – 5793 | ~427 linhas | Tema, layout, abas |
| `ui/componentes.py` | 6039 – 7044 | ~1.006 linhas | Widgets, animações, navegação |
| `ui/telas.py` | 7045 – 7576 | ~532 linhas | Telas individuais |

---

## Estratégia de execução (ordem segura)

A refatoração deve ser feita em fases, do mais simples ao mais complexo, testando o sistema a cada fase antes de continuar.

### Fase 1 — Funções puras (sem risco)
Extrair `utils.py` primeiro. São funções matemáticas e de formatação que não dependem de nenhum estado global. Se algo quebrar é imediatamente óbvio.

### Fase 2 — Banco de dados
Extrair `banco.py`. Tem dependência de `DB_PATH` (que virá do `config.py`) mas nenhum estado de UI.

### Fase 3 — Config e State
Criar `config.py` com as constantes e `state.py` com as variáveis mutáveis. Esta é a fase mais delicada porque todos os outros módulos vão depender deles — mas é também o que desbloqueia o resto.

### Fase 4 — Lógica de negócio
Extrair `documentos.py`, `sync.py` e `relatorios.py`. Dependem de `banco.py`, `utils.py` e `state.py`, mas não da UI.

### Fase 5 — Importação
Extrair `importacao.py`. É o maior bloco de lógica pura (~1.400 linhas de parsing de PDF/Excel).

### Fase 6 — Dashboard
Extrair `dashboard.py`. Separa a lógica de dados da apresentação gráfica.

### Fase 7 — Interface (UI)
Extrair os módulos de UI por último, quando toda a lógica já estiver estável.

---

## Regras para não quebrar nada

1. **Uma fase de cada vez.** Testar o sistema completo antes de avançar.
2. **Não mudar comportamento**, só mover código.
3. **Git commit a cada fase** concluída — facilita reverter se algo quebrar.
4. **O `sistema_faturamento.py` original** fica intocado até a Fase 7 estar 100% validada.

---

## O que NÃO fazer nessa refatoração

- Não renomear funções
- Não mudar a lógica de nenhuma função
- Não alterar o banco de dados
- Não mexer na interface visível ao usuário

Tudo isso pode vir depois, em refatorações menores e mais focadas.
