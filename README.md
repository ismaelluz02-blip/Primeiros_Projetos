# Sistema de Faturamento — Horizonte Logística

Aplicação desktop para gestão e acompanhamento do faturamento mensal da Horizonte Logística. O sistema opera em cima de um modelo de **franquia mensal por cliente**: cada cliente tem um valor contratado para consumir no mês, e o sistema controla os documentos fiscais (NF-e, CT-e) que compõem esse consumo, permitindo ajustes, exportação de relatórios e acompanhamento por período de competência.

---

## Como funciona o fluxo

Os documentos fiscais são emitidos no **Visual Rodopar** (sistema externo). Desse sistema é exportado um relatório em PDF ou Excel, que é importado aqui. A partir daí o sistema organiza os documentos, permite fazer ajustes manuais e gera os relatórios consolidados do mês.

---

## Tipos de documento

**Documento normal** — nota fiscal ou CT-e de um serviço de transporte convencional dentro da franquia do cliente.

**Intercompany** — documento de transporte entre estados (entrada ou saída). Sinalizado separadamente por ter tratamento fiscal diferenciado.

**Delta** — documento emitido para cobrar a diferença quando o cliente não consume a franquia completa no mês. Por exemplo: franquia de R$ 300.000, cliente consumiu R$ 250.000 → é emitida uma nota delta de R$ 50.000 para fechar o pacote mensal.

**Spot** — documento de serviço avulso, pago fora do valor da franquia. Não entra no cálculo do consumo mensal contratado.

---

## Competência manual

O período de competência de um documento nem sempre coincide com o mês da data de emissão. O caso mais comum é a **nota delta**: ela só pode ser emitida depois que o mês fecha, ou seja, no primeiro dia útil do mês seguinte. Para que esse documento seja contabilizado no mês ao qual pertence (e não no mês em que foi emitido), a competência é ajustada manualmente dentro do sistema.

---

## Funcionalidades

- **Importação de relatórios** — lê PDFs e planilhas Excel exportados do Visual Rodopar e salva os documentos no banco local.
- **Dashboard** — exibe totais do período, gráficos de faturamento e comparativo por tipo de documento.
- **Gestão de documentos** — cancela, substitui, altera competência e declara modalidade (intercompany, delta, spot) em cada documento.
- **Exportação** — gera planilha Excel filtrada por período, pronta para entrega ou arquivo.
- **Busca** — localiza documentos por número, tipo ou período.
- **Sincronização / backup** — exporta e importa toda a base em JSON, facilitando migração entre máquinas.
- **Temas** — suporte a modo claro e escuro via CustomTkinter.
- **Interface personalizável** — ordem das abas e layout dos botões configuráveis pelo usuário.
- **Instância única** — impede que o sistema abra duas vezes ao mesmo tempo via arquivo de lock.

---

## Tecnologias

| Camada | Biblioteca |
|---|---|
| Interface | `customtkinter` + `tkinter` |
| Banco de dados | SQLite (`sqlite3`) |
| PDFs | `PyMuPDF` (fitz) |
| Planilhas | `pandas` + `openpyxl` |
| Gráficos | `matplotlib` |
| Imagens | `Pillow` |
| Build (.exe) | `cx_Freeze` |

Requer **Python 3.12**.

---

## Instalação (modo desenvolvimento)

```bash
# 1. Clone ou copie o projeto
cd "FATURAMENTO HORIZONTE"

# 2. Crie e ative um ambiente virtual (recomendado)
python -m venv .venv
.venv\Scripts\activate

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Execute
python sistema_faturamento.py
```

> O banco de dados (`faturamento.db`) é criado automaticamente no primeiro uso em `%LOCALAPPDATA%\Horizonte Logistica\Sistema de Faturamento\`. Se esse caminho não tiver permissão de escrita, o sistema usa a pasta `_dados_app\` local como fallback.

---

## Gerar o executável (.exe)

```bash
build.bat
```

O executável é gerado em `build\exe.win-amd64-3.12\`. O instalador `.msi` pode ser gerado com:

```bash
gerar_instalador.bat
```

---

## Estrutura do projeto

```
FATURAMENTO HORIZONTE/
│
├── sistema_faturamento.py     # Código principal (único arquivo, ~7.500 linhas)
├── requirements.txt           # Dependências Python
├── setup.py                   # Configuração do cx_Freeze para build
├── build.bat                  # Atalho para gerar o .exe
├── gerar_instalador.bat       # Gera o instalador .msi
├── logo.png / logo.ico        # Identidade visual
│
├── _dados_app/                # Dados em runtime (banco SQLite, cache)
│   └── faturamento.db
│
├── RELATORIOS/                # Relatórios Excel de saída
│
├── build/                     # Saída do cx_Freeze (gerado, não versionar)
└── dist/                      # Instalador .msi (gerado, não versionar)
```

---

## Observações

- O `.gitignore` já exclui `build/`, `dist/`, `__pycache__/` e `_dados_app/`.
- A pasta `_tmp/` é usada internamente pelo sistema em runtime; não precisa ser versionada.
- O arquivo `faturamento.db` contém dados reais — **não versionar**.
