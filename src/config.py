"""
Configuração de ambiente — caminhos, constantes e inicialização de diretórios.

Este módulo é importado cedo no ciclo de vida da aplicação.
Ao ser importado, chama automaticamente configurar_diretorio_dados()
para garantir que DB_PATH aponte para um diretório gravável.
"""

import os
import shutil
import sys
import time

# ─────────────────────────────────────────────
#  Caminhos base
# ─────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__ if not getattr(sys, "frozen", False) else sys.executable))
# Quando empacotado com cx_Freeze, __file__ aponta para dentro do .zip;
# usamos sys.executable como referência de diretório.
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
    APP_DIR = SCRIPT_DIR
else:
    SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    APP_DIR = SCRIPT_DIR

BASE_DIR = SCRIPT_DIR
PROJECT_DIR = APP_DIR if getattr(sys, "frozen", False) else os.path.dirname(BASE_DIR)
RELATORIOS_DIR = os.path.join(PROJECT_DIR, "RELATORIOS")

DEFAULT_APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", BASE_DIR),
    "Horizonte Logistica",
    "Sistema de Faturamento",
)
FALLBACK_APP_DATA_DIR = os.path.join(BASE_DIR, "_dados_app")
LEGACY_DB_PATH = os.path.join(APP_DIR, "faturamento.db")
LOGO_PATH = os.path.join(APP_DIR, "logo.png")
APP_USER_MODEL_ID = "horizonte.logistica.sistema.faturamento"

# ─────────────────────────────────────────────
#  Variáveis de caminho mutáveis (atualizadas por configurar_diretorio_dados)
# ─────────────────────────────────────────────

APP_DATA_DIR = DEFAULT_APP_DATA_DIR
DB_PATH = os.path.join(APP_DATA_DIR, "faturamento.db")
LOCK_PATH = os.path.join(APP_DATA_DIR, ".sistema_faturamento.lock")

# ─────────────────────────────────────────────
#  Constantes de importação
# ─────────────────────────────────────────────

# Código de filial aceito na importação de relatórios (planilha/PDF).
# Linhas com filial diferente deste valor são ignoradas.
FILIAL_PADRAO = "88"


# ─────────────────────────────────────────────
#  Constantes de sincronização
# ─────────────────────────────────────────────

SYNC_CONFIG_SCHEMA_VERSION = 1
SYNC_STATE_DIR = os.path.join(APP_DIR, "sync_state")
SYNC_STATE_PATH = os.path.join(SYNC_STATE_DIR, "configuracoes_manuais.json")
SYNC_DOCUMENT_FIELDS = [
    "tipo",
    "numero",
    "numero_original",
    "data_emissao",
    "valor_inicial",
    "valor_final",
    "frete",
    "status",
    "competencia",
    "valor_inicial_original",
    "valor_final_original",
    "status_original",
    "cancelado_manual",
    "competencia_manual",
    "frete_manual",
    "frete_revisado_manual",
]


# ─────────────────────────────────────────────
#  Helpers de diretório
# ─────────────────────────────────────────────

def _diretorio_gravavel(caminho_dir):
    try:
        os.makedirs(caminho_dir, exist_ok=True)
        arquivo_teste = os.path.join(caminho_dir, f".write_test_{os.getpid()}_{time.time_ns()}.tmp")
        with open(arquivo_teste, "w", encoding="utf-8") as f:
            f.write("ok")
        try:
            os.remove(arquivo_teste)
        except OSError:
            pass
        return True
    except OSError:
        return False


def configurar_diretorio_dados():
    """
    Escolhe o diretório de dados gravável e atualiza APP_DATA_DIR,
    DB_PATH e LOCK_PATH como atributos deste módulo.
    Ordem de preferência: AppData → _dados_app → BASE_DIR.
    """
    global APP_DATA_DIR, DB_PATH, LOCK_PATH

    dir_original = APP_DATA_DIR
    dir_escolhido = None
    for candidato in (APP_DATA_DIR, FALLBACK_APP_DATA_DIR, BASE_DIR):
        if _diretorio_gravavel(candidato):
            dir_escolhido = candidato
            break

    if not dir_escolhido:
        dir_escolhido = BASE_DIR

    APP_DATA_DIR = dir_escolhido
    DB_PATH = os.path.join(APP_DATA_DIR, "faturamento.db")
    LOCK_PATH = os.path.join(APP_DATA_DIR, ".sistema_faturamento.lock")

    # Migra banco se a pasta padrão não estava acessível na última execução.
    if APP_DATA_DIR != dir_original:
        origem_db = os.path.join(dir_original, "faturamento.db")
        if os.path.exists(origem_db) and not os.path.exists(DB_PATH):
            try:
                shutil.copy2(origem_db, DB_PATH)
            except OSError:
                pass


def configurar_cache_matplotlib():
    cache_dir = os.path.join(APP_DATA_DIR, "matplotlib")
    if _diretorio_gravavel(cache_dir):
        os.environ["MPLCONFIGDIR"] = cache_dir


# Executa ao importar o modulo, garantindo que DB_PATH seja sempre valido.
configurar_diretorio_dados()
configurar_cache_matplotlib()
