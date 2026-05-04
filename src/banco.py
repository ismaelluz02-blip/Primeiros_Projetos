"""
Camada de banco de dados — SQLite.

Responsável por: validação e recuperação do banco, conexão,
criação/migração de tabelas e leitura/gravação de configurações.
Depende de src.config para os caminhos (DB_PATH, etc.).
"""

import os
import shutil
import sqlite3

import src.config as config
from src.utils import _normalizar_numero_original_nf


# ─────────────────────────────────────────────
#  Validação e recuperação
# ─────────────────────────────────────────────

def _sqlite_db_valido(caminho_db):
    if not caminho_db or not os.path.exists(caminho_db):
        return True
    try:
        conn = sqlite3.connect(caminho_db)
        cursor = conn.cursor()
        cursor.execute("PRAGMA quick_check")
        row = cursor.fetchone()
        conn.close()
        return bool(row) and str(row[0]).strip().lower() == "ok"
    except sqlite3.Error:
        return False


def _candidatos_recuperacao_banco():
    caminhos = []
    vistos = set()
    for caminho in (
        os.path.join(config.DEFAULT_APP_DATA_DIR, "faturamento.db"),
        os.path.join(config.FALLBACK_APP_DATA_DIR, "faturamento.db"),
        config.LEGACY_DB_PATH,
        os.path.join(config.BASE_DIR, "faturamento.db"),
    ):
        caminho_abs = os.path.abspath(caminho)
        if caminho_abs in vistos or caminho_abs == os.path.abspath(config.DB_PATH):
            continue
        vistos.add(caminho_abs)
        caminhos.append(caminho_abs)
    return caminhos


def _tentar_recuperar_banco():
    journal_path = f"{config.DB_PATH}-journal"
    try:
        if os.path.exists(journal_path):
            os.remove(journal_path)
    except OSError:
        pass

    if _sqlite_db_valido(config.DB_PATH):
        return True

    for candidato in _candidatos_recuperacao_banco():
        if not os.path.exists(candidato):
            continue
        if not _sqlite_db_valido(candidato):
            continue
        try:
            os.makedirs(os.path.dirname(config.DB_PATH), exist_ok=True)
            if os.path.exists(config.DB_PATH):
                backup_path = f"{config.DB_PATH}.corrompido.bak"
                try:
                    if os.path.exists(backup_path):
                        os.remove(backup_path)
                except OSError:
                    pass
                try:
                    os.replace(config.DB_PATH, backup_path)
                except OSError:
                    pass
            shutil.copy2(candidato, config.DB_PATH)
            try:
                if os.path.exists(journal_path):
                    os.remove(journal_path)
            except OSError:
                pass
            if _sqlite_db_valido(config.DB_PATH):
                return True
        except OSError:
            continue

    if os.path.exists(config.DB_PATH):
        backup_path = f"{config.DB_PATH}.corrompido.bak"
        try:
            if os.path.exists(backup_path):
                os.remove(backup_path)
        except OSError:
            pass
        try:
            os.replace(config.DB_PATH, backup_path)
        except OSError:
            pass
    try:
        if os.path.exists(journal_path):
            os.remove(journal_path)
    except OSError:
        pass
    return True


# ─────────────────────────────────────────────
#  Conexão
# ─────────────────────────────────────────────

def obter_conexao_banco():
    conn = sqlite3.connect(config.DB_PATH)
    try:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=MEMORY")
        cursor.execute("PRAGMA temp_store=MEMORY")
    except sqlite3.Error:
        pass
    return conn


# ─────────────────────────────────────────────
#  Configurações (chave/valor)
# ─────────────────────────────────────────────

def obter_configuracao(chave, padrao=""):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave=?", (chave,))
    row = cursor.fetchone()
    conn.close()
    if row and row[0] is not None:
        return str(row[0])
    return padrao


def salvar_configuracao(chave, valor):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO configuracoes (chave, valor)
        VALUES (?, ?)
        ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor
        """,
        (chave, "" if valor is None else str(valor)),
    )
    conn.commit()
    conn.close()


# ─────────────────────────────────────────────
#  Inicialização e migração
# ─────────────────────────────────────────────

def iniciar_banco():
    ultima_exc = None
    for tentativa in range(2):
        conn = None
        try:
            conn = obter_conexao_banco()
            cursor = conn.cursor()

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS documentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero INTEGER,
                    numero_original TEXT,
                    tipo TEXT,
                    data_emissao TEXT,
                    valor_inicial REAL,
                    valor_final REAL,
                    frete TEXT,
                    status TEXT,
                    competencia TEXT
                )
                """
            )

            cursor.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS idx_documentos_numero_tipo
                ON documentos (numero, tipo)
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS configuracoes (
                    chave TEXT PRIMARY KEY,
                    valor TEXT
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS historico_alteracoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_hora TEXT NOT NULL,
                    acao TEXT NOT NULL,
                    tipo TEXT,
                    numero INTEGER,
                    numero_original TEXT,
                    campo TEXT,
                    valor_anterior TEXT,
                    valor_novo TEXT,
                    usuario TEXT,
                    host TEXT
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS seguros (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    data_cadastro TEXT NOT NULL
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS seguro_controle_competencia (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    seguro_id INTEGER NOT NULL,
                    competencia_mes INTEGER NOT NULL,
                    competencia_ano INTEGER NOT NULL,
                    status TEXT NOT NULL DEFAULT 'PENDENTE',
                    observacao TEXT,
                    data_atualizacao TEXT NOT NULL,
                    UNIQUE(seguro_id, competencia_mes, competencia_ano),
                    FOREIGN KEY(seguro_id) REFERENCES seguros(id)
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS tarefas_categorias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    data_cadastro TEXT NOT NULL
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS tarefas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    titulo TEXT NOT NULL,
                    descricao TEXT,
                    categoria_id INTEGER,
                    responsavel TEXT,
                    data_criacao TEXT NOT NULL,
                    prazo TEXT,
                    prioridade TEXT NOT NULL DEFAULT 'MEDIA',
                    status TEXT NOT NULL DEFAULT 'A_FAZER',
                    ordem INTEGER NOT NULL DEFAULT 0,
                    data_atualizacao TEXT NOT NULL,
                    concluida_em TEXT,
                    tags TEXT,
                    excluida INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY(categoria_id) REFERENCES tarefas_categorias(id)
                )
                """
            )

            colunas_existentes = {linha[1] for linha in cursor.execute("PRAGMA table_info(documentos)").fetchall()}
            if "valor_inicial_original" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN valor_inicial_original REAL")
            if "valor_final_original" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN valor_final_original REAL")
            if "status_original" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN status_original TEXT")
            if "cancelado_manual" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN cancelado_manual INTEGER DEFAULT 0")
            if "competencia_manual" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN competencia_manual INTEGER DEFAULT 0")
            if "frete_manual" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN frete_manual INTEGER DEFAULT 0")
            if "frete_revisado_manual" not in colunas_existentes:
                cursor.execute("ALTER TABLE documentos ADD COLUMN frete_revisado_manual INTEGER DEFAULT 0")

            cursor.execute(
                """
                SELECT id, numero, numero_original, data_emissao
                FROM documentos
                WHERE UPPER(COALESCE(tipo, ''))='NF'
                """
            )
            ajustes_numero_original = []
            for doc_id, numero_doc, numero_original_doc, data_emissao_doc in cursor.fetchall():
                numero_corrigido = _normalizar_numero_original_nf(
                    numero_doc,
                    numero_original_doc,
                    data_emissao_doc,
                )
                numero_atual = str(numero_original_doc or "").strip()
                if numero_corrigido and numero_corrigido != numero_atual:
                    ajustes_numero_original.append((numero_corrigido, int(doc_id)))

            if ajustes_numero_original:
                cursor.executemany(
                    "UPDATE documentos SET numero_original=? WHERE id=?",
                    ajustes_numero_original,
                )

            conn.commit()
            conn.close()
            return
        except sqlite3.Error as exc:
            ultima_exc = exc
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass
            if tentativa == 0 and _tentar_recuperar_banco():
                continue
            raise

    if ultima_exc:
        raise ultima_exc
