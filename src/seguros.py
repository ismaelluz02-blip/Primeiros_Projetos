"""
Controle mensal de comprovantes de seguro.

As funções deste módulo ficam isoladas da UI para facilitar futura migração
para sincronização em nuvem ou outra fonte de persistência.
"""

from datetime import datetime
import sqlite3

from src.banco import obter_conexao_banco


STATUS_SEGURO = ("PENDENTE", "RECEBIDO", "ENVIADO")


def _agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalizar_nome(nome):
    return " ".join(str(nome or "").strip().split())


def _normalizar_status(status):
    status_norm = str(status or "").upper().strip()
    if status_norm not in STATUS_SEGURO:
        return "PENDENTE"
    return status_norm


def adicionar_seguro(nome):
    nome_norm = _normalizar_nome(nome)
    if not nome_norm:
        raise ValueError("Informe o nome do seguro.")

    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            INSERT INTO seguros (nome, ativo, data_cadastro)
            VALUES (?, 1, ?)
            ON CONFLICT(nome) DO UPDATE SET ativo=1
            """,
            (nome_norm, _agora()),
        )
        conn.commit()
    finally:
        conn.close()
    return nome_norm


def inativar_seguro(seguro_id):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE seguros SET ativo=0 WHERE id=?", (int(seguro_id),))
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def listar_seguros():
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, nome, ativo, data_cadastro FROM seguros ORDER BY ativo DESC, nome COLLATE NOCASE")
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def _garantir_registro_controle(cursor, seguro_id, mes, ano):
    cursor.execute(
        """
        INSERT OR IGNORE INTO seguro_controle_competencia
        (seguro_id, competencia_mes, competencia_ano, status, observacao, data_atualizacao)
        VALUES (?, ?, ?, 'PENDENTE', '', ?)
        """,
        (int(seguro_id), int(mes), int(ano), _agora()),
    )


def listar_controle_competencia(mes, ano, filtro_status="TODOS"):
    mes = int(mes)
    ano = int(ano)
    filtro_status = str(filtro_status or "TODOS").upper().strip()

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id FROM seguros WHERE ativo=1")
        for row in cursor.fetchall():
            _garantir_registro_controle(cursor, int(row["id"]), mes, ano)
        conn.commit()

        params = [mes, ano]
        filtro_sql = ""
        if filtro_status in STATUS_SEGURO:
            filtro_sql = "AND c.status=?"
            params.append(filtro_status)

        cursor.execute(
            f"""
            SELECT
                s.id AS seguro_id,
                s.nome,
                s.ativo,
                c.id AS controle_id,
                c.competencia_mes,
                c.competencia_ano,
                c.status,
                COALESCE(c.observacao, '') AS observacao,
                c.data_atualizacao
            FROM seguro_controle_competencia c
            JOIN seguros s ON s.id = c.seguro_id
            WHERE c.competencia_mes=? AND c.competencia_ano=?
              AND (s.ativo=1 OR c.status <> 'PENDENTE' OR COALESCE(c.observacao, '') <> '')
              {filtro_sql}
            ORDER BY s.nome COLLATE NOCASE
            """,
            params,
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def resumo_competencia(mes, ano):
    registros = listar_controle_competencia(mes, ano, "TODOS")
    resumo = {"total": len(registros), "pendente": 0, "recebido": 0, "enviado": 0}
    for item in registros:
        status = _normalizar_status(item.get("status"))
        if status == "PENDENTE":
            resumo["pendente"] += 1
        elif status == "RECEBIDO":
            resumo["recebido"] += 1
        elif status == "ENVIADO":
            resumo["enviado"] += 1
    return resumo


def atualizar_status_seguro(seguro_id, mes, ano, status):
    status_norm = _normalizar_status(status)
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        _garantir_registro_controle(cursor, int(seguro_id), int(mes), int(ano))
        cursor.execute(
            """
            UPDATE seguro_controle_competencia
            SET status=?, data_atualizacao=?
            WHERE seguro_id=? AND competencia_mes=? AND competencia_ano=?
            """,
            (status_norm, _agora(), int(seguro_id), int(mes), int(ano)),
        )
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def atualizar_observacao_seguro(seguro_id, mes, ano, observacao):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        _garantir_registro_controle(cursor, int(seguro_id), int(mes), int(ano))
        cursor.execute(
            """
            UPDATE seguro_controle_competencia
            SET observacao=?, data_atualizacao=?
            WHERE seguro_id=? AND competencia_mes=? AND competencia_ano=?
            """,
            (str(observacao or "").strip(), _agora(), int(seguro_id), int(mes), int(ano)),
        )
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()
