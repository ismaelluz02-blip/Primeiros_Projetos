"""Busca global em documentos, tarefas, seguros e historico."""

import sqlite3

from src.banco import obter_conexao_banco


def _like(texto):
    return f"%{str(texto or '').strip().lower()}%"


def buscar_global(termo, limite_por_tipo=8):
    termo = str(termo or "").strip()
    if len(termo) < 2:
        return []

    like = _like(termo)
    limite = max(1, int(limite_por_tipo or 8))
    resultados = []
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            SELECT id, tipo, numero, numero_original, competencia, data_emissao,
                   valor_inicial, valor_final, frete, status
            FROM documentos
            WHERE LOWER(COALESCE(tipo, '')) LIKE ?
               OR LOWER(COALESCE(numero_original, CAST(numero AS TEXT), '')) LIKE ?
               OR LOWER(COALESCE(competencia, '')) LIKE ?
               OR LOWER(COALESCE(frete, '')) LIKE ?
               OR LOWER(COALESCE(status, '')) LIKE ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (like, like, like, like, like, limite),
        )
        for row in cursor.fetchall():
            numero = row["numero_original"] or row["numero"] or "-"
            resultados.append(
                {
                    "tipo": "Documento",
                    "titulo": f"{row['tipo'] or '-'} {numero}",
                    "detalhe": f"{row['competencia'] or '-'} | {row['frete'] or '-'} | {row['status'] or '-'}",
                    "origem": "documentos",
                    "id": row["id"],
                }
            )

        cursor.execute(
            """
            SELECT t.id, t.titulo, t.descricao, t.responsavel, t.prazo, t.prioridade,
                   t.status, COALESCE(c.nome, 'Sem categoria') AS categoria
            FROM tarefas t
            LEFT JOIN tarefas_categorias c ON c.id = t.categoria_id
            WHERE t.excluida=0
              AND (
                LOWER(t.titulo) LIKE ?
                OR LOWER(COALESCE(t.descricao, '')) LIKE ?
                OR LOWER(COALESCE(t.responsavel, '')) LIKE ?
                OR LOWER(COALESCE(c.nome, '')) LIKE ?
                OR LOWER(COALESCE(t.tags, '')) LIKE ?
              )
            ORDER BY t.data_atualizacao DESC, t.id DESC
            LIMIT ?
            """,
            (like, like, like, like, like, limite),
        )
        for row in cursor.fetchall():
            resultados.append(
                {
                    "tipo": "Tarefa",
                    "titulo": row["titulo"],
                    "detalhe": f"{row['categoria']} | {row['responsavel'] or 'Sem responsavel'} | {row['status']}",
                    "origem": "tarefas",
                    "id": row["id"],
                }
            )

        cursor.execute(
            """
            SELECT s.id, s.nome, s.ativo,
                   c.competencia_mes, c.competencia_ano, c.status, c.observacao
            FROM seguros s
            LEFT JOIN seguro_controle_competencia c ON c.seguro_id = s.id
            WHERE LOWER(s.nome) LIKE ?
               OR LOWER(COALESCE(c.status, '')) LIKE ?
               OR LOWER(COALESCE(c.observacao, '')) LIKE ?
            ORDER BY c.competencia_ano DESC, c.competencia_mes DESC, s.nome
            LIMIT ?
            """,
            (like, like, like, limite),
        )
        for row in cursor.fetchall():
            comp = "-"
            if row["competencia_mes"] and row["competencia_ano"]:
                comp = f"{int(row['competencia_mes']):02d}/{row['competencia_ano']}"
            resultados.append(
                {
                    "tipo": "Seguro",
                    "titulo": row["nome"],
                    "detalhe": f"{comp} | {row['status'] or 'Sem controle'} | {row['observacao'] or ''}".strip(),
                    "origem": "seguros",
                    "id": row["id"],
                }
            )

        cursor.execute(
            """
            SELECT id, data_hora, acao, tipo, numero, numero_original, campo, valor_anterior, valor_novo
            FROM historico_alteracoes
            WHERE LOWER(COALESCE(acao, '')) LIKE ?
               OR LOWER(COALESCE(tipo, '')) LIKE ?
               OR LOWER(COALESCE(numero_original, CAST(numero AS TEXT), '')) LIKE ?
               OR LOWER(COALESCE(campo, '')) LIKE ?
               OR LOWER(COALESCE(valor_anterior, '')) LIKE ?
               OR LOWER(COALESCE(valor_novo, '')) LIKE ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (like, like, like, like, like, like, limite),
        )
        for row in cursor.fetchall():
            numero = row["numero_original"] or row["numero"] or "-"
            resultados.append(
                {
                    "tipo": "Historico",
                    "titulo": f"{row['acao'] or '-'} | {row['tipo'] or '-'} {numero}",
                    "detalhe": f"{row['campo'] or '-'}: {row['valor_anterior'] or '-'} -> {row['valor_novo'] or '-'}",
                    "origem": "alteracoes",
                    "id": row["id"],
                }
            )
    finally:
        conn.close()
    return resultados
