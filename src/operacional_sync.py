"""Sincronização versionada de Tarefas e Seguros em JSON local."""

import json
import os
import sqlite3
from datetime import datetime

import src.config as config
from src.banco import obter_conexao_banco


def _fetch_all(cursor, query, params=()):
    cursor.execute(query, params)
    return [dict(row) for row in cursor.fetchall()]


def exportar_estado_operacional(caminho_arquivo=None):
    caminho = caminho_arquivo or config.OPERATIONAL_STATE_PATH
    os.makedirs(os.path.dirname(caminho), exist_ok=True)

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        payload = {
            "metadata": {
                "schema_version": 1,
                "exportado_em": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
                "formato": "estado_operacional_seguros_tarefas",
            },
            "seguros": _fetch_all(cursor, "SELECT * FROM seguros ORDER BY id"),
            "seguro_controle_competencia": _fetch_all(cursor, "SELECT * FROM seguro_controle_competencia ORDER BY id"),
            "tarefas_categorias": _fetch_all(cursor, "SELECT * FROM tarefas_categorias ORDER BY id"),
            "tarefas": _fetch_all(cursor, "SELECT * FROM tarefas ORDER BY id"),
        }
    finally:
        conn.close()

    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return {
        "seguros": len(payload["seguros"]),
        "controles_seguro": len(payload["seguro_controle_competencia"]),
        "categorias": len(payload["tarefas_categorias"]),
        "tarefas": len(payload["tarefas"]),
    }


def importar_estado_operacional_se_existir(caminho_arquivo=None):
    caminho = caminho_arquivo or config.OPERATIONAL_STATE_PATH
    if not os.path.exists(caminho):
        return None

    with open(caminho, "r", encoding="utf-8") as f:
        payload = json.load(f)

    if not isinstance(payload, dict):
        raise ValueError("Estado operacional inválido.")

    conn = obter_conexao_banco()
    cursor = conn.cursor()
    resumo = {"seguros": 0, "controles_seguro": 0, "categorias": 0, "tarefas": 0}
    try:
        for item in payload.get("seguros", []):
            cursor.execute(
                """
                INSERT INTO seguros (id, nome, ativo, data_cadastro)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    nome=excluded.nome,
                    ativo=excluded.ativo,
                    data_cadastro=excluded.data_cadastro
                """,
                (item.get("id"), item.get("nome"), int(item.get("ativo", 1) or 0), item.get("data_cadastro") or ""),
            )
            resumo["seguros"] += 1

        for item in payload.get("seguro_controle_competencia", []):
            cursor.execute(
                """
                INSERT INTO seguro_controle_competencia
                (id, seguro_id, competencia_mes, competencia_ano, status, observacao, data_atualizacao)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(seguro_id, competencia_mes, competencia_ano) DO UPDATE SET
                    status=excluded.status,
                    observacao=excluded.observacao,
                    data_atualizacao=excluded.data_atualizacao
                """,
                (
                    item.get("id"),
                    item.get("seguro_id"),
                    item.get("competencia_mes"),
                    item.get("competencia_ano"),
                    item.get("status") or "PENDENTE",
                    item.get("observacao") or "",
                    item.get("data_atualizacao") or "",
                ),
            )
            resumo["controles_seguro"] += 1

        for item in payload.get("tarefas_categorias", []):
            cursor.execute(
                """
                INSERT INTO tarefas_categorias (id, nome, ativo, data_cadastro)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    nome=excluded.nome,
                    ativo=excluded.ativo,
                    data_cadastro=excluded.data_cadastro
                """,
                (item.get("id"), item.get("nome"), int(item.get("ativo", 1) or 0), item.get("data_cadastro") or ""),
            )
            resumo["categorias"] += 1

        for item in payload.get("tarefas", []):
            cursor.execute(
                """
                INSERT INTO tarefas
                (
                    id, titulo, descricao, categoria_id, responsavel, data_criacao,
                    prazo, prioridade, status, ordem, data_atualizacao, concluida_em,
                    tags, excluida
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    titulo=excluded.titulo,
                    descricao=excluded.descricao,
                    categoria_id=excluded.categoria_id,
                    responsavel=excluded.responsavel,
                    prazo=excluded.prazo,
                    prioridade=excluded.prioridade,
                    status=excluded.status,
                    ordem=excluded.ordem,
                    data_atualizacao=excluded.data_atualizacao,
                    concluida_em=excluded.concluida_em,
                    tags=excluded.tags,
                    excluida=excluded.excluida
                """,
                (
                    item.get("id"),
                    item.get("titulo"),
                    item.get("descricao") or "",
                    item.get("categoria_id"),
                    item.get("responsavel") or "",
                    item.get("data_criacao") or "",
                    item.get("prazo") or "",
                    item.get("prioridade") or "MEDIA",
                    item.get("status") or "A_FAZER",
                    int(item.get("ordem", 0) or 0),
                    item.get("data_atualizacao") or "",
                    item.get("concluida_em"),
                    item.get("tags") or "",
                    int(item.get("excluida", 0) or 0),
                ),
            )
            resumo["tarefas"] += 1

        conn.commit()
        return resumo
    finally:
        conn.close()
