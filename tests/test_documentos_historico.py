import sqlite3

import src.config as config
import src.documentos as documentos
from src.banco import iniciar_banco, obter_conexao_banco
from src.documentos import declarar_delta, salvar_alteracao_frete_manual


def _preparar_banco_temporario(monkeypatch, tmp_path):
    db_path = tmp_path / "faturamento.db"
    monkeypatch.setattr(config, "DB_PATH", str(db_path))
    documentos._on_change_callbacks.clear()
    iniciar_banco()
    return db_path


def _inserir_nf_base():
    conn = obter_conexao_banco()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO documentos
        (
            numero, numero_original, tipo, data_emissao,
            valor_inicial, valor_final, frete, status, competencia
        )
        VALUES (?,?,?,?,?,?,?,?,?)
        """,
        (20260129, "129", "NF", "02/04/2026", 62066.13, 58962.82, "FRANQUIA", "OK", "abril/2026"),
    )
    conn.commit()
    conn.close()


def test_declarar_delta_registra_historico_e_desfazer_limpa_marcacao(monkeypatch, tmp_path):
    _preparar_banco_temporario(monkeypatch, tmp_path)
    _inserir_nf_base()

    resultado = declarar_delta("NF", 20260129)
    assert resultado["encontrados"] == 1
    assert resultado["alterados"] == 1

    resultado_desfazer = salvar_alteracao_frete_manual("NF", 20260129, "FRANQUIA")
    assert resultado_desfazer["encontrados"] == 1
    assert resultado_desfazer["alterados"] == 1

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT frete, frete_manual, frete_revisado_manual FROM documentos WHERE numero=? AND tipo='NF'", (20260129,))
    doc = dict(cur.fetchone())
    assert doc == {"frete": "FRANQUIA", "frete_manual": 0, "frete_revisado_manual": 0}

    cur.execute("SELECT acao, campo, valor_anterior, valor_novo FROM historico_alteracoes ORDER BY id")
    historico = [dict(row) for row in cur.fetchall()]
    conn.close()

    assert historico[0] == {
        "acao": "DECLARAR_DELTA",
        "campo": "frete",
        "valor_anterior": "FRANQUIA",
        "valor_novo": "DELTA",
    }
    assert historico[1] == {
        "acao": "DESFAZER_FRETE",
        "campo": "frete",
        "valor_anterior": "DELTA",
        "valor_novo": "FRANQUIA",
    }
