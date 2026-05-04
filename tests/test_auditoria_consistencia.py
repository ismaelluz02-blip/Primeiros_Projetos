import src.config as config
from src.auditoria_consistencia import auditar_consistencia
from src.banco import iniciar_banco, obter_conexao_banco


def test_auditoria_consistencia_aponta_problemas_basicos(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "consistencia.db"))
    iniciar_banco()
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO documentos (numero, numero_original, tipo, competencia, frete, status, valor_inicial, valor_final)
        VALUES (10, 'NF 10', 'NF', '', '', 'OK', 0, 0)
        """
    )
    cursor.execute(
        """
        INSERT INTO documentos (numero, numero_original, tipo, competencia, frete, status, valor_inicial, valor_final, cancelado_manual)
        VALUES (11, 'CTE 11', 'CTE', 'maio/2026', 'FRANQUIA', 'CANCELADO', 100, 100, 1)
        """
    )
    conn.commit()
    conn.close()

    resultado = auditar_consistencia()
    codigos = {item["codigo"] for item in resultado["problemas"]}

    assert resultado["resumo"]["total"] >= 3
    assert "SEM_COMPETENCIA" in codigos
    assert "FRETE_VAZIO" in codigos
    assert "CANCELADO_COM_VALOR" in codigos
