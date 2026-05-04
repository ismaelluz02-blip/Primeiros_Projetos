import src.config as config
from src.banco import iniciar_banco, obter_conexao_banco
from src.busca_global import buscar_global
from src.seguros import adicionar_seguro, atualizar_observacao_seguro, listar_controle_competencia
from src.tarefas import criar_tarefa


def _preparar(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "busca.db"))
    iniciar_banco()


def test_busca_global_encontra_documentos_tarefas_e_seguros(monkeypatch, tmp_path):
    _preparar(monkeypatch, tmp_path)
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO documentos (numero, numero_original, tipo, competencia, frete, status, valor_inicial, valor_final)
        VALUES (129, 'NF 129', 'NF', 'marco/2026', 'DELTA', 'OK', 10, 10)
        """
    )
    conn.commit()
    conn.close()

    criar_tarefa("Cobrar NF 129", categoria="Financeiro", responsavel="Ismael")
    adicionar_seguro("Seguro Delta")
    seguro_id = listar_controle_competencia(3, 2026)[0]["seguro_id"]
    atualizar_observacao_seguro(seguro_id, 3, 2026, "Relacionado a NF 129")

    resultados = buscar_global("129")
    tipos = {r["tipo"] for r in resultados}

    assert "Documento" in tipos
    assert "Tarefa" in tipos
    assert "Seguro" in tipos
