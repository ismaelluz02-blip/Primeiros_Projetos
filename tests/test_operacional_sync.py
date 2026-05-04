import src.config as config
from src.banco import iniciar_banco
from src.operacional_sync import exportar_estado_operacional, importar_estado_operacional_se_existir
from src.seguros import adicionar_seguro, atualizar_status_seguro, listar_controle_competencia
from src.tarefas import criar_tarefa, listar_tarefas


def test_exporta_e_importa_estado_operacional(monkeypatch, tmp_path):
    origem_db = tmp_path / "origem.db"
    sync_path = tmp_path / "estado_operacional.json"
    monkeypatch.setattr(config, "DB_PATH", str(origem_db))
    iniciar_banco()

    adicionar_seguro("Fator")
    seguro_id = listar_controle_competencia(4, 2026)[0]["seguro_id"]
    atualizar_status_seguro(seguro_id, 4, 2026, "ENVIADO")
    criar_tarefa("Cobrar apolice", categoria="Seguros", responsavel="Ismael", prioridade="URGENTE")

    resumo_export = exportar_estado_operacional(str(sync_path))

    destino_db = tmp_path / "destino.db"
    monkeypatch.setattr(config, "DB_PATH", str(destino_db))
    iniciar_banco()
    resumo_import = importar_estado_operacional_se_existir(str(sync_path))

    seguros = listar_controle_competencia(4, 2026)
    tarefas = listar_tarefas("URGENTES")

    assert resumo_export["seguros"] == 1
    assert resumo_import["seguros"] == 1
    assert seguros[0]["nome"] == "Fator"
    assert seguros[0]["status"] == "ENVIADO"
    assert tarefas[0]["titulo"] == "Cobrar apolice"
