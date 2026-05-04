from datetime import date, timedelta

import src.config as config
from src.banco import iniciar_banco
from src.tarefas import (
    adicionar_categoria,
    classificar_prazo,
    criar_tarefa,
    excluir_tarefa,
    formatar_prazo_br,
    listar_categorias,
    listar_tarefas,
    mover_tarefa,
    resumo_tarefas,
    atualizar_tarefa,
)


def _preparar_banco(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "tarefas.db"))
    iniciar_banco()


def test_criar_tarefa_aparece_em_a_fazer_com_categoria(monkeypatch, tmp_path):
    _preparar_banco(monkeypatch, tmp_path)
    adicionar_categoria("Seguros")

    tarefa_id = criar_tarefa("Cobrar comprovante", categoria="Seguros", responsavel="Ismael", prioridade="URGENTE")
    tarefas = listar_tarefas()

    assert tarefa_id > 0
    assert tarefas[0]["titulo"] == "Cobrar comprovante"
    assert tarefas[0]["status"] == "A_FAZER"
    assert tarefas[0]["categoria"] == "Seguros"
    assert resumo_tarefas()["urgentes"] == 1


def test_mover_tarefa_conclui_e_permite_reabrir(monkeypatch, tmp_path):
    _preparar_banco(monkeypatch, tmp_path)
    tarefa_id = criar_tarefa("Enviar e-mail")

    mover_tarefa(tarefa_id, "CONCLUIDO")
    concluida = listar_tarefas("CONCLUIDO")[0]
    assert concluida["concluida_em"]

    mover_tarefa(tarefa_id, "EM_ANDAMENTO")
    reaberta = listar_tarefas("EM_ANDAMENTO")[0]
    assert reaberta["concluida_em"] is None


def test_prazo_e_busca(monkeypatch, tmp_path):
    _preparar_banco(monkeypatch, tmp_path)
    ontem = (date.today() - timedelta(days=1)).strftime("%d/%m/%Y")
    tarefa_id = criar_tarefa("Conferir seguro HDI", descricao="Cliente pendente", prazo=ontem, categoria="Seguros")

    tarefa = listar_tarefas("ATRASADAS", busca="hdi")[0]
    assert tarefa["id"] == tarefa_id
    assert classificar_prazo(tarefa["prazo"], tarefa["status"]) == "atrasada"
    assert formatar_prazo_br(tarefa["prazo"]) == ontem


def test_editar_e_excluir_tarefa(monkeypatch, tmp_path):
    _preparar_banco(monkeypatch, tmp_path)
    tarefa_id = criar_tarefa("Original", categoria="Outros")
    atualizar_tarefa(tarefa_id, titulo="Atualizada", categoria="Financeiro", status="AGUARDANDO")

    tarefa = listar_tarefas("AGUARDANDO", categoria="Financeiro")[0]
    assert tarefa["titulo"] == "Atualizada"

    excluir_tarefa(tarefa_id)
    assert listar_tarefas() == []
    assert "Financeiro" in [c["nome"] for c in listar_categorias()]
