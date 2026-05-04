import src.config as config
from src.banco import iniciar_banco
from src.seguros import (
    adicionar_seguro,
    atualizar_observacao_seguro,
    atualizar_status_seguro,
    inativar_seguro,
    listar_controle_competencia,
    resumo_competencia,
)


def _preparar_banco_temporario(monkeypatch, tmp_path):
    db_path = tmp_path / "seguros.db"
    monkeypatch.setattr(config, "DB_PATH", str(db_path))
    iniciar_banco()
    return db_path


def test_seguro_aparece_pendente_por_competencia(monkeypatch, tmp_path):
    _preparar_banco_temporario(monkeypatch, tmp_path)
    adicionar_seguro("Fator")

    registros = listar_controle_competencia(1, 2026)

    assert len(registros) == 1
    assert registros[0]["nome"] == "Fator"
    assert registros[0]["status"] == "PENDENTE"


def test_status_e_observacao_sao_individuais_por_competencia(monkeypatch, tmp_path):
    _preparar_banco_temporario(monkeypatch, tmp_path)
    adicionar_seguro("HDI")
    seguro_id = listar_controle_competencia(1, 2026)[0]["seguro_id"]

    atualizar_status_seguro(seguro_id, 1, 2026, "ENVIADO")
    atualizar_observacao_seguro(seguro_id, 1, 2026, "Enviado por e-mail dia 10")

    janeiro = listar_controle_competencia(1, 2026)[0]
    fevereiro = listar_controle_competencia(2, 2026)[0]

    assert janeiro["status"] == "ENVIADO"
    assert janeiro["observacao"] == "Enviado por e-mail dia 10"
    assert fevereiro["status"] == "PENDENTE"
    assert fevereiro["observacao"] == ""


def test_inativar_seguro_nao_apaga_historico_registrado(monkeypatch, tmp_path):
    _preparar_banco_temporario(monkeypatch, tmp_path)
    adicionar_seguro("Fator")
    seguro_id = listar_controle_competencia(3, 2026)[0]["seguro_id"]
    atualizar_status_seguro(seguro_id, 3, 2026, "RECEBIDO")

    inativar_seguro(seguro_id)

    historico_marco = listar_controle_competencia(3, 2026)
    nova_competencia = listar_controle_competencia(4, 2026)
    resumo = resumo_competencia(3, 2026)

    assert historico_marco[0]["nome"] == "Fator"
    assert historico_marco[0]["status"] == "RECEBIDO"
    assert nova_competencia == []
    assert resumo["recebido"] == 1
