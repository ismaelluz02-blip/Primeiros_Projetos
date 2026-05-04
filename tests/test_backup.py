import os

import src.config as config
from src.backup import criar_backup_automatico_se_necessario, criar_backup_local, listar_backups, restaurar_backup


def test_criar_backup_local_copia_arquivos_principais(monkeypatch, tmp_path):
    db = tmp_path / "faturamento.db"
    sync_dir = tmp_path / "sync_state"
    backup_dir = tmp_path / "backups"
    sync_dir.mkdir()
    db.write_text("db", encoding="utf-8")
    sync = sync_dir / "configuracoes_manuais.json"
    operacional = sync_dir / "estado_operacional.json"
    sync.write_text("{}", encoding="utf-8")
    operacional.write_text("{}", encoding="utf-8")

    monkeypatch.setattr(config, "DB_PATH", str(db))
    monkeypatch.setattr(config, "SYNC_STATE_PATH", str(sync))
    monkeypatch.setattr(config, "OPERATIONAL_STATE_PATH", str(operacional))
    monkeypatch.setattr(config, "BACKUP_DIR", str(backup_dir))

    resultado = criar_backup_local("teste")
    backups = listar_backups()

    assert resultado["arquivos"] == 3
    assert len(backups) == 1
    assert os.path.exists(os.path.join(resultado["pasta"], "faturamento.db"))
    assert os.path.exists(os.path.join(resultado["pasta"], "configuracoes_manuais.json"))
    assert os.path.exists(os.path.join(resultado["pasta"], "estado_operacional.json"))


def test_backup_automatico_respeita_intervalo(monkeypatch, tmp_path):
    db = tmp_path / "faturamento.db"
    db.write_text("db", encoding="utf-8")
    monkeypatch.setattr(config, "DB_PATH", str(db))
    monkeypatch.setattr(config, "SYNC_STATE_PATH", str(tmp_path / "ausente_config.json"))
    monkeypatch.setattr(config, "OPERATIONAL_STATE_PATH", str(tmp_path / "ausente_operacional.json"))
    monkeypatch.setattr(config, "BACKUP_DIR", str(tmp_path / "backups"))

    primeiro = criar_backup_automatico_se_necessario(intervalo_horas=12)
    segundo = criar_backup_automatico_se_necessario(intervalo_horas=12)

    assert primeiro["criado"] is True
    assert segundo["criado"] is False
    assert len(listar_backups()) == 1


def test_restaurar_backup_recoloca_arquivos(monkeypatch, tmp_path):
    db = tmp_path / "faturamento.db"
    sync = tmp_path / "configuracoes_manuais.json"
    operacional = tmp_path / "estado_operacional.json"
    backup_dir = tmp_path / "backups"
    db.write_text("versao antiga", encoding="utf-8")
    sync.write_text('{"a": 1}', encoding="utf-8")
    operacional.write_text('{"b": 1}', encoding="utf-8")
    monkeypatch.setattr(config, "DB_PATH", str(db))
    monkeypatch.setattr(config, "SYNC_STATE_PATH", str(sync))
    monkeypatch.setattr(config, "OPERATIONAL_STATE_PATH", str(operacional))
    monkeypatch.setattr(config, "BACKUP_DIR", str(backup_dir))

    backup = criar_backup_local("teste")
    db.write_text("versao quebrada", encoding="utf-8")
    sync.write_text('{"a": 99}', encoding="utf-8")
    operacional.write_text('{"b": 99}', encoding="utf-8")

    resultado = restaurar_backup(backup["pasta"])

    assert db.read_text(encoding="utf-8") == "versao antiga"
    assert sync.read_text(encoding="utf-8") == '{"a": 1}'
    assert operacional.read_text(encoding="utf-8") == '{"b": 1}'
    assert len(resultado["restaurados"]) == 3
    assert resultado["pre_restore"]["arquivos"] == 3
