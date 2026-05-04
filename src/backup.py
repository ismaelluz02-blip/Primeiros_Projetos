"""Backups locais versionados dos dados operacionais do sistema."""

import json
import os
import shutil
from datetime import datetime, timedelta

import src.config as config


def _agora():
    return datetime.now()


def _manifest_path(pasta_backup):
    return os.path.join(pasta_backup, "manifest.json")


def _ler_manifest(pasta_backup):
    caminho = _manifest_path(pasta_backup)
    if not os.path.exists(caminho):
        return None
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def listar_backups(limite=None):
    if not os.path.isdir(config.BACKUP_DIR):
        return []
    registros = []
    for nome in os.listdir(config.BACKUP_DIR):
        pasta = os.path.join(config.BACKUP_DIR, nome)
        if not os.path.isdir(pasta):
            continue
        manifest = _ler_manifest(pasta) or {}
        registros.append(
            {
                "nome": nome,
                "pasta": pasta,
                "criado_em": manifest.get("criado_em") or "",
                "acionado_por": manifest.get("acionado_por") or "",
                "arquivos": manifest.get("arquivos") or [],
            }
        )
    registros.sort(key=lambda item: item.get("criado_em") or item.get("nome") or "", reverse=True)
    if limite is not None:
        return registros[: int(limite)]
    return registros


def ultimo_backup():
    backups = listar_backups(limite=1)
    return backups[0] if backups else None


def _copiar_se_existir(origem, destino_dir, nome_destino):
    if not origem or not os.path.exists(origem):
        return None
    os.makedirs(destino_dir, exist_ok=True)
    destino = os.path.join(destino_dir, nome_destino)
    shutil.copy2(origem, destino)
    return {"origem": origem, "arquivo": nome_destino, "tamanho": os.path.getsize(destino)}


def limpar_backups_antigos(manter=20):
    backups = listar_backups()
    removidos = 0
    for item in backups[int(manter) :]:
        try:
            shutil.rmtree(item["pasta"])
            removidos += 1
        except OSError:
            pass
    return removidos


def criar_backup_local(acionado_por="manual", manter=20):
    os.makedirs(config.BACKUP_DIR, exist_ok=True)
    criado_em = _agora().strftime("%Y-%m-%dT%H:%M:%S")
    slug_base = _agora().strftime("backup_%Y%m%d_%H%M%S")
    slug = slug_base
    pasta = os.path.join(config.BACKUP_DIR, slug)
    contador = 2
    while os.path.exists(pasta):
        slug = f"{slug_base}_{contador}"
        pasta = os.path.join(config.BACKUP_DIR, slug)
        contador += 1
    os.makedirs(pasta, exist_ok=True)

    arquivos = []
    fontes = [
        (config.DB_PATH, "faturamento.db"),
        (config.SYNC_STATE_PATH, "configuracoes_manuais.json"),
        (config.OPERATIONAL_STATE_PATH, "estado_operacional.json"),
    ]
    for origem, nome_destino in fontes:
        copiado = _copiar_se_existir(origem, pasta, nome_destino)
        if copiado:
            arquivos.append(copiado)

    manifest = {
        "schema_version": 1,
        "criado_em": criado_em,
        "acionado_por": acionado_por,
        "app_dir": config.APP_DIR,
        "app_data_dir": config.APP_DATA_DIR,
        "arquivos": arquivos,
    }
    with open(_manifest_path(pasta), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    removidos = limpar_backups_antigos(manter=manter)
    return {"pasta": pasta, "criado_em": criado_em, "arquivos": len(arquivos), "removidos": removidos}


def criar_backup_automatico_se_necessario(intervalo_horas=12, manter=20, acionado_por="automatico"):
    ultimo = ultimo_backup()
    if ultimo and ultimo.get("criado_em"):
        try:
            criado = datetime.strptime(ultimo["criado_em"], "%Y-%m-%dT%H:%M:%S")
            if _agora() - criado < timedelta(hours=float(intervalo_horas)):
                return {"criado": False, "motivo": "intervalo", "ultimo": ultimo}
        except (TypeError, ValueError):
            pass
    resultado = criar_backup_local(acionado_por=acionado_por, manter=manter)
    resultado["criado"] = True
    return resultado


def restaurar_backup(pasta_backup, criar_pre_restore=True):
    pasta = os.path.abspath(str(pasta_backup or ""))
    if not pasta or not os.path.isdir(pasta):
        raise ValueError("Backup nao encontrado.")

    manifest = _ler_manifest(pasta)
    if not manifest:
        raise ValueError("Backup invalido ou sem manifest.json.")

    arquivos_backup = {item.get("arquivo"): item for item in manifest.get("arquivos", []) if item.get("arquivo")}
    destinos = {
        "faturamento.db": config.DB_PATH,
        "configuracoes_manuais.json": config.SYNC_STATE_PATH,
        "estado_operacional.json": config.OPERATIONAL_STATE_PATH,
    }

    disponiveis = []
    for nome_arquivo, destino in destinos.items():
        origem = os.path.join(pasta, nome_arquivo)
        if os.path.exists(origem) and (nome_arquivo in arquivos_backup or os.path.isfile(origem)):
            disponiveis.append((origem, destino, nome_arquivo))

    if not disponiveis:
        raise ValueError("Backup sem arquivos restauraveis.")

    pre_restore = None
    if criar_pre_restore:
        pre_restore = criar_backup_local("antes_de_restaurar")

    restaurados = []
    for origem, destino, nome_arquivo in disponiveis:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        shutil.copy2(origem, destino)
        restaurados.append({"arquivo": nome_arquivo, "destino": destino})

    return {
        "pasta": pasta,
        "criado_em": manifest.get("criado_em") or "",
        "restaurados": restaurados,
        "pre_restore": pre_restore,
    }
