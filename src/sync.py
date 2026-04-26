"""
Sincronização offline — exportação/importação de configurações manuais via JSON.

Funções puras de I/O e transformação. Sem dependências de UI.
Os wrappers de UI (exportar_configuracoes_ui, importar_configuracoes_ui)
permanecem em sistema_faturamento.py.
"""

import json
import sqlite3
import socket
import getpass
from datetime import datetime

import pandas as pd

import src.config as config
from src.banco import obter_conexao_banco
from src.utils import _coletar_numero_original_para_match, competencia_por_data
from src.documentos import _normalizar_modalidade_frete, _buscar_documento_existente_sync


# ─────────────────────────────────────────────
#  Helpers de conversão
# ─────────────────────────────────────────────

def _to_float(valor, padrao=0.0):
    try:
        if valor is None or str(valor).strip() == "":
            return float(padrao)
        return float(valor)
    except (TypeError, ValueError):
        return float(padrao)


def _to_optional_float(valor):
    if valor is None:
        return None
    if isinstance(valor, str) and not valor.strip():
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _to_manual_flag(valor, padrao=0):
    try:
        return 1 if int(valor) == 1 else 0
    except (TypeError, ValueError):
        return int(padrao)


def _normalizar_data_emissao_sync(data_txt, padrao_txt):
    valor = str(data_txt or "").strip()
    if not valor:
        valor = str(padrao_txt or "").strip()
    if not valor:
        return datetime.now().strftime("%d/%m/%Y")

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(valor[:19], fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass

    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime().strftime("%d/%m/%Y")
    except Exception:
        pass

    return datetime.now().strftime("%d/%m/%Y")


# ─────────────────────────────────────────────
#  Leitura de alterações manuais
# ─────────────────────────────────────────────

def _documento_possui_alteracao_manual(row):
    status_upper = str(row.get("status", "") or "").upper()
    frete_upper = str(row.get("frete", "") or "").upper().strip()
    return (
        int(row.get("cancelado_manual", 0) or 0) == 1
        or int(row.get("competencia_manual", 0) or 0) == 1
        or int(row.get("frete_manual", 0) or 0) == 1
        or int(row.get("frete_revisado_manual", 0) or 0) == 1
        or frete_upper in {"INTERCOMPANY", "DELTA", "SPOT"}
        or bool(row.get("valor_inicial_original") is not None)
        or bool(row.get("valor_final_original") is not None)
        or bool((row.get("status_original") or "").strip())
        or "DOCUMENTO SUBSTITUIDO POR" in status_upper
        or "DOCUMENTO SUBSTITUINDO DOCUMENTO" in status_upper
    )


def _listar_documentos_alterados_para_sync():
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT
            tipo, numero, numero_original, data_emissao,
            valor_inicial, valor_final, frete, status, competencia,
            valor_inicial_original, valor_final_original, status_original,
            cancelado_manual, competencia_manual, frete_manual, frete_revisado_manual
        FROM documentos
        ORDER BY id ASC
        """
    )
    docs = []
    for row in cursor.fetchall():
        item = dict(row)
        if _documento_possui_alteracao_manual(item):
            item["tipo"] = str(item.get("tipo", "")).upper().strip()
            item["numero"] = int(item.get("numero") or 0)
            item["numero_original"] = str(item.get("numero_original", "") or "").strip()
            item["cancelado_manual"] = int(item.get("cancelado_manual") or 0)
            item["competencia_manual"] = int(item.get("competencia_manual") or 0)
            item["frete_manual"] = int(item.get("frete_manual") or 0)
            item["frete_revisado_manual"] = int(item.get("frete_revisado_manual") or 0)
            # Mantém payload enxuto, mas com campos suficientes para reproduzir alterações.
            item = {campo: item.get(campo) for campo in config.SYNC_DOCUMENT_FIELDS}
            docs.append(item)
    conn.close()
    return docs


# ─────────────────────────────────────────────
#  Exportação
# ─────────────────────────────────────────────

def exportar_configuracoes_json(caminho_arquivo):
    documentos = _listar_documentos_alterados_para_sync()
    payload = {
        "metadata": {
            "schema_version": config.SYNC_CONFIG_SCHEMA_VERSION,
            "exportado_em": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "formato": "configuracoes_sync_documentos_manuais",
            "origem": {
                "host": socket.gethostname(),
                "usuario": getpass.getuser(),
                "app_data_dir": config.APP_DATA_DIR,
            },
            "total_documentos": len(documentos),
        },
        "documentos": documentos,
    }

    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return len(documentos)


def _extrair_documentos_payload_sync(payload):
    if isinstance(payload, list):
        return payload, {"schema_version": 0}

    if not isinstance(payload, dict):
        raise ValueError("Formato inválido de arquivo de configurações.")

    metadata = payload.get("metadata", {})
    if metadata is None:
        metadata = {}
    if not isinstance(metadata, dict):
        raise ValueError("Campo 'metadata' inválido no arquivo de configurações.")

    documentos = payload.get("documentos", [])
    if not isinstance(documentos, list):
        raise ValueError("O arquivo de configuração não contém lista de documentos.")

    return documentos, metadata


# ─────────────────────────────────────────────
#  Importação
# ─────────────────────────────────────────────

def importar_configuracoes_json(caminho_arquivo):
    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        payload = json.load(f)

    documentos, metadata = _extrair_documentos_payload_sync(payload)
    try:
        schema_version = int(metadata.get("schema_version", 0) or 0)
    except (TypeError, ValueError):
        schema_version = 0
    if schema_version > config.SYNC_CONFIG_SCHEMA_VERSION:
        raise ValueError(
            f"Arquivo de configurações em versão mais nova ({schema_version}) do que a suportada"
            f" ({config.SYNC_CONFIG_SCHEMA_VERSION})."
        )

    resumo = {"inseridos": 0, "atualizados": 0, "ignorados": 0, "erros": []}
    vistos = set()

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    for idx, item in enumerate(documentos, start=1):
        try:
            if not isinstance(item, dict):
                resumo["ignorados"] += 1
                continue

            tipo = str(item.get("tipo", "")).upper().strip()
            if tipo not in {"NF", "CTE"}:
                resumo["ignorados"] += 1
                continue

            try:
                numero = int(item.get("numero"))
            except (TypeError, ValueError):
                resumo["ignorados"] += 1
                continue

            numero_original_txt, _ = _coletar_numero_original_para_match(
                item.get("numero_original"), numero
            )
            chave = (tipo, numero, numero_original_txt)
            if chave in vistos:
                resumo["ignorados"] += 1
                continue
            vistos.add(chave)

            existente = _buscar_documento_existente_sync(cursor, tipo, numero, numero_original_txt)
            numero_chave = int(existente["numero"]) if existente else numero

            data_emissao = _normalizar_data_emissao_sync(
                item.get("data_emissao"),
                existente.get("data_emissao") if existente else "",
            )

            competencia = str(
                item.get("competencia") or (existente.get("competencia") if existente else "")
            ).strip().lower()
            if not competencia:
                competencia = competencia_por_data(datetime.now())

            frete = _normalizar_modalidade_frete(
                item.get("frete") or (existente.get("frete") if existente else "FRANQUIA")
            )

            status = str(
                item.get("status") or (existente.get("status") if existente else "OK")
            ).strip()
            if not status:
                status = "OK"

            valor_inicial = _to_float(
                item.get("valor_inicial"),
                existente.get("valor_inicial", 0.0) if existente else 0.0,
            )
            valor_final = _to_float(
                item.get("valor_final"),
                existente.get("valor_final", 0.0) if existente else 0.0,
            )

            valor_inicial_original = (
                _to_optional_float(item.get("valor_inicial_original"))
                if "valor_inicial_original" in item
                else _to_optional_float(existente.get("valor_inicial_original")) if existente else None
            )
            valor_final_original = (
                _to_optional_float(item.get("valor_final_original"))
                if "valor_final_original" in item
                else _to_optional_float(existente.get("valor_final_original")) if existente else None
            )

            status_original = item.get("status_original")
            if status_original is None and existente:
                status_original = existente.get("status_original")
            status_original = (
                str(status_original).strip() if status_original not in (None, "") else None
            )

            cancelado_manual = _to_manual_flag(
                item.get("cancelado_manual"),
                existente.get("cancelado_manual", 0) if existente else 0,
            )
            competencia_manual = _to_manual_flag(
                item.get("competencia_manual"),
                existente.get("competencia_manual", 0) if existente else 0,
            )
            frete_manual = _to_manual_flag(
                item.get("frete_manual"),
                existente.get("frete_manual", 0) if existente else (0 if frete == "FRANQUIA" else 1),
            )
            frete_revisado_manual = _to_manual_flag(
                item.get("frete_revisado_manual"),
                existente.get("frete_revisado_manual", 0) if existente else (
                    1 if (frete_manual == 1 or frete in {"INTERCOMPANY", "DELTA"}) else 0
                ),
            )

            cursor.execute(
                """
                INSERT INTO documentos
                (
                    numero, numero_original, tipo, data_emissao,
                    valor_inicial, valor_final, frete, status, competencia,
                    valor_inicial_original, valor_final_original, status_original,
                    cancelado_manual, competencia_manual, frete_manual, frete_revisado_manual
                )
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(numero, tipo) DO UPDATE SET
                    numero_original=excluded.numero_original,
                    data_emissao=excluded.data_emissao,
                    valor_inicial=excluded.valor_inicial,
                    valor_final=excluded.valor_final,
                    frete=excluded.frete,
                    status=excluded.status,
                    competencia=excluded.competencia,
                    valor_inicial_original=excluded.valor_inicial_original,
                    valor_final_original=excluded.valor_final_original,
                    status_original=excluded.status_original,
                    cancelado_manual=excluded.cancelado_manual,
                    competencia_manual=excluded.competencia_manual,
                    frete_manual=excluded.frete_manual,
                    frete_revisado_manual=excluded.frete_revisado_manual
                """,
                (
                    numero_chave,
                    numero_original_txt,
                    tipo,
                    data_emissao,
                    valor_inicial,
                    valor_final,
                    frete,
                    status,
                    competencia,
                    valor_inicial_original,
                    valor_final_original,
                    status_original,
                    cancelado_manual,
                    competencia_manual,
                    frete_manual,
                    frete_revisado_manual,
                ),
            )

            if existente:
                resumo["atualizados"] += 1
            else:
                resumo["inseridos"] += 1

        except Exception as exc:
            resumo["erros"].append(f"Linha {idx}: {exc}")

    conn.commit()
    conn.close()
    return resumo
