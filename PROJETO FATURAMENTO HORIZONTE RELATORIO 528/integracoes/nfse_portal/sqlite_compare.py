"""Comparacao de cancelamentos NFS-e entre Portal e banco SQLite local."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any
import sqlite3
import re


def _normalizar_numero_chave(valor: Any) -> str:
    txt = str(valor or "").strip()
    if not txt:
        return ""
    digitos = re.sub(r"\D", "", txt)
    if digitos:
        return str(int(digitos)) if digitos != "0" else "0"
    return txt.upper()


def _parse_data_flexivel(valor: Any) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    txt = str(valor or "").strip()
    if not txt:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:19], fmt).date()
        except ValueError:
            continue
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
        return dt.date()
    except Exception:
        return None


def _padronizar_data_texto(valor: Any) -> str:
    dt = _parse_data_flexivel(valor)
    return dt.strftime("%d/%m/%Y") if dt else ""


def _status_cancelado(status: Any, cancelado_manual: Any = 0) -> bool:
    try:
        if int(cancelado_manual or 0) == 1:
            return True
    except (TypeError, ValueError):
        pass
    txt = str(status or "").upper()
    return "CANCELAD" in txt


def _conectar(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def buscar_nota_banco_por_numero(db_path: str, numero: Any, tipo: str = "NF") -> dict[str, Any] | None:
    """Busca nota no banco pelo numero, com fallback por numero_original para NF."""
    tipo_norm = str(tipo or "NF").upper().strip()
    chave = _normalizar_numero_chave(numero)
    if not chave:
        return None

    conn = _conectar(db_path)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT * FROM documentos WHERE tipo=? ORDER BY id DESC",
            (tipo_norm,),
        )
        for row in cur.fetchall():
            item = dict(row)
            ch_num = _normalizar_numero_chave(item.get("numero"))
            ch_org = _normalizar_numero_chave(item.get("numero_original"))
            if chave in {ch_num, ch_org}:
                return item
    finally:
        conn.close()
    return None


def nota_cancelada_no_banco(nota: dict[str, Any] | None) -> bool:
    if not nota:
        return False
    return _status_cancelado(nota.get("status"), nota.get("cancelado_manual", 0))


def buscar_nota_no_banco(db_path: str, numero: Any, tipo: str = "NF") -> dict[str, Any] | None:
    """Alias semantico para busca de nota no SQLite local."""
    return buscar_nota_banco_por_numero(db_path=db_path, numero=numero, tipo=tipo)


def verificar_cancelamento_no_banco(nota: dict[str, Any] | None) -> bool:
    """Alias semantico para validar cancelamento no SQLite local."""
    return nota_cancelada_no_banco(nota)


def comparar_cancelamentos_portal_com_sqlite(
    notas_canceladas_portal: list[dict[str, Any]],
    db_path: str,
    tipo: str = "NF",
) -> dict[str, Any]:
    """Compara canceladas no portal com o SQLite e identifica divergencias.

    Regra de divergencia:
    - status no portal = CANCELADA
    - status no sistema local != CANCELADA
    """
    tipo_norm = str(tipo or "NF").upper().strip()
    total_analisadas = 0
    total_canceladas_portal = 0
    divergencias: list[dict[str, Any]] = []
    portal_processadas: list[dict[str, Any]] = []
    vistos: set[str] = set()

    for item in notas_canceladas_portal or []:
        if not isinstance(item, dict):
            continue

        numero = _normalizar_numero_chave(item.get("numero"))
        if not numero or numero in vistos:
            continue
        vistos.add(numero)
        total_analisadas += 1

        status_portal = str(item.get("status") or "CANCELADA").strip()
        if not _status_cancelado(
            status_portal,
            1 if "CANCELAD" in str(status_portal).upper() else 0,
        ):
            # Esta etapa recebe "canceladas do portal", mas mantemos validacao defensiva.
            continue
        total_canceladas_portal += 1

        nota_banco = buscar_nota_no_banco(db_path=db_path, numero=numero, tipo=tipo_norm)
        cancelada_no_banco = verificar_cancelamento_no_banco(nota_banco)
        status_sistema = str((nota_banco or {}).get("status") or "NAO ENCONTRADA").strip()

        portal_item = {
            "numero": numero,
            "data": _padronizar_data_texto(item.get("data")),
            "status_portal": status_portal or "CANCELADA",
            "status_sistema": status_sistema,
        }
        portal_processadas.append(portal_item)

        if cancelada_no_banco:
            continue

        observacao = (
            "Nota cancelada no portal e nao encontrada no banco local."
            if nota_banco is None
            else "Nota cancelada no portal, mas ainda nao esta cancelada no sistema local."
        )
        divergencias.append(
            {
                "numero": numero,
                "data": portal_item["data"],
                "status_portal": portal_item["status_portal"],
                "status_sistema": status_sistema,
                "observacao": observacao,
            }
        )

    return {
        "total_analisadas": total_analisadas,
        "total_canceladas_portal": total_canceladas_portal,
        "total_divergencias": len(divergencias),
        "portal_processadas": portal_processadas,
        "divergencias": divergencias,
    }


def listar_notas_canceladas_banco(
    db_path: str,
    tipo: str = "NF",
    data_inicio: Any | None = None,
    data_fim: Any | None = None,
) -> list[dict[str, Any]]:
    """Lista notas canceladas no banco em formato padronizado."""
    tipo_norm = str(tipo or "NF").upper().strip()
    dt_ini = _parse_data_flexivel(data_inicio) if data_inicio else None
    dt_fim = _parse_data_flexivel(data_fim) if data_fim else None

    conn = _conectar(db_path)
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM documentos WHERE tipo=? ORDER BY id DESC", (tipo_norm,))
        rows = [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()

    canceladas: list[dict[str, Any]] = []
    vistos: set[str] = set()

    for row in rows:
        if not _status_cancelado(row.get("status"), row.get("cancelado_manual", 0)):
            continue

        dt_emissao = _parse_data_flexivel(row.get("data_emissao"))
        if dt_ini and (not dt_emissao or dt_emissao < dt_ini):
            continue
        if dt_fim and (not dt_emissao or dt_emissao > dt_fim):
            continue

        chave_num = _normalizar_numero_chave(row.get("numero_original")) or _normalizar_numero_chave(row.get("numero"))
        if not chave_num or chave_num in vistos:
            continue
        vistos.add(chave_num)

        canceladas.append(
            {
                "numero": chave_num,
                "data": _padronizar_data_texto(row.get("data_emissao")),
                "status": str(row.get("status") or "").strip(),
                "chave": None,
                "tipo": str(row.get("tipo") or "").upper(),
                "id": row.get("id"),
            }
        )
    return canceladas


def comparar_canceladas_portal_banco(
    portal_canceladas: list[dict[str, Any]],
    db_path: str,
    tipo: str = "NF",
    data_inicio: Any | None = None,
    data_fim: Any | None = None,
) -> dict[str, Any]:
    """Compara canceladas do portal versus banco e retorna divergencias.

    Divergencias:
    - PORTAL_CANCELADA_NAO_ENCONTRADA_NO_BANCO
    - PORTAL_CANCELADA_NAO_MARCADA_CANCELADA_NO_BANCO
    - BANCO_CANCELADA_NAO_PRESENTE_NO_PORTAL
    """
    tipo_norm = str(tipo or "NF").upper().strip()
    portal_padronizadas: list[dict[str, Any]] = []
    portal_chaves: set[str] = set()
    ignorados_portal = 0

    for item in portal_canceladas or []:
        if not isinstance(item, dict):
            ignorados_portal += 1
            continue
        numero = _normalizar_numero_chave(item.get("numero"))
        if not numero:
            ignorados_portal += 1
            continue
        if numero in portal_chaves:
            continue
        portal_chaves.add(numero)
        portal_padronizadas.append(
            {
                "numero": numero,
                "data": _padronizar_data_texto(item.get("data")),
                "status": str(item.get("status") or "CANCELADA").strip(),
                "chave": (str(item.get("chave")).strip() if item.get("chave") else None),
            }
        )

    canceladas_banco = listar_notas_canceladas_banco(
        db_path=db_path,
        tipo=tipo_norm,
        data_inicio=data_inicio,
        data_fim=data_fim,
    )
    chaves_banco_canceladas = {str(x["numero"]) for x in canceladas_banco}

    divergencias: list[dict[str, Any]] = []
    consistentes: list[dict[str, Any]] = []

    for nota_portal in portal_padronizadas:
        numero = nota_portal["numero"]
        nota_local = buscar_nota_banco_por_numero(db_path=db_path, numero=numero, tipo=tipo_norm)
        if not nota_local:
            divergencias.append(
                {
                    "tipo_divergencia": "PORTAL_CANCELADA_NAO_ENCONTRADA_NO_BANCO",
                    "numero": numero,
                    "portal_status": nota_portal.get("status"),
                    "banco_status": None,
                    "data_portal": nota_portal.get("data"),
                    "data_banco": None,
                    "chave": nota_portal.get("chave"),
                }
            )
            continue
        if not nota_cancelada_no_banco(nota_local):
            divergencias.append(
                {
                    "tipo_divergencia": "PORTAL_CANCELADA_NAO_MARCADA_CANCELADA_NO_BANCO",
                    "numero": numero,
                    "portal_status": nota_portal.get("status"),
                    "banco_status": str(nota_local.get("status") or "").strip(),
                    "data_portal": nota_portal.get("data"),
                    "data_banco": _padronizar_data_texto(nota_local.get("data_emissao")),
                    "chave": nota_portal.get("chave"),
                }
            )
            continue
        consistentes.append(
            {
                "numero": numero,
                "portal_status": nota_portal.get("status"),
                "banco_status": str(nota_local.get("status") or "").strip(),
                "data_portal": nota_portal.get("data"),
                "data_banco": _padronizar_data_texto(nota_local.get("data_emissao")),
                "chave": nota_portal.get("chave"),
            }
        )

    for nota_banco in canceladas_banco:
        numero = str(nota_banco["numero"])
        if numero not in portal_chaves:
            divergencias.append(
                {
                    "tipo_divergencia": "BANCO_CANCELADA_NAO_PRESENTE_NO_PORTAL",
                    "numero": numero,
                    "portal_status": None,
                    "banco_status": nota_banco.get("status"),
                    "data_portal": None,
                    "data_banco": nota_banco.get("data"),
                    "chave": nota_banco.get("chave"),
                }
            )

    return {
        "portal_total_canceladas": len(portal_padronizadas),
        "banco_total_canceladas": len(chaves_banco_canceladas),
        "total_divergencias": len(divergencias),
        "total_consistentes": len(consistentes),
        "portal_ignorados": ignorados_portal,
        "canceladas_portal": portal_padronizadas,
        "canceladas_banco": canceladas_banco,
        "consistentes": consistentes,
        "divergencias": divergencias,
    }


__all__ = [
    "buscar_nota_no_banco",
    "buscar_nota_banco_por_numero",
    "verificar_cancelamento_no_banco",
    "nota_cancelada_no_banco",
    "comparar_cancelamentos_portal_com_sqlite",
    "listar_notas_canceladas_banco",
    "comparar_canceladas_portal_banco",
]
