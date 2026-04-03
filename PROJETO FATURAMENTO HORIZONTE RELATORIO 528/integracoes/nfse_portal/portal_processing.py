"""Processamento de registros capturados do Portal NFS-e.

Etapa anterior a comparacao com SQLite:
- recebe registros brutos capturados da tela
- identifica somente notas canceladas
- padroniza para: numero, data, status, chave
"""

from __future__ import annotations

from datetime import datetime
from typing import Any
import re
import unicodedata


def _normalizar_chave_coluna(texto: str) -> str:
    base = unicodedata.normalize("NFKD", str(texto or ""))
    base = "".join(ch for ch in base if not unicodedata.combining(ch))
    base = re.sub(r"\s+", " ", base.strip().lower())
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base or "coluna"


def _normalizar_texto_comparacao(texto: Any) -> str:
    valor = unicodedata.normalize("NFKD", str(texto or ""))
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"\s+", " ", valor).strip().upper()
    return valor


def _status_indica_cancelamento(status: str) -> bool:
    txt = _normalizar_texto_comparacao(status)
    return ("CANCELAD" in txt) or ("CANCELAMENTO" in txt)


def _coletar_valor_por_chaves(registro: dict[str, Any], chaves_preferenciais: tuple[str, ...]) -> str:
    for chave in chaves_preferenciais:
        if chave in registro:
            valor = str(registro.get(chave) or "").strip()
            if valor:
                return valor
    return ""


def _coletar_valor_por_fragmentos(registro: dict[str, Any], fragmentos: tuple[str, ...]) -> str:
    for chave, valor in registro.items():
        chave_norm = _normalizar_chave_coluna(chave)
        if any(fragmento in chave_norm for fragmento in fragmentos):
            txt = str(valor or "").strip()
            if txt:
                return txt
    return ""


def _padronizar_data_saida(data_txt: str) -> str:
    txt = str(data_txt or "").strip()
    if not txt:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:19], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return txt


def _extrair_numero_nota(registro: dict[str, Any]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "numero",
            "numero_nfse",
            "numero_nota",
            "n_nfse",
            "nro_nfse",
            "num_nfse",
            "n_da_nota",
            "n_nota",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("numero", "nota", "nfse"))
    if not valor:
        return ""
    numeros = re.findall(r"\d+", valor)
    if numeros:
        return numeros[-1]
    return valor.strip()


def _extrair_data_nota(registro: dict[str, Any]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "data",
            "data_emissao",
            "dt_emissao",
            "emissao",
            "data_de_emissao",
            "data_da_emissao",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("data", "emissao"))
    return _padronizar_data_saida(valor)


def _extrair_chave_nota(registro: dict[str, Any]) -> str | None:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "chave",
            "chave_acesso",
            "chave_nfse",
            "codigo_verificacao",
            "cod_verificacao",
            "codigo",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("chave", "verificacao", "codigo"))
    valor = str(valor or "").strip()
    return valor if valor else None


def _extrair_status_nota(registro: dict[str, Any]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "status",
            "situacao",
            "situacao_nfse",
            "status_nfse",
            "estado",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("status", "situacao"))
    return str(valor or "").strip()


def filtrar_notas_canceladas(registros: list[dict[str, Any]]) -> list[dict[str, str | None]]:
    """Retorna somente notas canceladas em estrutura padronizada.

    Estrutura:
    - numero
    - data
    - status
    - chave
    """
    canceladas: list[dict[str, str | None]] = []
    vistos: set[tuple[str, str | None]] = set()

    for registro in registros or []:
        if not isinstance(registro, dict):
            continue

        status_raw = _extrair_status_nota(registro)
        if not _status_indica_cancelamento(status_raw):
            continue

        numero = _extrair_numero_nota(registro)
        if not numero:
            continue

        data = _extrair_data_nota(registro)
        chave = _extrair_chave_nota(registro)
        assinatura = (numero, chave)
        if assinatura in vistos:
            continue
        vistos.add(assinatura)

        canceladas.append(
            {
                "numero": numero,
                "data": data,
                "status": "CANCELADA",
                "chave": chave,
            }
        )

    return canceladas


def processar_notas_canceladas_portal(registros: list[dict[str, Any]]) -> list[dict[str, str | None]]:
    """Alias semantico da etapa de processamento antes da comparacao."""
    return filtrar_notas_canceladas(registros)


__all__ = [
    "filtrar_notas_canceladas",
    "processar_notas_canceladas_portal",
]
