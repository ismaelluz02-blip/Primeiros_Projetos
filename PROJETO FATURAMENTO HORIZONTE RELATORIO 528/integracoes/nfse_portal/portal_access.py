"""Orquestracao de acesso ao Portal Nacional da NFS-e.

Modulo de alto nivel para:
- abrir navegador
- acessar portal
- permitir autenticacao manual (preparado para evoluir depois)
- navegar ate "Notas Emitidas"
- consultar por periodo
- retornar os dados capturados da tela
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from .portal_client import (
    ConsultaNfseResultado,
    NfsePortalClient,
    NfsePortalConfig,
    NfsePortalSelectors,
)


@dataclass(slots=True)
class ConsultaNotasEmitidasRequest:
    base_url: str
    data_inicio: Any
    data_fim: Any
    notas_emitidas_url: str | None = None
    browser: str = "chromium"
    headless: bool = False
    timeout_ms: int = 30_000
    auth_manual_wait_ms: int = 120_000
    max_linhas: int | None = None
    selectors: NfsePortalSelectors | None = None


def criar_config_portal(request: ConsultaNotasEmitidasRequest) -> NfsePortalConfig:
    selectors = request.selectors if request.selectors is not None else NfsePortalSelectors()
    return NfsePortalConfig(
        base_url=request.base_url,
        notas_emitidas_url=request.notas_emitidas_url,
        browser=request.browser,
        headless=request.headless,
        timeout_ms=request.timeout_ms,
        auth_manual_wait_ms=request.auth_manual_wait_ms,
        selectors=selectors,
    )


def consultar_notas_emitidas_portal(
    request: ConsultaNotasEmitidasRequest,
    callback_autenticacao: Callable[[Any], None] | None = None,
) -> ConsultaNfseResultado:
    """Executa o fluxo completo de consulta de Notas Emitidas no Portal.

    Etapas:
    1. abre o navegador
    2. acessa o portal
    3. aguarda autenticacao manual
    4. navega para "Notas Emitidas"
    5. consulta o periodo solicitado
    6. retorna os registros capturados
    """
    config = criar_config_portal(request)
    with NfsePortalClient(config=config) as client:
        return client.executar_consulta(
            data_inicio=request.data_inicio,
            data_fim=request.data_fim,
            callback_autenticacao=callback_autenticacao,
            max_linhas=request.max_linhas,
        )


__all__ = [
    "ConsultaNotasEmitidasRequest",
    "criar_config_portal",
    "consultar_notas_emitidas_portal",
]

