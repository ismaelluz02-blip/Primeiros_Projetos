"""Integracao com o Portal Nacional da NFS-e."""

from .portal_client import (
    ConsultaNfseResultado,
    NfsePeriodoConsulta,
    NfsePortalClient,
    NfsePortalConfig,
    NfsePortalSelectors,
    filtrar_notas_canceladas,
    processar_notas_canceladas_portal,
)
from .portal_access import (
    ConsultaNotasEmitidasRequest,
    consultar_notas_emitidas_portal,
    criar_config_portal,
)
from .sqlite_compare import (
    buscar_nota_no_banco,
    buscar_nota_banco_por_numero,
    comparar_cancelamentos_portal_com_sqlite,
    comparar_canceladas_portal_banco,
    listar_notas_canceladas_banco,
    nota_cancelada_no_banco,
    verificar_cancelamento_no_banco,
)
from .divergencias_report import (
    exportar_relatorio_divergencias,
    exportar_relatorio_divergencias_excel,
    exportar_relatorio_divergencias_json,
    gerar_relatorio_divergencias_cancelamento,
    resumo_texto_relatorio_divergencias,
)

__all__ = [
    "ConsultaNfseResultado",
    "NfsePeriodoConsulta",
    "NfsePortalClient",
    "NfsePortalConfig",
    "NfsePortalSelectors",
    "filtrar_notas_canceladas",
    "processar_notas_canceladas_portal",
    "ConsultaNotasEmitidasRequest",
    "criar_config_portal",
    "consultar_notas_emitidas_portal",
    "buscar_nota_no_banco",
    "buscar_nota_banco_por_numero",
    "verificar_cancelamento_no_banco",
    "nota_cancelada_no_banco",
    "comparar_cancelamentos_portal_com_sqlite",
    "listar_notas_canceladas_banco",
    "comparar_canceladas_portal_banco",
    "gerar_relatorio_divergencias_cancelamento",
    "exportar_relatorio_divergencias_json",
    "exportar_relatorio_divergencias_excel",
    "exportar_relatorio_divergencias",
    "resumo_texto_relatorio_divergencias",
]
