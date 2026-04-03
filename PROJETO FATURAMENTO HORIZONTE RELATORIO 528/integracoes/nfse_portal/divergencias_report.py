"""Geracao de relatorio de divergencias de cancelamento NFS-e.

Foco:
- listar apenas notas canceladas no portal e nao canceladas no sistema
- padronizar campos de saida
- exportar para Excel e JSON
- consolidar resumo da analise
"""

from __future__ import annotations

from datetime import datetime
from typing import Any
import json
import os

import pandas as pd

from .sqlite_compare import comparar_canceladas_portal_banco


TIPOS_DIVERGENCIA_PORTAL_CANCELADA = {
    "PORTAL_CANCELADA_NAO_ENCONTRADA_NO_BANCO",
    "PORTAL_CANCELADA_NAO_MARCADA_CANCELADA_NO_BANCO",
}


def _observacao_por_tipo_divergencia(tipo_divergencia: str) -> str:
    if tipo_divergencia == "PORTAL_CANCELADA_NAO_ENCONTRADA_NO_BANCO":
        return "Nota cancelada no portal e nao encontrada no sistema."
    if tipo_divergencia == "PORTAL_CANCELADA_NAO_MARCADA_CANCELADA_NO_BANCO":
        return "Nota cancelada no portal, mas sem cancelamento no sistema."
    return "Divergencia de cancelamento identificada."


def gerar_relatorio_divergencias_cancelamento(
    portal_canceladas: list[dict[str, Any]],
    db_path: str,
    tipo: str = "NF",
    data_inicio: Any | None = None,
    data_fim: Any | None = None,
) -> dict[str, Any]:
    """Gera relatorio com divergencias de cancelamento portal x sistema."""
    comparacao = comparar_canceladas_portal_banco(
        portal_canceladas=portal_canceladas,
        db_path=db_path,
        tipo=tipo,
        data_inicio=data_inicio,
        data_fim=data_fim,
    )

    divergencias_saida: list[dict[str, Any]] = []
    for item in comparacao.get("divergencias", []):
        tipo_div = str(item.get("tipo_divergencia") or "").strip()
        if tipo_div not in TIPOS_DIVERGENCIA_PORTAL_CANCELADA:
            continue
        divergencias_saida.append(
            {
                "numero": str(item.get("numero") or "").strip(),
                "data": str(item.get("data_portal") or item.get("data_banco") or "").strip(),
                "status_portal": str(item.get("portal_status") or "CANCELADA").strip(),
                "status_sistema": str(item.get("banco_status") or "NAO ENCONTRADA").strip(),
                "observacao": _observacao_por_tipo_divergencia(tipo_div),
            }
        )

    resumo = {
        "total_analisadas": int(comparacao.get("portal_total_canceladas", 0)),
        "total_canceladas": int(comparacao.get("portal_total_canceladas", 0)),
        "total_divergencias": int(len(divergencias_saida)),
    }

    return {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "tipo_documento": str(tipo or "NF").upper().strip(),
        "resumo": resumo,
        "divergencias": divergencias_saida,
    }


def exportar_relatorio_divergencias_json(relatorio: dict[str, Any], caminho_arquivo: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(caminho_arquivo)), exist_ok=True)
    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(relatorio, f, ensure_ascii=False, indent=2)
    return caminho_arquivo


def exportar_relatorio_divergencias_excel(relatorio: dict[str, Any], caminho_arquivo: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(caminho_arquivo)), exist_ok=True)
    divergencias = list(relatorio.get("divergencias", []))
    resumo = dict(relatorio.get("resumo", {}))

    colunas = ["numero", "data", "status_portal", "status_sistema", "observacao"]
    df = pd.DataFrame(divergencias)
    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = ""
    df = df[colunas]

    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Divergencias", index=False)

        df_resumo = pd.DataFrame(
            [
                {"metricas": "Total analisadas", "valor": int(resumo.get("total_analisadas", 0))},
                {"metricas": "Total canceladas", "valor": int(resumo.get("total_canceladas", 0))},
                {"metricas": "Total divergencias", "valor": int(resumo.get("total_divergencias", 0))},
            ]
        )
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

        ws_div = writer.sheets["Divergencias"]
        ws_res = writer.sheets["Resumo"]

        ws_div.column_dimensions["A"].width = 16
        ws_div.column_dimensions["B"].width = 14
        ws_div.column_dimensions["C"].width = 24
        ws_div.column_dimensions["D"].width = 28
        ws_div.column_dimensions["E"].width = 52

        ws_res.column_dimensions["A"].width = 24
        ws_res.column_dimensions["B"].width = 16

    return caminho_arquivo


def exportar_relatorio_divergencias(
    relatorio: dict[str, Any],
    caminho_arquivo: str,
    formato: str = "excel",
) -> str:
    fmt = str(formato or "excel").strip().lower()
    if fmt in {"excel", "xlsx"}:
        return exportar_relatorio_divergencias_excel(relatorio, caminho_arquivo)
    if fmt in {"json"}:
        return exportar_relatorio_divergencias_json(relatorio, caminho_arquivo)
    raise ValueError("Formato de exportacao invalido. Use 'excel' ou 'json'.")


def resumo_texto_relatorio_divergencias(relatorio: dict[str, Any]) -> str:
    resumo = dict(relatorio.get("resumo", {}))
    return (
        f"Total analisadas: {int(resumo.get('total_analisadas', 0))}\n"
        f"Total canceladas: {int(resumo.get('total_canceladas', 0))}\n"
        f"Total divergencias: {int(resumo.get('total_divergencias', 0))}"
    )


__all__ = [
    "gerar_relatorio_divergencias_cancelamento",
    "exportar_relatorio_divergencias_json",
    "exportar_relatorio_divergencias_excel",
    "exportar_relatorio_divergencias",
    "resumo_texto_relatorio_divergencias",
]

