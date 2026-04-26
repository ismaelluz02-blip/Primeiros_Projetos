"""
Funções puras de filtragem e montagem de DataFrames para relatórios.

Sem dependências de UI. Os orquestradores com messagebox/filedialog
(gerar_excel, exportar_relatorio_filtrado, abrir_relatorio, abrir_relatorio_filtrado)
permanecem em sistema_faturamento.py.
"""

from datetime import datetime

import pandas as pd

from src.banco import obter_conexao_banco
from src.utils import MESES, _numero_documento_exibicao, _chave_documento_compativel


# ─────────────────────────────────────────────
#  Filtragem de documentos por período
# ─────────────────────────────────────────────

def _obter_dataframe_relatorio_filtrado(data_inicial, data_final, docs_df_base=None):
    if isinstance(docs_df_base, pd.DataFrame):
        df = docs_df_base.copy()
    else:
        conn = obter_conexao_banco()
        df = pd.read_sql_query("SELECT * FROM documentos", conn)
        conn.close()

    if df.empty:
        return pd.DataFrame(), "Nenhum documento encontrado no banco."

    df["data_emissao"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao"])
    if df.empty:
        return pd.DataFrame(), "Não há datas válidas para gerar o relatório."

    df["numero"] = pd.to_numeric(df["numero"], errors="coerce")
    df = df.dropna(subset=["numero"])
    if df.empty:
        return pd.DataFrame(), "Não há números de documento válidos para gerar o relatório."
    df["numero"] = df["numero"].astype(int)

    def competencia_para_data(comp_str):
        try:
            partes = str(comp_str).lower().split("/")
            if len(partes) == 2:
                mes_nome = partes[0].strip()
                ano_str = partes[1].strip()
                ano = int(ano_str)
                mes_idx = MESES.index(mes_nome) + 1
                return datetime(ano, mes_idx, 1)
        except Exception:
            pass
        return None

    df["data_competencia"] = df["competencia"].apply(competencia_para_data)
    df = df.dropna(subset=["data_competencia"])
    df = df[(df["data_competencia"] >= data_inicial) & (df["data_competencia"] <= data_final)].copy()

    if df.empty:
        return pd.DataFrame(), ""

    df["numero_original_num"] = pd.to_numeric(df.get("numero_original"), errors="coerce")
    df["numero_exibicao"] = df.apply(
        lambda r: _numero_documento_exibicao(
            r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")
        ),
        axis=1,
    )
    df["chave_documento"] = df.apply(
        lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")
    return df, ""


# ─────────────────────────────────────────────
#  Montagem do DataFrame de exportação
# ─────────────────────────────────────────────

def _montar_dataframe_exportacao_periodo(df_filtrado):
    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame()

    dados = df_filtrado.copy()
    dados["numero_doc"] = dados.apply(
        lambda r: _numero_documento_exibicao(
            r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")
        ),
        axis=1,
    )
    dados["numero_doc"] = dados["numero_doc"].astype(str).str.strip()
    dados = dados[dados["numero_doc"] != ""].copy()
    if dados.empty:
        return pd.DataFrame()

    dados["numero_doc_ordem"] = pd.to_numeric(dados["numero_doc"], errors="coerce")
    # Evita alerta do Excel "numero armazenado como texto" para documentos numericos.
    dados["numero_doc"] = dados.apply(
        lambda r: int(r["numero_doc_ordem"]) if pd.notna(r["numero_doc_ordem"]) else r["numero_doc"],
        axis=1,
    )
    dados["numero_ordem"] = pd.to_numeric(dados["numero"], errors="coerce")
    dados["numero_doc_ordem"] = dados["numero_doc_ordem"].fillna(dados["numero_ordem"])
    dados["tipo"] = dados["tipo"].astype(str).str.upper()
    dados["frete"] = dados["frete"].astype(str)
    dados["status"] = dados["status"].astype(str)
    dados["mes_referencia"] = pd.to_datetime(dados["data_competencia"], errors="coerce")
    dados = dados.sort_values(
        ["data_emissao", "numero_doc_ordem", "numero_doc"], ascending=[True, True, True]
    )

    export_df = dados[
        [
            "data_emissao",
            "mes_referencia",
            "numero_doc",
            "tipo",
            "frete",
            "valor_inicial",
            "valor_final",
            "status",
        ]
    ].copy()
    export_df.columns = [
        "Data Emissao",
        "Mes Referencia",
        "Numero Doc",
        "Tipo Doc",
        "Frete",
        "Valor Inicial",
        "Valor Final",
        "Status",
    ]
    return export_df
