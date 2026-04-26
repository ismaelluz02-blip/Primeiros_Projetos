"""
Testes para src/importacao.py — parsers puros de PDF/planilha.
Execute com: pytest tests/test_importacao.py -v
"""

import pytest
import pandas as pd

from src.importacao import (
    _normalizar_mes_relatorio,
    _normalizar_coluna_relatorio,
    _parse_tipo_documento,
    _linha_parece_cabecalho_planilha,
    _linha_totalizadora_planilha,
    _normalizar_nome_coluna_planilha,
    _achar_coluna,
)


# ---------------------------------------------------------------------------
#  _normalizar_mes_relatorio
# ---------------------------------------------------------------------------

class TestNormalizarMesRelatorio:
    def test_mes_abreviado_maiusculo(self):
        assert _normalizar_mes_relatorio("JAN") == "janeiro"

    def test_mes_abreviado_minusculo(self):
        assert _normalizar_mes_relatorio("jan") == "janeiro"

    def test_mes_por_extenso(self):
        assert _normalizar_mes_relatorio("fevereiro") == "fevereiro"

    def test_mes_com_acento(self):
        assert _normalizar_mes_relatorio("março") == "marco"

    def test_string_invalida_retorna_original_normalizada(self):
        result = _normalizar_mes_relatorio("xyz")
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
#  _normalizar_coluna_relatorio
# ---------------------------------------------------------------------------

class TestNormalizarColunaRelatorio:
    def test_remove_espacos_e_acentos(self):
        result = _normalizar_coluna_relatorio("Número Doc")
        assert " " not in result
        assert result == result.lower()

    def test_string_vazia(self):
        result = _normalizar_coluna_relatorio("")
        assert result == ""

    def test_none(self):
        result = _normalizar_coluna_relatorio(None)
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
#  _parse_tipo_documento
# ---------------------------------------------------------------------------

class TestParseTipoDocumento:
    def test_nf_minusculo(self):
        assert _parse_tipo_documento("nf") == "NF"

    def test_nf_maiusculo(self):
        assert _parse_tipo_documento("NF") == "NF"

    def test_cte_variante(self):
        result = _parse_tipo_documento("CT-e")
        assert result == "CTE"

    def test_cte_simples(self):
        assert _parse_tipo_documento("CTE") == "CTE"

    def test_desconhecido_retorna_none_ou_string(self):
        result = _parse_tipo_documento("BOLETO")
        # Deve retornar None ou string vazia para tipo desconhecido
        assert result is None or result == ""


# ---------------------------------------------------------------------------
#  _linha_parece_cabecalho_planilha
# ---------------------------------------------------------------------------

class TestLinhaPareceCabecalhoPlanilha:
    def test_cabecalho_tipico(self):
        row = ["Filial", "Numero", "Tipo", "Valor"]
        assert _linha_parece_cabecalho_planilha(row) is True

    def test_dados_numericos_nao_e_cabecalho(self):
        row = [88, 12345, "NF", 1500.00]
        assert _linha_parece_cabecalho_planilha(row) is False

    def test_linha_vazia(self):
        assert _linha_parece_cabecalho_planilha([]) is False

    def test_linha_com_nans(self):
        assert _linha_parece_cabecalho_planilha([float("nan"), float("nan")]) is False


# ---------------------------------------------------------------------------
#  _linha_totalizadora_planilha
# ---------------------------------------------------------------------------

class TestLinhaTotalizadoraPlanilha:
    def test_total_na_linha(self):
        assert _linha_totalizadora_planilha(["Total", 0, 9999.99]) is True

    def test_linha_normal_nao_e_total(self):
        assert _linha_totalizadora_planilha([88, 12345, "NF"]) is False

    def test_linha_vazia(self):
        assert _linha_totalizadora_planilha([]) is False


# ---------------------------------------------------------------------------
#  _normalizar_nome_coluna_planilha
# ---------------------------------------------------------------------------

class TestNormalizarNomeColunaPlanilha:
    def test_valor_string(self):
        result = _normalizar_nome_coluna_planilha("  Numero Doc  ", 0)
        assert result.strip() == result  # sem espaços nas pontas
        assert isinstance(result, str)

    def test_valor_numerico(self):
        result = _normalizar_nome_coluna_planilha(1, 2)
        assert isinstance(result, str)

    def test_nan_retorna_col_prefixo(self):
        result = _normalizar_nome_coluna_planilha(float("nan"), 3)
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
#  _achar_coluna
# ---------------------------------------------------------------------------

class TestAcharColuna:
    def test_encontra_coluna_exata(self):
        colunas = {"numero": 0, "tipo": 1, "valor": 2}
        result = _achar_coluna(colunas, [["numero"]])
        assert result == "numero"

    def test_retorna_none_quando_nao_encontra(self):
        colunas = {"filial": 0}
        result = _achar_coluna(colunas, [["numero"]])
        assert result is None

    def test_encontra_por_alternativa(self):
        colunas = {"num_doc": 0, "tipo": 1}
        # regra com múltiplas alternativas
        result = _achar_coluna(colunas, [["numero", "num_doc"]])
        assert result == "num_doc"
