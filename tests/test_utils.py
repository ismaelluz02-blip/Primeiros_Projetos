"""
Testes para src/utils.py — funções puras de formatação, parsing e data.
Execute com: pytest tests/test_utils.py -v
"""

import pytest
from datetime import datetime


# ---------------------------------------------------------------------------
#  Importações — isoladas para cada grupo de funções
# ---------------------------------------------------------------------------

from src.utils import (
    formatar_moeda_brl,
    formatar_moeda_brl_exata,
    parse_valor_monetario,
    normalizar_texto,
    _numero_documento_exibicao,
    _chave_documento_compativel,
    _competencia_para_data,
    _hex_para_rgb,
    _rgb_para_hex,
    _interpolar_cor,
    competencia_por_data,
    MESES,
)


# ---------------------------------------------------------------------------
#  formatar_moeda_brl
# ---------------------------------------------------------------------------

class TestFormatarMoedaBrl:
    def test_valor_inteiro(self):
        assert formatar_moeda_brl(1000) == "R$ 1.000,00"

    def test_valor_decimal(self):
        assert formatar_moeda_brl(1234.56) == "R$ 1.234,56"

    def test_zero(self):
        assert formatar_moeda_brl(0) == "R$ 0,00"

    def test_tipo_invalido_retorna_zero(self):
        assert formatar_moeda_brl("nao_numero") == "R$ 0,00"

    def test_none_retorna_zero(self):
        assert formatar_moeda_brl(None) == "R$ 0,00"


# ---------------------------------------------------------------------------
#  formatar_moeda_brl_exata
# ---------------------------------------------------------------------------

class TestFormatarMoedaBrlExata:
    def test_valor_inteiro_sem_centavos(self):
        result = formatar_moeda_brl_exata(5000)
        assert result == "R$ 5.000"

    def test_valor_com_centavos(self):
        result = formatar_moeda_brl_exata(5000.75)
        assert result == "R$ 5.000,75"

    def test_tipo_invalido(self):
        assert formatar_moeda_brl_exata("x") == "R$ 0,00"


# ---------------------------------------------------------------------------
#  parse_valor_monetario
# ---------------------------------------------------------------------------

class TestParseValorMonetario:
    def test_formato_brl(self):
        assert parse_valor_monetario("R$ 1.234,56") == pytest.approx(1234.56)

    def test_formato_sem_simbolo(self):
        assert parse_valor_monetario("1234,56") == pytest.approx(1234.56)

    def test_formato_ponto_decimal(self):
        assert parse_valor_monetario("1234.56") == pytest.approx(1234.56)

    def test_zero_string(self):
        assert parse_valor_monetario("0") == pytest.approx(0.0)

    def test_string_vazia_retorna_zero(self):
        assert parse_valor_monetario("") == pytest.approx(0.0)

    def test_none_retorna_zero(self):
        assert parse_valor_monetario(None) == pytest.approx(0.0)

    def test_valor_ja_float(self):
        assert parse_valor_monetario(99.9) == pytest.approx(99.9)


# ---------------------------------------------------------------------------
#  normalizar_texto
# ---------------------------------------------------------------------------

class TestNormalizarTexto:
    def test_remove_acentos(self):
        assert normalizar_texto("março") == "marco"

    def test_converte_maiusculas_para_minusculas(self):
        assert normalizar_texto("JANEIRO") == "janeiro"

    def test_texto_ja_normalizado(self):
        assert normalizar_texto("fevereiro") == "fevereiro"

    def test_string_vazia(self):
        assert normalizar_texto("") == ""


# ---------------------------------------------------------------------------
#  _competencia_para_data
# ---------------------------------------------------------------------------

class TestCompetenciaParaData:
    def test_formato_valido_minusculas(self):
        result = _competencia_para_data("janeiro/2025")
        assert result == datetime(2025, 1, 1)

    def test_formato_valido_maiusculas(self):
        result = _competencia_para_data("MARCO/2024")
        assert result == datetime(2024, 3, 1)

    def test_formato_valido_acento(self):
        result = _competencia_para_data("março/2024")
        assert result == datetime(2024, 3, 1)

    def test_todos_os_meses(self):
        """Todos os meses de MESES devem ser convertidos corretamente."""
        for idx, mes in enumerate(MESES, start=1):
            result = _competencia_para_data(f"{mes}/2023")
            assert result is not None, f"Falhou para {mes}"
            assert result.month == idx

    def test_formato_invalido_retorna_none(self):
        assert _competencia_para_data("nao_data") is None

    def test_none_retorna_none(self):
        assert _competencia_para_data(None) is None

    def test_string_vazia_retorna_none(self):
        assert _competencia_para_data("") is None


# ---------------------------------------------------------------------------
#  _numero_documento_exibicao e _chave_documento_compativel
# ---------------------------------------------------------------------------

class TestNumeroDocumentoExibicao:
    def test_nf_sem_substituicao(self):
        result = _numero_documento_exibicao("NF", 100, "", "")
        assert result == "100"

    def test_cte_sem_substituicao(self):
        result = _numero_documento_exibicao("CTE", 200, "", "")
        assert result == "200"

    def test_nf_com_numero_original(self):
        # Documento substituído — deve exibir número original
        result = _numero_documento_exibicao("NF", 100, "99", "01/01/2024")
        assert result != ""  # garante que retorna algo


class TestChaveDocumentoCompativel:
    def test_nf_retorna_chave_consistente(self):
        chave1 = _chave_documento_compativel("NF", 100, "")
        chave2 = _chave_documento_compativel("NF", 100, "")
        assert chave1 == chave2

    def test_tipos_diferentes_geram_chaves_diferentes(self):
        chave_nf = _chave_documento_compativel("NF", 100, "")
        chave_cte = _chave_documento_compativel("CTE", 100, "")
        assert chave_nf != chave_cte

    def test_numeros_diferentes_geram_chaves_diferentes(self):
        chave1 = _chave_documento_compativel("NF", 100, "")
        chave2 = _chave_documento_compativel("NF", 101, "")
        assert chave1 != chave2


# ---------------------------------------------------------------------------
#  Utilitários de cor
# ---------------------------------------------------------------------------

class TestCoresHex:
    def test_hex_para_rgb_branco(self):
        assert _hex_para_rgb("#FFFFFF") == (255, 255, 255)

    def test_hex_para_rgb_preto(self):
        assert _hex_para_rgb("#000000") == (0, 0, 0)

    def test_rgb_para_hex_branco(self):
        assert _rgb_para_hex((255, 255, 255)).upper() == "#FFFFFF"

    def test_rgb_para_hex_preto(self):
        assert _rgb_para_hex((0, 0, 0)).upper() == "#000000"

    def test_round_trip(self):
        cor_original = "#1F4E78"
        rgb = _hex_para_rgb(cor_original)
        hex_de_volta = _rgb_para_hex(rgb)
        assert hex_de_volta.upper() == cor_original.upper()

    def test_interpolar_t0_igual_a(self):
        """t=0 deve retornar a cor A."""
        cor_a = "#FF0000"
        cor_b = "#0000FF"
        result = _interpolar_cor(cor_a, cor_b, 0)
        assert _hex_para_rgb(result) == _hex_para_rgb(cor_a)

    def test_interpolar_t1_igual_b(self):
        """t=1 deve retornar a cor B."""
        cor_a = "#FF0000"
        cor_b = "#0000FF"
        result = _interpolar_cor(cor_a, cor_b, 1)
        assert _hex_para_rgb(result) == _hex_para_rgb(cor_b)


# ---------------------------------------------------------------------------
#  competencia_por_data
# ---------------------------------------------------------------------------

class TestCompetenciaPorData:
    def test_janeiro(self):
        result = competencia_por_data(datetime(2025, 1, 15))
        assert "janeiro" in result.lower()
        assert "2025" in result

    def test_dezembro(self):
        result = competencia_por_data(datetime(2024, 12, 1))
        assert "dezembro" in result.lower()
        assert "2024" in result
