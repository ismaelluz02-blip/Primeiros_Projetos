"""
Funções utilitárias puras — sem estado global, sem UI.

Formatação de moeda, normalização de texto, cálculos de data/competência
e manipulação de números de documento.
"""

import re
import unicodedata
from calendar import monthrange
from datetime import datetime

# ─────────────────────────────────────────────
#  Constantes
# ─────────────────────────────────────────────

MESES = [
    "janeiro",
    "fevereiro",
    "marco",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
]


# ─────────────────────────────────────────────
#  Formatação de moeda
# ─────────────────────────────────────────────

def valor_brasileiro(v):
    v = v.replace(".", "").replace(",", ".")
    return float(v)


def formatar_moeda_brl(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"


def formatar_moeda_brl_exata(valor):
    try:
        v = float(valor)
    except:
        return "R$ 0,00"
    if abs(v - round(v)) < 0.005:
        return f"R$ {int(round(v)):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_valor_monetario(valor_str):
    bruto = re.sub(r"[^\d,.\s]", "", str(valor_str)).replace(" ", "")
    if not bruto:
        return None

    if "," in bruto:
        try:
            return float(bruto.replace(".", "").replace(",", "."))
        except ValueError:
            return None

    if bruto.count(".") == 1:
        try:
            return float(bruto)
        except ValueError:
            return None

    if bruto.count(".") > 1:
        return None

    try:
        return float(bruto.replace(".", ""))
    except ValueError:
        return None


# ─────────────────────────────────────────────
#  Normalização de texto
# ─────────────────────────────────────────────

def normalizar_texto(texto):
    sem_acentos = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return sem_acentos.upper()


# ─────────────────────────────────────────────
#  Manipulação de números de documento
# ─────────────────────────────────────────────

def _numero_para_texto(numero):
    if numero in (None, ""):
        return ""

    if isinstance(numero, float):
        return str(int(numero)) if numero.is_integer() else str(numero)

    numero_txt = str(numero).strip()
    if not numero_txt:
        return ""

    if re.fullmatch(r"-?\d+\.0+", numero_txt):
        try:
            return str(int(float(numero_txt)))
        except ValueError:
            return numero_txt

    return numero_txt


def _extrair_ano_data_emissao(data_emissao):
    if isinstance(data_emissao, datetime):
        return int(data_emissao.year)

    data_txt = str(data_emissao or "").strip()
    if not data_txt:
        return None

    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(data_txt, formato).year
        except ValueError:
            continue
    return None


def _normalizar_numero_original_nf(numero, numero_original="", data_emissao=""):
    numero_txt = re.sub(r"\D", "", _numero_para_texto(numero))
    numero_original_txt = re.sub(r"\D", "", str(numero_original or "").strip())

    numero_legado_txt = ""
    ano_data = _extrair_ano_data_emissao(data_emissao)
    if numero_txt:
        if ano_data and numero_txt.startswith(str(ano_data)) and 0 < len(numero_txt[4:]) <= 4:
            numero_legado_txt = numero_txt[4:]
        elif len(numero_txt) == 8 and numero_txt.startswith("20"):
            numero_legado_txt = numero_txt[4:]

    numero_legado_txt = str(int(numero_legado_txt)) if numero_legado_txt else ""

    if numero_original_txt:
        if numero_legado_txt and numero_original_txt == numero_txt:
            return numero_legado_txt
        return str(int(numero_original_txt)) if numero_original_txt.isdigit() else numero_original_txt

    if numero_legado_txt:
        return numero_legado_txt

    if numero_txt:
        return str(int(numero_txt)) if numero_txt.isdigit() else numero_txt

    return ""


def _coletar_numero_original_para_match(numero_original, numero, data_emissao=""):
    numero_original_txt = _normalizar_numero_original_nf(numero, numero_original, data_emissao)
    try:
        numero_original_int = int(re.sub(r"\D", "", numero_original_txt))
    except ValueError:
        numero_original_int = None
    return numero_original_txt, numero_original_int


def _numero_documento_exibicao(tipo, numero, numero_original="", data_emissao=""):
    tipo_norm = str(tipo or "").upper().strip()
    if tipo_norm == "NF":
        return _normalizar_numero_original_nf(numero, numero_original, data_emissao)
    return _numero_para_texto(numero)


def _chave_documento_compativel(tipo, numero, numero_original=""):
    tipo_norm = str(tipo or "").upper().strip() or "DOC"
    if tipo_norm == "NF":
        numero_match_txt, numero_match_int = _coletar_numero_original_para_match(numero_original, numero)
        numero_txt = str(numero_match_int) if numero_match_int is not None else numero_match_txt
    else:
        numero_txt = _numero_para_texto(numero)
    return f"{tipo_norm}:{numero_txt or 'SEMNUM'}"


# ─────────────────────────────────────────────
#  Datas e competência
# ─────────────────────────────────────────────

def competencia_por_data(data):
    return f"{MESES[data.month - 1]}/{data.year}"


def periodo_padrao_mes_atual():
    hoje = datetime.now()
    primeiro_dia = datetime(hoje.year, hoje.month, 1)
    ultimo_dia = datetime(hoje.year, hoje.month, monthrange(hoje.year, hoje.month)[1])
    return primeiro_dia, ultimo_dia


def obter_periodo_padrao_dashboard():
    hoje = datetime.now()
    return datetime(hoje.year, 1, 1), hoje


def obter_periodo_padrao_relatorios():
    hoje = datetime.now()
    return datetime(hoje.year, hoje.month, 1), hoje


def ler_data_filtro(texto_data, nome_campo):
    try:
        return datetime.strptime(texto_data.strip(), "%d/%m/%Y")
    except ValueError as exc:
        raise ValueError(f"{nome_campo} invalida. Use o formato DD/MM/AAAA.") from exc


def _obter_periodo_por_entries(entry_inicio, entry_fim, contexto="Período", silencioso=False):
    if entry_inicio is None or entry_fim is None:
        return None, None

    try:
        data_inicial = ler_data_filtro(entry_inicio.get(), "Data inicial")
        data_final = ler_data_filtro(entry_fim.get(), "Data final")
    except ValueError as exc:
        if silencioso:
            return None, None
        raise ValueError(f"{contexto}: {exc}") from exc

    return data_inicial, data_final
