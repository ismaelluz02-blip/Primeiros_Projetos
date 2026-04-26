"""
Funções puras de parsing para importação de relatórios.

Parsers de PDF (PyMuPDF) e planilha (Excel/pandas) — sem dependências de UI.
Os orquestradores com UI (importar_relatorio_consolidado, importar_relatorio_planilha)
e os wrappers (selecionar_relatorio, importar_relatorio_ui) permanecem em
sistema_faturamento.py.
"""

import re
from datetime import datetime

import pandas as pd

import src.config as config
from src.utils import MESES, parse_valor_monetario, normalizar_texto, competencia_por_data


def _normalizar_mes_relatorio(mes_txt):
    mes_norm = normalizar_texto(mes_txt).lower()
    mapa = {normalizar_texto(m).lower(): m for m in MESES}
    return mapa.get(mes_norm, mes_norm)


def _extrair_docs_pagina_relatorio(texto_pagina, competencia_atual=None):
    docs = []
    linhas = [l.strip() for l in texto_pagina.splitlines() if l.strip()]

    header_re = re.compile(r"(?:C\.T\.R\.C\.|N\.F\.)\s*-\s*([A-Za-z?-?]+)\s*/\s*(\d{2,4})", re.IGNORECASE)
    money_re = re.compile(r"R\$\s*([0-9\.,]+)")
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    i = 0
    while i < len(linhas):
        linha = linhas[i]
        m_header = header_re.search(linha)
        if m_header:
            mes_txt = _normalizar_mes_relatorio(m_header.group(1))
            ano_txt = m_header.group(2)
            ano = int(ano_txt)
            if ano < 100:
                ano += 2000
            competencia_atual = f"{mes_txt}/{ano}"

        if linha.isdigit() and (i + 1) < len(linhas) and linhas[i + 1].isdigit():
            window_end = min(i + 35, len(linhas))
            tipo_idx = None
            tipo = None
            for j in range(i + 2, window_end):
                t = normalizar_texto(linhas[j])
                if t in {"CTRC", "NF"}:
                    tipo_idx = j
                    tipo = t
                    break

            if tipo_idx is not None:
                trecho = linhas[i:tipo_idx + 1]
                codigo = int(linha)

                valor = None
                for item in trecho:
                    mm = money_re.search(item)
                    if mm:
                        v = parse_valor_monetario(mm.group(1))
                        if v is not None and v > 0:
                            valor = v
                            break

                data = None
                for item in trecho:
                    if date_re.match(item):
                        try:
                            data = datetime.strptime(item, "%d/%m/%Y")
                            break
                        except ValueError:
                            pass

                if data and valor is not None:
                    if tipo == "NF":
                        numero = codigo
                        valor_final = valor * 0.95
                    else:
                        numero = codigo
                        valor_final = valor

                    docs.append({
                        "numero": numero,
                        "numero_original": linha.strip(),
                        "tipo": "NF" if tipo == "NF" else "CTE",
                        "data": data,
                        "valor_inicial": valor,
                        "valor_final": valor_final,
                        "frete": "FRANQUIA",
                        "status": "OK",
                        "competencia": competencia_atual.lower() if competencia_atual else competencia_por_data(data),
                    })

                i = tipo_idx + 1
                continue

        i += 1

    return docs, competencia_atual

def _normalizar_coluna_relatorio(nome_coluna):
    base = normalizar_texto(str(nome_coluna))
    base = re.sub(r"[^A-Z0-9 ]", " ", base)
    return " ".join(base.split())


def _achar_coluna(colunas_norm, regras):
    for regra in regras:
        for col_original, col_norm in colunas_norm.items():
            if all(chave in col_norm for chave in regra):
                return col_original
    return None


def _achar_coluna_exata(colunas_norm, nomes_exatos):
    nomes = {n.upper() for n in nomes_exatos}
    for col_original, col_norm in colunas_norm.items():
        if col_norm in nomes:
            return col_original
    return None


def _parse_tipo_documento(valor_tipo):
    t = normalizar_texto(str(valor_tipo))
    if "CTE" in t or "CTRC" in t:
        return "CTE"
    if "NF" in t:
        return "NF"
    return None


def _mapear_colunas_planilha(df):
    colunas_norm = {col: _normalizar_coluna_relatorio(col) for col in df.columns}

    col_tipo = _achar_coluna(colunas_norm, [
        ("TIPO", "DOC"),
        ("TIPO",),
    ])
    col_serie = _achar_coluna_exata(colunas_norm, ["SERIE"]) or _achar_coluna(colunas_norm, [("SERIE",)])
    col_numero = (
        _achar_coluna_exata(colunas_norm, ["CODIGO", "COD"])
        or _achar_coluna(colunas_norm, [
            ("NUMERO", "DOC"),
            ("NUM", "DOC"),
            ("NRO", "DOC"),
            ("NR", "DOC"),
            ("DOC", "NUM"),
            ("CODIGO",),
            ("COD",),
            ("NUMERO",),
            ("DOCUMENTO",),
            ("NR",),
        ])
    )

    # Data de emissao: sempre prioriza coluna "Data" (sem "Ref").
    col_data = _achar_coluna_exata(colunas_norm, ["DATA"])
    if not col_data:
        for col_original, col_norm in colunas_norm.items():
            if (("DATA" in col_norm) or (col_norm == "DT") or ("EMISSAO" in col_norm)) and ("REF" not in col_norm):
                col_data = col_original
                break

    # Data de referencia: apenas fallback quando Data estiver vazia/invalida.
    col_data_ref = _achar_coluna_exata(colunas_norm, ["DATA REF", "DT REF"])
    if not col_data_ref:
        for col_original, col_norm in colunas_norm.items():
            if ("REF" in col_norm) and (("DATA" in col_norm) or ("DT" in col_norm)):
                col_data_ref = col_original
                break

    col_valor = _achar_coluna(colunas_norm, [
        ("FRETE",),
        ("VALOR", "DOCUMENTO"),
        ("VALOR", "TOTAL"),
        ("VLR", "DOCUMENTO"),
        ("VLR", "TOTAL"),
        ("VALOR", "FRETE"),
        ("VLR", "FRETE"),
        ("TOTAL", "FRETE"),
        ("VALOR", "RECEBER"),
        ("VALOR",),
        ("VLR",),
    ])
    col_frete = _achar_coluna(colunas_norm, [
        ("FRETE",),
    ])
    col_status = _achar_coluna(colunas_norm, [
        ("STATUS",),
        ("SITUACAO",),
    ])
    col_filial = _achar_coluna_exata(colunas_norm, ["FILIAL"]) or _achar_coluna(colunas_norm, [
        ("FILIAL",),
    ])
    col_pagador = _achar_coluna_exata(colunas_norm, ["PAGADOR"]) or _achar_coluna(colunas_norm, [
        ("PAGADOR",),
        ("CLIENTE",),
    ])

    faltando = []
    if not col_numero:
        faltando.append("numero")
    if not col_data:
        faltando.append("data_emissao")
    if not col_valor:
        faltando.append("valor")

    return {
        "tipo": col_tipo,
        "serie": col_serie,
        "numero": col_numero,
        "data": col_data,
        "data_ref": col_data_ref,
        "valor": col_valor,
        "frete": col_frete,
        "status": col_status,
        "filial": col_filial,
        "pagador": col_pagador,
        "faltando": faltando,
        "ordem_colunas": list(df.columns),
        "colunas_norm": colunas_norm,
    }



def _linha_valida_para_importacao(linha, mapa):
    def _normalizar_filial(valor_raw):
        if pd.isna(valor_raw):
            return ""
        txt = str(valor_raw).strip()
        # Excel pode trazer 88 como 88.0
        if re.fullmatch(r"\d+(?:\.0+)?", txt):
            try:
                return str(int(float(txt)))
            except ValueError:
                return ""
        return re.sub(r"\D", "", txt)

    # 1) Filial deve corresponder a config.FILIAL_PADRAO (padrão "88")
    if not mapa.get("filial"):
        return False
    filial_txt = _normalizar_filial(linha.get(mapa["filial"], ""))
    if filial_txt != config.FILIAL_PADRAO:
        return False

    # 2) Numero do documento obrigatorio no campo codigo/documento
    numero_txt = re.sub(r"\D", "", str(linha.get(mapa["numero"], "")))
    if not numero_txt:
        return False
    # Evita capturar datas como codigo (ex.: 27022026 / 20260227) e valores fora do padrao esperado.
    if len(numero_txt) > 6:
        return False
    if re.fullmatch(r"\d{8}", numero_txt):
        try:
            datetime.strptime(numero_txt, "%d%m%Y")
            return False
        except ValueError:
            pass
        try:
            datetime.strptime(numero_txt, "%Y%m%d")
            return False
        except ValueError:
            pass

    # 3) Data de emissao obrigatoria
    data_raw = linha.get(mapa["data"])
    data = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
    if pd.isna(data) and mapa.get("data_ref"):
        data = pd.to_datetime(linha.get(mapa["data_ref"]), dayfirst=True, errors="coerce")
    if pd.isna(data):
        return False

    # 4) Pagador obrigatoriamente Energisa
    if mapa.get("pagador"):
        pagador_txt = normalizar_texto(str(linha.get(mapa["pagador"], "")))
    else:
        pagador_txt = normalizar_texto(" ".join(str(v) for v in linha.tolist()))
    if "ENERGISA" not in pagador_txt:
        return False

    # 5) Ignora linhas de total/somatorio
    texto_linha = normalizar_texto(" ".join(str(v) for v in linha.tolist()))
    if any(chave in texto_linha for chave in [
        "TOTAL FRETE",
        "TOTAL FILIAL",
        "TOTAL C T R C",
        "TOTAL N F",
        "TOTAL FRETE N F",
        "TOTAL FRETE MARCO",
        "=>",
    ]):
        return False

    return True

def _inferir_tipo_documento_linha(linha, mapa):
    tipo_secao = normalizar_texto(str(linha.get("__secao_tipo", "")))
    if tipo_secao in {"NF", "CTE"}:
        return tipo_secao

    if mapa.get("serie"):
        serie_txt = re.sub(r"\D", "", str(linha.get(mapa["serie"], "")))
        if serie_txt == "1":
            return "NF"
        if serie_txt == "2":
            return "CTE"

    if mapa.get("tipo"):
        tipo = _parse_tipo_documento(linha.get(mapa["tipo"]))
        if tipo:
            return tipo

    texto_linha = normalizar_texto(" ".join(str(v) for v in linha.tolist()))
    if "NF" in texto_linha or "NOTA" in texto_linha:
        return "NF"
    if "CTE" in texto_linha or "CTRC" in texto_linha:
        return "CTE"

    return "CTE"


def _extrair_valor_frete_linha(linha, mapa):
    col_frete = mapa.get("frete")
    if not col_frete or col_frete not in linha.index:
        return None

    valor_raw = linha.get(col_frete)
    valor = parse_valor_monetario(valor_raw)
    if valor is None or valor <= 0:
        return None
    if valor > 100_000:
        return None
    return float(valor)


def _normalizar_nome_coluna_planilha(valor, idx_coluna):
    nome = str(valor).strip()
    if not nome or nome.lower() == "nan":
        return f"COL_{idx_coluna + 1}"
    return nome


def _linha_parece_cabecalho_planilha(row_vals):
    tokens = {_normalizar_coluna_relatorio(v) for v in row_vals}
    if "FILIAL" not in tokens:
        return False

    tem_codigo = any(t in {"CODIGO", "COD"} for t in tokens)
    tem_serie = "SERIE" in tokens
    tem_data = any(t.startswith("DATA") for t in tokens)
    tem_pagador = "PAGADOR" in tokens

    return tem_codigo and tem_serie and tem_data and tem_pagador


def _identificar_secao_planilha(row_vals):
    texto = normalizar_texto(" ".join(str(v) for v in row_vals))
    # Considera secao apenas na linha de titulo do bloco: "C.T.R.C. - janeiro / 26" ou "N.F. - janeiro / 26".
    # Aceita variacoes com/sem ponto final e com hifen simples ou longo.
    if not (("/" in texto) and (("-" in texto) or ("–" in texto))):
        return None
    if re.search(r"\bC\s*\.?\s*T\s*\.?\s*R\s*\.?\s*C\s*\.?\s*[-–]\s*[A-Z ]+\s*/\s*\d{2,4}", texto):
        return "CTE"
    if re.search(r"\bN\s*\.?\s*F\s*\.?\s*[-–]\s*[A-Z ]+\s*/\s*\d{2,4}", texto):
        return "NF"
    return None


def _linha_totalizadora_planilha(row_vals):
    texto = normalizar_texto(" ".join(str(v) for v in row_vals))
    return any(chave in texto for chave in [
        "TOTAL FRETE",
        "TOTAL FILIAL",
        "TOTAL C T R C",
        "TOTAL N F",
        "TOTAL FRETE N F",
        "TOTAL FRETE MARCO",
        "=>",
    ])


def _preparar_dataframe_planilha(df_raw):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    blocos = []
    i = 0
    total_linhas = len(df_raw.index)

    encontrou_secao = False
    while i < total_linhas:
        row_vals = df_raw.iloc[i].tolist()
        tipo_secao = _identificar_secao_planilha(row_vals)
        if not tipo_secao:
            i += 1
            continue
        encontrou_secao = True

        idx_cabecalho = None
        limite_busca = min(i + 8, total_linhas)
        for idx_tentativa in range(i + 1, limite_busca):
            if _linha_parece_cabecalho_planilha(df_raw.iloc[idx_tentativa].tolist()):
                idx_cabecalho = idx_tentativa
                break

        if idx_cabecalho is None:
            i += 1
            continue

        row_vals = df_raw.iloc[idx_cabecalho].tolist()
        headers = []
        usados = {}
        for idx_coluna, valor in enumerate(row_vals):
            nome_base = _normalizar_nome_coluna_planilha(valor, idx_coluna)
            if nome_base in usados:
                usados[nome_base] += 1
                nome = f"{nome_base}_{usados[nome_base]}"
            else:
                usados[nome_base] = 1
                nome = nome_base
            headers.append(nome)

        j = idx_cabecalho + 1
        linhas_bloco = []
        while j < total_linhas:
            prox_vals = df_raw.iloc[j].tolist()

            if _identificar_secao_planilha(prox_vals):
                break

            if _linha_totalizadora_planilha(prox_vals):
                break

            if not all(pd.isna(v) or str(v).strip() == "" for v in prox_vals):
                linhas_bloco.append(prox_vals)

            j += 1

        if linhas_bloco:
            df_bloco = pd.DataFrame(linhas_bloco, columns=headers)
            df_bloco = df_bloco.dropna(how="all")
            if not df_bloco.empty:
                df_bloco["__secao_tipo"] = tipo_secao
                blocos.append(df_bloco)

        i = j

    if blocos:
        return pd.concat(blocos, ignore_index=True, sort=False)

    # Fallback: se nao montar blocos por secao (ou secao estiver ausente),
    # usa blocos por cabecalho e para em totalizadores.
    if not blocos:
        i = 0
        while i < total_linhas:
            row_vals = df_raw.iloc[i].tolist()
            if not _linha_parece_cabecalho_planilha(row_vals):
                i += 1
                continue

            headers = []
            usados = {}
            for idx_coluna, valor in enumerate(row_vals):
                nome_base = _normalizar_nome_coluna_planilha(valor, idx_coluna)
                if nome_base in usados:
                    usados[nome_base] += 1
                    nome = f"{nome_base}_{usados[nome_base]}"
                else:
                    usados[nome_base] = 1
                    nome = nome_base
                headers.append(nome)

            j = i + 1
            linhas_bloco = []
            while j < total_linhas:
                prox_vals = df_raw.iloc[j].tolist()
                if _linha_parece_cabecalho_planilha(prox_vals):
                    break
                if _linha_totalizadora_planilha(prox_vals):
                    break
                if not all(pd.isna(v) or str(v).strip() == "" for v in prox_vals):
                    linhas_bloco.append(prox_vals)
                j += 1

            if linhas_bloco:
                df_bloco = pd.DataFrame(linhas_bloco, columns=headers)
                df_bloco = df_bloco.dropna(how="all")
                if not df_bloco.empty:
                    df_bloco["__secao_tipo"] = ""
                    blocos.append(df_bloco)

            i = j

        if blocos:
            return pd.concat(blocos, ignore_index=True, sort=False)

    return pd.DataFrame()

