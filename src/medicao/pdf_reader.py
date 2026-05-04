import os
from .utils import normalize
from .pdf_audit import analyze_pdf_file, evidence_note, pdf_content_enabled, tags_from_pdf_analysis


def read_pdf_text(filepath):
    try:
        import pdfplumber
        text = ''
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + '\n'
        return text.strip(), None
    except ImportError:
        pass
    except Exception as e:
        pass

    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(filepath)
        text = ''
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + '\n'
        return text.strip(), None
    except ImportError:
        return '', 'Biblioteca PDF não instalada'
    except Exception as e:
        return '', str(e)


def identify_doc_types(filepath):
    fn = normalize(os.path.basename(filepath))
    tags = set()

    # ADMISSÃO
    if any(k in fn for k in ['contrato de trabalho', 'contrato trabalho', 'contrato_trabalho']):
        tags.add('CONTRATO')
    if any(k in fn for k in ['ficha de registro', 'ficha registro', 'registro de empregado', 'regisdtro', 'registro_empregado']):
        tags.add('FICHA_REGISTRO')
    if 'ctpscontratos' in fn or ('ctps' in fn and 'contrato' in fn):
        tags.add('CTPS')
    if any(k in fn for k in ['esocial', 'e-social', 's-social', 'e social']):
        if any(k in fn for k in ['demissao', 'recisao', 'rescisao', 'baixa']):
            tags.add('ESOCIAL_DEMISSAO')
        else:
            tags.add('ESOCIAL')
    if 'aso' in fn:
        if any(k in fn for k in ['admissional', 'admissao']):
            tags.add('ASO_ADMISSIONAL')
        elif any(k in fn for k in ['demissional', 'demissao']):
            tags.add('ASO_DEMISSIONAL')
        elif any(k in fn for k in ['mudanca', 'risco', 'funcao']):
            tags.add('ASO_MUDANCA')
        else:
            tags.add('ASO')
    if any(k in fn for k in ['epi', 'entrega de fardas', 'entrega epi', 'fardas', 'entrega_epi']):
        tags.add('EPI')
    if any(k in fn for k in ['vale transporte', 'opcao de vale', 'opcao vale', 'termo de opcao', 'opcao_vt', 'termo opcao']):
        tags.add('VT_DECLARACAO')

    # DEMISSÃO
    if any(k in fn for k in ['termo de rescisao', 'termo de recisao', 'termo de recissao', 'trct', 'rescisao', 'recisao', 'recissao']) and 'quitacao' not in fn:
        tags.add('TRCT')
    if 'quitacao' in fn:
        tags.add('QUITACAO')
    if any(k in fn for k in ['aviso previo', 'aviso de dispensa', 'pedido de demissao', 'aviso_previo', 'dispensa do colaborador']):
        tags.add('AVISO_PREVIO')
    if any(k in fn for k in ['multa 40', 'multa de 40', 'grrf', 'multa fgts', 'comprovante da multa']):
        tags.add('GRRF_MULTA')
    if any(k in fn for k in ['seguro desemprego', 'requerimento do seguro', 'requerimento seguro']):
        tags.add('SEGURO_DESEMPREGO')
    if any(k in fn for k in ['homologacao', 'comunicado de homologacao', 'comunicado_homologacao']):
        tags.add('HOMOLOGACAO')
    if any(k in fn for k in ['comprovante de pagamento das verbas', 'verbas recisorias', 'verbas rescisorias', 'pagamento das verbas',
                              'comprovante de pagamento de recis', 'comprovante de recisao', 'comprovante de recissao',
                              'comprovante de rescisao', 'comprovante rescisao', 'comprovante recisao',
                              'comprovante recissao', 'pagamento de recisao', 'pagamento de recissao',
                              'pagamento de rescisao']):
        tags.add('COMPROVANTE_RESCISAO')
    if any(k in fn for k in ['gfd guia', 'gfd_guia']):
        tags.add('GRRF_MULTA')

    # FÉRIAS
    if 'aviso' in fn and any(k in fn for k in ['ferias']):
        tags.add('AVISO_FERIAS')
    if 'recibo' in fn and 'ferias' in fn:
        tags.add('RECIBO_FERIAS')
    if any(k in fn for k in ['liq ferias', 'liq. ferias', 'ferias acre', 'comprovante ferias', 'ferias_acre']):
        tags.add('COMPROVANTE_FERIAS')

    # FOPAG
    if any(k in fn for k in ['folha de pagamento', 'fopag', 'folha_pagamento']):
        tags.add('FOPAG')
    if any(k in fn for k in ['liq folha', 'liq. folha', 'liquido folha', 'liq_folha']):
        tags.add('COMPROVANTE_SALARIO')
    if any(k in fn for k in ['adiantamento', 'adto', 'liquido adiantamento', 'liq adiantamento']):
        tags.add('ADIANTAMENTO')
    if any(k in fn for k in ['saldo de salario', 'saldo_salario', '1 saldo']):
        tags.add('SALDO_SALARIO')

    # INSS + FGTS
    if 'dctfweb' in fn or ('dctf' in fn and ('declaracao' in fn or 'web' in fn)):
        if 'recibo' in fn:
            tags.add('RECIBO_DCTFWEB')
        else:
            tags.add('DCTFWEB')
    if 'resumo' in fn and 'debito' in fn:
        tags.add('RESUMO_DEBITOS')
    if 'resumo' in fn and 'credito' in fn:
        tags.add('RESUMO_CREDITOS')
    if any(k in fn for k in ['darf']) and any(k in fn for k in ['inss', 'previdenci']):
        tags.add('DARF_INSS')
    if 'crf' in fn or 'certificado de regularidade' in fn:
        tags.add('CRF')
    if 'fgts' in fn:
        if 'detalhamento' in fn or 'detalhe da guia' in fn or 'detalhe guia' in fn:
            tags.add('DETALHAMENTO_FGTS')
        elif 'comprovante' in fn or 'pagamento' in fn:
            tags.add('COMPROVANTE_FGTS')
        elif 'consignado' in fn:
            tags.add('CONSIGNADO_FGTS')
        else:
            tags.add('GUIA_FGTS')

    # VA/VR
    if any(k in fn for k in ['relatoriocolaboradores', 'relatorio colaboradores']):
        tags.add('RELATORIO_VAVR')
    if 'boleto' in fn and any(k in fn for k in ['cafe', 'cesta', 'vr', 'pluxee', 'cnpj']):
        tags.add('BOLETO_VAVR')
    if any(k in fn for k in ['nf-e', 'nota fiscal']) and any(k in fn for k in ['cafe', 'cesta', 'vr', 'alimentacao', 'refeicao']):
        tags.add('NF_VAVR')
    if 'comprovante' in fn and any(k in fn for k in ['pluxee', 'cafe', 'cesta', 'vr']):
        tags.add('COMPROVANTE_VAVR')

    # VT
    if any(k in fn for k in ['ricco', 'stage five', 'pedido detalhado', 'emissao pedido', 'ordem de compra', 'emissao_pedido']):
        tags.add('RELATORIO_VT')
    if 'boleto' in fn and any(k in fn for k in ['ricco', 'stage five', 'empresa1', 'transporte', 'vt']):
        tags.add('BOLETO_VT')
    if 'comprovante' in fn and any(k in fn for k in ['ricco', 'transporte', 'vt']):
        tags.add('COMPROVANTE_VT')
    if any(k in fn for k in ['nf 10', 'nf_10']) and 'vavr' not in str(tags):
        tags.add('NF_VT')

    # ACORDO COLETIVO
    if any(k in fn for k in ['cct', 'cct_acre', 'acordo coletivo', 'convencao coletiva', 'convencao_coletiva']):
        tags.add('ACORDO_COLETIVO')

    # DECLARAÇÕES
    if 'declaracao' in fn or 'declarac' in fn or 'decaracao' in fn:
        if any(k in fn for k in ['admissao', 'admissao-alocacao', 'alocacao', 'mobilizacao']):
            tags.add('DECL_ADMISSAO')
        if any(k in fn for k in ['demissao', 'demissoes']):
            tags.add('DECL_DEMISSAO')
        if 'ferias' in fn:
            tags.add('DECL_FERIAS')
        if any(k in fn for k in ['transferencia', 'transferencia de contrato']):
            tags.add('DECL_TRANSFERENCIA')
        if 'subcontratacao' in fn:
            tags.add('DECL_SUBCONTRATACAO')
        if any(k in fn for k in ['acidente', 'registro de acidente']):
            tags.add('DECL_ACIDENTE')
        if 'mobilizacao' in fn and 'admissao' not in fn:
            tags.add('DECL_MOBILIZACAO')
        if any(k in fn for k in ['mudanca de funcao', 'mudanca_funcao', 'mudanca de funcao']):
            tags.add('DECL_MUDANCA_FUNCAO')

    return tags


def get_pdf_evidence_in_folder(folder_path, recurse=False):
    evidence = {}
    if not pdf_content_enabled():
        return evidence
    if not os.path.isdir(folder_path):
        return evidence
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path) and item.lower().endswith('.pdf'):
            analysis = analyze_pdf_file(item_path)
            _, pdf_evidence = tags_from_pdf_analysis(analysis)
            for tag, items in pdf_evidence.items():
                evidence.setdefault(tag, []).extend(items)
        elif recurse and os.path.isdir(item_path):
            child = get_pdf_evidence_in_folder(item_path, recurse=True)
            for tag, items in child.items():
                evidence.setdefault(tag, []).extend(items)
    return evidence


def get_evidence_note(evidence_map, tag):
    items = evidence_map.get(tag) or []
    if not items:
        return ''
    best = sorted(
        items,
        key=lambda e: {'alta': 3, 'media': 2, 'baixa': 1}.get(e.get('confidence'), 0),
        reverse=True,
    )[0]
    return evidence_note(best)


def get_tags_for_files(file_paths):
    all_tags = set()
    for item_path in file_paths:
        if not os.path.isfile(item_path):
            continue
        all_tags |= identify_doc_types(item_path)
        if pdf_content_enabled() and item_path.lower().endswith('.pdf'):
            analysis = analyze_pdf_file(item_path)
            pdf_tags, _ = tags_from_pdf_analysis(analysis)
            all_tags |= pdf_tags
    return all_tags


def get_all_tags_in_folder(folder_path, recurse=False):
    all_tags = set()
    if not os.path.isdir(folder_path):
        return all_tags
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            all_tags |= identify_doc_types(item_path)
            if pdf_content_enabled() and item.lower().endswith('.pdf'):
                analysis = analyze_pdf_file(item_path)
                pdf_tags, _ = tags_from_pdf_analysis(analysis)
                all_tags |= pdf_tags
        elif recurse and os.path.isdir(item_path):
            all_tags |= get_all_tags_in_folder(item_path, recurse=True)
    return all_tags
