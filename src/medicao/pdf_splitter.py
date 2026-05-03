"""
pdf_splitter.py
Lê um PDF com múltiplos documentos, detecta os tipos de documento por
texto de cada página e separa em arquivos individuais.
Usa PyMuPDF (fitz). Páginas sem texto (scans) são processadas via
Windows Built-in OCR (sem dependência extra além do sistema operacional).
"""

import os
import re
import subprocess
import tempfile

# ── Padrões de detecção por texto da página ───────────────────────────────────
# Cada entrada: (doc_type, [lista de strings que identificam o início do doc])
# A busca é feita nos primeiros ~700 caracteres da página, case-insensitive.
# Sempre inclua versão COM e SEM acento (OCR pode variar).

_PATTERNS = [
    # ── ADMISSÃO ──────────────────────────────────────────────────────────────
    ('CONTRATO',            ['contrato de trabalho', 'contrato individual de trabalho']),

    ('FICHA_REGISTRO',      ['ficha de registro', 'registro de empregado', 'registro de empregados',
                             'ficha registro', 'livro ou ficha de registro']),

    ('CTPS',                ['carteira de trabalho', 'ctps -', 'ctps–', 'ctps n°', 'ctps nº']),

    ('ESOCIAL',             ['esocial', 'e-social', 's-2200', 's2200',
                             'admissão de trabalhador', 'admissao de trabalhador']),

    ('ASO_ADMISSIONAL',     ['aso admissional', 'exame admissional',
                             'atestado de saúde ocupacional admissional',
                             'atestado de saude ocupacional admissional']),

    ('ASO_DEMISSIONAL',     ['aso demissional', 'exame demissional',
                             'atestado de saúde ocupacional demissional',
                             'atestado de saude ocupacional demissional']),

    ('ASO_MUDANCA',         ['aso mudança', 'aso mudanca',
                             'atestado de saúde ocupacional periódico',
                             'atestado de saude ocupacional periodico',
                             'exame periódico', 'exame periodico',
                             'mudança de função', 'mudanca de funcao',
                             'mudança de risco']),

    ('ASO',                 ['atestado de saúde ocupacional', 'atestado de saude ocupacional']),

    ('EPI',                 ['ficha de entrega de epi', 'entrega de equipamento de proteção',
                             'entrega de equipamento de protecao',
                             'equipamento de proteção individual',
                             'equipamento de protecao individual',
                             'entrega de farda', 'recibo de epi']),

    # ── BENEFÍCIOS / DECLARAÇÕES DO COLABORADOR ───────────────────────────────
    ('VT_DECLARACAO',       ['declaração de opção de vale transporte',
                             'declaracao de opcao de vale transporte',
                             'termo de opção de vale transporte',
                             'termo de opcao de vale transporte',
                             'opção de vale-transporte', 'opcao de vale-transporte',
                             'opção pelo vale transporte', 'opcao pelo vale transporte',
                             'opto pela utilização do vale', 'opto pela utilizacao do vale',
                             'não opto pela utilização do vale', 'nao opto pela utilizacao do vale']),

    ('AUTORIZACAO_DESCONTOS', ['autorização de descontos', 'autorizacao de descontos',
                               'autorizo os descontos', 'autorizo a empresa a descontar',
                               'desconto em folha', 'art. 462 da clt',
                               'desconto no salário', 'desconto no salario']),

    ('INTERVALO_INTRAJORNADA', ['intervalo intrajornada', 'intrajornada',
                                'concessão do intervalo', 'concessao do intervalo',
                                'ciência da concessão', 'ciencia da concessao',
                                'termo de ciência', 'termo de ciencia',
                                'intervalo para descanso e almoço',
                                'intervalo para descanso e almoco']),

    ('LGPD',                ['lei geral de proteção de dados', 'lei geral de protecao de dados',
                             'lgpd', 'tratamento de dados pessoais',
                             'autorização para tratamento de dados',
                             'autorizacao para tratamento de dados',
                             'consentimento para uso de dados']),

    ('DECLARACAO_BENEFICIARIOS', ['declaração de beneficiários', 'declaracao de beneficiarios',
                                  'beneficiário do seguro', 'beneficiario do seguro']),

    ('SALARIO_FAMILIA',     ['salário família', 'salario familia', 'cota de salário família',
                             'cota de salario familia']),

    ('DECLARACAO_DEPENDENTES', ['declaração de dependentes', 'declaracao de dependentes',
                                'dependentes para irrf', 'imposto de renda retido na fonte',
                                'dependente para ir']),

    ('REGULAMENTO_INTERNO', ['regulamento interno', 'normas internas', 'código de conduta',
                             'codigo de conduta', 'manual do colaborador', 'politica interna']),

    ('ACUMULO_CARGOS',      ['acumulação de cargos', 'acumulacao de cargos',
                             'não acumulação', 'nao acumulacao',
                             'declaração de não acumulação', 'declaracao de nao acumulacao']),

    # ── DEMISSÃO ──────────────────────────────────────────────────────────────
    ('TRCT',                ['termo de rescisão do contrato de trabalho',
                             'termo de rescisao do contrato de trabalho',
                             'trct', 'termo de rescisao']),

    ('AVISO_PREVIO',        ['aviso prévio', 'aviso previo',
                             'pedido de demissão', 'pedido de demissao',
                             'comunicação de dispensa', 'comunicacao de dispensa']),

    ('SEGURO_DESEMPREGO',   ['seguro desemprego', 'seguro-desemprego',
                             'requerimento do seguro-desemprego',
                             'requerimento de seguro desemprego']),

    ('GRRF_MULTA',          ['grrf', 'guia para recolhimento rescisório',
                             'guia para recolhimento rescisorio',
                             'guia rescisória', 'guia rescisoria', 'gfd']),

    ('COMPROVANTE_RESCISAO', ['comprovante de pagamento da rescisão',
                              'comprovante de pagamento de rescisão',
                              'comprovante de pagamento da rescisao',
                              'liquido da rescisão', 'líquido da rescisão',
                              'liquido da rescisao']),

    ('ESOCIAL_DEMISSAO',    ['s-2299', 's2299', 'desligamento de trabalhador',
                             'baixa do esocial', 'rescisão esocial', 'rescisao esocial']),

    # ── FÉRIAS ────────────────────────────────────────────────────────────────
    ('AVISO_FERIAS',        ['aviso de férias', 'aviso de ferias']),

    ('RECIBO_FERIAS',       ['recibo de férias', 'recibo de ferias']),

    ('COMPROVANTE_FERIAS',  ['comprovante de pagamento de férias',
                             'comprovante de pagamento de ferias',
                             'férias acrescidas', 'ferias acrescidas',
                             'líquido férias', 'liquido ferias']),

    # ── FOLHA DE PAGAMENTO ────────────────────────────────────────────────────
    ('FOPAG',               ['folha de pagamento', 'demonstrativo de pagamento',
                             'contra cheque', 'contracheque', 'holerite']),

    ('COMPROVANTE_SALARIO',  ['comprovante de pagamento de salário',
                              'comprovante de pagamento de salario',
                              'comprovante de salario', 'líquido folha', 'liquido folha']),

    ('ADIANTAMENTO',        ['adiantamento salarial', 'adiantamento de salário',
                             'adiantamento de salario', 'adiantamento quinzenal']),

    # ── INSS / FGTS ───────────────────────────────────────────────────────────
    ('DCTFWEB',             ['dctfweb', 'declaração de débitos e créditos tributários federais',
                             'declaracao de debitos e creditos tributarios federais']),

    ('DARF_INSS',           ['darf', 'documento de arrecadação de receitas federais',
                             'documento de arrecadacao', 'previdência social', 'previdencia social']),

    ('GUIA_FGTS',           ['guia do fgts', 'fgts mensal', 'recolhimento fgts',
                             'guia de recolhimento do fgts']),

    ('DETALHAMENTO_FGTS',   ['detalhamento do fgts', 'detalhe da guia fgts', 'extrato fgts']),

    ('COMPROVANTE_FGTS',    ['comprovante fgts', 'pagamento fgts']),

    ('CRF',                 ['certificado de regularidade do fgts', 'crf -',
                             'certidão de regularidade', 'certidao de regularidade']),

    # ── VALE ALIMENTAÇÃO / VR / VT ────────────────────────────────────────────
    ('RELATORIO_VAVR',      ['relatório de vale alimentação', 'relatório de vale refeição',
                             'relatorio de vale alimentacao', 'relatorio de vale refeicao']),

    ('BOLETO_VAVR',         ['boleto vale', 'boleto alimentação', 'boleto refeição',
                             'boleto alimentacao', 'boleto refeicao']),

    ('RELATORIO_VT',        ['relatório de vale transporte', 'relatorio de vale transporte',
                             'espelho de vale transporte']),

    ('BOLETO_VT',           ['boleto vale transporte']),

    # ── SEGURO DE VIDA ────────────────────────────────────────────────────────
    ('APOLICE_SEGURO',      ['apólice de seguro', 'apolice de seguro',
                             'seguro de vida coletivo', 'bilhete de seguro']),

    ('COBRANCA_SEGURO',     ['fatura de seguro', 'cobrança de seguro',
                             'cobranca de seguro', 'boleto seguro de vida']),

    # ── PONTO ─────────────────────────────────────────────────────────────────
    ('PONTO',               ['folha de ponto', 'controle de ponto', 'espelho de ponto',
                             'registro de ponto', 'relatório de frequência',
                             'relatorio de frequencia']),

    # ── ACORDO COLETIVO ───────────────────────────────────────────────────────
    ('ACORDO_COLETIVO',     ['acordo coletivo de trabalho', 'convenção coletiva',
                             'convencao coletiva', 'cct -']),
]

# Nomes legíveis para cada tipo
_LABELS = {
    'CONTRATO':               'Contrato de Trabalho',
    'FICHA_REGISTRO':         'Ficha de Registro',
    'CTPS':                   'CTPS',
    'ESOCIAL':                'eSocial - Admissão',
    'ESOCIAL_DEMISSAO':       'eSocial - Demissão',
    'ASO_ADMISSIONAL':        'ASO Admissional',
    'ASO_DEMISSIONAL':        'ASO Demissional',
    'ASO_MUDANCA':            'ASO Mudança de Função',
    'ASO':                    'ASO',
    'EPI':                    'Ficha de EPI',
    'VT_DECLARACAO':          'Declaração de VT',
    'AUTORIZACAO_DESCONTOS':  'Autorização de Descontos',
    'INTERVALO_INTRAJORNADA': 'Termo de Intervalo Intrajornada',
    'LGPD':                   'Termo LGPD',
    'DECLARACAO_BENEFICIARIOS': 'Declaração de Beneficiários',
    'SALARIO_FAMILIA':        'Salário Família',
    'DECLARACAO_DEPENDENTES': 'Declaração de Dependentes',
    'REGULAMENTO_INTERNO':    'Regulamento Interno',
    'ACUMULO_CARGOS':         'Declaração de Não Acumulação de Cargos',
    'TRCT':                   'TRCT - Termo de Rescisão',
    'AVISO_PREVIO':           'Aviso Prévio',
    'SEGURO_DESEMPREGO':      'Seguro Desemprego',
    'GRRF_MULTA':             'GRRF - Multa FGTS',
    'COMPROVANTE_RESCISAO':   'Comprovante de Rescisão',
    'AVISO_FERIAS':           'Aviso de Férias',
    'RECIBO_FERIAS':          'Recibo de Férias',
    'COMPROVANTE_FERIAS':     'Comprovante de Férias',
    'FOPAG':                  'Folha de Pagamento',
    'COMPROVANTE_SALARIO':    'Comprovante de Salário',
    'ADIANTAMENTO':           'Adiantamento',
    'DCTFWEB':                'DCTFWeb',
    'DARF_INSS':              'DARF INSS',
    'GUIA_FGTS':              'Guia FGTS',
    'DETALHAMENTO_FGTS':      'Detalhamento FGTS',
    'COMPROVANTE_FGTS':       'Comprovante FGTS',
    'CRF':                    'CRF FGTS',
    'RELATORIO_VAVR':         'Relatório VA e VR',
    'BOLETO_VAVR':            'Boleto VA e VR',
    'RELATORIO_VT':           'Relatório VT',
    'BOLETO_VT':              'Boleto VT',
    'APOLICE_SEGURO':         'Apólice Seguro de Vida',
    'COBRANCA_SEGURO':        'Cobrança Seguro de Vida',
    'PONTO':                  'Folha de Ponto',
    'ACORDO_COLETIVO':        'Acordo Coletivo',
    'DESCONHECIDO':           'Documento Desconhecido',
}

# ── OCR via Windows Built-in Engine ──────────────────────────────────────────

_PS_OCR_SCRIPT = r"""
param([string]$ImagePath)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]

function Await($WinRtTask) {
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() |
               Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod }
    $m = $methods | Select-Object -First 1
    $gm = $m.MakeGenericMethod($WinRtTask.GetType().GetGenericArguments()[0])
    $task = $gm.Invoke($null, @($WinRtTask))
    $task.Wait()
    return $task.Result
}

try {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        $lang = [Windows.Globalization.Language]::new('pt-BR')
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    }
    if ($null -eq $engine) { exit 1 }

    $absPath = [System.IO.Path]::GetFullPath($ImagePath)
    $fileTask = [Windows.Storage.StorageFile]::GetFileFromPathAsync($absPath)
    $file = Await $fileTask

    $streamTask = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
    $stream = Await $streamTask

    $decoderTask = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
    $decoder = Await $decoderTask

    $bitmapTask = $decoder.GetSoftwareBitmapAsync()
    $bitmap = Await $bitmapTask

    $ocrTask = $engine.RecognizeAsync($bitmap)
    $result = Await $ocrTask

    Write-Output $result.Text
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
"""


def _ocr_page_windows(img_path):
    """Usa o OCR nativo do Windows (via PowerShell) para extrair texto de uma imagem PNG."""
    try:
        flags = 0
        if hasattr(subprocess, 'CREATE_NO_WINDOW'):
            flags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive',
             '-Command', _PS_OCR_SCRIPT, '-ImagePath', img_path],
            capture_output=True, text=True, timeout=90,
            creationflags=flags
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except Exception:
        pass
    return ''


def _get_page_text(page, dpi=200):
    """
    Extrai texto de uma página fitz.
    Se a página não tiver texto (scan), usa Windows OCR sobre a imagem renderizada.
    """
    text = page.get_text().strip()
    if len(text) > 40:      # texto suficiente → usa direto
        return text

    # Página provavelmente é imagem → renderiza e faz OCR
    try:
        mat = page.get_pixmap(dpi=dpi)
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            tmp_path = tmp.name
        mat.save(tmp_path)
        try:
            ocr_text = _ocr_page_windows(tmp_path)
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
        return ocr_text or text
    except Exception:
        return text


# ── Detecção de tipo de documento ────────────────────────────────────────────

def detect_doc_type(text):
    """Retorna o tipo de documento detectado no texto, ou None se não reconhecido."""
    t = text.lower()[:700]
    for doc_type, keywords in _PATTERNS:
        for kw in keywords:
            if kw in t:
                return doc_type
    return None


def get_label(doc_type):
    return _LABELS.get(doc_type, doc_type or 'Documento Desconhecido')


# ── Análise principal ─────────────────────────────────────────────────────────

def analyze_pdf(pdf_path, progress_cb=None):
    """
    Analisa um PDF e retorna uma lista de grupos detectados.
    Cada grupo: {'type': str, 'label': str, 'pages': [int, ...], 'start_text': str}

    progress_cb: função opcional chamada com (pagina_atual, total_paginas).
    Requer PyMuPDF (fitz).
    """
    import fitz

    doc = fitz.open(pdf_path)
    total = len(doc)

    page_infos = []
    for i in range(total):
        if progress_cb:
            progress_cb(i + 1, total)
        page = doc[i]
        text = _get_page_text(page)
        detected = detect_doc_type(text)
        page_infos.append({'page': i, 'type': detected, 'text': text[:300].strip()})

    doc.close()

    if not page_infos:
        return []

    # ── Agrupa páginas ────────────────────────────────────────────────────────
    # Regras:
    #  1. Página com tipo RECONHECIDO diferente do grupo atual → NOVO grupo
    #  2. Página com tipo RECONHECIDO igual ao grupo atual    → continua
    #  3. Página SEM tipo (None) → continua no grupo atual, OU cria
    #     "Documento Desconhecido" se não há nenhum grupo ainda
    #
    # Heurística extra: se a página é sem tipo MAS tem texto significativo
    # E o grupo atual já tem ≥1 páginas detectadas, cria um novo DESCONHECIDO.
    # Isso evita fundir documentos de tipos não cadastrados com o anterior.

    groups = []
    current = None

    for pi in page_infos:
        detected = pi['type']
        has_text = len(pi['text']) > 30

        if detected is not None:
            # Tipo reconhecido
            if current is None or detected != current['type']:
                # Novo documento
                current = {
                    'type': detected,
                    'label': get_label(detected),
                    'pages': [pi['page']],
                    'start_text': pi['text'],
                }
                groups.append(current)
            else:
                # Continuação (mesmo tipo)
                current['pages'].append(pi['page'])
        else:
            # Tipo não reconhecido
            if current is None:
                # Nenhum grupo aberto → cria desconhecido
                current = {
                    'type': 'DESCONHECIDO',
                    'label': 'Documento Desconhecido',
                    'pages': [pi['page']],
                    'start_text': pi['text'],
                }
                groups.append(current)
            elif current['type'] == 'DESCONHECIDO' or not has_text:
                # Continua no grupo atual (sem texto = segunda página do mesmo doc)
                current['pages'].append(pi['page'])
            else:
                # Tem texto mas não reconhecemos o tipo → nova seção desconhecida
                # (provavelmente outro documento cujo tipo não está cadastrado)
                current = {
                    'type': 'DESCONHECIDO',
                    'label': 'Documento Desconhecido',
                    'pages': [pi['page']],
                    'start_text': pi['text'],
                }
                groups.append(current)

    return groups


def split_pdf(pdf_path, output_folder, groups, employee_name=''):
    """
    Salva cada grupo como um PDF separado na output_folder.
    Retorna lista de (caminho_salvo, label).
    """
    import fitz

    os.makedirs(output_folder, exist_ok=True)
    doc = fitz.open(pdf_path)
    saved = []
    name_count = {}

    for group in groups:
        label = group.get('label') or get_label(group.get('type', ''))
        pages = group.get('pages', [])
        if not pages:
            continue

        prefix = employee_name.strip().upper() + ' - ' if employee_name.strip() else ''
        base_name = f"{prefix}{label}"

        name_count[base_name] = name_count.get(base_name, 0) + 1
        if name_count[base_name] > 1:
            base_name = f"{base_name} ({name_count[base_name]})"

        safe_name = re.sub(r'[\\/:*?"<>|]', '_', base_name)
        out_path = os.path.join(output_folder, safe_name + '.pdf')

        new_doc = fitz.open()
        for pg in sorted(set(pages)):
            if 0 <= pg < len(doc):
                new_doc.insert_pdf(doc, from_page=pg, to_page=pg)

        new_doc.save(out_path)
        new_doc.close()
        saved.append((out_path, label))

    doc.close()
    return saved
