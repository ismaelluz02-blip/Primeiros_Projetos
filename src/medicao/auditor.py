import os
import re
from datetime import date
from .utils import (normalize, load_config, find_folder, find_files_by_keywords,
                   list_subfolders, has_any_file, employee_matches_folder)
from .excel_reader import read_forca_trabalho, get_month_employees
from .pdf_reader import (
    identify_doc_types,
    get_all_tags_in_folder,
    get_evidence_note,
    get_pdf_evidence_in_folder,
    get_tags_for_files,
)
from .pdf_audit import configure_pdf_audit, get_pdf_audit_diagnostics

FOLDER_PATTERNS = {
    'acordo_coletivo': ['acordo coletivo'],
    'admissao':        ['admissao-alocacao', 'admissao alocacao', 'admissao'],
    'declaracoes':     ['declaracoes', 'declaracao'],
    'demissao':        ['demissao-transferencia', 'demissao transferencia', 'demissao'],
    'ferias':          ['ferias'],
    'fopag':           ['fopag e comp', 'fopag'],
    'inss_fgts':       ['inss + fgts', 'inss+fgts', 'inss fgts', 'inss'],
    'parcelamentos':   ['parcelamentos'],
    'ponto':           ['ponto'],
    'va_vr':           ['va e vr', 'va vr', 'va_vr'],
    'vt':              ['vt'],
    'seguro_vida':     ['seguro de vida', 'seguro vida'],
}

MONTH_NAMES = {
    1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL',
    5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO',
    9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'
}


def _item(label, status, note=''):
    return {'label': label, 'status': status, 'note': note}


def _issue(msg, section=''):
    return {'msg': msg, 'section': section}


def _evidence(evidence_map, *tags):
    for tag in tags:
        note = get_evidence_note(evidence_map, tag)
        if note:
            return note
    return ''


def _tags_from_pending_text(text):
    n = normalize(text)
    checks = [
        ('CONTRATO', ['contrato de trabalho', 'contrato']),
        ('FICHA_REGISTRO', ['ficha de registro', 'registro de empregado']),
        ('CTPS', ['ctps']),
        ('ESOCIAL', ['esocial', 'e social', 'baixa digital']),
        ('ASO_ADMISSIONAL', ['aso admissional']),
        ('ASO_DEMISSIONAL', ['aso demissional']),
        ('ASO_MUDANCA', ['aso mudanca', 'mudanca de risco']),
        ('ASO', ['aso']),
        ('EPI', ['epi']),
        ('VT_DECLARACAO', ['declaracao de opcao vt', 'vale transporte']),
        ('TRCT', ['termo de rescisao', 'trct']),
        ('COMPROVANTE_RESCISAO', ['comprovante de pagamento da rescisao']),
        ('AVISO_PREVIO', ['aviso previo', 'pedido de demissao']),
        ('FGTS_GRRF', ['fgts grrf', 'grrf', 'multa 40']),
        ('SEGURO_DESEMPREGO', ['seguro desemprego']),
        ('RECIBO_FERIAS', ['recibo de ferias']),
        ('AVISO_FERIAS', ['aviso de ferias']),
        ('COMPROVANTE_FERIAS', ['comprovante de pagamento de ferias']),
        ('FOPAG', ['folha de pagamento analitica', 'fopag']),
        ('COMPROVANTE_SALARIO', ['comprovante de pagamento de salario', 'saldo salario', 'saldo/salario']),
        ('GUIA_FGTS', ['guia fgts']),
        ('COMPROVANTE_FGTS', ['comprovante de pagamento fgts']),
        ('DETALHAMENTO_FGTS', ['detalhamento da guia fgts']),
        ('CRF', ['certificado de regularidade fgts', 'crf']),
        ('DCTFWEB', ['dctfweb']),
        ('DARF_INSS', ['darf inss']),
    ]
    return [tag for tag, terms in checks if any(term in n for term in terms)]


def _best_global_evidence(evidence_map, tags, person=''):
    candidates = []
    person_norm = normalize(person)
    for tag in tags:
        for item in evidence_map.get(tag, []) or []:
            path = item.get('filePath') or item.get('fileName') or ''
            if person_norm and not employee_matches_folder(person, path):
                continue
            confidence_score = {'alta': 3, 'media': 2, 'baixa': 1}.get(item.get('confidence'), 0)
            candidates.append((confidence_score, item, tag))
    if not candidates:
        return None, None
    candidates.sort(key=lambda row: row[0], reverse=True)
    return candidates[0][1], candidates[0][2]


def _filename_evidence_in_folder(folder_path):
    evidence = {}
    if not os.path.isdir(folder_path):
        return evidence
    for root, _dirs, files in os.walk(folder_path):
        for filename in files:
            item_path = os.path.join(root, filename)
            if not os.path.isfile(item_path) or not filename.lower().endswith('.pdf'):
                continue
            for tag in identify_doc_types(item_path):
                evidence.setdefault(tag, []).append({
                    'documentType': tag,
                    'documentName': tag,
                    'filePath': item_path,
                    'fileName': filename,
                    'pages': [],
                    'confidence': 'alta',
                    'method': 'nome do arquivo',
                    'matchedKeywords': [],
                    'source': 'nome do arquivo',
                })
    return evidence


def _merge_evidence_maps(*maps):
    merged = {}
    seen = set()
    for evidence_map in maps:
        for tag, items in (evidence_map or {}).items():
            for item in items or []:
                key = (tag, item.get('filePath') or item.get('fileName'), tuple(item.get('pages') or []), item.get('method'))
                if key in seen:
                    continue
                seen.add(key)
                merged.setdefault(tag, []).append(item)
    return merged


def _resolve_pendencias_por_busca_profunda(sections, folder_path):
    evidence = _merge_evidence_maps(
        _filename_evidence_in_folder(folder_path),
        get_pdf_evidence_in_folder(folder_path, recurse=True),
    )
    if not evidence:
        return

    def _resolve_item(item, section_name='', person=''):
        if item.get('status') not in ('error', 'warning'):
            return False, []
        text = f"{item.get('label', '')} {item.get('note', '')}"
        tags = _tags_from_pending_text(text)
        found, found_tag = _best_global_evidence(evidence, tags, person=person)
        if not found:
            return False, tags
        note = get_evidence_note({found_tag: [found]}, found_tag)
        where = note or f"Encontrado em {found.get('fileName') or 'PDF da competencia'}"
        old_note = item.get('note', '')
        prefix = "Busca profunda encontrou evidência fora da conferência inicial"
        item['status'] = 'ok'
        item['note'] = f"{prefix}: {where}" if not old_note or old_note in ('FALTANDO', 'Não encontrado', 'Não identificado') else f"{old_note} | {prefix}: {where}"
        item['deep_resolved'] = True
        return True, tags

    def _issue_matches(issue_text, tags):
        issue_tags = set(_tags_from_pending_text(issue_text))
        return bool(issue_tags.intersection(tags))

    for sec in sections:
        resolved_tags = []
        for item in sec.get('items', []):
            resolved, tags = _resolve_item(item, section_name=sec.get('name', ''))
            if resolved:
                resolved_tags.extend(tags)

        for emp in sec.get('employees', []):
            emp_resolved_tags = []
            for item in emp.get('items', []):
                resolved, tags = _resolve_item(item, section_name=sec.get('name', ''), person=emp.get('name', ''))
                if resolved:
                    emp_resolved_tags.extend(tags)
                    resolved_tags.extend(tags)
            if emp_resolved_tags:
                emp['issues'] = [
                    issue for issue in emp.get('issues', [])
                    if not _issue_matches(issue, emp_resolved_tags)
                ]
                emp['status'] = 'error' if any(i.get('status') == 'error' for i in emp.get('items', [])) else \
                                'warning' if any(i.get('status') == 'warning' for i in emp.get('items', [])) else 'ok'

        if resolved_tags:
            sec['issues'] = [
                issue for issue in sec.get('issues', [])
                if not _issue_matches(issue.get('msg', '') if isinstance(issue, dict) else issue, resolved_tags)
            ]
            has_emp_error = any(emp.get('status') == 'error' for emp in sec.get('employees', []))
            has_emp_warning = any(emp.get('status') == 'warning' for emp in sec.get('employees', []))
            sec['status'] = 'error' if sec.get('issues') or any(i.get('status') == 'error' for i in sec.get('items', [])) or has_emp_error else \
                            'warning' if any(i.get('status') == 'warning' for i in sec.get('items', [])) or has_emp_warning else 'ok'


def _join_scope(scope, msg):
    return f'{scope}: {msg}' if scope else msg


def detect_competencia(folder_path):
    folder_name = os.path.basename(folder_path).upper()
    for num, name in MONTH_NAMES.items():
        if name in folder_name:
            year_match = re.search(r'(20\d{2})', folder_path)
            year = int(year_match.group(1)) if year_match else date.today().year
            return num, year, name
    # Try parent path
    parts = folder_path.replace('\\', '/').split('/')
    for part in reversed(parts):
        part_up = part.upper()
        for num, name in MONTH_NAMES.items():
            if name in part_up:
                year_match = re.search(r'(20\d{2})', folder_path)
                year = int(year_match.group(1)) if year_match else date.today().year
                return num, year, name
    return None, None, 'DESCONHECIDO'


def find_forca_trabalho(folder_path):
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            item_norm = normalize(item)
            if any(k in item_norm for k in ['forca de trabalho', 'forca_trabalho', 'força de trabalho']):
                if item.endswith(('.xlsx', '.xls')):
                    return item_path
    return None


def audit_acordo_coletivo(folder_path):
    section = {'name': 'Acordo Coletivo', 'icon': '📄', 'items': [], 'issues': []}
    ac_folder = find_folder(folder_path, FOLDER_PATTERNS['acordo_coletivo'])
    if not ac_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "Acordo Coletivo" não encontrada', 'Acordo Coletivo'))
        section['items'].append(_item('Pasta Acordo Coletivo', 'error', 'Não encontrada'))
        return section
    files = [f for f in os.listdir(ac_folder) if os.path.isfile(os.path.join(ac_folder, f))]
    if not files:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "Acordo Coletivo" está vazia', 'Acordo Coletivo'))
        section['items'].append(_item('Documento CCT/Acordo', 'error', 'Nenhum arquivo encontrado'))
    else:
        section['status'] = 'ok'
        for f in files:
            section['items'].append(_item(f, 'ok'))
    return section


def audit_declaracoes(folder_path, has_admissao, has_demissao, has_ferias, has_troca_funcao, config):
    section = {'name': 'Declarações', 'icon': '📋', 'items': [], 'issues': []}
    decl_folder = find_folder(folder_path, FOLDER_PATTERNS['declaracoes'])
    if not decl_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "Declarações" não encontrada', 'Declarações'))
        section['items'].append(_item('Pasta Declarações', 'error', 'Não encontrada'))
        return section

    all_tags = get_all_tags_in_folder(decl_folder)
    evidence = get_pdf_evidence_in_folder(decl_folder)
    files_in_folder = [f for f in os.listdir(decl_folder) if os.path.isfile(os.path.join(decl_folder, f))]

    required = [
        ('DECL_ACIDENTE', 'Declaração de Acidente de Trabalho', True),
        ('DECL_SUBCONTRATACAO', 'Declaração de Subcontratação', True),
        ('DECL_ADMISSAO', 'Declaração de Admissão/Alocação', has_admissao),
        ('DECL_DEMISSAO', 'Declaração de Demissão', has_demissao),
        ('DECL_FERIAS', 'Declaração de Férias', has_ferias),
        ('DECL_TRANSFERENCIA', 'Declaração de Transferência', False),
        ('DECL_MOBILIZACAO', 'Declaração de Mobilização', False),
        ('DECL_MUDANCA_FUNCAO', 'Declaração de Mudança de Função', has_troca_funcao),
    ]

    issues = []
    for tag, label, is_required in required:
        if tag in all_tags:
            section['items'].append(_item(label, 'ok', _evidence(evidence, tag)))
        elif is_required:
            section['items'].append(_item(label, 'error', 'FALTANDO'))
            issues.append(_issue(f'Declaração faltando: {label}', 'Declarações'))
        else:
            section['items'].append(_item(label, 'info', 'Não se aplica no mês'))

    # List all files found
    section['files_found'] = files_in_folder
    section['issues'] = issues
    section['status'] = 'error' if issues else 'ok'
    return section


def audit_admissao_person(person_folder, person_name, config):
    result = {'name': person_name, 'items': [], 'issues': []}
    tags = get_all_tags_in_folder(person_folder, recurse=True)
    evidence = get_pdf_evidence_in_folder(person_folder, recurse=True)
    files = [f for f in os.listdir(person_folder) if os.path.isfile(os.path.join(person_folder, f))]

    ponto_exempt = [normalize(e) for e in config.get('ponto_exempt', [])]
    epi_exempt_names = [normalize(e) for e in config.get('epi_exempt', [])]
    esocial_ok = config.get('esocial_substitutes_ctps', True)
    person_norm = normalize(person_name)

    checks = [
        ('CONTRATO', 'Contrato de trabalho', True),
        ('FICHA_REGISTRO', 'Ficha de registro', True),
    ]
    for tag, label, required in checks:
        if tag in tags:
            result['items'].append(_item(label, 'ok', _evidence(evidence, tag)))
        elif required:
            result['items'].append(_item(label, 'error', 'FALTANDO'))
            result['issues'].append(f'{label} não encontrado')

    # CTPS / eSocial
    has_ctps = 'CTPS' in tags
    has_esocial = 'ESOCIAL' in tags
    if has_ctps or has_esocial:
        label = 'CTPS Digital / eSocial'
        note = _evidence(evidence, 'CTPS', 'ESOCIAL')
        if not note and has_esocial and not has_ctps and esocial_ok:
            note = '(eSocial aceito)'
        result['items'].append(_item(label, 'ok', note))
    else:
        result['items'].append(_item('CTPS Digital / eSocial', 'error', 'FALTANDO'))
        result['issues'].append('CTPS/eSocial não encontrado')

    # ASO
    has_aso = any(t in tags for t in ['ASO_ADMISSIONAL', 'ASO'])
    if has_aso:
        result['items'].append(_item('ASO admissional', 'ok', _evidence(evidence, 'ASO_ADMISSIONAL', 'ASO')))
    else:
        result['items'].append(_item('ASO admissional', 'error', 'FALTANDO'))
        result['issues'].append('ASO admissional não encontrado')

    # EPI — match by any word overlap between exempt list and folder name
    def _name_overlaps(a_norm, b_norm):
        words_a = [w for w in a_norm.split() if len(w) > 2]
        return any(w in b_norm for w in words_a)

    is_epi_exempt = any(_name_overlaps(e, person_norm) for e in epi_exempt_names)
    has_epi = 'EPI' in tags
    if has_epi:
        result['items'].append(_item('Ficha de entrega de EPI', 'ok', _evidence(evidence, 'EPI')))
    elif is_epi_exempt:
        result['items'].append(_item('Ficha de entrega de EPI', 'info', 'Isento por regra operacional'))
    else:
        result['items'].append(_item('Ficha de entrega de EPI', 'warning', 'Não encontrado'))
        result['issues'].append('Ficha de EPI não encontrada (verificar se é obrigatório)')

    # VT declaration
    has_vt_decl = 'VT_DECLARACAO' in tags
    if has_vt_decl:
        result['items'].append(_item('Declaração de opção VT', 'ok', _evidence(evidence, 'VT_DECLARACAO')))
    else:
        result['items'].append(_item('Declaração de opção VT', 'warning', 'Não encontrado'))
        result['issues'].append('Declaração de opção VT não encontrada')

    result['status'] = 'error' if any('FALTANDO' in i for i in result['issues']) else \
                       'warning' if result['issues'] else 'ok'
    return result


def audit_admissoes(folder_path, admitted_employees, config):
    section = {'name': 'Admissões', 'icon': '🆕', 'employees': [], 'issues': []}
    admissao_folder = find_folder(folder_path, FOLDER_PATTERNS['admissao'])

    if not admissao_folder:
        if admitted_employees:
            section['status'] = 'error'
            section['issues'].append(_issue('Pasta "Admissão-alocação" não encontrada mas há admissões na planilha', 'Admissões'))
        else:
            section['status'] = 'info'
            section['items'] = [_item('Sem admissões no mês', 'info')]
        return section

    subfolders = list_subfolders(admissao_folder)

    # Separate regular admissions from troca de função
    admission_folders = []
    troca_funcao_folders = []
    for sf in subfolders:
        sf_norm = normalize(os.path.basename(sf))
        if 'troca de funcao' in sf_norm or 'troca funcao' in sf_norm or 'mudanca de funcao' in sf_norm:
            # It's a troca de função parent - look inside
            for sub in list_subfolders(sf):
                troca_funcao_folders.append(sub)
        else:
            admission_folders.append(sf)

    # Audit each admission subfolder
    emp_results = []
    for sf in admission_folders:
        person_name = os.path.basename(sf)
        result = audit_admissao_person(sf, person_name, config)
        emp_results.append(result)

    # Audit troca de função subfolders
    troca_results = []
    for sf in troca_funcao_folders:
        person_name = os.path.basename(sf)
        tags = get_all_tags_in_folder(sf, recurse=True)
        evidence = get_pdf_evidence_in_folder(sf, recurse=True)
        result = {'name': person_name + ' (Mudança de função)', 'items': [], 'issues': []}
        has_aso = any(t in tags for t in ['ASO_MUDANCA', 'ASO'])
        if has_aso:
            result['items'].append(_item('ASO mudança de risco', 'ok', _evidence(evidence, 'ASO_MUDANCA', 'ASO')))
        else:
            result['items'].append(_item('ASO mudança de risco', 'warning', 'Não encontrado'))
            result['issues'].append('ASO de mudança de risco não encontrado')
        result['status'] = 'warning' if result['issues'] else 'ok'
        troca_results.append(result)

    section['employees'] = emp_results + troca_results
    has_issues = any(r['status'] in ('error', 'warning') for r in section['employees'])
    section['status'] = 'error' if any(r['status'] == 'error' for r in section['employees']) else \
                        'warning' if has_issues else 'ok'

    for r in section['employees']:
        for issue in r['issues']:
            section['issues'].append(_issue(f"{r['name']}: {issue}", 'Admissões'))

    return section, troca_funcao_folders


def audit_demissao_person(person_folder, person_name, config, employees):
    result = {'name': person_name, 'items': [], 'issues': []}
    tags = get_all_tags_in_folder(person_folder, recurse=True)
    evidence = get_pdf_evidence_in_folder(person_folder, recurse=True)

    # Check if ASO demissional can be exempt (admission < 3 months ago)
    aso_exempt = False
    exempt_months = config.get('aso_demissional_exempt_months', 3)
    person_norm = normalize(person_name)
    for emp in employees:
        if employee_matches_folder(emp['nome'], person_name):
            if emp.get('data_admissao') and emp.get('data_demissao'):
                diff_days = (emp['data_demissao'] - emp['data_admissao']).days
                if diff_days < exempt_months * 30:
                    aso_exempt = True
            break

    # TRCT
    if 'TRCT' in tags:
        result['items'].append(_item('Termo de rescisão (TRCT)', 'ok'))
    else:
        result['items'].append(_item('Termo de rescisão (TRCT)', 'error', 'FALTANDO'))
        result['issues'].append('Termo de rescisão (TRCT) não encontrado')

    # Comprovante pagamento rescisão
    if 'COMPROVANTE_RESCISAO' in tags:
        result['items'].append(_item('Comprovante de pagamento da rescisão', 'ok'))
    else:
        result['items'].append(_item('Comprovante de pagamento da rescisão', 'error', 'FALTANDO'))
        result['issues'].append('Comprovante de pagamento da rescisão não encontrado')

    # Aviso prévio - warning, not error (can be absent in fixed-term contracts)
    if 'AVISO_PREVIO' in tags:
        result['items'].append(_item('Aviso prévio / Pedido de demissão', 'ok'))
    else:
        result['items'].append(_item('Aviso prévio / Pedido de demissão', 'warning', 'Não encontrado (verifique tipo de contrato)'))
        result['issues'].append('Aviso prévio não encontrado — verificar se se aplica')

    # FGTS/GRRF — accept GRRF_MULTA, or GUIA_FGTS+COMPROVANTE_FGTS combo (rescisão sem justa causa vs prazo det.)
    has_grrf = 'GRRF_MULTA' in tags
    has_fgts_combo = ('GUIA_FGTS' in tags or 'DETALHAMENTO_FGTS' in tags) and 'COMPROVANTE_FGTS' not in tags
    has_fgts_pay = any(t in tags for t in ['GUIA_FGTS', 'COMPROVANTE_FGTS', 'DETALHAMENTO_FGTS'])
    if has_grrf:
        result['items'].append(_item('FGTS / GRRF (multa 40%)', 'ok'))
    elif has_fgts_pay:
        result['items'].append(_item('FGTS rescisório (guia/comprovante)', 'ok', 'Aceito como FGTS de rescisão'))
    else:
        result['items'].append(_item('FGTS / GRRF', 'error', 'FALTANDO'))
        result['issues'].append('FGTS/GRRF da rescisão não encontrado')

    # Seguro desemprego - warning, not error (absent in fixed-term contracts)
    if 'SEGURO_DESEMPREGO' in tags:
        result['items'].append(_item('Protocolo de Seguro Desemprego', 'ok'))
    else:
        result['items'].append(_item('Protocolo de Seguro Desemprego', 'warning', 'Não encontrado (verifique tipo de contrato)'))
        result['issues'].append('Seguro desemprego não encontrado — verificar se se aplica')

    # CTPS baixa / eSocial
    has_esocial_dem = any(t in tags for t in ['ESOCIAL_DEMISSAO', 'ESOCIAL'])
    if has_esocial_dem:
        result['items'].append(_item('eSocial / Baixa digital', 'ok'))
    else:
        result['items'].append(_item('eSocial / Baixa digital', 'warning', 'Não encontrado'))
        result['issues'].append('eSocial/baixa digital não encontrado')

    # ASO demissional
    has_aso_dem = any(t in tags for t in ['ASO_DEMISSIONAL', 'ASO'])
    if has_aso_dem:
        result['items'].append(_item('ASO demissional', 'ok', _evidence(evidence, 'ASO_DEMISSIONAL', 'ASO')))
    elif aso_exempt:
        result['items'].append(_item('ASO demissional', 'info', f'Dispensado — ASO admissional < {exempt_months} meses'))
    else:
        result['items'].append(_item('ASO demissional', 'warning', 'Não encontrado'))
        result['issues'].append('ASO demissional não encontrado')

    result['status'] = 'error' if any('FALTANDO' in i for i in result['issues']) else \
                       'warning' if result['issues'] else 'ok'
    return result


def audit_demissoes(folder_path, dismissed_employees, config, all_employees):
    section = {'name': 'Demissões', 'icon': '🔚', 'employees': [], 'issues': []}
    dem_folder = find_folder(folder_path, FOLDER_PATTERNS['demissao'])

    if not dem_folder:
        if dismissed_employees:
            section['status'] = 'error'
            section['issues'].append(_issue('Pasta "Demissão-transferência" não encontrada mas há demissões na planilha', 'Demissões'))
        else:
            section['status'] = 'info'
            section['items'] = [_item('Sem demissões no mês', 'info')]
        return section

    subfolders = list_subfolders(dem_folder)
    emp_results = []
    for sf in subfolders:
        person_name = os.path.basename(sf)
        result = audit_demissao_person(sf, person_name, config, all_employees)
        emp_results.append(result)

    section['employees'] = emp_results
    section['status'] = 'error' if any(r['status'] == 'error' for r in emp_results) else \
                        'warning' if any(r['status'] == 'warning' for r in emp_results) else \
                        'ok' if emp_results else 'info'

    for r in emp_results:
        for issue in r['issues']:
            section['issues'].append(_issue(f"{r['name']}: {issue}", 'Demissões'))

    return section


def _audit_ferias_scope(scope_path, scope_name, recurse=True, file_paths=None):
    result = {'name': scope_name, 'items': [], 'issues': []}
    tags = get_tags_for_files(file_paths) if file_paths is not None else get_all_tags_in_folder(scope_path, recurse=recurse)

    checks = [
        ('AVISO_FERIAS', 'Aviso de férias', 'error', 'Aviso de férias não encontrado'),
        ('RECIBO_FERIAS', 'Recibo de férias', 'error', 'Recibo de férias não encontrado'),
        ('COMPROVANTE_FERIAS', 'Comprovante de pagamento', 'warning',
         'Comprovante de pagamento de férias não identificado pelo nome — verificar manualmente'),
    ]

    for tag, label, missing_status, missing_msg in checks:
        if tag in tags:
            result['items'].append(_item(label, 'ok'))
        else:
            note = 'FALTANDO' if missing_status == 'error' else 'Não identificado'
            result['items'].append(_item(label, missing_status, note))
            result['issues'].append(missing_msg)

    result['status'] = 'error' if any(i['status'] == 'error' for i in result['items']) else \
                       'warning' if any(i['status'] == 'warning' for i in result['items']) else 'ok'
    return result


def _ferias_subjects(fer_folder, employees):
    subfolders = list_subfolders(fer_folder)
    if subfolders:
        return [(os.path.basename(sf), sf, True, None) for sf in subfolders]

    files = [f for f in os.listdir(fer_folder) if os.path.isfile(os.path.join(fer_folder, f))]
    subjects = []
    for emp in employees or []:
        matched = [
            os.path.join(fer_folder, f)
            for f in files
            if employee_matches_folder(emp.get('nome', ''), f)
        ]
        if matched:
            subjects.append((emp['nome'], fer_folder, False, matched))

    if subjects:
        return subjects
    return [('Competência inteira', fer_folder, True, None)]


def audit_ferias(folder_path, employees=None):
    section = {'name': 'Férias', 'icon': '🏖️', 'items': [], 'employees': [], 'issues': []}
    fer_folder = find_folder(folder_path, FOLDER_PATTERNS['ferias'])

    if not fer_folder:
        section['status'] = 'info'
        section['items'].append(_item('Pasta Férias não encontrada', 'info', 'Sem férias no mês'))
        return section

    subject_results = [
        _audit_ferias_scope(subject_path, subject_name, recurse=recurse, file_paths=file_paths)
        for subject_name, subject_path, recurse, file_paths in _ferias_subjects(fer_folder, employees or [])
    ]

    if len(subject_results) == 1 and subject_results[0]['name'] == 'Competência inteira':
        scope = subject_results[0]
        section['items'] = scope['items']
        section['status'] = scope['status']
        for issue in scope['issues']:
            section['issues'].append(_issue(_join_scope(scope['name'], issue), 'Férias'))
    else:
        section['employees'] = subject_results
        section['status'] = 'error' if any(r['status'] == 'error' for r in subject_results) else \
                            'warning' if any(r['status'] == 'warning' for r in subject_results) else 'ok'
        for r in subject_results:
            for issue in r['issues']:
                section['issues'].append(_issue(f"{r['name']}: {issue}", 'Férias'))
    return section


def audit_fopag(folder_path):
    section = {'name': 'FOPAG e Comp. de Pgto', 'icon': '💰', 'items': [], 'issues': []}
    fopag_folder = find_folder(folder_path, FOLDER_PATTERNS['fopag'])

    if not fopag_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "FOPAG e Comp. de Pgto" não encontrada', 'FOPAG'))
        section['items'].append(_item('Pasta FOPAG', 'error', 'Não encontrada'))
        return section

    tags = get_all_tags_in_folder(fopag_folder)

    if 'FOPAG' in tags:
        section['items'].append(_item('Folha de pagamento analítica (FOPAG)', 'ok'))
    else:
        section['items'].append(_item('Folha de pagamento analítica (FOPAG)', 'error', 'FALTANDO'))
        section['issues'].append(_issue('Folha de pagamento analítica não encontrada', 'FOPAG'))

    has_adto = 'ADIANTAMENTO' in tags
    has_saldo = 'SALDO_SALARIO' in tags
    has_comp = 'COMPROVANTE_SALARIO' in tags

    if has_adto:
        section['items'].append(_item('Comprovante de adiantamento', 'ok'))
    else:
        section['items'].append(_item('Comprovante de adiantamento', 'info', 'Não identificado (pode não haver adiantamento)'))

    if has_saldo or has_comp:
        section['items'].append(_item('Comprovante de pagamento (saldo/salário)', 'ok'))
    else:
        section['items'].append(_item('Comprovante de pagamento (saldo/salário)', 'error', 'FALTANDO'))
        section['issues'].append(_issue('Comprovante de pagamento de salário não encontrado', 'FOPAG'))

    section['status'] = 'error' if section['issues'] else \
                        'warning' if any(i['status'] == 'warning' for i in section['items']) else 'ok'
    return section


def audit_inss_fgts(folder_path, config):
    section = {'name': 'INSS + FGTS', 'icon': '🏦', 'items': [], 'issues': []}
    inss_folder = find_folder(folder_path, FOLDER_PATTERNS['inss_fgts'])

    if not inss_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "INSS + FGTS" não encontrada', 'INSS + FGTS'))
        section['items'].append(_item('Pasta INSS + FGTS', 'error', 'Não encontrada'))
        return section

    tags = get_all_tags_in_folder(inss_folder)
    dctfweb_req = config.get('dctfweb_required', False)
    darf_req = config.get('darf_inss_required', False)

    # Required FGTS docs
    fgts_checks = [
        ('GUIA_FGTS', 'Guia FGTS Digital', True),
        ('COMPROVANTE_FGTS', 'Comprovante de pagamento FGTS', True),
        ('DETALHAMENTO_FGTS', 'Detalhamento da guia FGTS', True),
        ('CRF', 'Certificado de Regularidade FGTS (CRF)', True),
    ]

    for tag, label, req in fgts_checks:
        if tag in tags:
            section['items'].append(_item(label, 'ok'))
        elif req:
            section['items'].append(_item(label, 'error', 'FALTANDO'))
            section['issues'].append(_issue(f'Competência/empresa: {label} não encontrado', 'INSS + FGTS'))

    # DCTFWeb - per config
    if dctfweb_req:
        has_dctf = any(t in tags for t in ['DCTFWEB', 'RECIBO_DCTFWEB', 'RESUMO_DEBITOS', 'RESUMO_CREDITOS'])
        if has_dctf:
            section['items'].append(_item('DCTFWeb (declaração completa)', 'ok'))
        else:
            section['items'].append(_item('DCTFWeb', 'error', 'FALTANDO'))
            section['issues'].append(_issue('Competência/empresa: DCTFWeb não encontrado', 'INSS + FGTS'))
    else:
        has_dctf = any(t in tags for t in ['DCTFWEB', 'RECIBO_DCTFWEB'])
        if has_dctf:
            section['items'].append(_item('DCTFWeb', 'ok', '(presente — bônus)'))
        else:
            section['items'].append(_item('DCTFWeb', 'info', 'Revisão futura — não exigido neste envio'))

    # DARF INSS
    if darf_req:
        if 'DARF_INSS' in tags:
            section['items'].append(_item('DARF INSS', 'ok'))
        else:
            section['items'].append(_item('DARF INSS', 'error', 'FALTANDO'))
            section['issues'].append(_issue('Competência/empresa: DARF INSS não encontrado', 'INSS + FGTS'))
    else:
        if 'DARF_INSS' in tags:
            section['items'].append(_item('DARF INSS', 'ok', '(presente — bônus)'))
        else:
            section['items'].append(_item('DARF INSS', 'info', 'Revisão futura — não exigido neste envio'))

    section['status'] = 'error' if section['issues'] else 'ok'
    return section


def audit_ponto(folder_path, active_employees, config):
    section = {'name': 'Ponto', 'icon': '⏰', 'items': [], 'issues': []}
    ponto_folder = find_folder(folder_path, FOLDER_PATTERNS['ponto'])
    ponto_exempt = [normalize(e) for e in config.get('ponto_exempt', [])]

    if ponto_folder:
        files = [f for f in os.listdir(ponto_folder) if os.path.isfile(os.path.join(ponto_folder, f))]
        if files:
            for f in files:
                section['items'].append(_item(f, 'ok'))
            section['status'] = 'ok'
        else:
            section['items'].append(_item('Folha de ponto', 'error', 'Pasta vazia'))
            section['issues'].append(_issue('Pasta Ponto está vazia', 'Ponto'))
            section['status'] = 'error'
    else:
        section['items'].append(_item('Pasta Ponto', 'error', 'Não encontrada'))
        section['issues'].append(_issue('Pasta "Ponto" não encontrada', 'Ponto'))
        section['status'] = 'error'

    if ponto_exempt:
        exempt_names = config.get('ponto_exempt', [])
        section['items'].append(_item(
            f'Isentos de ponto: {", ".join(exempt_names)}',
            'info',
            'Por regra operacional'
        ))

    return section


def audit_va_vr(folder_path):
    section = {'name': 'VA e VR', 'icon': '🎫', 'items': [], 'issues': []}
    va_folder = find_folder(folder_path, FOLDER_PATTERNS['va_vr'])

    if not va_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "VA e VR" não encontrada', 'VA e VR'))
        section['items'].append(_item('Pasta VA e VR', 'error', 'Não encontrada'))
        return section

    tags = get_all_tags_in_folder(va_folder)
    files = [f for f in os.listdir(va_folder) if os.path.isfile(os.path.join(va_folder, f))]

    has_relatorio = 'RELATORIO_VAVR' in tags
    has_boleto = 'BOLETO_VAVR' in tags
    has_nf = 'NF_VAVR' in tags
    has_comp = 'COMPROVANTE_VAVR' in tags

    if has_relatorio:
        section['items'].append(_item('Relatório de colaboradores', 'ok'))
    else:
        section['items'].append(_item('Relatório de colaboradores', 'error', 'FALTANDO'))
        section['issues'].append(_issue('Relatório de colaboradores VA/VR não encontrado', 'VA e VR'))

    if has_boleto or has_comp:
        section['items'].append(_item('Comprovante/boleto de pagamento', 'ok'))
    else:
        section['items'].append(_item('Comprovante/boleto de pagamento', 'warning', 'Não identificado'))

    if has_nf:
        section['items'].append(_item('Nota Fiscal', 'ok'))
    else:
        section['items'].append(_item('Nota Fiscal', 'warning', 'Não identificada'))

    section['status'] = 'error' if section['issues'] else \
                        'warning' if any(i['status'] == 'warning' for i in section['items']) else 'ok'
    return section


def audit_vt(folder_path):
    section = {'name': 'VT', 'icon': '🚌', 'items': [], 'issues': []}
    vt_folder = find_folder(folder_path, FOLDER_PATTERNS['vt'])

    if not vt_folder:
        section['status'] = 'error'
        section['issues'].append(_issue('Pasta "VT" não encontrada', 'VT'))
        section['items'].append(_item('Pasta VT', 'error', 'Não encontrada'))
        return section

    tags = get_all_tags_in_folder(vt_folder)
    files = [f for f in os.listdir(vt_folder) if os.path.isfile(os.path.join(vt_folder, f))]

    has_rel = any(t in tags for t in ['RELATORIO_VT', 'NF_VT'])
    has_bol = 'BOLETO_VT' in tags
    has_comp = 'COMPROVANTE_VT' in tags

    if has_rel:
        section['items'].append(_item('Relatório/Pedido de VT', 'ok'))
    else:
        section['items'].append(_item('Relatório/Pedido de VT', 'warning', 'Não identificado'))

    if has_bol or has_comp:
        section['items'].append(_item('Comprovante/boleto de pagamento', 'ok'))
    else:
        section['items'].append(_item('Comprovante/boleto de pagamento', 'warning', 'Não identificado'))

    # If there are any files at all, it's probably ok
    if files and not (has_rel or has_bol or has_comp):
        section['items'] = [_item(f, 'ok') for f in files]
        section['status'] = 'ok'
    else:
        section['status'] = 'warning' if any(i['status'] == 'warning' for i in section['items']) else 'ok'
    return section


def find_previous_month_folder(folder_path):
    parent = os.path.dirname(folder_path)
    month_num, year, _ = detect_competencia(folder_path)
    if not month_num:
        return None
    prev_month = month_num - 1 if month_num > 1 else 12
    prev_year = year if month_num > 1 else year - 1
    prev_name = normalize(MONTH_NAMES[prev_month])
    try:
        for item in os.listdir(parent):
            if os.path.isdir(os.path.join(parent, item)) and prev_name in normalize(item):
                return os.path.join(parent, item)
    except Exception:
        pass
    return None


def audit_prev_month_comparison(folder_path, current_employees, month_num, year):
    section = {'name': 'Comparação Mês Anterior', 'icon': '🔄', 'items': [], 'issues': []}

    prev_folder = find_previous_month_folder(folder_path)
    if not prev_folder:
        section['status'] = 'info'
        section['items'].append(_item('Mês anterior não encontrado', 'info', 'Comparação indisponível'))
        return section

    prev_ft = find_forca_trabalho(prev_folder)
    if not prev_ft:
        section['status'] = 'info'
        section['items'].append(_item(f'Planilha ausente em {os.path.basename(prev_folder)}', 'info', 'Comparação indisponível'))
        return section

    prev_employees, err = read_forca_trabalho(prev_ft)
    if err:
        section['status'] = 'warning'
        section['items'].append(_item('Erro ao ler planilha do mês anterior', 'warning', err))
        return section

    prev_month = month_num - 1 if month_num > 1 else 12
    prev_year = year if month_num > 1 else year - 1
    prev_month_emps = get_month_employees(prev_employees, prev_year, prev_month)

    prev_active = {normalize(e['nome']): e for e in prev_month_emps if e['situacao'] == 'A'}
    curr_active = {normalize(e['nome']): e for e in current_employees if e['situacao'] == 'A'}
    curr_all = {normalize(e['nome']): e for e in current_employees}

    section['items'].append(_item(
        f'Referência: {os.path.basename(prev_folder)}',
        'info',
        f'{len(prev_active)} ativos → {len(curr_active)} ativos no mês atual'
    ))

    # New in current that weren't in previous
    new_names = set(curr_active.keys()) - set(prev_active.keys())
    for name in sorted(new_names):
        emp = curr_active[name]
        admissao_folder = find_folder(folder_path, FOLDER_PATTERNS['admissao'])
        found_adm = False
        if admissao_folder:
            found_adm = any(
                employee_matches_folder(emp['nome'], os.path.basename(sf), loose=True)
                for sf in list_subfolders(admissao_folder)
            )
        if found_adm:
            section['items'].append(_item(f'Novo colaborador: {emp["nome"]}', 'ok', 'Pasta de admissão presente'))
        else:
            section['items'].append(_item(f'Novo colaborador: {emp["nome"]}', 'warning', 'Sem pasta de admissão — verificar'))
            section['issues'].append(_issue(f'Novo colaborador detectado sem pasta de admissão: {emp["nome"]}', 'Comparação'))

    # Active in previous but absent or inactive now
    gone_names = set(prev_active.keys()) - set(curr_active.keys())
    for name in sorted(gone_names):
        emp_prev = prev_active[name]
        emp_curr = curr_all.get(name)

        if emp_curr and emp_curr['situacao'] == 'I':
            # Confirmed demission — check if demissão folder exists
            dem_folder = find_folder(folder_path, FOLDER_PATTERNS['demissao'])
            found_dem = False
            if dem_folder:
                found_dem = any(
                    employee_matches_folder(emp_curr['nome'], os.path.basename(sf), loose=True)
                    for sf in list_subfolders(dem_folder)
                )
            if found_dem:
                section['items'].append(_item(f'Desligado: {emp_curr["nome"]}', 'ok', 'Pasta de demissão presente'))
            else:
                section['items'].append(_item(f'Desligado: {emp_curr["nome"]}', 'error', 'Sem pasta de demissão — VERIFICAR'))
                section['issues'].append(_issue(f'Colaborador desligado sem pasta de demissão: {emp_curr["nome"]}', 'Comparação'))
        else:
            # Disappeared without being marked inactive — flag for review
            section['items'].append(_item(f'Ausente: {emp_prev["nome"]}', 'warning',
                                          'Ativo no mês anterior — não consta no atual'))
            section['issues'].append(_issue(
                f'{emp_prev["nome"]} estava ativo no mês anterior mas não aparece no atual — verificar planilha',
                'Comparação'
            ))

    section['status'] = 'error' if any(i['status'] == 'error' for i in section['items']) else \
                       'warning' if section['issues'] else 'ok'
    return section


def _pdf_diagnostic_section(diagnostics):
    section = {'name': 'Diagnóstico da Auditoria PDF', 'icon': '🧪', 'items': [], 'issues': [], 'status': 'info'}
    pdf_total = int(diagnostics.get('pdf_files', 0) or 0)
    analyzed_now = int(diagnostics.get('analyzed_files', 0) or 0)
    cache_hits = int(diagnostics.get('cache_hits', 0) or 0)
    considered = max(pdf_total, analyzed_now + cache_hits)
    section['items'].append(_item('Modo de analise', 'info', diagnostics.get('mode', 'rapida sem OCR')))
    section['items'].append(_item('Cache da medicao', 'info', 'ignorado nesta execucao' if diagnostics.get('force_reprocess') else 'reutilizado quando valido'))
    section['items'].append(_item('Escopo do OCR', 'info', diagnostics.get('ocr_scope', '-')))
    section['items'].append(_item('PDFs considerados na auditoria', 'info', str(considered)))
    section['items'].append(_item('PDFs reprocessados nesta execucao', 'info', str(analyzed_now)))
    section['items'].append(_item('PDFs reutilizados do cache', 'info', str(cache_hits)))
    section['items'].append(_item('Paginas reprocessadas nesta execucao', 'info', str(diagnostics.get('pages_processed', 0))))
    section['items'].append(_item('Paginas com texto digital nesta execucao', 'info', str(diagnostics.get('digital_pages', 0))))
    section['items'].append(_item('Paginas com OCR nesta execucao', 'info', str(diagnostics.get('ocr_pages', 0))))

    docs_unicos = []
    vistos = set()
    for doc in diagnostics.get('documents_found', []):
        pages = doc.get('pages') or []
        chave = (
            doc.get('documentType') or doc.get('documentName') or '',
            doc.get('fileName') or '',
            tuple(pages),
            doc.get('confidence') or '',
            doc.get('method') or '',
        )
        if chave in vistos:
            continue
        vistos.add(chave)
        docs_unicos.append(doc)

    for doc in docs_unicos[:30]:
        pages = doc.get('pages') or []
        page_txt = f"páginas {min(pages)}-{max(pages)}" if pages else "páginas ?"
        confidence = doc.get('confidence') or 'media'
        label = f"{doc.get('documentName') or doc.get('documentType')} em {doc.get('fileName')}"
        note = f"{page_txt}; confiança {confidence}; método {doc.get('method') or 'conteúdo'}"
        if confidence != 'alta':
            note = f"{note}; achado tecnico para conferencia, nao e pendencia"
        status = 'ok' if confidence == 'alta' else 'info'
        section['items'].append(_item(label, status, note))

    for err in diagnostics.get('errors', [])[:10]:
        section['items'].append(_item('Aviso técnico', 'warning', err))

    if any(i['status'] == 'warning' for i in section['items']):
        section['status'] = 'warning'
    return section


def run_audit(folder_path, enable_ocr=False, analyze_pdf_content=False, progress_cb=None, force_reprocess=False):
    configure_pdf_audit(
        enable_ocr=enable_ocr,
        analyze_content=analyze_pdf_content or enable_ocr,
        progress_cb=progress_cb,
        force_reprocess=force_reprocess,
    )
    config = load_config()
    result = {
        'folder_path': folder_path,
        'timestamp': __import__('datetime').datetime.now().strftime('%d/%m/%Y %H:%M'),
        'sections': [],
        'all_issues': [],
        'overall_status': 'ok',
    }

    # Detect competência
    month_num, year, month_name = detect_competencia(folder_path)
    result['competencia'] = f'{month_name}/{year}' if year else month_name

    # --- Força de Trabalho ---
    ft_path = find_forca_trabalho(folder_path)
    ft_section = {'name': 'Força de Trabalho', 'icon': '📊', 'items': [], 'issues': []}
    all_employees = []
    admitted_this_month = []
    dismissed_this_month = []

    if ft_path:
        employees, err = read_forca_trabalho(ft_path)
        if err:
            ft_section['status'] = 'warning'
            ft_section['items'].append(_item('Erro ao ler planilha', 'warning', err))
        else:
            month_emps = get_month_employees(employees, year, month_num) if month_num else employees
            all_employees = month_emps

            active = [e for e in month_emps if e['situacao'] == 'A']
            inactive = [e for e in month_emps if e['situacao'] == 'I']

            if month_num:
                admitted_this_month = [
                    e for e in month_emps
                    if e.get('data_admissao') and
                    e['data_admissao'].month == month_num and
                    e['data_admissao'].year == year
                ]
                dismissed_this_month = [
                    e for e in month_emps
                    if e.get('data_demissao') and
                    e['data_demissao'].month == month_num and
                    e['data_demissao'].year == year
                ]

            ft_section['status'] = 'ok'
            ft_section['items'].append(_item(f'Planilha: {os.path.basename(ft_path)}', 'ok'))
            ft_section['items'].append(_item(f'Colaboradores ativos: {len(active)}', 'info',
                                             ', '.join(e['nome'] for e in active[:5]) + ('...' if len(active) > 5 else '')))
            if admitted_this_month:
                ft_section['items'].append(_item(f'Admissões no mês: {len(admitted_this_month)}', 'info',
                                                 ', '.join(e['nome'] for e in admitted_this_month)))
            if dismissed_this_month:
                ft_section['items'].append(_item(f'Demissões no mês: {len(dismissed_this_month)}', 'info',
                                                 ', '.join(e['nome'] for e in dismissed_this_month)))
    else:
        ft_section['status'] = 'error'
        ft_section['items'].append(_item('Planilha Força de Trabalho', 'error', 'FALTANDO na pasta'))
        ft_section['issues'].append(_issue('Planilha Força de Trabalho não encontrada', 'Força de Trabalho'))

    result['sections'].append(ft_section)

    # Detect what's present for conditional checks
    admissao_folder = find_folder(folder_path, FOLDER_PATTERNS['admissao'])
    demissao_folder = find_folder(folder_path, FOLDER_PATTERNS['demissao'])
    ferias_folder = find_folder(folder_path, FOLDER_PATTERNS['ferias'])

    has_admissao = admissao_folder is not None and bool(list_subfolders(admissao_folder))
    has_demissao = demissao_folder is not None and bool(list_subfolders(demissao_folder))
    has_ferias = ferias_folder is not None and has_any_file(ferias_folder)

    # Detect troca de função
    has_troca_funcao = False
    if admissao_folder:
        for sf in list_subfolders(admissao_folder):
            sf_norm = normalize(os.path.basename(sf))
            if 'troca de funcao' in sf_norm or 'mudanca de funcao' in sf_norm:
                has_troca_funcao = True
                break

    # --- Run all section audits ---
    if progress_cb:
        progress_cb('Lendo estrutura da competência')
    result['sections'].append(audit_acordo_coletivo(folder_path))
    result['sections'].append(audit_declaracoes(folder_path, has_admissao, has_demissao,
                                                 has_ferias, has_troca_funcao, config))

    if progress_cb:
        progress_cb('Auditando admissões')
    admissao_result = audit_admissoes(folder_path, admitted_this_month, config)
    admissao_section = admissao_result[0] if isinstance(admissao_result, tuple) else admissao_result
    result['sections'].append(admissao_section)

    if progress_cb:
        progress_cb('Auditando demissões, férias e documentos mensais')
    result['sections'].append(audit_demissoes(folder_path, dismissed_this_month, config, all_employees))
    result['sections'].append(audit_ferias(folder_path, all_employees))
    result['sections'].append(audit_fopag(folder_path))
    result['sections'].append(audit_inss_fgts(folder_path, config))
    result['sections'].append(audit_ponto(folder_path, all_employees, config))
    result['sections'].append(audit_va_vr(folder_path))
    result['sections'].append(audit_vt(folder_path))

    # Previous month comparison (only when spreadsheet was found)
    if all_employees and month_num:
        result['sections'].append(
            audit_prev_month_comparison(folder_path, all_employees, month_num, year)
        )

    diagnostics = get_pdf_audit_diagnostics()
    result['diagnostics'] = diagnostics
    result['sections'].append(_pdf_diagnostic_section(diagnostics))
    _resolve_pendencias_por_busca_profunda(result['sections'], folder_path)

    # Aggregate all issues (employees only — section-level issues already cover them)
    for sec in result['sections']:
        for issue in sec.get('issues', []):
            # Only add section-level issues that are NOT about specific employees
            if 'employees' not in sec or not sec['employees']:
                result['all_issues'].append(issue)
        for emp in sec.get('employees', []):
            for issue in emp.get('issues', []):
                result['all_issues'].append({'msg': f"{sec['name']} — {emp['name']}: {issue}", 'section': sec['name']})

    # Overall status: avisos tecnicos sem pendencia real nao devem gerar ATENCAO.
    statuses = [sec.get('status', 'ok') for sec in result['sections']]
    if not result['all_issues']:
        result['overall_status'] = 'ok'
    elif 'error' in statuses:
        result['overall_status'] = 'error'
    else:
        result['overall_status'] = 'warning'

    return result
