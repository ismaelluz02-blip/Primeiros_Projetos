import os
import tempfile
import hashlib
import re
import unicodedata
import html

STATUS_ICON = {'ok': '✅', 'error': '❌', 'warning': '⚠️', 'info': 'ℹ️'}
STATUS_COLOR = {'ok': '#27ae60', 'error': '#e74c3c', 'warning': '#f39c12', 'info': '#3498db'}
STATUS_BG = {'ok': '#eafaf1', 'error': '#fdedec', 'warning': '#fef9e7', 'info': '#eaf4fb'}


def _iid(folder, text):
    """Stable short ID for an item — used as localStorage sub-key."""
    return hashlib.md5(f"{folder}|{text}".encode('utf-8')).hexdigest()[:10]


def _norm(text):
    text = unicodedata.normalize('NFKD', str(text or ''))
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r'[^a-z0-9]+', ' ', text.lower())
    return ' '.join(text.split())


def _doc_code(text):
    n = _norm(text)
    for code, terms in [
        ('AVISO_PREVIO', ['aviso previo', 'pedido de demissao']),
        ('SEGURO_DESEMPREGO', ['seguro desemprego']),
        ('RECIBO_FERIAS', ['recibo de ferias']),
        ('AVISO_FERIAS', ['aviso de ferias']),
        ('COMPROVANTE_RESCISAO', ['comprovante de pagamento da rescisao']),
        ('COMPROVANTE_FERIAS', ['comprovante de pagamento de ferias', 'comprovante de pagamento']),
        ('TRCT', ['termo de rescisao', 'trct']),
        ('FGTS_GRRF', ['fgts grrf', 'grrf', 'multa 40']),
        ('ASO', ['aso demissional', 'aso admissional', 'aso']),
        ('ESOCIAL', ['esocial', 'baixa digital']),
        ('VT_DECLARACAO', ['declaracao de opcao vt', 'vale transporte']),
        ('EPI', ['ficha de epi', 'entrega de epi']),
        ('CTPS_ESOCIAL', ['ctps', 'ctps digital']),
        ('CONTRATO', ['contrato de trabalho']),
        ('FICHA_REGISTRO', ['ficha de registro']),
    ]:
        if any(term in n for term in terms):
            return code
    return n[:80] or 'pendencia'


def _issue_storage_id(folder, section='', person='', text=''):
    return _iid(folder, f"{_norm(section)}|{_norm(person)}|{_doc_code(text)}")


def _parse_issue(msg):
    section = ''
    person = ''
    text = msg or ''
    if ' — ' in text:
        section, text = text.split(' — ', 1)
    if ': ' in text:
        person, text = text.split(': ', 1)
    return section, person, text


def _h(text):
    return html.escape(str(text or ''), quote=True)


def _dedupe_issues(issues, folder):
    deduped = []
    seen = {}
    for issue in issues or []:
        issue_dict = issue if isinstance(issue, dict) else {'msg': str(issue)}
        msg = issue_dict.get('msg', '')
        section, person, text = _parse_issue(msg)
        key = _issue_storage_id(folder, section or issue_dict.get('section', ''), person, text)
        if key in seen:
            seen[key]['duplicate_count'] = seen[key].get('duplicate_count', 1) + 1
            continue
        item = dict(issue_dict)
        item['duplicate_count'] = int(item.get('duplicate_count', 1) or 1)
        item['_dedupe_key'] = key
        deduped.append(item)
        seen[key] = item
    return deduped


def _pending_explanation(text, section='', person='', note=''):
    raw = ' '.join(str(v or '') for v in (section, person, text, note))
    n = _norm(raw)
    subject = str(text or '').strip() or str(note or '').strip() or 'Item sinalizado'
    scope = []
    if section:
        scope.append(f"seção {section}")
    if person:
        scope.append(f"colaborador(a) {person}")
    scope_txt = ' / '.join(scope) if scope else 'escopo desta conferência'

    if 'pasta' in n and ('nao encontrada' in n or 'sem pasta' in n):
        reason = (
            f"O auditor esperava encontrar uma pasta obrigatória para {scope_txt}, mas a pasta não foi localizada "
            "com os nomes/padrões aceitos pela medição."
        )
        expected = (
            "A pasta deveria existir dentro da competência, com nome compatível com o tipo de documento exigido "
            "e, quando for por colaborador, contendo uma subpasta ou arquivos identificáveis para a pessoa correta."
        )
        check = (
            "Confira se a pasta foi criada, se não ficou dentro de outra pasta por engano, se o nome não está muito diferente "
            "do padrão e se os arquivos não foram enviados soltos em outro local."
        )
    elif 'vazia' in n or 'nenhum arquivo' in n:
        reason = (
            f"A estrutura foi encontrada em {scope_txt}, mas não havia arquivo suficiente para comprovar o requisito."
        )
        expected = (
            "Deveria haver pelo menos um PDF/arquivo válido dentro da pasta correta, com conteúdo legível e relacionado ao documento solicitado."
        )
        check = (
            "Verifique se o arquivo foi salvo na competência certa, se não está em uma subpasta inesperada, se não ficou com extensão diferente "
            "e se o PDF abre normalmente."
        )
    elif 'verificar se se aplica' in n or 'pode nao haver' in n or 'nao identificado' in n:
        reason = (
            "O auditor não encontrou evidência suficiente para confirmar automaticamente se este item se aplica ou se pode ser dispensado."
        )
        expected = (
            "Se o documento for obrigatório para este caso, ele deveria estar na pasta correspondente e identificado por nome ou conteúdo. "
            "Se não for obrigatório, precisa ficar justificado para auditoria."
        )
        check = (
            "Confira a situação do colaborador/competência, o tipo de contrato e a regra operacional. Se realmente não se aplicar, registre a justificativa."
        )
    elif 'nao encontrado' in n or 'faltando' in n:
        reason = (
            f"O auditor procurou por '{subject}' em {scope_txt}, mas não encontrou evidência confiável no nome do arquivo nem no conteúdo analisado."
        )
        expected = (
            "O documento deveria estar na pasta correta da competência, preferencialmente no grupo certo e, quando envolver colaborador, "
            "na pasta da pessoa correspondente. O arquivo também precisa ter texto digital ou OCR suficiente para ser reconhecido."
        )
        check = (
            "Verifique se o documento existe, se está na pasta errada, se está com nome genérico, se foi anexado no colaborador errado, "
            "se o PDF está escaneado com baixa qualidade ou se o conteúdo pertence a outro documento."
        )
    elif 'confianca media' in n or 'confiança media' in n:
        reason = (
            "O documento foi encontrado, mas a classificação ficou com confiança média. Isso significa que o sistema viu algumas evidências, "
            "mas não encontrou sinais fortes o bastante para considerar o item totalmente confirmado."
        )
        expected = (
            "O arquivo deveria conter termos claros do documento esperado, páginas coerentes e conteúdo legível para que a confiança fique alta."
        )
        check = (
            "Abra o PDF e confira se as páginas indicadas realmente correspondem ao documento. Se estiver correto, pode justificar; se não, reorganize ou renomeie."
        )
    else:
        reason = (
            f"O item foi marcado porque a regra de conferência da seção {section or 'atual'} não encontrou uma comprovação completa."
        )
        expected = (
            "O esperado é que a pasta, o arquivo, o nome e o conteúdo batam com o requisito da medição para a competência auditada."
        )
        check = (
            "Confira localização, nome do arquivo, legibilidade do PDF, páginas apontadas, colaborador vinculado e se o documento pertence ao mês correto."
        )

    return reason, expected, check


def _detail_box(text, section='', person='', note='', duplicate_count=1):
    reason, expected, check = _pending_explanation(text, section=section, person=person, note=note)
    dup = ''
    if duplicate_count and duplicate_count > 1:
        dup = f'<div style="margin-top:6px;color:#8a5a00"><strong>Ocorrências agrupadas:</strong> este mesmo problema apareceu {duplicate_count} vezes e foi consolidado aqui.</div>'
    path_html = ''
    match = re.search(r'caminho:\s*(.+?)(?:\s*\||$)', str(note or ''))
    if match:
        path = match.group(1).strip()
        path_html = f'''
      <div style="margin-top:6px"><strong>Endereço do arquivo:</strong></div>
      <div style="margin-top:3px;background:#fff;border:1px dashed #c9a94a;border-radius:4px;padding:6px 8px;font-family:Consolas,monospace;font-size:0.88em;color:#3d3d3d;user-select:text;word-break:break-all">{_h(path)}</div>'''
    return f'''
    <div class="why-box" style="margin-top:8px;background:#fffdf5;border:1px solid #f1d18a;border-left:4px solid #f39c12;border-radius:6px;padding:9px 11px;color:#5b4a1f;font-size:0.9em;line-height:1.45">
      <div><strong>Por que isso é pendência:</strong> {_h(reason)}</div>
      <div style="margin-top:5px"><strong>Como deveria estar:</strong> {_h(expected)}</div>
      <div style="margin-top:5px"><strong>O que conferir agora:</strong> {_h(check)}</div>
      {path_html}
      {dup}
    </div>'''


def _render_item(item, folder, justifiable=False, context=None):
    st = item.get('status', 'info')
    icon = STATUS_ICON.get(st, '•')
    color = STATUS_COLOR.get(st, '#555')
    label = item.get('label', '')
    note = item.get('note', '')
    note_html = f' <small style="color:#888">— {note}</small>' if note else ''

    if justifiable and st in ('error', 'warning'):
        context = context or {}
        storage_id = _issue_storage_id(folder, context.get('section', ''), context.get('person', ''), f'{label} {note}')
        iid = _iid(folder, f'{storage_id}|detail|{label}|{note}')
        detail_html = _detail_box(label, section=context.get('section', ''), person=context.get('person', ''), note=note)
        return f'''
<div id="row_{iid}" data-rowid="{iid}" data-justifiable="{storage_id}" style="display:flex;align-items:flex-start;gap:8px;margin:5px 0">
  <div style="flex:1">
    <span style="color:{color}">{icon} {label}{note_html}</span>
    <span id="badge_{iid}" style="display:none;margin-left:6px"></span>
    {detail_html}
    <div id="panel_{iid}" style="display:none;margin-top:7px;background:#fff;border:1px solid #dce3ea;border-radius:6px;padding:10px 12px">
      <div id="display_{iid}" style="display:none;color:#555;font-size:0.88em;margin-bottom:6px;font-style:italic;border-left:3px solid #3498db;padding-left:8px"></div>
      <textarea id="note_{iid}" rows="2" placeholder="Descreva a justificativa ou tratativa…"
                style="width:100%;border:1px solid #ccc;border-radius:4px;padding:6px 8px;font-size:0.9em;resize:vertical;font-family:inherit"></textarea>
      <div style="margin-top:7px;display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <label style="display:flex;align-items:center;gap:5px;font-size:0.9em;cursor:pointer">
          <input type="checkbox" id="chk_{iid}" style="width:15px;height:15px"> ✅ Marcar como tratado
        </label>
        <small id="date_{iid}" style="color:#aaa"></small>
        <button onclick="saveJust('{iid}')"
                style="margin-left:auto;background:#2980b9;color:white;border:none;padding:5px 14px;border-radius:4px;cursor:pointer;font-size:0.88em">
          Salvar
        </button>
      </div>
    </div>
  </div>
  <button onclick="togglePanel('{iid}')" title="Justificar / Tratar"
          style="background:#f39c12;color:white;border:none;padding:4px 9px;border-radius:4px;cursor:pointer;font-size:0.82em;white-space:nowrap;flex-shrink:0;margin-top:2px">
    📝 Justificar
  </button>
</div>'''
    else:
        return f'<li style="color:{color};margin:4px 0">{icon} {label}{note_html}</li>\n'


def _render_items(items, folder='', justifiable=False, context=None):
    if not items:
        return ''
    has_just = justifiable and any(i.get('status') in ('error', 'warning') for i in items)
    if has_just:
        rows = ''.join(_render_item(i, folder, justifiable=True, context=context) for i in items)
        return f'<div style="padding-left:8px;margin:6px 0">{rows}</div>'
    rows = ''.join(_render_item(i, folder, context=context) for i in items)
    return f'<ul style="list-style:none;padding-left:8px;margin:6px 0">{rows}</ul>'


def _render_employee(emp, folder, section_name=''):
    st = emp.get('status', 'ok')
    icon = STATUS_ICON.get(st, '•')
    color = STATUS_COLOR.get(st, '#333')
    bg = STATUS_BG.get(st, '#fff')
    name = emp.get('name', '')
    items_html = _render_items(
        emp.get('items', []),
        folder=folder,
        justifiable=True,
        context={'section': section_name, 'person': name},
    )
    return f'''
    <div class="emp-block" style="border-left:4px solid {color};background:{bg};padding:8px 12px;margin:8px 0;border-radius:4px">
      <strong style="color:{color}">{icon} {name}</strong>
      {items_html}
    </div>'''


def _render_section(sec, folder):
    st = sec.get('status', 'ok')
    icon_sec = sec.get('icon', '📁')
    name = sec.get('name', '')
    color = STATUS_COLOR.get(st, '#333')
    bg = STATUS_BG.get(st, '#fff')
    status_icon = STATUS_ICON.get(st, '•')

    items_html = _render_items(
        sec.get('items', []),
        folder=folder,
        justifiable=True,
        context={'section': name},
    )
    employees_html = ''.join(_render_employee(e, folder, section_name=name) for e in sec.get('employees', []))

    return f'''
  <div class="section" style="border:1px solid #ddd;border-radius:8px;margin:16px 0;overflow:hidden">
    <div class="sec-header" style="background:{color};color:white;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;cursor:pointer"
         onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'block':'none'">
      <span style="font-size:1.05em;font-weight:600">{icon_sec} {name}</span>
      <span style="font-size:1.1em">{status_icon}</span>
    </div>
    <div class="sec-body" style="padding:12px 16px;background:{bg}">
      {items_html}
      {employees_html}
    </div>
  </div>'''


def _render_issues_section(all_issues, folder):
    all_issues = _dedupe_issues(all_issues, folder)
    if not all_issues:
        return '''
  <div style="border:2px solid #27ae60;border-radius:8px;padding:14px 18px;background:#eafaf1;margin:16px 0;font-weight:600;color:#1e8449;font-size:1.05em">
    ✅ Nenhuma pendência encontrada. Pasta pronta para envio!
  </div>'''

    n = len(all_issues)
    items_html = ''
    for issue in all_issues:
        msg = issue.get('msg', '')
        section, person, text = _parse_issue(msg)
        storage_id = _issue_storage_id(folder, section, person, text)
        iid = _iid(folder, f'{storage_id}|summary|{msg}')
        detail_html = _detail_box(text or msg, section=section, person=person, note=issue.get('note', ''), duplicate_count=issue.get('duplicate_count', 1))
        items_html += f'''
<div id="row_{iid}" data-rowid="{iid}" data-justifiable="{storage_id}" style="display:flex;align-items:flex-start;gap:8px;padding:6px 0;border-bottom:1px solid #fad7d7">
  <div style="flex:1">
    <span style="color:#c0392b">❌ {msg}</span>
    <span id="badge_{iid}" style="display:none;margin-left:6px"></span>
    {detail_html}
    <div id="panel_{iid}" style="display:none;margin-top:7px;background:#fff;border:1px solid #dce3ea;border-radius:6px;padding:10px 12px">
      <div id="display_{iid}" style="display:none;color:#555;font-size:0.88em;margin-bottom:6px;font-style:italic;border-left:3px solid #3498db;padding-left:8px"></div>
      <textarea id="note_{iid}" rows="2" placeholder="Descreva a justificativa ou tratativa…"
                style="width:100%;border:1px solid #ccc;border-radius:4px;padding:6px 8px;font-size:0.9em;resize:vertical;font-family:inherit"></textarea>
      <div style="margin-top:7px;display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <label style="display:flex;align-items:center;gap:5px;font-size:0.9em;cursor:pointer">
          <input type="checkbox" id="chk_{iid}" style="width:15px;height:15px"> ✅ Marcar como tratado
        </label>
        <small id="date_{iid}" style="color:#aaa"></small>
        <button onclick="saveJust('{iid}')"
                style="margin-left:auto;background:#2980b9;color:white;border:none;padding:5px 14px;border-radius:4px;cursor:pointer;font-size:0.88em">
          Salvar
        </button>
      </div>
    </div>
  </div>
  <button onclick="togglePanel('{iid}')" title="Justificar / Tratar"
          style="background:#f39c12;color:white;border:none;padding:4px 9px;border-radius:4px;cursor:pointer;font-size:0.82em;white-space:nowrap;flex-shrink:0;margin-top:2px">
    📝 Justificar
  </button>
</div>'''

    return f'''
  <div class="section" style="border:2px solid #e74c3c;border-radius:8px;margin:16px 0;overflow:hidden">
    <div style="background:#e74c3c;color:white;padding:10px 16px;font-weight:600;font-size:1.05em;display:flex;align-items:center;justify-content:space-between">
      <span>⚠️ Resumo de Pendências</span>
      <span id="just-counter" style="font-size:0.9em;font-weight:normal"></span>
    </div>
    <div id="just-summary" style="background:#fff3f3;padding:8px 16px;font-size:0.9em;color:#666;min-height:10px"></div>
    <div style="padding:12px 16px;background:#fdedec">
      {items_html}
    </div>
  </div>'''


def generate_report(audit_result, output_dir=None):
    competencia = audit_result.get('competencia', '—')
    folder = audit_result.get('folder_path', '')
    timestamp = audit_result.get('timestamp', '')
    overall = audit_result.get('overall_status', 'ok')
    all_issues = _dedupe_issues(audit_result.get('all_issues', []), folder)
    if not all_issues:
        overall = 'ok'

    overall_color = STATUS_COLOR.get(overall, '#333')
    overall_icon = STATUS_ICON.get(overall, '•')

    if overall == 'ok':
        overall_text = 'OK PARA ENVIO'
    elif overall == 'warning':
        overall_text = 'ATENÇÃO — Verificar pontos sinalizados'
    else:
        overall_text = f'PENDÊNCIAS — {len(all_issues)} item(ns) a resolver'

    sections_html = ''.join(_render_section(s, folder) for s in audit_result.get('sections', []))
    issues_html = _render_issues_section(all_issues, folder)

    folder_js = folder.replace('\\', '\\\\').replace("'", "\\'")

    html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Auditoria — {competencia}</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background:#f4f6f8; margin:0; padding:20px; color:#2c3e50; }}
  .container {{ max-width: 920px; margin: 0 auto; }}
  .header {{ background: linear-gradient(135deg, #1a5276, #2980b9); color: white; border-radius:10px; padding:24px 30px; margin-bottom:20px; }}
  .header h1 {{ margin:0 0 6px 0; font-size:1.6em; }}
  .header p {{ margin:0; opacity:0.85; font-size:0.9em; }}
  .overall-banner {{ border-radius:8px; padding:16px 20px; margin:16px 0; font-size:1.15em; font-weight:700;
                     background:{overall_color}; color:white; text-align:center; }}
  .badge {{ border-radius:12px; padding:2px 8px; font-size:0.8em; font-weight:600; white-space:nowrap; }}
  .sec-header {{ transition: background 0.2s; }}
  .sec-header:hover {{ opacity: 0.9; }}
  small {{ font-size:0.85em; }}
  textarea:focus {{ outline:none; border-color:#2980b9 !important; box-shadow:0 0 0 2px rgba(41,128,185,0.15); }}
  @media print {{ .sec-header {{ cursor:default; }} .sec-body {{ display:block !important; }} button {{ display:none; }} }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>🔍 Auditoria Energisa — Horizonte Logística</h1>
    <p>Competência: <strong>{competencia}</strong> &nbsp;|&nbsp; Gerado em: {timestamp}</p>
    <p style="font-size:0.8em;opacity:0.7;margin-top:6px">📁 {folder}</p>
  </div>

  <div class="overall-banner" id="overall-banner">
    {overall_icon} STATUS GERAL: {overall_text}
  </div>

  {issues_html}
  {sections_html}

  <div style="text-align:center;color:#aaa;font-size:0.8em;margin-top:30px;padding:10px">
    Auditoria gerada automaticamente — Apenas leitura, nenhum arquivo foi alterado.
  </div>
</div>

<script>
const _FKEY = 'medicao||{folder_js}';

function _jkey(id) {{ return _FKEY + '||' + id; }}
function _load(id) {{ try {{ return JSON.parse(localStorage.getItem(_jkey(id))); }} catch(e) {{ return null; }} }}
function _save(id, data) {{
  if (!data.note && !data.treated) localStorage.removeItem(_jkey(id));
  else localStorage.setItem(_jkey(id), JSON.stringify(data));
}}

function togglePanel(id) {{
  var p = document.getElementById('panel_' + id);
  p.style.display = p.style.display === 'none' ? 'block' : 'none';
}}

function saveJust(id) {{
  var row = document.getElementById('row_' + id);
  var key = row ? row.dataset.justifiable : id;
  var note = (document.getElementById('note_' + id).value || '').trim();
  var treated = document.getElementById('chk_' + id).checked;
  var now = new Date().toLocaleDateString('pt-BR');
  _save(key, {{note: note, treated: treated, date: now}});
  _applyJust(key, note, treated, now);
  _updateCounters();
}}

function _applyJust(key, note, treated, date) {{
  document.querySelectorAll('[data-justifiable="' + key + '"]').forEach(function(row) {{
    var id = row.dataset.rowid;
    var badge = document.getElementById('badge_' + id);
    var display = document.getElementById('display_' + id);
    var dateEl = document.getElementById('date_' + id);
    var noteEl = document.getElementById('note_' + id);
    var chkEl = document.getElementById('chk_' + id);

    if (noteEl) noteEl.value = note || '';
    if (chkEl) chkEl.checked = !!treated;

    if (treated) {{
      row.style.display = 'none';
      row.dataset.treated = '1';
      if (badge) {{
        badge.innerHTML = '<span style="background:#27ae60;color:white;padding:2px 9px;border-radius:10px;font-size:0.82em">✓ TRATADO</span>';
        badge.style.display = 'inline';
      }}
    }} else if (note) {{
      row.style.display = '';
      row.dataset.treated = '0';
      if (badge) {{
        badge.innerHTML = '<span style="background:#3498db;color:white;padding:2px 9px;border-radius:10px;font-size:0.82em">📝 JUSTIFICADO</span>';
        badge.style.display = 'inline';
      }}
    }} else {{
      row.style.display = '';
      row.dataset.treated = '0';
      if (badge) badge.style.display = 'none';
    }}

    if (display) {{
      if (note) {{ display.textContent = '💬 ' + note; display.style.display = treated ? 'none' : 'block'; }}
      else {{ display.style.display = 'none'; }}
    }}
    if (dateEl && date) dateEl.textContent = date;
  }});
}}

function _updateCounters() {{
  var rows = document.querySelectorAll('[data-justifiable]');
  var states = {{}};
  rows.forEach(function(r) {{
    states[r.dataset.justifiable] = _load(r.dataset.justifiable) || {{}};
  }});
  var keys = Object.keys(states);
  var total = keys.length, treated = 0, justified = 0;
  keys.forEach(function(k) {{
    var d = states[k];
    if (d && d.treated) treated++;
    else if (d && d.note) justified++;
  }});
  var pending = total - treated;

  var counter = document.getElementById('just-counter');
  if (counter) {{
    counter.textContent = treated + '/' + total + ' tratadas';
  }}

  var summary = document.getElementById('just-summary');
  if (summary && total > 0) {{
    var parts = [];
    if (treated > 0) parts.push('<span style="color:#27ae60;font-weight:600">✓ ' + treated + ' tratada(s)</span>');
    if (justified > 0) parts.push('<span style="color:#3498db;font-weight:600">📝 ' + justified + ' justificada(s)</span>');
    if (pending > 0) parts.push('<span style="color:#e74c3c;font-weight:600">⏳ ' + pending + ' em aberto</span>');
    summary.innerHTML = parts.join(' &nbsp;|&nbsp; ');
  }}

  var banner = document.getElementById('overall-banner');
  if (banner && treated === total && total > 0) {{
    banner.style.background = '#27ae60';
    banner.textContent = '✅ TODAS AS PENDÊNCIAS TRATADAS';
  }}
}}

document.addEventListener('DOMContentLoaded', function() {{
  document.querySelectorAll('[data-justifiable]').forEach(function(row) {{
    var id = row.dataset.justifiable;
    var d = _load(id);
    if (d) {{
      _applyJust(id, d.note || '', !!d.treated, d.date || '');
    }}
  }});
  _updateCounters();
}});
</script>
</body>
</html>'''

    if output_dir is None:
        output_dir = tempfile.gettempdir()

    safe_comp = competencia.replace('/', '_').replace(' ', '_')
    report_path = os.path.join(output_dir, f'auditoria_{safe_comp}.html')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return report_path
