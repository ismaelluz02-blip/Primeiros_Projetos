import os
import tempfile
import hashlib

STATUS_ICON = {'ok': '✅', 'error': '❌', 'warning': '⚠️', 'info': 'ℹ️'}
STATUS_COLOR = {'ok': '#27ae60', 'error': '#e74c3c', 'warning': '#f39c12', 'info': '#3498db'}
STATUS_BG = {'ok': '#eafaf1', 'error': '#fdedec', 'warning': '#fef9e7', 'info': '#eaf4fb'}


def _iid(folder, text):
    """Stable short ID for an item — used as localStorage sub-key."""
    return hashlib.md5(f"{folder}|{text}".encode('utf-8')).hexdigest()[:10]


def _render_item(item, folder, justifiable=False):
    st = item.get('status', 'info')
    icon = STATUS_ICON.get(st, '•')
    color = STATUS_COLOR.get(st, '#555')
    label = item.get('label', '')
    note = item.get('note', '')
    note_html = f' <small style="color:#888">— {note}</small>' if note else ''

    if justifiable and st in ('error', 'warning'):
        iid = _iid(folder, label + note)
        return f'''
<div id="row_{iid}" data-justifiable="{iid}" style="display:flex;align-items:flex-start;gap:8px;margin:5px 0">
  <div style="flex:1">
    <span style="color:{color}">{icon} {label}{note_html}</span>
    <span id="badge_{iid}" style="display:none;margin-left:6px"></span>
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


def _render_items(items, folder='', justifiable=False):
    if not items:
        return ''
    has_just = justifiable and any(i.get('status') in ('error', 'warning') for i in items)
    if has_just:
        rows = ''.join(_render_item(i, folder, justifiable=True) for i in items)
        return f'<div style="padding-left:8px;margin:6px 0">{rows}</div>'
    rows = ''.join(_render_item(i, folder) for i in items)
    return f'<ul style="list-style:none;padding-left:8px;margin:6px 0">{rows}</ul>'


def _render_employee(emp, folder):
    st = emp.get('status', 'ok')
    icon = STATUS_ICON.get(st, '•')
    color = STATUS_COLOR.get(st, '#333')
    bg = STATUS_BG.get(st, '#fff')
    name = emp.get('name', '')
    items_html = _render_items(emp.get('items', []), folder=folder, justifiable=True)
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

    items_html = _render_items(sec.get('items', []), folder=folder, justifiable=True)
    employees_html = ''.join(_render_employee(e, folder) for e in sec.get('employees', []))

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
    if not all_issues:
        return '''
  <div style="border:2px solid #27ae60;border-radius:8px;padding:14px 18px;background:#eafaf1;margin:16px 0;font-weight:600;color:#1e8449;font-size:1.05em">
    ✅ Nenhuma pendência encontrada. Pasta pronta para envio!
  </div>'''

    n = len(all_issues)
    items_html = ''
    for issue in all_issues:
        msg = issue.get('msg', '')
        iid = _iid(folder, msg)
        items_html += f'''
<div id="row_{iid}" data-justifiable="{iid}" style="display:flex;align-items:flex-start;gap:8px;padding:6px 0;border-bottom:1px solid #fad7d7">
  <div style="flex:1">
    <span style="color:#c0392b">❌ {msg}</span>
    <span id="badge_{iid}" style="display:none;margin-left:6px"></span>
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
    all_issues = audit_result.get('all_issues', [])

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
  var note = (document.getElementById('note_' + id).value || '').trim();
  var treated = document.getElementById('chk_' + id).checked;
  var now = new Date().toLocaleDateString('pt-BR');
  _save(id, {{note: note, treated: treated, date: now}});
  _applyJust(id, note, treated, now);
  _updateCounters();
}}

function _applyJust(id, note, treated, date) {{
  var row = document.getElementById('row_' + id);
  var badge = document.getElementById('badge_' + id);
  var display = document.getElementById('display_' + id);
  var dateEl = document.getElementById('date_' + id);
  if (!row) return;

  if (treated) {{
    row.style.opacity = '0.55';
    badge.innerHTML = '<span style="background:#27ae60;color:white;padding:2px 9px;border-radius:10px;font-size:0.82em">✓ TRATADO</span>';
    badge.style.display = 'inline';
  }} else if (note) {{
    row.style.opacity = '1';
    badge.innerHTML = '<span style="background:#3498db;color:white;padding:2px 9px;border-radius:10px;font-size:0.82em">📝 JUSTIFICADO</span>';
    badge.style.display = 'inline';
  }} else {{
    row.style.opacity = '1';
    badge.style.display = 'none';
  }}

  if (display) {{
    if (note) {{ display.textContent = '💬 ' + note; display.style.display = 'block'; }}
    else {{ display.style.display = 'none'; }}
  }}
  if (dateEl && date) dateEl.textContent = date;
}}

function _updateCounters() {{
  var rows = document.querySelectorAll('[data-justifiable]');
  var treated = 0, justified = 0, total = rows.length;
  rows.forEach(function(r) {{
    var d = _load(r.dataset.justifiable);
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
      var noteEl = document.getElementById('note_' + id);
      var chkEl = document.getElementById('chk_' + id);
      if (noteEl) noteEl.value = d.note || '';
      if (chkEl) chkEl.checked = !!d.treated;
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
