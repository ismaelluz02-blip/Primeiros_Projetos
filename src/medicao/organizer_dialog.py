import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .utils import normalize, find_folder
from .auditor import find_forca_trabalho, FOLDER_PATTERNS
from .excel_reader import read_forca_trabalho

# ── Mapeamento: tipo → (chave_pasta, precisa_colaborador) ──────────────────
DOC_TYPE_DESTINATION = {
    'CONTRATO':            ('admissao',       True),
    'FICHA_REGISTRO':      ('admissao',       True),
    'CTPS':                ('admissao',       True),
    'ESOCIAL':             ('admissao',       True),
    'ASO_ADMISSIONAL':     ('admissao',       True),
    'EPI':                 ('admissao',       True),
    'VT_DECLARACAO':       ('admissao',       True),
    'TRCT':                ('demissao',       True),
    'COMPROVANTE_RESCISAO':('demissao',       True),
    'AVISO_PREVIO':        ('demissao',       True),
    'GRRF_MULTA':          ('demissao',       True),
    'SEGURO_DESEMPREGO':   ('demissao',       True),
    'ESOCIAL_DEMISSAO':    ('demissao',       True),
    'ASO_DEMISSIONAL':     ('demissao',       True),
    'QUITACAO':            ('demissao',       True),
    'HOMOLOGACAO':         ('demissao',       True),
    'AVISO_FERIAS':        ('ferias',         True),
    'RECIBO_FERIAS':       ('ferias',         True),
    'COMPROVANTE_FERIAS':  ('ferias',         False),
    'FOPAG':               ('fopag',          False),
    'COMPROVANTE_SALARIO': ('fopag',          False),
    'ADIANTAMENTO':        ('fopag',          False),
    'SALDO_SALARIO':       ('fopag',          False),
    'GUIA_FGTS':           ('inss_fgts',      False),
    'COMPROVANTE_FGTS':    ('inss_fgts',      False),
    'DETALHAMENTO_FGTS':   ('inss_fgts',      False),
    'CRF':                 ('inss_fgts',      False),
    'DCTFWEB':             ('inss_fgts',      False),
    'RECIBO_DCTFWEB':      ('inss_fgts',      False),
    'DARF_INSS':           ('inss_fgts',      False),
    'RELATORIO_VT':        ('vt',             False),
    'BOLETO_VT':           ('vt',             False),
    'NF_VT':               ('vt',             False),
    'COMPROVANTE_VT':      ('vt',             False),
    'RELATORIO_VAVR':      ('va_vr',          False),
    'BOLETO_VAVR':         ('va_vr',          False),
    'NF_VAVR':             ('va_vr',          False),
    'COMPROVANTE_VAVR':    ('va_vr',          False),
    'ACORDO_COLETIVO':     ('acordo_coletivo',False),
    'DECL_ADMISSAO':       ('declaracoes',    False),
    'DECL_DEMISSAO':       ('declaracoes',    False),
    'DECL_FERIAS':         ('declaracoes',    False),
    'DECL_ACIDENTE':       ('declaracoes',    False),
    'DECL_MOBILIZACAO':    ('declaracoes',    False),
    'DECL_SUBCONTRATACAO': ('declaracoes',    False),
    'DECL_MUDANCA_FUNCAO': ('declaracoes',    False),
    'DECL_TRANSFERENCIA':  ('declaracoes',    False),
}

DOC_TYPE_LABELS = {
    '':                    '— Selecionar tipo —',
    'CONTRATO':            'Contrato de Trabalho',
    'FICHA_REGISTRO':      'Ficha de Registro',
    'CTPS':                'CTPS Digital',
    'ESOCIAL':             'eSocial (Admissão)',
    'ASO_ADMISSIONAL':     'ASO Admissional',
    'EPI':                 'Entrega de EPI / Fardas',
    'VT_DECLARACAO':       'Declaração de Opção VT',
    'TRCT':                'Termo de Rescisão (TRCT)',
    'COMPROVANTE_RESCISAO':'Comprovante Pagto. Rescisão',
    'AVISO_PREVIO':        'Aviso Prévio / Pedido de Demissão',
    'GRRF_MULTA':          'GRRF – Multa 40% FGTS',
    'SEGURO_DESEMPREGO':   'Seguro Desemprego',
    'ESOCIAL_DEMISSAO':    'eSocial / Baixa Digital',
    'ASO_DEMISSIONAL':     'ASO Demissional',
    'QUITACAO':            'Termo de Quitação',
    'HOMOLOGACAO':         'Comunicado de Homologação',
    'AVISO_FERIAS':        'Aviso de Férias',
    'RECIBO_FERIAS':       'Recibo de Férias',
    'COMPROVANTE_FERIAS':  'Comprovante Pgto. Férias',
    'FOPAG':               'Folha de Pagamento (FOPAG)',
    'COMPROVANTE_SALARIO': 'Comprovante de Salário',
    'ADIANTAMENTO':        'Adiantamento de Salário',
    'SALDO_SALARIO':       'Saldo de Salário',
    'GUIA_FGTS':           'Guia FGTS',
    'COMPROVANTE_FGTS':    'Comprovante FGTS',
    'DETALHAMENTO_FGTS':   'Detalhamento FGTS',
    'CRF':                 'CRF – Certificado Regularidade FGTS',
    'DCTFWEB':             'DCTFWeb',
    'RECIBO_DCTFWEB':      'Recibo DCTFWeb',
    'DARF_INSS':           'DARF INSS',
    'RELATORIO_VT':        'Relatório / Pedido VT',
    'BOLETO_VT':           'Boleto VT',
    'NF_VT':               'Nota Fiscal VT',
    'COMPROVANTE_VT':      'Comprovante VT',
    'RELATORIO_VAVR':      'Relatório Colaboradores VA/VR',
    'BOLETO_VAVR':         'Boleto VA/VR',
    'NF_VAVR':             'Nota Fiscal VA/VR',
    'COMPROVANTE_VAVR':    'Comprovante VA/VR',
    'ACORDO_COLETIVO':     'Acordo Coletivo (CCT)',
    'DECL_ADMISSAO':       'Declaração de Admissão',
    'DECL_DEMISSAO':       'Declaração de Demissão',
    'DECL_FERIAS':         'Declaração de Férias',
    'DECL_ACIDENTE':       'Declaração de Acidente',
    'DECL_MOBILIZACAO':    'Declaração de Mobilização',
    'DECL_SUBCONTRATACAO': 'Declaração de Subcontratação',
    'DECL_MUDANCA_FUNCAO': 'Declaração de Mudança de Função',
    'DECL_TRANSFERENCIA':  'Declaração de Transferência',
}

# Sorted list for the comboboxes
TYPE_OPTIONS = [''] + sorted(
    (k for k in DOC_TYPE_LABELS if k),
    key=lambda k: DOC_TYPE_LABELS[k]
)

BG = '#f4f6f8'
BG_DARK = '#1a3a5c'
BG_MID = '#2c5f8a'
WHITE = '#ffffff'
ACCENT = '#2980b9'
GREEN = '#27ae60'
RED = '#e74c3c'
TEXT_DARK = '#2c3e50'
TEXT_GRAY = '#7f8c8d'


def _safe_name(s):
    return s.replace('/', '-').replace('\\', '-').replace(':', '-').strip()


def _unique_path(dest_path):
    if not os.path.exists(dest_path):
        return dest_path
    base, ext = os.path.splitext(dest_path)
    i = 2
    while os.path.exists(f'{base} ({i}){ext}'):
        i += 1
    return f'{base} ({i}){ext}'


def compute_destination(comp_folder, doc_type, employee_name):
    if not doc_type or doc_type not in DOC_TYPE_DESTINATION:
        return '', ''

    section_key, per_employee = DOC_TYPE_DESTINATION[doc_type]
    section_folder = find_folder(comp_folder, FOLDER_PATTERNS.get(section_key, [section_key]))

    label = DOC_TYPE_LABELS.get(doc_type, doc_type)

    if per_employee and employee_name:
        emp_safe = _safe_name(employee_name)
        if section_folder:
            dest_dir = os.path.join(section_folder, emp_safe)
        else:
            dest_dir = os.path.join(comp_folder, _get_default_folder_name(section_key), emp_safe)
        new_name = f'{label} - {emp_safe}.pdf'
    elif per_employee and not employee_name:
        return '⚠ Selecione um colaborador', ''
    else:
        if section_folder:
            dest_dir = section_folder
        else:
            dest_dir = os.path.join(comp_folder, _get_default_folder_name(section_key))
        new_name = f'{label}.pdf'

    return dest_dir, new_name


def _get_default_folder_name(section_key):
    defaults = {
        'admissao': 'Admissão-alocação',
        'demissao': 'Demissão-transferência',
        'ferias': 'FÉRIAS',
        'fopag': 'FOPAG e Comp. de Pgto',
        'inss_fgts': 'INSS + FGTS',
        'ponto': 'Ponto',
        'va_vr': 'VA E VR',
        'vt': 'VT',
        'acordo_coletivo': 'ACORDO COLETIVO',
        'declaracoes': 'DECLARAÇÕES',
    }
    return defaults.get(section_key, section_key.upper())


class EditRowDialog(tk.Toplevel):
    def __init__(self, parent, filename, doc_type, employee, employee_list, comp_folder, callback):
        super().__init__(parent)
        self.title('Editar Classificação')
        self.resizable(False, False)
        self.configure(bg=WHITE)
        self.grab_set()
        self.callback = callback
        self.comp_folder = comp_folder
        self.employee_list = employee_list

        pad = {'padx': 12, 'pady': 6}

        tk.Label(self, text='Arquivo:', font=('Segoe UI', 9, 'bold'), bg=WHITE, fg=TEXT_DARK).grid(
            row=0, column=0, sticky='w', **pad)
        tk.Label(self, text=filename, font=('Segoe UI', 9), bg=WHITE, fg=ACCENT).grid(
            row=0, column=1, sticky='w', **pad)

        tk.Label(self, text='Tipo de Documento:', font=('Segoe UI', 9, 'bold'), bg=WHITE, fg=TEXT_DARK).grid(
            row=1, column=0, sticky='w', **pad)
        self.type_var = tk.StringVar(value=doc_type)
        type_cb = ttk.Combobox(self, textvariable=self.type_var, width=38,
                               values=[DOC_TYPE_LABELS[k] for k in TYPE_OPTIONS], state='readonly')
        type_cb.grid(row=1, column=1, sticky='w', **pad)
        if doc_type in DOC_TYPE_LABELS:
            type_cb.set(DOC_TYPE_LABELS[doc_type])
        type_cb.bind('<<ComboboxSelected>>', self._update_dest)

        tk.Label(self, text='Colaborador:', font=('Segoe UI', 9, 'bold'), bg=WHITE, fg=TEXT_DARK).grid(
            row=2, column=0, sticky='w', **pad)
        self.emp_var = tk.StringVar(value=employee)
        emp_opts = ['— Nenhum (doc. coletivo) —'] + employee_list
        emp_cb = ttk.Combobox(self, textvariable=self.emp_var, width=38,
                              values=emp_opts, state='readonly')
        emp_cb.grid(row=2, column=1, sticky='w', **pad)
        emp_cb.set(employee if employee else '— Nenhum (doc. coletivo) —')
        emp_cb.bind('<<ComboboxSelected>>', self._update_dest)

        tk.Label(self, text='Destino:', font=('Segoe UI', 9, 'bold'), bg=WHITE, fg=TEXT_DARK).grid(
            row=3, column=0, sticky='w', **pad)
        self.dest_var = tk.StringVar()
        tk.Label(self, textvariable=self.dest_var, font=('Segoe UI', 8), bg=WHITE,
                 fg=TEXT_GRAY, wraplength=340, justify='left').grid(row=3, column=1, sticky='w', **pad)

        tk.Label(self, text='Novo nome:', font=('Segoe UI', 9, 'bold'), bg=WHITE, fg=TEXT_DARK).grid(
            row=4, column=0, sticky='w', **pad)
        self.name_var = tk.StringVar()
        tk.Label(self, textvariable=self.name_var, font=('Segoe UI', 8), bg=WHITE,
                 fg=GREEN, wraplength=340).grid(row=4, column=1, sticky='w', **pad)

        btn_frame = tk.Frame(self, bg=WHITE)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=12)
        tk.Button(btn_frame, text='  OK  ', bg=GREEN, fg=WHITE, font=('Segoe UI', 9, 'bold'),
                  relief='flat', padx=12, command=self._ok).pack(side='left', padx=6)
        tk.Button(btn_frame, text='  Cancelar  ', bg=TEXT_GRAY, fg=WHITE, font=('Segoe UI', 9),
                  relief='flat', padx=8, command=self.destroy).pack(side='left', padx=6)

        self._type_key = doc_type
        self._update_dest()

    def _get_type_key(self):
        label = self.type_var.get()
        for k, v in DOC_TYPE_LABELS.items():
            if v == label:
                return k
        return ''

    def _get_employee(self):
        v = self.emp_var.get()
        return '' if v.startswith('—') else v

    def _update_dest(self, *_):
        key = self._get_type_key()
        emp = self._get_employee()
        dest_dir, new_name = compute_destination(self.comp_folder, key, emp)
        if dest_dir:
            rel = os.path.relpath(dest_dir, self.comp_folder) if self.comp_folder else dest_dir
            self.dest_var.set(rel)
        else:
            self.dest_var.set('—')
        self.name_var.set(new_name or '—')
        self._type_key = key

    def _ok(self):
        key = self._get_type_key()
        emp = self._get_employee()
        self.callback(key, emp)
        self.destroy()


class OrganizerDialog(tk.Toplevel):
    def __init__(self, parent, comp_folder=''):
        super().__init__(parent)
        self.title('Organizar Documentos Escaneados')
        self.geometry('920x600')
        self.configure(bg=BG)
        self.comp_folder = tk.StringVar(value=comp_folder)
        self.source_folder = tk.StringVar()
        self.employee_list = []
        self.rows = []  # list of dicts: {filepath, doc_type, employee, iid}
        self._build_ui()
        if comp_folder:
            self._load_employees()

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=BG_DARK, pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='📂  Organizar Documentos Escaneados', font=('Segoe UI', 13, 'bold'),
                 bg=BG_DARK, fg=WHITE).pack()
        tk.Label(hdr, text='Selecione a pasta dos arquivos escaneados e a competência destino',
                 font=('Segoe UI', 9), bg=BG_DARK, fg='#aed6f1').pack()

        # Folder selectors
        sel = tk.Frame(self, bg=BG, pady=8, padx=16)
        sel.pack(fill='x')

        for row_i, (label, var, cmd) in enumerate([
            ('Pasta dos escaneados:', self.source_folder, self._pick_source),
            ('Competência (destino):', self.comp_folder, self._pick_comp),
        ]):
            tk.Label(sel, text=label, font=('Segoe UI', 9, 'bold'), bg=BG, fg=TEXT_DARK, width=22,
                     anchor='w').grid(row=row_i, column=0, sticky='w', pady=3)
            tk.Entry(sel, textvariable=var, font=('Segoe UI', 9), width=60,
                     relief='solid', bd=1).grid(row=row_i, column=1, sticky='ew', padx=(4, 6))
            tk.Button(sel, text='Procurar…', font=('Segoe UI', 8), bg=ACCENT, fg=WHITE,
                      relief='flat', command=cmd, padx=6).grid(row=row_i, column=2)

        tk.Button(sel, text='  🔍  Carregar Arquivos  ', font=('Segoe UI', 9, 'bold'),
                  bg=BG_MID, fg=WHITE, relief='flat', command=self._load_files, pady=4
                  ).grid(row=2, column=0, columnspan=3, sticky='e', pady=(8, 0))

        # Treeview
        tv_frame = tk.Frame(self, bg=BG)
        tv_frame.pack(fill='both', expand=True, padx=16, pady=(0, 4))

        cols = ('arquivo', 'tipo', 'colaborador', 'destino')
        self.tree = ttk.Treeview(tv_frame, columns=cols, show='headings', selectmode='browse')
        self.tree.heading('arquivo',     text='Arquivo Original')
        self.tree.heading('tipo',        text='Tipo do Documento')
        self.tree.heading('colaborador', text='Colaborador')
        self.tree.heading('destino',     text='Destino / Novo Nome')
        self.tree.column('arquivo',     width=200, stretch=False)
        self.tree.column('tipo',        width=240, stretch=False)
        self.tree.column('colaborador', width=180, stretch=False)
        self.tree.column('destino',     width=260)

        vsb = ttk.Scrollbar(tv_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', self._on_double_click)

        # Tag colors
        self.tree.tag_configure('ok',      background='#eafaf1')
        self.tree.tag_configure('pending', background='#fef9e7')
        self.tree.tag_configure('error',   background='#fdedec')

        # Footer buttons
        foot = tk.Frame(self, bg=BG, pady=8, padx=16)
        foot.pack(fill='x')
        tk.Label(foot, text='Duplo clique para editar uma linha',
                 font=('Segoe UI', 8), bg=BG, fg=TEXT_GRAY).pack(side='left')
        tk.Button(foot, text='  ✅  Organizar Arquivos  ', font=('Segoe UI', 10, 'bold'),
                  bg=GREEN, fg=WHITE, relief='flat', pady=6, padx=12,
                  command=self._organize).pack(side='right', padx=(8, 0))
        tk.Button(foot, text='Fechar', font=('Segoe UI', 9), bg=TEXT_GRAY, fg=WHITE,
                  relief='flat', pady=6, padx=10, command=self.destroy).pack(side='right')

    def _pick_source(self):
        d = filedialog.askdirectory(title='Pasta dos arquivos escaneados', parent=self)
        if d:
            self.source_folder.set(d)

    def _pick_comp(self):
        d = filedialog.askdirectory(title='Pasta da competência destino', parent=self)
        if d:
            self.comp_folder.set(d)
            self._load_employees()

    def _load_employees(self):
        comp = self.comp_folder.get()
        if not comp or not os.path.isdir(comp):
            return
        ft_path = find_forca_trabalho(comp)
        if ft_path:
            employees, _ = read_forca_trabalho(ft_path)
            self.employee_list = sorted({e['nome'] for e in employees if e['nome']})

    def _load_files(self):
        src = self.source_folder.get().strip()
        comp = self.comp_folder.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showwarning('Atenção', 'Selecione a pasta dos arquivos escaneados.', parent=self)
            return
        if not comp or not os.path.isdir(comp):
            messagebox.showwarning('Atenção', 'Selecione a pasta da competência destino.', parent=self)
            return

        self._load_employees()
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.rows.clear()

        from .pdf_reader import identify_doc_types
        pdfs = [f for f in os.listdir(src)
                if os.path.isfile(os.path.join(src, f)) and f.lower().endswith('.pdf')]

        if not pdfs:
            messagebox.showinfo('Sem arquivos', 'Nenhum PDF encontrado na pasta selecionada.', parent=self)
            return

        for f in sorted(pdfs):
            filepath = os.path.join(src, f)
            # Try filename-based classification first
            tags = identify_doc_types(filepath)
            doc_type = next(iter(tags)) if tags else ''
            employee = ''
            self._add_row(filepath, f, doc_type, employee)

    def _add_row(self, filepath, filename, doc_type, employee):
        dest_dir, new_name = compute_destination(self.comp_folder.get(), doc_type, employee)
        type_label = DOC_TYPE_LABELS.get(doc_type, '— Selecionar tipo —')
        dest_text = ''
        if dest_dir and new_name:
            rel = os.path.relpath(dest_dir, self.comp_folder.get()) if self.comp_folder.get() else dest_dir
            dest_text = f'{rel}  ›  {new_name}'
        tag = 'ok' if (doc_type and (not DOC_TYPE_DESTINATION.get(doc_type, ('', False))[1] or employee)) \
              else 'pending' if doc_type else 'error'
        iid = self.tree.insert('', 'end',
                               values=(filename, type_label, employee or '—', dest_text or '⚠ Classificar'),
                               tags=(tag,))
        self.rows.append({
            'filepath': filepath,
            'filename': filename,
            'doc_type': doc_type,
            'employee': employee,
            'iid': iid,
        })

    def _on_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        row = next((r for r in self.rows if r['iid'] == iid), None)
        if not row:
            return

        def _cb(new_type, new_emp):
            row['doc_type'] = new_type
            row['employee'] = new_emp
            dest_dir, new_name = compute_destination(self.comp_folder.get(), new_type, new_emp)
            type_label = DOC_TYPE_LABELS.get(new_type, '— Selecionar tipo —')
            dest_text = ''
            if dest_dir and new_name:
                rel = os.path.relpath(dest_dir, self.comp_folder.get()) if self.comp_folder.get() else dest_dir
                dest_text = f'{rel}  ›  {new_name}'
            tag = 'ok' if (new_type and (not DOC_TYPE_DESTINATION.get(new_type, ('', False))[1] or new_emp)) \
                  else 'pending' if new_type else 'error'
            self.tree.item(iid, values=(row['filename'], type_label, new_emp or '—',
                                        dest_text or '⚠ Classificar'), tags=(tag,))

        EditRowDialog(self, row['filename'], row['doc_type'], row['employee'],
                      self.employee_list, self.comp_folder.get(), _cb)

    def _organize(self):
        comp = self.comp_folder.get().strip()
        if not comp or not os.path.isdir(comp):
            messagebox.showwarning('Atenção', 'Competência destino inválida.', parent=self)
            return

        ready = [r for r in self.rows if r['doc_type']]
        skip = [r for r in self.rows if not r['doc_type']]

        if not ready:
            messagebox.showwarning('Nada a organizar',
                                   'Nenhum arquivo foi classificado ainda.\nDê duplo clique nas linhas para classificar.',
                                   parent=self)
            return

        # Check for missing employees
        needs_emp = [r for r in ready
                     if DOC_TYPE_DESTINATION.get(r['doc_type'], ('', False))[1] and not r['employee']]
        if needs_emp:
            names = '\n'.join(f'  • {r["filename"]}' for r in needs_emp)
            if not messagebox.askyesno('Colaborador faltando',
                                       f'Os arquivos abaixo precisam de colaborador e serão ignorados:\n{names}\n\n'
                                       f'Continuar com os demais?', parent=self):
                return
            ready = [r for r in ready if r not in needs_emp]

        if skip:
            names = '\n'.join(f'  • {r["filename"]}' for r in skip)
            if not messagebox.askyesno('Arquivos sem classificação',
                                       f'Os arquivos abaixo NÃO serão movidos (sem tipo definido):\n{names}\n\n'
                                       f'Continuar com os classificados?', parent=self):
                return

        moved, errors = [], []
        for row in ready:
            dest_dir, new_name = compute_destination(comp, row['doc_type'], row['employee'])
            if not dest_dir or not new_name:
                errors.append(f'{row["filename"]}: destino não determinado')
                continue
            try:
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = _unique_path(os.path.join(dest_dir, new_name))
                shutil.move(row['filepath'], dest_path)
                moved.append(f'{row["filename"]}  →  {os.path.relpath(dest_path, comp)}')
            except Exception as e:
                errors.append(f'{row["filename"]}: {e}')

        summary = f'✅ {len(moved)} arquivo(s) organizado(s).'
        if errors:
            summary += f'\n\n❌ {len(errors)} erro(s):\n' + '\n'.join(errors)
        messagebox.showinfo('Concluído', summary, parent=self)

        # Refresh tree — remove successfully moved files
        moved_names = {r['filename'] for r in ready if not any(r['filename'] in e for e in errors)}
        for row in list(self.rows):
            if row['filename'] in moved_names:
                self.tree.delete(row['iid'])
                self.rows.remove(row)
