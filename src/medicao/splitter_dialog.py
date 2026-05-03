"""
splitter_dialog.py
Dialog para separar um PDF com múltiplos documentos em arquivos individuais.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .pdf_splitter import analyze_pdf, split_pdf, get_label, _LABELS


BG = '#f4f6f8'
HEADER_BG = '#1a3a5c'
WHITE = '#ffffff'
ACCENT = '#2980b9'
GREEN = '#27ae60'
RED = '#e74c3c'
TEXT = '#2c3e50'
GRAY = '#7f8c8d'


class SplitterDialog(tk.Toplevel):
    def __init__(self, parent, comp_folder=''):
        super().__init__(parent)
        self.title('✂️  Separar PDF em Documentos')
        self.geometry('860x580')
        self.resizable(True, True)
        self.configure(bg=BG)
        self.grab_set()

        self.comp_folder = comp_folder
        self.pdf_path = tk.StringVar()
        self.emp_name = tk.StringVar()
        self._groups = []      # lista de dicts editáveis
        self._total_pages = 0

        self._build_ui()

    # ── Layout ───────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=HEADER_BG, pady=12)
        hdr.pack(fill='x')
        tk.Label(hdr, text='✂️  Separar PDF em Documentos', font=('Segoe UI', 15, 'bold'),
                 bg=HEADER_BG, fg=WHITE).pack()
        tk.Label(hdr, text='Detecta automaticamente o tipo de cada documento e separa em arquivos individuais',
                 font=('Segoe UI', 9), bg=HEADER_BG, fg='#aed6f1').pack(pady=(2, 0))

        body = tk.Frame(self, bg=BG, padx=18, pady=10)
        body.pack(fill='both', expand=True)

        # ── Seleção do PDF ────────────────────────────────────────────────
        tk.Label(body, text='PDF de origem:', font=('Segoe UI', 10, 'bold'),
                 bg=BG, fg=TEXT).grid(row=0, column=0, sticky='w', pady=(6, 2))
        row_pdf = tk.Frame(body, bg=BG)
        row_pdf.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        body.columnconfigure(0, weight=1)

        tk.Entry(row_pdf, textvariable=self.pdf_path, font=('Segoe UI', 9),
                 relief='solid', bd=1).pack(side='left', fill='x', expand=True, ipady=4)
        tk.Button(row_pdf, text='Procurar…', font=('Segoe UI', 9), bg=ACCENT, fg=WHITE,
                  relief='flat', bd=0, padx=10, cursor='hand2',
                  command=self._browse_pdf).pack(side='left', padx=(6, 0), ipady=4)

        # ── Nome do colaborador (opcional) ────────────────────────────────
        tk.Label(body, text='Colaborador (opcional — prefixo nos nomes dos arquivos):',
                 font=('Segoe UI', 10, 'bold'), bg=BG, fg=TEXT
                 ).grid(row=2, column=0, sticky='w', pady=(0, 2))
        tk.Entry(body, textvariable=self.emp_name, font=('Segoe UI', 9),
                 relief='solid', bd=1).grid(row=3, column=0, columnspan=2, sticky='ew', ipady=4, pady=(0, 10))

        # ── Botão Analisar ────────────────────────────────────────────────
        self.btn_analyze = tk.Button(body, text='🔍  ANALISAR PDF',
                                     font=('Segoe UI', 11, 'bold'), bg=ACCENT, fg=WHITE,
                                     relief='flat', bd=0, pady=8, cursor='hand2',
                                     command=self._start_analyze)
        self.btn_analyze.grid(row=4, column=0, columnspan=2, sticky='ew', pady=(0, 6))

        self.status_var = tk.StringVar(value='Selecione um PDF e clique em Analisar.')
        self.status_lbl = tk.Label(body, textvariable=self.status_var,
                                   font=('Segoe UI', 9), bg=BG, fg=GRAY, wraplength=800)
        self.status_lbl.grid(row=5, column=0, columnspan=2, sticky='w', pady=(0, 4))

        # ── Progress bar ──────────────────────────────────────────────────
        self.progress = ttk.Progressbar(body, mode='determinate', maximum=100)
        self.progress.grid(row=6, column=0, columnspan=2, sticky='ew', pady=(0, 8))

        # ── Tabela de grupos detectados ───────────────────────────────────
        tk.Label(body, text='Documentos detectados (clique duas vezes para editar o nome):',
                 font=('Segoe UI', 10, 'bold'), bg=BG, fg=TEXT
                 ).grid(row=7, column=0, columnspan=2, sticky='w', pady=(4, 2))

        tree_frame = tk.Frame(body, bg=BG)
        tree_frame.grid(row=8, column=0, columnspan=2, sticky='nsew', pady=(0, 8))
        body.rowconfigure(8, weight=1)

        cols = ('#', 'Tipo detectado', 'Nome do arquivo (editável)', 'Páginas', 'Trecho inicial')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=10)
        for col, w in zip(cols, [40, 160, 220, 80, 300]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=30)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', self._on_edit_row)

        # ── Botão Separar ─────────────────────────────────────────────────
        self.btn_split = tk.Button(body, text='✂️  SEPARAR E SALVAR ARQUIVOS',
                                   font=('Segoe UI', 11, 'bold'), bg=GREEN, fg=WHITE,
                                   relief='flat', bd=0, pady=8, cursor='hand2',
                                   state='disabled', command=self._start_split)
        self.btn_split.grid(row=9, column=0, columnspan=2, sticky='ew', pady=(4, 0))

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _browse_pdf(self):
        init = os.path.dirname(self.pdf_path.get()) if self.pdf_path.get() else \
               self.comp_folder or os.path.expanduser('~')
        path = filedialog.askopenfilename(
            title='Selecione o PDF',
            initialdir=init,
            filetypes=[('PDF', '*.pdf'), ('Todos', '*.*')]
        )
        if path:
            self.pdf_path.set(path)
            # Tenta inferir nome do colaborador pelo nome do arquivo
            base = os.path.splitext(os.path.basename(path))[0]
            if not self.emp_name.get():
                self.emp_name.set(base)

    def _set_status(self, msg, color=None):
        self.status_var.set(msg)
        self.status_lbl.configure(fg=color or GRAY)

    # ── Análise (thread) ──────────────────────────────────────────────────────
    def _start_analyze(self):
        path = self.pdf_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning('Atenção', 'Selecione um arquivo PDF válido.', parent=self)
            return

        self.btn_analyze.configure(state='disabled')
        self.btn_split.configure(state='disabled')
        self.progress['value'] = 0
        self._set_status('Analisando PDF… (páginas escaneadas serão processadas com OCR)', ACCENT)
        self._clear_tree()

        def _progress(current, total):
            pct = int(current / total * 100) if total else 0
            self.after(0, lambda c=current, t=total, p=pct: self._on_progress(c, t, p))

        def _worker():
            try:
                groups = analyze_pdf(path, progress_cb=_progress)
                self.after(0, lambda g=groups: self._analyze_done(g))
            except Exception as exc:
                import traceback
                err = traceback.format_exc()
                self.after(0, lambda e=str(exc): self._analyze_error(e))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_progress(self, current, total, pct):
        self.progress['value'] = pct
        self._set_status(f'Analisando página {current}/{total}…', ACCENT)

    def _analyze_done(self, groups):
        self.progress['value'] = 100
        self.btn_analyze.configure(state='normal')
        self._groups = groups

        if not groups:
            self._set_status('Nenhum documento reconhecido. Verifique se o PDF contém texto ou imagens legíveis.', RED)
            return

        self._populate_tree(groups)
        n = len(groups)
        self._set_status(f'{n} documento(s) detectado(s). Revise os nomes e clique em Separar.', GREEN)
        self.btn_split.configure(state='normal')

    def _analyze_error(self, err):
        self.progress['value'] = 0
        self.btn_analyze.configure(state='normal')
        self._set_status(f'Erro: {err}', RED)
        messagebox.showerror('Erro na análise', err, parent=self)

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _populate_tree(self, groups):
        self._clear_tree()
        emp = self.emp_name.get().strip().upper()

        for i, g in enumerate(groups):
            label = g.get('label', get_label(g.get('type', '')))
            pages = g.get('pages', [])
            page_str = f'{pages[0]+1}–{pages[-1]+1}' if pages else '—'
            trecho = g.get('start_text', '')[:80].replace('\n', ' ')
            prefix = f"{emp} - " if emp else ''
            display_name = f"{prefix}{label}"
            self.tree.insert('', 'end', iid=str(i),
                             values=(i + 1, g.get('type', '?'), display_name, page_str, trecho))

    # ── Edição inline de linha ────────────────────────────────────────────────
    def _on_edit_row(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = int(iid)
        row_vals = self.tree.item(iid, 'values')
        current_name = row_vals[2] if len(row_vals) > 2 else ''

        dlg = tk.Toplevel(self)
        dlg.title('Editar nome do documento')
        dlg.geometry('500x160')
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg=BG)

        tk.Label(dlg, text='Nome do arquivo (sem .pdf):', font=('Segoe UI', 10),
                 bg=BG, fg=TEXT).pack(anchor='w', padx=18, pady=(18, 4))

        var = tk.StringVar(value=current_name)
        entry = tk.Entry(dlg, textvariable=var, font=('Segoe UI', 10),
                         relief='solid', bd=1, width=55)
        entry.pack(fill='x', padx=18, ipady=5)
        entry.select_range(0, 'end')
        entry.focus_set()

        def _apply():
            new_name = var.get().strip()
            if not new_name:
                return
            vals = list(self.tree.item(iid, 'values'))
            vals[2] = new_name
            self.tree.item(iid, values=vals)
            self._groups[idx]['label'] = new_name
            dlg.destroy()

        btn_row = tk.Frame(dlg, bg=BG)
        btn_row.pack(fill='x', padx=18, pady=12)
        tk.Button(btn_row, text='Cancelar', font=('Segoe UI', 9), bg='#bdc3c7',
                  relief='flat', padx=14, command=dlg.destroy).pack(side='right', padx=(6, 0))
        tk.Button(btn_row, text='Confirmar', font=('Segoe UI', 9, 'bold'), bg=ACCENT, fg=WHITE,
                  relief='flat', padx=14, command=_apply).pack(side='right')
        dlg.bind('<Return>', lambda _: _apply())

    # ── Split (thread) ────────────────────────────────────────────────────────
    def _start_split(self):
        pdf = self.pdf_path.get().strip()
        emp = self.emp_name.get().strip()

        if not pdf or not os.path.isfile(pdf):
            messagebox.showwarning('Atenção', 'PDF de origem inválido.', parent=self)
            return
        if not self._groups:
            messagebox.showwarning('Atenção', 'Analise o PDF primeiro.', parent=self)
            return

        # Pasta de saída = mesma pasta do PDF de origem
        out_folder = os.path.dirname(os.path.abspath(pdf))

        # Atualiza labels da tabela para os grupos (o usuário pode ter editado)
        for i, iid in enumerate(self.tree.get_children()):
            vals = self.tree.item(iid, 'values')
            if i < len(self._groups) and len(vals) > 2:
                self._groups[i]['label'] = vals[2]

        self.btn_split.configure(state='disabled')
        self.btn_analyze.configure(state='disabled')
        self.progress['value'] = 0
        self._set_status('Separando e salvando arquivos…', ACCENT)

        def _worker():
            try:
                saved = split_pdf(pdf, out_folder, self._groups, employee_name=emp)
                self.after(0, lambda s=saved, o=out_folder: self._split_done(s, o))
            except Exception as exc:
                self.after(0, lambda e=str(exc): self._split_error(e))

        threading.Thread(target=_worker, daemon=True).start()

    def _split_done(self, saved, out_folder):
        self.progress['value'] = 100
        self.btn_analyze.configure(state='normal')
        self.btn_split.configure(state='normal')
        n = len(saved)
        self._set_status(f'✅ {n} arquivo(s) salvo(s) em: {out_folder}', GREEN)
        names = '\n'.join(f'  ✓ {os.path.basename(p)}' for p, _ in saved)
        messagebox.showinfo(
            'Concluído',
            f'{n} arquivo(s) criado(s) com sucesso!\n\n{names}',
            parent=self
        )

    def _split_error(self, err):
        self.progress['value'] = 0
        self.btn_analyze.configure(state='normal')
        self.btn_split.configure(state='normal')
        self._set_status(f'Erro: {err}', RED)
        messagebox.showerror('Erro ao separar', err, parent=self)
