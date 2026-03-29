import os
import glob
import re
import atexit
import ctypes
import sqlite3
import shutil
import sys
import unicodedata
import calendar
import time
import json
import getpass
import socket
from calendar import monthrange
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import fitz
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageFilter

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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else SCRIPT_DIR
BASE_DIR = SCRIPT_DIR
PROJECT_DIR = APP_DIR if getattr(sys, "frozen", False) else os.path.dirname(BASE_DIR)
RELATORIOS_DIR = os.path.join(PROJECT_DIR, "RELATORIOS")
DEFAULT_APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", BASE_DIR),
    "Horizonte Logistica",
    "Sistema de Faturamento",
)
APP_DATA_DIR = DEFAULT_APP_DATA_DIR
SYNC_CONFIG_SCHEMA_VERSION = 1
DB_PATH = os.path.join(APP_DATA_DIR, "faturamento.db")
LOCK_PATH = os.path.join(APP_DATA_DIR, ".sistema_faturamento.lock")
LEGACY_DB_PATH = os.path.join(APP_DIR, "faturamento.db")
LOGO_PATH = os.path.join(APP_DIR, "logo.png")
APP_USER_MODEL_ID = "horizonte.logistica.sistema.faturamento"
FALLBACK_APP_DATA_DIR = os.path.join(BASE_DIR, "_dados_app")


def _diretorio_gravavel(caminho_dir):
    try:
        os.makedirs(caminho_dir, exist_ok=True)
        arquivo_teste = os.path.join(caminho_dir, ".write_test.tmp")
        with open(arquivo_teste, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(arquivo_teste)
        return True
    except OSError:
        return False


def configurar_diretorio_dados():
    global APP_DATA_DIR, DB_PATH, LOCK_PATH

    dir_original = APP_DATA_DIR
    dir_escolhido = None
    for candidato in (APP_DATA_DIR, FALLBACK_APP_DATA_DIR, BASE_DIR):
        if _diretorio_gravavel(candidato):
            dir_escolhido = candidato
            break

    if not dir_escolhido:
        dir_escolhido = BASE_DIR

    APP_DATA_DIR = dir_escolhido
    DB_PATH = os.path.join(APP_DATA_DIR, "faturamento.db")
    LOCK_PATH = os.path.join(APP_DATA_DIR, ".sistema_faturamento.lock")

    # Se o AppData estiver sem permissao de escrita, migra o banco para pasta local.
    if APP_DATA_DIR != dir_original:
        origem_db = os.path.join(dir_original, "faturamento.db")
        if os.path.exists(origem_db) and not os.path.exists(DB_PATH):
            try:
                shutil.copy2(origem_db, DB_PATH)
            except OSError:
                pass


configurar_diretorio_dados()


def _processo_ativo(pid):
    try:
        os.kill(pid, 0)
        return True
    except:
        return False


def preparar_arquivos_aplicacao():
    try:
        os.makedirs(APP_DATA_DIR, exist_ok=True)
    except OSError:
        configurar_diretorio_dados()
        os.makedirs(APP_DATA_DIR, exist_ok=True)

    def _contar_documentos(caminho_db):
        try:
            conn = sqlite3.connect(caminho_db)
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM documentos")
            total = int(cursor.fetchone()[0])
            conn.close()
            return total
        except Exception:
            return None

    # Migra automaticamente o banco legado da pasta do executavel
    # para a pasta de dados do usuario na primeira execucao.
    if not os.path.exists(DB_PATH):
        if os.path.exists(LEGACY_DB_PATH):
            try:
                shutil.copy2(LEGACY_DB_PATH, DB_PATH)
            except OSError:
                pass
        return

    # Se o banco novo existir mas estiver vazio e o legado possuir dados,
    # reaproveita o legado para evitar "sumir" com historico em atualizacoes.
    if os.path.exists(LEGACY_DB_PATH):
        docs_novo = _contar_documentos(DB_PATH)
        docs_legado = _contar_documentos(LEGACY_DB_PATH)
        if docs_novo == 0 and isinstance(docs_legado, int) and docs_legado > 0:
            try:
                shutil.copy2(LEGACY_DB_PATH, DB_PATH)
            except OSError:
                pass


def adquirir_lock_instancia():
    pid_existente = None
    if os.path.exists(LOCK_PATH):
        try:
            with open(LOCK_PATH, "r", encoding="utf-8") as f:
                pid_txt = f.read().strip()
            pid_existente = int(pid_txt)
            if pid_existente != os.getpid() and _processo_ativo(pid_existente):
                return False, pid_existente
        except (ValueError, OSError):
            pass

        try:
            os.remove(LOCK_PATH)
        except OSError:
            # Se nao houver permissao para manipular o lock, nao bloqueia inicializacao.
            return True, None

    try:
        with open(LOCK_PATH, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
    except OSError:
        # Em ambientes com pasta protegida, continua sem lock para nao impedir o uso.
        return True, None

    return True, None


def liberar_lock_instancia():
    try:
        if os.path.exists(LOCK_PATH):
            with open(LOCK_PATH, "r", encoding="utf-8") as f:
                pid_txt = f.read().strip()
            if pid_txt == str(os.getpid()):
                os.remove(LOCK_PATH)
    except OSError:
        pass


def alertar_instancia_em_execucao(pid_existente=None):
    msg = "O sistema de faturamento ja esta em execucao.\nFeche a janela atual antes de abrir novamente."
    if pid_existente:
        msg += f"\n\nPID em execucao: {pid_existente}"
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, "Sistema ja em execucao", 0x30)
    except Exception:
        try:
            messagebox.showwarning("Sistema ja em execucao", msg)
        except Exception:
            print(msg)


# ------------------------
# BANCO
# ------------------------


def obter_conexao_banco():
    return sqlite3.connect(DB_PATH)


def obter_configuracao(chave, padrao=""):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave=?", (chave,))
    row = cursor.fetchone()
    conn.close()
    if row and row[0] is not None:
        return str(row[0])
    return padrao


def salvar_configuracao(chave, valor):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO configuracoes (chave, valor)
        VALUES (?, ?)
        ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor
        """,
        (chave, "" if valor is None else str(valor)),
    )
    conn.commit()
    conn.close()


def iniciar_banco():
    conn = obter_conexao_banco()
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS documentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero INTEGER,
            numero_original TEXT,
            tipo TEXT,
            data_emissao TEXT,
            valor_inicial REAL,
            valor_final REAL,
            frete TEXT,
            status TEXT,
            competencia TEXT
        )
        """
    )

    cursor.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_documentos_numero_tipo
        ON documentos (numero, tipo)
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS configuracoes (
            chave TEXT PRIMARY KEY,
            valor TEXT
        )
        """
    )

    colunas_existentes = {linha[1] for linha in cursor.execute("PRAGMA table_info(documentos)").fetchall()}
    if "valor_inicial_original" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN valor_inicial_original REAL")
    if "valor_final_original" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN valor_final_original REAL")
    if "status_original" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN status_original TEXT")
    if "cancelado_manual" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN cancelado_manual INTEGER DEFAULT 0")
    if "competencia_manual" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN competencia_manual INTEGER DEFAULT 0")
    if "frete_manual" not in colunas_existentes:
        cursor.execute("ALTER TABLE documentos ADD COLUMN frete_manual INTEGER DEFAULT 0")

    conn.commit()
    conn.close()


# ------------------------
# VARIAVEIS
# ------------------------

relatorio_selecionado = ""
pasta_relatorios_saida = RELATORIOS_DIR

paginas_lidas = 0
docs_encontrados = 0
dashboard_update_after_id = None
dashboard_update_running = False


# ------------------------
# UTIL
# ------------------------

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


def normalizar_texto(texto):
    sem_acentos = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return sem_acentos.upper()


def competencia_por_data(data):
    return f"{MESES[data.month - 1]}/{data.year}"


def periodo_padrao_mes_atual():
    hoje = datetime.now()
    primeiro_dia = datetime(hoje.year, hoje.month, 1)
    ultimo_dia = datetime(hoje.year, hoje.month, monthrange(hoje.year, hoje.month)[1])
    return primeiro_dia, ultimo_dia


def ler_data_filtro(texto_data, nome_campo):
    try:
        return datetime.strptime(texto_data.strip(), "%d/%m/%Y")
    except ValueError as exc:
        raise ValueError(f"{nome_campo} invalida. Use o formato DD/MM/AAAA.") from exc


def centralizar_janela(janela, largura, altura):
    janela.update_idletasks()
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    x = int((screen_width / 2) - (largura / 2))
    y = int((screen_height / 2) - (altura / 2))
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def manter_interface_responsiva():
    try:
        if "app" in globals() and app.winfo_exists():
            app.update()
    except Exception:
        pass


def solicitar_atualizacao_dashboard(_event=None, delay_ms=180):
    global dashboard_update_after_id

    def _executar():
        global dashboard_update_after_id
        dashboard_update_after_id = None
        try:
            if "atualizar_dashboard" in globals():
                atualizar_dashboard()
        except Exception:
            pass

    try:
        if "app" in globals() and app.winfo_exists():
            if dashboard_update_after_id:
                try:
                    app.after_cancel(dashboard_update_after_id)
                except Exception:
                    pass
                dashboard_update_after_id = None
            if delay_ms and delay_ms > 0:
                dashboard_update_after_id = app.after(int(delay_ms), _executar)
            else:
                _executar()
        else:
            _executar()
    except Exception:
        pass


def obter_pasta_saida_relatorios():
    global pasta_relatorios_saida
    if not pasta_relatorios_saida:
        pasta_relatorios_saida = RELATORIOS_DIR
    os.makedirs(pasta_relatorios_saida, exist_ok=True)
    return pasta_relatorios_saida


def atualizar_label_pasta_saida():
    if "pasta_saida_label" in globals():
        pasta_saida_label.configure(text=f"📂 Pasta de saida: {obter_pasta_saida_relatorios()}")


def carregar_pasta_saida_relatorios():
    global pasta_relatorios_saida
    caminho_salvo = obter_configuracao("pasta_relatorios_saida", "").strip()
    if caminho_salvo:
        pasta_relatorios_saida = caminho_salvo
    obter_pasta_saida_relatorios()
    atualizar_label_pasta_saida()


def selecionar_pasta_saida_relatorios():
    global pasta_relatorios_saida
    caminho = filedialog.askdirectory(
        title="Selecionar pasta para salvar o relatorio",
        initialdir=obter_pasta_saida_relatorios(),
        mustexist=True,
    )
    if not caminho:
        return

    pasta_relatorios_saida = caminho
    salvar_configuracao("pasta_relatorios_saida", pasta_relatorios_saida)
    obter_pasta_saida_relatorios()
    atualizar_label_pasta_saida()
    messagebox.showinfo("Pasta de saida", f"Relatorios serao salvos em:\n{pasta_relatorios_saida}")


def salvar_ultimo_relatorio(caminho):
    if not caminho:
        return
    caminho_abs = os.path.abspath(caminho)
    salvar_configuracao("ultimo_relatorio_caminho", caminho_abs)
    salvar_configuracao("ultimo_relatorio_nome", os.path.basename(caminho_abs))
    salvar_configuracao("ultimo_relatorio_diretorio", os.path.dirname(caminho_abs))


def _resolver_ultimo_relatorio_salvo():
    caminho_salvo = obter_configuracao("ultimo_relatorio_caminho", "").strip()
    if caminho_salvo and os.path.exists(caminho_salvo):
        return caminho_salvo

    diretorio_salvo = obter_configuracao("ultimo_relatorio_diretorio", "").strip()
    nome_salvo = obter_configuracao("ultimo_relatorio_nome", "").strip()
    if diretorio_salvo and nome_salvo:
        candidato = os.path.join(diretorio_salvo, nome_salvo)
        if os.path.exists(candidato):
            salvar_ultimo_relatorio(candidato)
            return candidato

    return ""


def atualizar_label_relatorio():
    if "pasta_label" in globals():
        if relatorio_selecionado:
            pasta_label.configure(text=f"📄 Relatorio selecionado: {relatorio_selecionado}")
        else:
            pasta_label.configure(text="📄 Nenhum relatorio selecionado")


def definir_relatorio_selecionado(caminho, persistir=False):
    global relatorio_selecionado
    relatorio_selecionado = os.path.normpath(caminho) if caminho else ""
    if persistir and relatorio_selecionado:
        salvar_ultimo_relatorio(relatorio_selecionado)
    atualizar_label_relatorio()
    return relatorio_selecionado


def carregar_ultimo_relatorio():
    caminho = _resolver_ultimo_relatorio_salvo()
    if caminho:
        definir_relatorio_selecionado(caminho, persistir=False)


def aplicar_icone_aplicacao(janela):
    # Forca um AppUserModelID proprio para evitar agrupamento/icone do Python na barra.
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass

    try:
        icon_path = os.path.join(BASE_DIR, "logo.ico")
        if os.path.exists(icon_path):
            janela.iconbitmap(default=icon_path)
    except Exception as e:
        print(f"Aviso: nao foi possivel carregar o icone .ico - {e}")

    # Fallback para garantir icone em janelas Tk quando .ico falhar.
    try:
        if os.path.exists(LOGO_PATH):
            janela._app_icon_photo = tk.PhotoImage(file=LOGO_PATH)
            janela.iconphoto(True, janela._app_icon_photo)
    except Exception:
        pass


def abrir_seletor_data(entry_widget):
    texto_atual = entry_widget.get().strip()
    try:
        data_base = datetime.strptime(texto_atual, "%d/%m/%Y")
    except ValueError:
        data_base = datetime.now()

    picker = ctk.CTkToplevel(app)
    picker.title("Selecionar data")
    centralizar_janela(picker, 290, 280)
    picker.grab_set()

    estado = {"ano": data_base.year, "mes": data_base.month}
    semana_nomes = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]

    topo = ctk.CTkFrame(picker, fg_color="transparent")
    topo.pack(fill="x", padx=8, pady=(8, 4))

    def mudar_mes(delta):
        mes = estado["mes"] + delta
        ano = estado["ano"]
        if mes < 1:
            mes = 12
            ano -= 1
        elif mes > 12:
            mes = 1
            ano += 1
        estado["mes"] = mes
        estado["ano"] = ano
        renderizar()

    ctk.CTkButton(topo, text="<", width=30, command=lambda: mudar_mes(-1)).pack(side="left")
    titulo_mes = ctk.CTkLabel(topo, text="", font=ctk.CTkFont(weight="bold"))
    titulo_mes.pack(side="left", expand=True)
    ctk.CTkButton(topo, text=">", width=30, command=lambda: mudar_mes(1)).pack(side="right")

    corpo = ctk.CTkFrame(picker)
    corpo.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def selecionar_dia(dia):
        valor = datetime(estado["ano"], estado["mes"], dia).strftime("%d/%m/%Y")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, valor)
        solicitar_atualizacao_dashboard(delay_ms=0)
        picker.destroy()

    def renderizar():
        for child in corpo.winfo_children():
            child.destroy()

        titulo_mes.configure(text=f"{MESES[estado['mes'] - 1].capitalize()}/{estado['ano']}")

        for c, nome in enumerate(semana_nomes):
            ctk.CTkLabel(corpo, text=nome).grid(row=0, column=c, padx=2, pady=(4, 2))

        semanas = calendar.monthcalendar(estado["ano"], estado["mes"])
        for r, semana in enumerate(semanas, start=1):
            for c, dia in enumerate(semana):
                if dia == 0:
                    ctk.CTkLabel(corpo, text="", width=34).grid(row=r, column=c, padx=2, pady=2)
                else:
                    ctk.CTkButton(
                        corpo,
                        text=str(dia),
                        width=34,
                        height=28,
                        command=lambda d=dia: selecionar_dia(d),
                    ).grid(row=r, column=c, padx=2, pady=2)

    renderizar()


# ------------------------
# PORTAL NFS-E
# ------------------------


# ------------------------
# EXTRAIR CTE
# ------------------------

# ------------------------
# EXTRAIR NF
# ------------------------

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
                        numero = int(f"{data.year}{codigo:04d}")
                        valor_final = valor * 0.95
                    else:
                        numero = codigo
                        valor_final = valor

                    docs.append({
                        "numero": numero,
                        "numero_original": codigo,
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


def importar_relatorio_consolidado(caminho_pdf):
    global docs_encontrados, paginas_lidas

    docs_encontrados = 0
    paginas_lidas = 0

    vistos = {}
    competencia_ctx = None

    try:
        with fitz.open(caminho_pdf) as pdf:
            total_pag = pdf.page_count
            progress.set(0)
            for idx, pagina in enumerate(pdf, start=1):
                texto = pagina.get_text("text")
                docs, competencia_ctx = _extrair_docs_pagina_relatorio(texto, competencia_ctx)
                for doc in docs:
                    chave = (doc["tipo"], doc["numero"])
                    if chave in vistos:
                        continue
                    vistos.add(chave)
                    salvar_documento(doc)
                    docs_encontrados += 1

                paginas_lidas = idx
                progress.set(idx / total_pag)
                status_label.configure(text=f"Importando relatorio: pagina {idx}/{total_pag} | Docs:{docs_encontrados}")
                manter_interface_responsiva()
    except Exception as exc:
        messagebox.showerror("Relatorio", f"Falha ao importar relatorio.\n\n{exc}")
        return

    messagebox.showinfo("Relatorio", f"Importacao concluida. {docs_encontrados} documento(s) carregado(s).")

    atualizar_dashboard()


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

    # 1) Filial obrigatoriamente 88
    if not mapa.get("filial"):
        return False
    filial_txt = _normalizar_filial(linha.get(mapa["filial"], ""))
    if filial_txt != "88":
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
def importar_relatorio_planilha(caminho_planilha):
    global docs_encontrados, paginas_lidas

    docs_encontrados = 0
    paginas_lidas = 0

    try:
        planilhas = pd.read_excel(caminho_planilha, sheet_name=None, header=None)
    except ImportError as exc:
        messagebox.showerror(
            "Relatorio",
            "Falha ao abrir planilha. Para arquivo .xls, instale 'xlrd' ou exporte para .xlsx.\n\n" + str(exc),
        )
        return
    except Exception as exc:
        messagebox.showerror("Relatorio", f"Falha ao abrir planilha.\n\n{exc}")
        return

    if not planilhas:
        messagebox.showwarning("Relatorio", "A planilha nao possui abas para importar.")
        return

    # Limpa importacoes automaticas anteriores para evitar residuos no novo relatorio.
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        DELETE FROM documentos
        WHERE COALESCE(cancelado_manual,0)=0
          AND COALESCE(competencia_manual,0)=0
          AND UPPER(COALESCE(status,'')) NOT LIKE '%CANCELADO%'
          AND UPPER(COALESCE(status,'')) NOT LIKE '%SUBSTITUIDO%'
          AND UPPER(COALESCE(status,'')) NOT LIKE 'DOCUMENTO SUBSTITUINDO%'
        """
    )
    conn.commit()
    conn.close()

    vistos = {}

    conn = obter_conexao_banco()
    cursor = conn.cursor()

    erros = []
    total_linhas = 0
    planilhas_preparadas = {}
    for nome_aba, df in planilhas.items():
        df_prep = _preparar_dataframe_planilha(df)
        planilhas_preparadas[nome_aba] = df_prep
        if df_prep is not None and not df_prep.empty:
            total_linhas += len(df_prep.index)

    if total_linhas == 0:
        messagebox.showwarning("Relatorio", "Nao ha linhas com dados na planilha selecionada.")
        return

    processadas = 0
    total_abas = len(planilhas)

    for idx_aba, (nome_aba, df_aba) in enumerate(planilhas_preparadas.items(), start=1):
        paginas_lidas = idx_aba

        if df_aba is None or df_aba.empty:
            continue

        mapa = _mapear_colunas_planilha(df_aba)
        if mapa["faltando"]:
            preview_cols = ", ".join([str(c) for c in list(df_aba.columns)[:20]])
            erros.append(f"Aba '{nome_aba}' sem colunas obrigatorias: {', '.join(mapa['faltando'])}. Colunas detectadas: {preview_cols}")
            continue

        for _, linha in df_aba.iterrows():
            processadas += 1

            if not _linha_valida_para_importacao(linha, mapa):
                continue

            tipo = _inferir_tipo_documento_linha(linha, mapa)

            numero_txt = re.sub(r"\D", "", str(linha.get(mapa["numero"], "")))
            if not numero_txt:
                continue

            try:
                numero = int(numero_txt)
            except ValueError:
                continue

            data_raw = linha.get(mapa["data"])
            data = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
            if pd.isna(data) and mapa.get("data_ref"):
                data = pd.to_datetime(linha.get(mapa["data_ref"]), dayfirst=True, errors="coerce")
            if pd.isna(data):
                continue
            data = data.to_pydatetime()

            valor = _extrair_valor_frete_linha(linha, mapa)
            status = "OK"
            if valor is None:
                valor_inicial = 0.0
                status = "ERRO NO LANCAMENTO NO SISTEMA"
            else:
                valor_inicial = float(valor)

            if tipo == "NF":
                numero = int(f"{data.year}{int(numero_txt):04d}")
                valor_final = round(valor_inicial * 0.95, 2) if valor_inicial > 0 else 0.0
            else:
                numero = int(numero_txt)
                valor_final = valor_inicial

            frete = "FRANQUIA"
            if mapa["status"]:
                status_raw = str(linha.get(mapa["status"], "")).strip()
                if status == "OK" and status_raw and status_raw.lower() != "nan":
                    status = status_raw

            chave = (tipo, numero_txt) if tipo == "NF" else (tipo, numero)
            status_chave = "OK" if status == "OK" else "ERRO"
            if chave in vistos:
                status_existente = vistos[chave]
                # Mantem o registro mais confiavel: OK > ERRO.
                if status_existente == "OK":
                    continue
                if status_existente == "ERRO" and status_chave == "ERRO":
                    continue

            salvar_documento(
                {
                    "numero": numero,
                    "numero_original": numero_txt,
                    "tipo": tipo,
                    "data": data,
                    "valor_inicial": valor_inicial,
                    "valor_final": valor_final,
                    "frete": frete,
                    "status": status,
                    "competencia": competencia_por_data(data),
                },
                cursor
            )
            vistos[chave] = status_chave
            docs_encontrados += 1

            progress.set(processadas / total_linhas)
            status_label.configure(
                text=f"Importando planilha: linha {processadas}/{total_linhas} | Docs:{docs_encontrados}"
            )
            manter_interface_responsiva()

        progress.set(idx_aba / total_abas)
        status_label.configure(
            text=f"Importando planilha: aba {idx_aba}/{total_abas} | Docs:{docs_encontrados}"
        )
        manter_interface_responsiva()

    conn.commit()
    conn.close()

    resumo = f"Importacao concluida. {docs_encontrados} documento(s) carregado(s)."
    if erros:
        resumo += "\n\nAvisos:\n- " + "\n- ".join(erros[:5])
        if len(erros) > 5:
            resumo += f"\n- ... e mais {len(erros) - 5} aviso(s)."

    messagebox.showinfo("Relatorio", resumo)

    atualizar_dashboard()


def selecionar_relatorio():
    diretorio_inicial = (
        os.path.dirname(relatorio_selecionado)
        if relatorio_selecionado
        else obter_configuracao("ultimo_relatorio_diretorio", RELATORIOS_DIR).strip() or RELATORIOS_DIR
    )
    caminho = filedialog.askopenfilename(
        title="Selecionar relatorio consolidado",
        initialdir=diretorio_inicial,
        filetypes=[
            ("Relatorios", "*.xlsx *.xls *.pdf"),
            ("Excel", "*.xlsx *.xls"),
            ("PDF", "*.pdf"),
            ("Todos os arquivos", "*.*"),
        ],
    )
    if caminho:
        return definir_relatorio_selecionado(caminho, persistir=True)
    return ""


def importar_relatorio_ui():
    try:
        caminho = relatorio_selecionado

        # 1) Tenta relatorio salvo em memoria/configuracao.
        if not caminho:
            caminho = _resolver_ultimo_relatorio_salvo()
            if caminho:
                definir_relatorio_selecionado(caminho, persistir=True)

        # 2) Se o caminho salvo nao existir mais, tenta recuperar pelo ultimo diretorio + nome.
        if not os.path.exists(caminho):
            nome_salvo = os.path.basename(caminho) if caminho else obter_configuracao("ultimo_relatorio_nome", "").strip()
            diretorio_salvo = obter_configuracao("ultimo_relatorio_diretorio", "").strip()
            if nome_salvo and diretorio_salvo:
                candidato = os.path.join(diretorio_salvo, nome_salvo)
                if os.path.exists(candidato):
                    caminho = definir_relatorio_selecionado(candidato, persistir=True)

        # 3) Fallback: pede selecao manual somente se necessario.
        if not caminho or not os.path.exists(caminho):
            caminho = selecionar_relatorio()
            if not caminho:
                messagebox.showwarning("Relatorio", "Nenhum relatorio selecionado para importacao.")
                return

        ext = os.path.splitext(caminho)[1].lower()
        if ext == ".pdf":
            definir_relatorio_selecionado(caminho, persistir=True)
            importar_relatorio_consolidado(caminho)
            return

        if ext in {".xlsx", ".xls"}:
            definir_relatorio_selecionado(caminho, persistir=True)
            importar_relatorio_planilha(caminho)
            return

        messagebox.showerror("Relatorio", "Formato de arquivo nao suportado. Use PDF, XLSX ou XLS.")
    except Exception as exc:
        try:
            status_label.configure(text="Falha ao importar relatorio.")
        except Exception:
            pass
        messagebox.showerror("Relatorio", f"Falha ao importar relatorio.\n\n{exc}")


# ------------------------
# SINCRONIZACAO OFFLINE
# ------------------------

def _documento_possui_alteracao_manual(row):
    status_upper = str(row.get("status", "") or "").upper()
    return (
        int(row.get("cancelado_manual", 0) or 0) == 1
        or int(row.get("competencia_manual", 0) or 0) == 1
        or int(row.get("frete_manual", 0) or 0) == 1
        or bool(row.get("valor_inicial_original") is not None)
        or bool(row.get("valor_final_original") is not None)
        or bool((row.get("status_original") or "").strip())
        or "DOCUMENTO SUBSTITUIDO POR" in status_upper
        or "DOCUMENTO SUBSTITUINDO DOCUMENTO" in status_upper
    )


def _listar_documentos_alterados_para_sync():
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT
            tipo, numero, numero_original, data_emissao,
            valor_inicial, valor_final, frete, status, competencia,
            valor_inicial_original, valor_final_original, status_original,
            cancelado_manual, competencia_manual, frete_manual
        FROM documentos
        ORDER BY id ASC
        """
    )
    docs = []
    for row in cursor.fetchall():
        item = dict(row)
        if _documento_possui_alteracao_manual(item):
            item["tipo"] = str(item.get("tipo", "")).upper().strip()
            item["numero"] = int(item.get("numero") or 0)
            item["numero_original"] = str(item.get("numero_original", "") or "").strip()
            item["cancelado_manual"] = int(item.get("cancelado_manual") or 0)
            item["competencia_manual"] = int(item.get("competencia_manual") or 0)
            item["frete_manual"] = int(item.get("frete_manual") or 0)
            docs.append(item)
    conn.close()
    return docs


def exportar_configuracoes_json(caminho_arquivo):
    documentos = _listar_documentos_alterados_para_sync()
    payload = {
        "metadata": {
            "schema_version": SYNC_CONFIG_SCHEMA_VERSION,
            "exportado_em": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "origem": {
                "host": socket.gethostname(),
                "usuario": getpass.getuser(),
                "app_data_dir": APP_DATA_DIR,
            },
            "total_documentos": len(documentos),
        },
        "documentos": documentos,
    }

    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return len(documentos)


def _coletar_numero_original_para_match(numero_original, numero):
    numero_original_txt = str(numero_original or "").strip()
    if numero_original_txt:
        numero_original_txt = re.sub(r"\D", "", numero_original_txt) or numero_original_txt
    if not numero_original_txt:
        numero_original_txt = str(numero)
    try:
        numero_original_int = int(re.sub(r"\D", "", numero_original_txt))
    except ValueError:
        numero_original_int = None
    return numero_original_txt, numero_original_int


def _buscar_documento_existente_sync(cursor, tipo, numero, numero_original):
    cursor.execute(
        "SELECT * FROM documentos WHERE tipo=? AND numero=? ORDER BY id DESC LIMIT 1",
        (tipo, numero),
    )
    row = cursor.fetchone()
    if row:
        return dict(row)

    if tipo == "NF":
        numero_original_txt, numero_original_int = _coletar_numero_original_para_match(numero_original, numero)
        if numero_original_int is not None:
            cursor.execute(
                """
                SELECT * FROM documentos
                WHERE tipo='NF'
                  AND (
                      numero_original=?
                      OR CAST(numero_original AS INTEGER)=?
                  )
                ORDER BY id DESC
                LIMIT 1
                """,
                (numero_original_txt, numero_original_int),
            )
        else:
            cursor.execute(
                """
                SELECT * FROM documentos
                WHERE tipo='NF' AND numero_original=?
                ORDER BY id DESC
                LIMIT 1
                """,
                (numero_original_txt,),
            )
        row = cursor.fetchone()
        if row:
            return dict(row)

    return None


def _to_float(valor, padrao=0.0):
    try:
        if valor is None or str(valor).strip() == "":
            return float(padrao)
        return float(valor)
    except (TypeError, ValueError):
        return float(padrao)


def _to_optional_float(valor):
    if valor is None:
        return None
    if isinstance(valor, str) and not valor.strip():
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _to_manual_flag(valor, padrao=0):
    try:
        return 1 if int(valor) == 1 else 0
    except (TypeError, ValueError):
        return int(padrao)


def importar_configuracoes_json(caminho_arquivo):
    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        payload = json.load(f)

    if isinstance(payload, list):
        documentos = payload
    elif isinstance(payload, dict):
        documentos = payload.get("documentos", [])
    else:
        raise ValueError("Formato invalido de arquivo de configuracao.")

    if not isinstance(documentos, list):
        raise ValueError("O arquivo de configuracao nao contem lista de documentos.")

    resumo = {"inseridos": 0, "atualizados": 0, "ignorados": 0, "erros": []}
    vistos = set()

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    for idx, item in enumerate(documentos, start=1):
        try:
            if not isinstance(item, dict):
                resumo["ignorados"] += 1
                continue

            tipo = str(item.get("tipo", "")).upper().strip()
            if tipo not in {"NF", "CTE"}:
                resumo["ignorados"] += 1
                continue

            try:
                numero = int(item.get("numero"))
            except (TypeError, ValueError):
                resumo["ignorados"] += 1
                continue

            numero_original_txt, _ = _coletar_numero_original_para_match(item.get("numero_original"), numero)
            chave = (tipo, numero, numero_original_txt)
            if chave in vistos:
                resumo["ignorados"] += 1
                continue
            vistos.add(chave)

            existente = _buscar_documento_existente_sync(cursor, tipo, numero, numero_original_txt)
            numero_chave = int(existente["numero"]) if existente else numero

            data_emissao = str(item.get("data_emissao") or (existente.get("data_emissao") if existente else "")).strip()
            if not data_emissao:
                data_emissao = datetime.now().strftime("%d/%m/%Y")

            competencia = str(item.get("competencia") or (existente.get("competencia") if existente else "")).strip().lower()
            if not competencia:
                competencia = competencia_por_data(datetime.now())

            frete = str(item.get("frete") or (existente.get("frete") if existente else "FRANQUIA")).strip().upper()
            if not frete:
                frete = "FRANQUIA"

            status = str(item.get("status") or (existente.get("status") if existente else "OK")).strip()
            if not status:
                status = "OK"

            valor_inicial = _to_float(item.get("valor_inicial"), existente.get("valor_inicial", 0.0) if existente else 0.0)
            valor_final = _to_float(item.get("valor_final"), existente.get("valor_final", 0.0) if existente else 0.0)

            valor_inicial_original = (
                _to_optional_float(item.get("valor_inicial_original"))
                if "valor_inicial_original" in item
                else _to_optional_float(existente.get("valor_inicial_original")) if existente else None
            )
            valor_final_original = (
                _to_optional_float(item.get("valor_final_original"))
                if "valor_final_original" in item
                else _to_optional_float(existente.get("valor_final_original")) if existente else None
            )

            status_original = item.get("status_original")
            if status_original is None and existente:
                status_original = existente.get("status_original")
            status_original = str(status_original).strip() if status_original not in (None, "") else None

            cancelado_manual = _to_manual_flag(
                item.get("cancelado_manual"),
                existente.get("cancelado_manual", 0) if existente else 0,
            )
            competencia_manual = _to_manual_flag(
                item.get("competencia_manual"),
                existente.get("competencia_manual", 0) if existente else 0,
            )
            frete_manual = _to_manual_flag(
                item.get("frete_manual"),
                existente.get("frete_manual", 0) if existente else (0 if frete == "FRANQUIA" else 1),
            )

            cursor.execute(
                """
                INSERT INTO documentos
                (
                    numero,numero_original,tipo,data_emissao,valor_inicial,valor_final,frete,status,competencia,
                    valor_inicial_original,valor_final_original,status_original,cancelado_manual,competencia_manual,frete_manual
                )
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(numero,tipo) DO UPDATE SET
                    numero_original=excluded.numero_original,
                    data_emissao=excluded.data_emissao,
                    valor_inicial=excluded.valor_inicial,
                    valor_final=excluded.valor_final,
                    frete=excluded.frete,
                    status=excluded.status,
                    competencia=excluded.competencia,
                    valor_inicial_original=excluded.valor_inicial_original,
                    valor_final_original=excluded.valor_final_original,
                    status_original=excluded.status_original,
                    cancelado_manual=excluded.cancelado_manual,
                    competencia_manual=excluded.competencia_manual,
                    frete_manual=excluded.frete_manual
                """,
                (
                    numero_chave,
                    numero_original_txt,
                    tipo,
                    data_emissao,
                    valor_inicial,
                    valor_final,
                    frete,
                    status,
                    competencia,
                    valor_inicial_original,
                    valor_final_original,
                    status_original,
                    cancelado_manual,
                    competencia_manual,
                    frete_manual,
                ),
            )

            if existente:
                resumo["atualizados"] += 1
            else:
                resumo["inseridos"] += 1
        except Exception as exc:
            resumo["erros"].append(f"Linha {idx}: {exc}")

    conn.commit()
    conn.close()
    return resumo


def exportar_configuracoes_ui():
    try:
        docs = _listar_documentos_alterados_para_sync()
        if not docs:
            messagebox.showwarning("Configuracoes", "Nao ha alteracoes manuais para exportar.")
            return

        diretorio_inicial = obter_configuracao("ultimo_sync_diretorio", obter_pasta_saida_relatorios()).strip() or obter_pasta_saida_relatorios()
        nome_padrao = f"configuracoes_faturamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        caminho = filedialog.asksaveasfilename(
            title="Exportar configuracoes",
            defaultextension=".json",
            initialdir=diretorio_inicial,
            initialfile=nome_padrao,
            filetypes=[("Arquivo JSON", "*.json"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        total = exportar_configuracoes_json(caminho)
        salvar_configuracao("ultimo_sync_diretorio", os.path.dirname(caminho))
        messagebox.showinfo("Configuracoes", f"Exportacao concluida.\n\nDocumentos exportados: {total}\nArquivo: {caminho}")
    except Exception as exc:
        messagebox.showerror("Configuracoes", f"Falha ao exportar configuracoes.\n\n{exc}")


def importar_configuracoes_ui():
    try:
        diretorio_inicial = obter_configuracao("ultimo_sync_diretorio", obter_pasta_saida_relatorios()).strip() or obter_pasta_saida_relatorios()
        caminho = filedialog.askopenfilename(
            title="Importar configuracoes",
            initialdir=diretorio_inicial,
            filetypes=[("Arquivo JSON", "*.json"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        salvar_configuracao("ultimo_sync_diretorio", os.path.dirname(caminho))
        resumo = importar_configuracoes_json(caminho)
        atualizar_dashboard()

        mensagem = (
            "Importacao concluida.\n\n"
            f"Inseridos: {resumo['inseridos']}\n"
            f"Atualizados: {resumo['atualizados']}\n"
            f"Ignorados: {resumo['ignorados']}\n"
            f"Erros: {len(resumo['erros'])}"
        )
        if resumo["erros"]:
            mensagem += "\n\nPrimeiros erros:\n- " + "\n- ".join(resumo["erros"][:5])
            if len(resumo["erros"]) > 5:
                mensagem += f"\n- ... e mais {len(resumo['erros']) - 5} erro(s)."

        messagebox.showinfo("Configuracoes", mensagem)
    except Exception as exc:
        messagebox.showerror("Configuracoes", f"Falha ao importar configuracoes.\n\n{exc}")


# ------------------------
# SALVAR
# ------------------------

def salvar_documento(doc, cursor=None):
    conn_externo = cursor is not None
    if not conn_externo:
        conn = obter_conexao_banco()
        cursor = conn.cursor()

    try:
        tipo_doc = str(doc.get("tipo", "")).upper()
        if tipo_doc == "NF":
            numero_original = str(doc.get("numero_original", "")).strip()
            try:
                numero_original_int = int(re.sub(r"\D", "", numero_original))
            except ValueError:
                numero_original_int = None

            if numero_original:
                if numero_original_int is not None:
                    cursor.execute(
                        """
                        DELETE FROM documentos
                        WHERE tipo='NF' AND numero<>? AND (numero_original=? OR numero=?)
                        """,
                        (doc["numero"], numero_original, numero_original_int),
                    )
                else:
                    cursor.execute(
                        """
                        DELETE FROM documentos
                        WHERE tipo='NF' AND numero<>? AND numero_original=?
                        """,
                        (doc["numero"], numero_original),
                    )

        cursor.execute(
            """
            INSERT INTO documentos
            (
                numero,numero_original,tipo,data_emissao,valor_inicial,valor_final,frete,status,competencia,
                valor_inicial_original,valor_final_original,status_original,cancelado_manual,competencia_manual,frete_manual
            )
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(numero,tipo) DO UPDATE SET
                numero_original=excluded.numero_original,
                data_emissao=excluded.data_emissao,
                frete=CASE
                    WHEN documentos.frete_manual=1 THEN documentos.frete
                    ELSE excluded.frete
                END,
                competencia=CASE
                    WHEN documentos.competencia_manual=1 THEN documentos.competencia
                    ELSE excluded.competencia
                END,
                valor_inicial=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.valor_inicial
                    ELSE excluded.valor_inicial
                END,
                valor_final=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.valor_final
                    ELSE excluded.valor_final
                END,
                status=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.status
                    ELSE excluded.status
                END,
                valor_inicial_original=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.valor_inicial_original
                    ELSE NULL
                END,
                valor_final_original=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.valor_final_original
                    ELSE NULL
                END,
                status_original=CASE
                    WHEN documentos.cancelado_manual=1 THEN documentos.status_original
                    ELSE NULL
                END
            """,
            (
                doc["numero"],
                doc["numero_original"],
                doc["tipo"],
                doc["data"].strftime("%d/%m/%Y"),
                doc["valor_inicial"],
                doc["valor_final"],
                doc["frete"],
                doc["status"],
                doc.get("competencia", competencia_por_data(doc["data"])),
                None,
                None,
                None,
                0,
                0,
                int(doc.get("frete_manual", 0) or 0),
            ),
        )
        if not conn_externo:
            conn.commit()
    finally:
        if not conn_externo:
            conn.close()


def alterar_competencia_documento(tipo, numero, mes_competencia, ano_competencia):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    competencia = f"{mes_competencia}/{ano_competencia}"
    cursor.execute(
        "UPDATE documentos SET competencia=?, competencia_manual=1 WHERE tipo=? AND numero=?",
        (competencia, tipo, numero),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()

    atualizar_dashboard()

    return alterados


def declarar_documento_frete(tipo, numero, novo_frete):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    tipo = str(tipo).upper()
    frete_manual = 0 if str(novo_frete).upper() == "FRANQUIA" else 1

    if tipo == "NF":
        numero_txt = str(numero).strip()
        try:
            numero_num = int(numero_txt)
        except ValueError:
            numero_num = numero
        cursor.execute(
            """
            UPDATE documentos
            SET frete=?, frete_manual=?
            WHERE tipo='NF'
              AND (
                  numero=?
                  OR numero_original=?
                  OR CAST(numero_original AS INTEGER)=?
              )
            """,
            (novo_frete, frete_manual, numero_num, numero_txt, numero_num),
        )
    else:
        cursor.execute(
            "UPDATE documentos SET frete=?, frete_manual=? WHERE tipo=? AND numero=?",
            (novo_frete, frete_manual, tipo, numero),
        )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()

    atualizar_dashboard()

    return alterados


def registrar_substituicao(tipo_antigo, numero_antigo, tipo_novo, numero_novo):
    conn = obter_conexao_banco()
    cursor = conn.cursor()

    status_novo = f"DOCUMENTO SUBSTITUINDO DOCUMENTO {numero_antigo} {tipo_antigo}"
    status_antigo = f"DOCUMENTO SUBSTITUIDO POR {numero_novo} {tipo_novo}"

    cursor.execute(
        """
        UPDATE documentos
        SET
            status_original=COALESCE(status_original, status),
            status=?
        WHERE tipo=? AND numero=?
        """,
        (status_novo, tipo_novo, numero_novo),
    )
    novo_alterado = cursor.rowcount

    cursor.execute(
        """
        UPDATE documentos
        SET
            valor_inicial_original=COALESCE(valor_inicial_original, valor_inicial),
            valor_final_original=COALESCE(valor_final_original, valor_final),
            status_original=COALESCE(status_original, status),
            valor_inicial=0,
            valor_final=0,
            status=?
        WHERE tipo=? AND numero=?
        """,
        (status_antigo, tipo_antigo, numero_antigo),
    )
    antigo_alterado = cursor.rowcount

    conn.commit()
    conn.close()

    atualizar_dashboard()

    return novo_alterado, antigo_alterado


def desfazer_substituicao(tipo_antigo, numero_antigo, tipo_novo, numero_novo):
    conn = obter_conexao_banco()
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE documentos
        SET
            valor_inicial=COALESCE(valor_inicial_original, valor_inicial),
            valor_final=COALESCE(valor_final_original, valor_final),
            status=COALESCE(status_original, 'OK'),
            valor_inicial_original=NULL,
            valor_final_original=NULL,
            status_original=NULL
        WHERE tipo=? AND numero=? AND UPPER(status) LIKE 'DOCUMENTO SUBSTITUIDO POR%'
        """,
        (tipo_antigo, numero_antigo),
    )
    antigo_restaurado = cursor.rowcount

    cursor.execute(
        """
        UPDATE documentos
        SET
            status=COALESCE(status_original, 'OK'),
            status_original=NULL
        WHERE tipo=? AND numero=? AND UPPER(status) LIKE 'DOCUMENTO SUBSTITUINDO DOCUMENTO%'
        """,
        (tipo_novo, numero_novo),
    )
    novo_restaurado = cursor.rowcount

    conn.commit()
    conn.close()

    atualizar_dashboard()

    return antigo_restaurado, novo_restaurado


def cancelar_documento(tipo, numero):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE documentos
        SET
            valor_inicial_original=COALESCE(valor_inicial_original, valor_inicial),
            valor_final_original=COALESCE(valor_final_original, valor_final),
            status_original=COALESCE(status_original, status),
            valor_inicial=0,
            valor_final=0,
            status='CANCELADO MANUALMENTE',
            cancelado_manual=1
        WHERE tipo=? AND numero=?
        """,
        (tipo, numero),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()

    atualizar_dashboard()

    return alterados


def desfazer_cancelamento_documento(tipo, numero):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE documentos
        SET
            valor_inicial=COALESCE(valor_inicial_original, valor_inicial),
            valor_final=COALESCE(valor_final_original, valor_final),
            status=COALESCE(status_original, 'OK'),
            valor_inicial_original=NULL,
            valor_final_original=NULL,
            status_original=NULL,
            cancelado_manual=0
        WHERE tipo=? AND numero=? AND UPPER(status)='CANCELADO MANUALMENTE'
        """,
        (tipo, numero),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()

    atualizar_dashboard()

    return alterados


def abrir_relatorio():
    nome = os.path.join(obter_pasta_saida_relatorios(), "Faturamento_AC.xlsx")
    if os.path.exists(nome):
        try:
            os.startfile(nome)
        except OSError as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir o relatorio: {e}")
    else:
        messagebox.showwarning("Relatorio", "Relatorio nao encontrado. Gere o relatorio primeiro.")


# ------------------------
# ATUALIZAR
# ------------------------

# ------------------------
# STATUS
# ------------------------

# ------------------------
# EXCEL
# ------------------------

def gerar_excel():
    conn = obter_conexao_banco()
    df = pd.read_sql_query("SELECT * FROM documentos", conn)
    conn.close()

    if df.empty:
        messagebox.showwarning("Excel", "Nenhum documento encontrado no banco.")
        return

    df["data_emissao"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao"])

    if df.empty:
        messagebox.showwarning("Excel", "Nao ha datas validas para gerar o relatorio.")
        return

    df["numero"] = pd.to_numeric(df["numero"], errors="coerce")
    df = df.dropna(subset=["numero"])

    if df.empty:
        messagebox.showwarning("Excel", "Nao ha numeros de documento validos para gerar o relatorio.")
        return

    df["numero"] = df["numero"].astype(int)

    try:
        data_inicial = ler_data_filtro(data_inicio_entry.get(), "Data inicial")
        data_final = ler_data_filtro(data_fim_entry.get(), "Data final")
    except ValueError as erro_data:
        messagebox.showwarning("Excel", str(erro_data))
        return

    if data_inicial > data_final:
        messagebox.showwarning("Excel", "A data inicial nao pode ser maior que a data final.")
        return

    # Converte competencia para data (primeiro dia do mês) para filtro
    def competencia_para_data(comp_str):
        try:
            partes = comp_str.lower().split("/")
            if len(partes) == 2:
                mes_nome = partes[0].strip()
                ano_str = partes[1].strip()
                ano = int(ano_str)
                mes_idx = MESES.index(mes_nome) + 1
                return datetime(ano, mes_idx, 1)
        except:
            pass
        return None

    df["data_competencia"] = df["competencia"].apply(competencia_para_data)
    df = df.dropna(subset=["data_competencia"])

    # Filtra por competência (período selecionado)
    df = df[(df["data_competencia"] >= data_inicial) & (df["data_competencia"] <= data_final)].copy()

    # Evita duplicidade de NF quando existir numero antigo (ex.: 20260092) e numero real (92).
    df["numero_original_num"] = pd.to_numeric(df.get("numero_original"), errors="coerce")
    df["chave_documento"] = df.apply(
        lambda r: (
            f"NF:{int(r['numero_original_num'])}"
            if str(r["tipo"]).upper() == "NF" and pd.notna(r["numero_original_num"])
            else f"{str(r['tipo']).upper()}:{int(r['numero'])}"
        ),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")

    if df.empty:
        messagebox.showwarning(
            "Excel",
            f"Nao ha documentos para o periodo selecionado ({data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}).",
        )
        return

    df = df.sort_values(["data_emissao", "numero"], ascending=[True, True])
    df["Concat"] = df["numero"].astype(str) + " " + df["tipo"]
    # Usa a data de competencia ja validada para evitar coluna "Mes Referencia" vazia no Excel.
    df["competencia_excel"] = df["data_competencia"]

    # Remove arquivo anterior se existir
    pasta_saida = obter_pasta_saida_relatorios()
    nome = os.path.join(pasta_saida, "Faturamento_AC.xlsx")
    for antigo in glob.glob(os.path.join(pasta_saida, "Faturamento_AC*.xlsx")):
        try:
            if os.path.exists(antigo):
                os.remove(antigo)
        except (OSError, PermissionError) as e:
            try:
                os.rename(antigo, antigo + ".bak")
            except:
                pass

    colunas_base = [
        "data_emissao",
        "competencia_excel",
        "numero",
        "numero_original_num",
        "tipo",
        "frete",
        "valor_inicial",
        "valor_final",
        "status",
    ]

    df_base = df[colunas_base].copy()

    def montar_df_relatorio(df_in, usar_numero_real_nf=False):
        dados = df_in.copy()
        if usar_numero_real_nf:
            dados["numero_doc"] = dados.apply(
                lambda r: int(r["numero_original_num"])
                if str(r["tipo"]).upper() == "NF" and pd.notna(r["numero_original_num"])
                else int(r["numero"]) if pd.notna(r["numero"]) else None,
                axis=1,
            )
        else:
            dados["numero_doc"] = dados["numero"]

        dados["numero_doc"] = pd.to_numeric(dados["numero_doc"], errors="coerce")
        dados = dados.dropna(subset=["numero_doc"]).copy()
        dados["numero_doc"] = dados["numero_doc"].astype(int)
        dados["Concat"] = dados["numero_doc"].astype(str) + " " + dados["tipo"].astype(str)

        dados = dados[
            [
                "data_emissao",
                "competencia_excel",
                "numero_doc",
                "tipo",
                "Concat",
                "frete",
                "valor_inicial",
                "valor_final",
                "status",
            ]
        ]

        dados.columns = [
            "Data Emissao",
            "Mes Referencia",
            "Numero Doc",
            "Tipo Doc",
            "Concat",
            "Frete",
            "Valor Inicial",
            "Valor Final",
            "Status",
        ]
        return dados

    df_relatorio_1 = montar_df_relatorio(df_base, usar_numero_real_nf=False)
    df_relatorio_2 = montar_df_relatorio(df_base, usar_numero_real_nf=True)

    try:
        with pd.ExcelWriter(nome, engine="openpyxl") as writer:
            nome_aba_1 = "Faturamento AC"
            nome_aba_2 = "Faturamento AC 2"
            df_relatorio_1.to_excel(writer, index=False, startcol=1, sheet_name=nome_aba_1)
            df_relatorio_2.to_excel(writer, index=False, startcol=1, sheet_name=nome_aba_2)

            def formatar_aba(nome_aba, df_aba):
                ws = writer.sheets[nome_aba]
                ws.sheet_view.showGridLines = False
                ws.freeze_panes = "B2"

                ultima_linha = len(df_aba) + 1
                linha_inicial = 1
                col_inicial = 2  # B
                col_final = 10   # J

                cor_cabecalho = PatternFill("solid", fgColor="1F4E78")
                fonte_cabecalho = Font(color="FFFFFF", bold=True)
                cor_cancelado = PatternFill("solid", fgColor="F8D7DA")
                cor_b_dados = PatternFill("solid", fgColor="E6E6E6")
                cor_f_dados = PatternFill("solid", fgColor="FFF2CC")
                borda_fina = Border(
                    left=Side(style="thin", color="D9D9D9"),
                    right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"),
                    bottom=Side(style="thin", color="D9D9D9"),
                )

                for col in range(col_inicial, col_final + 1):
                    celula = ws.cell(row=1, column=col)
                    celula.fill = cor_cabecalho
                    celula.font = fonte_cabecalho
                    celula.alignment = Alignment(horizontal="center", vertical="center")

                for row in range(linha_inicial, ultima_linha + 1):
                    for col in range(col_inicial, col_final + 1):
                        ws.cell(row=row, column=col).border = borda_fina

                for row in range(2, ultima_linha + 1):
                    for col in range(col_inicial, col_final + 1):
                        ws.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="center")

                    ws[f"B{row}"].fill = cor_b_dados
                    ws[f"F{row}"].fill = cor_f_dados

                    status_valor = str(ws[f"J{row}"].value or "").upper()
                    if ("CANCELADO" in status_valor) or ("SUBSTITUIDO" in status_valor):
                        for col in range(col_inicial, col_final + 1):
                            ws.cell(row=row, column=col).fill = cor_cancelado

                    ws[f"B{row}"].number_format = "DD/MM/YYYY"
                    ws[f"C{row}"].number_format = '[$-pt-BR]mmmm/yyyy'
                    ws[f"D{row}"].number_format = "0"
                    ws[f"H{row}"].number_format = "R$ #,##0.00"
                    ws[f"I{row}"].number_format = "R$ #,##0.00"

                for col in range(col_inicial, col_final + 1):
                    letra_coluna = ws.cell(row=1, column=col).column_letter
                    maior = 0
                    for row in range(1, ultima_linha + 1):
                        valor = ws.cell(row=row, column=col).value
                        tamanho = len(str(valor)) if valor is not None else 0
                        if tamanho > maior:
                            maior = tamanho
                    ws.column_dimensions[letra_coluna].width = min(maior + 2, 45)

            formatar_aba(nome_aba_1, df_relatorio_1)
            formatar_aba(nome_aba_2, df_relatorio_2)

        messagebox.showinfo(
            "Excel",
            f"Relatorio gerado com {len(df_relatorio_1)} documento(s) no periodo {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}\n\nAbas: Faturamento AC e Faturamento AC 2\nArquivo: {nome}",
        )
    except Exception as e:
        messagebox.showerror("Excel", f"Erro ao gerar o relatorio: {e}")


# ------------------------
# INTERFACE
# ------------------------

preparar_arquivos_aplicacao()

lock_ok, pid_existente = adquirir_lock_instancia()
if not lock_ok:
    alertar_instancia_em_execucao(pid_existente)
    raise SystemExit(0)

atexit.register(liberar_lock_instancia)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.screens = {}
        self.current_screen = ""
        self.screen_host = None
        self.title("Sistema de Faturamento - Horizonte Logistica")
        self.resizable(True, True)
        aplicar_icone_aplicacao(self)

    def registrar_tela(self, nome_tela, frame_tela):
        self.screens[nome_tela] = frame_tela

    def mostrar_tela(self, nome_tela):
        if nome_tela not in self.screens:
            return
        for nome, frame in self.screens.items():
            if frame.winfo_exists():
                if nome == nome_tela:
                    frame.pack(fill="both", expand=True, padx=2, pady=2)
                else:
                    frame.pack_forget()
        self.current_screen = nome_tela
        _configurar_nav_tela_ativa(nome_tela)


app = App()

largura_tela = app.winfo_screenwidth()
altura_tela = app.winfo_screenheight()

largura = max(860, min(1120, largura_tela - 80))
altura = max(650, min(860, altura_tela - 90))
largura = min(largura, largura_tela - 10)
altura = min(altura, altura_tela - 10)

centralizar_janela(app, largura, altura)


def _encerrar_aplicacao():
    liberar_lock_instancia()
    app.destroy()


app.protocol("WM_DELETE_WINDOW", _encerrar_aplicacao)

iniciar_banco()
carregar_pasta_saida_relatorios()
carregar_ultimo_relatorio()

primeiro_dia_padrao, ultimo_dia_padrao = periodo_padrao_mes_atual()

APP_THEMES = {
    "light": {
        "app_bg": "#EDF3FA",
        "header_bg": "#E8F1FB",
        "surface": "#FFFFFF",
        "surface_alt": "#F2F7FD",
        "border": "#C9D8EA",
        "divider": "#CFDBEA",
        "text_primary": "#1A2A3B",
        "text_secondary": "#607286",
        "accent": "#2F80D0",
        "accent_hover": "#256FBD",
        "on_accent": "#FFFFFF",
        "success_bg": "#E6F4EB",
        "success_text": "#246B43",
        "danger_bg": "#FCEBED",
        "danger_text": "#A33B4B",
        "scroll_btn": "#BFD1E4",
        "scroll_btn_hover": "#A8BED7",
        "progress_bg": "#D7E4F3",
        "cta_border": "#2A70B8",
        "cta_border_hover": "#205D9E",
        "cta_press": "#1F5FA6",
        "tab_hover": "#E2EDF9",
        "tab_press": "#D4E4F7",
        "tab_active_hover": "#2A77C6",
        "tab_active_press": "#215EA5",
        "metric_initial_bg": "#E1ECFA",
        "metric_initial_title": "#245A94",
        "metric_initial_value": "#113F72",
        "metric_final_bg": "#E2F2E8",
        "metric_final_title": "#2E7950",
        "metric_final_value": "#1D5F3C",
        "metric_nf_bg": "#E5F5EC",
        "metric_nf_title": "#2E7950",
        "metric_nf_value": "#1F6340",
        "metric_docs_bg": "#E5EFFA",
        "metric_docs_title": "#245A94",
        "metric_docs_value": "#164A7E",
        "metric_cte_bg": "#EDF2F9",
        "metric_cte_title": "#415C79",
        "metric_cte_value": "#2E445F",
        "metric_cancelados_bg": "#EDF2F9",
        "metric_cancelados_title": "#415C79",
        "metric_cancelados_value": "#2E445F",
        "metric_default_bg": "#EDF2F9",
        "metric_default_title": "#415C79",
        "metric_default_value": "#2E445F",
        "chart_bg": "#FFFFFF",
        "chart_plot_bg": "#F6F9FD",
        "chart_grid": "#C8D6E7",
        "chart_axis": "#4C637D",
        "chart_bar_primary": "#2F80D0",
        "chart_bar_secondary": "#78AEDD",
        "chart_line": "#2A6EAF",
        "chart_cancelados": "#D56A7A",
    },
    "dark": {
        "app_bg": "#0C131C",
        "header_bg": "#121E2C",
        "surface": "#111B27",
        "surface_alt": "#182433",
        "border": "#2A3D53",
        "divider": "#2A3D53",
        "text_primary": "#E7EEF8",
        "text_secondary": "#9BB0C8",
        "accent": "#3E96E0",
        "accent_hover": "#3183CC",
        "on_accent": "#F7FBFF",
        "success_bg": "#173A2A",
        "success_text": "#7BD2A6",
        "danger_bg": "#3A2026",
        "danger_text": "#F39AA8",
        "scroll_btn": "#384E66",
        "scroll_btn_hover": "#4A6380",
        "progress_bg": "#27384D",
        "cta_border": "#2B76BF",
        "cta_border_hover": "#3C86D0",
        "cta_press": "#266EB3",
        "tab_hover": "#24364B",
        "tab_press": "#2B4260",
        "tab_active_hover": "#3A8ED6",
        "tab_active_press": "#2D76BD",
        "metric_initial_bg": "#163150",
        "metric_initial_title": "#9CC5EC",
        "metric_initial_value": "#E6F2FF",
        "metric_final_bg": "#163829",
        "metric_final_title": "#9ED8B6",
        "metric_final_value": "#DDF6EA",
        "metric_nf_bg": "#193D2D",
        "metric_nf_title": "#9ED8B6",
        "metric_nf_value": "#D8F2E6",
        "metric_docs_bg": "#193552",
        "metric_docs_title": "#9CC5EC",
        "metric_docs_value": "#E6F2FF",
        "metric_cte_bg": "#1A2E43",
        "metric_cte_title": "#A9BED6",
        "metric_cte_value": "#E0EAF6",
        "metric_cancelados_bg": "#1A2E43",
        "metric_cancelados_title": "#A9BED6",
        "metric_cancelados_value": "#E0EAF6",
        "metric_default_bg": "#1A2E43",
        "metric_default_title": "#A9BED6",
        "metric_default_value": "#E0EAF6",
        "chart_bg": "#111B27",
        "chart_plot_bg": "#182433",
        "chart_grid": "#38506A",
        "chart_axis": "#A6BCD4",
        "chart_bar_primary": "#4A9CE4",
        "chart_bar_secondary": "#77B5EB",
        "chart_line": "#A4D0F4",
        "chart_cancelados": "#E07D8D",
    },
}

tema_salvo = obter_configuracao("tema_interface", "light").strip().lower()
current_theme_mode = tema_salvo if tema_salvo in APP_THEMES else "light"
UI_THEME = dict(APP_THEMES[current_theme_mode])
ctk.set_appearance_mode("dark" if current_theme_mode == "dark" else "light")


def _gerar_tab_styles():
    return {
        "normal_fg": UI_THEME["surface_alt"],
        "normal_text": UI_THEME["text_primary"],
        "hover_fg": UI_THEME["tab_hover"],
        "press_fg": UI_THEME["tab_press"],
        "active_fg": UI_THEME["accent"],
        "active_hover_fg": UI_THEME["tab_active_hover"],
        "active_press_fg": UI_THEME["tab_active_press"],
        "active_text": UI_THEME["on_accent"],
    }


TAB_STYLES = _gerar_tab_styles()
app.configure(fg_color=UI_THEME["app_bg"])


def abrir_dialogo_alterar_competencia():
    dialog = ctk.CTkToplevel(app)
    dialog.title("Alterar competencia de documento")
    centralizar_janela(dialog, 620, 270)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Tipo de documento").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Numero do documento").grid(row=0, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=260)
    tipo_combo.set("NF")
    tipo_combo.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 10))

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 20260089 ou 2390")
    numero_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 10))

    ctk.CTkLabel(form, text="Novo mes de competencia").grid(row=2, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Ano da competencia").grid(row=2, column=1, sticky="w", padx=6, pady=(0, 4))

    mes_combo = ctk.CTkComboBox(form, values=MESES, width=260)
    mes_combo.set(MESES[datetime.now().month - 1])
    mes_combo.grid(row=3, column=0, sticky="ew", padx=6)

    anos_competencia = [str(ano) for ano in range(datetime.now().year - 3, datetime.now().year + 5)]
    ano_combo = ctk.CTkComboBox(form, values=anos_competencia, width=260)
    ano_combo.set(str(datetime.now().year))
    ano_combo.grid(row=3, column=1, sticky="ew", padx=6)

    def salvar_alteracao():
        tipo = tipo_combo.get().strip().upper()
        mes_novo = mes_combo.get().strip().lower()
        ano_novo = ano_combo.get().strip()
        numero_texto = numero_entry.get().strip()

        if tipo not in {"NF", "CTE"}:
            messagebox.showwarning("Aviso", "Tipo de documento invalido.")
            return

        if mes_novo not in MESES:
            messagebox.showwarning("Aviso", "Selecione um mes valido.")
            return

        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um numero de documento valido.")
            return

        if not ano_novo.isdigit():
            messagebox.showwarning("Aviso", "Informe um ano valido.")
            return

        numero = int(numero_texto)
        alterados = alterar_competencia_documento(tipo, numero, mes_novo, int(ano_novo))
        if alterados == 0:
            messagebox.showwarning("Aviso", "Documento nao encontrado para alterar competencia.")
            return

        messagebox.showinfo("Sucesso", "Competencia atualizada com sucesso.")
        dialog.destroy()

    ctk.CTkButton(form, text="Salvar alteracao", command=salvar_alteracao, width=200).grid(row=4, column=0, columnspan=2, pady=(14, 0))


def _abrir_dialogo_declarar_frete(rotulo_frete):
    dialog = ctk.CTkToplevel(app)
    dialog.title(rotulo_frete)
    centralizar_janela(dialog, 620, 220)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Tipo de documento").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Numero do documento").grid(row=0, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=260)
    tipo_combo.set("NF")
    tipo_combo.grid(row=1, column=0, sticky="ew", padx=6)

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 20260092 ou 2390")
    numero_entry.grid(row=1, column=1, sticky="ew", padx=6)

    desfazer_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(
        form,
        text=f"Des{rotulo_frete.lower()}",
        variable=desfazer_var,
        onvalue=True,
        offvalue=False,
    ).grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def confirmar():
        tipo = tipo_combo.get().strip().upper()
        numero_texto = numero_entry.get().strip()

        if tipo not in {"NF", "CTE"}:
            messagebox.showwarning("Aviso", "Tipo de documento invalido.")
            return
        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um numero de documento valido.")
            return

        if desfazer_var.get():
            # Desdeclarar - volta para FRANQUIA
            alterados = declarar_documento_frete(tipo, int(numero_texto), "FRANQUIA")
            if alterados == 0:
                messagebox.showwarning("Aviso", f"Documento nao encontrado para des{rotulo_frete.lower()}.")
                return
            messagebox.showinfo(
                "Sucesso",
                f"Documento des{rotulo_frete.lower()} com sucesso.\n\nPara refletir no Excel, gere o faturamento novamente.",
            )
        else:
            # Declarar
            alterados = declarar_documento_frete(tipo, int(numero_texto), rotulo_frete.upper())
            if alterados == 0:
                messagebox.showwarning("Aviso", f"Documento nao encontrado para {rotulo_frete.lower()}.")
                return
            messagebox.showinfo(
                "Sucesso",
                f"Documento declarado como {rotulo_frete}.\n\nPara refletir no Excel, gere o faturamento novamente.",
            )
        dialog.destroy()

    ctk.CTkButton(form, text="Confirmar", command=confirmar, width=220).grid(row=3, column=0, columnspan=2, pady=(14, 0))


def abrir_dialogo_declarar_intercompany():
    _abrir_dialogo_declarar_frete("Intercompany")


def abrir_dialogo_declarar_delta():
    _abrir_dialogo_declarar_frete("Delta")


def abrir_dialogo_cancelar_documento():
    dialog = ctk.CTkToplevel(app)
    dialog.title("Cancelar documento")
    centralizar_janela(dialog, 620, 220)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Tipo de documento").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Numero do documento").grid(row=0, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=260)
    tipo_combo.set("NF")
    tipo_combo.grid(row=1, column=0, sticky="ew", padx=6)

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 20260089 ou 2390")
    numero_entry.grid(row=1, column=1, sticky="ew", padx=6)

    desfazer_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(
        form,
        text="Desfazer cancelamento",
        variable=desfazer_var,
        onvalue=True,
        offvalue=False,
    ).grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def confirmar_cancelamento():
        tipo = tipo_combo.get().strip().upper()
        numero_texto = numero_entry.get().strip()

        if tipo not in {"NF", "CTE"}:
            messagebox.showwarning("Aviso", "Tipo de documento invalido.")
            return

        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um numero de documento valido.")
            return

        numero = int(numero_texto)
        if desfazer_var.get():
            alterados = desfazer_cancelamento_documento(tipo, numero)
            if alterados == 0:
                messagebox.showwarning("Aviso", "Documento nao encontrado ou nao esta cancelado.")
                return
            messagebox.showinfo("Sucesso", "Cancelamento desfeito com sucesso.")
        else:
            alterados = cancelar_documento(tipo, numero)
            if alterados == 0:
                messagebox.showwarning("Aviso", "Documento nao encontrado.")
                return
            messagebox.showinfo("Sucesso", "Documento cancelado com sucesso.")
        dialog.destroy()

    ctk.CTkButton(form, text="Confirmar", command=confirmar_cancelamento, width=200).grid(row=3, column=0, columnspan=2, pady=(14, 0))





def abrir_dialogo_substituir_documento():
    dialog = ctk.CTkToplevel(app)
    dialog.title("Substituir documento")
    centralizar_janela(dialog, 700, 320)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Documento antigo", font=ctk.CTkFont(weight="bold")).grid(
        row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 4)
    )
    ctk.CTkLabel(form, text="Tipo").grid(row=1, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Numero").grid(row=1, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_antigo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=320)
    tipo_antigo_combo.set("NF")
    tipo_antigo_combo.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 10))

    numero_antigo_entry = ctk.CTkEntry(form, width=320, placeholder_text="Ex.: 20260089 ou 2390")
    numero_antigo_entry.grid(row=2, column=1, sticky="ew", padx=6, pady=(0, 10))

    ctk.CTkLabel(form, text="Documento substituto", font=ctk.CTkFont(weight="bold")).grid(
        row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 4)
    )
    ctk.CTkLabel(form, text="Tipo").grid(row=4, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Numero").grid(row=4, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_novo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=320)
    tipo_novo_combo.set("NF")
    tipo_novo_combo.grid(row=5, column=0, sticky="ew", padx=6)

    numero_novo_entry = ctk.CTkEntry(form, width=320, placeholder_text="Ex.: 20260090 ou 2391")
    numero_novo_entry.grid(row=5, column=1, sticky="ew", padx=6)

    desfazer_var = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(
        form,
        text="Desfazer substituicao",
        variable=desfazer_var,
        onvalue=True,
        offvalue=False,
    ).grid(row=6, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def confirmar_substituicao():
        tipo_antigo = tipo_antigo_combo.get().strip().upper()
        tipo_novo = tipo_novo_combo.get().strip().upper()
        numero_antigo_texto = numero_antigo_entry.get().strip()
        numero_novo_texto = numero_novo_entry.get().strip()

        if tipo_antigo not in {"NF", "CTE"} or tipo_novo not in {"NF", "CTE"}:
            messagebox.showwarning("Aviso", "Tipo de documento invalido.")
            return

        if not numero_antigo_texto.isdigit() or not numero_novo_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe numeros de documento validos.")
            return

        numero_antigo = int(numero_antigo_texto)
        numero_novo = int(numero_novo_texto)

        if tipo_antigo == tipo_novo and numero_antigo == numero_novo:
            messagebox.showwarning("Aviso", "Documento antigo e substituto nao podem ser iguais.")
            return

        if desfazer_var.get():
            antigo_restaurado, novo_restaurado = desfazer_substituicao(
                tipo_antigo, numero_antigo, tipo_novo, numero_novo
            )
            if antigo_restaurado == 0:
                messagebox.showwarning("Aviso", "Documento antigo nao encontrado ou nao esta substituido.")
                return
            if novo_restaurado == 0:
                messagebox.showwarning("Aviso", "Documento substituto nao encontrado ou nao esta como substituto.")
                return
            messagebox.showinfo("Sucesso", "Substituicao desfeita com sucesso.")
        else:
            novo_alterado, antigo_alterado = registrar_substituicao(
                tipo_antigo, numero_antigo, tipo_novo, numero_novo
            )
            if novo_alterado == 0:
                messagebox.showwarning("Aviso", "Documento substituto nao encontrado.")
                return
            if antigo_alterado == 0:
                messagebox.showwarning("Aviso", "Documento antigo nao encontrado.")
                return
            messagebox.showinfo("Sucesso", "Substituicao registrada com sucesso.")
        dialog.destroy()

    ctk.CTkButton(form, text="Confirmar", command=confirmar_substituicao, width=220).grid(
        row=7, column=0, columnspan=2, pady=(14, 0)
    )


def abrir_relatorio_cancelados():
    try:
        data_inicial = ler_data_filtro(data_inicio_entry.get(), "Data inicial")
        data_final = ler_data_filtro(data_fim_entry.get(), "Data final")
    except ValueError as exc:
        messagebox.showwarning("Filtro de periodo", str(exc))
        return

    if data_inicial > data_final:
        messagebox.showwarning("Filtro de periodo", "Data inicial nao pode ser maior que a data final.")
        return

    conn = obter_conexao_banco()
    df = pd.read_sql_query(
        """
        SELECT tipo, numero, data_emissao, valor_final, status
        FROM documentos
        WHERE UPPER(status) LIKE '%CANCELADO%'
        """,
        conn,
    )
    conn.close()

    df["data_emissao_dt"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao_dt"])
    df = df[(df["data_emissao_dt"] >= data_inicial) & (df["data_emissao_dt"] <= data_final)]

    if df.empty:
        messagebox.showinfo(
            "Relatorio",
            (
                "Nenhum documento cancelado encontrado no periodo selecionado.\n\n"
                f"Periodo: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}"
            ),
        )
        return

    df = df.sort_values(["data_emissao_dt", "numero"], ascending=[True, True])

    janela = ctk.CTkToplevel(app)
    janela.title("Relatorio de documentos cancelados")
    centralizar_janela(janela, 860, 520)
    janela.grab_set()

    ctk.CTkLabel(
        janela,
        text=(
            "Documentos Cancelados\n"
            f"{data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}"
        ),
        font=ctk.CTkFont(size=16, weight="bold"),
    ).pack(pady=(12, 8))

    container = ctk.CTkFrame(janela)
    container.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    style = ttk.Style()
    style.configure("Cancelados.Treeview", rowheight=28, borderwidth=1, relief="solid")
    style.configure(
        "Cancelados.Treeview.Heading",
        font=("Segoe UI", 10, "bold"),
        relief="solid",
        borderwidth=1,
    )

    colunas = ("tipo", "numero", "data_emissao", "valor", "status")
    tabela = ttk.Treeview(container, columns=colunas, show="headings", style="Cancelados.Treeview")

    tabela.heading("tipo", text="Tipo Doc")
    tabela.heading("numero", text="Numero Doc")
    tabela.heading("data_emissao", text="Data Emissao")
    tabela.heading("valor", text="Valor")
    tabela.heading("status", text="Status")

    tabela.column("tipo", width=110, anchor="center")
    tabela.column("numero", width=170, anchor="center")
    tabela.column("data_emissao", width=170, anchor="center")
    tabela.column("valor", width=180, anchor="e")
    tabela.column("status", width=240, anchor="center")

    scroll_y = ttk.Scrollbar(container, orient="vertical", command=tabela.yview)
    scroll_x = ttk.Scrollbar(container, orient="horizontal", command=tabela.xview)
    tabela.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    tabela.grid(row=0, column=0, sticky="nsew")
    scroll_y.grid(row=0, column=1, sticky="ns")
    scroll_x.grid(row=1, column=0, sticky="ew")

    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    for _, linha in df.iterrows():
        data_txt = (
            linha["data_emissao_dt"].strftime("%d/%m/%Y")
            if pd.notna(linha["data_emissao_dt"])
            else str(linha["data_emissao"])
        )
        valor = float(linha["valor_final"]) if pd.notna(linha["valor_final"]) else 0.0
        tabela.insert(
            "",
            "end",
            values=(
                str(linha["tipo"]),
                int(linha["numero"]),
                data_txt,
                formatar_moeda_brl(valor),
                str(linha["status"]),
            ),
        )




def abrir_busca_documentos():

    dialog = ctk.CTkToplevel(app)
    dialog.title("Buscar documentos")
    centralizar_janela(dialog, 750, 500)
    dialog.grab_set()

    ctk.CTkLabel(
        dialog,
        text="Buscar documento",
        font=ctk.CTkFont(size=16, weight="bold"),
    ).pack(pady=(12, 8))

    busca_entry = ctk.CTkEntry(dialog, width=300, placeholder_text="Digite o numero do documento")
    busca_entry.pack(pady=(0, 10))

    container = ctk.CTkFrame(dialog)
    container.pack(fill="both", expand=True, padx=10, pady=10)

    tabela = ttk.Treeview(
        container,
        columns=("tipo","numero","data","valor","status"),
        show="headings"
    )

    tabela.heading("tipo", text="Tipo")
    tabela.heading("numero", text="Numero")
    tabela.heading("data", text="Data")
    tabela.heading("valor", text="Valor")
    tabela.heading("status", text="Status")

    tabela.column("tipo", width=80, anchor="center")
    tabela.column("numero", width=120, anchor="center")
    tabela.column("data", width=120, anchor="center")
    tabela.column("valor", width=120, anchor="e")
    tabela.column("status", width=200, anchor="center")

    scroll = ttk.Scrollbar(container, orient="vertical", command=tabela.yview)
    tabela.configure(yscrollcommand=scroll.set)

    tabela.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")

    def buscar():

        termo = busca_entry.get().strip()

        for item in tabela.get_children():
            tabela.delete(item)

        conn = obter_conexao_banco()

        df = pd.read_sql_query(
            """
            SELECT tipo, numero, data_emissao, valor_final, status
            FROM documentos
            WHERE numero LIKE ?
            ORDER BY data_emissao
            """,
            conn,
            params=(f"%{termo}%",),
        )

        conn.close()

        for _, row in df.iterrows():

            tabela.insert(
                "",
                "end",
                values=(
                    row["tipo"],
                    row["numero"],
                    row["data_emissao"],
                    formatar_moeda_brl(row["valor_final"]),
                    row["status"],
                ),
            )

    ctk.CTkButton(dialog, text="Buscar", command=buscar, width=200).pack(pady=6)


def _competencia_para_data(comp_str):
    try:
        partes = str(comp_str).lower().split("/")
        if len(partes) == 2:
            mes_nome = normalizar_texto(partes[0].strip()).lower()
            ano = int(partes[1].strip())
            mes_idx = MESES.index(mes_nome) + 1
            return datetime(ano, mes_idx, 1)
    except Exception:
        return None
    return None


def _obter_dataframe_dashboard_filtrado():
    try:
        data_inicial = datetime.strptime(data_inicio_entry.get().strip(), "%d/%m/%Y")
        data_final = datetime.strptime(data_fim_entry.get().strip(), "%d/%m/%Y")
    except ValueError:
        return None, None, None

    if data_inicial > data_final:
        return pd.DataFrame(), data_inicial, data_final

    conn = obter_conexao_banco()
    df = pd.read_sql_query(
        """
        SELECT id, tipo, numero, numero_original, valor_inicial, valor_final, status, competencia
        FROM documentos
        """,
        conn,
    )
    conn.close()

    if df.empty:
        return df, data_inicial, data_final

    df["data_competencia"] = df["competencia"].apply(_competencia_para_data)
    df = df.dropna(subset=["data_competencia"])
    df = df[(df["data_competencia"] >= data_inicial) & (df["data_competencia"] <= data_final)].copy()

    if df.empty:
        return df, data_inicial, data_final

    df["numero"] = pd.to_numeric(df["numero"], errors="coerce")
    df["numero_original_num"] = pd.to_numeric(df.get("numero_original"), errors="coerce")
    df["valor_inicial"] = pd.to_numeric(df["valor_inicial"], errors="coerce")
    df["valor_final"] = pd.to_numeric(df["valor_final"], errors="coerce")
    df["tipo"] = df["tipo"].astype(str).str.upper()
    df["status"] = df["status"].astype(str)

    df["chave_documento"] = df.apply(
        lambda r: (
            f"NF:{int(r['numero_original_num'])}"
            if r["tipo"] == "NF" and pd.notna(r["numero_original_num"])
            else f"{r['tipo']}:{int(r['numero'])}" if pd.notna(r["numero"]) else f"{r['tipo']}:SEMNUM"
        ),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")
    return df, data_inicial, data_final


def _garantir_matplotlib_dashboard():
    if dashboard_chart_state.get("import_ok") is not None:
        return bool(dashboard_chart_state.get("import_ok"))

    try:
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.ticker import FuncFormatter

        dashboard_chart_state.update(
            {
                "import_ok": True,
                "import_error": "",
                "plt": plt,
                "FigureCanvasTkAgg": FigureCanvasTkAgg,
                "FuncFormatter": FuncFormatter,
            }
        )
    except Exception as exc:
        dashboard_chart_state.update(
            {
                "import_ok": False,
                "import_error": str(exc),
                "plt": None,
                "FigureCanvasTkAgg": None,
                "FuncFormatter": None,
            }
        )
    return bool(dashboard_chart_state.get("import_ok"))


def _liberar_graficos_dashboard():
    plt = dashboard_chart_state.get("plt")
    for cfg in dashboard_chart_widgets.values():
        fig = cfg.get("fig")
        if fig is not None and plt is not None:
            try:
                plt.close(fig)
            except Exception:
                pass
        canvas = cfg.get("canvas")
        if canvas is not None:
            try:
                widget = canvas.get_tk_widget()
                if widget.winfo_exists():
                    widget.destroy()
            except Exception:
                pass
    dashboard_chart_widgets.clear()


def _limpar_host_grafico(host):
    if host is None or not host.winfo_exists():
        return
    for child in host.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass


def _mostrar_placeholder_grafico(chave, mensagem):
    cfg = dashboard_chart_widgets.get(chave)
    if not cfg:
        return

    _limpar_host_grafico(cfg.get("host"))
    lbl = ctk.CTkLabel(
        cfg["host"],
        text=mensagem,
        justify="center",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color=UI_THEME["text_secondary"],
    )
    lbl.pack(expand=True, fill="both", padx=8, pady=8)
    cfg["placeholder"] = lbl

    fig_antiga = cfg.get("fig")
    if fig_antiga is not None and dashboard_chart_state.get("plt") is not None:
        try:
            dashboard_chart_state["plt"].close(fig_antiga)
        except Exception:
            pass
    cfg["fig"] = None
    cfg["canvas"] = None


def _renderizar_figura_dashboard(chave, fig):
    cfg = dashboard_chart_widgets.get(chave)
    if not cfg:
        if dashboard_chart_state.get("plt") is not None:
            try:
                dashboard_chart_state["plt"].close(fig)
            except Exception:
                pass
        return

    host = cfg.get("host")
    if host is None or not host.winfo_exists():
        if dashboard_chart_state.get("plt") is not None:
            try:
                dashboard_chart_state["plt"].close(fig)
            except Exception:
                pass
        return

    fig_antiga = cfg.get("fig")
    if fig_antiga is not None and dashboard_chart_state.get("plt") is not None:
        try:
            dashboard_chart_state["plt"].close(fig_antiga)
        except Exception:
            pass

    _limpar_host_grafico(host)
    canvas = dashboard_chart_state["FigureCanvasTkAgg"](fig, master=host)
    canvas.draw()
    widget_canvas = canvas.get_tk_widget()
    try:
        widget_canvas.configure(
            bg=UI_THEME["chart_bg"],
            highlightthickness=0,
            bd=0,
        )
    except Exception:
        pass
    widget_canvas.pack(fill="both", expand=True, padx=8, pady=8)
    cfg["fig"] = fig
    cfg["canvas"] = canvas
    cfg["placeholder"] = None


def _criar_figura_faturamento_periodo(df):
    plt = dashboard_chart_state["plt"]
    from matplotlib.colors import LinearSegmentedColormap
    from matplotlib.patches import FancyBboxPatch, Rectangle

    fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=100)
    fig.patch.set_facecolor(UI_THEME["chart_bg"])
    ax.set_facecolor(UI_THEME["chart_plot_bg"])

    resumo = (
        df.groupby(df["data_competencia"].dt.to_period("M"))["valor_final"]
        .sum()
        .sort_index()
        .reset_index()
    )
    resumo = resumo[resumo["valor_final"].abs() > 0.0001].copy()

    if resumo.empty:
        ax.set_title("Faturamento por periodo", fontsize=11, fontweight="bold", color=UI_THEME["text_primary"], pad=8)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.text(
            0.5,
            0.5,
            "Sem meses com faturamento para o periodo.",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=10,
            color=UI_THEME["text_secondary"],
        )
        fig.tight_layout(pad=1.0)
        return fig

    meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    resumo["periodo_label"] = resumo["data_competencia"].dt.to_timestamp().apply(
        lambda dt: f"{meses_abrev[dt.month - 1]}/{dt.strftime('%y')}"
    )
    resumo["idx"] = range(len(resumo))
    valores = [float(v) for v in resumo["valor_final"].fillna(0).tolist()]
    indices = resumo["idx"].tolist()
    max_valor = max(valores) if valores else 0.0
    limite_superior = max(max_valor * 1.20, 1.0)
    largura_barra = 0.46

    def _misturar_cor(cor_a, cor_b, t):
        c1 = _hex_para_rgb(cor_a)
        c2 = _hex_para_rgb(cor_b)
        mix = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
        return _rgb_para_hex(mix)

    def _gerar_cores_gradiente(qtd):
        if qtd <= 0:
            return []
        if qtd == 1:
            return [_misturar_cor(UI_THEME["chart_bar_secondary"], UI_THEME["chart_bar_primary"], 0.6)]
        return [
            _misturar_cor(UI_THEME["chart_bar_secondary"], UI_THEME["chart_bar_primary"], i / (qtd - 1))
            for i in range(qtd)
        ]

    def _fmt_brl_exato(v):
        return formatar_moeda_brl_exata(v)

    def _calcular_tendencia_linear(valores_seq):
        n = len(valores_seq)
        if n < 2:
            return []
        xs = list(range(n))
        soma_x = sum(xs)
        soma_y = sum(valores_seq)
        soma_x2 = sum(x * x for x in xs)
        soma_xy = sum(x * y for x, y in zip(xs, valores_seq))
        denominador = (n * soma_x2) - (soma_x * soma_x)
        if denominador == 0:
            return []
        m = ((n * soma_xy) - (soma_x * soma_y)) / denominador
        b = (soma_y - (m * soma_x)) / n
        return [(m * x) + b for x in xs]

    ax.set_title("Faturamento por periodo", fontsize=11, fontweight="bold", color=UI_THEME["text_primary"], pad=8)
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_xticks(indices)
    ax.set_xticklabels(resumo["periodo_label"], rotation=0, ha="center", color=UI_THEME["text_secondary"], fontsize=9.5, fontweight="bold")
    ax.set_yticks([])
    ax.tick_params(axis="y", left=False, labelleft=False)
    ax.tick_params(axis="x", pad=6)
    ax.grid(False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_visible(False)
    ax.set_xlim(-0.5, len(indices) - 0.5)
    ax.set_ylim(0, limite_superior)
    ax.margins(x=0.08)

    fundo_topo = _misturar_cor(UI_THEME["chart_plot_bg"], "#FFFFFF", 0.06)
    fundo_base = _misturar_cor(UI_THEME["chart_plot_bg"], "#000000", 0.08)
    gradiente_bg = LinearSegmentedColormap.from_list("dashboard_bg_grad", [fundo_topo, fundo_base])
    ax.imshow(
        [[0], [1]],
        extent=[-0.6, len(indices) - 0.4, 0, limite_superior],
        aspect="auto",
        cmap=gradiente_bg,
        interpolation="bicubic",
        alpha=0.52,
        zorder=0,
    )

    cores = _gerar_cores_gradiente(len(valores))
    for idx_barra, (pos_x, altura, cor) in enumerate(zip(indices, valores, cores)):
        cor_base = cor
        if idx_barra == len(indices) - 1:
            cor_base = _misturar_cor(cor_base, UI_THEME["accent"], 0.34)

        esquerda = pos_x - (largura_barra / 2)
        barra = FancyBboxPatch(
            (esquerda, 0),
            largura_barra,
            max(altura, 0.0001),
            boxstyle=f"round,pad=0,rounding_size={largura_barra * 0.20}",
            linewidth=0,
            facecolor=cor_base,
            edgecolor=cor_base,
            alpha=0.90,
            zorder=2,
        )
        ax.add_patch(barra)

        brilho_topo = Rectangle(
            (esquerda, max(altura * 0.42, 0)),
            largura_barra,
            max(altura * 0.58, 0.0001),
            linewidth=0,
            facecolor=_misturar_cor(cor_base, "#FFFFFF", 0.28),
            alpha=0.36,
            zorder=2.25,
        )
        brilho_topo.set_clip_path(barra)
        ax.add_patch(brilho_topo)

    tendencia = _calcular_tendencia_linear(valores)
    if tendencia:
        ax.plot(
            indices,
            tendencia,
            color=UI_THEME["text_secondary"],
            linewidth=1.1,
            linestyle="--",
            alpha=0.45,
            zorder=3,
        )

    for pos_x, valor in zip(indices, valores):
        if valor <= 0:
            continue
        y_texto = valor * 0.52
        alinhamento_vertical = "center"
        cor_texto = UI_THEME["on_accent"]
        tamanho_fonte = 8.5
        if valor < (limite_superior * 0.11):
            y_texto = valor + (limite_superior * 0.02)
            alinhamento_vertical = "bottom"
            cor_texto = UI_THEME["text_primary"]
            tamanho_fonte = 8
        ax.text(
            pos_x,
            y_texto,
            _fmt_brl_exato(valor),
            ha="center",
            va=alinhamento_vertical,
            fontsize=tamanho_fonte + 0.6,
            color=cor_texto,
            fontweight="bold",
            zorder=4,
        )

    fig.tight_layout(pad=1.0)
    return fig


def _criar_figura_comparativo_tipos(df):
    plt = dashboard_chart_state["plt"]
    fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=100)
    fig.patch.set_facecolor(UI_THEME["chart_bg"])
    ax.set_facecolor(UI_THEME["chart_plot_bg"])

    total_nf = int((df["tipo"] == "NF").sum())
    total_cte = int((df["tipo"] == "CTE").sum())
    total_cancelados = int(df["status"].str.upper().str.contains("CANCELADO", na=False).sum())

    labels = ["NF", "CTE", "Cancelados"]
    valores = [total_nf, total_cte, total_cancelados]
    cores = [
        UI_THEME["metric_nf_value"],
        UI_THEME["metric_cte_value"],
        UI_THEME["chart_cancelados"],
    ]

    barras = ax.bar(labels, valores, color=cores, edgecolor=UI_THEME["chart_axis"], linewidth=0.6, width=0.55, zorder=2)
    ax.set_title("Comparativo de documentos", fontsize=11, fontweight="bold", color=UI_THEME["text_primary"], pad=8)
    ax.set_ylabel("Quantidade", fontsize=9, color=UI_THEME["chart_axis"])
    ax.tick_params(axis="x", colors=UI_THEME["chart_axis"])
    ax.tick_params(axis="y", colors=UI_THEME["chart_axis"])
    ax.grid(axis="y", linestyle="--", linewidth=0.8, alpha=0.28, color=UI_THEME["chart_grid"], zorder=1)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(UI_THEME["chart_grid"])
    ax.spines["bottom"].set_color(UI_THEME["chart_grid"])
    ax.set_ylim(0, max(valores + [1]) * 1.22)

    for barra, valor in zip(barras, valores):
        ax.text(
            barra.get_x() + barra.get_width() / 2,
            barra.get_height() + 0.05,
            f"{valor}",
            ha="center",
            va="bottom",
            fontsize=11,
            color=UI_THEME["text_primary"],
            fontweight="bold",
        )

    fig.tight_layout(pad=1.0)
    return fig


def _atualizar_graficos_dashboard(df, data_inicial=None, data_final=None):
    if not dashboard_chart_widgets:
        return

    if not _garantir_matplotlib_dashboard():
        erro = dashboard_chart_state.get("import_error") or "Erro desconhecido ao carregar matplotlib."
        _mostrar_placeholder_grafico("faturamento", f"Matplotlib indisponivel.\n{erro}")
        _mostrar_placeholder_grafico("comparativo", f"Matplotlib indisponivel.\n{erro}")
        return

    if data_inicial is not None and data_final is not None and data_inicial > data_final:
        msg = "Periodo invalido.\nA data inicial precisa ser menor ou igual a data final."
        _mostrar_placeholder_grafico("faturamento", msg)
        _mostrar_placeholder_grafico("comparativo", msg)
        return

    if df is None:
        return

    if df.empty:
        msg = "Sem dados para o periodo selecionado."
        _mostrar_placeholder_grafico("faturamento", msg)
        _mostrar_placeholder_grafico("comparativo", msg)
        return

    try:
        fig_faturamento = _criar_figura_faturamento_periodo(df)
        fig_comparativo = _criar_figura_comparativo_tipos(df)
    except Exception as exc:
        _mostrar_placeholder_grafico("faturamento", f"Falha ao desenhar grafico.\n{exc}")
        _mostrar_placeholder_grafico("comparativo", f"Falha ao desenhar grafico.\n{exc}")
        return

    _renderizar_figura_dashboard("faturamento", fig_faturamento)
    _renderizar_figura_dashboard("comparativo", fig_comparativo)


def abrir_grafico_faturamento():
    try:
        import matplotlib.pyplot as plt
        from matplotlib.colors import LinearSegmentedColormap
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.patches import FancyBboxPatch, Rectangle
    except Exception as exc:
        messagebox.showerror("Grafico", f"Falha ao carregar bibliotecas do grafico.\n\n{exc}")
        return

    conn = obter_conexao_banco()

    df = pd.read_sql_query(
        """
        SELECT data_emissao, valor_final
        FROM documentos
        WHERE status NOT LIKE '%CANCELADO%'
        """,
        conn,
    )

    conn.close()

    if df.empty:
        messagebox.showwarning("Grafico", "Nao ha dados para gerar grafico.")
        return

    df["data_emissao"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao"])
    df["valor_final"] = pd.to_numeric(df["valor_final"], errors="coerce").fillna(0)

    df["mes"] = df["data_emissao"].dt.to_period("M")

    resumo_base = df.groupby("mes")["valor_final"].sum().reset_index().sort_values("mes").reset_index(drop=True)
    resumo_base = resumo_base[resumo_base["valor_final"].abs() > 0.0001].copy().reset_index(drop=True)
    resumo_base["mes_dt"] = resumo_base["mes"].dt.to_timestamp()
    meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

    def _formatar_mes_ano(valor_dt):
        return f"{meses_abrev[valor_dt.month - 1]}/{valor_dt.strftime('%y')}"

    resumo_base["mes_label"] = resumo_base["mes_dt"].apply(_formatar_mes_ano)
    ano_atual = datetime.now().year

    if resumo_base.empty:
        messagebox.showwarning("Grafico", "Nao ha dados para gerar grafico.")
        return

    janela = ctk.CTkToplevel(app)
    janela.title("Faturamento mensal")
    centralizar_janela(janela, 1020, 620)
    janela.grab_set()

    conteudo = ctk.CTkFrame(janela, fg_color="transparent")
    conteudo.pack(fill="both", expand=True, padx=12, pady=12)

    filtro_card = ctk.CTkFrame(conteudo, fg_color=UI_THEME["surface"], corner_radius=12)
    filtro_card.pack(fill="x", pady=(0, 6))

    topo_filtro = ctk.CTkFrame(filtro_card, fg_color="transparent")
    topo_filtro.pack(fill="x", padx=10, pady=10)

    ctk.CTkLabel(
        topo_filtro,
        text="Filtro",
        font=("Segoe UI", 13, "bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(side="left")

    info_meses_label = ctk.CTkLabel(
        topo_filtro,
        text="",
        font=("Segoe UI", 11),
        text_color=UI_THEME["text_secondary"],
    )
    info_meses_label.pack(side="left", padx=(12, 8))

    botao_meses = ctk.CTkButton(
        topo_filtro,
        text="Meses ▾",
        width=130,
        fg_color=UI_THEME["accent"],
        hover_color=UI_THEME["accent_hover"],
        text_color=UI_THEME["on_accent"],
        border_width=1,
        border_color=UI_THEME["cta_border"],
    )
    botao_meses.pack(side="right")

    dropdown_meses_frame = ctk.CTkFrame(
        conteudo,
        fg_color=UI_THEME["surface_alt"],
        corner_radius=10,
        border_width=1,
        border_color=UI_THEME["border"],
        width=250,
        height=196,
    )
    dropdown_meses_frame.pack_propagate(False)
    altura_lista_meses = min(112, max(72, len(resumo_base) * 24))
    frame_checkboxes = ctk.CTkScrollableFrame(
        dropdown_meses_frame,
        height=altura_lista_meses,
        fg_color=UI_THEME["surface_alt"],
        corner_radius=10,
    )
    frame_checkboxes.pack(fill="both", expand=True, padx=8, pady=(8, 6))

    check_vars = {}
    periodos_disponiveis = resumo_base["mes"].tolist()
    periodos_ano_atual = [p for p in periodos_disponiveis if p.year == ano_atual]
    periodos_padrao = set(periodos_ano_atual if periodos_ano_atual else periodos_disponiveis)

    for _, linha in resumo_base.iterrows():
        periodo = linha["mes"]
        check_vars[periodo] = tk.BooleanVar(value=(periodo in periodos_padrao))
        ctk.CTkCheckBox(
            frame_checkboxes,
            text=linha["mes_label"],
            variable=check_vars[periodo],
            onvalue=True,
            offvalue=False,
            command=lambda: _desenhar_grafico(fechar_dropdown=False, animar=True),
            font=("Segoe UI", 11),
            text_color=UI_THEME["text_primary"],
            checkbox_height=18,
            checkbox_width=18,
        ).pack(anchor="w", padx=8, pady=4)

    acoes_dropdown = ctk.CTkFrame(dropdown_meses_frame, fg_color="transparent")
    acoes_dropdown.pack(fill="x", padx=8, pady=(0, 8))

    botao_marcar_todos = ctk.CTkButton(
        acoes_dropdown,
        text="Marcar todos",
        width=120,
        fg_color=UI_THEME["tab_hover"],
        hover_color=UI_THEME["tab_press"],
        text_color=UI_THEME["text_primary"],
        border_width=1,
        border_color=UI_THEME["border"],
    )
    botao_marcar_todos.pack(side="left", padx=(0, 8))

    botao_limpar = ctk.CTkButton(
        acoes_dropdown,
        text="Limpar",
        width=90,
        fg_color=UI_THEME["surface"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        border_width=1,
        border_color=UI_THEME["border"],
    )
    botao_limpar.pack(side="left")

    dropdown_visivel = False

    frame_grafico = ctk.CTkFrame(conteudo, fg_color="transparent")
    frame_grafico.pack(fill="both", expand=True)

    fig, ax = plt.subplots(figsize=(9.0, 4.8), dpi=110)
    fig.patch.set_facecolor(UI_THEME["chart_bg"])
    canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
    widget_canvas_popup = canvas.get_tk_widget()
    try:
        widget_canvas_popup.configure(
            bg=UI_THEME["chart_bg"],
            highlightthickness=0,
            bd=0,
        )
    except Exception:
        pass
    widget_canvas_popup.pack(fill="both", expand=True)
    animacao_estado = {"job": None}

    def _misturar_cor(cor_a, cor_b, t):
        c1 = _hex_para_rgb(cor_a)
        c2 = _hex_para_rgb(cor_b)
        mix = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
        return _rgb_para_hex(mix)

    def _fmt_brl_exato(v):
        return formatar_moeda_brl_exata(v)

    def _gerar_cores_gradiente(qtd):
        if qtd <= 0:
            return []
        if qtd == 1:
            return [_misturar_cor(UI_THEME["chart_bar_secondary"], UI_THEME["chart_bar_primary"], 0.6)]
        return [
            _misturar_cor(
                UI_THEME["chart_bar_secondary"],
                UI_THEME["chart_bar_primary"],
                i / (qtd - 1),
            )
            for i in range(qtd)
        ]

    def _calcular_tendencia_linear(valores_seq):
        n = len(valores_seq)
        if n < 2:
            return []
        xs = list(range(n))
        soma_x = sum(xs)
        soma_y = sum(valores_seq)
        soma_x2 = sum(x * x for x in xs)
        soma_xy = sum(x * y for x, y in zip(xs, valores_seq))
        denominador = (n * soma_x2) - (soma_x * soma_x)
        if denominador == 0:
            return []
        m = ((n * soma_xy) - (soma_x * soma_y)) / denominador
        b = (soma_y - (m * soma_x)) / n
        return [(m * x) + b for x in xs]

    def _obter_resumo_filtrado():
        periodos_marcados = [periodo for periodo, var in check_vars.items() if bool(var.get())]
        if not periodos_marcados:
            return pd.DataFrame(columns=list(resumo_base.columns) + ["idx"])
        resumo_filtrado = resumo_base[resumo_base["mes"].isin(periodos_marcados)].copy()
        resumo_filtrado = resumo_filtrado.sort_values("mes").reset_index(drop=True)
        resumo_filtrado["idx"] = range(len(resumo_filtrado))
        return resumo_filtrado

    def _atualizar_resumo_meses():
        meses_selecionados = [
            linha["mes_label"] for _, linha in resumo_base.iterrows() if bool(check_vars[linha["mes"]].get())
        ]
        if not meses_selecionados:
            info_meses_label.configure(text="Nenhum mes selecionado")
            return
        if len(meses_selecionados) <= 3:
            info_meses_label.configure(text=", ".join(meses_selecionados))
            return
        info_meses_label.configure(text=f"{len(meses_selecionados)} meses selecionados")

    def _posicionar_dropdown():
        if not dropdown_visivel:
            return
        try:
            conteudo.update_idletasks()
            filtro_card.update_idletasks()
            y_base = filtro_card.winfo_y() + filtro_card.winfo_height() + 2
            dropdown_meses_frame.place(relx=1.0, x=-10, y=y_base, anchor="ne")
            dropdown_meses_frame.lift()
        except Exception:
            pass

    def _alternar_dropdown_meses():
        nonlocal dropdown_visivel
        if dropdown_visivel:
            dropdown_meses_frame.place_forget()
            dropdown_visivel = False
            botao_meses.configure(text="Meses ▾")
            return
        dropdown_visivel = True
        botao_meses.configure(text="Meses ▴")
        _posicionar_dropdown()

    def _cancelar_animacao():
        job = animacao_estado.get("job")
        if job:
            try:
                janela.after_cancel(job)
            except Exception:
                pass
        animacao_estado["job"] = None

    def _desenhar_estado_vazio():
        ax.clear()
        ax.set_facecolor(UI_THEME["chart_plot_bg"])
        ax.set_title("Faturamento Mensal", fontsize=15, fontweight="bold", color=UI_THEME["text_primary"], pad=12)
        ax.set_xlabel("")
        ax.set_yticks([])
        ax.tick_params(axis="y", left=False, labelleft=False)
        ax.grid(False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["bottom"].set_visible(False)
        ax.set_xticks([])
        ax.text(
            0.5,
            0.5,
            "Selecione ao menos um mes para exibir o grafico.",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=11,
            color=UI_THEME["text_secondary"],
        )
        fig.tight_layout(pad=1.4)
        canvas.draw()

    def _renderizar_grafico(resumo, progresso=1.0, exibir_rotulos=False):
        ax.clear()
        ax.set_facecolor(UI_THEME["chart_plot_bg"])

        valores_reais = [float(v) for v in resumo["valor_final"].tolist()]
        indices = resumo["idx"].tolist()
        labels = resumo["mes_label"].tolist()
        qtd = len(valores_reais)
        if qtd == 0:
            _desenhar_estado_vazio()
            return

        fator = max(0.0, min(1.0, float(progresso)))
        valores_animados = [v * fator for v in valores_reais]
        cores = _gerar_cores_gradiente(qtd)

        valor_max = max(valores_reais) if valores_reais else 0
        limite_superior = max(valor_max * 1.18, 1.0)
        largura_barra = 0.46

        ax.set_title("Faturamento Mensal", fontsize=15, fontweight="bold", color=UI_THEME["text_primary"], pad=12)
        ax.set_xlabel("")
        ax.set_yticks([])
        ax.tick_params(axis="y", left=False, labelleft=False)
        ax.grid(False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["bottom"].set_visible(False)
        ax.tick_params(axis="x", colors=UI_THEME["text_secondary"], pad=7)
        ax.set_xticks(indices)
        ax.set_xticklabels(labels, rotation=0, ha="center", fontsize=10.5, fontweight="bold")
        ax.set_xlim(-0.5, qtd - 0.5)
        ax.set_ylim(0, limite_superior)
        ax.margins(x=0.08)

        fundo_topo = _misturar_cor(UI_THEME["chart_plot_bg"], "#FFFFFF", 0.06)
        fundo_base = _misturar_cor(UI_THEME["chart_plot_bg"], "#000000", 0.08)
        gradiente_bg = LinearSegmentedColormap.from_list("popup_bg_grad", [fundo_topo, fundo_base])
        ax.imshow(
            [[0], [1]],
            extent=[-0.6, qtd - 0.4, 0, limite_superior],
            aspect="auto",
            cmap=gradiente_bg,
            interpolation="bicubic",
            alpha=0.54,
            zorder=0,
        )

        alpha_barra = 0.35 + (0.60 * fator)
        for idx_barra, (pos_x, altura, cor) in enumerate(zip(indices, valores_animados, cores)):
            cor_base = cor
            if idx_barra == len(indices) - 1:
                cor_base = _misturar_cor(cor_base, UI_THEME["accent"], 0.34)

            esquerda = pos_x - (largura_barra / 2)
            altura_visivel = max(altura, 0.0001)
            barra = FancyBboxPatch(
                (esquerda, 0),
                largura_barra,
                altura_visivel,
                boxstyle=f"round,pad=0,rounding_size={largura_barra * 0.20}",
                linewidth=0,
                facecolor=cor_base,
                edgecolor=cor_base,
                alpha=max(0.52, alpha_barra),
                zorder=2,
            )
            ax.add_patch(barra)

            brilho_topo = Rectangle(
                (esquerda, max(altura_visivel * 0.42, 0)),
                largura_barra,
                max(altura_visivel * 0.58, 0.0001),
                linewidth=0,
                facecolor=_misturar_cor(cor_base, "#FFFFFF", 0.28),
                alpha=0.36,
                zorder=2.25,
            )
            brilho_topo.set_clip_path(barra)
            ax.add_patch(brilho_topo)

        tendencia = _calcular_tendencia_linear(valores_reais)
        if tendencia:
            ax.plot(
                indices,
                tendencia,
                color=UI_THEME["text_secondary"],
                linewidth=1.15,
                linestyle="--",
                alpha=0.45,
                zorder=3,
            )

        if exibir_rotulos:
            for pos_x, valor in zip(indices, valores_reais):
                if valor <= 0:
                    continue
                y_texto = valor * 0.52
                alinhamento_vertical = "center"
                cor_texto = UI_THEME["on_accent"]
                tamanho_fonte = 10
                if valor < (limite_superior * 0.10):
                    # Evita corte em barras pequenas.
                    y_texto = valor + (limite_superior * 0.025)
                    alinhamento_vertical = "bottom"
                    cor_texto = UI_THEME["text_primary"]
                    tamanho_fonte = 9
                ax.text(
                    pos_x,
                    y_texto,
                    _fmt_brl_exato(valor),
                    ha="center",
                    va=alinhamento_vertical,
                    fontsize=tamanho_fonte + 0.9,
                    fontweight="bold",
                    color=cor_texto,
                    zorder=4,
                )

        periodo_ini = _formatar_mes_ano(resumo["mes_dt"].min())
        periodo_fim = _formatar_mes_ano(resumo["mes_dt"].max())
        resumo_txt = f"Periodo exibido: {periodo_ini} a {periodo_fim}"
        ax.text(
            0.01,
            1.02,
            resumo_txt,
            transform=ax.transAxes,
            fontsize=9,
            color=UI_THEME["text_secondary"],
            ha="left",
            va="bottom",
        )

        fig.tight_layout(pad=1.4)
        canvas.draw()

    def _animar_grafico(resumo):
        _cancelar_animacao()
        total_frames = 6
        estado = {"frame": 0}

        def _passo():
            frame_atual = estado["frame"]
            t = (frame_atual + 1) / total_frames
            progresso = 1 - ((1 - t) ** 2)
            _renderizar_grafico(
                resumo,
                progresso=progresso,
                exibir_rotulos=(frame_atual >= total_frames - 1),
            )
            if frame_atual >= total_frames - 1:
                animacao_estado["job"] = None
                return
            estado["frame"] += 1
            animacao_estado["job"] = janela.after(16, _passo)

        _passo()

    def _desenhar_grafico(fechar_dropdown=True, animar=False):
        nonlocal dropdown_visivel
        resumo = _obter_resumo_filtrado()
        _atualizar_resumo_meses()

        if resumo.empty:
            _cancelar_animacao()
            _desenhar_estado_vazio()
            return

        if animar:
            _animar_grafico(resumo)
        else:
            _cancelar_animacao()
            _renderizar_grafico(
                resumo,
                progresso=1.0,
                exibir_rotulos=True,
            )

        if fechar_dropdown and dropdown_visivel:
            dropdown_meses_frame.place_forget()
            dropdown_visivel = False
            botao_meses.configure(text="Meses ▾")

    def _marcar_todos():
        for var in check_vars.values():
            var.set(True)
        _desenhar_grafico(fechar_dropdown=False, animar=True)

    def _limpar_todos():
        for var in check_vars.values():
            var.set(False)
        _desenhar_grafico(fechar_dropdown=False, animar=True)

    botao_meses.configure(command=_alternar_dropdown_meses)
    botao_marcar_todos.configure(command=_marcar_todos)
    botao_limpar.configure(command=_limpar_todos)
    _atualizar_resumo_meses()

    janela.bind("<Configure>", lambda _e: _posicionar_dropdown(), add="+")

    _desenhar_grafico(animar=False)

    def _fechar_janela():
        _cancelar_animacao()
        try:
            plt.close(fig)
        except Exception:
            pass
        janela.destroy()

    janela.protocol("WM_DELETE_WINDOW", _fechar_janela)


def atualizar_dashboard():
    global dashboard_update_running
    if dashboard_update_running:
        return
    dashboard_update_running = True

    def atualizar_resumo(total_inicial=0.0, total_final=0.0, total_docs=0, total_nf=0, total_cte=0, total_cancelados=0):
        total_inicial_valor_label.configure(text=formatar_moeda_brl(total_inicial))
        total_final_valor_label.configure(text=formatar_moeda_brl(total_final))
        total_documentos_label.configure(text=f"{total_docs}")
        nf_label.configure(text=f"{total_nf}")
        cte_label.configure(text=f"{total_cte}")
        cancelados_label.configure(text=f"{total_cancelados}")

        diferenca = total_inicial - total_final
        if diferenca > 0:
            diferenca_label.configure(
                text=f"Diferenca referente a Impostos de NFS-e: {formatar_moeda_brl(diferenca)}",
                text_color=UI_THEME["success_text"],
            )
        elif diferenca < 0:
            diferenca_label.configure(
                text=f"Diferenca referente a Impostos de NFS-e: +{formatar_moeda_brl(abs(diferenca))}",
                text_color=UI_THEME["danger_text"],
            )
        else:
            diferenca_label.configure(
                text="Diferenca referente a Impostos de NFS-e: R$ 0,00",
                text_color=UI_THEME["text_secondary"],
            )

        if "cancelados_card" in globals():
            if total_cancelados > 0:
                cancelados_card.configure(fg_color=UI_THEME["danger_bg"])
                cancelados_label.configure(text_color=UI_THEME["danger_text"])
            else:
                cfg_cancel = summary_metric_widgets.get("cancelados", {})
                cancelados_card.configure(fg_color=UI_THEME.get(cfg_cancel.get("bg_key", "metric_default_bg"), UI_THEME["metric_default_bg"]))
                cancelados_label.configure(text_color=UI_THEME.get(cfg_cancel.get("value_key", "metric_default_value"), UI_THEME["metric_default_value"]))

    try:
        df, data_inicial, data_final = _obter_dataframe_dashboard_filtrado()
        if df is None:
            # Enquanto o usuario edita parcialmente a data, nao recalcula.
            return
        if data_inicial is not None and data_final is not None and data_inicial > data_final:
            atualizar_resumo()
            _atualizar_graficos_dashboard(pd.DataFrame(), data_inicial, data_final)
            return
        if df.empty:
            atualizar_resumo()
            _atualizar_graficos_dashboard(df, data_inicial, data_final)
            return

        total_inicial = df["valor_inicial"].fillna(0).sum()
        total_final = df["valor_final"].fillna(0).sum()
        total_docs = len(df)
        total_nf = int((df["tipo"] == "NF").sum())
        total_cte = int((df["tipo"] == "CTE").sum())
        total_cancelados = int(df["status"].str.upper().str.contains("CANCELADO", na=False).sum())

        atualizar_resumo(
            float(total_inicial),
            float(total_final),
            int(total_docs),
            total_nf,
            total_cte,
            total_cancelados,
        )
        _atualizar_graficos_dashboard(df, data_inicial, data_final)
    finally:
        dashboard_update_running = False


tab_buttons = {}
tab_hover_state = {}
tab_active_id = "principal"
tab_indicator = None
tab_indicator_host = None
menu_feedback_label = None
theme_toggle_button = None
ui_refs = {}
summary_metric_widgets = {}
dashboard_chart_widgets = {}
dashboard_chart_state = {
    "import_ok": None,
    "import_error": "",
    "plt": None,
    "FigureCanvasTkAgg": None,
    "FuncFormatter": None,
}
_anim_jobs = {}
screen_nav_buttons = {}
SCREEN_NAV_STYLES = {}
action_buttons = []


def _safe_config(widget, **kwargs):
    try:
        if widget is not None and widget.winfo_exists():
            widget.configure(**kwargs)
    except Exception:
        pass


def _texto_botao_tema():
    return "🌙 Tema escuro" if current_theme_mode == "light" else "☀ Tema claro"


def _definir_tema_interface(modo, persistir=True):
    global current_theme_mode, UI_THEME, TAB_STYLES

    novo_modo = "dark" if str(modo).strip().lower() == "dark" else "light"
    current_theme_mode = novo_modo
    UI_THEME = dict(APP_THEMES[novo_modo])
    TAB_STYLES = _gerar_tab_styles()
    ctk.set_appearance_mode("dark" if novo_modo == "dark" else "light")

    if persistir:
        try:
            salvar_configuracao("tema_interface", novo_modo)
        except Exception:
            pass


def aplicar_tema_interface():
    app.configure(fg_color=UI_THEME["app_bg"])

    _safe_config(
        container_scroll if "container_scroll" in globals() else None,
        fg_color=UI_THEME["app_bg"],
        scrollbar_button_color=UI_THEME["scroll_btn"],
        scrollbar_button_hover_color=UI_THEME["scroll_btn_hover"],
    )
    _safe_config(ui_refs.get("main_frame"), fg_color=UI_THEME["app_bg"])
    _safe_config(ui_refs.get("screen_host"), fg_color=UI_THEME["app_bg"])
    if ui_refs.get("logo_watermark_label") is not None:
        try:
            ui_refs["logo_watermark_label"].lower()
        except Exception:
            pass
    for frame_tela in getattr(app, "screens", {}).values():
        _safe_config(frame_tela, fg_color=UI_THEME["app_bg"])

    _safe_config(ui_refs.get("header_card"), fg_color=UI_THEME["header_bg"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("header_title"), text_color=UI_THEME["text_primary"])
    _safe_config(
        theme_toggle_button,
        text=_texto_botao_tema(),
        fg_color=UI_THEME["surface_alt"],
        text_color=UI_THEME["text_primary"],
        hover_color=UI_THEME["tab_hover"],
        border_color=UI_THEME["border"],
    )

    _safe_config(ui_refs.get("filtro_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("filtro_titulo"), text_color=UI_THEME["text_primary"])
    _safe_config(ui_refs.get("filtro_subtitulo"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("filtro_ate"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("filtro_icon_inicio"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("filtro_icon_fim"), text_color=UI_THEME["text_secondary"])
    _safe_config(data_inicio_entry if "data_inicio_entry" in globals() else None, border_color=UI_THEME["border"])
    _safe_config(data_fim_entry if "data_fim_entry" in globals() else None, border_color=UI_THEME["border"])

    buscar_btn = ui_refs.get("buscar_btn")
    _safe_config(
        buscar_btn,
        fg_color=UI_THEME["accent"],
        hover_color=UI_THEME["accent_hover"],
        text_color=UI_THEME["on_accent"],
        border_color=UI_THEME["cta_border"],
    )
    if buscar_btn is not None:
        _aplicar_microinteracao_cta(
            buscar_btn,
            UI_THEME["accent"],
            UI_THEME["accent_hover"],
            UI_THEME["cta_press"],
            UI_THEME["cta_border"],
            UI_THEME["cta_border_hover"],
        )

    _safe_config(ui_refs.get("info_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("info_divider"), fg_color=UI_THEME["divider"])
    _safe_config(pasta_saida_label if "pasta_saida_label" in globals() else None, text_color=UI_THEME["text_secondary"])
    _safe_config(pasta_label if "pasta_label" in globals() else None, text_color=UI_THEME["text_primary"])
    _safe_config(
        progress if "progress" in globals() else None,
        fg_color=UI_THEME["progress_bg"],
        progress_color=UI_THEME["accent"],
    )
    _safe_config(status_label if "status_label" in globals() else None, text_color=UI_THEME["text_secondary"])

    _safe_config(ui_refs.get("tabs_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    for tab_id, btn in tab_buttons.items():
        _safe_config(btn, hover_color=TAB_STYLES["hover_fg"], border_color=UI_THEME["border"])
        _aplicar_microinteracao_botao(btn, tab_id)
    _safe_config(tab_indicator, fg_color=TAB_STYLES["active_fg"])
    _safe_config(menu_feedback_label, text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("screen_nav_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    SCREEN_NAV_STYLES.update(
        {
            "normal_fg": UI_THEME["surface_alt"],
            "normal_text": UI_THEME["text_primary"],
            "hover_fg": UI_THEME["tab_hover"],
            "active_fg": UI_THEME["accent"],
            "active_hover_fg": UI_THEME["tab_active_hover"],
            "active_press_fg": UI_THEME["tab_active_press"],
            "active_text": UI_THEME["on_accent"],
        }
    )
    for btn in action_buttons:
        _safe_config(
            btn,
            fg_color=UI_THEME["accent"],
            hover_color=UI_THEME["accent_hover"],
            text_color=UI_THEME["on_accent"],
            border_color=UI_THEME["cta_border"],
        )
        _aplicar_microinteracao_cta(
            btn,
            UI_THEME["accent"],
            UI_THEME["accent_hover"],
            UI_THEME["cta_press"],
            UI_THEME["cta_border"],
            UI_THEME["cta_border_hover"],
        )
    _configurar_nav_tela_ativa(app.current_screen or "dashboard")

    _safe_config(ui_refs.get("resumo_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("resumo_titulo"), text_color=UI_THEME["text_primary"])
    _safe_config(diferenca_label if "diferenca_label" in globals() else None, text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("graficos_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("graficos_titulo"), text_color=UI_THEME["text_primary"])
    _safe_config(ui_refs.get("graficos_subtitulo"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("grafico_faturamento_card"), fg_color=UI_THEME["surface_alt"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("grafico_comparativo_card"), fg_color=UI_THEME["surface_alt"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("grafico_faturamento_titulo"), text_color=UI_THEME["text_primary"])
    _safe_config(ui_refs.get("grafico_comparativo_titulo"), text_color=UI_THEME["text_primary"])
    for cfg in dashboard_chart_widgets.values():
        _safe_config(cfg.get("placeholder"), text_color=UI_THEME["text_secondary"])

    for metric_key, cfg in summary_metric_widgets.items():
        _safe_config(cfg["card"], fg_color=UI_THEME[cfg["bg_key"]], border_color=UI_THEME["border"])
        _safe_config(cfg["title"], text_color=UI_THEME[cfg["title_key"]])
        _safe_config(cfg["value"], text_color=UI_THEME[cfg["value_key"]])

    _configurar_tab_ativo(tab_active_id)
    atualizar_dashboard()


def alternar_tema_interface():
    novo_modo = "dark" if current_theme_mode == "light" else "light"
    _definir_tema_interface(novo_modo, persistir=True)
    aplicar_tema_interface()

def _criar_card(parent, fg_color=None, corner_radius=18, border_width=1):
    return ctk.CTkFrame(
        parent,
        fg_color=fg_color or UI_THEME["surface"],
        corner_radius=corner_radius,
        border_width=border_width,
        border_color=UI_THEME["border"],
    )


def _normalizar_hex_cor(cor):
    if isinstance(cor, (tuple, list)) and cor:
        cor = cor[0]
    cor = str(cor).strip()
    if not cor.startswith("#"):
        return "#000000"
    if len(cor) == 4:
        return "#" + "".join(ch * 2 for ch in cor[1:])
    if len(cor) != 7:
        return "#000000"
    return cor


def _hex_para_rgb(cor):
    c = _normalizar_hex_cor(cor)
    return tuple(int(c[i:i + 2], 16) for i in (1, 3, 5))


def _rgb_para_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in rgb]
    return f"#{r:02x}{g:02x}{b:02x}"


def _interpolar_cor(cor_a, cor_b, t):
    a = _hex_para_rgb(cor_a)
    b = _hex_para_rgb(cor_b)
    return _rgb_para_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


def _animar_estilo_botao(botao, alvo_fg, alvo_texto, alvo_borda, passos=7, delay_ms=16):
    chave = f"btn_{id(botao)}"
    job = _anim_jobs.get(chave)
    if job:
        try:
            app.after_cancel(job)
        except Exception:
            pass

    try:
        inicio_fg = _normalizar_hex_cor(botao.cget("fg_color"))
        inicio_texto = _normalizar_hex_cor(botao.cget("text_color"))
        inicio_borda = _normalizar_hex_cor(botao.cget("border_color"))
    except Exception:
        inicio_fg = _normalizar_hex_cor(alvo_fg)
        inicio_texto = _normalizar_hex_cor(alvo_texto)
        inicio_borda = _normalizar_hex_cor(alvo_borda)

    alvo_fg = _normalizar_hex_cor(alvo_fg)
    alvo_texto = _normalizar_hex_cor(alvo_texto)
    alvo_borda = _normalizar_hex_cor(alvo_borda)

    def _passo(idx):
        if not botao.winfo_exists():
            return
        t = idx / float(passos)
        botao.configure(
            fg_color=_interpolar_cor(inicio_fg, alvo_fg, t),
            text_color=_interpolar_cor(inicio_texto, alvo_texto, t),
            border_color=_interpolar_cor(inicio_borda, alvo_borda, t),
        )
        if idx < passos:
            _anim_jobs[chave] = app.after(delay_ms, lambda: _passo(idx + 1))

    _passo(1)


def _animar_indicador_tab(tab_id):
    global tab_indicator
    if not tab_indicator_host or not tab_indicator:
        return
    botao = tab_buttons.get(tab_id)
    if not botao or not botao.winfo_exists():
        return

    tab_indicator_host.update_idletasks()
    x_alvo = botao.winfo_x() + 8
    w_alvo = max(24, botao.winfo_width() - 16)
    y_alvo = botao.winfo_y() + botao.winfo_height() + 6

    if not tab_indicator.winfo_ismapped():
        tab_indicator.configure(width=w_alvo, height=3)
        tab_indicator.place(x=x_alvo, y=y_alvo)
        return

    try:
        x_ini = int(tab_indicator.place_info().get("x", x_alvo))
        w_ini = int(tab_indicator.place_info().get("width", w_alvo))
    except Exception:
        x_ini, w_ini = x_alvo, w_alvo

    passos = 8
    delay_ms = 14

    def _passo(i):
        if not tab_indicator.winfo_exists():
            return
        t = i / float(passos)
        x = int(x_ini + (x_alvo - x_ini) * t)
        w = int(w_ini + (w_alvo - w_ini) * t)
        tab_indicator.configure(width=w, height=3)
        tab_indicator.place(x=x, y=y_alvo)
        if i < passos:
            app.after(delay_ms, lambda: _passo(i + 1))

    _passo(1)


def _aplicar_microinteracao_botao(botao, tab_id):
    tab_hover_state[tab_id] = False

    def _on_enter(_evt):
        tab_hover_state[tab_id] = True
        ativo = tab_active_id == tab_id
        _animar_estilo_botao(
            botao,
            TAB_STYLES["active_hover_fg"] if ativo else TAB_STYLES["hover_fg"],
            TAB_STYLES["active_text"] if ativo else TAB_STYLES["normal_text"],
            TAB_STYLES["active_fg"] if ativo else UI_THEME["border"],
        )

    def _on_leave(_evt):
        tab_hover_state[tab_id] = False
        ativo = tab_active_id == tab_id
        _animar_estilo_botao(
            botao,
            TAB_STYLES["active_fg"] if ativo else TAB_STYLES["normal_fg"],
            TAB_STYLES["active_text"] if ativo else TAB_STYLES["normal_text"],
            TAB_STYLES["active_fg"] if ativo else UI_THEME["border"],
        )

    def _on_press(_evt):
        ativo = tab_active_id == tab_id
        _animar_estilo_botao(
            botao,
            TAB_STYLES["active_press_fg"] if ativo else TAB_STYLES["press_fg"],
            TAB_STYLES["active_text"] if ativo else TAB_STYLES["normal_text"],
            TAB_STYLES["active_fg"] if ativo else UI_THEME["border"],
            passos=4,
            delay_ms=12,
        )

    def _on_release(_evt):
        ativo = tab_active_id == tab_id
        hovered = tab_hover_state.get(tab_id, False)
        _animar_estilo_botao(
            botao,
            TAB_STYLES["active_hover_fg"] if (ativo and hovered) else (
                TAB_STYLES["active_fg"] if ativo else (
                    TAB_STYLES["hover_fg"] if hovered else TAB_STYLES["normal_fg"]
                )
            ),
            TAB_STYLES["active_text"] if ativo else TAB_STYLES["normal_text"],
            TAB_STYLES["active_fg"] if ativo else UI_THEME["border"],
        )

    botao.bind("<Enter>", _on_enter)
    botao.bind("<Leave>", _on_leave)
    botao.bind("<ButtonPress-1>", _on_press)
    botao.bind("<ButtonRelease-1>", _on_release)


def _aplicar_microinteracao_cta(botao, base_fg, hover_fg, press_fg, base_border, hover_border):
    def _on_enter(_evt):
        _animar_estilo_botao(
            botao,
            hover_fg,
            UI_THEME["on_accent"],
            hover_border,
            passos=6,
            delay_ms=14,
        )

    def _on_leave(_evt):
        _animar_estilo_botao(
            botao,
            base_fg,
            UI_THEME["on_accent"],
            base_border,
            passos=6,
            delay_ms=14,
        )

    def _on_press(_evt):
        _animar_estilo_botao(
            botao,
            press_fg,
            UI_THEME["on_accent"],
            hover_border,
            passos=4,
            delay_ms=10,
        )

    def _on_release(_evt):
        _animar_estilo_botao(
            botao,
            hover_fg,
            UI_THEME["on_accent"],
            hover_border,
            passos=5,
            delay_ms=12,
        )

    botao.bind("<Enter>", _on_enter)
    botao.bind("<Leave>", _on_leave)
    botao.bind("<ButtonPress-1>", _on_press)
    botao.bind("<ButtonRelease-1>", _on_release)


def _configurar_nav_tela_ativa(nome_tela):
    if not SCREEN_NAV_STYLES:
        return
    for chave, botao in screen_nav_buttons.items():
        ativo = chave == nome_tela
        alvo_fg = SCREEN_NAV_STYLES["active_fg"] if ativo else SCREEN_NAV_STYLES["normal_fg"]
        alvo_text = SCREEN_NAV_STYLES["active_text"] if ativo else SCREEN_NAV_STYLES["normal_text"]
        alvo_border = SCREEN_NAV_STYLES["active_fg"] if ativo else UI_THEME["border"]
        _safe_config(
            botao,
            hover_color=SCREEN_NAV_STYLES["active_hover_fg"] if ativo else SCREEN_NAV_STYLES["hover_fg"],
        )
        _animar_estilo_botao(
            botao,
            alvo_fg,
            alvo_text,
            alvo_border,
            passos=5,
            delay_ms=12,
        )


def _navegar_tela_com_feedback(nome_tela):
    botao = screen_nav_buttons.get(nome_tela)
    if botao is not None and botao.winfo_exists() and SCREEN_NAV_STYLES:
        _animar_estilo_botao(
            botao,
            SCREEN_NAV_STYLES["active_press_fg"],
            SCREEN_NAV_STYLES["active_text"],
            SCREEN_NAV_STYLES["active_fg"],
            passos=4,
            delay_ms=10,
        )
    app.after(55, lambda: app.mostrar_tela(nome_tela))


def _criar_botao_acao(parent, texto, comando):
    botao = ctk.CTkButton(
        parent,
        text=texto,
        height=40,
        corner_radius=12,
        border_width=1,
        border_color=UI_THEME["cta_border"],
        fg_color=UI_THEME["accent"],
        hover_color=UI_THEME["accent_hover"],
        text_color=UI_THEME["on_accent"],
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        command=comando,
    )
    _aplicar_microinteracao_cta(
        botao,
        UI_THEME["accent"],
        UI_THEME["accent_hover"],
        UI_THEME["cta_press"],
        UI_THEME["cta_border"],
        UI_THEME["cta_border_hover"],
    )
    action_buttons.append(botao)
    return botao


def _configurar_tab_ativo(tab_id):
    global tab_active_id
    tab_active_id = tab_id
    for chave, botao in tab_buttons.items():
        ativo = chave == tab_id
        hovered = tab_hover_state.get(chave, False)
        alvo_fg = (
            TAB_STYLES["active_hover_fg"] if (ativo and hovered) else
            TAB_STYLES["active_fg"] if ativo else
            TAB_STYLES["hover_fg"] if hovered else
            TAB_STYLES["normal_fg"]
        )
        _animar_estilo_botao(
            botao,
            alvo_fg,
            TAB_STYLES["active_text"] if ativo else TAB_STYLES["normal_text"],
            TAB_STYLES["active_fg"] if ativo else UI_THEME["border"],
        )
    _animar_indicador_tab(tab_id)


def _executar_acao_menu(tab_id, titulo_item, comando):
    if menu_feedback_label and menu_feedback_label.winfo_exists():
        menu_feedback_label.configure(text=f"✔ {titulo_item}")

    botao = tab_buttons.get(tab_id)
    if botao and botao.winfo_exists():
        _animar_estilo_botao(
            botao,
            TAB_STYLES["active_press_fg"],
            TAB_STYLES["active_text"],
            TAB_STYLES["active_fg"],
            passos=5,
            delay_ms=12,
        )
        app.after(70, lambda: _configurar_tab_ativo(tab_id))

    comando()


def abrir_menu_dropdown(botao, opcoes, tab_id, tab_titulo):
    _configurar_tab_ativo(tab_id)
    if menu_feedback_label and menu_feedback_label.winfo_exists():
        menu_feedback_label.configure(text=f"Abrindo menu de {tab_titulo}...")

    menu = tk.Menu(
        app,
        tearoff=0,
        bg=UI_THEME["surface"],
        fg=UI_THEME["text_primary"],
        activebackground=UI_THEME["accent"],
        activeforeground=UI_THEME["on_accent"],
        selectcolor=UI_THEME["accent"],
        borderwidth=1,
        relief="solid",
    )
    menu.configure(font=("Segoe UI", 10))
    for item in opcoes:
        if item is None:
            menu.add_separator()
            continue
        titulo, comando = item
        menu.add_command(
            label=f"  {titulo}",
            command=lambda t=titulo, c=comando: _executar_acao_menu(tab_id, t, c),
        )
    try:
        menu.tk_popup(botao.winfo_rootx(), botao.winfo_rooty() + botao.winfo_height() + 4)
    finally:
        menu.grab_release()


def _obter_logo_watermark():
    if not os.path.exists(LOGO_PATH):
        return None

    cache_attr = "_logo_watermark"
    if hasattr(app, cache_attr):
        return getattr(app, cache_attr)

    try:
        with Image.open(LOGO_PATH) as origem:
            base = origem.convert("RGBA")

        def _processar_logo(opacidade):
            img = base.copy()
            novos_pixels = []
            for r, g, b, a in img.getdata():
                # Remove fundo claro para evitar "placa" branca.
                if r > 242 and g > 242 and b > 242:
                    novos_pixels.append((r, g, b, 0))
                    continue
                alpha = int(a * opacidade)
                novos_pixels.append((r, g, b, max(0, min(255, alpha))))
            img.putdata(novos_pixels)
            return img.filter(ImageFilter.GaussianBlur(radius=0.9))

        # Opacidade bem baixa (5% a 10%), com leve ajuste por tema.
        logo_light = _processar_logo(0.05)
        logo_dark = _processar_logo(0.08)

        logo_wm = ctk.CTkImage(
            light_image=logo_light,
            dark_image=logo_dark,
            size=(760, 320),
        )
        setattr(app, cache_attr, logo_wm)
        return logo_wm
    except Exception as e:
        print(f"Aviso: nao foi possivel preparar a watermark da logo - {e}")
        return None


def _aplicar_logo_watermark(container):
    if container is None or not container.winfo_exists():
        return

    logo_wm = _obter_logo_watermark()
    if logo_wm is None:
        return

    lbl_existente = ui_refs.get("logo_watermark_label")
    if lbl_existente is not None and lbl_existente.winfo_exists():
        try:
            lbl_existente.destroy()
        except Exception:
            pass

    lbl = ctk.CTkLabel(container, image=logo_wm, text="", fg_color="transparent")
    lbl.place(relx=0.5, rely=0.52, anchor="center")
    lbl.lower()
    ui_refs["logo_watermark_label"] = lbl


def _criar_header(parent):
    global theme_toggle_button

    header = _criar_card(parent, fg_color=UI_THEME["header_bg"], corner_radius=22)
    header.pack(fill="x", padx=18, pady=(8, 8))
    ui_refs["header_card"] = header

    header_top = ctk.CTkFrame(header, fg_color="transparent")
    header_top.pack(fill="x", padx=14, pady=(8, 0))
    theme_toggle_button = ctk.CTkButton(
        header_top,
        text=_texto_botao_tema(),
        width=120,
        height=30,
        corner_radius=12,
        border_width=1,
        fg_color=UI_THEME["surface_alt"],
        border_color=UI_THEME["border"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        command=alternar_tema_interface,
    )
    theme_toggle_button.pack(side="right")

    titulo = ctk.CTkLabel(
        header,
        text="Sistema de Faturamento - Horizonte Logistica",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.pack(pady=(10, 12))
    ui_refs["header_title"] = titulo


def _criar_filtro_periodo(parent):
    global data_inicio_entry, data_fim_entry

    card = _criar_card(parent, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 10))
    card.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
    ui_refs["filtro_card"] = card

    titulo = ctk.CTkLabel(
        card,
        text="Periodo de emissao",
        font=ctk.CTkFont(family="Segoe UI", size=19, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.grid(row=0, column=0, columnspan=6, sticky="w", padx=20, pady=(16, 4))
    ui_refs["filtro_titulo"] = titulo

    subtitulo = ctk.CTkLabel(
        card,
        text="Filtre os documentos por intervalo de emissao",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    )
    subtitulo.grid(row=1, column=0, columnspan=6, sticky="w", padx=20, pady=(0, 12))
    ui_refs["filtro_subtitulo"] = subtitulo

    icon_inicio = ctk.CTkLabel(card, text="📅", font=ctk.CTkFont(size=16), text_color=UI_THEME["text_secondary"])
    icon_inicio.grid(
        row=2, column=0, sticky="e", padx=(20, 6), pady=(0, 16)
    )
    ui_refs["filtro_icon_inicio"] = icon_inicio

    data_inicio_entry = ctk.CTkEntry(
        card,
        width=150,
        height=44,
        corner_radius=14,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    data_inicio_entry.insert(0, primeiro_dia_padrao.strftime("%d/%m/%Y"))
    data_inicio_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(0, 16))

    lbl_ate = ctk.CTkLabel(card, text="ate", font=ctk.CTkFont(family="Segoe UI", size=12), text_color=UI_THEME["text_secondary"])
    lbl_ate.grid(
        row=2, column=2, padx=4, pady=(0, 16)
    )
    ui_refs["filtro_ate"] = lbl_ate

    icon_fim = ctk.CTkLabel(card, text="📅", font=ctk.CTkFont(size=16), text_color=UI_THEME["text_secondary"])
    icon_fim.grid(
        row=2, column=3, sticky="e", padx=(8, 6), pady=(0, 16)
    )
    ui_refs["filtro_icon_fim"] = icon_fim

    data_fim_entry = ctk.CTkEntry(
        card,
        width=150,
        height=44,
        corner_radius=14,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    data_fim_entry.insert(0, ultimo_dia_padrao.strftime("%d/%m/%Y"))
    data_fim_entry.grid(row=2, column=4, sticky="ew", padx=(0, 10), pady=(0, 16))

    buscar_btn = ctk.CTkButton(
        card,
        text="Buscar",
        width=126,
        height=44,
        corner_radius=14,
        border_width=1,
        border_color=UI_THEME["cta_border"],
        fg_color=UI_THEME["accent"],
        hover_color=UI_THEME["accent_hover"],
        text_color=UI_THEME["on_accent"],
        font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
        command=atualizar_dashboard,
    )
    buscar_btn.grid(row=2, column=5, sticky="e", padx=(6, 20), pady=(0, 16))
    ui_refs["buscar_btn"] = buscar_btn
    _aplicar_microinteracao_cta(
        buscar_btn,
        UI_THEME["accent"],
        UI_THEME["accent_hover"],
        UI_THEME["cta_press"],
        UI_THEME["cta_border"],
        UI_THEME["cta_border_hover"],
    )

    data_inicio_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(data_inicio_entry))
    data_inicio_entry.bind("<FocusOut>", solicitar_atualizacao_dashboard)
    data_inicio_entry.bind("<Return>", solicitar_atualizacao_dashboard)
    data_fim_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(data_fim_entry))
    data_fim_entry.bind("<FocusOut>", solicitar_atualizacao_dashboard)
    data_fim_entry.bind("<Return>", solicitar_atualizacao_dashboard)


def _criar_info_relatorio(parent):
    global pasta_saida_label, pasta_label, progress, status_label

    info_card = _criar_card(parent, corner_radius=20)
    info_card.pack(fill="x", padx=18, pady=(0, 10))
    ui_refs["info_card"] = info_card

    pasta_saida_label = ctk.CTkLabel(
        info_card,
        text="",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
        wraplength=920,
    )
    pasta_saida_label.pack(fill="x", padx=18, pady=(14, 3))

    divisor = ctk.CTkFrame(info_card, height=1, fg_color=UI_THEME["divider"], corner_radius=1)
    divisor.pack(fill="x", padx=18, pady=(2, 6))
    ui_refs["info_divider"] = divisor

    pasta_label = ctk.CTkLabel(
        info_card,
        text="📄 Nenhum relatorio selecionado",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=12),
        wraplength=920,
    )
    pasta_label.pack(fill="x", padx=18, pady=(0, 8))

    progress = ctk.CTkProgressBar(
        info_card,
        height=10,
        corner_radius=6,
        fg_color=UI_THEME["progress_bg"],
        progress_color=UI_THEME["accent"],
    )
    progress.pack(fill="x", padx=18, pady=(0, 8))
    progress.set(0)

    status_label = ctk.CTkLabel(
        info_card,
        text="Relatorio:0 pagina(s) | Docs:0",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
    )
    status_label.pack(fill="x", padx=18, pady=(0, 14))

    atualizar_label_pasta_saida()
    atualizar_label_relatorio()


def _criar_tabs_menu(parent):
    global tab_indicator, tab_indicator_host, menu_feedback_label

    tabs_card = _criar_card(parent, corner_radius=20)
    tabs_card.pack(fill="x", padx=18, pady=(0, 10))
    tabs_card.grid_columnconfigure((0, 1, 2), weight=1)
    tab_indicator_host = tabs_card
    ui_refs["tabs_card"] = tabs_card

    acoes_menu_principal = [
        ("📂 Pasta do relatorio", selecionar_pasta_saida_relatorios),
        ("🔎 Buscar documento", abrir_busca_documentos),
        ("📈 Grafico faturamento", abrir_grafico_faturamento),
    ]
    acoes_menu_relatorios = [
        ("📄 Selecionar Relatorio", selecionar_relatorio),
        ("⬇️ Importar Relatorio", importar_relatorio_ui),
        ("🧾 Gerar Faturamento", gerar_excel),
        ("📤 Abrir Relatorio", abrir_relatorio),
        None,
        ("🚫 Relatorio Cancelados", abrir_relatorio_cancelados),
        ("💾 Exportar configuracoes", exportar_configuracoes_ui),
        ("📥 Importar configuracoes", importar_configuracoes_ui),
    ]
    acoes_menu_alteracoes = [
        ("🗓️ Alterar competencia", abrir_dialogo_alterar_competencia),
        ("🔁 Substituir documento", abrir_dialogo_substituir_documento),
        ("❌ Cancelar documento", abrir_dialogo_cancelar_documento),
        ("🏢 Declarar intercompany", abrir_dialogo_declarar_intercompany),
        ("📦 Declarar delta", abrir_dialogo_declarar_delta),
    ]

    tabs = [
        ("principal", "🏠 Principal", acoes_menu_principal),
        ("relatorios", "📄 Relatorios", acoes_menu_relatorios),
        ("alteracoes", "⚙️ Alteracoes", acoes_menu_alteracoes),
    ]

    for idx, (tab_id, titulo, opcoes) in enumerate(tabs):
        botao_tab = ctk.CTkButton(
            tabs_card,
            text=titulo,
            height=42,
            corner_radius=14,
            border_width=1,
            fg_color=TAB_STYLES["normal_fg"],
            border_color=UI_THEME["border"],
            hover_color=TAB_STYLES["hover_fg"],
            text_color=TAB_STYLES["normal_text"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )
        botao_tab.configure(command=lambda b=botao_tab, ops=opcoes, t=tab_id, ttl=titulo: abrir_menu_dropdown(b, ops, t, ttl))
        botao_tab.grid(row=0, column=idx, padx=8, pady=12, sticky="ew")
        tab_buttons[tab_id] = botao_tab
        _aplicar_microinteracao_botao(botao_tab, tab_id)

    tab_indicator = ctk.CTkFrame(
        tabs_card,
        fg_color=TAB_STYLES["active_fg"],
        height=3,
        corner_radius=3,
        border_width=0,
    )
    menu_feedback_label = ctk.CTkLabel(
        tabs_card,
        text="Escolha uma aba para acessar as acoes",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
    )
    menu_feedback_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 10))

    _configurar_tab_ativo("principal")


def _criar_bloco_metrica(parent, metric_key, titulo, valor_padrao):
    bg_key = f"metric_{metric_key}_bg" if f"metric_{metric_key}_bg" in UI_THEME else "metric_default_bg"
    title_key = f"metric_{metric_key}_title" if f"metric_{metric_key}_title" in UI_THEME else "metric_default_title"
    value_key = f"metric_{metric_key}_value" if f"metric_{metric_key}_value" in UI_THEME else "metric_default_value"

    card = ctk.CTkFrame(
        parent,
        fg_color=UI_THEME[bg_key],
        corner_radius=16,
        border_width=1,
        border_color=UI_THEME["border"],
    )
    card.configure(height=118 if metric_key in {"initial", "final"} else 106)
    card.pack_propagate(False)

    title_size = 12 if metric_key in {"initial", "final"} else 11
    value_size = 30 if metric_key in {"initial", "final"} else 26
    titulo_lbl = ctk.CTkLabel(
        card,
        text=titulo,
        font=ctk.CTkFont(family="Segoe UI", size=title_size),
        text_color=UI_THEME[title_key],
    )
    titulo_lbl.pack(pady=(12, 4))
    valor_label = ctk.CTkLabel(
        card,
        text=valor_padrao,
        font=ctk.CTkFont(family="Segoe UI", size=value_size, weight="bold"),
        text_color=UI_THEME[value_key],
    )
    valor_label.pack(pady=(0, 12))
    summary_metric_widgets[metric_key] = {
        "card": card,
        "title": titulo_lbl,
        "value": valor_label,
        "bg_key": bg_key,
        "title_key": title_key,
        "value_key": value_key,
    }
    return card, titulo_lbl, valor_label


def _criar_resumo_periodo(parent):
    global total_inicial_valor_label, total_final_valor_label, diferenca_label
    global total_documentos_label, nf_label, cte_label, cancelados_label, cancelados_card

    resumo_card = _criar_card(parent, corner_radius=20)
    resumo_card.pack(fill="x", padx=18, pady=(0, 16))
    ui_refs["resumo_card"] = resumo_card

    resumo_titulo = ctk.CTkLabel(
        resumo_card,
        text="Resumo do periodo",
        font=ctk.CTkFont(family="Segoe UI", size=21, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    resumo_titulo.pack(anchor="w", padx=20, pady=(16, 10))
    ui_refs["resumo_titulo"] = resumo_titulo

    linha_valores = ctk.CTkFrame(resumo_card, fg_color="transparent")
    linha_valores.pack(fill="x", padx=18, pady=(0, 8))
    linha_valores.grid_columnconfigure((0, 1), weight=1)

    inicial_card, _, total_inicial_valor_label = _criar_bloco_metrica(
        linha_valores, "initial", "Valor inicial", "R$ 0,00"
    )
    inicial_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

    final_card, _, total_final_valor_label = _criar_bloco_metrica(
        linha_valores, "final", "Valor final", "R$ 0,00"
    )
    final_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

    diferenca_label = ctk.CTkLabel(
        resumo_card,
        text="Diferenca referente a Impostos de NFS-e: R$ 0,00",
        font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        text_color=UI_THEME["text_secondary"],
    )
    diferenca_label.pack(fill="x", padx=20, pady=(6, 12))

    linha_metricas = ctk.CTkFrame(resumo_card, fg_color="transparent")
    linha_metricas.pack(fill="x", padx=18, pady=(0, 16))
    for c in range(4):
        linha_metricas.grid_columnconfigure(c, weight=1)

    docs_card, _, total_documentos_label = _criar_bloco_metrica(
        linha_metricas, "docs", "Total de documentos", "0"
    )
    docs_card.grid(row=0, column=0, padx=5, sticky="nsew")

    nf_card, _, nf_label = _criar_bloco_metrica(
        linha_metricas, "nf", "NF", "0"
    )
    nf_card.grid(row=0, column=1, padx=5, sticky="nsew")

    cte_card, _, cte_label = _criar_bloco_metrica(
        linha_metricas, "cte", "CTE", "0"
    )
    cte_card.grid(row=0, column=2, padx=5, sticky="nsew")

    cancelados_card, _, cancelados_label = _criar_bloco_metrica(
        linha_metricas, "cancelados", "Cancelados", "0"
    )
    cancelados_card.grid(row=0, column=3, padx=5, sticky="nsew")


def _criar_graficos_dashboard(parent):
    graficos_card = _criar_card(parent, corner_radius=20)
    graficos_card.pack(fill="x", padx=18, pady=(0, 16))
    ui_refs["graficos_card"] = graficos_card

    titulo = ctk.CTkLabel(
        graficos_card,
        text="Analise grafica",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.pack(anchor="w", padx=20, pady=(16, 4))
    ui_refs["graficos_titulo"] = titulo

    subtitulo = ctk.CTkLabel(
        graficos_card,
        text="Visao de faturamento por periodo e comparativo de documentos.",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color=UI_THEME["text_secondary"],
    )
    subtitulo.pack(anchor="w", padx=20, pady=(0, 10))
    ui_refs["graficos_subtitulo"] = subtitulo

    grade = ctk.CTkFrame(graficos_card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1, uniform="graf")

    card_faturamento = ctk.CTkFrame(
        grade,
        fg_color=UI_THEME["surface_alt"],
        corner_radius=16,
        border_width=1,
        border_color=UI_THEME["border"],
    )
    card_faturamento.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
    ui_refs["grafico_faturamento_card"] = card_faturamento

    titulo_fat = ctk.CTkLabel(
        card_faturamento,
        text="Faturamento por periodo",
        font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo_fat.pack(anchor="w", padx=12, pady=(10, 4))
    ui_refs["grafico_faturamento_titulo"] = titulo_fat

    host_fat = ctk.CTkFrame(card_faturamento, fg_color="transparent", height=230)
    host_fat.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    card_comp = ctk.CTkFrame(
        grade,
        fg_color=UI_THEME["surface_alt"],
        corner_radius=16,
        border_width=1,
        border_color=UI_THEME["border"],
    )
    card_comp.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
    ui_refs["grafico_comparativo_card"] = card_comp

    titulo_comp = ctk.CTkLabel(
        card_comp,
        text="NF x CTE x Cancelados",
        font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo_comp.pack(anchor="w", padx=12, pady=(10, 4))
    ui_refs["grafico_comparativo_titulo"] = titulo_comp

    host_comp = ctk.CTkFrame(card_comp, fg_color="transparent", height=230)
    host_comp.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    _liberar_graficos_dashboard()
    dashboard_chart_widgets["faturamento"] = {
        "host": host_fat,
        "fig": None,
        "canvas": None,
        "placeholder": None,
    }
    dashboard_chart_widgets["comparativo"] = {
        "host": host_comp,
        "fig": None,
        "canvas": None,
        "placeholder": None,
    }

    _mostrar_placeholder_grafico("faturamento", "Aguardando dados para montar o grafico.")
    _mostrar_placeholder_grafico("comparativo", "Aguardando dados para montar o grafico.")


def _criar_navegacao_telas(parent):
    global theme_toggle_button

    SCREEN_NAV_STYLES.update(
        {
            "normal_fg": UI_THEME["surface_alt"],
            "normal_text": UI_THEME["text_primary"],
            "hover_fg": UI_THEME["tab_hover"],
            "active_fg": UI_THEME["accent"],
            "active_hover_fg": UI_THEME["tab_active_hover"],
            "active_press_fg": UI_THEME["tab_active_press"],
            "active_text": UI_THEME["on_accent"],
        }
    )

    nav_card = _criar_card(parent, corner_radius=18)
    nav_card.pack(fill="x", padx=18, pady=(0, 10))
    nav_card.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
    nav_card.grid_columnconfigure(5, weight=0)
    ui_refs["screen_nav_card"] = nav_card

    telas = [
        ("dashboard", "📊 Dashboard"),
        ("relatorios", "📄 Relatorios"),
        ("faturamento", "🧾 Faturamento"),
        ("alteracoes", "⚙️ Alteracoes"),
        ("configuracoes", "🔧 Configuracoes"),
    ]

    for idx, (id_tela, titulo) in enumerate(telas):
        btn = ctk.CTkButton(
            nav_card,
            text=titulo,
            height=42,
            corner_radius=13,
            border_width=1,
            fg_color=SCREEN_NAV_STYLES["normal_fg"],
            border_color=UI_THEME["border"],
            hover_color=SCREEN_NAV_STYLES["hover_fg"],
            text_color=SCREEN_NAV_STYLES["normal_text"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda n=id_tela: _navegar_tela_com_feedback(n),
        )
        btn.grid(row=0, column=idx, padx=6, pady=10, sticky="ew")
        screen_nav_buttons[id_tela] = btn

    theme_toggle_button = ctk.CTkButton(
        nav_card,
        text=_texto_botao_tema(),
        width=124,
        height=38,
        corner_radius=12,
        border_width=1,
        fg_color=UI_THEME["surface_alt"],
        border_color=UI_THEME["border"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        command=alternar_tema_interface,
    )
    theme_toggle_button.grid(row=0, column=5, padx=(8, 10), pady=10, sticky="e")


def _criar_tela_dashboard(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    _criar_filtro_periodo(tela)
    _criar_info_relatorio(tela)
    _criar_resumo_periodo(tela)
    _criar_graficos_dashboard(tela)
    return tela


def _criar_tela_relatorios(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    card = _criar_card(tela, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(
        card,
        text="Relatorios",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 10))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)

    _criar_botao_acao(grade, "📄 Selecionar Relatorio", selecionar_relatorio).grid(row=0, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "⬇️ Importar Relatorio", importar_relatorio_ui).grid(row=0, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📤 Abrir Relatorio", abrir_relatorio).grid(row=1, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "🚫 Relatorio Cancelados", abrir_relatorio_cancelados).grid(row=1, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📈 Grafico faturamento", abrir_grafico_faturamento).grid(row=2, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "🔎 Buscar documento", abrir_busca_documentos).grid(row=2, column=1, padx=6, pady=6, sticky="ew")
    return tela


def _criar_tela_faturamento(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    card = _criar_card(tela, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(
        card,
        text="Faturamento",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 10))

    info = ctk.CTkLabel(
        card,
        text="Gere e abra os relatorios de faturamento com um clique.",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color=UI_THEME["text_secondary"],
    )
    info.pack(anchor="w", padx=18, pady=(0, 12))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)
    _criar_botao_acao(grade, "🧾 Gerar Faturamento", gerar_excel).grid(row=0, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📂 Pasta de saida", selecionar_pasta_saida_relatorios).grid(row=0, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📤 Abrir Relatorio", abrir_relatorio).grid(row=1, column=0, padx=6, pady=6, sticky="ew")
    return tela


def _criar_tela_alteracoes(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    card = _criar_card(tela, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(
        card,
        text="Alteracoes manuais",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 10))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)
    _criar_botao_acao(grade, "🗓️ Alterar competencia", abrir_dialogo_alterar_competencia).grid(row=0, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "🔁 Substituir documento", abrir_dialogo_substituir_documento).grid(row=0, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "❌ Cancelar documento", abrir_dialogo_cancelar_documento).grid(row=1, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "🏢 Declarar intercompany", abrir_dialogo_declarar_intercompany).grid(row=1, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📦 Declarar delta", abrir_dialogo_declarar_delta).grid(row=2, column=0, padx=6, pady=6, sticky="ew")
    return tela


def _criar_tela_configuracoes(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    card = _criar_card(tela, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(
        card,
        text="Configuracoes do sistema",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 10))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)
    _criar_botao_acao(grade, "🌗 Alternar tema", alternar_tema_interface).grid(row=0, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📂 Pasta de saida", selecionar_pasta_saida_relatorios).grid(row=0, column=1, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "💾 Exportar configuracoes", exportar_configuracoes_ui).grid(row=1, column=0, padx=6, pady=6, sticky="ew")
    _criar_botao_acao(grade, "📥 Importar configuracoes", importar_configuracoes_ui).grid(row=1, column=1, padx=6, pady=6, sticky="ew")
    return tela


def construir_tela_principal():
    global container_scroll, main_frame

    app.screens = {}
    app.current_screen = ""
    screen_nav_buttons.clear()
    action_buttons.clear()
    tab_buttons.clear()
    tab_hover_state.clear()
    summary_metric_widgets.clear()
    _liberar_graficos_dashboard()
    ui_refs.clear()

    container_scroll = ctk.CTkScrollableFrame(
        app,
        corner_radius=0,
        fg_color=UI_THEME["app_bg"],
        scrollbar_button_color=UI_THEME["scroll_btn"],
        scrollbar_button_hover_color=UI_THEME["scroll_btn_hover"],
    )
    container_scroll.pack(fill="both", expand=True)
    ui_refs["container_scroll"] = container_scroll

    main_frame = ctk.CTkFrame(container_scroll, fg_color=UI_THEME["app_bg"])
    main_frame.pack(fill="both", expand=True, padx=6, pady=(2, 6))
    ui_refs["main_frame"] = main_frame

    _aplicar_logo_watermark(main_frame)
    _criar_navegacao_telas(main_frame)

    screen_host = ctk.CTkFrame(main_frame, fg_color=UI_THEME["app_bg"])
    screen_host.pack(fill="both", expand=True)
    app.screen_host = screen_host
    ui_refs["screen_host"] = screen_host

    dashboard_frame = _criar_tela_dashboard(screen_host)
    relatorios_frame = _criar_tela_relatorios(screen_host)
    faturamento_frame = _criar_tela_faturamento(screen_host)
    alteracoes_frame = _criar_tela_alteracoes(screen_host)
    configuracoes_frame = _criar_tela_configuracoes(screen_host)

    app.registrar_tela("dashboard", dashboard_frame)
    app.registrar_tela("relatorios", relatorios_frame)
    app.registrar_tela("faturamento", faturamento_frame)
    app.registrar_tela("alteracoes", alteracoes_frame)
    app.registrar_tela("configuracoes", configuracoes_frame)
    app.mostrar_tela("dashboard")


construir_tela_principal()
aplicar_tema_interface()

# Ajusta a janela ao novo layout dashboard.
app.update_idletasks()
largura_ideal = min(max(900, app.winfo_reqwidth()), largura_tela - 10)
altura_ideal = min(max(680, app.winfo_reqheight()), altura_tela - 10)
centralizar_janela(app, largura_ideal, altura_ideal)

app.after(120, atualizar_dashboard)

try:
    app.mainloop()
finally:
    liberar_lock_instancia()


