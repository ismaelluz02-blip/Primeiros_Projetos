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

from src.utils import (
    MESES,
    valor_brasileiro,
    formatar_moeda_brl,
    formatar_moeda_brl_exata,
    parse_valor_monetario,
    normalizar_texto,
    _numero_para_texto,
    _extrair_ano_data_emissao,
    _normalizar_numero_original_nf,
    _coletar_numero_original_para_match,
    _numero_documento_exibicao,
    _chave_documento_compativel,
    competencia_por_data,
    periodo_padrao_mes_atual,
    obter_periodo_padrao_dashboard,
    obter_periodo_padrao_relatorios,
    ler_data_filtro,
    _obter_periodo_por_entries,
)

import src.config as _cfg
from src.banco import (
    obter_conexao_banco,
    obter_configuracao,
    salvar_configuracao,
    iniciar_banco,
)
from src.documentos import (
    salvar_documento,
    alterar_competencia_documento,
    _normalizar_modalidade_frete,
    _coletar_ids_documentos_por_numero,
    _coletar_ids_documentos_para_frete,
    atualizar_modalidade_frete_documento,
    declarar_documento_frete,
    salvar_alteracao_frete_manual,
    declarar_intercompany,
    declarar_delta,
    declarar_spot,
    registrar_substituicao,
    desfazer_substituicao,
    cancelar_documento,
    desfazer_cancelamento_documento,
    _buscar_documento_existente_sync,
    register_on_change as _register_doc_on_change,
)
from src.sync import (
    exportar_configuracoes_json,
    importar_configuracoes_json,
    _listar_documentos_alterados_para_sync,
)

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
SYNC_DOCUMENT_FIELDS = [
    "tipo",
    "numero",
    "numero_original",
    "data_emissao",
    "valor_inicial",
    "valor_final",
    "frete",
    "status",
    "competencia",
    "valor_inicial_original",
    "valor_final_original",
    "status_original",
    "cancelado_manual",
    "competencia_manual",
    "frete_manual",
    "frete_revisado_manual",
]
DB_PATH = os.path.join(APP_DATA_DIR, "faturamento.db")
LOCK_PATH = os.path.join(APP_DATA_DIR, ".sistema_faturamento.lock")
LEGACY_DB_PATH = os.path.join(APP_DIR, "faturamento.db")
LOGO_PATH = os.path.join(APP_DIR, "logo.png")
APP_USER_MODEL_ID = "horizonte.logistica.sistema.faturamento"
FALLBACK_APP_DATA_DIR = os.path.join(BASE_DIR, "_dados_app")


def _diretorio_gravavel(caminho_dir):
    try:
        os.makedirs(caminho_dir, exist_ok=True)
        arquivo_teste = os.path.join(caminho_dir, f".write_test_{os.getpid()}_{time.time_ns()}.tmp")
        with open(arquivo_teste, "w", encoding="utf-8") as f:
            f.write("ok")
        try:
            os.remove(arquivo_teste)
        except OSError:
            pass
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


def configurar_cache_matplotlib():
    cache_dir = os.path.join(APP_DATA_DIR, "matplotlib")
    if _diretorio_gravavel(cache_dir):
        os.environ["MPLCONFIGDIR"] = cache_dir


configurar_cache_matplotlib()


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

    fontes_banco = []
    for origem in (
        os.path.join(DEFAULT_APP_DATA_DIR, "faturamento.db"),
        LEGACY_DB_PATH,
    ):
        origem_abs = os.path.abspath(origem)
        if origem_abs == os.path.abspath(DB_PATH):
            continue
        if origem_abs not in fontes_banco:
            fontes_banco.append(origem_abs)

    # Migra automaticamente um banco saudável para a pasta atual na primeira execução
    # ou quando o fallback local estiver vazio.
    if not os.path.exists(DB_PATH):
        for origem_db in fontes_banco:
            if not os.path.exists(origem_db):
                continue
            try:
                shutil.copy2(origem_db, DB_PATH)
                break
            except OSError:
                continue
        return

    docs_novo = _contar_documentos(DB_PATH)
    if docs_novo == 0:
        for origem_db in fontes_banco:
            docs_origem = _contar_documentos(origem_db)
            if isinstance(docs_origem, int) and docs_origem > 0:
                try:
                    shutil.copy2(origem_db, DB_PATH)
                    break
                except OSError:
                    continue


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
# BANCO  →  src/banco.py
# ------------------------


# _sqlite_db_valido / _tentar_recuperar_banco / obter_conexao_banco
# obter_configuracao / salvar_configuracao / iniciar_banco  →  src/banco.py


# ------------------------
# VARIAVEIS
# ------------------------

relatorio_selecionado = ""
pasta_relatorios_saida = RELATORIOS_DIR
caminho_relatorio_atual = ""
dados_importados = None
relatorio_carregado = False
relatorios_total_intervalo = None
dashboard_data_inicio = ""
dashboard_data_fim = ""
relatorio_data_inicio = ""
relatorio_data_fim = ""

paginas_lidas = 0
docs_encontrados = 0
dashboard_update_after_id = None
dashboard_update_running = False
scroll_refresh_after_id = None
scroll_dragging = False
watermark_hidden_by_scroll = False
ui_animations_paused = False

WATERMARK_POS = {
    "relx": 0.5,
    "rely": 0.52,
    "anchor": "center",
}


# ------------------------
# UTIL  →  src/utils.py
# ------------------------


def obter_periodo_dashboard(silencioso=False):
    global dashboard_data_inicio, dashboard_data_fim
    data_inicial, data_final = _obter_periodo_por_entries(
        globals().get("dashboard_data_inicio_entry"),
        globals().get("dashboard_data_fim_entry"),
        contexto="Filtro do Dashboard",
        silencioso=silencioso,
    )
    if data_inicial and data_final:
        dashboard_data_inicio = data_inicial.strftime("%d/%m/%Y")
        dashboard_data_fim = data_final.strftime("%d/%m/%Y")
    return data_inicial, data_final


def obter_periodo_relatorios(silencioso=False):
    global relatorio_data_inicio, relatorio_data_fim
    data_inicial, data_final = _obter_periodo_por_entries(
        globals().get("relatorio_data_inicio_entry"),
        globals().get("relatorio_data_fim_entry"),
        contexto="Filtro de Relatórios",
        silencioso=silencioso,
    )
    if data_inicial and data_final:
        relatorio_data_inicio = data_inicial.strftime("%d/%m/%Y")
        relatorio_data_fim = data_final.strftime("%d/%m/%Y")
    return data_inicial, data_final


def inicializar_filtros_dashboard():
    global dashboard_data_inicio, dashboard_data_fim
    data_inicial, data_final = obter_periodo_padrao_dashboard()
    dashboard_data_inicio = data_inicial.strftime("%d/%m/%Y")
    dashboard_data_fim = data_final.strftime("%d/%m/%Y")

    entry_inicio = globals().get("dashboard_data_inicio_entry")
    entry_fim = globals().get("dashboard_data_fim_entry")
    if entry_inicio is not None and entry_fim is not None:
        entry_inicio.delete(0, "end")
        entry_inicio.insert(0, dashboard_data_inicio)
        entry_fim.delete(0, "end")
        entry_fim.insert(0, dashboard_data_fim)


def inicializar_filtros_relatorios():
    global relatorio_data_inicio, relatorio_data_fim
    data_inicial, data_final = obter_periodo_padrao_relatorios()
    relatorio_data_inicio = data_inicial.strftime("%d/%m/%Y")
    relatorio_data_fim = data_final.strftime("%d/%m/%Y")

    entry_inicio = globals().get("relatorio_data_inicio_entry")
    entry_fim = globals().get("relatorio_data_fim_entry")
    if entry_inicio is not None and entry_fim is not None:
        entry_inicio.delete(0, "end")
        entry_inicio.insert(0, relatorio_data_inicio)
        entry_fim.delete(0, "end")
        entry_fim.insert(0, relatorio_data_fim)


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
            _atualizar_status_relatorios_ui()
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


def aplicar_filtro_dashboard(_event=None):
    try:
        data_inicial, data_final = obter_periodo_dashboard(silencioso=False)
    except ValueError as exc:
        messagebox.showwarning("Filtro do Dashboard", str(exc))
        return

    if data_inicial is None or data_final is None:
        messagebox.showwarning("Filtro do Dashboard", "Preencha as datas do Dashboard.")
        return

    if data_inicial > data_final:
        messagebox.showwarning("Filtro do Dashboard", "A data inicial não pode ser maior que a data final.")
        return

    atualizar_dashboard()


def aplicar_filtro_relatorios(_event=None, mensagem_sucesso=True):
    global relatorios_total_intervalo
    try:
        data_inicial, data_final = obter_periodo_relatorios(silencioso=False)
    except ValueError as exc:
        _atualizar_status_relatorios_ui(str(exc), tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", str(exc))
        return False

    if data_inicial is None or data_final is None:
        mensagem = "Preencha as datas da aba Relatórios."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", mensagem)
        return False

    if data_inicial > data_final:
        mensagem = "A data inicial não pode ser maior que a data final."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", mensagem)
        return False

    try:
        df_filtrado, msg_erro = _obter_dataframe_relatorio_filtrado(
            data_inicial,
            data_final,
            docs_df_base=_obter_documentos_em_memoria(force=False),
        )
    except Exception:
        df_filtrado, msg_erro = pd.DataFrame(), ""

    if msg_erro:
        _atualizar_status_relatorios_ui(msg_erro, tipo="erro", total_registros=0)
        if mensagem_sucesso:
            messagebox.showwarning("Filtro de Relatórios", msg_erro)
        return False

    total_registros = int(len(df_filtrado)) if isinstance(df_filtrado, pd.DataFrame) else 0
    relatorios_total_intervalo = total_registros
    if mensagem_sucesso:
        _atualizar_status_relatorios_ui(
            "Período de Relatórios atualizado",
            tipo="ok",
            total_registros=total_registros,
        )
    else:
        _atualizar_status_relatorios_ui(total_registros=total_registros)
    return True


def _forcar_redesenho_pos_scroll():
    # Mitiga artefatos visuais durante/apos arraste da barra lateral no Windows.
    canvas_scroll = globals().get("ui_refs", {}).get("scroll_canvas")
    if canvas_scroll is None:
        scroll_frame = globals().get("container_scroll")
        canvas_scroll = getattr(scroll_frame, "_parent_canvas", None) if scroll_frame is not None else None
    if canvas_scroll is not None:
        try:
            canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
            canvas_scroll.update_idletasks()
        except Exception:
            pass

    for cfg in globals().get("dashboard_chart_widgets", {}).values():
        canvas = cfg.get("canvas")
        if canvas is None:
            continue
        try:
            canvas.draw_idle()
        except Exception:
            try:
                canvas.draw()
            except Exception:
                pass
    try:
        if "app" in globals() and app.winfo_exists():
            app.update_idletasks()
    except Exception:
        pass


def _ocultar_watermark_durante_scroll():
    global watermark_hidden_by_scroll

    lbl = globals().get("ui_refs", {}).get("logo_watermark_label")
    if lbl is None or not lbl.winfo_exists():
        return
    if watermark_hidden_by_scroll:
        return
    try:
        lbl.place_forget()
        watermark_hidden_by_scroll = True
    except Exception:
        pass


def _agendar_restauro_visual_scroll(delay_ms=110):
    global scroll_refresh_after_id

    try:
        if "app" not in globals() or not app.winfo_exists():
            return
        if scroll_refresh_after_id:
            try:
                app.after_cancel(scroll_refresh_after_id)
            except Exception:
                pass
            scroll_refresh_after_id = None
        scroll_refresh_after_id = app.after(max(40, int(delay_ms)), _restaurar_visual_apos_scroll)
    except Exception:
        pass


def _restaurar_visual_apos_scroll():
    global scroll_refresh_after_id, watermark_hidden_by_scroll, ui_animations_paused

    scroll_refresh_after_id = None
    if scroll_dragging:
        _agendar_restauro_visual_scroll(140)
        return

    if watermark_hidden_by_scroll:
        container = globals().get("ui_refs", {}).get("main_frame")
        if container is not None and container.winfo_exists():
            try:
                _aplicar_logo_watermark(container)
            except Exception:
                pass
        watermark_hidden_by_scroll = False

    ui_animations_paused = False
    _forcar_redesenho_pos_scroll()


def _on_scrollbar_press(_event=None):
    global scroll_dragging, ui_animations_paused
    scroll_dragging = True
    ui_animations_paused = True
    _ocultar_watermark_durante_scroll()
    _agendar_restauro_visual_scroll(180)


def _on_scrollbar_drag(_event=None):
    _ocultar_watermark_durante_scroll()
    _agendar_restauro_visual_scroll(180)


def _on_scrollbar_release(_event=None):
    global scroll_dragging
    scroll_dragging = False
    _agendar_restauro_visual_scroll(70)


def _configurar_mitigacao_artefato_scroll(scroll_frame):
    if scroll_frame is None:
        return
    if getattr(scroll_frame, "_artefato_scroll_configurado", False):
        return

    barra = getattr(scroll_frame, "_scrollbar", None)
    canvas = getattr(scroll_frame, "_parent_canvas", None)
    if canvas is not None and canvas.winfo_exists():
        try:
            # Menos steps por arraste reduz "ghosting" no Windows.
            canvas.configure(yscrollincrement=5)
        except Exception:
            pass

    if barra is not None and barra.winfo_exists():
        comando_original = getattr(canvas, "yview", None) if canvas is not None else None
        if callable(comando_original):
            def _comando_scrollbar(*args):
                _on_scrollbar_press()
                try:
                    comando_original(*args)
                finally:
                    _on_scrollbar_drag()
                    _forcar_redesenho_pos_scroll()

            try:
                barra.configure(command=_comando_scrollbar)
            except Exception:
                pass

        barra.bind("<ButtonPress-1>", _on_scrollbar_press, add="+")
        barra.bind("<B1-Motion>", _on_scrollbar_drag, add="+")
        barra.bind("<ButtonRelease-1>", _on_scrollbar_release, add="+")

    if canvas is not None and canvas.winfo_exists():
        canvas.bind("<ButtonRelease-1>", _on_scrollbar_release, add="+")
    try:
        if "app" in globals() and app.winfo_exists():
            app.bind_all("<ButtonRelease-1>", _on_scrollbar_release, add="+")
    except Exception:
        pass

    setattr(scroll_frame, "_artefato_scroll_configurado", True)


def obter_pasta_saida_relatorios():
    global pasta_relatorios_saida
    if not pasta_relatorios_saida:
        pasta_relatorios_saida = RELATORIOS_DIR
    os.makedirs(pasta_relatorios_saida, exist_ok=True)
    return pasta_relatorios_saida


def atualizar_label_pasta_saida():
    pasta_txt = f"Pasta de saída: {obter_pasta_saida_relatorios()}"
    if "pasta_saida_label" in globals():
        pasta_saida_label.configure(text=pasta_txt)
    lbl_saida = None
    if "ui_refs" in globals():
        lbl_saida = ui_refs.get("relatorios_saida_label")
    try:
        if lbl_saida is not None and lbl_saida.winfo_exists():
            if "UI_THEME" in globals():
                lbl_saida.configure(text=pasta_txt, text_color=UI_THEME["text_secondary"])
            else:
                lbl_saida.configure(text=pasta_txt)
    except Exception:
        pass


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
        title="Selecionar pasta para salvar o relatório",
        initialdir=obter_pasta_saida_relatorios(),
        mustexist=True,
    )
    if not caminho:
        return

    pasta_relatorios_saida = caminho
    salvar_configuracao("pasta_relatorios_saida", pasta_relatorios_saida)
    obter_pasta_saida_relatorios()
    atualizar_label_pasta_saida()
    messagebox.showinfo("Pasta de saída", f"Relatórios serão salvos em:\n{pasta_relatorios_saida}")


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
    if "pasta_label" in globals() and pasta_label is not None:
        if relatorio_selecionado:
            nome_arquivo = os.path.basename(relatorio_selecionado)
            pasta_label.configure(text=f"Arquivo: {nome_arquivo}")
        else:
            pasta_label.configure(text="Arquivo: Nenhum selecionado")
    _atualizar_status_relatorios_ui()


def definir_relatorio_selecionado(caminho, persistir=False):
    global relatorio_selecionado, caminho_relatorio_atual
    relatorio_selecionado = os.path.normpath(caminho) if caminho else ""
    caminho_relatorio_atual = relatorio_selecionado
    if persistir and relatorio_selecionado:
        salvar_ultimo_relatorio(relatorio_selecionado)
    atualizar_label_relatorio()
    return relatorio_selecionado


def _periodo_dashboard_relatorio_texto():
    try:
        ini, fim = obter_periodo_relatorios(silencioso=True)
        if ini is None and relatorio_data_inicio and relatorio_data_fim:
            ini = ler_data_filtro(relatorio_data_inicio, "Data inicial")
            fim = ler_data_filtro(relatorio_data_fim, "Data final")
        if ini and fim:
            meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
            ini_txt = f"{meses_abrev[ini.month - 1]}/{ini.strftime('%y')}"
            fim_txt = f"{meses_abrev[fim.month - 1]}/{fim.strftime('%y')}"
            return f"{ini_txt} a {fim_txt}"
    except Exception:
        pass
    return "Período indefinido"


def _atualizar_status_relatorios_ui(mensagem=None, tipo="ok", total_registros=None):
    global relatorios_total_intervalo
    if "UI_THEME" not in globals() or "ui_refs" not in globals():
        return

    lbl_status = ui_refs.get("relatorios_status_label")
    lbl_arquivo = ui_refs.get("relatorios_arquivo_label")
    lbl_periodo = ui_refs.get("relatorios_periodo_label")
    lbl_registros = ui_refs.get("relatorios_registros_label")
    lbl_saida = ui_refs.get("relatorios_saida_label")

    if relatorio_selecionado and relatorio_carregado:
        nome_arquivo = os.path.basename(relatorio_selecionado)
        arquivo_txt = f"Arquivo: {nome_arquivo}"
        status_padrao = "Relatório carregado com sucesso"
    elif relatorio_selecionado:
        nome_arquivo = os.path.basename(relatorio_selecionado)
        arquivo_txt = f"Arquivo: {nome_arquivo}"
        status_padrao = "Arquivo selecionado"
    else:
        arquivo_txt = "Arquivo: Nenhum selecionado"
        status_padrao = "Nenhum relatório carregado"
        if mensagem is None:
            tipo = "neutral"

    status_txt = mensagem or status_padrao
    periodo_txt = f"Período aplicado: {_periodo_dashboard_relatorio_texto()}"
    if total_registros is not None:
        relatorios_total_intervalo = int(total_registros)
    registros_txt = (
        f"Registros no período: {relatorios_total_intervalo}"
        if isinstance(relatorios_total_intervalo, int)
        else "Registros no período: -"
    )

    cor_status = UI_THEME["success_text"] if tipo == "ok" else (
        UI_THEME["danger_text"] if tipo == "erro" else UI_THEME["text_secondary"]
    )

    _safe_config(lbl_status, text=status_txt, text_color=cor_status)
    _safe_config(lbl_arquivo, text=arquivo_txt, text_color=UI_THEME["text_primary"])
    _safe_config(lbl_saida, text=f"Pasta de saída: {obter_pasta_saida_relatorios()}", text_color=UI_THEME["text_secondary"])
    _safe_config(lbl_periodo, text=periodo_txt, text_color=UI_THEME["text_secondary"])
    _safe_config(lbl_registros, text=registros_txt, text_color=UI_THEME["text_secondary"])


def carregar_ultimo_relatorio():
    global relatorio_carregado, dados_importados, caminho_relatorio_atual
    caminho = _resolver_ultimo_relatorio_salvo()
    if caminho:
        definir_relatorio_selecionado(caminho, persistir=False)
    caminho_relatorio_atual = relatorio_selecionado
    try:
        conn = obter_conexao_banco()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM documentos")
        relatorio_carregado = int(cursor.fetchone()[0] or 0) > 0
        conn.close()
    except Exception:
        relatorio_carregado = False
    if relatorio_carregado:
        dados_importados = {"origem": "banco_local", "carregado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
        try:
            _obter_documentos_em_memoria(force=True)
        except Exception:
            pass


def _carregar_documentos_para_memoria():
    conn = obter_conexao_banco()
    try:
        df = pd.read_sql_query("SELECT * FROM documentos", conn)
    finally:
        conn.close()
    return df


def _obter_documentos_em_memoria(force=False):
    global dados_importados

    if not isinstance(dados_importados, dict):
        dados_importados = {}

    if not force:
        df_mem = dados_importados.get("df_documentos")
        if isinstance(df_mem, pd.DataFrame):
            return df_mem.copy()

    df_mem = _carregar_documentos_para_memoria()
    dados_importados["df_documentos"] = df_mem
    dados_importados["memoria_atualizada_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    return df_mem.copy()


def _atualizar_cache_documentos_pos_alteracao():
    global dados_importados, relatorio_carregado
    try:
        df_mem = _obter_documentos_em_memoria(force=True)
        relatorio_carregado = not df_mem.empty
        if not isinstance(dados_importados, dict):
            dados_importados = {}
        dados_importados["total_documentos"] = int(len(df_mem))
        dados_importados["memoria_atualizada_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        pass


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
        print(f"Aviso: não foi possível carregar o ícone .ico - {e}")

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


def importar_relatorio_consolidado(caminho_pdf):
    global docs_encontrados, paginas_lidas

    docs_encontrados = 0
    paginas_lidas = 0

    vistos = set()
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
                status_label.configure(text=f"Importando relatório: página {idx}/{total_pag} | Docs: {docs_encontrados}")
                manter_interface_responsiva()
    except Exception as exc:
        messagebox.showerror("Relatório", f"Falha ao importar relatório.\n\n{exc}")
        return

    messagebox.showinfo("Relatório", f"Importação concluída. {docs_encontrados} documento(s) carregado(s).")

    atualizar_dashboard()
    _atualizar_status_relatorios_ui("Relatório carregado com sucesso", tipo="ok")


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
            "Relatório",
            "Falha ao abrir planilha. Para arquivo .xls, instale 'xlrd' ou exporte para .xlsx.\n\n" + str(exc),
        )
        return
    except Exception as exc:
        messagebox.showerror("Relatório", f"Falha ao abrir planilha.\n\n{exc}")
        return

    if not planilhas:
        messagebox.showwarning("Relatório", "A planilha não possui abas para importar.")
        return

    # Limpa importacoes automaticas anteriores para evitar residuos no novo relatorio.
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    cursor.execute(
        """
        DELETE FROM documentos
        WHERE COALESCE(cancelado_manual,0)=0
          AND COALESCE(competencia_manual,0)=0
          AND COALESCE(frete_manual,0)=0
          AND UPPER(COALESCE(status,'')) NOT LIKE '%CANCELADO%'
          AND UPPER(COALESCE(status,'')) NOT LIKE '%SUBSTITUIDO%'
          AND UPPER(COALESCE(status,'')) NOT LIKE 'DOCUMENTO SUBSTITUINDO%'
        """
    )
    conn.commit()
    conn.close()

    vistos = {}

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
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
        messagebox.showwarning("Relatório", "Não há linhas com dados na planilha selecionada.")
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
                numero = int(numero_txt)
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

    resumo = f"Importação concluída. {docs_encontrados} documento(s) carregado(s)."
    if erros:
        resumo += "\n\nAvisos:\n- " + "\n- ".join(erros[:5])
        if len(erros) > 5:
            resumo += f"\n- ... e mais {len(erros) - 5} aviso(s)."

    messagebox.showinfo("Relatório", resumo)

    atualizar_dashboard()
    _atualizar_status_relatorios_ui("Relatório carregado com sucesso", tipo="ok")


def _importar_relatorio_por_caminho(caminho):
    global relatorio_carregado, dados_importados, caminho_relatorio_atual, relatorios_total_intervalo
    ext = os.path.splitext(caminho)[1].lower()
    try:
        if ext == ".pdf":
            importar_relatorio_consolidado(caminho)
        elif ext in {".xlsx", ".xls"}:
            importar_relatorio_planilha(caminho)
        else:
            messagebox.showerror("Relatório", "Formato de arquivo não suportado. Use PDF, XLSX ou XLS.")
            _atualizar_status_relatorios_ui("Falha ao carregar relatório", tipo="erro")
            return False
    except Exception as exc:
        relatorio_carregado = False
        dados_importados = None
        relatorios_total_intervalo = None
        _atualizar_status_relatorios_ui("Falha ao carregar relatório", tipo="erro")
        messagebox.showerror("Relatório", f"Falha ao importar relatório.\n\n{exc}")
        return False

    caminho_relatorio_atual = os.path.normpath(caminho)
    relatorio_carregado = True
    relatorios_total_intervalo = None
    dados_importados = {
        "caminho": caminho_relatorio_atual,
        "fonte": ext,
        "importado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    try:
        dados_importados["df_documentos"] = _obter_documentos_em_memoria(force=True)
        dados_importados["total_documentos"] = int(len(dados_importados["df_documentos"]))
    except Exception:
        pass
    _atualizar_status_relatorios_ui("Relatório carregado com sucesso", tipo="ok")
    return True


def importar_relatorio_automaticamente(caminho):
    return _importar_relatorio_por_caminho(caminho)


def selecionar_relatorio():
    diretorio_inicial = (
        os.path.dirname(relatorio_selecionado)
        if relatorio_selecionado
        else obter_configuracao("ultimo_relatorio_diretorio", RELATORIOS_DIR).strip() or RELATORIOS_DIR
    )
    caminho = filedialog.askopenfilename(
        title="Selecionar relatório consolidado",
        initialdir=diretorio_inicial,
        filetypes=[
            ("Relatórios", "*.xlsx *.xls *.pdf"),
            ("Excel", "*.xlsx *.xls"),
            ("PDF", "*.pdf"),
            ("Todos os arquivos", "*.*"),
        ],
    )
    if caminho:
        caminho = definir_relatorio_selecionado(caminho, persistir=True)
        importar_relatorio_automaticamente(caminho)
        return caminho
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
                messagebox.showwarning("Relatório", "Nenhum relatório selecionado para importação.")
                return

        definir_relatorio_selecionado(caminho, persistir=True)
        _importar_relatorio_por_caminho(caminho)
    except Exception as exc:
        try:
            status_label.configure(text="Falha ao importar relatório.")
        except Exception:
            pass
        _atualizar_status_relatorios_ui("Falha ao carregar relatório", tipo="erro")
        messagebox.showerror("Relatório", f"Falha ao importar relatório.\n\n{exc}")


# ─── src/sync.py ────────────────────────────────────────────────────────────
# _documento_possui_alteracao_manual, _listar_documentos_alterados_para_sync,
# exportar_configuracoes_json, _extrair_documentos_payload_sync,
# importar_configuracoes_json,
# _to_float, _to_optional_float, _to_manual_flag, _normalizar_data_emissao_sync
# ─────────────────────────────────────────────────────────────────────────────

# _coletar_numero_original_para_match  →  src/utils.py


def exportar_configuracoes_ui():
    try:
        docs = _listar_documentos_alterados_para_sync()
        if not docs:
            messagebox.showwarning("Configurações", "Não há alterações manuais para exportar.")
            return

        diretorio_inicial = (
            obter_configuracao("ultimo_sync_diretorio", obter_pasta_saida_relatorios()).strip()
            or obter_pasta_saida_relatorios()
        )
        nome_padrao = f"configuracoes_faturamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        caminho = filedialog.asksaveasfilename(
            title="Exportar configurações",
            defaultextension=".json",
            initialdir=diretorio_inicial,
            initialfile=nome_padrao,
            filetypes=[("Arquivo JSON", "*.json"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        total = exportar_configuracoes_json(caminho)
        salvar_configuracao("ultimo_sync_diretorio", os.path.dirname(caminho))
        messagebox.showinfo(
            "Configurações",
            "Configurações exportadas com sucesso.\n\n"
            f"Documentos exportados: {total}\n"
            f"Arquivo: {caminho}",
        )
    except Exception as exc:
        messagebox.showerror("Configurações", f"Falha ao exportar configurações.\n\n{exc}")


def importar_configuracoes_ui():
    try:
        diretorio_inicial = (
            obter_configuracao("ultimo_sync_diretorio", obter_pasta_saida_relatorios()).strip()
            or obter_pasta_saida_relatorios()
        )
        caminho = filedialog.askopenfilename(
            title="Importar configurações",
            initialdir=diretorio_inicial,
            filetypes=[("Arquivo JSON", "*.json"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        salvar_configuracao("ultimo_sync_diretorio", os.path.dirname(caminho))
        resumo = importar_configuracoes_json(caminho)
        _atualizar_cache_documentos_pos_alteracao()
        atualizar_dashboard()
        _atualizar_status_relatorios_ui("Configurações importadas com sucesso", tipo="ok")

        mensagem = (
            "Importação concluída com sucesso.\n\n"
            f"Inseridos: {resumo['inseridos']}\n"
            f"Atualizados: {resumo['atualizados']}\n"
            f"Ignorados: {resumo['ignorados']}\n"
            f"Erros: {len(resumo['erros'])}"
        )
        if resumo["erros"]:
            mensagem += "\n\nPrimeiros erros:\n- " + "\n- ".join(resumo["erros"][:5])
            if len(resumo["erros"]) > 5:
                mensagem += f"\n- ... e mais {len(resumo['erros']) - 5} erro(s)."

        messagebox.showinfo("Configurações", mensagem)
    except Exception as exc:
        messagebox.showerror("Configurações", f"Falha ao importar configurações.\n\n{exc}")



# ─── src/documentos.py ──────────────────────────────────────────────────────
# _buscar_documento_existente_sync, salvar_documento, alterar_competencia_documento,
# _normalizar_modalidade_frete, _coletar_ids_documentos_*, atualizar_modalidade_frete_documento,
# declarar_intercompany/delta/spot, registrar_substituicao, desfazer_substituicao,
# cancelar_documento, desfazer_cancelamento_documento
# ─────────────────────────────────────────────────────────────────────────────


def _obter_dataframe_relatorio_filtrado(data_inicial, data_final, docs_df_base=None):
    if isinstance(docs_df_base, pd.DataFrame):
        df = docs_df_base.copy()
    else:
        try:
            df = _obter_documentos_em_memoria(force=False)
        except Exception:
            conn = obter_conexao_banco()
            df = pd.read_sql_query("SELECT * FROM documentos", conn)
            conn.close()

    if df.empty:
        return pd.DataFrame(), "Nenhum documento encontrado no banco."

    df["data_emissao"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao"])
    if df.empty:
        return pd.DataFrame(), "Não há datas válidas para gerar o relatório."

    df["numero"] = pd.to_numeric(df["numero"], errors="coerce")
    df = df.dropna(subset=["numero"])
    if df.empty:
        return pd.DataFrame(), "Não há números de documento válidos para gerar o relatório."
    df["numero"] = df["numero"].astype(int)

    def competencia_para_data(comp_str):
        try:
            partes = str(comp_str).lower().split("/")
            if len(partes) == 2:
                mes_nome = partes[0].strip()
                ano_str = partes[1].strip()
                ano = int(ano_str)
                mes_idx = MESES.index(mes_nome) + 1
                return datetime(ano, mes_idx, 1)
        except Exception:
            pass
        return None

    df["data_competencia"] = df["competencia"].apply(competencia_para_data)
    df = df.dropna(subset=["data_competencia"])
    df = df[(df["data_competencia"] >= data_inicial) & (df["data_competencia"] <= data_final)].copy()

    if df.empty:
        return pd.DataFrame(), ""

    df["numero_original_num"] = pd.to_numeric(df.get("numero_original"), errors="coerce")
    df["numero_exibicao"] = df.apply(
        lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
        axis=1,
    )
    df["chave_documento"] = df.apply(
        lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")
    return df, ""


def abrir_relatorio():
    global relatorio_carregado, dados_importados

    base_disponivel = bool(relatorio_carregado)
    if not relatorio_selecionado:
        caminho_salvo = _resolver_ultimo_relatorio_salvo()
        if caminho_salvo:
            definir_relatorio_selecionado(caminho_salvo, persistir=True)

    if not relatorio_selecionado and not base_disponivel:
        mensagem = "Selecione um relatório para importar automaticamente antes de abrir."
        _atualizar_status_relatorios_ui(mensagem, tipo="neutral", total_registros=None)
        messagebox.showwarning("Relatório", mensagem)
        return

    try:
        data_inicial, data_final = obter_periodo_relatorios(silencioso=False)
    except ValueError as exc:
        _atualizar_status_relatorios_ui(str(exc), tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", str(exc))
        return

    if data_inicial is None or data_final is None:
        mensagem = "Preencha o período da aba Relatórios para abrir o arquivo filtrado."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", mensagem)
        return

    if data_inicial > data_final:
        mensagem = "A data inicial não pode ser maior que a data final."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Filtro de Relatórios", mensagem)
        return

    resultado = gerar_excel(data_inicial=data_inicial, data_final=data_final, exibir_mensagem=False)
    if not resultado.get("ok"):
        _atualizar_status_relatorios_ui(
            resultado.get("mensagem", "Não foi possível abrir o relatório."),
            tipo="erro",
            total_registros=0,
        )
        messagebox.showwarning("Relatório", resultado.get("mensagem", "Não foi possível abrir o relatório."))
        return

    nome = resultado.get("arquivo", "")
    try:
        os.startfile(nome)
    except OSError as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o relatório: {e}")
        return

    relatorio_carregado = True
    if not isinstance(dados_importados, dict):
        dados_importados = {}
    dados_importados["ultimo_relatorio_filtrado"] = nome
    dados_importados["ultimo_periodo_relatorios"] = (
        data_inicial.strftime("%d/%m/%Y"),
        data_final.strftime("%d/%m/%Y"),
    )
    dados_importados["ultimo_total_registros"] = int(resultado.get("total_documentos", 0))
    _atualizar_status_relatorios_ui(
        f"Relatório aberto com {resultado.get('total_documentos', 0)} registro(s)",
        tipo="ok",
        total_registros=resultado.get("total_documentos", 0),
    )


def abrir_relatorio_filtrado():
    abrir_relatorio()


def _montar_dataframe_exportacao_periodo(df_filtrado):
    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame()

    dados = df_filtrado.copy()
    dados["numero_doc"] = dados.apply(
        lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
        axis=1,
    )
    dados["numero_doc"] = dados["numero_doc"].astype(str).str.strip()
    dados = dados[dados["numero_doc"] != ""].copy()
    if dados.empty:
        return pd.DataFrame()

    dados["numero_doc_ordem"] = pd.to_numeric(dados["numero_doc"], errors="coerce")
    # Evita alerta do Excel "numero armazenado como texto" para documentos numericos.
    dados["numero_doc"] = dados.apply(
        lambda r: int(r["numero_doc_ordem"]) if pd.notna(r["numero_doc_ordem"]) else r["numero_doc"],
        axis=1,
    )
    dados["numero_ordem"] = pd.to_numeric(dados["numero"], errors="coerce")
    dados["numero_doc_ordem"] = dados["numero_doc_ordem"].fillna(dados["numero_ordem"])
    dados["tipo"] = dados["tipo"].astype(str).str.upper()
    dados["frete"] = dados["frete"].astype(str)
    dados["status"] = dados["status"].astype(str)
    dados["mes_referencia"] = pd.to_datetime(dados["data_competencia"], errors="coerce")
    dados = dados.sort_values(["data_emissao", "numero_doc_ordem", "numero_doc"], ascending=[True, True, True])

    export_df = dados[
        [
            "data_emissao",
            "mes_referencia",
            "numero_doc",
            "tipo",
            "frete",
            "valor_inicial",
            "valor_final",
            "status",
        ]
    ].copy()
    export_df.columns = [
        "Data Emissao",
        "Mes Referencia",
        "Numero Doc",
        "Tipo Doc",
        "Frete",
        "Valor Inicial",
        "Valor Final",
        "Status",
    ]
    return export_df


def exportar_relatorio_filtrado():
    global relatorio_carregado, dados_importados

    if not relatorio_carregado:
        mensagem = "Selecione e importe um relatório antes de exportar."
        _atualizar_status_relatorios_ui(mensagem, tipo="neutral", total_registros=None)
        messagebox.showwarning("Exportação", mensagem)
        return

    try:
        data_inicial, data_final = obter_periodo_relatorios(silencioso=False)
    except ValueError as exc:
        _atualizar_status_relatorios_ui(str(exc), tipo="erro", total_registros=None)
        messagebox.showwarning("Exportação", str(exc))
        return

    if data_inicial is None or data_final is None:
        mensagem = "Preencha o período da aba Relatórios antes de exportar."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Exportação", mensagem)
        return

    if data_inicial > data_final:
        mensagem = "A data inicial não pode ser maior que a data final."
        _atualizar_status_relatorios_ui(mensagem, tipo="erro", total_registros=None)
        messagebox.showwarning("Exportação", mensagem)
        return

    df_memoria = _obter_documentos_em_memoria(force=False)
    df_filtrado, msg_erro = _obter_dataframe_relatorio_filtrado(
        data_inicial,
        data_final,
        docs_df_base=df_memoria,
    )
    if msg_erro:
        _atualizar_status_relatorios_ui(msg_erro, tipo="erro", total_registros=0)
        messagebox.showwarning("Exportação", msg_erro)
        return
    if df_filtrado.empty:
        mensagem = "Nenhum registro encontrado para o período selecionado."
        _atualizar_status_relatorios_ui(mensagem, tipo="neutral", total_registros=0)
        messagebox.showinfo("Exportação", mensagem)
        return

    df_export = _montar_dataframe_exportacao_periodo(df_filtrado)
    if df_export.empty:
        mensagem = "Nenhum registro válido para exportação no período selecionado."
        _atualizar_status_relatorios_ui(mensagem, tipo="neutral", total_registros=0)
        messagebox.showinfo("Exportação", mensagem)
        return

    pasta_saida = obter_pasta_saida_relatorios()
    nome_arquivo = "Faturamento_AC.xlsx"
    caminho_saida = os.path.join(pasta_saida, nome_arquivo)

    try:
        with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Relatório Filtrado")
            ws = writer.sheets["Relatório Filtrado"]
            ws.sheet_view.showGridLines = False
            ws.freeze_panes = "A2"
            for idx, coluna in enumerate(df_export.columns, start=1):
                maior = max(
                    len(str(coluna)),
                    int(df_export[coluna].astype(str).str.len().max()) if not df_export.empty else 0,
                )
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(maior + 2, 42)
            # Coluna E = Frete. Força largura mínima para INTERCOMPANY sem corte visual.
            ws.column_dimensions["E"].width = max(float(ws.column_dimensions["E"].width or 0), 18.0)
            for linha in range(2, len(df_export) + 2):
                ws[f"A{linha}"].number_format = "DD/MM/YYYY"
                ws[f"B{linha}"].number_format = '[$-pt-BR]mmmm/yyyy'
                ws[f"C{linha}"].number_format = "0"
                ws[f"E{linha}"].alignment = Alignment(horizontal="left", vertical="center")
                ws[f"F{linha}"].number_format = "R$ #,##0.00"
                ws[f"G{linha}"].number_format = "R$ #,##0.00"
    except Exception as exc:
        _atualizar_status_relatorios_ui("Falha ao exportar relatório", tipo="erro", total_registros=len(df_export))
        messagebox.showerror("Exportação", f"Erro ao exportar relatório do período.\n\n{exc}")
        return

    if not isinstance(dados_importados, dict):
        dados_importados = {}
    dados_importados["ultimo_export_periodo"] = caminho_saida
    dados_importados["ultimo_total_registros"] = int(len(df_export))

    _atualizar_status_relatorios_ui(
        f"Relatório atualizado com {len(df_export)} registro(s)",
        tipo="ok",
        total_registros=len(df_export),
    )
    messagebox.showinfo(
        "Exportação",
        "Relatório atualizado com sucesso.\n\n"
        f"Arquivo: {nome_arquivo}\n"
        f"Registros: {len(df_export)}",
    )


# ------------------------
# ATUALIZAR
# ------------------------

# ------------------------
# STATUS
# ------------------------

# ------------------------
# EXCEL
# ------------------------

def gerar_excel(data_inicial=None, data_final=None, exibir_mensagem=True):
    def _falha(msg, erro=False):
        if exibir_mensagem:
            if erro:
                messagebox.showerror("Excel", msg)
            else:
                messagebox.showwarning("Excel", msg)
        return {"ok": False, "mensagem": msg, "arquivo": "", "total_documentos": 0}

    if data_inicial is None or data_final is None:
        try:
            data_inicial, data_final = obter_periodo_relatorios(silencioso=False)
        except ValueError as erro_data:
            return _falha(str(erro_data))

    if data_inicial is None or data_final is None:
        return _falha("Período de Relatórios não configurado.")

    if data_inicial > data_final:
        return _falha("A data inicial não pode ser maior que a data final.")

    df, msg_erro = _obter_dataframe_relatorio_filtrado(
        data_inicial,
        data_final,
        docs_df_base=_obter_documentos_em_memoria(force=False),
    )
    if msg_erro:
        return _falha(msg_erro)
    if df.empty:
        return _falha(
            f"Não há documentos para o período selecionado ({data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')})."
        )

    df = df.sort_values(["data_emissao", "numero"], ascending=[True, True])
    df["numero_exibicao"] = df.apply(
        lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
        axis=1,
    )
    df["Concat"] = df["numero_exibicao"].astype(str) + " " + df["tipo"].astype(str)
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
        dados["numero_doc"] = dados.apply(
            lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
            axis=1,
        )
        dados["numero_doc"] = dados["numero_doc"].astype(str).str.strip()
        dados = dados[dados["numero_doc"] != ""].copy()
        dados["numero_doc_ordem"] = pd.to_numeric(dados["numero_doc"], errors="coerce")
        dados["numero_doc"] = dados.apply(
            lambda r: int(r["numero_doc_ordem"]) if pd.notna(r["numero_doc_ordem"]) else r["numero_doc"],
            axis=1,
        )
        dados["numero_ordem"] = pd.to_numeric(dados["numero"], errors="coerce")
        dados["numero_doc_ordem"] = dados["numero_doc_ordem"].fillna(dados["numero_ordem"])
        dados["Concat"] = dados["numero_doc"].astype(str) + " " + dados["tipo"].astype(str)
        dados = dados.sort_values(["data_emissao", "numero_doc_ordem", "numero_doc"], ascending=[True, True, True])

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

    df_relatorio_1 = montar_df_relatorio(df_base, usar_numero_real_nf=True)
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
                    # Coluna G = Frete. Mantém alinhamento à esquerda para leitura sem corte.
                    ws[f"G{row}"].alignment = Alignment(horizontal="left", vertical="center")

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
                # Coluna G (Frete): largura mínima para acomodar INTERCOMPANY sem corte.
                ws.column_dimensions["G"].width = max(float(ws.column_dimensions["G"].width or 0), 18.0)

            formatar_aba(nome_aba_1, df_relatorio_1)
            formatar_aba(nome_aba_2, df_relatorio_2)

        if exibir_mensagem:
            messagebox.showinfo(
                "Excel",
                f"Relatório gerado com {len(df_relatorio_1)} documento(s) no período {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}\n\nAbas: Faturamento AC e Faturamento AC 2\nArquivo: {nome}",
            )
        return {
            "ok": True,
            "mensagem": "",
            "arquivo": nome,
            "total_documentos": int(len(df_relatorio_1)),
            "periodo": (
                data_inicial.strftime("%d/%m/%Y"),
                data_final.strftime("%d/%m/%Y"),
            ),
        }
    except Exception as e:
        return _falha(f"Erro ao gerar o relatório: {e}", erro=True)


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
        self.title("Sistema de Faturamento - Horizonte Logística")
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
inicializar_filtros_dashboard()
inicializar_filtros_relatorios()

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

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 89 ou 2390")
    numero_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 10))

    ctk.CTkLabel(form, text="Novo mês de competência").grid(row=2, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Ano da competência").grid(row=2, column=1, sticky="w", padx=6, pady=(0, 4))

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
            messagebox.showwarning("Aviso", "Tipo de documento inválido.")
            return

        if mes_novo not in MESES:
            messagebox.showwarning("Aviso", "Selecione um mês válido.")
            return

        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um número de documento válido.")
            return

        if not ano_novo.isdigit():
            messagebox.showwarning("Aviso", "Informe um ano válido.")
            return

        numero = int(numero_texto)
        alterados = alterar_competencia_documento(tipo, numero, mes_novo, int(ano_novo))
        if alterados == 0:
            messagebox.showwarning("Aviso", "Documento não encontrado para alterar competência.")
            return

        messagebox.showinfo("Sucesso", "Competência atualizada com sucesso.")
        dialog.destroy()

    ctk.CTkButton(form, text="Salvar alteração", command=salvar_alteracao, width=200).grid(row=4, column=0, columnspan=2, pady=(14, 0))


def _abrir_dialogo_declarar_frete(rotulo_frete):
    dialog = ctk.CTkToplevel(app)
    dialog.title(rotulo_frete)
    centralizar_janela(dialog, 620, 220)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Tipo de documento").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Número do documento").grid(row=0, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=260)
    tipo_combo.set("NF")
    tipo_combo.grid(row=1, column=0, sticky="ew", padx=6)

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 92 ou 2390")
    numero_entry.grid(row=1, column=1, sticky="ew", padx=6)

    desfazer_var = ctk.BooleanVar(value=False)
    texto_desfazer = f"Des{rotulo_frete.lower()} (voltar para FRANQUIA)"
    ctk.CTkCheckBox(
        form,
        text=texto_desfazer,
        variable=desfazer_var,
        onvalue=True,
        offvalue=False,
    ).grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 0))

    def confirmar():
        tipo = tipo_combo.get().strip().upper()
        numero_texto = numero_entry.get().strip()

        if tipo not in {"NF", "CTE"}:
            messagebox.showwarning("Aviso", "Tipo de documento inválido.")
            return
        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um número de documento válido.")
            return

        numero_doc = int(numero_texto)
        acao = rotulo_frete.upper().strip()

        if desfazer_var.get():
            # Desfaz a modalidade manual e retorna para FRANQUIA.
            resultado = salvar_alteracao_frete_manual(tipo, numero_doc, "FRANQUIA")
            if resultado.get("encontrados", 0) == 0:
                messagebox.showwarning("Aviso", f"Documento não encontrado para des{rotulo_frete.lower()}.")
                return
            messagebox.showinfo(
                "Sucesso",
                f"Des{rotulo_frete.lower()} aplicado com sucesso.\n"
                "A modalidade de frete foi revertida para FRANQUIA.\n\n"
                "A alteração foi salva e será refletida nas próximas consultas e exportações.",
            )
        else:
            # Declaração manual da modalidade.
            if acao == "INTERCOMPANY":
                resultado = declarar_intercompany(tipo, numero_doc)
            elif acao == "DELTA":
                resultado = declarar_delta(tipo, numero_doc)
            elif acao == "SPOT":
                resultado = declarar_spot(tipo, numero_doc)
            else:
                resultado = salvar_alteracao_frete_manual(tipo, numero_doc, acao)

            modalidade = str(resultado.get("modalidade", _normalizar_modalidade_frete(acao))).upper()
            if resultado.get("encontrados", 0) == 0:
                messagebox.showwarning("Aviso", f"Documento não encontrado para {rotulo_frete.lower()}.")
                return
            messagebox.showinfo(
                "Sucesso",
                f"{modalidade} aplicado com sucesso.\n"
                f"A modalidade de frete do documento agora é {modalidade}.\n\n"
                "A alteração foi salva e será refletida nas próximas consultas e exportações.",
            )
        dialog.destroy()

    ctk.CTkButton(form, text="Confirmar", command=confirmar, width=220).grid(row=3, column=0, columnspan=2, pady=(14, 0))


def abrir_dialogo_declarar_intercompany():
    _abrir_dialogo_declarar_frete("Intercompany")


def abrir_dialogo_declarar_delta():
    _abrir_dialogo_declarar_frete("Delta")


def abrir_dialogo_declarar_spot():
    _abrir_dialogo_declarar_frete("Spot")


def abrir_dialogo_cancelar_documento():
    dialog = ctk.CTkToplevel(app)
    dialog.title("Cancelar documento")
    centralizar_janela(dialog, 620, 220)
    dialog.grab_set()

    form = ctk.CTkFrame(dialog, fg_color="transparent")
    form.pack(fill="both", expand=True, padx=16, pady=12)
    form.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkLabel(form, text="Tipo de documento").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Número do documento").grid(row=0, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=260)
    tipo_combo.set("NF")
    tipo_combo.grid(row=1, column=0, sticky="ew", padx=6)

    numero_entry = ctk.CTkEntry(form, width=260, placeholder_text="Ex.: 89 ou 2390")
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
            messagebox.showwarning("Aviso", "Tipo de documento inválido.")
            return

        if not numero_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe um número de documento válido.")
            return

        numero = int(numero_texto)
        if desfazer_var.get():
            alterados = desfazer_cancelamento_documento(tipo, numero)
            if alterados == 0:
                messagebox.showwarning("Aviso", "Documento não encontrado ou não está cancelado.")
                return
            messagebox.showinfo("Sucesso", "Cancelamento desfeito com sucesso.")
        else:
            alterados = cancelar_documento(tipo, numero)
            if alterados == 0:
                messagebox.showwarning("Aviso", "Documento não encontrado.")
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
    ctk.CTkLabel(form, text="Número").grid(row=1, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_antigo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=320)
    tipo_antigo_combo.set("NF")
    tipo_antigo_combo.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 10))

    numero_antigo_entry = ctk.CTkEntry(form, width=320, placeholder_text="Ex.: 89 ou 2390")
    numero_antigo_entry.grid(row=2, column=1, sticky="ew", padx=6, pady=(0, 10))

    ctk.CTkLabel(form, text="Documento substituto", font=ctk.CTkFont(weight="bold")).grid(
        row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 4)
    )
    ctk.CTkLabel(form, text="Tipo").grid(row=4, column=0, sticky="w", padx=6, pady=(0, 4))
    ctk.CTkLabel(form, text="Número").grid(row=4, column=1, sticky="w", padx=6, pady=(0, 4))

    tipo_novo_combo = ctk.CTkComboBox(form, values=["NF", "CTE"], width=320)
    tipo_novo_combo.set("NF")
    tipo_novo_combo.grid(row=5, column=0, sticky="ew", padx=6)

    numero_novo_entry = ctk.CTkEntry(form, width=320, placeholder_text="Ex.: 90 ou 2391")
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
            messagebox.showwarning("Aviso", "Tipo de documento inválido.")
            return

        if not numero_antigo_texto.isdigit() or not numero_novo_texto.isdigit():
            messagebox.showwarning("Aviso", "Informe números de documento válidos.")
            return

        numero_antigo = int(numero_antigo_texto)
        numero_novo = int(numero_novo_texto)

        if tipo_antigo == tipo_novo and numero_antigo == numero_novo:
            messagebox.showwarning("Aviso", "Documento antigo e substituto não podem ser iguais.")
            return

        if desfazer_var.get():
            antigo_restaurado, novo_restaurado = desfazer_substituicao(
                tipo_antigo, numero_antigo, tipo_novo, numero_novo
            )
            if antigo_restaurado == 0:
                messagebox.showwarning("Aviso", "Documento antigo não encontrado ou não está substituído.")
                return
            if novo_restaurado == 0:
                messagebox.showwarning("Aviso", "Documento substituto não encontrado ou não está como substituto.")
                return
            messagebox.showinfo("Sucesso", "Substituição desfeita com sucesso.")
        else:
            novo_alterado, antigo_alterado = registrar_substituicao(
                tipo_antigo, numero_antigo, tipo_novo, numero_novo
            )
            if novo_alterado == 0:
                messagebox.showwarning("Aviso", "Documento substituto não encontrado.")
                return
            if antigo_alterado == 0:
                messagebox.showwarning("Aviso", "Documento antigo não encontrado.")
                return
            messagebox.showinfo("Sucesso", "Substituição registrada com sucesso.")
        dialog.destroy()

    ctk.CTkButton(form, text="Confirmar", command=confirmar_substituicao, width=220).grid(
        row=7, column=0, columnspan=2, pady=(14, 0)
    )


def abrir_relatorio_cancelados():
    try:
        data_inicial, data_final = obter_periodo_relatorios(silencioso=False)
    except ValueError as exc:
        messagebox.showwarning("Filtro de período", str(exc))
        return

    if data_inicial is None or data_final is None:
        messagebox.showwarning("Filtro de período", "Preencha as datas na aba Relatórios.")
        return

    if data_inicial > data_final:
        messagebox.showwarning("Filtro de período", "A data inicial não pode ser maior que a data final.")
        return

    conn = obter_conexao_banco()
    df = pd.read_sql_query(
        """
        SELECT id, tipo, numero, numero_original, data_emissao, valor_final, status
        FROM documentos
        WHERE UPPER(status) LIKE '%CANCELADO%'
        """,
        conn,
    )
    conn.close()

    df["data_emissao_dt"] = pd.to_datetime(df["data_emissao"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data_emissao_dt"])
    df = df[(df["data_emissao_dt"] >= data_inicial) & (df["data_emissao_dt"] <= data_final)]
    df["numero_exibicao"] = df.apply(
        lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
        axis=1,
    )
    df["chave_documento"] = df.apply(
        lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")

    if df.empty:
        messagebox.showinfo(
            "Relatório",
            (
                "Nenhum documento cancelado encontrado no período selecionado.\n\n"
                f"Período: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}"
            ),
        )
        return

    df = df.sort_values(["data_emissao_dt", "numero"], ascending=[True, True])

    janela = ctk.CTkToplevel(app)
    janela.title("Relatório de documentos cancelados")
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
                str(linha["numero_exibicao"]),
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

    busca_entry = ctk.CTkEntry(dialog, width=300, placeholder_text="Digite o numero original do documento")
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
            SELECT id, tipo, numero, numero_original, data_emissao, valor_final, status
            FROM documentos
            WHERE CAST(numero AS TEXT) LIKE ?
               OR COALESCE(numero_original, '') LIKE ?
            ORDER BY data_emissao, id
            """,
            conn,
            params=(f"%{termo}%", f"%{termo}%"),
        )

        conn.close()
        if df.empty:
            return

        df["numero_exibicao"] = df.apply(
            lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
            axis=1,
        )
        df["chave_documento"] = df.apply(
            lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
            axis=1,
        )
        df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")

        for _, row in df.iterrows():

            tabela.insert(
                "",
                "end",
                values=(
                    row["tipo"],
                    row["numero_exibicao"],
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
    data_inicial, data_final = obter_periodo_dashboard(silencioso=True)
    if data_inicial is None or data_final is None:
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
    df["numero_exibicao"] = df.apply(
        lambda r: _numero_documento_exibicao(r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")),
        axis=1,
    )

    df["chave_documento"] = df.apply(
        lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
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
        ax.set_title("Faturamento por período", fontsize=11, fontweight="bold", color=UI_THEME["text_primary"], pad=8)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.text(
            0.5,
            0.5,
            "Sem meses com faturamento para o período.",
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

    ax.set_title("Faturamento por período", fontsize=11, fontweight="bold", color=UI_THEME["text_primary"], pad=8)
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
        msg = "Período inválido.\nA data inicial precisa ser menor ou igual à data final."
        _mostrar_placeholder_grafico("faturamento", msg)
        _mostrar_placeholder_grafico("comparativo", msg)
        return

    if df is None:
        return

    if df.empty:
        msg = "Sem dados para o período selecionado."
        _mostrar_placeholder_grafico("faturamento", msg)
        _mostrar_placeholder_grafico("comparativo", msg)
        return

    try:
        fig_faturamento = _criar_figura_faturamento_periodo(df)
        fig_comparativo = _criar_figura_comparativo_tipos(df)
    except Exception as exc:
        _mostrar_placeholder_grafico("faturamento", f"Falha ao desenhar gráfico.\n{exc}")
        _mostrar_placeholder_grafico("comparativo", f"Falha ao desenhar gráfico.\n{exc}")
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
        messagebox.showerror("Gráfico", f"Falha ao carregar bibliotecas do gráfico.\n\n{exc}")
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
        messagebox.showwarning("Gráfico", "Não há dados para gerar gráfico.")
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
        messagebox.showwarning("Gráfico", "Não há dados para gerar gráfico.")
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
            "Selecione ao menos um mês para exibir o gráfico.",
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
        resumo_txt = f"Período exibido: {periodo_ini} a {periodo_fim}"
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
SCREEN_NAV_CATALOG = {
    "dashboard": "Dashboard",
    "relatorios": "Relatórios",
    "alteracoes": "Alterações",
    "configuracoes": "Configurações",
}
SCREEN_NAV_DEFAULT_ORDER = list(SCREEN_NAV_CATALOG.keys())
ACTION_HOST_SCREENS = ["relatorios", "alteracoes", "configuracoes"]
SCREEN_ORDER_ATUAL = SCREEN_NAV_DEFAULT_ORDER.copy()
ACTION_LAYOUT_ATUAL = {}

# Registrar callbacks para src.documentos notificar mudanças na UI
_register_doc_on_change(_atualizar_cache_documentos_pos_alteracao)
_register_doc_on_change(atualizar_dashboard)


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

    _safe_config(container_scroll if "container_scroll" in globals() else None, fg_color=UI_THEME["app_bg"])
    scroll_canvas = ui_refs.get("scroll_canvas")
    if scroll_canvas is not None:
        try:
            scroll_canvas.configure(
                bg=UI_THEME["app_bg"],
                highlightthickness=0,
                bd=0,
            )
        except Exception:
            pass
    scroll_vertical = ui_refs.get("scrollbar_vertical")
    if scroll_vertical is not None:
        try:
            scroll_vertical.configure(
                bg=UI_THEME["scroll_btn"],
                activebackground=UI_THEME["scroll_btn_hover"],
                troughcolor=UI_THEME["app_bg"],
            )
        except Exception:
            pass
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
    _safe_config(
        ui_refs.get("btn_reordenar_interface"),
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
    _safe_config(
        dashboard_data_inicio_entry if "dashboard_data_inicio_entry" in globals() else None,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        text_color=UI_THEME["text_primary"],
    )
    _safe_config(
        dashboard_data_fim_entry if "dashboard_data_fim_entry" in globals() else None,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        text_color=UI_THEME["text_primary"],
    )

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
    _safe_config(ui_refs.get("relatorios_status_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("relatorios_status_label"), text_color=UI_THEME["success_text"])
    _safe_config(ui_refs.get("relatorios_arquivo_label"), text_color=UI_THEME["text_primary"])
    _safe_config(ui_refs.get("relatorios_saida_label"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_periodo_label"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_registros_label"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_filtro_card"), fg_color=UI_THEME["surface"], border_color=UI_THEME["border"])
    _safe_config(ui_refs.get("relatorios_filtro_titulo"), text_color=UI_THEME["text_primary"])
    _safe_config(ui_refs.get("relatorios_filtro_subtitulo"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_filtro_ate"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_filtro_icon_inicio"), text_color=UI_THEME["text_secondary"])
    _safe_config(ui_refs.get("relatorios_filtro_icon_fim"), text_color=UI_THEME["text_secondary"])
    _safe_config(
        relatorio_data_inicio_entry if "relatorio_data_inicio_entry" in globals() else None,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        text_color=UI_THEME["text_primary"],
    )
    _safe_config(
        relatorio_data_fim_entry if "relatorio_data_fim_entry" in globals() else None,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        text_color=UI_THEME["text_primary"],
    )
    relatorios_aplicar_filtro_btn = ui_refs.get("relatorios_aplicar_filtro_btn")
    _safe_config(
        relatorios_aplicar_filtro_btn,
        fg_color=UI_THEME["surface_alt"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        border_color=UI_THEME["border"],
    )
    if relatorios_aplicar_filtro_btn is not None:
        _aplicar_microinteracao_botao(relatorios_aplicar_filtro_btn, "secondary")

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
        variant = getattr(btn, "_action_variant", "primary")
        if variant == "secondary":
            _safe_config(
                btn,
                fg_color=UI_THEME["surface_alt"],
                hover_color=UI_THEME["tab_hover"],
                text_color=UI_THEME["text_primary"],
                border_color=UI_THEME["border"],
            )
            _aplicar_microinteracao_cta(
                btn,
                UI_THEME["surface_alt"],
                UI_THEME["tab_hover"],
                UI_THEME["tab_press"],
                UI_THEME["border"],
                UI_THEME["border"],
                text_color=UI_THEME["text_primary"],
            )
        else:
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
                text_color=UI_THEME["on_accent"],
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
    _atualizar_status_relatorios_ui()


def alternar_tema_interface():
    novo_modo = "dark" if current_theme_mode == "light" else "light"
    _definir_tema_interface(novo_modo, persistir=True)
    aplicar_tema_interface()


def _catalogo_acoes_interface():
    return {
        "selecionar_relatorio": {"titulo": "Selecionar relatório", "comando": selecionar_relatorio},
        "abrir_relatorio": {"titulo": "Abrir relatório", "comando": abrir_relatorio},
        "exportar_relatorio_periodo": {"titulo": "Exportar relatório do período", "comando": exportar_relatorio_filtrado},
        "pasta_saida": {"titulo": "Pasta de saída", "comando": selecionar_pasta_saida_relatorios},
        "relatorio_cancelados": {"titulo": "Relatórios cancelados", "comando": abrir_relatorio_cancelados},
        "grafico_faturamento": {"titulo": "Gráfico de faturamento", "comando": abrir_grafico_faturamento},
        "buscar_documento": {"titulo": "Buscar documento", "comando": abrir_busca_documentos},
        "alterar_competencia": {"titulo": "Alterar competência", "comando": abrir_dialogo_alterar_competencia},
        "substituir_documento": {"titulo": "Substituir documento", "comando": abrir_dialogo_substituir_documento},
        "cancelar_documento": {"titulo": "Cancelar documento", "comando": abrir_dialogo_cancelar_documento},
        "declarar_intercompany": {"titulo": "Declarar intercompany", "comando": abrir_dialogo_declarar_intercompany},
        "declarar_delta": {"titulo": "Declarar delta", "comando": abrir_dialogo_declarar_delta},
        "declarar_spot": {"titulo": "Declarar spot", "comando": abrir_dialogo_declarar_spot},
        "alternar_tema": {"titulo": "Alternar tema", "comando": alternar_tema_interface},
        "exportar_configuracoes": {"titulo": "Exportar configurações", "comando": exportar_configuracoes_ui},
        "importar_configuracoes": {"titulo": "Importar configurações", "comando": importar_configuracoes_ui},
    }


def _layout_botoes_padrao():
    return {
        "relatorios": [
            "selecionar_relatorio",
            "abrir_relatorio",
            "exportar_relatorio_periodo",
            "pasta_saida",
            "relatorio_cancelados",
            "grafico_faturamento",
            "buscar_documento",
        ],
        "alteracoes": [
            "alterar_competencia",
            "substituir_documento",
            "cancelar_documento",
            "declarar_intercompany",
            "declarar_delta",
            "declarar_spot",
        ],
        "configuracoes": [
            "alternar_tema",
            "exportar_configuracoes",
            "importar_configuracoes",
        ],
    }


def _normalizar_ordem_abas(ordem_raw):
    ordem = []
    for item in ordem_raw or []:
        chave = str(item).strip().lower()
        if chave in SCREEN_NAV_CATALOG and chave not in ordem:
            ordem.append(chave)
    for chave in SCREEN_NAV_DEFAULT_ORDER:
        if chave not in ordem:
            ordem.append(chave)
    return ordem


def _carregar_ordem_abas_interface():
    try:
        raw = obter_configuracao("ui_ordem_abas", "").strip()
        if not raw:
            return SCREEN_NAV_DEFAULT_ORDER.copy()
        dados = json.loads(raw)
        if not isinstance(dados, list):
            return SCREEN_NAV_DEFAULT_ORDER.copy()
        return _normalizar_ordem_abas(dados)
    except Exception:
        return SCREEN_NAV_DEFAULT_ORDER.copy()


def _normalizar_layout_botoes(layout_raw):
    catalogo = _catalogo_acoes_interface()
    layout_padrao = _layout_botoes_padrao()
    validos = set(catalogo.keys())
    usados = set()
    layout = {aba: [] for aba in ACTION_HOST_SCREENS}

    if isinstance(layout_raw, dict):
        for aba in ACTION_HOST_SCREENS:
            ids = layout_raw.get(aba, [])
            if not isinstance(ids, list):
                continue
            for acao_id in ids:
                chave = str(acao_id).strip().lower()
                if chave in validos and chave not in usados:
                    layout[aba].append(chave)
                    usados.add(chave)

    for aba in ACTION_HOST_SCREENS:
        for acao_id in layout_padrao.get(aba, []):
            if acao_id not in usados:
                layout[aba].append(acao_id)
                usados.add(acao_id)

    for acao_id in catalogo.keys():
        if acao_id not in usados:
            layout["configuracoes"].append(acao_id)
            usados.add(acao_id)

    # Mantém fluxo principal sempre disponível na aba Relatórios.
    obrigatorios_relatorios = ["selecionar_relatorio", "abrir_relatorio"]
    for acao_id in obrigatorios_relatorios:
        for aba in ACTION_HOST_SCREENS:
            if acao_id in layout.get(aba, []):
                layout[aba] = [x for x in layout[aba] if x != acao_id]
        layout["relatorios"].insert(0, acao_id)

    # Remove possiveis duplicidades apos reforco dos obrigatorios.
    vistos_rel = set()
    rel_limpo = []
    for acao_id in layout["relatorios"]:
        if acao_id not in vistos_rel:
            rel_limpo.append(acao_id)
            vistos_rel.add(acao_id)
    layout["relatorios"] = rel_limpo

    # Mantém alternância de tema concentrada em Configurações.
    for aba in ACTION_HOST_SCREENS:
        if aba != "configuracoes" and "alternar_tema" in layout.get(aba, []):
            layout[aba] = [x for x in layout[aba] if x != "alternar_tema"]
    if "alternar_tema" not in layout["configuracoes"]:
        layout["configuracoes"].insert(0, "alternar_tema")

    return layout


def _carregar_layout_botoes_interface():
    try:
        raw = obter_configuracao("ui_layout_botoes", "").strip()
        if not raw:
            return _normalizar_layout_botoes(_layout_botoes_padrao())
        dados = json.loads(raw)
        return _normalizar_layout_botoes(dados)
    except Exception:
        return _normalizar_layout_botoes(_layout_botoes_padrao())


def _recarregar_layout_interface():
    global SCREEN_ORDER_ATUAL, ACTION_LAYOUT_ATUAL
    SCREEN_ORDER_ATUAL = _carregar_ordem_abas_interface()
    ACTION_LAYOUT_ATUAL = _carregar_layout_botoes_interface()


def _salvar_layout_interface(ordem_abas, layout_botoes):
    salvar_configuracao(
        "ui_ordem_abas",
        json.dumps(_normalizar_ordem_abas(ordem_abas), ensure_ascii=False),
    )
    salvar_configuracao(
        "ui_layout_botoes",
        json.dumps(_normalizar_layout_botoes(layout_botoes), ensure_ascii=False),
    )


def _renderizar_grade_acoes_por_tela(grade, tela_id):
    catalogo = _catalogo_acoes_interface()
    ordem = ACTION_LAYOUT_ATUAL.get(tela_id, [])
    linha = 0
    coluna = 0
    for acao_id in ordem:
        item = catalogo.get(acao_id)
        if not item:
            continue
        _criar_botao_acao(grade, item["titulo"], item["comando"]).grid(
            row=linha,
            column=coluna,
            padx=6,
            pady=6,
            sticky="ew",
        )
        coluna += 1
        if coluna > 1:
            coluna = 0
            linha += 1
    return bool(ordem)


def abrir_dialogo_reordenar_interface():
    ordem_abas = list(SCREEN_ORDER_ATUAL)
    layout_local = {aba: list(ids) for aba, ids in ACTION_LAYOUT_ATUAL.items()}
    catalogo = _catalogo_acoes_interface()
    labels_abas = {sid: SCREEN_NAV_CATALOG[sid] for sid in SCREEN_NAV_DEFAULT_ORDER}
    id_por_label_aba = {v: k for k, v in labels_abas.items()}
    labels_host = [labels_abas[sid] for sid in ACTION_HOST_SCREENS]

    dialog = ctk.CTkToplevel(app)
    dialog.title("Organizar interface")
    centralizar_janela(dialog, 860, 520)
    dialog.grab_set()

    container = ctk.CTkFrame(dialog, fg_color="transparent")
    container.pack(fill="both", expand=True, padx=14, pady=14)
    container.grid_columnconfigure((0, 1), weight=1)

    card_abas = _criar_card(container, corner_radius=14)
    card_abas.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    card_abas.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(
        card_abas,
        text="Ordem das abas",
        font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=12, pady=(10, 8))

    lista_abas = tk.Listbox(
        card_abas,
        height=10,
        activestyle="none",
        bd=0,
        highlightthickness=0,
        bg=UI_THEME["surface_alt"],
        fg=UI_THEME["text_primary"],
        selectbackground=UI_THEME["accent"],
        selectforeground=UI_THEME["on_accent"],
        font=("Segoe UI", 11),
    )
    lista_abas.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    acoes_abas = ctk.CTkFrame(card_abas, fg_color="transparent")
    acoes_abas.pack(fill="x", padx=12, pady=(0, 12))

    card_botoes = _criar_card(container, corner_radius=14)
    card_botoes.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    card_botoes.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(
        card_botoes,
        text="Botoes por aba",
        font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=12, pady=(10, 8))

    combo_aba_origem = ctk.CTkComboBox(
        card_botoes,
        values=labels_host,
        width=260,
        state="readonly",
    )
    combo_aba_origem.set(labels_host[0])
    combo_aba_origem.pack(anchor="w", padx=12, pady=(0, 8))

    lista_botoes = tk.Listbox(
        card_botoes,
        height=10,
        activestyle="none",
        bd=0,
        highlightthickness=0,
        bg=UI_THEME["surface_alt"],
        fg=UI_THEME["text_primary"],
        selectbackground=UI_THEME["accent"],
        selectforeground=UI_THEME["on_accent"],
        font=("Segoe UI", 11),
    )
    lista_botoes.pack(fill="both", expand=True, padx=12, pady=(0, 10))
    ctk.CTkLabel(
        card_botoes,
        text="Dica: arraste um botao na lista para reordenar.",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    ).pack(anchor="w", padx=12, pady=(0, 8))

    acoes_botoes = ctk.CTkFrame(card_botoes, fg_color="transparent")
    acoes_botoes.pack(fill="x", padx=12, pady=(0, 6))

    mover_frame = ctk.CTkFrame(card_botoes, fg_color="transparent")
    mover_frame.pack(fill="x", padx=12, pady=(0, 12))

    combo_aba_destino = ctk.CTkComboBox(
        mover_frame,
        values=labels_host,
        width=220,
        state="readonly",
    )
    combo_aba_destino.set(labels_host[1] if len(labels_host) > 1 else labels_host[0])
    combo_aba_destino.pack(side="left", padx=(0, 8))

    def _id_aba_origem():
        return id_por_label_aba.get(combo_aba_origem.get(), "relatorios")

    def _id_aba_destino():
        return id_por_label_aba.get(combo_aba_destino.get(), "alteracoes")

    def _atualizar_lista_abas(selecao_idx=None):
        lista_abas.delete(0, tk.END)
        for sid in ordem_abas:
            lista_abas.insert(tk.END, labels_abas[sid])
        if not ordem_abas:
            return
        idx = 0 if selecao_idx is None else max(0, min(selecao_idx, len(ordem_abas) - 1))
        lista_abas.selection_clear(0, tk.END)
        lista_abas.selection_set(idx)
        lista_abas.activate(idx)

    def _atualizar_lista_botoes(selecao_idx=None):
        lista_botoes.delete(0, tk.END)
        origem = _id_aba_origem()
        ids = layout_local.get(origem, [])
        for acao_id in ids:
            titulo = catalogo.get(acao_id, {}).get("titulo", acao_id)
            lista_botoes.insert(tk.END, titulo)
        if not ids:
            return
        idx = 0 if selecao_idx is None else max(0, min(selecao_idx, len(ids) - 1))
        lista_botoes.selection_clear(0, tk.END)
        lista_botoes.selection_set(idx)
        lista_botoes.activate(idx)

    drag_state = {"indice": None}

    def _iniciar_arraste_botao(event):
        origem = _id_aba_origem()
        ids = layout_local.get(origem, [])
        if not ids:
            drag_state["indice"] = None
            return
        idx = lista_botoes.nearest(event.y)
        idx = max(0, min(idx, len(ids) - 1))
        drag_state["indice"] = idx
        _atualizar_lista_botoes(idx)

    def _arrastar_botao(event):
        origem = _id_aba_origem()
        ids = layout_local.get(origem, [])
        idx_origem = drag_state.get("indice")
        if idx_origem is None or not ids:
            return
        idx_alvo = lista_botoes.nearest(event.y)
        idx_alvo = max(0, min(idx_alvo, len(ids) - 1))
        if idx_alvo == idx_origem:
            return
        item = ids.pop(idx_origem)
        ids.insert(idx_alvo, item)
        drag_state["indice"] = idx_alvo
        _atualizar_lista_botoes(idx_alvo)

    def _finalizar_arraste_botao(_event=None):
        drag_state["indice"] = None

    def _mover_aba(delta):
        selecao = lista_abas.curselection()
        if not selecao:
            return
        idx = int(selecao[0])
        novo = idx + delta
        if novo < 0 or novo >= len(ordem_abas):
            return
        ordem_abas[idx], ordem_abas[novo] = ordem_abas[novo], ordem_abas[idx]
        _atualizar_lista_abas(novo)

    def _mover_botao(delta):
        origem = _id_aba_origem()
        ids = layout_local.get(origem, [])
        selecao = lista_botoes.curselection()
        if not ids or not selecao:
            return
        idx = int(selecao[0])
        novo = idx + delta
        if novo < 0 or novo >= len(ids):
            return
        ids[idx], ids[novo] = ids[novo], ids[idx]
        _atualizar_lista_botoes(novo)

    def _mover_para_outra_aba():
        origem = _id_aba_origem()
        destino = _id_aba_destino()
        if origem == destino:
            return
        ids_origem = layout_local.get(origem, [])
        selecao = lista_botoes.curselection()
        if not ids_origem or not selecao:
            return
        idx = int(selecao[0])
        acao_id = ids_origem.pop(idx)
        layout_local.setdefault(destino, []).append(acao_id)
        _atualizar_lista_botoes(max(0, idx - 1))

    def _restaurar_padrao():
        ordem_abas[:] = SCREEN_NAV_DEFAULT_ORDER.copy()
        layout_padrao = _layout_botoes_padrao()
        layout_local.clear()
        for aba in ACTION_HOST_SCREENS:
            layout_local[aba] = list(layout_padrao.get(aba, []))
        combo_aba_origem.set(labels_abas["relatorios"])
        combo_aba_destino.set(labels_abas["alteracoes"])
        _atualizar_lista_abas(0)
        _atualizar_lista_botoes(0)

    def _salvar_e_aplicar():
        _salvar_layout_interface(ordem_abas, layout_local)
        _recarregar_layout_interface()
        tela_atual = app.current_screen or "dashboard"
        dialog.destroy()
        construir_tela_principal()
        aplicar_tema_interface()
        if tela_atual in app.screens:
            app.mostrar_tela(tela_atual)
        solicitar_atualizacao_dashboard(delay_ms=120)
        messagebox.showinfo("Layout atualizado", "Nova sequencia aplicada com sucesso.")

    ctk.CTkButton(acoes_abas, text="Subir aba", width=120, command=lambda: _mover_aba(-1)).pack(side="left", padx=(0, 8))
    ctk.CTkButton(acoes_abas, text="Descer aba", width=120, command=lambda: _mover_aba(1)).pack(side="left")

    ctk.CTkButton(acoes_botoes, text="Subir botao", width=120, command=lambda: _mover_botao(-1)).pack(side="left", padx=(0, 8))
    ctk.CTkButton(acoes_botoes, text="Descer botao", width=120, command=lambda: _mover_botao(1)).pack(side="left")
    ctk.CTkButton(mover_frame, text="Mover para aba", width=140, command=_mover_para_outra_aba).pack(side="left")

    combo_aba_origem.configure(command=lambda _v: _atualizar_lista_botoes(0))
    lista_botoes.bind("<ButtonPress-1>", _iniciar_arraste_botao, add="+")
    lista_botoes.bind("<B1-Motion>", _arrastar_botao, add="+")
    lista_botoes.bind("<ButtonRelease-1>", _finalizar_arraste_botao, add="+")

    rodape = ctk.CTkFrame(dialog, fg_color="transparent")
    rodape.pack(fill="x", padx=14, pady=(0, 14))
    ctk.CTkButton(rodape, text="Restaurar padrao", width=150, command=_restaurar_padrao).pack(side="left")
    ctk.CTkButton(rodape, text="Cancelar", width=120, command=dialog.destroy).pack(side="right", padx=(8, 0))
    ctk.CTkButton(rodape, text="Salvar", width=140, command=_salvar_e_aplicar).pack(side="right")

    _atualizar_lista_abas(0)
    _atualizar_lista_botoes(0)


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
    if ui_animations_paused or scroll_dragging:
        _safe_config(
            botao,
            fg_color=_normalizar_hex_cor(alvo_fg),
            text_color=_normalizar_hex_cor(alvo_texto),
            border_color=_normalizar_hex_cor(alvo_borda),
        )
        return

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


def _aplicar_microinteracao_cta(botao, base_fg, hover_fg, press_fg, base_border, hover_border, text_color=None):
    cor_texto = text_color or UI_THEME["on_accent"]

    def _on_enter(_evt):
        _animar_estilo_botao(
            botao,
            hover_fg,
            cor_texto,
            hover_border,
            passos=6,
            delay_ms=14,
        )

    def _on_leave(_evt):
        _animar_estilo_botao(
            botao,
            base_fg,
            cor_texto,
            base_border,
            passos=6,
            delay_ms=14,
        )

    def _on_press(_evt):
        _animar_estilo_botao(
            botao,
            press_fg,
            cor_texto,
            hover_border,
            passos=4,
            delay_ms=10,
        )

    def _on_release(_evt):
        _animar_estilo_botao(
            botao,
            hover_fg,
            cor_texto,
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


def _criar_botao_acao(parent, texto, comando, variant="primary", height=40):
    eh_secundario = str(variant).strip().lower() == "secondary"
    fg_color = UI_THEME["surface_alt"] if eh_secundario else UI_THEME["accent"]
    hover_color = UI_THEME["tab_hover"] if eh_secundario else UI_THEME["accent_hover"]
    press_color = UI_THEME["tab_press"] if eh_secundario else UI_THEME["cta_press"]
    border_color = UI_THEME["border"] if eh_secundario else UI_THEME["cta_border"]
    border_hover = UI_THEME["border"] if eh_secundario else UI_THEME["cta_border_hover"]
    text_color = UI_THEME["text_primary"] if eh_secundario else UI_THEME["on_accent"]

    botao = ctk.CTkButton(
        parent,
        text=texto,
        height=max(height, 42),
        corner_radius=14,
        border_width=1,
        border_color=border_color,
        fg_color=fg_color,
        hover_color=hover_color,
        text_color=text_color,
        font=ctk.CTkFont(family="Segoe UI Semibold", size=13, weight="bold"),
        command=comando,
    )
    botao._action_variant = "secondary" if eh_secundario else "primary"
    _aplicar_microinteracao_cta(
        botao,
        fg_color,
        hover_color,
        press_color,
        border_color,
        border_hover,
        text_color=text_color,
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
        menu_feedback_label.configure(text=f"Ação selecionada: {titulo_item}")

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
        menu_feedback_label.configure(text=f"Menu: {tab_titulo}")

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
        print(f"Aviso: não foi possível preparar a watermark da logo - {e}")
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
    lbl.place(**WATERMARK_POS)
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
        text="Sistema de Faturamento - Horizonte Logística",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.pack(pady=(10, 12))
    ui_refs["header_title"] = titulo


def _criar_filtro_periodo(parent):
    global dashboard_data_inicio_entry, dashboard_data_fim_entry

    card = _criar_card(parent, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 10))
    card.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
    ui_refs["filtro_card"] = card

    titulo = ctk.CTkLabel(
        card,
        text="Período de emissão",
        font=ctk.CTkFont(family="Segoe UI", size=19, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.grid(row=0, column=0, columnspan=6, sticky="w", padx=20, pady=(16, 4))
    ui_refs["filtro_titulo"] = titulo

    subtitulo = ctk.CTkLabel(
        card,
        text="Filtre os documentos por intervalo de emissão",
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

    dashboard_data_inicio_entry = ctk.CTkEntry(
        card,
        width=150,
        height=44,
        corner_radius=14,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    dashboard_data_inicio_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(0, 16))

    lbl_ate = ctk.CTkLabel(card, text="até", font=ctk.CTkFont(family="Segoe UI", size=12), text_color=UI_THEME["text_secondary"])
    lbl_ate.grid(
        row=2, column=2, padx=4, pady=(0, 16)
    )
    ui_refs["filtro_ate"] = lbl_ate

    icon_fim = ctk.CTkLabel(card, text="📅", font=ctk.CTkFont(size=16), text_color=UI_THEME["text_secondary"])
    icon_fim.grid(
        row=2, column=3, sticky="e", padx=(8, 6), pady=(0, 16)
    )
    ui_refs["filtro_icon_fim"] = icon_fim

    dashboard_data_fim_entry = ctk.CTkEntry(
        card,
        width=150,
        height=44,
        corner_radius=14,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    dashboard_data_fim_entry.grid(row=2, column=4, sticky="ew", padx=(0, 10), pady=(0, 16))

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
        command=aplicar_filtro_dashboard,
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

    inicializar_filtros_dashboard()

    dashboard_data_inicio_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(dashboard_data_inicio_entry))
    dashboard_data_inicio_entry.bind("<FocusOut>", solicitar_atualizacao_dashboard)
    dashboard_data_inicio_entry.bind("<Return>", solicitar_atualizacao_dashboard)
    dashboard_data_fim_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(dashboard_data_fim_entry))
    dashboard_data_fim_entry.bind("<FocusOut>", solicitar_atualizacao_dashboard)
    dashboard_data_fim_entry.bind("<Return>", solicitar_atualizacao_dashboard)


def _criar_info_relatorio(parent):
    global pasta_saida_label, pasta_label, progress, status_label

    info_card = _criar_card(parent, corner_radius=20)
    info_card.pack(fill="x", padx=18, pady=(0, 10))
    ui_refs["info_card"] = info_card
    ui_refs["relatorios_status_card"] = info_card

    titulo = ctk.CTkLabel(
        info_card,
        text="Informações operacionais",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
    )
    titulo.pack(fill="x", padx=18, pady=(12, 6))

    lbl_status = ctk.CTkLabel(
        info_card,
        text="Relatório carregado com sucesso",
        anchor="w",
        justify="left",
        text_color=UI_THEME["success_text"],
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
    )
    lbl_status.pack(fill="x", padx=18, pady=(0, 4))

    pasta_label = ctk.CTkLabel(
        info_card,
        text="Arquivo: Nenhum selecionado",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=12),
        wraplength=920,
    )
    pasta_label.pack(fill="x", padx=18, pady=(0, 4))

    pasta_saida_label = ctk.CTkLabel(
        info_card,
        text="Pasta de saída: -",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
        wraplength=920,
    )
    pasta_saida_label.pack(fill="x", padx=18, pady=(0, 4))

    divisor = ctk.CTkFrame(info_card, height=1, fg_color=UI_THEME["divider"], corner_radius=1)
    divisor.pack(fill="x", padx=18, pady=(4, 8))
    ui_refs["info_divider"] = divisor

    lbl_periodo = ctk.CTkLabel(
        info_card,
        text="Período aplicado: Período indefinido",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
        wraplength=920,
    )
    lbl_periodo.pack(fill="x", padx=18, pady=(0, 4))

    lbl_registros = ctk.CTkLabel(
        info_card,
        text="Registros no período: -",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
        wraplength=920,
    )
    lbl_registros.pack(fill="x", padx=18, pady=(0, 8))

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
        text="Processamento: Relatório 0 página(s) | Docs: 0",
        anchor="w",
        justify="left",
        text_color=UI_THEME["text_secondary"],
        font=ctk.CTkFont(family="Segoe UI", size=11),
    )
    status_label.pack(fill="x", padx=18, pady=(0, 14))

    ui_refs["relatorios_status_label"] = lbl_status
    ui_refs["relatorios_arquivo_label"] = pasta_label
    ui_refs["relatorios_saida_label"] = pasta_saida_label
    ui_refs["relatorios_periodo_label"] = lbl_periodo
    ui_refs["relatorios_registros_label"] = lbl_registros

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
        ("Pasta do relatório", selecionar_pasta_saida_relatorios),
        ("Buscar documento", abrir_busca_documentos),
        ("Gráfico de faturamento", abrir_grafico_faturamento),
    ]
    acoes_menu_relatorios = [
        ("Selecionar relatório", selecionar_relatorio),
        ("Abrir relatório", abrir_relatorio),
        None,
        ("Relatórios cancelados", abrir_relatorio_cancelados),
        ("Exportar configurações", exportar_configuracoes_ui),
        ("Importar configurações", importar_configuracoes_ui),
    ]
    acoes_menu_alteracoes = [
        ("Alterar competência", abrir_dialogo_alterar_competencia),
        ("Substituir documento", abrir_dialogo_substituir_documento),
        ("Cancelar documento", abrir_dialogo_cancelar_documento),
        ("Declarar intercompany", abrir_dialogo_declarar_intercompany),
        ("Declarar delta", abrir_dialogo_declarar_delta),
        ("Declarar spot", abrir_dialogo_declarar_spot),
    ]

    tabs = [
        ("principal", "Principal", acoes_menu_principal),
        ("relatorios", "Relatórios", acoes_menu_relatorios),
        ("alteracoes", "Alterações", acoes_menu_alteracoes),
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
        text="Escolha uma aba para acessar as ações",
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
        text="Resumo do período",
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
        text="Diferença referente a impostos de NFS-e: R$ 0,00",
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
        text="Análise gráfica",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.pack(anchor="w", padx=20, pady=(16, 4))
    ui_refs["graficos_titulo"] = titulo

    subtitulo = ctk.CTkLabel(
        graficos_card,
        text="Visão de faturamento por período e comparativo de documentos.",
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
    ui_refs["screen_nav_card"] = nav_card

    telas = [(sid, SCREEN_NAV_CATALOG.get(sid, sid.title())) for sid in SCREEN_ORDER_ATUAL]
    for idx in range(len(telas)):
        nav_card.grid_columnconfigure(idx, weight=1)
    nav_card.grid_columnconfigure(len(telas), weight=0)

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

    btn_reordenar = ctk.CTkButton(
        nav_card,
        text="Organizar botões",
        width=220,
        height=38,
        corner_radius=12,
        border_width=1,
        fg_color=UI_THEME["surface_alt"],
        border_color=UI_THEME["border"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        command=abrir_dialogo_reordenar_interface,
    )
    btn_reordenar.grid(row=0, column=len(telas), padx=(0, 10), pady=10, sticky="e")
    ui_refs["btn_reordenar_interface"] = btn_reordenar


def _criar_tela_dashboard(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    _criar_filtro_periodo(tela)
    _criar_graficos_dashboard(tela)
    _criar_resumo_periodo(tela)
    return tela


def _obter_acoes_relatorios_ordenadas():
    catalogo = _catalogo_acoes_interface()
    permitidas = {
        "selecionar_relatorio",
        "abrir_relatorio",
        "exportar_relatorio_periodo",
        "pasta_saida",
        "relatorio_cancelados",
        "buscar_documento",
        "grafico_faturamento",
    }
    ids = [
        acao_id
        for acao_id in ACTION_LAYOUT_ATUAL.get("relatorios", [])
        if acao_id in catalogo and acao_id in permitidas
    ]
    if not ids:
        ids = [
            acao_id
            for acao_id in _layout_botoes_padrao().get("relatorios", [])
            if acao_id in catalogo and acao_id in permitidas
        ]
    for acao_id in ("selecionar_relatorio", "abrir_relatorio", "exportar_relatorio_periodo"):
        if acao_id in catalogo and acao_id not in ids:
            ids.append(acao_id)
    return ids


def _criar_filtro_relatorios(parent):
    global relatorio_data_inicio_entry, relatorio_data_fim_entry

    card = _criar_card(parent, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 12))
    card.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
    ui_refs["relatorios_filtro_card"] = card

    titulo = ctk.CTkLabel(
        card,
        text="Período da aba Relatórios",
        font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
        text_color=UI_THEME["text_primary"],
    )
    titulo.grid(row=0, column=0, columnspan=6, sticky="w", padx=18, pady=(14, 4))
    ui_refs["relatorios_filtro_titulo"] = titulo

    subtitulo = ctk.CTkLabel(
        card,
        text="Esse período afeta a consulta, a abertura e a exportação do relatório.",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    )
    subtitulo.grid(row=1, column=0, columnspan=6, sticky="w", padx=18, pady=(0, 10))
    ui_refs["relatorios_filtro_subtitulo"] = subtitulo

    icon_inicio = ctk.CTkLabel(card, text="📅", font=ctk.CTkFont(size=15), text_color=UI_THEME["text_secondary"])
    icon_inicio.grid(row=2, column=0, sticky="e", padx=(18, 6), pady=(0, 14))
    ui_refs["relatorios_filtro_icon_inicio"] = icon_inicio

    relatorio_data_inicio_entry = ctk.CTkEntry(
        card,
        height=42,
        corner_radius=12,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    relatorio_data_inicio_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(0, 14))

    lbl_ate = ctk.CTkLabel(card, text="até", font=ctk.CTkFont(family="Segoe UI", size=12), text_color=UI_THEME["text_secondary"])
    lbl_ate.grid(row=2, column=2, padx=4, pady=(0, 14))
    ui_refs["relatorios_filtro_ate"] = lbl_ate

    icon_fim = ctk.CTkLabel(card, text="📅", font=ctk.CTkFont(size=15), text_color=UI_THEME["text_secondary"])
    icon_fim.grid(row=2, column=3, sticky="e", padx=(8, 6), pady=(0, 14))
    ui_refs["relatorios_filtro_icon_fim"] = icon_fim

    relatorio_data_fim_entry = ctk.CTkEntry(
        card,
        height=42,
        corner_radius=12,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        border_width=1,
        font=ctk.CTkFont(family="Segoe UI", size=13),
    )
    relatorio_data_fim_entry.grid(row=2, column=4, sticky="ew", padx=(0, 8), pady=(0, 14))

    btn_aplicar = ctk.CTkButton(
        card,
        text="Aplicar filtro",
        width=126,
        height=42,
        corner_radius=12,
        border_width=1,
        border_color=UI_THEME["border"],
        fg_color=UI_THEME["surface_alt"],
        hover_color=UI_THEME["tab_hover"],
        text_color=UI_THEME["text_primary"],
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        command=aplicar_filtro_relatorios,
    )
    btn_aplicar.grid(row=2, column=5, sticky="e", padx=(6, 18), pady=(0, 14))
    ui_refs["relatorios_aplicar_filtro_btn"] = btn_aplicar
    _aplicar_microinteracao_botao(btn_aplicar, "secondary")

    inicializar_filtros_relatorios()

    relatorio_data_inicio_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(relatorio_data_inicio_entry))
    relatorio_data_fim_entry.bind("<Button-1>", lambda _e: abrir_seletor_data(relatorio_data_fim_entry))
    relatorio_data_inicio_entry.bind("<FocusOut>", lambda _e: _atualizar_status_relatorios_ui())
    relatorio_data_fim_entry.bind("<FocusOut>", lambda _e: _atualizar_status_relatorios_ui())
    relatorio_data_inicio_entry.bind("<Return>", lambda _e: aplicar_filtro_relatorios(mensagem_sucesso=False))
    relatorio_data_fim_entry.bind("<Return>", lambda _e: aplicar_filtro_relatorios(mensagem_sucesso=False))


def _criar_bloco_relatorios_principal(parent, acoes_relatorios):
    card = _criar_card(parent, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 12))

    ctk.CTkLabel(
        card,
        text="Relatórios",
        font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 4))

    ctk.CTkLabel(
        card,
        text="Fluxo principal: selecionar, abrir e exportar o relatório filtrado pelo período desta aba.",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color=UI_THEME["text_secondary"],
    ).pack(anchor="w", padx=18, pady=(0, 12))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1, 2), weight=1)

    catalogo = _catalogo_acoes_interface()
    principal_ordem = ["selecionar_relatorio", "abrir_relatorio", "exportar_relatorio_periodo"]
    principal_ids = [aid for aid in principal_ordem if aid in acoes_relatorios]
    for idx, acao_id in enumerate(principal_ids):
        item = catalogo.get(acao_id)
        if not item:
            continue
        _criar_botao_acao(
            grade,
            item["titulo"],
            item["comando"],
            variant="primary",
            height=52,
        ).grid(row=0, column=idx, padx=6, pady=6, sticky="ew")


def _criar_bloco_relatorios_complementar(parent, acoes_relatorios):
    card = _criar_card(parent, corner_radius=18)
    card.pack(fill="x", padx=18, pady=(0, 12))

    ctk.CTkLabel(
        card,
        text="Ações complementares",
        font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(12, 4))

    ctk.CTkLabel(
        card,
        text="Ferramentas auxiliares para consulta e apoio.",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    ).pack(anchor="w", padx=18, pady=(0, 10))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)

    catalogo = _catalogo_acoes_interface()
    secundario_ids = [
        aid for aid in acoes_relatorios
        if aid in {"pasta_saida", "relatorio_cancelados", "buscar_documento", "grafico_faturamento"}
    ]

    linha = 0
    coluna = 0
    for acao_id in secundario_ids:
        item = catalogo.get(acao_id)
        if not item:
            continue
        _criar_botao_acao(
            grade,
            item["titulo"],
            item["comando"],
            variant="secondary",
            height=40,
        ).grid(row=linha, column=coluna, padx=6, pady=6, sticky="ew")
        coluna += 1
        if coluna > 1:
            coluna = 0
            linha += 1


def _criar_bloco_status_relatorios(parent):
    card = _criar_card(parent, corner_radius=16)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ui_refs["relatorios_status_card"] = card

    lbl_status = ctk.CTkLabel(
        card,
        text="Relatório carregado com sucesso",
        anchor="w",
        justify="left",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        text_color=UI_THEME["success_text"],
    )
    lbl_status.pack(fill="x", padx=18, pady=(12, 4))

    lbl_arquivo = ctk.CTkLabel(
        card,
        text="Arquivo: nenhum selecionado",
        anchor="w",
        justify="left",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color=UI_THEME["text_primary"],
    )
    lbl_arquivo.pack(fill="x", padx=18, pady=(0, 4))

    lbl_saida = ctk.CTkLabel(
        card,
        text="Pasta de saída: -",
        anchor="w",
        justify="left",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    )
    lbl_saida.pack(fill="x", padx=18, pady=(0, 4))

    lbl_periodo = ctk.CTkLabel(
        card,
        text="Período aplicado: período indefinido",
        anchor="w",
        justify="left",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    )
    lbl_periodo.pack(fill="x", padx=18, pady=(0, 12))

    lbl_registros = ctk.CTkLabel(
        card,
        text="Registros no período: -",
        anchor="w",
        justify="left",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        text_color=UI_THEME["text_secondary"],
    )
    lbl_registros.pack(fill="x", padx=18, pady=(0, 12))

    ui_refs["relatorios_status_label"] = lbl_status
    ui_refs["relatorios_arquivo_label"] = lbl_arquivo
    ui_refs["relatorios_saida_label"] = lbl_saida
    ui_refs["relatorios_periodo_label"] = lbl_periodo
    ui_refs["relatorios_registros_label"] = lbl_registros


def _criar_tela_relatorios(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    acoes_relatorios = _obter_acoes_relatorios_ordenadas()
    _criar_filtro_relatorios(tela)
    _criar_bloco_relatorios_principal(tela, acoes_relatorios)
    _criar_info_relatorio(tela)
    _criar_bloco_relatorios_complementar(tela, acoes_relatorios)
    _atualizar_status_relatorios_ui()
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

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 12))
    grade.grid_columnconfigure((0, 1), weight=1)
    tem_botoes = _renderizar_grade_acoes_por_tela(grade, "faturamento")

    if not tem_botoes:
        info = ctk.CTkLabel(
            card,
            text="Sem botões nesta aba. Use 'Trocar sequência dos botões' para organizar.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=UI_THEME["text_secondary"],
        )
        info.pack(anchor="w", padx=18, pady=(0, 16))
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
    tem_botoes = _renderizar_grade_acoes_por_tela(grade, "alteracoes")
    if not tem_botoes:
        ctk.CTkLabel(
            card,
            text="Sem botões nesta aba. Use 'Trocar sequência dos botões' para organizar.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=UI_THEME["text_secondary"],
        ).pack(anchor="w", padx=18, pady=(0, 16))
    return tela


def _criar_tela_configuracoes(parent):
    tela = ctk.CTkFrame(parent, fg_color=UI_THEME["app_bg"])
    card = _criar_card(tela, corner_radius=20)
    card.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(
        card,
        text="Configurações do sistema",
        font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        text_color=UI_THEME["text_primary"],
    ).pack(anchor="w", padx=18, pady=(14, 10))

    grade = ctk.CTkFrame(card, fg_color="transparent")
    grade.pack(fill="x", padx=16, pady=(0, 16))
    grade.grid_columnconfigure((0, 1), weight=1)
    tem_botoes = _renderizar_grade_acoes_por_tela(grade, "configuracoes")
    if not tem_botoes:
        ctk.CTkLabel(
            card,
            text="Sem botões nesta aba. Use 'Trocar sequência dos botões' para organizar.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=UI_THEME["text_secondary"],
        ).pack(anchor="w", padx=18, pady=(0, 16))
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
    _recarregar_layout_interface()

    container_scroll = ctk.CTkFrame(app, fg_color=UI_THEME["app_bg"], corner_radius=0)
    container_scroll.pack(fill="both", expand=True)
    ui_refs["container_shell"] = container_scroll

    scroll_canvas = tk.Canvas(
        container_scroll,
        highlightthickness=0,
        bd=0,
        bg=UI_THEME["app_bg"],
    )
    scroll_vertical = tk.Scrollbar(
        container_scroll,
        orient="vertical",
        jump=1,
        relief="flat",
        bd=0,
        highlightthickness=0,
        width=12,
        bg=UI_THEME["scroll_btn"],
        activebackground=UI_THEME["scroll_btn_hover"],
        troughcolor=UI_THEME["app_bg"],
    )
    scroll_state = {"job": None, "last_args": None, "dragging": False}

    def _executar_scroll(*args):
        try:
            if args and args[0] == "moveto":
                scroll_canvas.yview_moveto(float(args[1]))
            else:
                scroll_canvas.yview(*args)
        except Exception:
            return
        _forcar_redesenho_pos_scroll()

    def _flush_scroll():
        scroll_state["job"] = None
        args = scroll_state.get("last_args")
        if args:
            _executar_scroll(*args)

    def _comando_scroll(*args):
        scroll_state["last_args"] = args
        if scroll_state.get("dragging"):
            if scroll_state.get("job") is None:
                # Limita frequencia de redraw no arraste rapido para evitar ghosting.
                scroll_state["job"] = app.after(24, _flush_scroll)
            return
        _executar_scroll(*args)

    scroll_vertical.configure(command=_comando_scroll)
    scroll_canvas.configure(yscrollcommand=scroll_vertical.set)

    container_scroll.grid_columnconfigure(0, weight=1)
    container_scroll.grid_rowconfigure(0, weight=1)
    scroll_canvas.grid(row=0, column=0, sticky="nsew")
    scroll_vertical.grid(row=0, column=1, sticky="ns")
    ui_refs["scroll_canvas"] = scroll_canvas
    ui_refs["scrollbar_vertical"] = scroll_vertical

    main_frame = ctk.CTkFrame(scroll_canvas, fg_color=UI_THEME["app_bg"], corner_radius=0)
    scroll_window_id = scroll_canvas.create_window((0, 0), window=main_frame, anchor="nw")
    ui_refs["scroll_window_id"] = scroll_window_id

    def _atualizar_scrollregion(_event=None):
        try:
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        except Exception:
            pass

    def _ajustar_largura_frame(event):
        try:
            scroll_canvas.itemconfigure(scroll_window_id, width=event.width)
        except Exception:
            pass

    def _widget_no_mainframe(widget):
        atual = widget
        while atual is not None:
            if atual == main_frame:
                return True
            atual = getattr(atual, "master", None)
        return False

    def _rolar_mouse(event):
        if not _widget_no_mainframe(event.widget):
            return
        delta = int(getattr(event, "delta", 0))
        if delta == 0:
            return "break"
        passos = -int(delta / 120)
        if passos == 0:
            passos = -1 if delta > 0 else 1
        try:
            scroll_canvas.yview_scroll(passos, "units")
            _forcar_redesenho_pos_scroll()
        except Exception:
            pass
        return "break"

    def _press_scrollbar(_event=None):
        scroll_state["dragging"] = True

    def _release_scrollbar(_event=None):
        scroll_state["dragging"] = False
        if scroll_state.get("job") is not None:
            try:
                app.after_cancel(scroll_state["job"])
            except Exception:
                pass
            scroll_state["job"] = None
        _flush_scroll()

    main_frame.bind("<Configure>", _atualizar_scrollregion, add="+")
    scroll_canvas.bind("<Configure>", _ajustar_largura_frame, add="+")
    scroll_vertical.bind("<ButtonPress-1>", _press_scrollbar, add="+")
    scroll_vertical.bind("<ButtonRelease-1>", _release_scrollbar, add="+")
    if not getattr(app, "_bind_rolagem_global_principal", False):
        app.bind_all("<MouseWheel>", _rolar_mouse, add="+")
        setattr(app, "_bind_rolagem_global_principal", True)

    ui_refs["main_frame"] = main_frame

    _aplicar_logo_watermark(main_frame)
    _criar_navegacao_telas(main_frame)

    screen_host = ctk.CTkFrame(main_frame, fg_color=UI_THEME["app_bg"])
    screen_host.pack(fill="both", expand=True)
    app.screen_host = screen_host
    ui_refs["screen_host"] = screen_host

    dashboard_frame = _criar_tela_dashboard(screen_host)
    relatorios_frame = _criar_tela_relatorios(screen_host)
    alteracoes_frame = _criar_tela_alteracoes(screen_host)
    configuracoes_frame = _criar_tela_configuracoes(screen_host)

    app.registrar_tela("dashboard", dashboard_frame)
    app.registrar_tela("relatorios", relatorios_frame)
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
