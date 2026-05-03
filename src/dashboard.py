"""
Lógica pura do dashboard — filtragem de dados e criação de figuras matplotlib.

Sem dependências de UI (CustomTkinter, messagebox, globals de widget).
O tema e o objeto plt são recebidos como parâmetros, tornando as funções
testáveis e independentes do estado global da aplicação.

Os orquestradores com estado UI (_garantir_matplotlib_dashboard,
_liberar_graficos_dashboard, _renderizar_figura_dashboard, atualizar_dashboard)
permanecem em sistema_faturamento.py.
"""

import sqlite3

import pandas as pd

from src.banco import obter_conexao_banco
from src.logger import get_logger
from src.utils import (
    MESES,
    _numero_documento_exibicao,
    _chave_documento_compativel,
    _competencia_para_data,
    _hex_para_rgb,
    _rgb_para_hex,
    formatar_moeda_brl_exata,
)

logger = get_logger(__name__)


# ─────────────────────────────────────────────
#  Filtragem de dados do dashboard
# ─────────────────────────────────────────────

def obter_dataframe_dashboard(data_inicial, data_final):
    """
    Retorna (df, data_inicial, data_final).
    df pode ser None (período inválido) ou DataFrame (possivelmente vazio).
    """
    if data_inicial is None or data_final is None:
        return None, data_inicial, data_final

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
        lambda r: _numero_documento_exibicao(
            r["tipo"], r["numero"], r.get("numero_original", ""), r.get("data_emissao", "")
        ),
        axis=1,
    )
    df["chave_documento"] = df.apply(
        lambda r: _chave_documento_compativel(r["tipo"], r["numero"], r.get("numero_original", "")),
        axis=1,
    )
    df = df.sort_values(["id"]).drop_duplicates(subset=["chave_documento"], keep="last")
    return df, data_inicial, data_final


# ─────────────────────────────────────────────
#  Figuras matplotlib
# ─────────────────────────────────────────────

def criar_figura_faturamento_periodo(df, plt, theme):
    """Cria e retorna a figura matplotlib de faturamento por período."""
    from matplotlib.colors import LinearSegmentedColormap
    from matplotlib.patches import FancyBboxPatch, Rectangle

    fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=100)
    fig.patch.set_facecolor(theme["chart_bg"])
    ax.set_facecolor(theme["chart_plot_bg"])

    coluna_valor = "valor_inicial"
    resumo = (
        df.groupby(df["data_competencia"].dt.to_period("M"))[coluna_valor]
        .sum()
        .sort_index()
        .reset_index()
    )
    resumo = resumo[resumo[coluna_valor].abs() > 0.0001].copy()

    if resumo.empty:
        ax.set_title("Faturamento por período", fontsize=11, fontweight="bold", color=theme["text_primary"], pad=8)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.text(
            0.5, 0.5,
            "Sem meses com faturamento para o período.",
            transform=ax.transAxes, ha="center", va="center",
            fontsize=10, color=theme["text_secondary"],
        )
        fig.tight_layout(pad=1.0)
        return fig

    meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    resumo["periodo_label"] = resumo["data_competencia"].dt.to_timestamp().apply(
        lambda dt: f"{meses_abrev[dt.month - 1]}/{dt.strftime('%y')}"
    )
    resumo["idx"] = range(len(resumo))
    valores = [float(v) for v in resumo[coluna_valor].fillna(0).tolist()]
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
            return [_misturar_cor(theme["chart_bar_secondary"], theme["chart_bar_primary"], 0.6)]
        return [
            _misturar_cor(theme["chart_bar_secondary"], theme["chart_bar_primary"], i / (qtd - 1))
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
        b_val = (soma_y - (m * soma_x)) / n
        return [(m * x) + b_val for x in xs]

    ax.set_title("Faturamento por período", fontsize=11, fontweight="bold", color=theme["text_primary"], pad=8)
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_xticks(indices)
    ax.set_xticklabels(
        resumo["periodo_label"], rotation=0, ha="center",
        color=theme["text_secondary"], fontsize=9.5, fontweight="bold",
    )
    ax.set_yticks([])
    ax.tick_params(axis="y", left=False, labelleft=False)
    ax.tick_params(axis="x", pad=6)
    ax.grid(False)
    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)
    ax.set_xlim(-0.5, len(indices) - 0.5)
    ax.set_ylim(0, limite_superior)
    ax.margins(x=0.08)

    fundo_topo = _misturar_cor(theme["chart_plot_bg"], "#FFFFFF", 0.06)
    fundo_base = _misturar_cor(theme["chart_plot_bg"], "#000000", 0.08)
    gradiente_bg = LinearSegmentedColormap.from_list("dashboard_bg_grad", [fundo_topo, fundo_base])
    ax.imshow(
        [[0], [1]],
        extent=[-0.6, len(indices) - 0.4, 0, limite_superior],
        aspect="auto", cmap=gradiente_bg, interpolation="bicubic",
        alpha=0.52, zorder=0,
    )

    cores = _gerar_cores_gradiente(len(valores))
    for idx_barra, (pos_x, altura, cor) in enumerate(zip(indices, valores, cores)):
        eh_atual = (idx_barra == len(indices) - 1)
        cor_base = theme["accent"] if eh_atual else cor
        larg = largura_barra * 1.12 if eh_atual else largura_barra

        esquerda = pos_x - (larg / 2)
        barra = FancyBboxPatch(
            (esquerda, 0), larg, max(altura, 0.0001),
            boxstyle=f"round,pad=0,rounding_size={larg * 0.20}",
            linewidth=0, facecolor=cor_base, edgecolor=cor_base,
            alpha=1.0 if eh_atual else 0.90, zorder=2,
        )
        ax.add_patch(barra)

        # Halo de destaque no mês atual
        if eh_atual:
            halo = FancyBboxPatch(
                (esquerda - 0.018, -limite_superior * 0.012), larg + 0.036, max(altura, 0.0001) + limite_superior * 0.024,
                boxstyle=f"round,pad=0,rounding_size={larg * 0.22}",
                linewidth=0, facecolor=cor_base, alpha=0.18, zorder=1.5,
            )
            ax.add_patch(halo)

        brilho_topo = Rectangle(
            (esquerda, max(altura * 0.42, 0)), larg, max(altura * 0.58, 0.0001),
            linewidth=0, facecolor=_misturar_cor(cor_base, "#FFFFFF", 0.32),
            alpha=0.40 if eh_atual else 0.36, zorder=2.25,
        )
        brilho_topo.set_clip_path(barra)
        ax.add_patch(brilho_topo)

    tendencia = _calcular_tendencia_linear(valores)
    if tendencia:
        ax.plot(
            indices, tendencia,
            color=theme["text_secondary"], linewidth=1.1,
            linestyle="--", alpha=0.45, zorder=3,
        )

    for pos_x, valor in zip(indices, valores):
        if valor <= 0:
            continue
        y_texto = valor * 0.52
        alinhamento_vertical = "center"
        cor_texto = theme["on_accent"]
        tamanho_fonte = 8.5
        if valor < (limite_superior * 0.11):
            y_texto = valor + (limite_superior * 0.02)
            alinhamento_vertical = "bottom"
            cor_texto = theme["text_primary"]
            tamanho_fonte = 8
        ax.text(
            pos_x, y_texto,
            formatar_moeda_brl_exata(valor),
            ha="center", va=alinhamento_vertical,
            fontsize=tamanho_fonte + 0.6,
            color=cor_texto, fontweight="bold", zorder=4,
        )

    fig.tight_layout(pad=1.0)
    return fig


def criar_figura_comparativo_tipos(df, plt, theme):
    """Donut chart — NF vs CTE vs Cancelados."""
    fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=100)
    fig.patch.set_facecolor(theme["chart_bg"])
    ax.set_facecolor(theme["chart_bg"])

    total_nf = int((df["tipo"] == "NF").sum())
    total_cte = int((df["tipo"] == "CTE").sum())
    total_cancelados = int(df["status"].str.upper().str.contains("CANCELADO", na=False).sum())
    total = total_nf + total_cte + total_cancelados

    ax.set_title("Comparativo de documentos", fontsize=11, fontweight="bold",
                 color=theme["text_primary"], pad=8)

    if total == 0:
        ax.text(0.5, 0.5, "Sem documentos", transform=ax.transAxes,
                ha="center", va="center", fontsize=10, color=theme["text_secondary"])
        ax.axis("off")
        fig.tight_layout(pad=1.0)
        return fig

    labels = ["NF", "CTE", "Cancelados"]
    valores = [total_nf, total_cte, total_cancelados]
    cores = [
        theme["metric_nf_value"],
        theme["metric_cte_value"],
        theme["chart_cancelados"],
    ]

    # Remove fatias zero para evitar artefatos
    dados = [(l, v, c) for l, v, c in zip(labels, valores, cores) if v > 0]
    if not dados:
        ax.axis("off")
        fig.tight_layout(pad=1.0)
        return fig

    lbs, vals, cors = zip(*dados)

    wedge_props = {"linewidth": 2.5, "edgecolor": theme["chart_bg"]}
    wedges, _ = ax.pie(
        vals,
        colors=cors,
        startangle=90,
        wedgeprops=wedge_props,
        radius=1.0,
    )

    # Furo central — transforma em donut
    centro = plt.Circle((0, 0), 0.58, color=theme["chart_bg"])
    ax.add_patch(centro)

    # Texto central: total de documentos
    ax.text(0, 0.10, str(total), ha="center", va="center",
            fontsize=22, fontweight="bold", color=theme["text_primary"])
    ax.text(0, -0.18, "docs", ha="center", va="center",
            fontsize=9, color=theme["text_secondary"])

    # Legenda lateral compacta
    legenda_x = 1.22
    legenda_y_inicio = 0.38
    passo = 0.28
    for i, (lbl, val, cor) in enumerate(zip(lbs, vals, cors)):
        pct = val / total * 100
        y = legenda_y_inicio - i * passo
        ax.add_patch(plt.Circle((legenda_x - 0.13, y + 0.04), 0.07,
                                color=cor, transform=ax.transData, zorder=5))
        ax.text(legenda_x, y + 0.04, f"{lbl}  {val}",
                ha="left", va="center", fontsize=10,
                fontweight="bold", color=theme["text_primary"])
        ax.text(legenda_x, y - 0.10, f"{pct:.1f}%",
                ha="left", va="center", fontsize=8.5,
                color=theme["text_secondary"])

    ax.set_xlim(-1.4, 1.9)
    ax.set_ylim(-1.2, 1.2)
    ax.axis("equal")
    ax.axis("off")
    fig.tight_layout(pad=0.6)
    return fig
