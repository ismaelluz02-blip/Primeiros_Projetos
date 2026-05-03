from datetime import datetime

import pandas as pd
import pytest

from src.dashboard import criar_figura_faturamento_periodo


def _theme():
    return {
        "chart_bg": "#08121C",
        "chart_plot_bg": "#122031",
        "text_primary": "#F4F7FB",
        "text_secondary": "#AEBBCB",
        "on_accent": "#FFFFFF",
        "chart_bar_secondary": "#0F899E",
        "chart_bar_primary": "#20D4B0",
        "accent": "#20D4B0",
    }


def test_grafico_faturamento_usa_valor_inicial():
    plt = pytest.importorskip("matplotlib.pyplot")
    plt.switch_backend("Agg")

    df = pd.DataFrame(
        [
            {
                "data_competencia": datetime(2026, 3, 1),
                "valor_inicial": 300000.00,
                "valor_final": 295000.00,
            },
            {
                "data_competencia": datetime(2026, 3, 1),
                "valor_inicial": 25091.85,
                "valor_final": 25737.98,
            },
        ]
    )

    fig = criar_figura_faturamento_periodo(df, plt, _theme())
    try:
        textos = [texto.get_text() for ax in fig.axes for texto in ax.texts]
        assert any("325.091,85" in texto for texto in textos)
        assert not any("320.737,98" in texto for texto in textos)
    finally:
        plt.close(fig)
