"""
Configuração centralizada de logging para o sistema de faturamento.

Uso nos módulos src/:
    from src.logger import get_logger
    logger = get_logger(__name__)
    logger.warning("mensagem %s", valor)

O arquivo de log é gravado em APP_DATA_DIR/faturamento.log
com rotação automática (máx. 1 MB × 3 backups).
"""

import logging
import logging.handlers
import os
import sys


def _caminho_log():
    """Resolve o caminho do arquivo de log sem importar src.config diretamente
    (evita importação circular caso config importe logger no futuro)."""
    try:
        import src.config as cfg
        return os.path.join(cfg.APP_DATA_DIR, "faturamento.log")
    except Exception:
        # Fallback: diretório do executável / script
        base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
        return os.path.join(base, "faturamento.log")


def _configurar_root_logger():
    root = logging.getLogger("faturamento")
    if root.handlers:
        return root  # já configurado

    root.setLevel(logging.DEBUG)

    # ── handler arquivo (rotativo) ─────────────────────────────────────
    try:
        caminho = _caminho_log()
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        fh = logging.handlers.RotatingFileHandler(
            caminho,
            maxBytes=1 * 1024 * 1024,  # 1 MB
            backupCount=3,
            encoding="utf-8",
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(
            logging.Formatter(
                "%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        root.addHandler(fh)
    except Exception:
        pass  # se não conseguir criar o arquivo, segue sem handler de arquivo

    # ── handler console (apenas WARNING+, só em modo dev) ─────────────
    if not getattr(sys, "frozen", False):
        ch = logging.StreamHandler()
        ch.setLevel(logging.WARNING)
        ch.setFormatter(logging.Formatter("%(levelname)s %(name)s: %(message)s"))
        root.addHandler(ch)

    return root


_configurar_root_logger()


def get_logger(name: str) -> logging.Logger:
    """
    Retorna um logger filho de 'faturamento' com o nome fornecido.
    Uso: logger = get_logger(__name__)
    """
    return logging.getLogger(f"faturamento.{name}")
