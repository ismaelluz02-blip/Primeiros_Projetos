"""
Cache em memória para o DataFrame de documentos.

Substitui o acesso direto ao dict global `dados_importados` nas funções
_obter_documentos_em_memoria e _atualizar_cache_documentos_pos_alteracao.

Uso:
    from src.cache import doc_cache

    df = doc_cache.get()           # retorna cópia do df em cache (ou None)
    doc_cache.set(df)              # armazena novo df
    doc_cache.invalidate()         # força recarga no próximo get()
    doc_cache.total                # int — número de documentos
    doc_cache.atualizado_em        # str — timestamp da última carga
"""

from datetime import datetime

import pandas as pd

from src.logger import get_logger

logger = get_logger(__name__)


class DocumentCache:
    """Cache simples para o DataFrame principal de documentos."""

    def __init__(self):
        self._df: pd.DataFrame | None = None
        self._total: int = 0
        self._atualizado_em: str = ""
        self._valido: bool = False

    # ── leitura ─────────────────────────────────────────────────────────

    @property
    def valido(self) -> bool:
        return self._valido and isinstance(self._df, pd.DataFrame)

    @property
    def total(self) -> int:
        return self._total

    @property
    def atualizado_em(self) -> str:
        return self._atualizado_em

    def get(self) -> "pd.DataFrame | None":
        """Retorna uma cópia do DataFrame em cache, ou None se inválido."""
        if self.valido:
            return self._df.copy()
        return None

    # ── escrita ─────────────────────────────────────────────────────────

    def set(self, df: pd.DataFrame) -> None:
        """Armazena o DataFrame e atualiza os metadados."""
        self._df = df
        self._total = int(len(df)) if isinstance(df, pd.DataFrame) else 0
        self._atualizado_em = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self._valido = True
        logger.debug("DocumentCache.set: %d documentos em cache", self._total)

    def invalidate(self) -> None:
        """Marca o cache como inválido — próximo get() retorna None."""
        self._valido = False
        logger.debug("DocumentCache.invalidate: cache invalidado")

    def reset(self) -> None:
        """Limpa completamente o cache."""
        self._df = None
        self._total = 0
        self._atualizado_em = ""
        self._valido = False
        logger.debug("DocumentCache.reset: cache limpo")

    # ── compatibilidade com dados_importados (dict legado) ───────────────

    def para_dict_legado(self) -> dict:
        """Retorna um dict compatível com o formato antigo de dados_importados."""
        return {
            "df_documentos": self._df,
            "total_documentos": self._total,
            "memoria_atualizada_em": self._atualizado_em,
        }


# Singleton — importado pelos módulos que precisam do cache
doc_cache = DocumentCache()
