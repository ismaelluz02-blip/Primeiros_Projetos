"""Cliente Playwright para consulta de notas emitidas no Portal Nacional da NFS-e.

Modulo preparado para:
- abrir navegador
- permitir autenticacao manual
- navegar para "Notas Emitidas"
- consultar por periodo
- capturar dados exibidos na tela
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Callable
import importlib
import re
import unicodedata

from .portal_processing import (
    filtrar_notas_canceladas as _filtrar_notas_canceladas_processamento,
    processar_notas_canceladas_portal,
)


def _coerce_date(valor: Any) -> date:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        txt = valor.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(txt[:19], fmt).date()
            except ValueError:
                continue
    raise ValueError(f"Data invalida para consulta: {valor!r}")


def _formatar_data_portal(valor: Any) -> str:
    return _coerce_date(valor).strftime("%d/%m/%Y")


def _normalizar_chave_coluna(texto: str) -> str:
    base = re.sub(r"\s+", " ", str(texto or "").strip().lower())
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base or "coluna"


def _normalizar_texto_comparacao(texto: Any) -> str:
    valor = unicodedata.normalize("NFKD", str(texto or ""))
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"\s+", " ", valor).strip().upper()
    return valor


def _coletar_valor_por_chaves(registro: dict[str, str], chaves_preferenciais: tuple[str, ...]) -> str:
    for chave in chaves_preferenciais:
        if chave in registro:
            valor = str(registro.get(chave) or "").strip()
            if valor:
                return valor
    return ""


def _coletar_valor_por_fragmentos(registro: dict[str, str], fragmentos: tuple[str, ...]) -> str:
    for chave, valor in registro.items():
        chave_norm = _normalizar_chave_coluna(chave)
        if any(fragmento in chave_norm for fragmento in fragmentos):
            txt = str(valor or "").strip()
            if txt:
                return txt
    return ""


def _status_indica_cancelamento(status: str) -> bool:
    txt = _normalizar_texto_comparacao(status)
    return "CANCELAD" in txt


def _padronizar_data_saida(data_txt: str) -> str:
    txt = str(data_txt or "").strip()
    if not txt:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:19], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return txt


def _extrair_numero_nota(registro: dict[str, str]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "numero",
            "numero_nfse",
            "numero_nota",
            "n_nfse",
            "nro_nfse",
            "num_nfse",
            "n_da_nota",
            "n_nota",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("numero", "nota", "nfse"))
    if not valor:
        return ""
    numeros = re.findall(r"\d+", valor)
    if numeros:
        return numeros[-1]
    return valor.strip()


def _extrair_data_nota(registro: dict[str, str]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "data",
            "data_emissao",
            "dt_emissao",
            "emissao",
            "data_de_emissao",
            "data_da_emissao",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("data", "emissao"))
    return _padronizar_data_saida(valor)


def _extrair_chave_nota(registro: dict[str, str]) -> str | None:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "chave",
            "chave_acesso",
            "chave_nfse",
            "codigo_verificacao",
            "cod_verificacao",
            "codigo",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("chave", "verificacao", "codigo"))
    valor = str(valor or "").strip()
    return valor if valor else None


def _extrair_status_nota(registro: dict[str, str]) -> str:
    valor = _coletar_valor_por_chaves(
        registro,
        (
            "status",
            "situacao",
            "situacao_nfse",
            "status_nfse",
            "estado",
        ),
    )
    if not valor:
        valor = _coletar_valor_por_fragmentos(registro, ("status", "situacao"))
    return str(valor or "").strip()


def filtrar_notas_canceladas(registros: list[dict[str, str]]) -> list[dict[str, str | None]]:
    """Wrapper de compatibilidade para etapa isolada de processamento."""
    return _filtrar_notas_canceladas_processamento(registros)


@dataclass(slots=True)
class NfsePeriodoConsulta:
    data_inicio: Any
    data_fim: Any

    @property
    def inicio_date(self) -> date:
        return _coerce_date(self.data_inicio)

    @property
    def fim_date(self) -> date:
        return _coerce_date(self.data_fim)

    @property
    def inicio_portal(self) -> str:
        return _formatar_data_portal(self.data_inicio)

    @property
    def fim_portal(self) -> str:
        return _formatar_data_portal(self.data_fim)


@dataclass(slots=True)
class NfsePortalSelectors:
    menu_notas_emitidas: tuple[str, ...] = (
        "a:has-text('Notas Emitidas')",
        "button:has-text('Notas Emitidas')",
        "[role='menuitem']:has-text('Notas Emitidas')",
        "text=Notas Emitidas",
    )
    campo_data_inicio: tuple[str, ...] = (
        "input[name='dataInicio']",
        "input[id*='dataInicio']",
        "input[placeholder*='Data inicial']",
        "input[aria-label*='Data inicial']",
    )
    campo_data_fim: tuple[str, ...] = (
        "input[name='dataFim']",
        "input[id*='dataFim']",
        "input[placeholder*='Data final']",
        "input[aria-label*='Data final']",
    )
    botao_consultar: tuple[str, ...] = (
        "button:has-text('Consultar')",
        "button:has-text('Pesquisar')",
        "button:has-text('Buscar')",
        "[role='button']:has-text('Consultar')",
    )
    tabela_notas_emitidas: tuple[str, ...] = (
        "table:has(thead th)",
        "table",
        "[role='table']",
    )
    linha_tabela: tuple[str, ...] = (
        "tbody tr",
        "[role='row']",
    )
    estado_logado: tuple[str, ...] = (
        "text=Notas Emitidas",
        "text=Emitidas",
        "text=Consultar",
    )
    aviso_sem_dados: tuple[str, ...] = (
        "text=Nenhum registro encontrado",
        "text=Sem dados",
        "text=Nao ha dados",
    )


@dataclass(slots=True)
class NfsePortalConfig:
    base_url: str
    notas_emitidas_url: str | None = None
    browser: str = "chromium"  # chromium | firefox | webkit
    headless: bool = False
    timeout_ms: int = 30_000
    auth_manual_wait_ms: int = 120_000
    selectors: NfsePortalSelectors = field(default_factory=NfsePortalSelectors)


@dataclass(slots=True)
class ConsultaNfseResultado:
    periodo: NfsePeriodoConsulta
    registros: list[dict[str, str]]
    total_registros: int
    url_final: str
    coletado_em: str
    html_tabela: str = ""


class NfsePortalClient:
    """Cliente reutilizavel para consulta do Portal Nacional NFS-e via Playwright."""

    def __init__(self, config: NfsePortalConfig):
        self.config = config
        self._playwright: Any | None = None
        self._browser: Any | None = None
        self._pw_timeout_error: Any | None = None
        self.context: Any | None = None
        self.page: Any | None = None

    # ---------- ciclo de vida ----------
    def iniciar(self) -> None:
        if self.page is not None:
            return
        module = importlib.import_module("playwright.sync_api")
        self._pw_timeout_error = module.TimeoutError
        self._playwright = module.sync_playwright().start()
        browser_type = getattr(self._playwright, self.config.browser, None)
        if browser_type is None:
            raise ValueError(f"Browser Playwright invalido: {self.config.browser}")

        self._browser = browser_type.launch(headless=self.config.headless)
        self.context = self._browser.new_context(locale="pt-BR")
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.config.timeout_ms)

    def encerrar(self) -> None:
        if self.context is not None:
            try:
                self.context.close()
            except Exception:
                pass
        self.context = None
        self.page = None

        if self._browser is not None:
            try:
                self._browser.close()
            except Exception:
                pass
        self._browser = None

        if self._playwright is not None:
            try:
                self._playwright.stop()
            except Exception:
                pass
        self._playwright = None

    def __enter__(self) -> "NfsePortalClient":
        self.iniciar()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.encerrar()

    # ---------- util internos ----------
    def _garantir_page(self) -> None:
        if self.page is None:
            self.iniciar()

    def _esperar_primeiro_selector(
        self,
        selectors: tuple[str, ...],
        timeout_ms: int | None = None,
        state: str = "visible",
    ) -> str | None:
        self._garantir_page()
        timeout = timeout_ms if timeout_ms is not None else self.config.timeout_ms
        for selector in selectors:
            try:
                self.page.wait_for_selector(selector, timeout=timeout, state=state)
                return selector
            except Exception:
                continue
        return None

    def _preencher_campo_data(self, selectors: tuple[str, ...], valor: str) -> None:
        selector = self._esperar_primeiro_selector(selectors, timeout_ms=3_000)
        if not selector:
            raise RuntimeError(f"Campo de data nao encontrado. Selectors testados: {selectors}")
        campo = self.page.locator(selector).first
        campo.click()
        campo.fill("")
        campo.type(valor, delay=20)

    # ---------- fluxo portal ----------
    def abrir_portal(self) -> None:
        self._garantir_page()
        self.page.goto(self.config.base_url, wait_until="domcontentloaded")

    def aguardar_autenticacao_manual(
        self,
        callback_autenticacao: Callable[[Any], None] | None = None,
        timeout_ms: int | None = None,
    ) -> None:
        """Ponto de extensao para login manual/certificado.

        - callback_autenticacao(page): opcional para UI externa guiar o usuario.
        - sem callback, aguarda ate aparecer um estado "logado" configurado.
        """
        self._garantir_page()
        if callback_autenticacao:
            callback_autenticacao(self.page)

        timeout = timeout_ms if timeout_ms is not None else self.config.auth_manual_wait_ms
        selector_ok = self._esperar_primeiro_selector(
            self.config.selectors.estado_logado,
            timeout_ms=timeout,
            state="visible",
        )
        if not selector_ok:
            raise TimeoutError(
                "Autenticacao manual nao confirmada no prazo. "
                "Ajuste os seletores de estado logado ou aumente o timeout."
            )

    def navegar_para_notas_emitidas(self) -> None:
        self._garantir_page()

        if self.config.notas_emitidas_url:
            self.page.goto(self.config.notas_emitidas_url, wait_until="domcontentloaded")
        else:
            menu_selector = self._esperar_primeiro_selector(
                self.config.selectors.menu_notas_emitidas,
                timeout_ms=8_000,
            )
            if not menu_selector:
                raise RuntimeError("Nao foi possivel localizar o menu 'Notas Emitidas'.")
            self.page.locator(menu_selector).first.click()

        self.page.wait_for_load_state("domcontentloaded")
        # Aguarda filtro/tabela aparecer.
        ok_filtro = self._esperar_primeiro_selector(
            self.config.selectors.campo_data_inicio,
            timeout_ms=10_000,
        )
        ok_tabela = self._esperar_primeiro_selector(
            self.config.selectors.tabela_notas_emitidas,
            timeout_ms=5_000,
            state="attached",
        )
        if not ok_filtro and not ok_tabela:
            raise RuntimeError("Tela de Notas Emitidas nao carregou como esperado.")

    def consultar_periodo(self, periodo: NfsePeriodoConsulta) -> None:
        self._garantir_page()
        self._preencher_campo_data(self.config.selectors.campo_data_inicio, periodo.inicio_portal)
        self._preencher_campo_data(self.config.selectors.campo_data_fim, periodo.fim_portal)

        botao_selector = self._esperar_primeiro_selector(
            self.config.selectors.botao_consultar,
            timeout_ms=4_000,
        )
        if not botao_selector:
            raise RuntimeError("Botao de consulta nao encontrado na tela de Notas Emitidas.")

        self.page.locator(botao_selector).first.click()
        try:
            self.page.wait_for_load_state("networkidle", timeout=12_000)
        except Exception:
            # Alguns portais nao chegam em networkidle; segue com fallback.
            pass

    def capturar_dados_notas_emitidas(self, max_linhas: int | None = None) -> tuple[list[dict[str, str]], str]:
        self._garantir_page()

        # Se houver aviso de sem dados, retorna vazio com seguranca.
        sem_dados = self._esperar_primeiro_selector(
            self.config.selectors.aviso_sem_dados,
            timeout_ms=1_200,
        )
        if sem_dados:
            return [], ""

        tabela_selector = self._esperar_primeiro_selector(
            self.config.selectors.tabela_notas_emitidas,
            timeout_ms=6_000,
            state="attached",
        )
        if not tabela_selector:
            return [], ""

        tabela = self.page.locator(tabela_selector).first
        html_tabela = tabela.inner_html()

        headers_loc = tabela.locator("thead tr th")
        total_headers = headers_loc.count()
        headers: list[str] = []
        for i in range(total_headers):
            headers.append(headers_loc.nth(i).inner_text().strip())

        rows_loc = tabela.locator("tbody tr")
        total_rows = rows_loc.count()
        limite = min(total_rows, max_linhas) if isinstance(max_linhas, int) and max_linhas > 0 else total_rows

        registros: list[dict[str, str]] = []
        for idx in range(limite):
            row = rows_loc.nth(idx)
            cells = row.locator("td")
            total_cells = cells.count()
            if total_cells == 0:
                continue
            valores = [cells.nth(i).inner_text().strip() for i in range(total_cells)]

            if headers and len(headers) == len(valores):
                item = {_normalizar_chave_coluna(k): v for k, v in zip(headers, valores)}
            else:
                item = {f"coluna_{i+1}": valor for i, valor in enumerate(valores)}
            registros.append(item)

        return registros, html_tabela

    def filtrar_notas_canceladas(self, registros: list[dict[str, str]]) -> list[dict[str, str | None]]:
        return filtrar_notas_canceladas(registros)

    # ---------- orquestracao ----------
    def executar_consulta(
        self,
        data_inicio: Any,
        data_fim: Any,
        callback_autenticacao: Callable[[Any], None] | None = None,
        max_linhas: int | None = None,
    ) -> ConsultaNfseResultado:
        periodo = NfsePeriodoConsulta(data_inicio=data_inicio, data_fim=data_fim)
        if periodo.inicio_date > periodo.fim_date:
            raise ValueError("Data inicial nao pode ser maior que data final.")

        self._garantir_page()
        self.abrir_portal()
        self.aguardar_autenticacao_manual(callback_autenticacao=callback_autenticacao)
        self.navegar_para_notas_emitidas()
        self.consultar_periodo(periodo)
        registros, html_tabela = self.capturar_dados_notas_emitidas(max_linhas=max_linhas)

        return ConsultaNfseResultado(
            periodo=periodo,
            registros=registros,
            total_registros=len(registros),
            url_final=self.page.url if self.page is not None else "",
            coletado_em=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            html_tabela=html_tabela,
        )


__all__ = [
    "ConsultaNfseResultado",
    "NfsePeriodoConsulta",
    "NfsePortalClient",
    "NfsePortalConfig",
    "NfsePortalSelectors",
    "filtrar_notas_canceladas",
    "processar_notas_canceladas_portal",
]
