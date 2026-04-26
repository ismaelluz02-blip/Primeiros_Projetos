"""
Operações CRUD em documentos fiscais.

Funções de banco de dados para salvar, alterar, cancelar e substituir
documentos. Não contém lógica de UI — quem chama é responsável por
atualizar a interface após cada operação.

Para receber notificações de mudança, use register_on_change():
    import src.documentos as documentos
    documentos.register_on_change(atualizar_dashboard)
    documentos.register_on_change(_atualizar_cache_documentos_pos_alteracao)
"""

import re
import sqlite3

from src.banco import obter_conexao_banco
from src.utils import (
    _coletar_numero_original_para_match,
    _numero_para_texto,
    competencia_por_data,
    normalizar_texto,
)

# ─────────────────────────────────────────────
#  Callback registry — notifica mudanças sem acoplamento à UI
# ─────────────────────────────────────────────

_on_change_callbacks: list = []


def register_on_change(fn):
    """Registra uma função a ser chamada sempre que um documento for alterado."""
    if fn not in _on_change_callbacks:
        _on_change_callbacks.append(fn)


def _fire_on_change():
    for fn in _on_change_callbacks:
        try:
            fn()
        except Exception:
            pass


# ─────────────────────────────────────────────
#  Helpers internos
# ─────────────────────────────────────────────

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


def _coletar_ids_documentos_por_numero(cursor, tipo, numero):
    tipo_norm = str(tipo or "").upper().strip()
    numero_txt = str(numero or "").strip()
    numero_digitos = re.sub(r"\D", "", numero_txt)

    valores_texto = {v for v in (numero_txt, numero_digitos) if v}
    valores_int = set()
    for candidato in (numero_digitos, numero_txt):
        if not candidato:
            continue
        try:
            valores_int.add(int(candidato))
        except ValueError:
            pass

    ids_encontrados = []
    vistos = set()

    def _registrar(linhas):
        for linha in linhas:
            doc_id = int(linha["id"])
            if doc_id in vistos:
                continue
            vistos.add(doc_id)
            ids_encontrados.append(doc_id)

    if tipo_norm == "NF":
        if valores_int:
            placeholders = ",".join("?" for _ in valores_int)
            cursor.execute(
                f"SELECT id FROM documentos WHERE tipo='NF' AND numero IN ({placeholders})",
                tuple(valores_int),
            )
            _registrar(cursor.fetchall())

        if valores_texto:
            placeholders = ",".join("?" for _ in valores_texto)
            cursor.execute(
                f"SELECT id FROM documentos WHERE tipo='NF' AND numero_original IN ({placeholders})",
                tuple(valores_texto),
            )
            _registrar(cursor.fetchall())

        if valores_int:
            placeholders = ",".join("?" for _ in valores_int)
            cursor.execute(
                f"""
                SELECT id FROM documentos
                WHERE tipo='NF' AND CAST(numero_original AS INTEGER) IN ({placeholders})
                """,
                tuple(valores_int),
            )
            _registrar(cursor.fetchall())
    else:
        if not valores_int:
            return []
        numero_ref = next(iter(valores_int))
        cursor.execute(
            "SELECT id FROM documentos WHERE tipo=? AND numero=?",
            (tipo_norm, numero_ref),
        )
        _registrar(cursor.fetchall())

    return ids_encontrados


def _coletar_ids_documentos_para_frete(cursor, tipo, numero):
    ids_alvo = _coletar_ids_documentos_por_numero(cursor, tipo, numero)
    if not ids_alvo:
        return []

    placeholders = ",".join("?" for _ in ids_alvo)
    cursor.execute(
        f"""
        SELECT id, frete, COALESCE(frete_manual,0) AS frete_manual
             , COALESCE(frete_revisado_manual,0) AS frete_revisado_manual
        FROM documentos
        WHERE id IN ({placeholders})
        """,
        tuple(ids_alvo),
    )
    return [
        {
            "id": int(linha["id"]),
            "frete": str(linha["frete"] or "").upper().strip(),
            "frete_manual": int(linha["frete_manual"] or 0),
            "frete_revisado_manual": int(linha["frete_revisado_manual"] or 0),
        }
        for linha in cursor.fetchall()
    ]


def _normalizar_modalidade_frete(nova_modalidade):
    modalidade = normalizar_texto(str(nova_modalidade or "")).upper().strip()
    if modalidade in {"INTERCOMPANY", "DELTA", "SPOT", "FRANQUIA"}:
        return modalidade
    if "INTER" in modalidade and "COMPANY" in modalidade:
        return "INTERCOMPANY"
    if "DELTA" in modalidade:
        return "DELTA"
    if "SPOT" in modalidade:
        return "SPOT"
    return "FRANQUIA"


# ─────────────────────────────────────────────
#  Operações públicas
# ─────────────────────────────────────────────

def salvar_documento(doc, cursor=None):
    conn_externo = cursor is not None
    if not conn_externo:
        conn = obter_conexao_banco()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
    else:
        conn = None

    try:
        tipo_doc = str(doc.get("tipo", "")).upper()
        numero_chave = int(doc["numero"])
        numero_original_txt = str(doc.get("numero_original", "")).strip()

        if tipo_doc == "NF":
            numero_original_txt, _ = _coletar_numero_original_para_match(numero_original_txt, numero_chave)
            existente = _buscar_documento_existente_sync(cursor, tipo_doc, numero_chave, numero_original_txt)
            if existente:
                try:
                    numero_chave = int(existente["numero"])
                except (TypeError, ValueError, KeyError):
                    numero_chave = int(doc["numero"])
        elif not numero_original_txt:
            numero_original_txt = _numero_para_texto(numero_chave)

        cursor.execute(
            """
            INSERT INTO documentos
            (
                numero,numero_original,tipo,data_emissao,valor_inicial,valor_final,frete,status,competencia,
                valor_inicial_original,valor_final_original,status_original,cancelado_manual,competencia_manual,frete_manual,frete_revisado_manual
            )
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
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
                END,
                frete_revisado_manual=CASE
                    WHEN documentos.frete_revisado_manual=1 THEN documentos.frete_revisado_manual
                    ELSE excluded.frete_revisado_manual
                END
            """,
            (
                numero_chave,
                numero_original_txt,
                doc["tipo"],
                doc["data"].strftime("%d/%m/%Y"),
                doc["valor_inicial"],
                doc["valor_final"],
                doc["frete"],
                doc["status"],
                doc.get("competencia", competencia_por_data(doc["data"])),
                None, None, None, 0, 0,
                int(doc.get("frete_manual", 0) or 0),
                int(doc.get("frete_revisado_manual", 0) or 0),
            ),
        )
        if not conn_externo:
            conn.commit()
    finally:
        if not conn_externo:
            conn.close()


def alterar_competencia_documento(tipo, numero, mes_competencia, ano_competencia):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    competencia = f"{mes_competencia}/{ano_competencia}"
    ids_alvo = _coletar_ids_documentos_por_numero(cursor, tipo, numero)
    if not ids_alvo:
        conn.close()
        return 0

    placeholders = ",".join("?" for _ in ids_alvo)
    cursor.execute(
        f"UPDATE documentos SET competencia=?, competencia_manual=1 WHERE id IN ({placeholders})",
        (competencia, *ids_alvo),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()
    _fire_on_change()
    return alterados


def atualizar_modalidade_frete_documento(tipo, numero, nova_modalidade):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    tipo = str(tipo).upper().strip()
    modalidade_frete = _normalizar_modalidade_frete(nova_modalidade)
    frete_manual = 0 if modalidade_frete == "FRANQUIA" else 1
    docs_alvo = _coletar_ids_documentos_para_frete(cursor, tipo, numero)

    if not docs_alvo:
        conn.close()
        return {
            "ok": False,
            "encontrados": 0,
            "alterados": 0,
            "modalidade": modalidade_frete,
            "frete_manual": frete_manual,
            "ja_estavam_na_modalidade": 0,
        }

    alterados = 0
    for doc in docs_alvo:
        frete_atual = str(doc.get("frete") or "").upper().strip()
        frete_manual_atual = int(doc.get("frete_manual") or 0)
        frete_revisado_atual = int(doc.get("frete_revisado_manual") or 0)
        if (
            frete_atual == modalidade_frete
            and frete_manual_atual == frete_manual
            and frete_revisado_atual == 1
        ):
            continue
        cursor.execute(
            "UPDATE documentos SET frete=?, frete_manual=?, frete_revisado_manual=1 WHERE id=?",
            (modalidade_frete, frete_manual, int(doc["id"])),
        )
        alterados += cursor.rowcount

    encontrados = len(docs_alvo)
    conn.commit()
    conn.close()
    _fire_on_change()
    return {
        "ok": True,
        "encontrados": encontrados,
        "alterados": alterados,
        "modalidade": modalidade_frete,
        "frete_manual": frete_manual,
        "ja_estavam_na_modalidade": max(encontrados - alterados, 0),
    }


def declarar_documento_frete(tipo, numero, novo_frete):
    resultado = atualizar_modalidade_frete_documento(tipo, numero, novo_frete)
    return int(resultado.get("encontrados", 0))


def salvar_alteracao_frete_manual(tipo, numero, nova_modalidade):
    return atualizar_modalidade_frete_documento(tipo, numero, nova_modalidade)


def declarar_intercompany(tipo, numero):
    return salvar_alteracao_frete_manual(tipo, numero, "INTERCOMPANY")


def declarar_delta(tipo, numero):
    return salvar_alteracao_frete_manual(tipo, numero, "DELTA")


def declarar_spot(tipo, numero):
    return salvar_alteracao_frete_manual(tipo, numero, "SPOT")


def registrar_substituicao(tipo_antigo, numero_antigo, tipo_novo, numero_novo):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    status_novo = f"DOCUMENTO SUBSTITUINDO DOCUMENTO {numero_antigo} {tipo_antigo}"
    status_antigo = f"DOCUMENTO SUBSTITUIDO POR {numero_novo} {tipo_novo}"
    ids_novo = _coletar_ids_documentos_por_numero(cursor, tipo_novo, numero_novo)
    ids_antigo = _coletar_ids_documentos_por_numero(cursor, tipo_antigo, numero_antigo)

    if ids_novo:
        placeholders_novo = ",".join("?" for _ in ids_novo)
        cursor.execute(
            f"""
            UPDATE documentos
            SET status_original=COALESCE(status_original, status), status=?
            WHERE id IN ({placeholders_novo})
            """,
            (status_novo, *ids_novo),
        )
        novo_alterado = cursor.rowcount
    else:
        novo_alterado = 0

    if ids_antigo:
        placeholders_antigo = ",".join("?" for _ in ids_antigo)
        cursor.execute(
            f"""
            UPDATE documentos
            SET
                valor_inicial_original=COALESCE(valor_inicial_original, valor_inicial),
                valor_final_original=COALESCE(valor_final_original, valor_final),
                status_original=COALESCE(status_original, status),
                valor_inicial=0, valor_final=0, status=?
            WHERE id IN ({placeholders_antigo})
            """,
            (status_antigo, *ids_antigo),
        )
        antigo_alterado = cursor.rowcount
    else:
        antigo_alterado = 0

    conn.commit()
    conn.close()
    _fire_on_change()
    return novo_alterado, antigo_alterado


def desfazer_substituicao(tipo_antigo, numero_antigo, tipo_novo, numero_novo):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    ids_antigo = _coletar_ids_documentos_por_numero(cursor, tipo_antigo, numero_antigo)
    ids_novo = _coletar_ids_documentos_por_numero(cursor, tipo_novo, numero_novo)

    if ids_antigo:
        placeholders_antigo = ",".join("?" for _ in ids_antigo)
        cursor.execute(
            f"""
            UPDATE documentos
            SET
                valor_inicial=COALESCE(valor_inicial_original, valor_inicial),
                valor_final=COALESCE(valor_final_original, valor_final),
                status=COALESCE(status_original, 'OK'),
                valor_inicial_original=NULL, valor_final_original=NULL, status_original=NULL
            WHERE id IN ({placeholders_antigo}) AND UPPER(status) LIKE 'DOCUMENTO SUBSTITUIDO POR%'
            """,
            tuple(ids_antigo),
        )
        antigo_restaurado = cursor.rowcount
    else:
        antigo_restaurado = 0

    if ids_novo:
        placeholders_novo = ",".join("?" for _ in ids_novo)
        cursor.execute(
            f"""
            UPDATE documentos
            SET status=COALESCE(status_original, 'OK'), status_original=NULL
            WHERE id IN ({placeholders_novo}) AND UPPER(status) LIKE 'DOCUMENTO SUBSTITUINDO DOCUMENTO%'
            """,
            tuple(ids_novo),
        )
        novo_restaurado = cursor.rowcount
    else:
        novo_restaurado = 0

    conn.commit()
    conn.close()
    _fire_on_change()
    return antigo_restaurado, novo_restaurado


def cancelar_documento(tipo, numero):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    ids_alvo = _coletar_ids_documentos_por_numero(cursor, tipo, numero)
    if not ids_alvo:
        conn.close()
        return 0

    placeholders = ",".join("?" for _ in ids_alvo)
    cursor.execute(
        f"""
        UPDATE documentos
        SET
            valor_inicial_original=COALESCE(valor_inicial_original, valor_inicial),
            valor_final_original=COALESCE(valor_final_original, valor_final),
            status_original=COALESCE(status_original, status),
            valor_inicial=0, valor_final=0,
            status='CANCELADO MANUALMENTE', cancelado_manual=1
        WHERE id IN ({placeholders})
        """,
        tuple(ids_alvo),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()
    _fire_on_change()
    return alterados


def desfazer_cancelamento_documento(tipo, numero):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    ids_alvo = _coletar_ids_documentos_por_numero(cursor, tipo, numero)
    if not ids_alvo:
        conn.close()
        return 0

    placeholders = ",".join("?" for _ in ids_alvo)
    cursor.execute(
        f"""
        UPDATE documentos
        SET
            valor_inicial=COALESCE(valor_inicial_original, valor_inicial),
            valor_final=COALESCE(valor_final_original, valor_final),
            status=COALESCE(status_original, 'OK'),
            valor_inicial_original=NULL, valor_final_original=NULL, status_original=NULL,
            cancelado_manual=0
        WHERE id IN ({placeholders}) AND UPPER(status)='CANCELADO MANUALMENTE'
        """,
        tuple(ids_alvo),
    )
    alterados = cursor.rowcount
    conn.commit()
    conn.close()
    _fire_on_change()
    return alterados
