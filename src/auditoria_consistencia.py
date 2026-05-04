"""Auditoria rapida de inconsistencias nos dados de faturamento."""

import sqlite3

from src.banco import obter_conexao_banco


def _doc_titulo(row):
    numero = row.get("numero_original") or row.get("numero") or "-"
    return f"{row.get('tipo') or '-'} {numero}"


def auditar_consistencia(limite_por_regra=20):
    limite = max(1, int(limite_por_regra or 20))
    problemas = []
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        regras = [
            (
                "SEM_COMPETENCIA",
                "Documento sem competencia",
                "alta",
                """
                SELECT * FROM documentos
                WHERE TRIM(COALESCE(competencia, ''))=''
                ORDER BY id DESC
                LIMIT ?
                """,
                lambda r: "Preencher competencia antes de exportar relatorios.",
            ),
            (
                "FRETE_VAZIO",
                "Documento sem modalidade de frete",
                "media",
                """
                SELECT * FROM documentos
                WHERE TRIM(COALESCE(frete, ''))=''
                ORDER BY id DESC
                LIMIT ?
                """,
                lambda r: "Conferir se e Franquia, Delta, Spot ou Intercompany.",
            ),
            (
                "VALOR_INICIAL_ZERADO",
                "Valor inicial zerado",
                "media",
                """
                SELECT * FROM documentos
                WHERE COALESCE(valor_inicial, 0)=0
                  AND COALESCE(cancelado_manual, 0)=0
                  AND UPPER(COALESCE(status, '')) NOT LIKE '%CANCEL%'
                ORDER BY id DESC
                LIMIT ?
                """,
                lambda r: "Conferir se o documento veio sem valor ou foi importado errado.",
            ),
            (
                "VALOR_FINAL_MAIOR_INICIAL",
                "Valor final maior que valor inicial",
                "baixa",
                """
                SELECT * FROM documentos
                WHERE COALESCE(valor_final, 0) > COALESCE(valor_inicial, 0)
                  AND COALESCE(valor_inicial, 0) > 0
                ORDER BY id DESC
                LIMIT ?
                """,
                lambda r: "Validar se a revisao manual ou regra de imposto esta correta.",
            ),
            (
                "CANCELADO_COM_VALOR",
                "Cancelado ainda com valor",
                "alta",
                """
                SELECT * FROM documentos
                WHERE (COALESCE(cancelado_manual, 0)=1 OR UPPER(COALESCE(status, '')) LIKE '%CANCEL%')
                  AND (COALESCE(valor_inicial, 0)<>0 OR COALESCE(valor_final, 0)<>0)
                ORDER BY id DESC
                LIMIT ?
                """,
                lambda r: "Cancelados normalmente nao deveriam somar no faturamento.",
            ),
        ]

        for codigo, titulo, severidade, sql, sugestao_fn in regras:
            cursor.execute(sql, (limite,))
            for row in cursor.fetchall():
                item = dict(row)
                problemas.append(
                    {
                        "codigo": codigo,
                        "titulo": titulo,
                        "severidade": severidade,
                        "documento": _doc_titulo(item),
                        "competencia": item.get("competencia") or "-",
                        "detalhe": f"Inicial: {item.get('valor_inicial') or 0} | Final: {item.get('valor_final') or 0} | Frete: {item.get('frete') or '-'} | Status: {item.get('status') or '-'}",
                        "sugestao": sugestao_fn(item),
                        "origem": "documentos",
                        "id": item.get("id"),
                    }
                )

        cursor.execute(
            """
            SELECT tipo, numero, COUNT(*) AS qtd, GROUP_CONCAT(id) AS ids
            FROM documentos
            WHERE numero IS NOT NULL AND TRIM(COALESCE(tipo, ''))<>''
            GROUP BY tipo, numero
            HAVING COUNT(*) > 1
            LIMIT ?
            """,
            (limite,),
        )
        for row in cursor.fetchall():
            problemas.append(
                {
                    "codigo": "DUPLICADO",
                    "titulo": "Documento duplicado",
                    "severidade": "alta",
                    "documento": f"{row['tipo']} {row['numero']}",
                    "competencia": "-",
                    "detalhe": f"{row['qtd']} registros com o mesmo tipo e numero. IDs: {row['ids']}",
                    "sugestao": "Conferir se houve importacao duplicada.",
                    "origem": "documentos",
                    "id": None,
                }
            )
    finally:
        conn.close()

    ordem = {"alta": 0, "media": 1, "baixa": 2}
    problemas.sort(key=lambda item: (ordem.get(item.get("severidade"), 9), item.get("titulo", "")))
    resumo = {"total": len(problemas), "alta": 0, "media": 0, "baixa": 0}
    for item in problemas:
        sev = item.get("severidade")
        if sev in resumo:
            resumo[sev] += 1
    return {"resumo": resumo, "problemas": problemas}
