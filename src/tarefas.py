"""Serviço local para quadro Kanban de tarefas."""

from datetime import datetime, date
import sqlite3

from src.banco import obter_conexao_banco


STATUS_TAREFA = ("A_FAZER", "EM_ANDAMENTO", "AGUARDANDO", "CONCLUIDO")
PRIORIDADES_TAREFA = ("BAIXA", "MEDIA", "ALTA", "URGENTE")
STATUS_LABELS = {
    "A_FAZER": "A Fazer",
    "EM_ANDAMENTO": "Em andamento",
    "AGUARDANDO": "Aguardando",
    "CONCLUIDO": "Concluído",
}
PRIORIDADE_LABELS = {
    "BAIXA": "Baixa",
    "MEDIA": "Média",
    "ALTA": "Alta",
    "URGENTE": "Urgente",
}
DEFAULT_CATEGORIAS = ("Seguros", "Documentos", "Financeiro", "Operação", "Pessoal", "Outros")


def _agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalizar_nome(texto):
    return " ".join(str(texto or "").strip().split())


def normalizar_status(status):
    status_norm = str(status or "").upper().strip()
    return status_norm if status_norm in STATUS_TAREFA else "A_FAZER"


def normalizar_prioridade(prioridade):
    prioridade_norm = str(prioridade or "").upper().strip()
    prioridade_norm = prioridade_norm.replace("É", "E")
    return prioridade_norm if prioridade_norm in PRIORIDADES_TAREFA else "MEDIA"


def parse_data_prazo(valor):
    texto = str(valor or "").strip()
    if not texto:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError("Prazo inválido. Use o formato dd/mm/aaaa.")


def formatar_prazo_br(valor):
    texto = str(valor or "").strip()
    if not texto:
        return "Sem prazo"
    try:
        return datetime.strptime(texto[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return texto


def classificar_prazo(valor, status):
    if normalizar_status(status) == "CONCLUIDO":
        return "concluido"
    texto = str(valor or "").strip()
    if not texto:
        return "sem_prazo"
    try:
        prazo = datetime.strptime(texto[:10], "%Y-%m-%d").date()
    except ValueError:
        return "sem_prazo"
    hoje = date.today()
    if prazo < hoje:
        return "atrasada"
    if prazo == hoje:
        return "hoje"
    return "futura"


def garantir_categorias_padrao():
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        agora = _agora()
        for nome in DEFAULT_CATEGORIAS:
            cursor.execute(
                """
                INSERT INTO tarefas_categorias (nome, ativo, data_cadastro)
                VALUES (?, 1, ?)
                ON CONFLICT(nome) DO NOTHING
                """,
                (nome, agora),
            )
        conn.commit()
    finally:
        conn.close()


def listar_categorias(ativas_apenas=True):
    garantir_categorias_padrao()
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        where = "WHERE ativo=1" if ativas_apenas else ""
        cursor.execute(f"SELECT id, nome, ativo FROM tarefas_categorias {where} ORDER BY nome COLLATE NOCASE")
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def adicionar_categoria(nome):
    nome_norm = _normalizar_nome(nome)
    if not nome_norm:
        raise ValueError("Informe o nome da categoria.")
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            INSERT INTO tarefas_categorias (nome, ativo, data_cadastro)
            VALUES (?, 1, ?)
            ON CONFLICT(nome) DO UPDATE SET ativo=1
            """,
            (nome_norm, _agora()),
        )
        conn.commit()
    finally:
        conn.close()
    return nome_norm


def _categoria_id_por_nome(cursor, nome):
    nome_norm = _normalizar_nome(nome)
    if not nome_norm:
        return None
    cursor.execute("SELECT id FROM tarefas_categorias WHERE nome=? COLLATE NOCASE", (nome_norm,))
    row = cursor.fetchone()
    if row:
        return int(row["id"] if isinstance(row, sqlite3.Row) else row[0])
    cursor.execute(
        "INSERT INTO tarefas_categorias (nome, ativo, data_cadastro) VALUES (?, 1, ?)",
        (nome_norm, _agora()),
    )
    return int(cursor.lastrowid)


def _proxima_ordem(cursor, status):
    cursor.execute(
        "SELECT COALESCE(MAX(ordem), 0) + 1 FROM tarefas WHERE status=? AND excluida=0",
        (normalizar_status(status),),
    )
    return int(cursor.fetchone()[0] or 1)


def criar_tarefa(titulo, descricao="", categoria="", responsavel="", prazo="", prioridade="MEDIA", status="A_FAZER", tags=""):
    titulo_norm = _normalizar_nome(titulo)
    if not titulo_norm:
        raise ValueError("Informe o título da tarefa.")
    status_norm = normalizar_status(status)
    prioridade_norm = normalizar_prioridade(prioridade)
    prazo_norm = parse_data_prazo(prazo)

    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        categoria_id = _categoria_id_por_nome(cursor, categoria)
        ordem = _proxima_ordem(cursor, status_norm)
        agora = _agora()
        concluida_em = agora if status_norm == "CONCLUIDO" else None
        cursor.execute(
            """
            INSERT INTO tarefas
            (titulo, descricao, categoria_id, responsavel, data_criacao, prazo, prioridade, status, ordem, data_atualizacao, concluida_em, tags)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                titulo_norm,
                str(descricao or "").strip(),
                categoria_id,
                str(responsavel or "").strip(),
                agora,
                prazo_norm,
                prioridade_norm,
                status_norm,
                ordem,
                agora,
                concluida_em,
                str(tags or "").strip(),
            ),
        )
        conn.commit()
        return int(cursor.lastrowid)
    finally:
        conn.close()


def atualizar_tarefa(tarefa_id, **campos):
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT * FROM tarefas WHERE id=? AND excluida=0", (int(tarefa_id),))
        atual = cursor.fetchone()
        if not atual:
            return 0
        atual = dict(atual)
        titulo = _normalizar_nome(campos.get("titulo", atual["titulo"]))
        if not titulo:
            raise ValueError("Informe o título da tarefa.")
        status = normalizar_status(campos.get("status", atual["status"]))
        prioridade = normalizar_prioridade(campos.get("prioridade", atual["prioridade"]))
        prazo = parse_data_prazo(campos.get("prazo", atual.get("prazo") or ""))
        categoria_id = _categoria_id_por_nome(cursor, campos.get("categoria", "")) if "categoria" in campos else atual.get("categoria_id")
        concluida_em = atual.get("concluida_em")
        if status == "CONCLUIDO" and not concluida_em:
            concluida_em = _agora()
        if status != "CONCLUIDO":
            concluida_em = None

        cursor.execute(
            """
            UPDATE tarefas
            SET titulo=?, descricao=?, categoria_id=?, responsavel=?, prazo=?, prioridade=?,
                status=?, data_atualizacao=?, concluida_em=?, tags=?
            WHERE id=?
            """,
            (
                titulo,
                str(campos.get("descricao", atual.get("descricao") or "") or "").strip(),
                categoria_id,
                str(campos.get("responsavel", atual.get("responsavel") or "") or "").strip(),
                prazo,
                prioridade,
                status,
                _agora(),
                concluida_em,
                str(campos.get("tags", atual.get("tags") or "") or "").strip(),
                int(tarefa_id),
            ),
        )
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def mover_tarefa(tarefa_id, novo_status):
    status_norm = normalizar_status(novo_status)
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        ordem = _proxima_ordem(cursor, status_norm)
        concluida_em = _agora() if status_norm == "CONCLUIDO" else None
        cursor.execute(
            """
            UPDATE tarefas
            SET status=?, ordem=?, data_atualizacao=?, concluida_em=?
            WHERE id=? AND excluida=0
            """,
            (status_norm, ordem, _agora(), concluida_em, int(tarefa_id)),
        )
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def excluir_tarefa(tarefa_id):
    conn = obter_conexao_banco()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE tarefas SET excluida=1, data_atualizacao=? WHERE id=?", (_agora(), int(tarefa_id)))
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def listar_tarefas(filtro="TODAS", busca="", categoria="Todas"):
    garantir_categorias_padrao()
    conn = obter_conexao_banco()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        params = []
        where = ["t.excluida=0"]
        filtro_norm = str(filtro or "TODAS").upper().strip()
        if filtro_norm in STATUS_TAREFA:
            where.append("t.status=?")
            params.append(filtro_norm)
        if categoria and str(categoria).strip().lower() != "todas":
            where.append("COALESCE(c.nome, '')=? COLLATE NOCASE")
            params.append(str(categoria).strip())
        busca_txt = str(busca or "").strip().lower()
        if busca_txt:
            where.append(
                "(LOWER(t.titulo) LIKE ? OR LOWER(COALESCE(t.descricao,'')) LIKE ? OR LOWER(COALESCE(t.responsavel,'')) LIKE ? OR LOWER(COALESCE(c.nome,'')) LIKE ? OR LOWER(COALESCE(t.tags,'')) LIKE ?)"
            )
            like = f"%{busca_txt}%"
            params.extend([like, like, like, like, like])

        cursor.execute(
            f"""
            SELECT t.*, COALESCE(c.nome, 'Sem categoria') AS categoria
            FROM tarefas t
            LEFT JOIN tarefas_categorias c ON c.id = t.categoria_id
            WHERE {' AND '.join(where)}
            ORDER BY t.status, t.ordem, t.id
            """,
            params,
        )
        tarefas = [dict(row) for row in cursor.fetchall()]

        if filtro_norm == "URGENTES":
            tarefas = [t for t in tarefas if normalizar_prioridade(t.get("prioridade")) == "URGENTE"]
        elif filtro_norm == "ATRASADAS":
            tarefas = [t for t in tarefas if classificar_prazo(t.get("prazo"), t.get("status")) == "atrasada"]
        elif filtro_norm == "HOJE":
            tarefas = [t for t in tarefas if classificar_prazo(t.get("prazo"), t.get("status")) == "hoje"]
        return tarefas
    finally:
        conn.close()


def resumo_tarefas():
    tarefas = listar_tarefas()
    resumo = {
        "total": len(tarefas),
        "a_fazer": 0,
        "em_andamento": 0,
        "aguardando": 0,
        "concluido": 0,
        "urgentes": 0,
        "atrasadas": 0,
    }
    for tarefa in tarefas:
        status = normalizar_status(tarefa.get("status"))
        if status == "A_FAZER":
            resumo["a_fazer"] += 1
        elif status == "EM_ANDAMENTO":
            resumo["em_andamento"] += 1
        elif status == "AGUARDANDO":
            resumo["aguardando"] += 1
        elif status == "CONCLUIDO":
            resumo["concluido"] += 1
        if normalizar_prioridade(tarefa.get("prioridade")) == "URGENTE":
            resumo["urgentes"] += 1
        if classificar_prazo(tarefa.get("prazo"), status) == "atrasada":
            resumo["atrasadas"] += 1
    return resumo
