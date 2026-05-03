import os
from datetime import datetime
from .utils import normalize


def _parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if hasattr(val, 'date'):
        return val.date()
    s = str(val).strip()
    if not s or s == 'None':
        return None
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_forca_trabalho(filepath):
    try:
        import openpyxl
    except ImportError:
        return [], "openpyxl não instalado. Execute: pip install openpyxl"

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        return [], f"Erro ao abrir planilha: {e}"

    employees = []
    header_found = False
    header_row_idx = 0

    # Find header row (look for "COLABORADOR" in column H = index 7)
    for i, row in enumerate(all_rows):
        if len(row) > 7 and row[7] is not None:
            cell_val = normalize(str(row[7]))
            if cell_val in ('colaborador', 'nome', 'funcionario'):
                header_found = True
                header_row_idx = i
                break
        # Also try: row with 'empresa' and 'contrato'
        if len(row) > 1:
            r = [normalize(str(c)) if c is not None else '' for c in row]
            if 'empresa' in r and 'contrato' in r:
                header_found = True
                header_row_idx = i
                break

    data_start = header_row_idx + 1 if header_found else 1

    for row in all_rows[data_start:]:
        if not any(c for c in row if c is not None):
            continue
        if len(row) <= 7:
            continue

        nome = row[7]
        if not nome or str(nome).strip() == '' or str(nome).strip() == 'None':
            continue

        nome_str = str(nome).strip().upper()
        if nome_str in ('COLABORADOR', 'NOME', 'FUNCIONARIO'):
            continue

        try:
            situacao = str(row[8]).strip().upper() if len(row) > 8 and row[8] is not None else 'A'
            cpf = str(row[11]).strip() if len(row) > 11 and row[11] is not None else ''
            funcao = str(row[10]).strip() if len(row) > 10 and row[10] is not None else ''

            ano_comp = row[5] if len(row) > 5 else None
            mes_comp = row[6] if len(row) > 6 else None

            data_admissao = _parse_date(row[16] if len(row) > 16 else None)
            data_demissao = _parse_date(row[17] if len(row) > 17 else None)

            afastado_raw = row[18] if len(row) > 18 else None
            afastado = str(afastado_raw).strip().upper() if afastado_raw is not None else 'NAO'

            emp = {
                'nome': nome_str,
                'situacao': situacao,
                'cpf': cpf,
                'funcao': funcao,
                'data_admissao': data_admissao,
                'data_demissao': data_demissao,
                'afastado': afastado,
                'ano_comp': ano_comp,
                'mes_comp': mes_comp,
            }
            employees.append(emp)
        except Exception:
            continue

    return employees, None


def get_month_employees(employees, year=None, month=None):
    if not employees:
        return []
    if year is None and month is None:
        return employees

    result = []
    for emp in employees:
        emp_year = emp.get('ano_comp')
        emp_month = emp.get('mes_comp')

        # Try to match by competencia columns
        try:
            emp_year_int = int(float(str(emp_year))) if emp_year else None
            emp_month_int = int(float(str(emp_month))) if emp_month else None
        except (ValueError, TypeError):
            emp_year_int = None
            emp_month_int = None

        if emp_year_int is None and emp_month_int is None:
            result.append(emp)
            continue

        if year and emp_year_int and emp_year_int != year:
            continue
        if month and emp_month_int and emp_month_int != month:
            continue

        result.append(emp)

    return result if result else employees
