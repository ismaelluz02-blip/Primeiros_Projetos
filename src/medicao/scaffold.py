import os

STANDARD_FOLDERS = [
    'ACORDO COLETIVO',
    'Admissão-alocação',
    'DECLARAÇÕES',
    'Demissão-transferência',
    'FÉRIAS',
    'FOPAG e Comp. de Pgto',
    'INSS + FGTS',
    'PARCELAMENTOS',
    'Ponto',
    'SEGURO DE VIDA',
    'VA E VR',
    'VT',
]


def create_competencia_structure(parent_dir, competencia_name):
    comp_path = os.path.join(parent_dir, competencia_name.strip())
    os.makedirs(comp_path, exist_ok=True)
    created = []
    for folder in STANDARD_FOLDERS:
        fpath = os.path.join(comp_path, folder)
        os.makedirs(fpath, exist_ok=True)
        created.append(fpath)
    return comp_path, created
