from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def default_steps() -> list[dict[str, Any]]:
    return [
        {
            "id": "materiais",
            "label": "Materiais",
            "enabled": True,
            "action": "click",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "movimentacao",
            "label": "Movimentacao",
            "enabled": True,
            "action": "click",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "requisicao_compras",
            "label": "Requisicao de Compras",
            "enabled": True,
            "action": "click",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "manutencao",
            "label": "Manutencao",
            "enabled": True,
            "action": "click",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "buscar_requisicao",
            "label": "Buscar",
            "enabled": False,
            "action": "click",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "fornecedor_f2",
            "label": "Campo Fornecedor + F2",
            "enabled": False,
            "action": "click_f2",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 1,
            "delay_after": 1.0,
        },
        {
            "id": "campo_cnpj",
            "label": "Campo CNPJ (Pesquisa de Parceiros)",
            "enabled": False,
            "action": "fill_identifier",
            "identifier_type": "cnpj",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 2,
            "delay_after": 1.0,
        },
        {
            "id": "campo_cpf",
            "label": "Campo CPF (Pesquisa de Parceiros)",
            "enabled": False,
            "action": "fill_identifier",
            "identifier_type": "cpf",
            "template_path": None,
            "match_threshold": 0.84,
            "capture_size": 120,
            "x": None,
            "y": None,
            "clicks": 2,
            "delay_after": 1.0,
        },
    ]


def default_config() -> dict[str, Any]:
    return {
        "version": 1,
        "flow_name": "Criacao de Requisicao",
        "flow_path_text": "Materiais > Movimentacao > Requisicao de Compras > Manutencao",
        "app_monitor_index": 2,
        "target_monitor_index": 1,
        "pre_start_delay": 3.0,
        "default_step_delay": 1.0,
        "last_inputs": {
            "observacao": "",
            "cnpj": "",
            "cpf": "",
            "codigo_fornecedor": "",
            "linhas": "",
        },
        "steps": default_steps(),
    }


class ConfigStore:
    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path

    def load(self) -> dict[str, Any]:
        if not self.file_path.exists():
            config = default_config()
            self.save(config)
            return config

        # Accept JSON files saved with BOM by Windows editors (utf-8-sig)
        with self.file_path.open("r", encoding="utf-8-sig") as fp:
            loaded = json.load(fp)
        return normalize_config(loaded)

    def save(self, config: dict[str, Any]) -> None:
        normalized = normalize_config(config)
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        with self.file_path.open("w", encoding="utf-8") as fp:
            json.dump(normalized, fp, indent=2, ensure_ascii=False)


def normalize_config(raw: dict[str, Any]) -> dict[str, Any]:
    base = default_config()

    for key in (
        "version",
        "flow_name",
        "flow_path_text",
        "app_monitor_index",
        "target_monitor_index",
        "pre_start_delay",
        "default_step_delay",
    ):
        if key in raw:
            base[key] = raw[key]

    if isinstance(raw.get("last_inputs"), dict):
        base["last_inputs"] = base["last_inputs"] | raw["last_inputs"]

    steps_by_id = {step["id"]: step for step in raw.get("steps", []) if "id" in step}
    merged_steps: list[dict[str, Any]] = []
    for default_step in default_steps():
        saved_step = steps_by_id.get(default_step["id"], {})
        merged = default_step | saved_step
        merged_steps.append(merged)
    base["steps"] = merged_steps

    return base
