from __future__ import annotations

import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any

from .config import ConfigStore
from .vision import (
    is_vision_available,
    locate_template_center,
    missing_vision_message,
    save_template_around_point,
)
from .windows_input import (
    get_mouse_position,
    left_click,
    list_monitors,
    press_backspace,
    press_f2,
    set_dpi_awareness,
    set_mouse_position,
    type_text,
)


class RequisicaoApp:
    def __init__(self, root: tk.Tk, store: ConfigStore, project_root: Path) -> None:
        self.root = root
        self.store = store
        self.project_root = project_root
        self.templates_dir = self.project_root / "templates"

        self.config_data = self.store.load()
        self.monitors = list_monitors()

        self.stop_event = threading.Event()
        self.is_running = False
        self.step_table: ttk.Treeview | None = None
        self.settings_window: tk.Toplevel | None = None
        self.flow_window: tk.Toplevel | None = None
        self.current_inputs: dict[str, str] | None = None
        self.step_delay_var = tk.StringVar(value="")

        self.status_var = tk.StringVar(value="Pronto para executar.")
        self.monitor_summary_var = tk.StringVar(value="Detectando monitores...")

        self.flow_name_var = tk.StringVar(value=str(self.config_data.get("flow_name", "Criacao de Requisicao")))
        self.flow_path_var = tk.StringVar(
            value=str(
                self.config_data.get(
                    "flow_path_text",
                    "Materiais > Movimentacao > Requisicao de Compras > Manutencao",
                )
            )
        )
        self.flow_active_var = tk.StringVar(value=f"Fluxo ativo: {self.flow_name_var.get()}")

        self.app_monitor_var = tk.StringVar(value=str(self.config_data.get("app_monitor_index", 2)))
        self.target_monitor_var = tk.StringVar(value=str(self.config_data.get("target_monitor_index", 1)))
        self.pre_start_var = tk.StringVar(value=str(self.config_data.get("pre_start_delay", 3.0)))
        self.default_delay_var = tk.StringVar(value=str(self.config_data.get("default_step_delay", 0.8)))

        self.palette = {
            "bg": "#eceff1",
            "card": "#d7dbdf",
            "text": "#1f2a37",
            "primary": "#3f8fd2",
            "primary_hover": "#2e7dbd",
            "success": "#2d7f48",
            "success_hover": "#23653a",
            "muted": "#9aa8b8",
            "muted_hover": "#8796a8",
        }

        self.root.title("Automacao de Requisicao - Visual Rodopar")
        self.root.geometry("500x300")
        self.root.minsize(460, 260)
        self.root.resizable(True, True)
        self.root.configure(bg=self.palette["bg"])

        if len(self.monitors) >= 2:
            secondary_index = next(
                (idx + 1 for idx, monitor in enumerate(self.monitors) if not monitor.get("is_primary")),
                2,
            )
            self.app_monitor_var.set(str(secondary_index))
            self.config_data["app_monitor_index"] = int(secondary_index)
        else:
            self.app_monitor_var.set("1")
            self.config_data["app_monitor_index"] = 1

        self._build_ui()
        self.refresh_monitor_labels()
        self.apply_app_monitor_position(initial=True)

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=self.palette["bg"])
        container.pack(fill="both", expand=True, padx=14, pady=12)

        top_row = tk.Frame(container, bg=self.palette["bg"])
        top_row.pack(fill="x")

        title_wrap = tk.Frame(top_row, bg=self.palette["bg"])
        title_wrap.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_wrap,
            text="Automacao de Requisicao",
            bg=self.palette["bg"],
            fg=self.palette["text"],
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")

        mode = "ATIVO" if is_vision_available() else "INDISPONIVEL"
        tk.Label(
            title_wrap,
            text=f"Modo imagem: {mode}",
            bg=self.palette["bg"],
            fg="#5a6673",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        actions_top = tk.Frame(top_row, bg=self.palette["bg"])
        actions_top.pack(side="right", anchor="ne")

        self._make_action_button(
            actions_top,
            text="Fluxo",
            command=self.open_flow_dialog,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
            padx=10,
        ).pack(side="right")
        self._make_action_button(
            actions_top,
            text="Configuracoes",
            command=self.open_settings_dialog,
            color=self.palette["muted"],
            hover=self.palette["muted_hover"],
            padx=10,
        ).pack(side="right", padx=(0, 8))

        main_card = tk.Frame(container, bg=self.palette["card"])
        main_card.pack(fill="both", expand=True, pady=(10, 8))

        tk.Label(
            main_card,
            textvariable=self.flow_active_var,
            bg=self.palette["card"],
            fg=self.palette["text"],
            anchor="w",
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=10,
        ).pack(fill="x")

        tk.Label(
            main_card,
            text="Use 'Fluxo' para treinar os passos e 'Configuracoes' para tela/tempos.",
            bg=self.palette["card"],
            fg="#425466",
            anchor="w",
            font=("Segoe UI", 9),
            padx=12,
            pady=6,
        ).pack(fill="x")

        run_row = tk.Frame(main_card, bg=self.palette["card"])
        run_row.pack(fill="x", padx=12, pady=(0, 12))

        self._make_action_button(
            run_row,
            text="Parar",
            command=self.stop_execution,
            color=self.palette["muted"],
            hover=self.palette["muted_hover"],
        ).pack(side="right")

        self.run_button = self._make_action_button(
            run_row,
            text="CRIAR REQUISICAO",
            command=self.run_requisicao_flow,
            color=self.palette["success"],
            hover=self.palette["success_hover"],
            bold=True,
            padx=16,
        )
        self.run_button.pack(side="right", padx=(0, 8))

        status_card = tk.Frame(container, bg=self.palette["card"])
        status_card.pack(fill="x")
        tk.Label(
            status_card,
            textvariable=self.status_var,
            bg=self.palette["card"],
            fg="#334155",
            anchor="w",
            font=("Segoe UI", 10),
            padx=12,
            pady=10,
        ).pack(fill="x")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _make_action_button(
        self,
        parent: tk.Misc,
        text: str,
        command,
        color: str,
        hover: str,
        bold: bool = False,
        padx: int = 12,
    ) -> tk.Button:
        font = ("Segoe UI", 10, "bold") if bold else ("Segoe UI", 10)
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg="white",
            activebackground=hover,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=padx,
            pady=7,
            font=font,
            cursor="hand2",
        )
        button.bind("<Enter>", lambda _evt: button.configure(bg=hover))
        button.bind("<Leave>", lambda _evt: button.configure(bg=color))
        return button
    def open_settings_dialog(self) -> None:
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.settings_window = window
        window.title("Configuracoes")
        window.geometry("650x470")
        window.minsize(620, 420)
        window.configure(bg=self.palette["bg"])
        window.transient(self.root)
        window.grab_set()

        wrapper = tk.Frame(window, bg=self.palette["bg"])
        wrapper.pack(fill="both", expand=True, padx=14, pady=12)

        tk.Label(
            wrapper,
            text="Lista de configuracoes",
            bg=self.palette["bg"],
            fg=self.palette["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")
        tk.Label(
            wrapper,
            text="- Configuracoes da Tela\n- Tempos",
            justify="left",
            bg=self.palette["bg"],
            fg="#4f5d6b",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(2, 8))

        tela_box = tk.Frame(wrapper, bg=self.palette["card"])
        tela_box.pack(fill="x", pady=(0, 8))
        tk.Label(
            tela_box,
            text="Configuracoes da Tela",
            bg=self.palette["card"],
            fg=self.palette["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

        tk.Label(
            tela_box,
            textvariable=self.monitor_summary_var,
            bg="#f4f5f6",
            fg="#374151",
            justify="left",
            anchor="w",
            font=("Consolas", 9),
            padx=10,
            pady=8,
            wraplength=600,
        ).pack(fill="x", padx=10, pady=(0, 8))

        grid = tk.Frame(tela_box, bg=self.palette["card"])
        grid.pack(fill="x", padx=10, pady=(0, 8))
        grid.grid_columnconfigure(1, weight=1)
        grid.grid_columnconfigure(3, weight=1)

        tk.Label(grid, text="Monitor do App", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w"
        )
        app_combo = ttk.Combobox(grid, textvariable=self.app_monitor_var, width=8, state="readonly")
        app_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))

        tk.Label(grid, text="Monitor alvo", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=0, column=2, sticky="w", padx=(18, 0)
        )
        target_combo = ttk.Combobox(
            grid,
            textvariable=self.target_monitor_var,
            width=8,
            state="readonly",
        )
        target_combo.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self._fill_monitor_combo_values(app_combo, target_combo)

        row_buttons = tk.Frame(tela_box, bg=self.palette["card"])
        row_buttons.pack(fill="x", padx=10, pady=(0, 10))
        self._make_action_button(
            row_buttons,
            text="Atualizar Monitores",
            command=lambda: self._reload_from_settings(app_combo, target_combo),
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left")
        self._make_action_button(
            row_buttons,
            text="Centralizar no Monitor",
            command=lambda: self.apply_app_monitor_position(initial=False),
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left", padx=(8, 0))

        tempo_box = tk.Frame(wrapper, bg=self.palette["card"])
        tempo_box.pack(fill="x", pady=(0, 10))
        tk.Label(
            tempo_box,
            text="Tempos",
            bg=self.palette["card"],
            fg=self.palette["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

        tempo_grid = tk.Frame(tempo_box, bg=self.palette["card"])
        tempo_grid.pack(fill="x", padx=10, pady=(0, 12))
        tk.Label(tempo_grid, text="Contagem inicial (s)", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w"
        )
        tk.Entry(tempo_grid, textvariable=self.pre_start_var, width=10, font=("Segoe UI", 10)).grid(
            row=0, column=1, sticky="w", padx=(8, 20)
        )

        tk.Label(tempo_grid, text="Delay por passo (s)", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=0, column=2, sticky="w"
        )
        tk.Entry(tempo_grid, textvariable=self.default_delay_var, width=10, font=("Segoe UI", 10)).grid(
            row=0, column=3, sticky="w", padx=(8, 0)
        )

        bottom = tk.Frame(wrapper, bg=self.palette["bg"])
        bottom.pack(fill="x")
        self._make_action_button(
            bottom,
            text="Fechar",
            command=self._close_settings_dialog,
            color=self.palette["muted"],
            hover=self.palette["muted_hover"],
        ).pack(side="right")
        self._make_action_button(
            bottom,
            text="Salvar",
            command=lambda: self._save_from_settings(close_after=False),
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="right", padx=(0, 8))

        window.protocol("WM_DELETE_WINDOW", self._close_settings_dialog)

    def _fill_monitor_combo_values(self, app_combo: ttk.Combobox, target_combo: ttk.Combobox) -> None:
        options = [str(idx) for idx in range(1, max(1, len(self.monitors)) + 1)]
        app_combo["values"] = options
        target_combo["values"] = options
        if self.app_monitor_var.get() not in options and options:
            self.app_monitor_var.set(options[0])
        if self.target_monitor_var.get() not in options and options:
            self.target_monitor_var.set(options[0])

    def _reload_from_settings(self, app_combo: ttk.Combobox, target_combo: ttk.Combobox) -> None:
        self.monitors = list_monitors()
        self.refresh_monitor_labels()
        self._fill_monitor_combo_values(app_combo, target_combo)
        self._set_status("Monitores atualizados.")

    def _save_from_settings(self, close_after: bool) -> None:
        if not self.save_config(silent=False):
            return
        if close_after:
            if self.settings_window is not None and self.settings_window.winfo_exists():
                self.settings_window.destroy()
                self.settings_window = None

    def _close_settings_dialog(self) -> None:
        if not self.save_config(silent=True):
            return
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.destroy()
        self.settings_window = None
    def open_flow_dialog(self) -> None:
        if self.flow_window is not None and self.flow_window.winfo_exists():
            self.flow_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.flow_window = window
        window.title("Fluxo")
        window.geometry("700x520")
        window.minsize(660, 460)
        window.configure(bg=self.palette["bg"])
        window.transient(self.root)
        window.grab_set()

        wrapper = tk.Frame(window, bg=self.palette["bg"])
        wrapper.pack(fill="both", expand=True, padx=14, pady=12)

        tk.Label(
            wrapper,
            text="Fluxo",
            bg=self.palette["bg"],
            fg=self.palette["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        head_box = tk.Frame(wrapper, bg=self.palette["card"])
        head_box.pack(fill="x", pady=(6, 8))
        form = tk.Frame(head_box, bg=self.palette["card"])
        form.pack(fill="x", padx=10, pady=10)

        tk.Label(form, text="Nome", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=self.flow_name_var, width=35, font=("Segoe UI", 10)).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )

        tk.Label(form, text="Caminho", bg=self.palette["card"], font=("Segoe UI", 10)).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        tk.Entry(form, textvariable=self.flow_path_var, width=70, font=("Segoe UI", 10)).grid(
            row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0)
        )

        steps_box = tk.Frame(wrapper, bg=self.palette["card"])
        steps_box.pack(fill="both", expand=True, pady=(0, 8))
        tk.Label(
            steps_box,
            text="Passos da Criacao de Requisicao",
            bg=self.palette["card"],
            fg=self.palette["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Flow.Treeview", font=("Segoe UI", 10), rowheight=24)
        style.configure("Flow.Treeview.Heading", font=("Segoe UI", 10, "bold"))

        table_wrap = tk.Frame(steps_box, bg=self.palette["card"])
        table_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("ordem", "passo", "ativo", "acao", "template", "fallback", "delay")
        self.step_table = ttk.Treeview(
            table_wrap,
            columns=columns,
            show="headings",
            height=7,
            style="Flow.Treeview",
        )
        self.step_table.heading("ordem", text="#")
        self.step_table.heading("passo", text="Passo")
        self.step_table.heading("ativo", text="Ativo")
        self.step_table.heading("acao", text="Acao")
        self.step_table.heading("template", text="Template")
        self.step_table.heading("fallback", text="Fallback XY")
        self.step_table.heading("delay", text="Delay")
        self.step_table.column("ordem", width=40, anchor="center")
        self.step_table.column("passo", width=190, anchor="w")
        self.step_table.column("ativo", width=55, anchor="center")
        self.step_table.column("acao", width=120, anchor="center")
        self.step_table.column("template", width=85, anchor="center")
        self.step_table.column("fallback", width=90, anchor="center")
        self.step_table.column("delay", width=70, anchor="center")
        self.step_table.pack(fill="both", expand=True)
        self.refresh_step_table()
        self.step_table.bind("<<TreeviewSelect>>", self._on_step_table_select)

        delay_editor = tk.Frame(wrapper, bg=self.palette["bg"])
        delay_editor.pack(fill="x", pady=(0, 8))
        tk.Label(
            delay_editor,
            text="Delay do passo selecionado (s)",
            bg=self.palette["bg"],
            fg=self.palette["text"],
            font=("Segoe UI", 10),
        ).pack(side="left")
        delay_entry = tk.Entry(delay_editor, textvariable=self.step_delay_var, width=10, font=("Segoe UI", 10))
        delay_entry.pack(side="left", padx=(8, 8))
        delay_entry.bind("<Return>", lambda _evt: self._apply_selected_step_delay())
        self._make_action_button(
            delay_editor,
            text="Atualizar Delay",
            command=self._apply_selected_step_delay,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
            padx=10,
        ).pack(side="left")

        actions = tk.Frame(wrapper, bg=self.palette["bg"])
        actions.pack(fill="x")
        self._make_action_button(
            actions,
            text="Treinar Imagem (3s)",
            command=self.capture_selected_step,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left")
        self._make_action_button(
            actions,
            text="Limpar Passo",
            command=self.clear_selected_step,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left", padx=(8, 0))
        self._make_action_button(
            actions,
            text="Ativar/Desativar",
            command=self.toggle_selected_step_enabled,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left", padx=(8, 0))
        self._make_action_button(
            actions,
            text="Salvar Fluxo",
            command=lambda: self.save_config(silent=False),
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="left", padx=(8, 0))
        self._make_action_button(
            actions,
            text="Fechar",
            command=self._close_flow_dialog,
            color=self.palette["muted"],
            hover=self.palette["muted_hover"],
        ).pack(side="right")

        window.protocol("WM_DELETE_WINDOW", self._close_flow_dialog)

    def _close_flow_dialog(self) -> None:
        self.save_config(silent=True)
        if self.flow_window is not None and self.flow_window.winfo_exists():
            self.flow_window.destroy()
        self.flow_window = None
        self.step_table = None

    def refresh_monitor_labels(self) -> None:
        lines = []
        for idx, monitor in enumerate(self.monitors, start=1):
            primary_tag = " (Principal)" if monitor["is_primary"] else ""
            lines.append(
                f"Monitor {idx}{primary_tag}: x={monitor['x']}, y={monitor['y']}, "
                f"w={monitor['width']}, h={monitor['height']}"
            )
        if not lines:
            lines.append("Nenhum monitor detectado.")
        self.monitor_summary_var.set("\n".join(lines))

    def _resolve_template_path(self, step: dict[str, Any]) -> Path | None:
        template_path = step.get("template_path")
        if not template_path:
            return None

        path = Path(str(template_path))
        if not path.is_absolute():
            path = self.project_root / path
        return path

    def _target_monitor(self) -> dict[str, Any] | None:
        if not self.monitors:
            return None
        try:
            monitor_index = int(self.target_monitor_var.get()) - 1
        except ValueError:
            return None
        monitor_index = max(0, min(monitor_index, len(self.monitors) - 1))
        return self.monitors[monitor_index]

    def _action_label(self, action: str) -> str:
        if action == "click_f2":
            return "Click + F2"
        if action == "fill_identifier":
            return "Preencher ID"
        return "Click"

    def refresh_step_table(self) -> None:
        if self.step_table is None:
            return

        self.step_table.delete(*self.step_table.get_children())
        for idx, step in enumerate(self.config_data.get("steps", []), start=1):
            template = self._resolve_template_path(step)
            template_status = "OK" if template and template.exists() else "-"
            fallback_xy = "OK" if step.get("x") is not None and step.get("y") is not None else "-"
            enabled_status = "Sim" if bool(step.get("enabled", True)) else "Nao"
            action = self._action_label(str(step.get("action", "click")))
            self.step_table.insert(
                "",
                "end",
                values=(
                    idx,
                    step.get("label"),
                    enabled_status,
                    action,
                    template_status,
                    fallback_xy,
                    step.get("delay_after", self.config_data.get("default_step_delay", 1.0)),
                ),
            )

    def _selected_step_index(self) -> int | None:
        if self.step_table is None:
            return None
        selected = self.step_table.selection()
        if not selected:
            return None
        return self.step_table.index(selected[0])

    def _on_step_table_select(self, _event=None) -> None:
        step_idx = self._selected_step_index()
        if step_idx is None:
            self.step_delay_var.set("")
            return

        step = self.config_data["steps"][step_idx]
        delay = float(step.get("delay_after", self.config_data.get("default_step_delay", 1.0)))
        self.step_delay_var.set(str(delay))

    def _reselect_step(self, step_idx: int) -> None:
        if self.step_table is None:
            return
        items = self.step_table.get_children()
        if 0 <= step_idx < len(items):
            item = items[step_idx]
            self.step_table.selection_set(item)
            self.step_table.focus(item)

    def _apply_selected_step_delay(self) -> None:
        step_idx = self._selected_step_index()
        if step_idx is None:
            messagebox.showwarning("Delay", "Selecione um passo antes de alterar o delay.")
            return

        raw = self.step_delay_var.get().strip().replace(",", ".")
        try:
            delay = float(raw)
            if delay < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Delay", "Informe um numero valido (ex: 1.0).")
            return

        self.config_data["steps"][step_idx]["delay_after"] = delay
        self.refresh_step_table()
        self._reselect_step(step_idx)
        self.save_config(silent=True)
        self._set_status(
            f"Delay do passo '{self.config_data['steps'][step_idx]['label']}' atualizado para {delay:.2f}s."
        )

    def toggle_selected_step_enabled(self) -> None:
        step_idx = self._selected_step_index()
        if step_idx is None:
            messagebox.showwarning("Passo", "Selecione um passo antes de alterar.")
            return

        step = self.config_data["steps"][step_idx]
        current = bool(step.get("enabled", True))
        step["enabled"] = not current
        status = "ativado" if step["enabled"] else "desativado"

        self.refresh_step_table()
        self._reselect_step(step_idx)
        self.save_config(silent=True)
        self._set_status(f"Passo '{step['label']}' {status}.")
    def capture_selected_step(self) -> None:
        if not is_vision_available():
            messagebox.showerror("Modo imagem", missing_vision_message())
            return

        step_idx = self._selected_step_index()
        if step_idx is None:
            messagebox.showwarning("Treino", "Selecione um passo da lista antes de treinar.")
            return

        if self.is_running:
            messagebox.showwarning("Execucao", "Pare a execucao atual antes de treinar.")
            return

        step_label = self.config_data["steps"][step_idx]["label"]
        self._set_status(f"Treino do passo '{step_label}'. Posicione o mouse no alvo.")
        thread = threading.Thread(target=self._capture_worker, args=(step_idx,), daemon=True)
        thread.start()

    def _capture_worker(self, step_idx: int) -> None:
        for seconds in (3, 2, 1):
            self._set_status_threadsafe(f"Capturando template em {seconds}s...")
            time.sleep(1)

        x, y = get_mouse_position()
        self.root.after(0, lambda: self._apply_captured_template(step_idx, x, y))

    def _apply_captured_template(self, step_idx: int, x: int, y: int) -> None:
        step = self.config_data["steps"][step_idx]
        step_id = str(step["id"])
        target_monitor = self._target_monitor()

        bounds = None
        if target_monitor:
            bounds = {
                "x": int(target_monitor["x"]),
                "y": int(target_monitor["y"]),
                "width": int(target_monitor["width"]),
                "height": int(target_monitor["height"]),
            }

        output_path = self.templates_dir / f"{step_id}.png"
        capture_size = int(step.get("capture_size", 120))

        try:
            save_template_around_point(
                x=x,
                y=y,
                output_path=output_path,
                capture_size=capture_size,
                bounds=bounds,
            )
        except Exception as exc:
            messagebox.showerror("Treino", f"Falha ao capturar template: {exc}")
            self._set_status("Falha no treino.")
            return

        step["template_path"] = output_path.relative_to(self.project_root).as_posix()
        step["x"] = x
        step["y"] = y

        self.refresh_step_table()
        self.save_config(silent=True)
        self._set_status(f"Template salvo para '{step['label']}'.")

    def clear_selected_step(self) -> None:
        step_idx = self._selected_step_index()
        if step_idx is None:
            messagebox.showwarning("Treino", "Selecione um passo para limpar.")
            return

        step = self.config_data["steps"][step_idx]
        template = self._resolve_template_path(step)
        if template and template.exists():
            try:
                template.unlink()
            except OSError:
                pass

        step["template_path"] = None
        step["x"] = None
        step["y"] = None
        self.refresh_step_table()
        self.save_config(silent=True)
        self._set_status(f"Passo '{step['label']}' limpo.")

    def save_config(self, silent: bool = False) -> bool:
        try:
            self.config_data["app_monitor_index"] = int(self.app_monitor_var.get())
            self.config_data["target_monitor_index"] = int(self.target_monitor_var.get())
            self.config_data["pre_start_delay"] = float(self.pre_start_var.get())
            self.config_data["default_step_delay"] = float(self.default_delay_var.get())
        except ValueError:
            messagebox.showerror("Configuracao", "Tempos e monitores precisam ser validos.")
            return False

        flow_name = self.flow_name_var.get().strip() or "Criacao de Requisicao"
        flow_path = self.flow_path_var.get().strip() or "Fluxo nao informado"
        self.config_data["flow_name"] = flow_name
        self.config_data["flow_path_text"] = flow_path

        self.flow_active_var.set(f"Fluxo ativo: {flow_name}")
        self.store.save(self.config_data)

        if not silent:
            self._set_status("Configuracao salva.")
        return True

    def apply_app_monitor_position(self, initial: bool) -> None:
        if not self.monitors:
            return

        try:
            monitor_index = int(self.app_monitor_var.get()) - 1
        except ValueError:
            monitor_index = 0

        monitor_index = max(0, min(monitor_index, len(self.monitors) - 1))
        monitor = self.monitors[monitor_index]

        self.root.update_idletasks()
        width = max(460, self.root.winfo_width())
        height = max(260, self.root.winfo_height())
        monitor_x = int(monitor["x"])
        monitor_y = int(monitor["y"])
        monitor_w = int(monitor["width"])
        monitor_h = int(monitor["height"])

        x = monitor_x + max(0, (monitor_w - width) // 2)
        y = monitor_y + max(0, (monitor_h - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

        if not initial:
            self._set_status(f"App centralizado no monitor {monitor_index + 1}.")

    def _validate_steps(self) -> list[str]:
        missing = []
        for step in self.config_data["steps"]:
            if not bool(step.get("enabled", True)):
                continue

            action = str(step.get("action", "click"))
            if action == "fill_identifier" and self.current_inputs is not None:
                identifier_type = str(step.get("identifier_type", ""))
                identifier_value = str(self.current_inputs.get(identifier_type, "")).strip()
                if not identifier_value:
                    continue

            template = self._resolve_template_path(step)
            has_template = template is not None and template.exists()
            has_xy = step.get("x") is not None and step.get("y") is not None
            if not has_template and not has_xy:
                missing.append(step["label"])
        return missing

    def _digits_only(self, value: str) -> str:
        return "".join(ch for ch in value if ch.isdigit())

    def _format_cpf(self, value: str) -> str:
        digits = self._digits_only(value)[:11]
        if len(digits) <= 3:
            return digits
        if len(digits) <= 6:
            return f"{digits[:3]}.{digits[3:]}"
        if len(digits) <= 9:
            return f"{digits[:3]}.{digits[3:6]}.{digits[6:]}"
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"

    def _format_cnpj(self, value: str) -> str:
        digits = self._digits_only(value)[:14]
        if len(digits) <= 2:
            return digits
        if len(digits) <= 5:
            return f"{digits[:2]}.{digits[2:]}"
        if len(digits) <= 8:
            return f"{digits[:2]}.{digits[2:5]}.{digits[5:]}"
        if len(digits) <= 12:
            return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:]}"
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"

    def _apply_mask(self, var: tk.StringVar, formatter) -> None:
        current = var.get()
        formatted = formatter(current)
        if current != formatted:
            var.set(formatted)

    def _locate_step_point(self, step: dict[str, Any], monitor: dict[str, Any]) -> tuple[int, int] | None:
        template_path = self._resolve_template_path(step)
        threshold = float(step.get("match_threshold", 0.84))
        step_name = str(step.get("label", "Passo"))

        if template_path and template_path.exists():
            if is_vision_available():
                found = locate_template_center(
                    template_path=template_path,
                    monitor=monitor,
                    threshold=threshold,
                )
                if found:
                    x, y, score = found
                    self._set_status_threadsafe(
                        f"{step_name}: template encontrado (score {score:.2f})."
                    )
                    return x, y
                self._set_status_threadsafe(
                    f"{step_name}: template nao encontrado, tentando fallback XY."
                )
            else:
                self._set_status_threadsafe(
                    f"{step_name}: visao por imagem indisponivel, tentando fallback XY."
                )

        if step.get("x") is not None and step.get("y") is not None:
            return int(step["x"]), int(step["y"])

        return None

    def _execute_step_action(self, step: dict[str, Any], monitor: dict[str, Any]) -> str:
        step_name = str(step.get("label", "Passo"))
        action = str(step.get("action", "click"))
        clicks = int(step.get("clicks", 1))

        identifier_value = ""
        identifier_type = str(step.get("identifier_type", ""))
        if action == "fill_identifier":
            if self.current_inputs:
                identifier_value = str(self.current_inputs.get(identifier_type, "")).strip()
            if not identifier_value:
                if identifier_type:
                    self._set_status_threadsafe(
                        f"{step_name}: {identifier_type.upper()} vazio. Passo ignorado."
                    )
                else:
                    self._set_status_threadsafe(f"{step_name}: identificador vazio. Passo ignorado.")
                return "skipped"

        point = self._locate_step_point(step, monitor)
        if point is None:
            return "failed"

        x, y = point
        set_mouse_position(x, y)
        time.sleep(0.08)

        if action == "click_f2":
            left_click(clicks=max(1, clicks))
            time.sleep(0.12)
            press_f2()
            return "done"

        if action == "fill_identifier":
            left_click(clicks=max(2, clicks), interval=0.05)
            time.sleep(0.08)
            press_backspace(times=1)
            time.sleep(0.05)
            type_text(identifier_value)
            return "done"

        left_click(clicks=max(1, clicks))
        return "done"
    def _collect_required_inputs(self) -> dict[str, str] | None:
        defaults = self.config_data.get("last_inputs", {})

        dialog = tk.Toplevel(self.root)
        dialog.title("Dados obrigatorios")
        dialog.geometry("560x360")
        dialog.minsize(540, 340)
        dialog.configure(bg=self.palette["bg"])
        dialog.transient(self.root)
        dialog.grab_set()

        obs_var = tk.StringVar(value=str(defaults.get("observacao", "")))
        cnpj_var = tk.StringVar(value=self._format_cnpj(str(defaults.get("cnpj", ""))))
        cpf_var = tk.StringVar(value=self._format_cpf(str(defaults.get("cpf", ""))))
        codigo_fornecedor_var = tk.StringVar(value=str(defaults.get("codigo_fornecedor", "")))
        linhas_var = tk.StringVar(value=str(defaults.get("linhas", "")))

        result: dict[str, str] | None = None

        frame = tk.Frame(dialog, bg=self.palette["bg"])
        frame.pack(fill="both", expand=True, padx=14, pady=12)

        tk.Label(
            frame,
            text="Antes de executar, informe os dados (CPF ou CNPJ e obrigatorio):",
            bg=self.palette["bg"],
            fg=self.palette["text"],
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        form = tk.Frame(frame, bg=self.palette["bg"])
        form.pack(fill="x")

        tk.Label(form, text="Observacao", bg=self.palette["bg"], font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=obs_var, width=52, font=("Segoe UI", 10)).grid(
            row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 8)
        )

        tk.Label(form, text="CNPJ procurado", bg=self.palette["bg"], font=("Segoe UI", 10)).grid(
            row=1, column=0, sticky="w"
        )
        cnpj_entry = tk.Entry(form, textvariable=cnpj_var, width=30, font=("Segoe UI", 10))
        cnpj_entry.grid(
            row=1, column=1, sticky="w", padx=(8, 0), pady=(0, 8)
        )
        cnpj_entry.bind("<KeyRelease>", lambda _evt: self._apply_mask(cnpj_var, self._format_cnpj))

        tk.Label(form, text="CPF procurado", bg=self.palette["bg"], font=("Segoe UI", 10)).grid(
            row=2, column=0, sticky="w"
        )
        cpf_entry = tk.Entry(form, textvariable=cpf_var, width=30, font=("Segoe UI", 10))
        cpf_entry.grid(
            row=2, column=1, sticky="w", padx=(8, 0), pady=(0, 8)
        )
        cpf_entry.bind("<KeyRelease>", lambda _evt: self._apply_mask(cpf_var, self._format_cpf))

        tk.Label(form, text="Codigo do fornecedor", bg=self.palette["bg"], font=("Segoe UI", 10)).grid(
            row=3, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=codigo_fornecedor_var, width=30, font=("Segoe UI", 10)).grid(
            row=3, column=1, sticky="w", padx=(8, 0), pady=(0, 8)
        )

        tk.Label(form, text="Linhas do processo", bg=self.palette["bg"], font=("Segoe UI", 10)).grid(
            row=4, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=linhas_var, width=30, font=("Segoe UI", 10)).grid(
            row=4, column=1, sticky="w", padx=(8, 0)
        )

        def confirm() -> None:
            nonlocal result
            payload = {
                "observacao": obs_var.get().strip(),
                "cnpj": cnpj_var.get().strip(),
                "cpf": cpf_var.get().strip(),
                "codigo_fornecedor": codigo_fornecedor_var.get().strip(),
                "linhas": linhas_var.get().strip(),
            }
            if not payload["cnpj"] and not payload["cpf"]:
                messagebox.showwarning("Campos obrigatorios", "Informe CPF ou CNPJ.")
                return

            result = payload
            self.config_data["last_inputs"] = payload
            self.save_config(silent=True)
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        bottom = tk.Frame(frame, bg=self.palette["bg"])
        bottom.pack(fill="x", pady=(14, 0))
        self._make_action_button(
            bottom,
            text="Cancelar",
            command=cancel,
            color=self.palette["muted"],
            hover=self.palette["muted_hover"],
        ).pack(side="right")
        self._make_action_button(
            bottom,
            text="Confirmar",
            command=confirm,
            color=self.palette["primary"],
            hover=self.palette["primary_hover"],
        ).pack(side="right", padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.root.wait_window(dialog)
        return result

    def run_requisicao_flow(self) -> None:
        if self.is_running:
            messagebox.showinfo("Execucao", "O fluxo ja esta em execucao.")
            return

        if not self.save_config(silent=True):
            return

        run_inputs = self._collect_required_inputs()
        if run_inputs is None:
            self._set_status("Execucao cancelada pelo usuario.")
            return
        self.current_inputs = run_inputs

        missing_steps = self._validate_steps()
        if missing_steps:
            joined = ", ".join(missing_steps)
            messagebox.showwarning(
                "Treino incompleto",
                f"Treine os passos antes de executar: {joined}",
            )
            return

        self.stop_event.clear()
        self.is_running = True
        self.run_button.configure(state="disabled")
        worker = threading.Thread(target=self._run_flow_worker, daemon=True)
        worker.start()

    def _run_flow_worker(self) -> None:
        try:
            pre_start = max(0.0, float(self.config_data.get("pre_start_delay", 3.0)))
            default_delay = max(0.0, float(self.config_data.get("default_step_delay", 0.8)))
            monitor = self._target_monitor()
            target_monitor_index = self.config_data.get("target_monitor_index", 1)

            if monitor is None:
                raise RuntimeError("Monitor alvo nao encontrado.")

            if self.current_inputs:
                cnpj = self.current_inputs.get("cnpj", "")
                cpf = self.current_inputs.get("cpf", "")
                self._set_status_threadsafe(
                    f"Dados recebidos. CNPJ: {cnpj or '-'} | CPF: {cpf or '-'}"
                )

            countdown = int(pre_start)
            while countdown > 0:
                if self.stop_event.is_set():
                    self._set_status_threadsafe("Execucao interrompida.")
                    return
                self._set_status_threadsafe(
                    f"Iniciando no monitor {target_monitor_index} em {countdown}s..."
                )
                time.sleep(1)
                countdown -= 1

            for step in self.config_data["steps"]:
                if self.stop_event.is_set():
                    self._set_status_threadsafe("Execucao interrompida.")
                    return

                if not bool(step.get("enabled", True)):
                    continue

                step_name = step["label"]
                delay_after = max(0.0, float(step.get("delay_after", default_delay)))
                result = self._execute_step_action(step, monitor)
                if result == "failed":
                    raise RuntimeError(
                        f"Nao foi possivel executar '{step_name}'. Treine novamente o passo."
                    )

                elapsed = 0.0
                while elapsed < delay_after:
                    if self.stop_event.is_set():
                        self._set_status_threadsafe("Execucao interrompida.")
                        return
                    chunk = min(0.1, delay_after - elapsed)
                    time.sleep(chunk)
                    elapsed += chunk

            self._set_status_threadsafe("Fluxo concluido com sucesso.")
        except Exception as exc:
            self._set_status_threadsafe(f"Erro durante a execucao: {exc}")
        finally:
            self.root.after(0, self._finish_run_state)

    def _finish_run_state(self) -> None:
        self.is_running = False
        self.run_button.configure(state="normal")

    def stop_execution(self) -> None:
        self.stop_event.set()
        self._set_status("Sinal de parada enviado.")

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)

    def _set_status_threadsafe(self, message: str) -> None:
        self.root.after(0, lambda: self.status_var.set(message))

    def on_close(self) -> None:
        if self.is_running:
            if not messagebox.askyesno(
                "Fechar app",
                "A execucao ainda esta ativa. Deseja parar e fechar?",
            ):
                return
            self.stop_event.set()

        self.save_config(silent=True)
        self.root.destroy()


def run_gui() -> None:
    set_dpi_awareness()
    root = tk.Tk()
    project_root = Path(__file__).resolve().parents[2]
    config_path = project_root / "config" / "requisicao_config.json"
    app = RequisicaoApp(root, ConfigStore(config_path), project_root)
    root.mainloop()

