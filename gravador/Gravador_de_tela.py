from __future__ import annotations

import ctypes
import subprocess
import tempfile
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, simpledialog
import tkinter as tk
from tkinter import ttk

try:
    import cv2
    import imageio_ffmpeg
    import mss
    import numpy as np
    from pynput import keyboard as pynput_keyboard
    import soundcard as sc
    import soundfile as sf
except ImportError as exc:
    raise SystemExit(
        "Dependencia ausente. Execute:\n"
        "pip install -r gravador/requirements.txt\n\n"
        f"Detalhe: {exc}"
    ) from exc

try:
    import sounddevice as sd
except ImportError:
    sd = None


CUSTOM_MONITOR_INDEX = -1
COUNTDOWN_SECONDS = 3
AUDIO_SAMPLE_RATE = 48_000
AUDIO_SAMPLE_RATE_CANDIDATES = (48_000, 44_100, 32_000, 24_000, 22_050, 16_000, 12_000, 11_025, 8_000)
AUDIO_BLOCK_SIZE = 2_048
VIDEO_PRESET = "veryfast"
COLOR_BG = "#EAF3FF"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT = "#0D3B66"
COLOR_BLUE = "#1565C0"
COLOR_BLUE_DARK = "#0E4C99"
COLOR_ORANGE = "#F57C00"
COLOR_ORANGE_DARK = "#D96C00"
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000

AUDIO_MODES = (
    "Sem audio",
    "Apenas microfone",
    "Apenas som do sistema",
    "Microfone + som do sistema",
)

QUALITY_PROFILES = {
    "Alta (60 FPS)": {"fps": 60, "crf": 18},
    "Boa (30 FPS)": {"fps": 30, "crf": 21},
    "Leve (20 FPS)": {"fps": 20, "crf": 24},
}


@dataclass
class MicTuning:
    gain_percent: float
    sensitivity_percent: float
    noise_suppression: bool


@dataclass
class RecordingOptions:
    output_file: Path
    region: dict[str, int]
    fps: int
    crf: int
    microphone: object | None
    system_audio: object | None
    mic_tuning: MicTuning


class ScreenRecorderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Gravador de Tela Inteligente - Base")
        self.geometry("760x500")
        self.minsize(740, 480)

        self.monitor_map: dict[str, int] = {}
        self.microphone_map: dict[str, object] = {}
        self.desktop_map: dict[str, object] = {}
        self.custom_region: dict[str, int] | None = None
        self.active_region: dict[str, int] | None = None

        default_dir = Path("C:/Projetos/Gravacoes")
        default_name = f"gravacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        self.output_dir_var = tk.StringVar(value=str(default_dir))
        self.file_name_var = tk.StringVar(value=default_name)
        self.monitor_var = tk.StringVar()
        self.quality_var = tk.StringVar(value="Alta (60 FPS)")
        self.audio_mode_var = tk.StringVar(value="Microfone + som do sistema")
        self.mic_var = tk.StringVar()
        self.desktop_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Pronto para iniciar.")

        self.hotkey_start_var = tk.StringVar(value="r")
        self.hotkey_pause_var = tk.StringVar(value="p")
        self.hotkey_stop_var = tk.StringVar(value="e")
        self.hotkey_mute_var = tk.StringVar(value="m")
        self.mic_gain_var = tk.DoubleVar(value=100.0)
        self.mic_sensitivity_var = tk.DoubleVar(value=65.0)
        self.mic_noise_suppress_var = tk.BooleanVar(value=True)
        self.mic_gain_text_var = tk.StringVar(value="100%")
        self.mic_sensitivity_text_var = tk.StringVar(value="65%")

        self.worker_thread: threading.Thread | None = None
        self.hotkey_listener: object | None = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.mute_event = threading.Event()
        self.recording = False
        self.paused = False
        self.muted = False
        self.countdown_active = False
        self.countdown_left = 0
        self.countdown_job: str | None = None
        self.countdown_sequence: list[str] = []
        self.countdown_index = 0
        self.countdown_overlay: tk.Toplevel | None = None
        self.countdown_label: tk.Label | None = None
        self.pending_options: RecordingOptions | None = None
        self.audio_errors: list[str] = []
        self.last_recording_warning: str | None = None
        self.last_output_file: Path | None = None
        self.audio_error_lock = threading.Lock()
        self._latest_refresh_token = 0
        self._choosing_directory = False
        self.overlay_windows: list[tk.Toplevel] = []
        self.mic_tuning_controls: list[ttk.Widget] = []

        self._configure_theme()
        self._build_ui()
        self._refresh_mic_tuning_labels()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._refresh_sources_async)
        self.after(150, lambda: self._apply_hotkeys(show_feedback=False))
        self._apply_audio_mode()

    def _configure_theme(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.configure(bg=COLOR_BG)
        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("App.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=COLOR_BG, foreground=COLOR_BLUE, font=("Segoe UI", 10, "bold"))
        style.configure("Card.TLabelframe", background=COLOR_CARD, bordercolor="#BFD9FF", relief="solid")
        style.configure(
            "Card.TLabelframe.Label",
            background=COLOR_CARD,
            foreground=COLOR_BLUE,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure("Card.TLabel", background=COLOR_CARD, foreground=COLOR_TEXT, font=("Segoe UI", 9))
        style.configure("App.TEntry", fieldbackground="#FFFFFF", foreground=COLOR_TEXT, padding=2)
        style.configure("App.TCombobox", fieldbackground="#FFFFFF", foreground=COLOR_TEXT, padding=1)
        style.map(
            "App.TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            selectbackground=[("readonly", "#FFFFFF")],
            selectforeground=[("readonly", COLOR_TEXT)],
        )
        style.configure(
            "Primary.TButton",
            background=COLOR_ORANGE,
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            padding=(8, 3),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLOR_ORANGE_DARK), ("disabled", "#DDDDDD")],
            foreground=[("disabled", "#7A7A7A")],
        )
        style.configure(
            "Secondary.TButton",
            background=COLOR_BLUE,
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            padding=(8, 3),
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLOR_BLUE_DARK), ("disabled", "#DDDDDD")],
            foreground=[("disabled", "#7A7A7A")],
        )
        style.configure(
            "Ghost.TButton",
            background="#DDEBFF",
            foreground=COLOR_BLUE_DARK,
            font=("Segoe UI", 9),
            padding=(8, 3),
            borderwidth=0,
        )
        style.map(
            "Ghost.TButton",
            background=[("active", "#CFE2FF"), ("disabled", "#ECECEC")],
            foreground=[("disabled", "#9B9B9B")],
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=8, style="App.TFrame")
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="Gravador Inteligente", style="Title.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )

        row = 1
        ttk.Label(root, text="Pasta de destino", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(root, textvariable=self.output_dir_var, style="App.TEntry").grid(
            row=row, column=1, sticky="ew", padx=6, pady=2
        )
        self.choose_dir_btn = ttk.Button(
            root, text="Escolher", command=self._choose_output_dir, style="Ghost.TButton"
        )
        self.choose_dir_btn.grid(row=row, column=2, pady=2)

        row += 1
        ttk.Label(root, text="Nome do arquivo", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(root, textvariable=self.file_name_var, style="App.TEntry").grid(
            row=row, column=1, sticky="ew", padx=6, pady=2
        )

        row += 1
        ttk.Label(root, text="Monitor", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        self.monitor_combo = ttk.Combobox(root, textvariable=self.monitor_var, state="readonly", style="App.TCombobox")
        self.monitor_combo.grid(row=row, column=1, sticky="ew", padx=6, pady=2)

        row += 1
        ttk.Label(root, text="Qualidade", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        ttk.Combobox(
            root,
            textvariable=self.quality_var,
            values=list(QUALITY_PROFILES.keys()),
            state="readonly",
            style="App.TCombobox",
        ).grid(row=row, column=1, sticky="ew", padx=6, pady=2)

        row += 1
        ttk.Label(root, text="Audio", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        audio_combo = ttk.Combobox(
            root,
            textvariable=self.audio_mode_var,
            values=AUDIO_MODES,
            state="readonly",
            style="App.TCombobox",
        )
        audio_combo.grid(row=row, column=1, sticky="ew", padx=6, pady=2)
        audio_combo.bind("<<ComboboxSelected>>", lambda _: self._apply_audio_mode())

        row += 1
        ttk.Label(root, text="Microfone", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        self.mic_combo = ttk.Combobox(root, textvariable=self.mic_var, state="readonly", style="App.TCombobox")
        self.mic_combo.grid(row=row, column=1, sticky="ew", padx=6, pady=2)

        row += 1
        ttk.Label(root, text="Som do sistema", style="App.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        self.desktop_combo = ttk.Combobox(
            root, textvariable=self.desktop_var, state="readonly", style="App.TCombobox"
        )
        self.desktop_combo.grid(row=row, column=1, sticky="ew", padx=6, pady=2)

        row += 1
        mic_frame = ttk.LabelFrame(root, text="Microfone (avancado)", padding=6, style="Card.TLabelframe")
        mic_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(2, 4))
        mic_frame.columnconfigure(1, weight=1)

        ttk.Label(mic_frame, text="Volume", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        self.mic_gain_scale = ttk.Scale(
            mic_frame,
            from_=50,
            to=250,
            orient="horizontal",
            variable=self.mic_gain_var,
            command=lambda _value: self._refresh_mic_tuning_labels(),
        )
        self.mic_gain_scale.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(mic_frame, textvariable=self.mic_gain_text_var, style="Card.TLabel", width=6).grid(
            row=0, column=2, sticky="e"
        )

        ttk.Label(mic_frame, text="Sensibilidade", style="Card.TLabel").grid(row=1, column=0, sticky="w")
        self.mic_sensitivity_scale = ttk.Scale(
            mic_frame,
            from_=0,
            to=100,
            orient="horizontal",
            variable=self.mic_sensitivity_var,
            command=lambda _value: self._refresh_mic_tuning_labels(),
        )
        self.mic_sensitivity_scale.grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Label(mic_frame, textvariable=self.mic_sensitivity_text_var, style="Card.TLabel", width=6).grid(
            row=1, column=2, sticky="e"
        )

        self.mic_noise_check = ttk.Checkbutton(
            mic_frame,
            text="Supressor de ruido (reduz chiado e som ambiente)",
            variable=self.mic_noise_suppress_var,
        )
        self.mic_noise_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.mic_tuning_controls = [self.mic_gain_scale, self.mic_sensitivity_scale, self.mic_noise_check]

        row += 1
        ttk.Button(
            root,
            text="Atualizar dispositivos",
            command=self._refresh_sources_async,
            style="Ghost.TButton",
        ).grid(row=row, column=0, pady=5, sticky="w")

        actions = ttk.Frame(root, style="App.TFrame")
        actions.grid(row=row, column=1, columnspan=2, sticky="e", pady=5)
        self.print_btn = ttk.Button(actions, text="Print", command=self._take_screenshot, style="Secondary.TButton")
        self.print_btn.pack(side="left", padx=4)

        self.open_last_btn = ttk.Button(
            actions,
            text="Abrir pasta",
            command=self._open_last_video_folder,
            state="disabled",
            style="Ghost.TButton",
        )
        self.open_last_btn.pack(side="left", padx=4)

        self.start_btn = ttk.Button(actions, text="Gravar", command=self._start_recording, style="Primary.TButton")
        self.start_btn.pack(side="left", padx=4)

        self.pause_btn = ttk.Button(
            actions, text="Pausar", command=self._toggle_pause, state="disabled", style="Secondary.TButton"
        )
        self.pause_btn.pack(side="left", padx=4)

        self.mute_btn = ttk.Button(
            actions, text="Mudo", command=self._toggle_mute, state="disabled", style="Ghost.TButton"
        )
        self.mute_btn.pack(side="left", padx=4)

        self.stop_btn = ttk.Button(
            actions,
            text="Parar",
            command=self._stop_recording,
            state="disabled",
            style="Secondary.TButton",
            width=8,
        )
        self.stop_btn.pack(side="left", padx=4)

        row += 1
        status_frame = ttk.LabelFrame(root, text="Status", padding=6, style="Card.TLabelframe")
        status_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status_var, style="Card.TLabel").grid(row=0, column=0, sticky="w")

        row += 1
        hotkey_frame = ttk.LabelFrame(root, text="Atalhos (Ctrl+Shift+tecla)", padding=6, style="Card.TLabelframe")
        hotkey_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        for col in (1, 3, 5, 7):
            hotkey_frame.columnconfigure(col, weight=1)

        ttk.Label(hotkey_frame, text="Iniciar", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(hotkey_frame, textvariable=self.hotkey_start_var, width=5, style="App.TEntry").grid(
            row=0, column=1, sticky="w", padx=6
        )
        ttk.Label(hotkey_frame, text="Pausar", style="Card.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(hotkey_frame, textvariable=self.hotkey_pause_var, width=5, style="App.TEntry").grid(
            row=0, column=3, sticky="w", padx=6
        )
        ttk.Label(hotkey_frame, text="Encerrar", style="Card.TLabel").grid(row=0, column=4, sticky="w")
        ttk.Entry(hotkey_frame, textvariable=self.hotkey_stop_var, width=5, style="App.TEntry").grid(
            row=0, column=5, sticky="w", padx=6
        )
        ttk.Label(hotkey_frame, text="Mudo", style="Card.TLabel").grid(row=0, column=6, sticky="w")
        ttk.Entry(hotkey_frame, textvariable=self.hotkey_mute_var, width=5, style="App.TEntry").grid(
            row=0, column=7, sticky="w", padx=6
        )
        ttk.Button(
            hotkey_frame, text="Aplicar", command=self._apply_hotkeys, style="Secondary.TButton"
        ).grid(row=0, column=8, padx=(8, 0))

    def _choose_output_dir(self) -> None:
        if self._choosing_directory:
            return

        self._choosing_directory = True
        self.choose_dir_btn.configure(state="disabled")
        self.status_var.set("Abrindo seletor de pasta...")
        initial_dir = self.output_dir_var.get().strip() or "C:\\"

        threading.Thread(
            target=self._choose_output_dir_worker,
            args=(initial_dir,),
            daemon=True,
        ).start()

    def _choose_output_dir_worker(self, initial_dir: str) -> None:
        chosen_dir = ""
        error_msg = ""
        try:
            chosen_dir = self._pick_directory_via_powershell(initial_dir)
        except Exception as exc:
            error_msg = str(exc)

        self.after(
            0,
            lambda: self._finish_choose_output_dir(chosen_dir, error_msg, initial_dir),
        )

    def _finish_choose_output_dir(self, chosen_dir: str, error_msg: str, initial_dir: str) -> None:
        self._choosing_directory = False
        self.choose_dir_btn.configure(state="normal")

        if chosen_dir:
            self.output_dir_var.set(chosen_dir)
            self.status_var.set(f"Pasta selecionada: {chosen_dir}")
            return

        if error_msg:
            manual = simpledialog.askstring(
                "Pasta de destino",
                "Falha ao abrir o seletor. Informe o caminho manualmente:",
                initialvalue=initial_dir,
                parent=self,
            )
            if manual and manual.strip():
                self.output_dir_var.set(manual.strip())
                self.status_var.set(f"Pasta selecionada: {manual.strip()}")
            else:
                self.status_var.set("Selecao de pasta cancelada.")
            return

        self.status_var.set("Selecao de pasta cancelada.")

    @staticmethod
    def _pick_directory_via_powershell(initial_dir: str) -> str:
        initial_escaped = initial_dir.replace("'", "''")
        script = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$dlg = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$dlg.Description = 'Selecione a pasta de destino'; "
            "$dlg.ShowNewFolderButton = $true; "
            f"$start = '{initial_escaped}'; "
            "if ([System.IO.Directory]::Exists($start)) { $dlg.SelectedPath = $start }; "
            "if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { "
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
            "Write-Output $dlg.SelectedPath }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-STA", "-Command", script],
            capture_output=True,
            check=False,
            timeout=300,
        )
        if result.returncode != 0:
            stderr = ScreenRecorderApp._decode_process_output(result.stderr).strip()
            raise RuntimeError(stderr or "erro ao abrir seletor de pasta")
        return ScreenRecorderApp._decode_process_output(result.stdout).strip()

    def _refresh_sources_async(self) -> None:
        token = time.time_ns()
        self._latest_refresh_token = token
        self.status_var.set("Atualizando monitores e dispositivos de audio...")
        threading.Thread(
            target=self._refresh_sources_worker,
            args=(token,),
            daemon=True,
        ).start()

    def _refresh_sources_worker(self, token: int) -> None:
        monitor_items: list[tuple[str, int]] = []
        microphone_items: list[tuple[str, object]] = []
        desktop_items: list[tuple[str, object]] = []
        errors: list[str] = []

        try:
            monitor_items = self._scan_monitors()
        except Exception as exc:
            errors.append(f"Monitores: {exc}")

        try:
            microphone_items, desktop_items = self._scan_audio_devices()
        except Exception as exc:
            errors.append(f"Audio: {exc}")

        self.after(
            0,
            lambda: self._apply_refresh_results(
                token, monitor_items, microphone_items, desktop_items, errors
            ),
        )

    @staticmethod
    def _scan_monitors() -> list[tuple[str, int]]:
        with mss.mss() as sct:
            monitors = sct.monitors
            if len(monitors) <= 1:
                raise RuntimeError("nenhum monitor detectado")

            monitor_names = ScreenRecorderApp._get_monitor_device_names()
            items: list[tuple[str, int]] = [("Todos os monitores", 0)]
            for idx, mon in enumerate(monitors[1:], start=1):
                monitor_name = monitor_names[idx - 1] if idx - 1 < len(monitor_names) else f"DISPLAY{idx}"
                label = (
                    f"Monitor {idx} ({monitor_name}): "
                    f"{mon['width']}x{mon['height']} @ ({mon['left']},{mon['top']})"
                )
                items.append((label, idx))
            items.append(("Area personalizada", CUSTOM_MONITOR_INDEX))
            return items

    @staticmethod
    def _scan_audio_devices() -> tuple[list[tuple[str, object]], list[tuple[str, object]]]:
        microphones = list(sc.all_microphones(include_loopback=False))
        speakers = list(sc.all_speakers())

        default_mic_id = ""
        try:
            default_mic = sc.default_microphone()
            default_mic_id = str(getattr(default_mic, "id", ""))
        except Exception:
            default_mic_id = ""

        if default_mic_id:
            microphones.sort(key=lambda d: 0 if str(getattr(d, "id", "")) == default_mic_id else 1)

        microphone_items: list[tuple[str, object]] = []
        for idx, device in enumerate(microphones, start=1):
            is_default = str(getattr(device, "id", "")) == default_mic_id and default_mic_id != ""
            suffix = " (padrao)" if is_default else ""
            microphone_items.append((f"{idx}. {device.name}{suffix}", device))

        default_speaker_id = ""
        try:
            default_speaker = sc.default_speaker()
            default_speaker_id = str(getattr(default_speaker, "id", ""))
        except Exception:
            default_speaker_id = ""

        if default_speaker_id:
            speakers.sort(key=lambda d: 0 if str(getattr(d, "id", "")) == default_speaker_id else 1)

        desktop_items: list[tuple[str, object]] = []
        for idx, speaker in enumerate(speakers, start=1):
            loopback = ScreenRecorderApp._get_loopback_for_speaker(speaker)
            if loopback is not None:
                is_default = str(getattr(speaker, "id", "")) == default_speaker_id and default_speaker_id != ""
                suffix = " (padrao do sistema)" if is_default else ""
                desktop_items.append((f"{idx}. {speaker.name} (som do sistema){suffix}", loopback))

        if not desktop_items:
            loopbacks = sc.all_microphones(include_loopback=True)
            for idx, device in enumerate(loopbacks, start=1):
                if "loopback" in device.name.lower():
                    desktop_items.append((f"{idx}. {device.name}", device))

        return microphone_items, desktop_items

    @staticmethod
    def _get_loopback_for_speaker(speaker: object) -> object | None:
        candidates: list[str] = []
        speaker_id = getattr(speaker, "id", None)
        speaker_name = getattr(speaker, "name", None)
        if speaker_id:
            candidates.append(speaker_id)
        if speaker_name:
            candidates.append(speaker_name)

        for candidate in candidates:
            try:
                return sc.get_microphone(candidate, include_loopback=True)
            except Exception:
                continue
        return None

    @staticmethod
    def _get_monitor_device_names() -> list[str]:
        class _RECT(ctypes.Structure):
            _fields_ = [
                ("left", wintypes.LONG),
                ("top", wintypes.LONG),
                ("right", wintypes.LONG),
                ("bottom", wintypes.LONG),
            ]

        class _MONITORINFOEXW(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", _RECT),
                ("rcWork", _RECT),
                ("dwFlags", wintypes.DWORD),
                ("szDevice", wintypes.WCHAR * 32),
            ]

        names: list[str] = []
        try:
            user32 = ctypes.windll.user32
            enum_proc_type = ctypes.WINFUNCTYPE(
                wintypes.BOOL,
                ctypes.c_void_p,
                ctypes.c_void_p,
                ctypes.POINTER(_RECT),
                wintypes.LPARAM,
            )

            def _enum_proc(h_monitor, _hdc, _rect_ptr, _lparam):
                info = _MONITORINFOEXW()
                info.cbSize = ctypes.sizeof(_MONITORINFOEXW)
                if user32.GetMonitorInfoW(h_monitor, ctypes.byref(info)):
                    names.append(info.szDevice)
                return True

            callback = enum_proc_type(_enum_proc)
            user32.EnumDisplayMonitors(0, 0, callback, 0)
        except Exception:
            return []
        return names

    def _apply_refresh_results(
        self,
        token: int,
        monitor_items: list[tuple[str, int]],
        microphone_items: list[tuple[str, object]],
        desktop_items: list[tuple[str, object]],
        errors: list[str],
    ) -> None:
        if token != self._latest_refresh_token:
            return

        if monitor_items:
            previous_monitor = self.monitor_var.get()
            self.monitor_map = {label: idx for label, idx in monitor_items}
            monitor_values = list(self.monitor_map.keys())
            self.monitor_combo["values"] = monitor_values
            if previous_monitor in self.monitor_map:
                self.monitor_var.set(previous_monitor)
            else:
                self.monitor_var.set(self._pick_default_monitor_label(monitor_values))

        previous_mic = self.mic_var.get()
        self.microphone_map = {label: device for label, device in microphone_items}
        mic_values = list(self.microphone_map.keys()) or ["(nenhum microfone encontrado)"]
        self.mic_combo["values"] = mic_values
        if previous_mic in self.microphone_map:
            self.mic_var.set(previous_mic)
        else:
            self.mic_var.set(self._pick_default_label(mic_values, "(padrao)"))

        previous_desktop = self.desktop_var.get()
        self.desktop_map = {label: device for label, device in desktop_items}
        desktop_values = list(self.desktop_map.keys()) or ["(nenhum som do sistema encontrado)"]
        self.desktop_combo["values"] = desktop_values
        if previous_desktop in self.desktop_map:
            self.desktop_var.set(previous_desktop)
        else:
            self.desktop_var.set(self._pick_default_label(desktop_values, "(padrao do sistema)"))

        self._apply_audio_mode()
        if errors:
            self.status_var.set("Atualizacao concluida com alertas.")
            messagebox.showwarning("Atualizacao parcial", "\n".join(errors))
        else:
            self.status_var.set("Pronto para iniciar.")

    @staticmethod
    def _pick_default_monitor_label(monitor_values: list[str]) -> str:
        for label in monitor_values:
            if label.startswith("Monitor 1 "):
                return label
            if label.startswith("Monitor 1 ("):
                return label
        return monitor_values[0] if monitor_values else ""

    @staticmethod
    def _pick_default_label(values: list[str], marker: str) -> str:
        for label in values:
            if marker in label:
                return label
        return values[0] if values else ""

    def _apply_audio_mode(self) -> None:
        mode = self.audio_mode_var.get()
        need_mic = mode in ("Apenas microfone", "Microfone + som do sistema")
        need_desktop = mode in ("Apenas som do sistema", "Microfone + som do sistema")
        self.mic_combo.configure(state="readonly" if need_mic else "disabled")
        self.desktop_combo.configure(state="readonly" if need_desktop else "disabled")
        for control in self.mic_tuning_controls:
            if need_mic:
                control.state(["!disabled"])
            else:
                control.state(["disabled"])

    def _refresh_mic_tuning_labels(self) -> None:
        gain_value = int(round(float(self.mic_gain_var.get())))
        sensitivity_value = int(round(float(self.mic_sensitivity_var.get())))
        self.mic_gain_text_var.set(f"{gain_value}%")
        self.mic_sensitivity_text_var.set(f"{sensitivity_value}%")

    def _current_mic_tuning(self) -> MicTuning:
        gain = max(50.0, min(250.0, float(self.mic_gain_var.get())))
        sensitivity = max(0.0, min(100.0, float(self.mic_sensitivity_var.get())))
        noise_suppression = bool(self.mic_noise_suppress_var.get())
        return MicTuning(
            gain_percent=gain,
            sensitivity_percent=sensitivity,
            noise_suppression=noise_suppression,
        )

    def _start_recording(self) -> None:
        if self.recording or self.countdown_active:
            return

        out_dir = Path(self.output_dir_var.get()).expanduser()
        if self._is_onedrive_path(out_dir):
            messagebox.showerror(
                "Destino invalido",
                "Salvar em OneDrive foi bloqueado. Escolha uma pasta fora do OneDrive.",
            )
            return
        out_dir.mkdir(parents=True, exist_ok=True)

        file_stem = self._sanitize_file_stem(self.file_name_var.get())
        output_file = out_dir / f"{file_stem}.mp4"
        if output_file.exists():
            stamp = datetime.now().strftime("%H%M%S")
            output_file = out_dir / f"{file_stem}_{stamp}.mp4"

        monitor_index = self.monitor_map.get(self.monitor_var.get())
        if monitor_index is None:
            messagebox.showerror("Erro", "Selecione um monitor valido.")
            return

        region = self._resolve_region_for_monitor(monitor_index, require_select=True)
        if region is None:
            self.status_var.set("Selecao de area cancelada.")
            return

        quality = QUALITY_PROFILES[self.quality_var.get()]
        mode = self.audio_mode_var.get()

        mic_device = None
        system_device = None
        if mode in ("Apenas microfone", "Microfone + som do sistema"):
            mic_device = self.microphone_map.get(self.mic_var.get())
            if mic_device is None:
                messagebox.showerror("Erro", "Nenhum microfone valido selecionado.")
                return

        if mode in ("Apenas som do sistema", "Microfone + som do sistema"):
            system_device = self.desktop_map.get(self.desktop_var.get())
            if system_device is None:
                messagebox.showerror("Erro", "Nenhum dispositivo de som do sistema valido selecionado.")
                return

        options = RecordingOptions(
            output_file=output_file,
            region=region,
            fps=quality["fps"],
            crf=quality["crf"],
            microphone=mic_device,
            system_audio=system_device,
            mic_tuning=self._current_mic_tuning(),
        )
        self._begin_countdown(options, COUNTDOWN_SECONDS)

    def _begin_countdown(self, options: RecordingOptions, seconds: int) -> None:
        self.pending_options = options
        self.countdown_active = True
        self.countdown_left = max(1, seconds)
        self.countdown_sequence = [str(n) for n in range(self.countdown_left, 0, -1)] + ["GO"]
        self.countdown_index = 0
        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="disabled")
        self.mute_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.stop_btn.configure(text="Cancelar")
        self._show_countdown_overlay(options.region)
        self._countdown_tick()

    def _countdown_tick(self) -> None:
        if not self.countdown_active:
            return

        if self.countdown_index >= len(self.countdown_sequence):
            self.countdown_job = None
            self._begin_recording_now()
            return

        token = self.countdown_sequence[self.countdown_index]
        self._update_countdown_overlay(token)
        if token == "GO":
            self.status_var.set("GO!")
            delay_ms = 450
        else:
            self.status_var.set(f"Iniciando gravacao em {token}...")
            delay_ms = 1000
        self.countdown_index += 1
        self.countdown_job = self.after(delay_ms, self._countdown_tick)

    def _begin_recording_now(self) -> None:
        options = self.pending_options
        if options is None:
            self._cancel_countdown_ui("Contagem cancelada.")
            return

        self.pending_options = None
        self.countdown_active = False
        self.countdown_job = None
        self._hide_countdown_overlay()
        self.stop_event.clear()
        self.pause_event.clear()
        self.mute_event.clear()
        self.recording = True
        self.paused = False
        self.muted = False
        self.active_region = options.region

        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="normal", text="Pausar")
        self.mute_btn.configure(state="normal", text="Mudo")
        self.stop_btn.configure(state="normal", text="Parar")
        self.status_var.set("Gravando... clique em Parar para finalizar.")
        self._show_recording_overlays(options.region)

        try:
            self.iconify()
        except Exception:
            pass

        self.worker_thread = threading.Thread(target=self._record_worker, args=(options,), daemon=True)
        self.worker_thread.start()

    def _stop_recording(self) -> None:
        if self.countdown_active:
            self._cancel_countdown_ui("Contagem cancelada.")
            return
        if not self.recording:
            return
        self.status_var.set("Finalizando gravacao...")
        self.stop_event.set()

    def _cancel_countdown_ui(self, status_text: str) -> None:
        self.countdown_active = False
        self.pending_options = None
        if self.countdown_job is not None:
            try:
                self.after_cancel(self.countdown_job)
            except Exception:
                pass
            self.countdown_job = None

        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="Pausar")
        self.mute_btn.configure(state="disabled", text="Mudo")
        self.stop_btn.configure(state="disabled", text="Parar")
        self._hide_countdown_overlay()
        self.status_var.set(status_text)

    def _show_countdown_overlay(self, region: dict[str, int]) -> None:
        self._hide_countdown_overlay()
        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        try:
            overlay.attributes("-alpha", 0.86)
        except Exception:
            pass
        overlay.configure(bg="black")

        box_w = 300
        box_h = 190
        center_x = int(region["left"]) + int(region["width"]) // 2
        center_y = int(region["top"]) + int(region["height"]) // 2
        pos_x = center_x - (box_w // 2)
        pos_y = center_y - (box_h // 2)
        overlay.geometry(f"{box_w}x{box_h}+{pos_x}+{pos_y}")

        label = tk.Label(
            overlay,
            text="",
            font=("Segoe UI", 110, "bold"),
            fg="#ff2a2a",
            bg="black",
        )
        label.place(relx=0.5, rely=0.5, anchor="center")

        self._make_window_click_through(overlay)
        self.countdown_overlay = overlay
        self.countdown_label = label

    def _update_countdown_overlay(self, token: str) -> None:
        if self.countdown_label is not None:
            self.countdown_label.configure(text=token)
            self.countdown_label.update_idletasks()

    def _hide_countdown_overlay(self) -> None:
        if self.countdown_overlay is not None:
            try:
                self.countdown_overlay.destroy()
            except Exception:
                pass
        self.countdown_overlay = None
        self.countdown_label = None

    def _toggle_pause(self) -> None:
        if not self.recording:
            return
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.paused = False
            self.pause_btn.configure(text="Pausar")
            self.status_var.set("Gravando...")
        else:
            self.pause_event.set()
            self.paused = True
            self.pause_btn.configure(text="Retomar")
            self.status_var.set("Pausado.")

    def _toggle_mute(self) -> None:
        if not self.recording:
            return
        if self.mute_event.is_set():
            self.mute_event.clear()
            self.muted = False
            self.mute_btn.configure(text="Mudo")
            self.status_var.set("Audio reativado.")
        else:
            self.mute_event.set()
            self.muted = True
            self.mute_btn.configure(text="Com som")
            self.status_var.set("Audio mutado.")

    def _record_worker(self, options: RecordingOptions) -> None:
        try:
            output, warning = self._run_recording(options)
            self.after(0, lambda out=output, warn=warning: self._record_success(out, warn))
        except Exception as exc:
            error_msg = str(exc)
            self.after(0, lambda msg=error_msg: self._record_error(msg))

    def _record_success(self, output: Path, warning: str | None = None) -> None:
        self._reset_runtime_ui()
        self.last_output_file = output
        self.open_last_btn.configure(state="normal")
        self.status_var.set(f"Concluido: {output}")
        if warning:
            messagebox.showwarning("Gravacao concluida com alerta", warning)
        messagebox.showinfo("Gravacao finalizada", f"Arquivo salvo em:\n{output}")

    def _record_error(self, error: str) -> None:
        self._reset_runtime_ui()
        self.status_var.set("Falha na gravacao.")
        messagebox.showerror("Erro na gravacao", error)

    def _reset_runtime_ui(self) -> None:
        self.recording = False
        self.paused = False
        self.muted = False
        self.active_region = None
        self.stop_event.clear()
        self.pause_event.clear()
        self.mute_event.clear()
        self._hide_recording_overlays()
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="Pausar")
        self.mute_btn.configure(state="disabled", text="Mudo")
        self.stop_btn.configure(state="disabled", text="Parar")
        if self.last_output_file is not None:
            self.open_last_btn.configure(state="normal")

    def _open_last_video_folder(self) -> None:
        if self.last_output_file is None:
            messagebox.showinfo("Abrir pasta", "Nenhum video foi salvo nesta sessao ainda.")
            return

        folder = self.last_output_file.parent
        if not folder.exists():
            messagebox.showerror("Abrir pasta", f"Pasta nao encontrada:\n{folder}")
            return

        try:
            subprocess.run(["explorer", str(folder)], check=False)
        except Exception as exc:
            messagebox.showerror("Abrir pasta", f"Nao foi possivel abrir a pasta:\n{exc}")

    def _show_recording_overlays(self, region: dict[str, int]) -> None:
        self._hide_recording_overlays()
        thickness = 3
        left = int(region["left"])
        top = int(region["top"])
        width = int(region["width"])
        height = int(region["height"])
        self.overlay_windows.extend(
            [
                self._create_overlay_line(left, top, width, thickness),
                self._create_overlay_line(left, top + height - thickness, width, thickness),
                self._create_overlay_line(left, top, thickness, height),
                self._create_overlay_line(left + width - thickness, top, thickness, height),
            ]
        )

    def _hide_recording_overlays(self) -> None:
        for window in self.overlay_windows:
            try:
                window.destroy()
            except Exception:
                pass
        self.overlay_windows.clear()

    def _create_overlay_line(self, x: int, y: int, width: int, height: int) -> tk.Toplevel:
        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.configure(bg="#ff0000")
        overlay.attributes("-topmost", True)
        overlay.geometry(f"{max(1, width)}x{max(1, height)}+{x}+{y}")
        self._make_window_click_through(overlay)
        return overlay

    @staticmethod
    def _make_window_click_through(window: tk.Toplevel) -> None:
        try:
            window.update_idletasks()
            hwnd = int(window.winfo_id())
            user32 = ctypes.windll.user32
            ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ex_style |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)
        except Exception:
            pass

    def _take_screenshot(self) -> None:
        out_dir = Path(self.output_dir_var.get()).expanduser()
        if self._is_onedrive_path(out_dir):
            messagebox.showerror(
                "Destino invalido",
                "Salvar em OneDrive foi bloqueado. Escolha uma pasta fora do OneDrive.",
            )
            return
        out_dir.mkdir(parents=True, exist_ok=True)

        region: dict[str, int] | None = None
        if self.recording and self.active_region is not None:
            region = self.active_region
        else:
            monitor_index = self.monitor_map.get(self.monitor_var.get())
            if monitor_index is None:
                messagebox.showerror("Erro", "Selecione um monitor valido.")
                return
            region = self._resolve_region_for_monitor(monitor_index, require_select=True)
            if region is None:
                self.status_var.set("Selecao de area cancelada.")
                return

        file_name = f"print_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        output_file = out_dir / file_name
        with mss.mss() as sct:
            frame = np.array(sct.grab(region), dtype=np.uint8)[:, :, :3]
        if not cv2.imwrite(str(output_file), frame):
            raise RuntimeError("Nao foi possivel salvar o print.")

        self.status_var.set(f"Print salvo: {output_file}")

    def _resolve_region_for_monitor(
        self,
        monitor_index: int,
        require_select: bool,
    ) -> dict[str, int] | None:
        if monitor_index == CUSTOM_MONITOR_INDEX:
            if require_select or self.custom_region is None:
                selected = self._select_custom_region()
                if selected is None:
                    return None
                self.custom_region = selected
            return self.custom_region

        with mss.mss() as sct:
            monitors = sct.monitors
            if monitor_index < 0 or monitor_index >= len(monitors):
                return None
            mon = monitors[monitor_index]
            return {
                "left": int(mon["left"]),
                "top": int(mon["top"]),
                "width": int(mon["width"]),
                "height": int(mon["height"]),
            }

    def _select_custom_region(self) -> dict[str, int] | None:
        self.status_var.set("Selecione a area personalizada e pressione ENTER.")
        with mss.mss() as sct:
            full = sct.monitors[0]
            frame = np.array(sct.grab(full), dtype=np.uint8)[:, :, :3]

        window_name = "Selecione a area - ENTER confirma / ESC cancela"
        cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
        cv2.setWindowProperty(window_name, cv2.WND_PROP_TOPMOST, 1)
        x, y, w, h = cv2.selectROI(window_name, frame, showCrosshair=True, fromCenter=False)
        cv2.destroyWindow(window_name)

        if w <= 0 or h <= 0:
            return None

        return {
            "left": int(full["left"] + x),
            "top": int(full["top"] + y),
            "width": int(w),
            "height": int(h),
        }

    def _run_recording(self, options: RecordingOptions) -> tuple[Path, str | None]:
        temp_dir = Path(tempfile.mkdtemp(prefix="gravador_tela_"))
        temp_video = temp_dir / "video.avi"
        mic_wav = temp_dir / "mic.wav"
        desktop_wav = temp_dir / "desktop.wav"
        mixed_wav = temp_dir / "mix.wav"

        self.audio_errors = []
        audio_threads: list[threading.Thread] = []

        if options.microphone is not None:
            t_mic = threading.Thread(
                target=self._record_audio_source,
                args=(options.microphone, mic_wav, "microfone", options.mic_tuning),
                daemon=True,
            )
            audio_threads.append(t_mic)
            t_mic.start()

        if options.system_audio is not None:
            t_sys = threading.Thread(
                target=self._record_audio_source,
                args=(options.system_audio, desktop_wav, "som do sistema", None),
                daemon=True,
            )
            audio_threads.append(t_sys)
            t_sys.start()

        self._record_video(temp_video, options.region, options.fps)
        self.stop_event.set()

        for thread in audio_threads:
            thread.join(timeout=3)

        has_mic_audio = mic_wav.exists() and mic_wav.stat().st_size > 44
        has_system_audio = desktop_wav.exists() and desktop_wav.stat().st_size > 44
        audio_inputs = [p for p in (mic_wav, desktop_wav) if p.exists() and p.stat().st_size > 44]
        warning: str | None = None
        if self.audio_errors:
            warning = "Audio parcial: " + " | ".join(self.audio_errors)

        if audio_inputs:
            self._mix_audio(audio_inputs, mixed_wav)
            self._encode_video_audio(temp_video, mixed_wav, options.output_file, options.crf)
        else:
            if options.microphone is not None or options.system_audio is not None:
                if warning:
                    warning += " | Nenhum audio capturado, arquivo salvo sem audio."
                else:
                    warning = "Nenhum audio capturado, arquivo salvo sem audio."
            self._encode_video_only(temp_video, options.output_file, options.crf)

        for path in temp_dir.glob("*"):
            try:
                path.unlink()
            except OSError:
                pass
        try:
            temp_dir.rmdir()
        except OSError:
            pass

        return options.output_file, warning

    def _record_video(self, output_video: Path, region: dict[str, int], fps: int) -> None:
        writer = cv2.VideoWriter(
            str(output_video),
            cv2.VideoWriter_fourcc(*"MJPG"),
            fps,
            (int(region["width"]), int(region["height"])),
        )
        if not writer.isOpened():
            raise RuntimeError("Nao foi possivel iniciar a escrita de video.")

        try:
            with mss.mss() as sct:
                frame_interval = 1.0 / max(1, fps)
                next_write_time = time.perf_counter()
                record_started_at = next_write_time
                pause_started_at: float | None = None
                paused_total = 0.0
                max_catchup_writes = 3
                frames_written = 0
                last_frame: np.ndarray | None = None

                while not self.stop_event.is_set():
                    if self.pause_event.is_set():
                        if pause_started_at is None:
                            pause_started_at = time.perf_counter()
                        time.sleep(0.03)
                        continue

                    now = time.perf_counter()
                    if pause_started_at is not None:
                        # Prevent paused time from accelerating output timeline.
                        pause_duration = now - pause_started_at
                        paused_total += pause_duration
                        next_write_time += pause_duration
                        pause_started_at = None

                    frame = np.array(sct.grab(region), dtype=np.uint8)[:, :, :3]
                    last_frame = frame
                    now_after_grab = time.perf_counter()
                    writes = 0
                    while (
                        now_after_grab >= next_write_time
                        and writes < max_catchup_writes
                        and not self.stop_event.is_set()
                    ):
                        writer.write(frame)
                        writes += 1
                        frames_written += 1
                        next_write_time += frame_interval

                    # Drop excessive backlog to avoid slideshow effect on slower machines.
                    if now_after_grab >= next_write_time:
                        next_write_time = now_after_grab + frame_interval

                    if writes == 0:
                        wait = next_write_time - now_after_grab
                        if wait > 0:
                            time.sleep(min(wait, frame_interval / 2))
        finally:
            if pause_started_at is not None:
                paused_total += max(0.0, time.perf_counter() - pause_started_at)
            if last_frame is not None and frames_written > 0:
                active_duration = max(0.0, time.perf_counter() - record_started_at - paused_total)
                expected_frames = max(frames_written, int(round(active_duration * max(1, fps))))
                for _ in range(expected_frames - frames_written):
                    writer.write(last_frame)
            writer.release()

    def _record_audio_source(
        self,
        device: object,
        wav_path: Path,
        source_name: str,
        mic_tuning: MicTuning | None = None,
    ) -> None:
        device_attempts = self._build_audio_device_attempts(device, source_name)
        channel_attempts = (1, 2) if source_name == "microfone" else (2, 1)

        last_error = ""
        for device_label, audio_device in device_attempts:
            for sample_rate in AUDIO_SAMPLE_RATE_CANDIDATES:
                for channels in channel_attempts:
                    try:
                        self._capture_audio(
                            audio_device,
                            wav_path,
                            channels=channels,
                            samplerate=sample_rate,
                            mic_tuning=mic_tuning if source_name == "microfone" else None,
                        )
                        return
                    except Exception as exc:
                        detail = self._format_exception(exc)
                        last_error = f"{device_label} {sample_rate}Hz/{channels}ch -> {detail}"

        if source_name == "microfone":
            sd_attempts = self._build_sounddevice_attempts(device)
            for sd_label, sd_device_index, max_channels, preferred_rate, hostapi_name in sd_attempts:
                channel_candidates: list[int] = [1]
                if max_channels >= 2:
                    channel_candidates.append(2)

                for sample_rate in self._build_samplerate_attempts(preferred_rate):
                    for channels in channel_candidates:
                        try:
                            self._capture_audio_sounddevice(
                                sd_device_index,
                                wav_path,
                                channels=channels,
                                samplerate=sample_rate,
                                hostapi_name=hostapi_name,
                                mic_tuning=mic_tuning,
                            )
                            return
                        except Exception as exc:
                            detail = self._format_exception(exc)
                            last_error = (
                                f"sounddevice {sd_label} {sample_rate}Hz/{channels}ch"
                                f" ({hostapi_name}) -> {detail}"
                            )
            if sd is None:
                last_error = (
                    f"{last_error} | fallback sounddevice indisponivel"
                    if last_error
                    else "fallback sounddevice indisponivel"
                )

        with self.audio_error_lock:
            if not last_error:
                last_error = "erro desconhecido"
            self.audio_errors.append(f"Falha ao capturar audio ({source_name}): {last_error}")

    def _build_audio_device_attempts(
        self,
        selected_device: object,
        source_name: str,
    ) -> list[tuple[str, object]]:
        attempts: list[tuple[str, object]] = []
        seen_ids: set[str] = set()

        def _add(label: str, audio_device: object | None) -> None:
            if audio_device is None:
                return
            dev_id = str(getattr(audio_device, "id", ""))
            if dev_id and dev_id in seen_ids:
                return
            if dev_id:
                seen_ids.add(dev_id)
            attempts.append((label, audio_device))

        _add("selecionado", selected_device)

        if source_name == "microfone":
            try:
                _add("padrao", sc.default_microphone())
            except Exception:
                pass
            try:
                for idx, mic in enumerate(sc.all_microphones(include_loopback=False), start=1):
                    _add(f"fallback#{idx}", mic)
            except Exception:
                pass
        else:
            try:
                default_speaker = sc.default_speaker()
                _add("padrao", self._get_loopback_for_speaker(default_speaker))
            except Exception:
                pass

        return attempts

    @staticmethod
    def _build_samplerate_attempts(preferred_rate: int) -> list[int]:
        ordered: list[int] = []
        seen: set[int] = set()

        def _add(rate: int | float | None) -> None:
            if rate is None:
                return
            try:
                value = int(round(float(rate)))
            except Exception:
                return
            if value < 8_000 or value > 192_000:
                return
            if value in seen:
                return
            seen.add(value)
            ordered.append(value)

        _add(preferred_rate)
        for rate in AUDIO_SAMPLE_RATE_CANDIDATES:
            _add(rate)
        if not ordered:
            ordered.append(AUDIO_SAMPLE_RATE)
        return ordered

    def _build_sounddevice_attempts(self, selected_device: object) -> list[tuple[str, int, int, int, str]]:
        if sd is None:
            return []

        attempts: list[tuple[str, int, int, int, str]] = []
        seen_indexes: set[int] = set()

        def _add(label: str, idx: int | None, devices: list[object], hostapis: list[object]) -> None:
            if idx is None:
                return
            if idx < 0:
                return
            if idx in seen_indexes:
                return
            if idx >= len(devices):
                return

            dev = devices[idx]
            try:
                max_input = int(dev.get("max_input_channels", 0))
            except Exception:
                max_input = 0
            if max_input <= 0:
                return

            try:
                preferred_rate = int(round(float(dev.get("default_samplerate", 0))))
            except Exception:
                preferred_rate = AUDIO_SAMPLE_RATE

            hostapi_name = "hostapi?"
            try:
                hostapi_idx = int(dev.get("hostapi", -1))
                if 0 <= hostapi_idx < len(hostapis):
                    hostapi_name = str(hostapis[hostapi_idx].get("name", "hostapi?"))
            except Exception:
                hostapi_name = "hostapi?"

            seen_indexes.add(idx)
            attempts.append((label, idx, max_input, preferred_rate, hostapi_name))

        try:
            default_input = int(sd.default.device[0]) if sd.default.device is not None else -1
        except Exception:
            default_input = -1

        try:
            devices = sd.query_devices()
            hostapis = sd.query_hostapis()
        except Exception:
            return attempts

        _add("default", default_input, devices, hostapis)

        selected_name = str(getattr(selected_device, "name", "")).lower().strip()
        selected_core = selected_name.split("(")[0].strip()

        for idx, dev in enumerate(devices):
            try:
                max_input = int(dev.get("max_input_channels", 0))
            except Exception:
                max_input = 0
            if max_input <= 0:
                continue
            dev_name = str(dev.get("name", "")).lower()
            if selected_core and selected_core in dev_name:
                _add(f"match#{idx}", idx, devices, hostapis)

        for idx, dev in enumerate(devices):
            try:
                max_input = int(dev.get("max_input_channels", 0))
            except Exception:
                max_input = 0
            if max_input > 0:
                _add(f"fallback#{idx}", idx, devices, hostapis)

        return attempts

    def _capture_audio(
        self,
        device: object,
        wav_path: Path,
        channels: int,
        samplerate: int,
        mic_tuning: MicTuning | None = None,
    ) -> None:
        with sf.SoundFile(
            str(wav_path),
            mode="w",
            samplerate=samplerate,
            channels=channels,
            subtype="PCM_16",
        ) as wav_file:
            with device.recorder(
                samplerate=samplerate,
                channels=channels,
            ) as recorder:
                while not self.stop_event.is_set():
                    if self.pause_event.is_set():
                        time.sleep(0.03)
                        continue

                    chunk = recorder.record(numframes=AUDIO_BLOCK_SIZE)
                    if chunk is None or not len(chunk):
                        continue
                    if mic_tuning is not None:
                        chunk = self._process_microphone_chunk(chunk, mic_tuning)
                    if self.mute_event.is_set():
                        chunk = np.zeros_like(chunk)
                    wav_file.write(chunk)

    def _capture_audio_sounddevice(
        self,
        device_index: int,
        wav_path: Path,
        channels: int,
        samplerate: int,
        hostapi_name: str = "",
        mic_tuning: MicTuning | None = None,
    ) -> None:
        if sd is None:
            raise RuntimeError("sounddevice indisponivel")

        extra_settings = None
        hostapi_upper = hostapi_name.upper()
        if "WASAPI" in hostapi_upper and hasattr(sd, "WasapiSettings"):
            try:
                extra_settings = sd.WasapiSettings(exclusive=False, auto_convert=True)
            except Exception:
                extra_settings = None

        with sf.SoundFile(
            str(wav_path),
            mode="w",
            samplerate=samplerate,
            channels=channels,
            subtype="PCM_16",
        ) as wav_file:
            try:
                sd.check_input_settings(
                    device=device_index,
                    channels=channels,
                    samplerate=samplerate,
                    dtype="float32",
                    extra_settings=extra_settings,
                )
            except TypeError:
                sd.check_input_settings(
                    device=device_index,
                    channels=channels,
                    samplerate=samplerate,
                    dtype="float32",
                )

            stream_kwargs = {
                "device": device_index,
                "samplerate": samplerate,
                "channels": channels,
                "dtype": "float32",
                "blocksize": AUDIO_BLOCK_SIZE,
            }
            if extra_settings is not None:
                stream_kwargs["extra_settings"] = extra_settings

            with sd.InputStream(**stream_kwargs) as stream:
                while not self.stop_event.is_set():
                    if self.pause_event.is_set():
                        time.sleep(0.03)
                        continue

                    chunk, _overflow = stream.read(AUDIO_BLOCK_SIZE)
                    if chunk is None or len(chunk) == 0:
                        continue
                    if mic_tuning is not None:
                        chunk = self._process_microphone_chunk(chunk, mic_tuning)
                    if self.mute_event.is_set():
                        chunk = np.zeros_like(chunk)
                    wav_file.write(chunk)

    @staticmethod
    def _process_microphone_chunk(chunk: np.ndarray, mic_tuning: MicTuning) -> np.ndarray:
        processed = np.array(chunk, dtype=np.float32, copy=True)

        gain = max(0.5, min(2.5, mic_tuning.gain_percent / 100.0))
        sensitivity = max(0.0, min(100.0, mic_tuning.sensitivity_percent))
        sensitivity_gain = 0.65 + (sensitivity / 100.0) * 0.85
        processed *= gain * sensitivity_gain

        if mic_tuning.noise_suppression:
            threshold = 0.05 - (sensitivity / 100.0) * 0.045
            threshold = max(0.003, min(0.06, threshold))
            low_amp = np.abs(processed) < threshold
            processed[low_amp] *= 0.08

        return np.clip(processed, -1.0, 1.0)

    @staticmethod
    def _format_exception(exc: Exception) -> str:
        message = str(exc).strip()
        if isinstance(exc, AssertionError) and not message:
            return "AssertionError (dispositivo ocupado/incompativel; comum em headset Bluetooth)"
        if message:
            return f"{exc.__class__.__name__}: {message}"
        return exc.__class__.__name__

    def _mix_audio(self, sources: list[Path], output_wav: Path) -> None:
        mixed: list[np.ndarray] = []
        target_sr: int | None = None

        for source in sources:
            data, sample_rate = sf.read(str(source), dtype="float32", always_2d=True)
            if data.size == 0:
                continue

            if target_sr is None:
                target_sr = sample_rate
            elif sample_rate != target_sr:
                data = self._resample_audio(data, sample_rate, target_sr)

            if data.shape[1] == 1:
                data = np.repeat(data, 2, axis=1)
            elif data.shape[1] > 2:
                data = data[:, :2]

            mixed.append(data)

        if not mixed or target_sr is None:
            raise RuntimeError("Nenhum audio util foi capturado.")

        max_len = max(chunk.shape[0] for chunk in mixed)
        mix = np.zeros((max_len, 2), dtype=np.float32)
        for chunk in mixed:
            mix[: chunk.shape[0], : chunk.shape[1]] += chunk

        peak = float(np.max(np.abs(mix)))
        if peak > 1.0:
            mix /= peak
        sf.write(str(output_wav), mix, target_sr, subtype="PCM_16")

    @staticmethod
    def _resample_audio(data: np.ndarray, source_sr: int, target_sr: int) -> np.ndarray:
        if source_sr == target_sr:
            return data
        old_count = data.shape[0]
        new_count = int(round(old_count * (target_sr / source_sr)))
        if new_count <= 0:
            return data

        old_x = np.linspace(0, 1, old_count, endpoint=False)
        new_x = np.linspace(0, 1, new_count, endpoint=False)
        resampled = np.zeros((new_count, data.shape[1]), dtype=np.float32)
        for channel in range(data.shape[1]):
            resampled[:, channel] = np.interp(new_x, old_x, data[:, channel]).astype(np.float32)
        return resampled

    def _encode_video_audio(self, video_file: Path, audio_file: Path, output_file: Path, crf: int) -> None:
        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
        cmd = [
            ffmpeg,
            "-y",
            "-i",
            str(video_file),
            "-i",
            str(audio_file),
            "-c:v",
            "libx264",
            "-preset",
            VIDEO_PRESET,
            "-crf",
            str(crf),
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-shortest",
            str(output_file),
        ]
        self._run_ffmpeg(cmd)

    def _encode_video_only(self, video_file: Path, output_file: Path, crf: int) -> None:
        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
        cmd = [
            ffmpeg,
            "-y",
            "-i",
            str(video_file),
            "-c:v",
            "libx264",
            "-preset",
            VIDEO_PRESET,
            "-crf",
            str(crf),
            "-pix_fmt",
            "yuv420p",
            str(output_file),
        ]
        self._run_ffmpeg(cmd)

    @staticmethod
    def _run_ffmpeg(cmd: list[str]) -> None:
        result = subprocess.run(cmd, capture_output=True, check=False)
        if result.returncode != 0:
            stderr = ScreenRecorderApp._decode_process_output(result.stderr)
            stderr_tail = stderr[-1400:] if stderr else "(sem erro detalhado)"
            raise RuntimeError(f"Falha na codificacao de video/audio.\n{stderr_tail}")

    @staticmethod
    def _decode_process_output(raw: bytes | None) -> str:
        if not raw:
            return ""
        for encoding in ("utf-8", "cp1252", "latin-1"):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace")

    def _apply_hotkeys(self, show_feedback: bool = True) -> None:
        keys = {
            "start": self._normalize_hotkey_key(self.hotkey_start_var.get()),
            "pause": self._normalize_hotkey_key(self.hotkey_pause_var.get()),
            "stop": self._normalize_hotkey_key(self.hotkey_stop_var.get()),
            "mute": self._normalize_hotkey_key(self.hotkey_mute_var.get()),
        }
        if any(value is None for value in keys.values()):
            messagebox.showerror(
                "Atalho invalido",
                "Use apenas letras, numeros ou F1..F12 (sem Ctrl/Shift).",
            )
            return

        if len(set(keys.values())) != 4:
            messagebox.showerror("Atalho invalido", "As teclas de atalho precisam ser diferentes.")
            return

        mapping = {
            self._hotkey_combo(keys["start"]): lambda: self.after(0, self._hotkey_start),
            self._hotkey_combo(keys["pause"]): lambda: self.after(0, self._hotkey_pause),
            self._hotkey_combo(keys["stop"]): lambda: self.after(0, self._hotkey_stop),
            self._hotkey_combo(keys["mute"]): lambda: self.after(0, self._hotkey_mute),
        }

        self._stop_hotkeys()
        try:
            self.hotkey_listener = pynput_keyboard.GlobalHotKeys(mapping)
            self.hotkey_listener.start()
        except Exception as exc:
            self.hotkey_listener = None
            messagebox.showwarning("Atalhos", f"Nao foi possivel ativar atalhos globais:\n{exc}")
            return

        if show_feedback:
            self.status_var.set(
                "Atalhos aplicados: "
                f"Ctrl+Shift+{keys['start']} / Ctrl+Shift+{keys['pause']} / "
                f"Ctrl+Shift+{keys['stop']} / Ctrl+Shift+{keys['mute']}"
            )

    @staticmethod
    def _normalize_hotkey_key(raw_key: str) -> str | None:
        key = raw_key.strip().lower()
        if len(key) == 1 and key.isalnum():
            return key
        if key.startswith("f") and key[1:].isdigit():
            number = int(key[1:])
            if 1 <= number <= 12:
                return f"<f{number}>"
        return None

    @staticmethod
    def _hotkey_combo(key: str) -> str:
        return f"<ctrl>+<shift>+{key}"

    def _hotkey_start(self) -> None:
        if self.recording or self.countdown_active:
            return
        self._start_recording()

    def _hotkey_pause(self) -> None:
        if self.recording:
            self._toggle_pause()

    def _hotkey_stop(self) -> None:
        self._stop_recording()

    def _hotkey_mute(self) -> None:
        if self.recording:
            self._toggle_mute()

    def _stop_hotkeys(self) -> None:
        if self.hotkey_listener is None:
            return
        try:
            self.hotkey_listener.stop()
        except Exception:
            pass
        self.hotkey_listener = None

    @staticmethod
    def _sanitize_file_stem(value: str) -> str:
        banned = '<>:"/\\|?*'
        clean = "".join(ch for ch in value.strip() if ch not in banned)
        return clean or f"gravacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    @staticmethod
    def _is_onedrive_path(path: Path) -> bool:
        normalized = str(path).lower().replace("/", "\\")
        return "\\onedrive\\" in normalized or normalized.endswith("\\onedrive")

    def _on_close(self) -> None:
        self.countdown_active = False
        if self.countdown_job is not None:
            try:
                self.after_cancel(self.countdown_job)
            except Exception:
                pass
            self.countdown_job = None
        self.stop_event.set()
        self._hide_countdown_overlay()
        self._hide_recording_overlays()
        self._stop_hotkeys()
        self.destroy()


def main() -> None:
    app = ScreenRecorderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
