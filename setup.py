import sys
from pathlib import Path

from cx_Freeze import Executable, setup
from cx_Freeze._compat import IS_WINDOWS
from cx_Freeze.hooks import tkinter as tkinter_hook

APP_NAME = "Sistema de Faturamento"
APP_VERSION = "1.0.6"
UPGRADE_CODE = "{8C9B4E4D-2F66-4D89-8B16-4D2E2A4C6F10}"


def _patch_tkinter_hook():
    """Avoid tkinter.Tk() during freeze, which can fail on some Windows installs."""
    original = tkinter_hook.load_tkinter

    def _load_tkinter_safe(finder, module):
        base = Path(sys.base_prefix)
        tcl_root = base / "tcl"
        tcl_dir = tcl_root / "tcl8.6"
        tk_dir = tcl_root / "tk8.6"

        if tcl_dir.is_dir() and tk_dir.is_dir():
            folders = {"TCL_LIBRARY": tcl_dir, "TK_LIBRARY": tk_dir}
            for env_name, source_path in folders.items():
                target_path = f"share/{source_path.name}"
                finder.add_constant(env_name, target_path)
                finder.include_files(source_path, target_path)
                if IS_WINDOWS:
                    dll_name = source_path.name.replace(".", "") + "t.dll"
                    dll_path = Path(sys.base_prefix, "DLLs", dll_name)
                    if dll_path.exists():
                        finder.include_files(dll_path, f"lib/{dll_name}")
            return

        original(finder, module)

    tkinter_hook.load_tkinter = _load_tkinter_safe


_patch_tkinter_hook()

# Define o arquivo executavel
executavel = Executable(
    script="sistema_faturamento.py",
    base="Win32GUI",
    icon="logo.ico",
    target_name=f"{APP_NAME}.exe",
    shortcut_name=APP_NAME,
    shortcut_dir="ProgramMenuFolder",
)

# Dependencias
packages = [
    # pacote local — módulos extraídos em src/
    "src",
    "src.banco",
    "src.cache",
    "src.config",
    "src.dashboard",
    "src.importacao",
    "src.logger",
    "src.relatorios",
    "src.sync",
    "src.utils",
    # dependências externas
    "customtkinter",
    "fitz",
    "pandas",
    "openpyxl",
    "PIL",
    "sqlite3",
    "os",
    "glob",
    "re",
    "atexit",
    "ctypes",
    "unicodedata",
    "calendar",
    "datetime",
    "tkinter",
    "matplotlib",
]

include_files = [
    ("logo.png", "logo.png"),
    ("logo.ico", "logo.ico"),
    ("desinstalar_sistema.bat", "Desinstalar Sistema de Faturamento.bat"),
]

setup(
    name=APP_NAME,
    version=APP_VERSION,
    description="Sistema de gestao de faturamento - Horizonte Logistica",
    author="Horizonte Logistica",
    executables=[executavel],
    options={
        "build_exe": {
            "packages": packages,
            "include_files": include_files,
        },
        "bdist_msi": {
            "upgrade_code": UPGRADE_CODE,
            "all_users": False,
            "install_icon": "logo.ico",
            "initial_target_dir": rf"[LocalAppDataFolder]\Programs\{APP_NAME}.exe",
        },
    },
)
