from cx_Freeze import setup, Executable

APP_NAME = "Sistema de Faturamento"
APP_VERSION = "1.0.6"
UPGRADE_CODE = "{8C9B4E4D-2F66-4D89-8B16-4D2E2A4C6F10}"

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
            "initial_target_dir": rf"[LocalAppDataFolder]\Programs\{APP_NAME}",
        },
    },
)
