from cx_Freeze import setup, Executable
import os

# Diretório onde está o script principal
base = None

# Dependências adicionais
packages = ["os", "tkinter", "datetime", "shutil", "openpyxl"]

setup(
    name="Unificador de Planilhas",
    version="1.0",
    description="Unificador de planilhas Excel com interface Tkinter.",
    options={
        "build_exe": {
            "packages": packages,
            "include_files": [],
            "excludes": [],
        }
    },
    executables=[Executable("main.py", base=base)]
)
