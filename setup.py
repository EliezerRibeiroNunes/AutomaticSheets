import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["customtkinter", "openpyxl", "xlrd", "os", "tkinter", "json", "re"],
    "include_files": ["ARQUIVOMODELO.xlsx", "database.json"],
}

base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="SIGAFY-FATURAMENTO",
    version="1.0",
    description="Planilhas de faturamento",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)],
)
