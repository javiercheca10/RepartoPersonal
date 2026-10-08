"""Compilador automático del ejecutable portable para RepartoPersonal."""

import os
from pathlib import Path
import subprocess
import sys

BASE_DIR = Path(__file__).resolve().parent

def compilar():
    print("====================================================")
    print("  COMPILACIÓN DE APLICACIÓN PORTABLE CON PYINSTALLER")
    print("====================================================\n")

    iniciar_py = BASE_DIR / "iniciar.py"
    reparto_dir = BASE_DIR / "reparto"

    # Separador de add-data: ';' en Windows, ':' en Linux/Mac
    sep = ";" if sys.platform.startswith("win") else ":"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=RepartoPersonal",
        "--noconfirm",
        "--clean",
        "--collect-all=ortools",
        "--windowed",          # Sin ventana de consola negra
        "--onedir",            # Formato carpeta portable rápida
        f"--add-data={reparto_dir}{sep}reparto",
        f"--distpath={BASE_DIR / 'dist'}",
        f"--workpath={BASE_DIR / 'build'}",
        f"--specpath={BASE_DIR}",
        str(iniciar_py),
    ]

    print("Ejecutando comando:")
    print(" ".join(cmd), "\n")

    subprocess.run(cmd, check=True)

    print("\n----------------------------------------------------")
    print("¡Compilación finalizada con éxito!")
    print(f"El ejecutable se encuentra en: {BASE_DIR / 'dist' / 'RepartoPersonal'}")
    print("----------------------------------------------------\n")

if __name__ == "__main__":
    compilar()
