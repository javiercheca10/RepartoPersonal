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
    ico_path = reparto_dir / "assets" / "icono.ico"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=RepartoPersonal",
        "--noconfirm",
        "--clean",
        f"--icon={ico_path}",
        "--collect-all=ortools",
        "--collect-all=reportlab",
        "--exclude-module=tests",
        "--exclude-module=unittest",
        "--exclude-module=pytest",
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

    # Preparar la carpeta portable con las plantillas de configuración
    dist_app_dir = BASE_DIR / "dist" / "RepartoPersonal"
    dist_config = dist_app_dir / "configuracion"
    dist_results = dist_app_dir / "resultados"
    dist_results.mkdir(parents=True, exist_ok=True)

    config_src = BASE_DIR / "configuracion"
    if not config_src.exists():
        config_src = BASE_DIR / "configuracion_ejemplo"

    if config_src.exists() and not dist_config.exists():
        import shutil
        shutil.copytree(config_src, dist_config)
        print(f"Plantillas copiadas a: {dist_config}")

    print("\n----------------------------------------------------")
    print("¡Compilación finalizada con éxito!")
    print(f"El ejecutable se encuentra en: {BASE_DIR / 'dist' / 'RepartoPersonal'}")
    print("----------------------------------------------------\n")

if __name__ == "__main__":
    compilar()
