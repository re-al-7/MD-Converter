#!/usr/bin/env python3
"""
setup_startup.py — Configura el arranque automatico de MD-Converter en Windows.

Uso:
    python setup_startup.py install    # Activa el inicio automatico
    python setup_startup.py uninstall  # Desactiva el inicio automatico
    python setup_startup.py status     # Muestra si esta instalado
"""

import sys
import os
from pathlib import Path

APP_NAME = "MD-Converter"
REGISTRY_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _get_registry():
    """Importa winreg (solo disponible en Windows)."""
    try:
        import winreg
        return winreg
    except ImportError:
        print("ERROR: Este script solo funciona en Windows.")
        sys.exit(1)


def _build_command(script_dir: Path) -> str:
    """Devuelve el comando wscript que lanzara run_hidden.vbs."""
    vbs_path = script_dir / "run_hidden.vbs"
    if not vbs_path.exists():
        print(f"ERROR: No se encontro {vbs_path}")
        print("Asegurate de que run_hidden.vbs esta en la misma carpeta que este script.")
        sys.exit(1)
    return f'wscript.exe "{vbs_path}"'


def install():
    winreg = _get_registry()
    script_dir = Path(__file__).parent.resolve()
    command = _build_command(script_dir)

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            REGISTRY_KEY,
            0,
            winreg.KEY_SET_VALUE,
        )
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        winreg.CloseKey(key)
        print(f"[OK] '{APP_NAME}' agregado al inicio automatico de Windows.")
        print(f"     Comando registrado: {command}")
        print()
        print("La aplicacion se iniciara automaticamente la proxima vez que")
        print("inicies sesion en Windows y abrira http://localhost:5000 en el navegador.")
    except PermissionError:
        print("ERROR: Sin permisos para escribir en el registro.")
        print("Intenta ejecutar este script como administrador.")
        sys.exit(1)


def uninstall():
    winreg = _get_registry()

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            REGISTRY_KEY,
            0,
            winreg.KEY_SET_VALUE,
        )
        try:
            winreg.DeleteValue(key, APP_NAME)
            print(f"[OK] '{APP_NAME}' eliminado del inicio automatico.")
        except FileNotFoundError:
            print(f"'{APP_NAME}' no estaba registrado. No hay nada que eliminar.")
        finally:
            winreg.CloseKey(key)
    except PermissionError:
        print("ERROR: Sin permisos para modificar el registro.")
        sys.exit(1)


def status():
    winreg = _get_registry()

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            REGISTRY_KEY,
            0,
            winreg.KEY_READ,
        )
        try:
            value, _ = winreg.QueryValueEx(key, APP_NAME)
            print(f"[INSTALADO] '{APP_NAME}' esta configurado para iniciar con Windows.")
            print(f"  Comando: {value}")
        except FileNotFoundError:
            print(f"[NO INSTALADO] '{APP_NAME}' NO esta en el inicio automatico.")
        finally:
            winreg.CloseKey(key)
    except PermissionError:
        print("ERROR: Sin permisos para leer el registro.")
        sys.exit(1)


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ("install", "uninstall", "status"):
        print(__doc__)
        sys.exit(1)

    action = sys.argv[1]
    if action == "install":
        install()
    elif action == "uninstall":
        uninstall()
    elif action == "status":
        status()


if __name__ == "__main__":
    main()
