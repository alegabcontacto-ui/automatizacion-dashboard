"""
agregar_usuario.py
Agrega un nuevo usuario al archivo usuarios.yaml del dashboard.
Ejecutar: python agregar_usuario.py
"""

import sys
from pathlib import Path
import bcrypt
import yaml
from yaml.loader import SafeLoader

USUARIOS_PATH = Path(__file__).parent / "usuarios.yaml"


def main():
    print("=" * 45)
    print("  AGREGAR USUARIO — Monitor IMSS ADEMEX")
    print("=" * 45)

    usuario    = input("\nCorreo / usuario : ").strip()
    nombre     = input("Nombre completo  : ").strip()
    email      = input("Email            : ").strip()
    contrasena = input("Contrasena       : ").strip()

    if not all([usuario, nombre, contrasena]):
        print("\nERROR: todos los campos son obligatorios.")
        sys.exit(1)

    # Generar hash bcrypt
    hashed = bcrypt.hashpw(contrasena.encode(), bcrypt.gensalt()).decode()

    # Cargar YAML existente
    with open(USUARIOS_PATH, encoding="utf-8") as f:
        config = yaml.load(f, Loader=SafeLoader)

    if usuario in config["credentials"]["usernames"]:
        print(f"\nAVISO: el usuario '{usuario}' ya existe. Se sobreescribira.")

    config["credentials"]["usernames"][usuario] = {
        "name":     nombre,
        "email":    email or usuario,
        "password": hashed,
    }

    with open(USUARIOS_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)

    print(f"\nOK - Usuario '{usuario}' ({nombre}) agregado correctamente.")
    print(f"Archivo: {USUARIOS_PATH}")


if __name__ == "__main__":
    main()
