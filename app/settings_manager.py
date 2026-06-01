"""
Módulo de persistencia de configuración SMTP.
Guarda y carga la configuración desde un archivo JSON local.
"""

import json
import os
from pathlib import Path
from app.config import SETTINGS_FILE, SMTP_DEFAULTS


def get_settings_path() -> Path:
    """Devuelve la ruta al archivo de configuración."""
    return Path(SETTINGS_FILE)


def save_settings(data: dict) -> bool:
    """Guarda la configuración SMTP en un archivo JSON. Retorna True si tuvo éxito."""
    try:
        path = get_settings_path()
        # No guardar la contraseña en texto plano (omitir)
        safe_data = {k: v for k, v in data.items() if k != "password"}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(safe_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def load_settings() -> dict:
    """
    Carga la configuración SMTP guardada.
    Si no existe el archivo, retorna los valores por defecto.
    """
    path = get_settings_path()
    if not path.exists():
        return dict(SMTP_DEFAULTS)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Combinar con defaults para campos faltantes
        merged = dict(SMTP_DEFAULTS)
        merged.update(data)
        return merged
    except Exception:
        return dict(SMTP_DEFAULTS)
